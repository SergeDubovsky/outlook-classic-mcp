using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;
using ModelContextProtocol;
using ModelContextProtocol.Protocol;
using OutlookClassicMcp.Core.Outlook;
using OutlookClassicMcp.Core.Policy;

namespace OutlookClassicMcp.Transport
{
    internal sealed class OutlookReadToolHandler
    {
        internal const string BodyProtectedWarning =
            "The message body is protected and was not returned.";
        internal const string BodyTruncatedWarning =
            "The message body was truncated to the requested character limit.";
        internal const string RecipientsTruncatedWarning =
            "One or more recipient lists were truncated to the supported maximum.";
        internal const string ResultTruncatedWarning =
            "The result page was truncated to the supported response-size limit.";

        private static readonly IDictionary<string, JsonElement> EmptyArguments =
            new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        private static readonly string[] MailboxRefProperties = { "storeId" };
        private static readonly string[] FolderRefProperties = { "storeId", "entryId" };
        private static readonly string[] ItemRefProperties = { "storeId", "entryId", "itemClass" };
        private static readonly string[] SearchScopeProperties = { "mailbox", "folder" };
        private static readonly string[] SearchFilterProperties =
        {
            "sender",
            "recipient",
            "subject",
            "text",
            "receivedFromUtc",
            "receivedToUtc",
            "isUnread",
            "category",
            "hasAttachments",
        };

        private readonly IOutlookGateway _gateway;
        private readonly HmacCursorCodec _cursorCodec;
        private readonly TimeSpan _toolDeadline;

        public OutlookReadToolHandler(
            IOutlookGateway gateway,
            HmacCursorCodec cursorCodec,
            TimeSpan toolDeadline)
        {
            _gateway = gateway ?? throw new ArgumentNullException(nameof(gateway));
            _cursorCodec = cursorCodec ?? throw new ArgumentNullException(nameof(cursorCodec));
            if (toolDeadline <= TimeSpan.Zero ||
                toolDeadline > RequestLimits.DefaultToolDeadline)
            {
                throw new ArgumentOutOfRangeException(nameof(toolDeadline));
            }

            _toolDeadline = toolDeadline;
        }

        public async ValueTask<CallToolResult> HandleAsync(
            string toolName,
            IDictionary<string, JsonElement>? arguments,
            string operationId,
            CancellationToken cancellationToken)
        {
            if (!OutlookReadCatalog.IsReadTool(toolName))
            {
                throw new ArgumentOutOfRangeException(nameof(toolName));
            }

            ParsedReadCall parsed;
            try
            {
                parsed = ParseCall(toolName, arguments ?? EmptyArguments);
            }
            catch (CursorValidationException)
            {
                return CreateErrorResult(
                    operationId,
                    "INVALID_ARGUMENT",
                    "The tool arguments are invalid.",
                    retryable: false);
            }
            catch (Exception exception) when (IsInvalidArgumentException(exception))
            {
                return CreateErrorResult(
                    operationId,
                    "INVALID_ARGUMENT",
                    "The tool arguments are invalid.",
                    retryable: false);
            }

            object result;
            try
            {
                result = await ExecuteAsync(parsed, cancellationToken).ConfigureAwait(false);
            }
            catch (OutlookGatewayException exception)
            {
                cancellationToken.ThrowIfCancellationRequested();
                return CreateGatewayErrorResult(operationId, exception.Failure);
            }
            catch (Exception)
            {
                cancellationToken.ThrowIfCancellationRequested();
                return CreateErrorResult(
                    operationId,
                    "INTERNAL",
                    "The Outlook read operation failed.",
                    retryable: false);
            }

            if (result == null)
            {
                return CreateErrorResult(
                    operationId,
                    "OUTLOOK_NOT_READY",
                    "The Outlook read result is temporarily unavailable.",
                    retryable: true);
            }

            try
            {
                return CreateSuccessResult(parsed, operationId, result);
            }
            catch (Exception)
            {
                return CreateErrorResult(
                    operationId,
                    "INTERNAL",
                    "The Outlook read result could not be returned safely.",
                    retryable: false);
            }
        }

        private ParsedReadCall ParseCall(
            string toolName,
            IDictionary<string, JsonElement> arguments)
        {
            switch (toolName)
            {
                case ToolNames.OutlookListMailboxes:
                    return ParseListMailboxes(arguments);
                case ToolNames.OutlookListFolders:
                    return ParseListFolders(arguments);
                case ToolNames.OutlookListMessages:
                    return ParseListMessages(arguments);
                case ToolNames.OutlookSearchMessages:
                    return ParseSearchMessages(arguments);
                case ToolNames.OutlookGetMessage:
                    return ParseGetMessage(arguments);
                case ToolNames.OutlookGetConversation:
                    return ParseGetConversation(arguments);
                case ToolNames.OutlookListAttachments:
                    return ParseListAttachments(arguments);
                default:
                    throw new ArgumentOutOfRangeException(nameof(toolName));
            }
        }

        private ParsedReadCall ParseListMailboxes(IDictionary<string, JsonElement> arguments)
        {
            EnsureAllowedArguments(arguments, "pageSize", "cursor");
            var pageSize = ReadPageSize(arguments);
            var queryHash = ComputeQueryHash(
                ToolNames.OutlookListMailboxes,
                writer => { });
            var cursorPayload = DecodeMailboxCursor(ReadCursor(arguments), queryHash);
            return new ParsedReadCall(
                ToolNames.OutlookListMailboxes,
                new OutlookListMailboxesRequest(
                    pageSize,
                    DecodeMailboxAnchor(cursorPayload)),
                HmacCursorKind.ListMailboxes,
                queryHash);
        }

        private ParsedReadCall ParseListFolders(IDictionary<string, JsonElement> arguments)
        {
            EnsureAllowedArguments(arguments, "mailbox", "parentFolder", "pageSize", "cursor");
            var mailbox = ParseMailboxRef(RequireArgument(arguments, "mailbox"));
            var parentFolder = TryGetArgument(arguments, "parentFolder", out var parentValue)
                ? ParseNullableFolderRef(parentValue)
                : null;
            var pageSize = ReadPageSize(arguments);
            var queryHash = ComputeQueryHash(
                ToolNames.OutlookListFolders,
                writer =>
                {
                    WriteMailboxRef(writer, "mailbox", mailbox);
                    WriteFolderRef(writer, "parentFolder", parentFolder);
                });
            var anchor = DecodeFolderAnchor(ReadCursor(arguments), queryHash);
            return new ParsedReadCall(
                ToolNames.OutlookListFolders,
                new OutlookListFoldersRequest(mailbox, parentFolder, pageSize, anchor),
                HmacCursorKind.ListFolders,
                queryHash);
        }

        private ParsedReadCall ParseListMessages(IDictionary<string, JsonElement> arguments)
        {
            EnsureAllowedArguments(arguments, "folder", "pageSize", "cursor");
            var folder = ParseFolderRef(RequireArgument(arguments, "folder"));
            var pageSize = ReadPageSize(arguments);
            var queryHash = ComputeQueryHash(
                ToolNames.OutlookListMessages,
                writer => WriteFolderRef(writer, "folder", folder));
            var anchor = DecodeMessageAnchor(
                ReadCursor(arguments),
                HmacCursorKind.ListMessages,
                queryHash);
            return new ParsedReadCall(
                ToolNames.OutlookListMessages,
                new OutlookListMessagesRequest(folder, pageSize, anchor),
                HmacCursorKind.ListMessages,
                queryHash);
        }

        private ParsedReadCall ParseSearchMessages(IDictionary<string, JsonElement> arguments)
        {
            EnsureAllowedArguments(arguments, "scopes", "filter", "pageSize", "cursor");
            var scopes = ParseSearchScopes(RequireArgument(arguments, "scopes"));
            var filter = TryGetArgument(arguments, "filter", out var filterValue)
                ? ParseSearchFilter(filterValue)
                : new OutlookMessageSearchFilter(
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null);
            var pageSize = ReadPageSize(arguments);
            var queryHash = ComputeSearchQueryHash(scopes, filter);
            var anchor = DecodeMessageAnchor(
                ReadCursor(arguments),
                HmacCursorKind.SearchMessages,
                queryHash);
            return new ParsedReadCall(
                ToolNames.OutlookSearchMessages,
                new OutlookSearchMessagesRequest(scopes, filter, pageSize, anchor),
                HmacCursorKind.SearchMessages,
                queryHash);
        }

        private static ParsedReadCall ParseGetMessage(IDictionary<string, JsonElement> arguments)
        {
            EnsureAllowedArguments(arguments, "item", "bodyFormat", "maximumBodyCharacters");
            var item = ParseItemRef(RequireArgument(arguments, "item"));
            var bodyFormat = OutlookBodyFormat.PlainText;
            if (TryGetArgument(arguments, "bodyFormat", out var bodyFormatValue))
            {
                var label = ReadString(bodyFormatValue);
                switch (label)
                {
                    case "plainText":
                        bodyFormat = OutlookBodyFormat.PlainText;
                        break;
                    case "html":
                        bodyFormat = OutlookBodyFormat.Html;
                        break;
                    default:
                        throw new ArgumentException("The body format is invalid.");
                }
            }

            var maximumBodyCharacters = RequestLimits.DefaultMessageBodyCharacters;
            if (TryGetArgument(arguments, "maximumBodyCharacters", out var maximumBodyValue) &&
                (maximumBodyValue.ValueKind != JsonValueKind.Number ||
                    !maximumBodyValue.TryGetInt32(out maximumBodyCharacters)))
            {
                throw new ArgumentException("The body character limit is invalid.");
            }

            return new ParsedReadCall(
                ToolNames.OutlookGetMessage,
                new OutlookGetMessageRequest(item, bodyFormat, maximumBodyCharacters),
                null,
                string.Empty);
        }

        private ParsedReadCall ParseGetConversation(IDictionary<string, JsonElement> arguments)
        {
            EnsureAllowedArguments(arguments, "item", "pageSize", "cursor");
            var item = ParseItemRef(RequireArgument(arguments, "item"));
            var pageSize = ReadPageSize(arguments);
            var queryHash = ComputeQueryHash(
                ToolNames.OutlookGetConversation,
                writer => WriteItemRef(writer, "item", item));
            var anchor = DecodeMessageAnchor(
                ReadCursor(arguments),
                HmacCursorKind.GetConversation,
                queryHash);
            return new ParsedReadCall(
                ToolNames.OutlookGetConversation,
                new OutlookGetConversationRequest(item, pageSize, anchor),
                HmacCursorKind.GetConversation,
                queryHash);
        }

        private ParsedReadCall ParseListAttachments(IDictionary<string, JsonElement> arguments)
        {
            EnsureAllowedArguments(arguments, "item", "pageSize", "cursor");
            var item = ParseItemRef(RequireArgument(arguments, "item"));
            var pageSize = ReadPageSize(arguments);
            var queryHash = ComputeQueryHash(
                ToolNames.OutlookListAttachments,
                writer => WriteItemRef(writer, "item", item));
            var anchor = DecodeAttachmentAnchor(ReadCursor(arguments), queryHash);
            return new ParsedReadCall(
                ToolNames.OutlookListAttachments,
                new OutlookListAttachmentsRequest(item, pageSize, anchor),
                HmacCursorKind.ListAttachments,
                queryHash);
        }

        private async Task<object> ExecuteAsync(
            ParsedReadCall parsed,
            CancellationToken cancellationToken)
        {
            switch (parsed.ToolName)
            {
                case ToolNames.OutlookListMailboxes:
                    return await ExecuteWithDeadlineAsync(
                        token => _gateway.ListMailboxesAsync(
                            (OutlookListMailboxesRequest)parsed.Request,
                            token),
                        cancellationToken).ConfigureAwait(false);
                case ToolNames.OutlookListFolders:
                    return await ExecuteWithDeadlineAsync(
                        token => _gateway.ListFoldersAsync(
                            (OutlookListFoldersRequest)parsed.Request,
                            token),
                        cancellationToken).ConfigureAwait(false);
                case ToolNames.OutlookListMessages:
                    return await ExecuteWithDeadlineAsync(
                        token => _gateway.ListMessagesAsync(
                            (OutlookListMessagesRequest)parsed.Request,
                            token),
                        cancellationToken).ConfigureAwait(false);
                case ToolNames.OutlookSearchMessages:
                    return await ExecuteWithDeadlineAsync(
                        token => _gateway.SearchMessagesAsync(
                            (OutlookSearchMessagesRequest)parsed.Request,
                            token),
                        cancellationToken).ConfigureAwait(false);
                case ToolNames.OutlookGetMessage:
                    return await ExecuteWithDeadlineAsync(
                        token => _gateway.GetMessageAsync(
                            (OutlookGetMessageRequest)parsed.Request,
                            token),
                        cancellationToken).ConfigureAwait(false);
                case ToolNames.OutlookGetConversation:
                    return await ExecuteWithDeadlineAsync(
                        token => _gateway.GetConversationAsync(
                            (OutlookGetConversationRequest)parsed.Request,
                            token),
                        cancellationToken).ConfigureAwait(false);
                case ToolNames.OutlookListAttachments:
                    return await ExecuteWithDeadlineAsync(
                        token => _gateway.ListAttachmentsAsync(
                            (OutlookListAttachmentsRequest)parsed.Request,
                            token),
                        cancellationToken).ConfigureAwait(false);
                default:
                    throw new ArgumentOutOfRangeException(nameof(parsed));
            }
        }

        private async Task<T> ExecuteWithDeadlineAsync<T>(
            Func<CancellationToken, Task<T>> operation,
            CancellationToken cancellationToken)
        {
            using (var gatewayCancellation =
                CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
            using (var deadlineCancellation =
                CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
            {
                var gatewayTask = operation(gatewayCancellation.Token);
                if (gatewayTask == null)
                {
                    throw new InvalidOperationException("The Outlook gateway returned no task.");
                }

                var deadlineTask = Task.Delay(_toolDeadline, deadlineCancellation.Token);
                var completedTask = await Task.WhenAny(gatewayTask, deadlineTask).ConfigureAwait(false);
                if (completedTask == gatewayTask)
                {
                    TryCancel(deadlineCancellation);
                    cancellationToken.ThrowIfCancellationRequested();
                    return await gatewayTask.ConfigureAwait(false);
                }

                TryCancel(gatewayCancellation);
                ObserveGatewayCompletion(gatewayTask);
                cancellationToken.ThrowIfCancellationRequested();
                throw new OutlookGatewayException(OutlookGatewayFailure.Timeout);
            }
        }

        private CallToolResult CreateSuccessResult(
            ParsedReadCall parsed,
            string operationId,
            object result)
        {
            switch (parsed.ToolName)
            {
                case ToolNames.OutlookListMailboxes:
                    return CreateMailboxPageResult(
                        operationId,
                        parsed,
                        (OutlookMailboxPage)result);
                case ToolNames.OutlookListFolders:
                    return CreateFolderPageResult(
                        operationId,
                        parsed,
                        (OutlookFolderPage)result);
                case ToolNames.OutlookListMessages:
                case ToolNames.OutlookSearchMessages:
                case ToolNames.OutlookGetConversation:
                    return CreateMessagePageResult(
                        operationId,
                        parsed,
                        (OutlookMessagePage)result);
                case ToolNames.OutlookGetMessage:
                    return CreateMessageDetailResult(
                        operationId,
                        (OutlookMessageDetail)result);
                case ToolNames.OutlookListAttachments:
                    return CreateAttachmentPageResult(
                        operationId,
                        parsed,
                        (OutlookAttachmentPage)result);
                default:
                    throw new ArgumentOutOfRangeException(nameof(parsed));
            }
        }

        private CallToolResult CreateMailboxPageResult(
            string operationId,
            ParsedReadCall parsed,
            OutlookMailboxPage page)
        {
            var retainedItemCount = page.Items.Count;
            var resultTruncated = false;
            while (true)
            {
                var mailboxes = new JsonArray();
                for (var index = 0; index < retainedItemCount; index++)
                {
                    mailboxes.Add(SerializeMailbox(page.Items[index]));
                }

                OutlookMailboxKeysetAnchor? nextAnchor;
                if (resultTruncated)
                {
                    var last = page.Items[retainedItemCount - 1];
                    nextAnchor = new OutlookMailboxKeysetAnchor(last.DisplayName, last.Mailbox);
                }
                else
                {
                    nextAnchor = page.NextAnchor;
                }

                var nextCursor = nextAnchor == null
                    ? null
                    : _cursorCodec.Encode(new MailboxCursorPayload(
                        parsed.QueryHash,
                        nextAnchor.DisplayName,
                        nextAnchor.Mailbox.StoreId));
                var data = new JsonObject
                {
                    ["mailboxes"] = mailboxes,
                    ["nextCursor"] = nextCursor,
                    ["resultTruncated"] = resultTruncated,
                };
                var warnings = CreatePageWarnings(resultTruncated);
                var result = TryCreateBoundedSuccessResult(
                    operationId,
                    data,
                    resultTruncated,
                    warnings,
                    $"Outlook returned {retainedItemCount} mailbox records.");
                if (result != null)
                {
                    return result;
                }

                if (retainedItemCount <= 1)
                {
                    return CreateOversizedResultError(operationId);
                }

                retainedItemCount--;
                resultTruncated = true;
            }
        }

        private CallToolResult CreateFolderPageResult(
            string operationId,
            ParsedReadCall parsed,
            OutlookFolderPage page)
        {
            var retainedItemCount = page.Items.Count;
            var resultTruncated = false;
            while (true)
            {
                var folders = new JsonArray();
                for (var index = 0; index < retainedItemCount; index++)
                {
                    folders.Add(SerializeFolder(page.Items[index]));
                }

                OutlookFolderKeysetAnchor? nextAnchor;
                if (resultTruncated)
                {
                    var last = page.Items[retainedItemCount - 1];
                    nextAnchor = new OutlookFolderKeysetAnchor(last.DisplayName, last.Folder);
                }
                else
                {
                    nextAnchor = page.NextAnchor;
                }

                var nextCursor = nextAnchor == null
                    ? null
                    : _cursorCodec.Encode(new FolderCursorPayload(
                        parsed.QueryHash,
                        nextAnchor.DisplayName,
                        nextAnchor.Folder.StoreId,
                        nextAnchor.Folder.EntryId));
                var data = new JsonObject
                {
                    ["folders"] = folders,
                    ["nextCursor"] = nextCursor,
                    ["resultTruncated"] = resultTruncated,
                };
                var warnings = CreatePageWarnings(resultTruncated);
                var result = TryCreateBoundedSuccessResult(
                    operationId,
                    data,
                    resultTruncated,
                    warnings,
                    $"Outlook returned {retainedItemCount} folder records.");
                if (result != null)
                {
                    return result;
                }

                if (retainedItemCount <= 1)
                {
                    return CreateOversizedResultError(operationId);
                }

                retainedItemCount--;
                resultTruncated = true;
            }
        }

        private CallToolResult CreateMessagePageResult(
            string operationId,
            ParsedReadCall parsed,
            OutlookMessagePage page)
        {
            var failures = new JsonArray();
            for (var index = 0; index < page.Failures.Count; index++)
            {
                failures.Add(SerializeScopeFailure(
                    page.Failures[index],
                    FindScopeIndex(parsed.Request, page.Failures[index].Scope)));
            }

            var retainedItemCount = page.Items.Count;
            var resultTruncated = false;
            while (true)
            {
                var messages = new JsonArray();
                for (var index = 0; index < retainedItemCount; index++)
                {
                    messages.Add(SerializeMessage(page.Items[index]));
                }

                OutlookMessageKeysetAnchor? nextAnchor;
                if (page.IsPartial)
                {
                    nextAnchor = null;
                }
                else if (resultTruncated)
                {
                    var last = page.Items[retainedItemCount - 1];
                    nextAnchor = new OutlookMessageKeysetAnchor(
                        last.EffectiveTimestampUtc,
                        last.Item);
                }
                else
                {
                    nextAnchor = page.NextAnchor;
                }

                var nextCursor = nextAnchor == null
                    ? null
                    : _cursorCodec.Encode(new MessageCursorPayload(
                        parsed.CursorKind!.Value,
                        parsed.QueryHash,
                        nextAnchor.EffectiveTimestampUtc.Ticks,
                        nextAnchor.Item.StoreId,
                        nextAnchor.Item.EntryId,
                        nextAnchor.Item.ItemClass));
                var data = new JsonObject
                {
                    ["messages"] = messages,
                    ["nextCursor"] = nextCursor,
                    ["resultTruncated"] = resultTruncated,
                    ["totalScopeCount"] = page.TotalScopeCount,
                    ["scopeFailures"] = failures.DeepClone(),
                };
                var warnings = CreatePageWarnings(resultTruncated);
                var result = TryCreateBoundedSuccessResult(
                    operationId,
                    data,
                    page.IsPartial || resultTruncated,
                    warnings,
                    $"Outlook returned {retainedItemCount} message records.");
                if (result != null)
                {
                    return result;
                }

                if (retainedItemCount <= 1)
                {
                    return CreateOversizedResultError(operationId);
                }

                retainedItemCount--;
                resultTruncated = true;
            }
        }

        private static CallToolResult CreateMessageDetailResult(
            string operationId,
            OutlookMessageDetail detail)
        {
            var warnings = new JsonArray();
            if (detail.Body.IsProtected)
            {
                warnings.Add(BodyProtectedWarning);
            }
            if (detail.Body.IsTruncated)
            {
                warnings.Add(BodyTruncatedWarning);
            }
            if (detail.ToRecipientsTruncated ||
                detail.CcRecipientsTruncated ||
                detail.BccRecipientsTruncated)
            {
                warnings.Add(RecipientsTruncatedWarning);
            }

            var data = new JsonObject
            {
                ["message"] = SerializeMessage(detail.Summary),
                ["toRecipients"] = SerializeAddresses(detail.ToRecipients),
                ["ccRecipients"] = SerializeAddresses(detail.CcRecipients),
                ["bccRecipients"] = SerializeAddresses(detail.BccRecipients),
                ["totalToRecipientCount"] = detail.TotalToRecipientCount,
                ["totalCcRecipientCount"] = detail.TotalCcRecipientCount,
                ["totalBccRecipientCount"] = detail.TotalBccRecipientCount,
                ["toRecipientsTruncated"] = detail.ToRecipientsTruncated,
                ["ccRecipientsTruncated"] = detail.CcRecipientsTruncated,
                ["bccRecipientsTruncated"] = detail.BccRecipientsTruncated,
                ["body"] = new JsonObject
                {
                    ["format"] = MapBodyFormat(detail.Body.Format),
                    ["content"] = detail.Body.Content,
                    ["originalCharacterCount"] = detail.Body.OriginalCharacterCount,
                    ["isTruncated"] = detail.Body.IsTruncated,
                    ["isProtected"] = detail.Body.IsProtected,
                },
            };
            return CreateBoundedSuccessResult(
                operationId,
                data,
                warnings.Count > 0,
                warnings,
                "Outlook returned one bounded message record.");
        }

        private CallToolResult CreateAttachmentPageResult(
            string operationId,
            ParsedReadCall parsed,
            OutlookAttachmentPage page)
        {
            var retainedItemCount = page.Items.Count;
            var resultTruncated = false;
            while (true)
            {
                var attachments = new JsonArray();
                for (var index = 0; index < retainedItemCount; index++)
                {
                    attachments.Add(SerializeAttachment(page.Items[index]));
                }

                OutlookAttachmentKeysetAnchor? nextAnchor;
                if (resultTruncated)
                {
                    var last = page.Items[retainedItemCount - 1].Attachment;
                    nextAnchor = new OutlookAttachmentKeysetAnchor(
                        last.AttachmentIndex,
                        last.MetadataFingerprint);
                }
                else
                {
                    nextAnchor = page.NextAnchor;
                }

                var nextCursor = nextAnchor == null
                    ? null
                    : _cursorCodec.Encode(new AttachmentCursorPayload(
                        parsed.QueryHash,
                        nextAnchor.AttachmentIndex,
                        nextAnchor.MetadataFingerprint));
                var data = new JsonObject
                {
                    ["attachments"] = attachments,
                    ["nextCursor"] = nextCursor,
                    ["resultTruncated"] = resultTruncated,
                };
                var warnings = CreatePageWarnings(resultTruncated);
                var result = TryCreateBoundedSuccessResult(
                    operationId,
                    data,
                    resultTruncated,
                    warnings,
                    $"Outlook returned {retainedItemCount} attachment records.");
                if (result != null)
                {
                    return result;
                }

                if (retainedItemCount <= 1)
                {
                    return CreateOversizedResultError(operationId);
                }

                retainedItemCount--;
                resultTruncated = true;
            }
        }

        private static JsonObject SerializeMailbox(OutlookMailboxSummary mailbox)
        {
            return new JsonObject
            {
                ["mailbox"] = SerializeMailboxRef(mailbox.Mailbox),
                ["displayName"] = mailbox.DisplayName,
                ["storeType"] = MapStoreType(mailbox.StoreType),
                ["capabilities"] = new JsonObject
                {
                    ["isExchangeStore"] = mailbox.Capabilities.IsExchangeStore,
                    ["isDataFileStore"] = mailbox.Capabilities.IsDataFileStore,
                    ["isCachedExchange"] = mailbox.Capabilities.IsCachedExchange,
                },
                ["standardFolders"] = new JsonObject
                {
                    ["inbox"] = SerializeFolderRef(mailbox.StandardFolders.Inbox),
                    ["drafts"] = SerializeFolderRef(mailbox.StandardFolders.Drafts),
                    ["sent"] = SerializeFolderRef(mailbox.StandardFolders.Sent),
                    ["deleted"] = SerializeFolderRef(mailbox.StandardFolders.Deleted),
                    ["archive"] = SerializeFolderRef(mailbox.StandardFolders.Archive),
                },
            };
        }

        private static JsonObject SerializeFolder(OutlookFolderSummary folder)
        {
            return new JsonObject
            {
                ["folder"] = SerializeFolderRef(folder.Folder),
                ["parentFolder"] = SerializeFolderRef(folder.ParentFolder),
                ["displayName"] = folder.DisplayName,
                ["hasChildren"] = folder.HasChildren,
            };
        }

        private static JsonObject SerializeMessage(OutlookMessageSummary message)
        {
            return new JsonObject
            {
                ["item"] = SerializeItemRef(message.Item),
                ["folder"] = SerializeFolderRef(message.Folder),
                ["subject"] = message.Subject,
                ["senderDisplayName"] = message.SenderDisplayName,
                ["senderAddress"] = message.SenderAddress,
                ["effectiveTimestampUtc"] = FormatUtc(message.EffectiveTimestampUtc),
                ["receivedUtc"] = FormatUtc(message.ReceivedUtc),
                ["sentUtc"] = FormatUtc(message.SentUtc),
                ["isRead"] = message.IsRead,
                ["attachmentCount"] = message.AttachmentCount,
                ["hasAttachments"] = message.HasAttachments,
                ["conversationId"] = message.ConversationId,
            };
        }

        private static JsonObject SerializeAttachment(OutlookAttachmentSummary summary)
        {
            var attachment = summary.Attachment;
            return new JsonObject
            {
                ["attachment"] = new JsonObject
                {
                    ["item"] = SerializeItemRef(attachment.Item),
                    ["attachmentIndex"] = attachment.AttachmentIndex,
                    ["name"] = attachment.Name,
                    ["size"] = attachment.Size,
                    ["sizeIsKnown"] = attachment.SizeIsKnown,
                    ["metadataFingerprint"] = attachment.MetadataFingerprint,
                },
                ["contentType"] = summary.ContentType,
            };
        }

        private static JsonArray SerializeAddresses(
            IReadOnlyList<OutlookMessageAddress> addresses)
        {
            var values = new JsonArray();
            for (var index = 0; index < addresses.Count; index++)
            {
                values.Add(new JsonObject
                {
                    ["displayName"] = addresses[index].DisplayName,
                    ["address"] = addresses[index].Address,
                });
            }

            return values;
        }

        private static JsonObject SerializeScopeFailure(
            OutlookScopeFailure failure,
            int scopeIndex)
        {
            return new JsonObject
            {
                ["scopeIndex"] = scopeIndex,
                ["scope"] = new JsonObject
                {
                    ["mailbox"] = SerializeMailboxRef(failure.Scope.Mailbox),
                    ["folder"] = SerializeFolderRef(failure.Scope.Folder),
                },
                ["code"] = MapFailureCode(failure.Failure),
                ["message"] = MapFailureMessage(failure.Failure),
                ["retryable"] = IsRetryable(failure.Failure),
            };
        }

        private static CallToolResult CreateBoundedSuccessResult(
            string operationId,
            JsonObject data,
            bool partial,
            JsonArray warnings,
            string text)
        {
            return TryCreateBoundedSuccessResult(
                operationId,
                data,
                partial,
                warnings,
                text) ?? CreateOversizedResultError(operationId);
        }

        private static CallToolResult? TryCreateBoundedSuccessResult(
            string operationId,
            JsonObject data,
            bool partial,
            JsonArray warnings,
            string text)
        {
            var envelope = new JsonObject
            {
                ["ok"] = true,
                ["operationId"] = operationId,
                ["data"] = data,
                ["partial"] = partial,
                ["warnings"] = warnings,
            };
            var bytes = JsonSerializer.SerializeToUtf8Bytes(
                envelope,
                McpJsonUtilities.DefaultOptions);
            try
            {
                if (bytes.Length > RequestLimits.MaximumToolResultBytes)
                {
                    return null;
                }

                using (var document = JsonDocument.Parse(bytes))
                {
                    var result = new CallToolResult
                    {
                        IsError = false,
                        StructuredContent = document.RootElement.Clone(),
                        Content = new List<ContentBlock>
                        {
                            new TextContentBlock { Text = text },
                        },
                    };
                    var resultBytes = JsonSerializer.SerializeToUtf8Bytes(
                        result,
                        McpJsonUtilities.DefaultOptions);
                    try
                    {
                        return resultBytes.Length <= RequestLimits.MaximumToolResultBytes
                            ? result
                            : null;
                    }
                    finally
                    {
                        Array.Clear(resultBytes, 0, resultBytes.Length);
                    }
                }
            }
            finally
            {
                Array.Clear(bytes, 0, bytes.Length);
            }
        }

        private static JsonArray CreatePageWarnings(bool resultTruncated)
        {
            var warnings = new JsonArray();
            if (resultTruncated)
            {
                warnings.Add(ResultTruncatedWarning);
            }

            return warnings;
        }

        private static CallToolResult CreateOversizedResultError(string operationId)
        {
            return CreateErrorResult(
                operationId,
                "INTERNAL",
                "The Outlook result exceeded the supported size limit.",
                retryable: false);
        }

        private static CallToolResult CreateGatewayErrorResult(
            string operationId,
            OutlookGatewayFailure failure)
        {
            return CreateErrorResult(
                operationId,
                MapFailureCode(failure),
                MapFailureMessage(failure),
                IsRetryable(failure));
        }

        private static CallToolResult CreateErrorResult(
            string operationId,
            string code,
            string message,
            bool retryable)
        {
            var envelope = new JsonObject
            {
                ["ok"] = false,
                ["operationId"] = operationId,
                ["error"] = new JsonObject
                {
                    ["code"] = code,
                    ["message"] = message,
                    ["retryable"] = retryable,
                    ["details"] = new JsonObject(),
                },
            };
            return new CallToolResult
            {
                IsError = true,
                StructuredContent = JsonSerializer.SerializeToElement(
                    envelope,
                    McpJsonUtilities.DefaultOptions),
                Content = new List<ContentBlock>
                {
                    new TextContentBlock { Text = message },
                },
            };
        }

        private MailboxCursorPayload? DecodeMailboxCursor(string? cursor, string queryHash)
        {
            if (cursor == null)
            {
                return null;
            }

            if (!_cursorCodec.TryDecode(
                cursor,
                HmacCursorKind.ListMailboxes,
                queryHash,
                out var payload) ||
                !(payload is MailboxCursorPayload mailbox))
            {
                throw new CursorValidationException();
            }

            return mailbox;
        }

        private OutlookFolderKeysetAnchor? DecodeFolderAnchor(string? cursor, string queryHash)
        {
            if (cursor == null)
            {
                return null;
            }

            if (!_cursorCodec.TryDecode(
                cursor,
                HmacCursorKind.ListFolders,
                queryHash,
                out var payload) ||
                !(payload is FolderCursorPayload folder))
            {
                throw new CursorValidationException();
            }

            return new OutlookFolderKeysetAnchor(
                folder.DisplayName,
                new FolderRef(folder.StoreId, folder.EntryId));
        }

        private OutlookMessageKeysetAnchor? DecodeMessageAnchor(
            string? cursor,
            HmacCursorKind kind,
            string queryHash)
        {
            if (cursor == null)
            {
                return null;
            }

            if (!_cursorCodec.TryDecode(cursor, kind, queryHash, out var payload) ||
                !(payload is MessageCursorPayload message))
            {
                throw new CursorValidationException();
            }

            return new OutlookMessageKeysetAnchor(
                new DateTime(message.TimestampUtcTicks, DateTimeKind.Utc),
                new ItemRef(message.StoreId, message.EntryId, message.ItemClass));
        }

        private OutlookAttachmentKeysetAnchor? DecodeAttachmentAnchor(
            string? cursor,
            string queryHash)
        {
            if (cursor == null)
            {
                return null;
            }

            if (!_cursorCodec.TryDecode(
                cursor,
                HmacCursorKind.ListAttachments,
                queryHash,
                out var payload) ||
                !(payload is AttachmentCursorPayload attachment))
            {
                throw new CursorValidationException();
            }

            return new OutlookAttachmentKeysetAnchor(
                attachment.AttachmentIndex,
                attachment.MetadataFingerprint);
        }

        private static OutlookMailboxKeysetAnchor? DecodeMailboxAnchor(
            MailboxCursorPayload? payload)
        {
            return payload == null
                ? null
                : new OutlookMailboxKeysetAnchor(
                    payload.DisplayName,
                    new MailboxRef(payload.StoreId));
        }

        private static MailboxRef ParseMailboxRef(JsonElement value)
        {
            EnsureObjectProperties(value, MailboxRefProperties, "storeId");
            return new MailboxRef(ReadString(value.GetProperty("storeId")));
        }

        private static FolderRef ParseFolderRef(JsonElement value)
        {
            EnsureObjectProperties(value, FolderRefProperties, "storeId", "entryId");
            return new FolderRef(
                ReadString(value.GetProperty("storeId")),
                ReadString(value.GetProperty("entryId")));
        }

        private static FolderRef? ParseNullableFolderRef(JsonElement value)
        {
            return value.ValueKind == JsonValueKind.Null ? null : ParseFolderRef(value);
        }

        private static ItemRef ParseItemRef(JsonElement value)
        {
            EnsureObjectProperties(
                value,
                ItemRefProperties,
                "storeId",
                "entryId",
                "itemClass");
            return new ItemRef(
                ReadString(value.GetProperty("storeId")),
                ReadString(value.GetProperty("entryId")),
                ReadString(value.GetProperty("itemClass")));
        }

        private static ReadOnlyCollection<OutlookSearchScope> ParseSearchScopes(JsonElement value)
        {
            if (value.ValueKind != JsonValueKind.Array ||
                value.GetArrayLength() < 1 ||
                value.GetArrayLength() > OutlookReadLimits.MaximumSearchScopeCount)
            {
                throw new ArgumentException("The search scopes are invalid.");
            }

            var scopes = new List<OutlookSearchScope>(value.GetArrayLength());
            foreach (var scopeValue in value.EnumerateArray())
            {
                EnsureObjectProperties(
                    scopeValue,
                    SearchScopeProperties,
                    "mailbox");
                var mailbox = ParseMailboxRef(scopeValue.GetProperty("mailbox"));
                var folder = scopeValue.TryGetProperty("folder", out var folderValue)
                    ? ParseNullableFolderRef(folderValue)
                    : null;
                scopes.Add(new OutlookSearchScope(mailbox, folder));
            }

            return scopes.AsReadOnly();
        }

        private static OutlookMessageSearchFilter ParseSearchFilter(JsonElement value)
        {
            EnsureObjectProperties(value, SearchFilterProperties);
            return new OutlookMessageSearchFilter(
                ReadOptionalString(value, "sender"),
                ReadOptionalString(value, "recipient"),
                ReadOptionalString(value, "subject"),
                ReadOptionalString(value, "text"),
                ReadOptionalUtc(value, "receivedFromUtc"),
                ReadOptionalUtc(value, "receivedToUtc"),
                ReadOptionalBoolean(value, "isUnread"),
                ReadOptionalString(value, "category"),
                ReadOptionalBoolean(value, "hasAttachments"));
        }

        private static int ReadPageSize(IDictionary<string, JsonElement> arguments)
        {
            if (!TryGetArgument(arguments, "pageSize", out var value))
            {
                return RequestLimits.DefaultReadPageSize;
            }

            if (value.ValueKind != JsonValueKind.Number || !value.TryGetInt32(out var pageSize))
            {
                throw new ArgumentException("The page size is invalid.");
            }

            return pageSize;
        }

        private static string? ReadCursor(IDictionary<string, JsonElement> arguments)
        {
            if (!TryGetArgument(arguments, "cursor", out var value))
            {
                return null;
            }

            var cursor = ReadString(value);
            if (cursor.Length > HmacCursorCodec.MaximumCursorLength)
            {
                throw new CursorValidationException();
            }

            return cursor;
        }

        private static JsonElement RequireArgument(
            IDictionary<string, JsonElement> arguments,
            string name)
        {
            if (!TryGetArgument(arguments, name, out var value))
            {
                throw new ArgumentException("A required tool argument is missing.");
            }

            return value;
        }

        private static bool TryGetArgument(
            IDictionary<string, JsonElement> arguments,
            string name,
            out JsonElement value)
        {
            return arguments.TryGetValue(name, out value);
        }

        private static void EnsureAllowedArguments(
            IDictionary<string, JsonElement> arguments,
            params string[] allowedNames)
        {
            foreach (var argument in arguments)
            {
                var allowed = false;
                for (var index = 0; index < allowedNames.Length; index++)
                {
                    if (string.Equals(argument.Key, allowedNames[index], StringComparison.Ordinal))
                    {
                        allowed = true;
                        break;
                    }
                }

                if (!allowed)
                {
                    throw new ArgumentException("An unsupported tool argument was supplied.");
                }
            }
        }

        private static void EnsureObjectProperties(
            JsonElement value,
            string[] allowedNames,
            params string[] requiredNames)
        {
            if (value.ValueKind != JsonValueKind.Object)
            {
                throw new ArgumentException("A tool argument object is invalid.");
            }

            var observed = new HashSet<string>(StringComparer.Ordinal);
            foreach (var property in value.EnumerateObject())
            {
                if (!observed.Add(property.Name))
                {
                    throw new ArgumentException("A tool argument property was duplicated.");
                }

                var allowed = false;
                for (var index = 0; index < allowedNames.Length; index++)
                {
                    if (string.Equals(property.Name, allowedNames[index], StringComparison.Ordinal))
                    {
                        allowed = true;
                        break;
                    }
                }

                if (!allowed)
                {
                    throw new ArgumentException("An unsupported nested tool argument was supplied.");
                }
            }

            for (var index = 0; index < requiredNames.Length; index++)
            {
                if (!observed.Contains(requiredNames[index]))
                {
                    throw new ArgumentException("A required nested tool argument is missing.");
                }
            }
        }

        private static string ReadString(JsonElement value)
        {
            if (value.ValueKind != JsonValueKind.String)
            {
                throw new ArgumentException("A string tool argument is invalid.");
            }

            return value.GetString() ?? throw new ArgumentException("A string tool argument is invalid.");
        }

        private static string? ReadOptionalString(JsonElement value, string propertyName)
        {
            return value.TryGetProperty(propertyName, out var property)
                ? ReadString(property)
                : null;
        }

        private static bool? ReadOptionalBoolean(JsonElement value, string propertyName)
        {
            if (!value.TryGetProperty(propertyName, out var property))
            {
                return null;
            }

            if (property.ValueKind != JsonValueKind.True &&
                property.ValueKind != JsonValueKind.False)
            {
                throw new ArgumentException("A Boolean tool argument is invalid.");
            }

            return property.GetBoolean();
        }

        private static DateTime? ReadOptionalUtc(JsonElement value, string propertyName)
        {
            if (!value.TryGetProperty(propertyName, out var property))
            {
                return null;
            }

            var text = ReadString(property);
            if (!DateTimeOffset.TryParse(
                    text,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.RoundtripKind,
                    out var timestamp) ||
                timestamp.Offset != TimeSpan.Zero)
            {
                throw new ArgumentException("A UTC timestamp tool argument is invalid.");
            }

            return timestamp.UtcDateTime;
        }

        private static string ComputeSearchQueryHash(
            IReadOnlyList<OutlookSearchScope> scopes,
            OutlookMessageSearchFilter filter)
        {
            var sortedScopes = new List<OutlookSearchScope>(scopes);
            sortedScopes.Sort(CompareScopes);
            return ComputeQueryHash(
                ToolNames.OutlookSearchMessages,
                writer =>
                {
                    writer.WritePropertyName("scopes");
                    writer.WriteStartArray();
                    for (var index = 0; index < sortedScopes.Count; index++)
                    {
                        writer.WriteStartObject();
                        WriteMailboxRef(writer, "mailbox", sortedScopes[index].Mailbox);
                        WriteFolderRef(writer, "folder", sortedScopes[index].Folder);
                        writer.WriteEndObject();
                    }
                    writer.WriteEndArray();
                    writer.WritePropertyName("filter");
                    writer.WriteStartObject();
                    WriteNullableString(writer, "sender", filter.Sender);
                    WriteNullableString(writer, "recipient", filter.Recipient);
                    WriteNullableString(writer, "subject", filter.Subject);
                    WriteNullableString(writer, "text", filter.Text);
                    WriteNullableUtc(writer, "receivedFromUtc", filter.ReceivedFromUtc);
                    WriteNullableUtc(writer, "receivedToUtc", filter.ReceivedToUtc);
                    WriteNullableBoolean(writer, "isUnread", filter.IsUnread);
                    WriteNullableString(writer, "category", filter.Category);
                    WriteNullableBoolean(writer, "hasAttachments", filter.HasAttachments);
                    writer.WriteEndObject();
                });
        }

        private static string ComputeQueryHash(
            string toolName,
            Action<Utf8JsonWriter> writeArguments)
        {
            using (var stream = new MemoryStream())
            {
                using (var writer = new Utf8JsonWriter(stream))
                {
                    writer.WriteStartObject();
                    writer.WriteNumber("v", 1);
                    writer.WriteString("tool", toolName);
                    writeArguments(writer);
                    writer.WriteEndObject();
                    writer.Flush();
                }

                var bytes = stream.ToArray();
                byte[] hash;
                using (var sha256 = SHA256.Create())
                {
                    hash = sha256.ComputeHash(bytes);
                }

                try
                {
                    var builder = new StringBuilder(hash.Length * 2);
                    for (var index = 0; index < hash.Length; index++)
                    {
                        builder.Append(hash[index].ToString("x2", CultureInfo.InvariantCulture));
                    }

                    return builder.ToString();
                }
                finally
                {
                    Array.Clear(bytes, 0, bytes.Length);
                    Array.Clear(hash, 0, hash.Length);
                }
            }
        }

        private static void WriteMailboxRef(
            Utf8JsonWriter writer,
            string propertyName,
            MailboxRef mailbox)
        {
            writer.WritePropertyName(propertyName);
            writer.WriteStartObject();
            writer.WriteString("storeId", mailbox.StoreId);
            writer.WriteEndObject();
        }

        private static void WriteFolderRef(
            Utf8JsonWriter writer,
            string propertyName,
            FolderRef? folder)
        {
            writer.WritePropertyName(propertyName);
            if (folder == null)
            {
                writer.WriteNullValue();
                return;
            }

            writer.WriteStartObject();
            writer.WriteString("storeId", folder.StoreId);
            writer.WriteString("entryId", folder.EntryId);
            writer.WriteEndObject();
        }

        private static void WriteItemRef(
            Utf8JsonWriter writer,
            string propertyName,
            ItemRef item)
        {
            writer.WritePropertyName(propertyName);
            writer.WriteStartObject();
            writer.WriteString("storeId", item.StoreId);
            writer.WriteString("entryId", item.EntryId);
            writer.WriteString("itemClass", item.ItemClass);
            writer.WriteEndObject();
        }

        private static void WriteNullableString(
            Utf8JsonWriter writer,
            string propertyName,
            string? value)
        {
            if (value == null)
            {
                writer.WriteNull(propertyName);
            }
            else
            {
                writer.WriteString(propertyName, value);
            }
        }

        private static void WriteNullableUtc(
            Utf8JsonWriter writer,
            string propertyName,
            DateTime? value)
        {
            WriteNullableString(writer, propertyName, FormatUtc(value));
        }

        private static void WriteNullableBoolean(
            Utf8JsonWriter writer,
            string propertyName,
            bool? value)
        {
            if (value.HasValue)
            {
                writer.WriteBoolean(propertyName, value.Value);
            }
            else
            {
                writer.WriteNull(propertyName);
            }
        }

        private static int CompareScopes(OutlookSearchScope left, OutlookSearchScope right)
        {
            var storeComparison = string.Compare(
                left.Mailbox.StoreId,
                right.Mailbox.StoreId,
                StringComparison.Ordinal);
            if (storeComparison != 0)
            {
                return storeComparison;
            }

            return string.Compare(
                left.Folder?.EntryId ?? string.Empty,
                right.Folder?.EntryId ?? string.Empty,
                StringComparison.Ordinal);
        }

        private static int FindScopeIndex(object request, OutlookSearchScope scope)
        {
            if (!(request is OutlookSearchMessagesRequest searchRequest))
            {
                return 0;
            }

            for (var index = 0; index < searchRequest.Scopes.Count; index++)
            {
                var candidate = searchRequest.Scopes[index];
                if (string.Equals(
                        candidate.Mailbox.StoreId,
                        scope.Mailbox.StoreId,
                        StringComparison.Ordinal) &&
                    string.Equals(
                        candidate.Folder?.EntryId,
                        scope.Folder?.EntryId,
                        StringComparison.Ordinal))
                {
                    return index;
                }
            }

            throw new InvalidOperationException("The scope failure did not match the request.");
        }

        private static JsonObject SerializeMailboxRef(MailboxRef mailbox)
        {
            return new JsonObject { ["storeId"] = mailbox.StoreId };
        }

        private static JsonObject? SerializeFolderRef(FolderRef? folder)
        {
            return folder == null
                ? null
                : new JsonObject
                {
                    ["storeId"] = folder.StoreId,
                    ["entryId"] = folder.EntryId,
                };
        }

        private static JsonObject SerializeItemRef(ItemRef item)
        {
            return new JsonObject
            {
                ["storeId"] = item.StoreId,
                ["entryId"] = item.EntryId,
                ["itemClass"] = item.ItemClass,
            };
        }

        private static string FormatUtc(DateTime value)
        {
            return value.ToString("O", CultureInfo.InvariantCulture);
        }

        private static string? FormatUtc(DateTime? value)
        {
            return value.HasValue ? FormatUtc(value.Value) : null;
        }

        private static string MapBodyFormat(OutlookBodyFormat format)
        {
            switch (format)
            {
                case OutlookBodyFormat.PlainText:
                    return "plainText";
                case OutlookBodyFormat.Html:
                    return "html";
                default:
                    throw new ArgumentOutOfRangeException(nameof(format));
            }
        }

        private static string MapStoreType(OutlookStoreType storeType)
        {
            switch (storeType)
            {
                case OutlookStoreType.PrimaryExchangeMailbox:
                    return "primaryExchangeMailbox";
                case OutlookStoreType.ExchangeMailbox:
                    return "exchangeMailbox";
                case OutlookStoreType.ExchangePublicFolder:
                    return "exchangePublicFolder";
                case OutlookStoreType.AdditionalExchangeMailbox:
                    return "additionalExchangeMailbox";
                case OutlookStoreType.NonExchange:
                    return "nonExchange";
                case OutlookStoreType.Unknown:
                    return "unknown";
                default:
                    throw new ArgumentOutOfRangeException(nameof(storeType));
            }
        }

        private static string MapFailureCode(OutlookGatewayFailure failure)
        {
            switch (failure)
            {
                case OutlookGatewayFailure.NotReady:
                    return "OUTLOOK_NOT_READY";
                case OutlookGatewayFailure.Degraded:
                    return "HOST_DEGRADED";
                case OutlookGatewayFailure.Paused:
                    return "HOST_PAUSED";
                case OutlookGatewayFailure.Stopping:
                    return "HOST_STOPPING";
                case OutlookGatewayFailure.StoreNotFound:
                    return "STORE_NOT_FOUND";
                case OutlookGatewayFailure.FolderNotFound:
                    return "FOLDER_NOT_FOUND";
                case OutlookGatewayFailure.ItemNotFound:
                    return "ITEM_NOT_FOUND";
                case OutlookGatewayFailure.ItemMovedOrDeleted:
                    return "ITEM_MOVED_OR_DELETED";
                case OutlookGatewayFailure.UnsupportedStore:
                    return "UNSUPPORTED_STORE";
                case OutlookGatewayFailure.UnsupportedItemType:
                    return "UNSUPPORTED_ITEM_TYPE";
                case OutlookGatewayFailure.AccessDenied:
                    return "ACCESS_DENIED";
                case OutlookGatewayFailure.ObjectModelGuard:
                    return "OBJECT_MODEL_GUARD";
                case OutlookGatewayFailure.InvalidArgument:
                    return "INVALID_ARGUMENT";
                case OutlookGatewayFailure.QueueFull:
                    return "QUEUE_FULL";
                case OutlookGatewayFailure.Timeout:
                    return "TIMEOUT";
                case OutlookGatewayFailure.ComBusy:
                    return "COM_BUSY";
                case OutlookGatewayFailure.StaDispatchFailed:
                    return "STA_DISPATCH_FAILED";
                case OutlookGatewayFailure.CursorStale:
                    return "CURSOR_STALE";
                case OutlookGatewayFailure.Internal:
                    return "INTERNAL";
                default:
                    throw new ArgumentOutOfRangeException(nameof(failure));
            }
        }

        private static string MapFailureMessage(OutlookGatewayFailure failure)
        {
            switch (failure)
            {
                case OutlookGatewayFailure.NotReady:
                    return "Outlook is not ready.";
                case OutlookGatewayFailure.Degraded:
                    return "The Outlook MCP host is degraded.";
                case OutlookGatewayFailure.Paused:
                    return "The Outlook MCP host is paused.";
                case OutlookGatewayFailure.Stopping:
                    return "The Outlook MCP host is stopping.";
                case OutlookGatewayFailure.StoreNotFound:
                    return "The mailbox store no longer exists.";
                case OutlookGatewayFailure.FolderNotFound:
                    return "The folder no longer exists.";
                case OutlookGatewayFailure.ItemNotFound:
                    return "The message no longer exists.";
                case OutlookGatewayFailure.ItemMovedOrDeleted:
                    return "The message was moved or deleted.";
                case OutlookGatewayFailure.UnsupportedStore:
                    return "The mailbox store is not supported for this operation.";
                case OutlookGatewayFailure.UnsupportedItemType:
                    return "The Outlook item type is not supported.";
                case OutlookGatewayFailure.AccessDenied:
                    return "Outlook denied access to the requested data.";
                case OutlookGatewayFailure.ObjectModelGuard:
                    return "Outlook blocked access through the Object Model Guard.";
                case OutlookGatewayFailure.InvalidArgument:
                    return "The request is invalid.";
                case OutlookGatewayFailure.QueueFull:
                    return "The Outlook operation queue is full.";
                case OutlookGatewayFailure.Timeout:
                    return "The Outlook operation timed out.";
                case OutlookGatewayFailure.ComBusy:
                    return "Outlook is busy.";
                case OutlookGatewayFailure.StaDispatchFailed:
                    return "The Outlook UI dispatcher is unavailable.";
                case OutlookGatewayFailure.CursorStale:
                    return "The continuation cursor is stale.";
                case OutlookGatewayFailure.Internal:
                    return "The Outlook read operation failed.";
                default:
                    throw new ArgumentOutOfRangeException(nameof(failure));
            }
        }

        private static bool IsRetryable(OutlookGatewayFailure failure)
        {
            switch (failure)
            {
                case OutlookGatewayFailure.NotReady:
                case OutlookGatewayFailure.Degraded:
                case OutlookGatewayFailure.Paused:
                case OutlookGatewayFailure.Stopping:
                case OutlookGatewayFailure.QueueFull:
                case OutlookGatewayFailure.Timeout:
                case OutlookGatewayFailure.ComBusy:
                    return true;
                default:
                    return false;
            }
        }

        private static bool IsInvalidArgumentException(Exception exception)
        {
            return exception is ArgumentException ||
                exception is FormatException ||
                exception is OverflowException;
        }

        private static void TryCancel(CancellationTokenSource cancellation)
        {
            try
            {
                cancellation.Cancel();
            }
            catch (Exception)
            {
            }
        }

        private static void ObserveGatewayCompletion(Task gatewayTask)
        {
            _ = gatewayTask.ContinueWith(
                completedTask =>
                {
                    if (completedTask.IsFaulted)
                    {
                        _ = completedTask.Exception;
                    }
                },
                CancellationToken.None,
                TaskContinuationOptions.ExecuteSynchronously,
                TaskScheduler.Default);
        }

        private sealed class ParsedReadCall
        {
            public ParsedReadCall(
                string toolName,
                object request,
                HmacCursorKind? cursorKind,
                string queryHash)
            {
                ToolName = toolName;
                Request = request;
                CursorKind = cursorKind;
                QueryHash = queryHash;
            }

            public string ToolName { get; }

            public object Request { get; }

            public HmacCursorKind? CursorKind { get; }

            public string QueryHash { get; }
        }

        private sealed class CursorValidationException : Exception
        {
        }
    }
}
