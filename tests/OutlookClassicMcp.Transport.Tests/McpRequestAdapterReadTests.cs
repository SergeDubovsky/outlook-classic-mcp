using System;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ModelContextProtocol.Protocol;
using NUnit.Framework;
using OutlookClassicMcp.Core.Outlook;
using OutlookClassicMcp.Core.Policy;

namespace OutlookClassicMcp.Transport.Tests
{
    [TestFixture]
    public sealed class McpRequestAdapterReadTests
    {
        private const string TokenText =
            "AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh8";
        private const string Fingerprint =
            "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
        private static readonly string[] LineSeparators = { "\r\n", "\n" };
        private static readonly string[] ExpectedToolOrder =
        {
            ToolNames.OutlookStatus,
            ToolNames.OutlookProbe,
            ToolNames.OutlookListMailboxes,
            ToolNames.OutlookListFolders,
            ToolNames.OutlookListMessages,
            ToolNames.OutlookSearchMessages,
            ToolNames.OutlookGetMessage,
            ToolNames.OutlookGetConversation,
            ToolNames.OutlookListAttachments,
        };

        [Test]
        public async Task Phase4ToolListFollowsTheExactPolicyOrder()
        {
            var gateway = new FakeOutlookGateway();
            using (var codec = CreateCodec())
            {
                var adapter = CreateAdapter(gateway, codec);
                using (var response = await InvokeAsync(
                    adapter,
                    "{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"tools/list\",\"params\":{}}"))
                {
                    var tools = response.RootElement.GetProperty("result").GetProperty("tools");
                    Assert.That(tools.GetArrayLength(), Is.EqualTo(ExpectedToolOrder.Length));
                    for (var index = 0; index < ExpectedToolOrder.Length; index++)
                    {
                        Assert.That(
                            tools[index].GetProperty("name").GetString(),
                            Is.EqualTo(ExpectedToolOrder[index]));
                    }
                }
            }
        }

        [Test]
        public async Task Phase4StatusIncludesBoundedReadDiagnostics()
        {
            var diagnostics = new OutlookReadDiagnosticsSnapshot(15, 12, 3, 5, 40);
            using (var codec = CreateCodec())
            {
                var adapter = new McpRequestAdapter(
                    () => new OutlookStatusSnapshot(
                        "Online",
                        listenerReady: true,
                        "1.0.0",
                        diagnostics),
                    new FakeOutlookGateway(),
                    codec);
                using (var response = await CallToolAsync(
                    adapter,
                    3,
                    ToolNames.OutlookStatus,
                    "{}"))
                {
                    var data = GetStructured(response).GetProperty("data");
                    var actual = data.GetProperty("readDiagnostics");
                    Assert.That(actual.GetProperty("comAcquired").GetInt64(), Is.EqualTo(15));
                    Assert.That(actual.GetProperty("comReleased").GetInt64(), Is.EqualTo(12));
                    Assert.That(actual.GetProperty("comOutstanding").GetInt64(), Is.EqualTo(3));
                    Assert.That(actual.GetProperty("comPeak").GetInt64(), Is.EqualTo(5));
                    Assert.That(
                        actual.GetProperty("materializedItemHighWater").GetInt64(),
                        Is.EqualTo(40));
                }
            }
        }

        [TestCase(
            ToolNames.OutlookListMailboxes,
            "{\"unknown\":true}",
            "INVALID_ARGUMENT")]
        [TestCase(
            ToolNames.OutlookListFolders,
            "{\"pageSize\":25}",
            "INVALID_ARGUMENT")]
        [TestCase(
            ToolNames.OutlookSearchMessages,
            "{\"scopes\":[{\"mailbox\":{\"storeId\":\"store-a\",\"extra\":1}}]}",
            "INVALID_ARGUMENT")]
        [TestCase(
            ToolNames.OutlookGetMessage,
            "{\"item\":{\"storeId\":\"store-a\",\"entryId\":\"item-a\",\"itemClass\":\"IPM.Note\"},\"bodyFormat\":\"rtf\"}",
            "INVALID_ARGUMENT")]
        [TestCase(
            ToolNames.OutlookListMessages,
            "{\"folder\":null}",
            "INVALID_ARGUMENT")]
        [TestCase(
            ToolNames.OutlookListMessages,
            "{\"folder\":{\"storeId\":\"store-a\",\"entryId\":\"inbox-a\"},\"cursor\":\"not-a-cursor\"}",
            "INVALID_ARGUMENT")]
        public async Task InvalidArgumentsFailClosedBeforeGatewayDispatch(
            string toolName,
            string arguments,
            string expectedCode)
        {
            var gateway = new FakeOutlookGateway();
            using (var codec = CreateCodec())
            {
                var adapter = CreateAdapter(gateway, codec);
                using (var response = await CallToolAsync(adapter, 2, toolName, arguments))
                {
                    var result = response.RootElement.GetProperty("result");
                    Assert.That(result.GetProperty("isError").GetBoolean(), Is.True);
                    Assert.That(
                        result.GetProperty("structuredContent")
                            .GetProperty("error")
                            .GetProperty("code")
                            .GetString(),
                        Is.EqualTo(expectedCode));
                }
            }

            Assert.That(gateway.ReadCallCount, Is.Zero);
        }

        [Test]
        public async Task AllSevenReadToolsDispatchWithClosedDefaultsAndTypedOutputs()
        {
            var gateway = CreateConfiguredGateway();
            using (var codec = CreateCodec())
            {
                var adapter = CreateAdapter(gateway, codec);
                using (var mailboxes = await CallToolAsync(
                    adapter,
                    10,
                    ToolNames.OutlookListMailboxes,
                    "{}"))
                using (var folders = await CallToolAsync(
                    adapter,
                    11,
                    ToolNames.OutlookListFolders,
                    "{\"mailbox\":{\"storeId\":\"store-a\"}}"))
                using (var messages = await CallToolAsync(
                    adapter,
                    12,
                    ToolNames.OutlookListMessages,
                    "{\"folder\":{\"storeId\":\"store-a\",\"entryId\":\"inbox-a\"}}"))
                using (var search = await CallToolAsync(
                    adapter,
                    13,
                    ToolNames.OutlookSearchMessages,
                    "{\"scopes\":[{\"mailbox\":{\"storeId\":\"store-a\"}}]}"))
                using (var detail = await CallToolAsync(
                    adapter,
                    14,
                    ToolNames.OutlookGetMessage,
                    "{\"item\":{\"storeId\":\"store-a\",\"entryId\":\"item-a\",\"itemClass\":\"IPM.Note\"}}"))
                using (var conversation = await CallToolAsync(
                    adapter,
                    15,
                    ToolNames.OutlookGetConversation,
                    "{\"item\":{\"storeId\":\"store-a\",\"entryId\":\"item-a\",\"itemClass\":\"IPM.Note\"}}"))
                using (var attachments = await CallToolAsync(
                    adapter,
                    16,
                    ToolNames.OutlookListAttachments,
                    "{\"item\":{\"storeId\":\"store-a\",\"entryId\":\"item-a\",\"itemClass\":\"IPM.Note\"}}"))
                {
                    AssertSuccessArray(mailboxes, "mailboxes", 1);
                    AssertSuccessArray(folders, "folders", 1);
                    AssertSuccessArray(messages, "messages", 1);
                    AssertSuccessArray(search, "messages", 1);
                    AssertSuccessArray(conversation, "messages", 1);
                    AssertSuccessArray(attachments, "attachments", 1);
                    var detailData = GetStructured(detail).GetProperty("data");
                    Assert.That(detailData.GetProperty("message").ValueKind, Is.EqualTo(JsonValueKind.Object));
                    Assert.That(detailData.GetProperty("totalToRecipientCount").GetInt32(), Is.EqualTo(1));
                }
            }

            Assert.That(gateway.ReadCallCount, Is.EqualTo(7));
            Assert.That(gateway.LastListMailboxesRequest!.PageSize,
                Is.EqualTo(RequestLimits.DefaultReadPageSize));
            Assert.That(gateway.LastGetMessageRequest!.BodyFormat,
                Is.EqualTo(OutlookBodyFormat.PlainText));
            Assert.That(gateway.LastGetMessageRequest.MaximumBodyCharacters,
                Is.EqualTo(OutlookReadLimits.MaximumBodyCharacters));
        }

        [Test]
        public async Task MessageCursorRoundTripsItemClassAndRejectsDifferentFolderWithoutDispatch()
        {
            var gateway = new FakeOutlookGateway();
            var call = 0;
            gateway.ListMessagesHandler = (request, _) =>
            {
                call++;
                var message = CreateMessage();
                return Task.FromResult(new OutlookMessagePage(
                    new[] { message },
                    call == 1
                        ? new OutlookMessageKeysetAnchor(
                            message.EffectiveTimestampUtc,
                            message.Item)
                        : null,
                    1,
                    Array.Empty<OutlookScopeFailure>()));
            };

            using (var codec = CreateCodec())
            {
                var adapter = CreateAdapter(gateway, codec);
                string cursor;
                using (var first = await CallToolAsync(
                    adapter,
                    20,
                    ToolNames.OutlookListMessages,
                    "{\"folder\":{\"storeId\":\"store-a\",\"entryId\":\"inbox-a\"}}"))
                {
                    cursor = GetStructured(first).GetProperty("data").GetProperty("nextCursor").GetString()!;
                }

                using (var second = await CallToolAsync(
                    adapter,
                    21,
                    ToolNames.OutlookListMessages,
                    "{\"folder\":{\"storeId\":\"store-a\",\"entryId\":\"inbox-a\"}," +
                    "\"cursor\":\"" + cursor + "\"}"))
                {
                    Assert.That(GetStructured(second).GetProperty("ok").GetBoolean(), Is.True);
                    Assert.That(gateway.LastListMessagesRequest!.Anchor!.Item.ItemClass,
                        Is.EqualTo("IPM.Note"));
                }

                using (var mismatch = await CallToolAsync(
                    adapter,
                    22,
                    ToolNames.OutlookListMessages,
                    "{\"folder\":{\"storeId\":\"store-a\",\"entryId\":\"other-folder\"}," +
                    "\"cursor\":\"" + cursor + "\"}"))
                {
                    Assert.That(
                        GetStructured(mismatch).GetProperty("error").GetProperty("code").GetString(),
                        Is.EqualTo("INVALID_ARGUMENT"));
                }
            }

            Assert.That(gateway.ListMessagesCallCount, Is.EqualTo(2));
        }

        [Test]
        public async Task ReacquisitionFailureForValidCursorReturnsCursorStale()
        {
            var gateway = new FakeOutlookGateway();
            var call = 0;
            gateway.ListMessagesHandler = (_, __) =>
            {
                call++;
                if (call > 1)
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
                }

                var message = CreateMessage();
                return Task.FromResult(new OutlookMessagePage(
                    new[] { message },
                    new OutlookMessageKeysetAnchor(
                        message.EffectiveTimestampUtc,
                        message.Item),
                    1,
                    Array.Empty<OutlookScopeFailure>()));
            };

            using (var codec = CreateCodec())
            {
                var adapter = CreateAdapter(gateway, codec);
                string cursor;
                using (var first = await CallToolAsync(
                    adapter,
                    23,
                    ToolNames.OutlookListMessages,
                    "{\"folder\":{\"storeId\":\"store-a\",\"entryId\":\"inbox-a\"}}"))
                {
                    cursor = GetStructured(first)
                        .GetProperty("data")
                        .GetProperty("nextCursor")
                        .GetString()!;
                }

                using (var second = await CallToolAsync(
                    adapter,
                    24,
                    ToolNames.OutlookListMessages,
                    "{\"folder\":{\"storeId\":\"store-a\",\"entryId\":\"inbox-a\"}," +
                    "\"cursor\":\"" + cursor + "\"}"))
                {
                    var error = GetStructured(second).GetProperty("error");
                    Assert.That(error.GetProperty("code").GetString(), Is.EqualTo("CURSOR_STALE"));
                    Assert.That(error.GetProperty("retryable").GetBoolean(), Is.False);
                }
            }

            Assert.That(gateway.ListMessagesCallCount, Is.EqualTo(2));
        }

        [Test]
        public async Task SearchCursorHashUsesCanonicalScopeOrdering()
        {
            var gateway = new FakeOutlookGateway();
            var call = 0;
            gateway.SearchMessagesHandler = (request, _) =>
            {
                call++;
                var message = CreateMessage();
                return Task.FromResult(new OutlookMessagePage(
                    new[] { message },
                    call == 1
                        ? new OutlookMessageKeysetAnchor(message.EffectiveTimestampUtc, message.Item)
                        : null,
                    2,
                    Array.Empty<OutlookScopeFailure>()));
            };

            using (var codec = CreateCodec())
            {
                var adapter = CreateAdapter(gateway, codec);
                string cursor;
                using (var first = await CallToolAsync(
                    adapter,
                    30,
                    ToolNames.OutlookSearchMessages,
                    "{\"scopes\":[" +
                    "{\"mailbox\":{\"storeId\":\"store-b\"}}," +
                    "{\"mailbox\":{\"storeId\":\"store-a\"}}]}"))
                {
                    cursor = GetStructured(first).GetProperty("data").GetProperty("nextCursor").GetString()!;
                }

                using (var second = await CallToolAsync(
                    adapter,
                    31,
                    ToolNames.OutlookSearchMessages,
                    "{\"scopes\":[" +
                    "{\"mailbox\":{\"storeId\":\"store-a\"}}," +
                    "{\"mailbox\":{\"storeId\":\"store-b\"}}]," +
                    "\"cursor\":\"" + cursor + "\"}"))
                {
                    Assert.That(GetStructured(second).GetProperty("ok").GetBoolean(), Is.True);
                }
            }

            Assert.That(gateway.SearchMessagesCallCount, Is.EqualTo(2));
        }

        [Test]
        public async Task FolderConversationAndAttachmentCursorsRoundTripTypedAnchors()
        {
            var folder = CreateFolder();
            var message = CreateMessage();
            var attachment = CreateAttachment(message.Item);
            var folderCalls = 0;
            var conversationCalls = 0;
            var attachmentCalls = 0;
            var gateway = new FakeOutlookGateway
            {
                ListFoldersHandler = (_, __) => Task.FromResult(new OutlookFolderPage(
                    new[] { folder },
                    ++folderCalls == 1
                        ? new OutlookFolderKeysetAnchor(folder.DisplayName, folder.Folder)
                        : null)),
                GetConversationHandler = (_, __) => Task.FromResult(new OutlookMessagePage(
                    new[] { message },
                    ++conversationCalls == 1
                        ? new OutlookMessageKeysetAnchor(
                            message.EffectiveTimestampUtc,
                            message.Item)
                        : null,
                    1,
                    Array.Empty<OutlookScopeFailure>())),
                ListAttachmentsHandler = (_, __) => Task.FromResult(new OutlookAttachmentPage(
                    new[] { attachment },
                    ++attachmentCalls == 1
                        ? new OutlookAttachmentKeysetAnchor(
                            attachment.Attachment.AttachmentIndex,
                            attachment.Attachment.MetadataFingerprint)
                        : null)),
            };

            using (var codec = CreateCodec())
            {
                var adapter = CreateAdapter(gateway, codec);
                var folderArguments = "{\"mailbox\":{\"storeId\":\"store-a\"}}";
                var itemArguments =
                    "{\"item\":{\"storeId\":\"store-a\",\"entryId\":\"item-a\"," +
                    "\"itemClass\":\"IPM.Note\"}}";

                var folderCursor = await ReadNextCursorAsync(
                    adapter,
                    32,
                    ToolNames.OutlookListFolders,
                    folderArguments);
                using (var replay = await CallToolAsync(
                    adapter,
                    33,
                    ToolNames.OutlookListFolders,
                    AddCursor(folderArguments, folderCursor)))
                {
                    Assert.That(GetStructured(replay).GetProperty("ok").GetBoolean(), Is.True);
                }

                var conversationCursor = await ReadNextCursorAsync(
                    adapter,
                    34,
                    ToolNames.OutlookGetConversation,
                    itemArguments);
                using (var replay = await CallToolAsync(
                    adapter,
                    35,
                    ToolNames.OutlookGetConversation,
                    AddCursor(itemArguments, conversationCursor)))
                {
                    Assert.That(GetStructured(replay).GetProperty("ok").GetBoolean(), Is.True);
                }

                var attachmentCursor = await ReadNextCursorAsync(
                    adapter,
                    36,
                    ToolNames.OutlookListAttachments,
                    itemArguments);
                using (var replay = await CallToolAsync(
                    adapter,
                    37,
                    ToolNames.OutlookListAttachments,
                    AddCursor(itemArguments, attachmentCursor)))
                {
                    Assert.That(GetStructured(replay).GetProperty("ok").GetBoolean(), Is.True);
                }
            }

            Assert.That(gateway.LastListFoldersRequest!.Anchor!.Folder.EntryId,
                Is.EqualTo(folder.Folder.EntryId));
            Assert.That(gateway.LastGetConversationRequest!.Anchor!.Item.ItemClass,
                Is.EqualTo(message.Item.ItemClass));
            Assert.That(gateway.LastListAttachmentsRequest!.Anchor!.AttachmentIndex,
                Is.EqualTo(attachment.Attachment.AttachmentIndex));
            Assert.That(gateway.LastListAttachmentsRequest.Anchor.MetadataFingerprint,
                Is.EqualTo(attachment.Attachment.MetadataFingerprint));
        }

        [Test]
        public async Task PartialSearchSuppressesGlobalCursorAndReturnsSafeScopeFailure()
        {
            var failedScope = new OutlookSearchScope(new MailboxRef("store-b"), null);
            var message = CreateMessage();
            var gateway = new FakeOutlookGateway
            {
                SearchMessagesHandler = (_, __) => Task.FromResult(new OutlookMessagePage(
                    new[] { message },
                    null,
                    2,
                    new[]
                    {
                        new OutlookScopeFailure(
                            failedScope,
                            OutlookGatewayFailure.AccessDenied),
                    })),
            };

            using (var codec = CreateCodec())
            {
                var adapter = CreateAdapter(gateway, codec);
                using (var response = await CallToolAsync(
                    adapter,
                    40,
                    ToolNames.OutlookSearchMessages,
                    "{\"scopes\":[" +
                    "{\"mailbox\":{\"storeId\":\"store-a\"}}," +
                    "{\"mailbox\":{\"storeId\":\"store-b\"}}]}"))
                {
                    var structured = GetStructured(response);
                    Assert.That(structured.GetProperty("partial").GetBoolean(), Is.True);
                    var data = structured.GetProperty("data");
                    Assert.That(data.GetProperty("nextCursor").ValueKind, Is.EqualTo(JsonValueKind.Null));
                    var failure = data.GetProperty("scopeFailures")[0];
                    Assert.That(failure.GetProperty("scopeIndex").GetInt32(), Is.EqualTo(1));
                    Assert.That(failure.GetProperty("code").GetString(), Is.EqualTo("ACCESS_DENIED"));
                    Assert.That(failure.GetProperty("retryable").GetBoolean(), Is.False);
                    Assert.That(failure.GetProperty("message").GetString(),
                        Is.EqualTo("Outlook denied access to the requested data."));
                }
            }
        }

        [Test]
        public async Task ProtectedOrTruncatedContentStaysOutOfTextAndMarksPartial()
        {
            const string sensitiveBody = "message-body-sentinel";
            var message = CreateMessage();
            var detail = new OutlookMessageDetail(
                message,
                new[] { new OutlookMessageAddress("Recipient", "recipient@example.test") },
                Array.Empty<OutlookMessageAddress>(),
                Array.Empty<OutlookMessageAddress>(),
                totalToRecipientCount: 2,
                totalCcRecipientCount: 0,
                totalBccRecipientCount: 0,
                new OutlookMessageBody(
                    OutlookBodyFormat.PlainText,
                    sensitiveBody,
                    sensitiveBody.Length + 10,
                    isTruncated: true,
                    isProtected: false));
            var gateway = new FakeOutlookGateway
            {
                GetMessageHandler = (_, __) => Task.FromResult(detail),
            };

            using (var codec = CreateCodec())
            {
                var adapter = CreateAdapter(gateway, codec);
                using (var response = await CallToolAsync(
                    adapter,
                    50,
                    ToolNames.OutlookGetMessage,
                    "{\"item\":{\"storeId\":\"store-a\",\"entryId\":\"item-a\",\"itemClass\":\"IPM.Note\"}}"))
                {
                    var result = response.RootElement.GetProperty("result");
                    var structured = result.GetProperty("structuredContent");
                    Assert.That(structured.GetProperty("partial").GetBoolean(), Is.True);
                    Assert.That(
                        structured.GetProperty("data").GetProperty("body").GetProperty("content").GetString(),
                        Is.EqualTo(sensitiveBody));
                    Assert.That(structured.GetProperty("warnings").GetArrayLength(), Is.EqualTo(2));
                    var text = result.GetProperty("content")[0].GetProperty("text").GetString();
                    Assert.That(text, Does.Not.Contain(sensitiveBody));
                    Assert.That(text, Does.Not.Contain("item-a"));
                    Assert.That(text, Does.Not.Contain("store-a"));
                }
            }
        }

        [Test]
        public async Task GatewayExceptionDetailsAreAbsentFromTheMcpResult()
        {
            const string sensitiveFailure =
                "message-body-sentinel store-id-sentinel entry-id-sentinel";
            var gateway = new FakeOutlookGateway
            {
                ListMailboxesHandler = (_, __) =>
                    throw new InvalidOperationException(sensitiveFailure),
            };

            using (var codec = CreateCodec())
            {
                var adapter = CreateAdapter(gateway, codec);
                using (var response = await CallToolAsync(
                    adapter,
                    54,
                    ToolNames.OutlookListMailboxes,
                    "{}"))
                {
                    var rawResult = response.RootElement.GetRawText();
                    Assert.That(rawResult, Does.Not.Contain(sensitiveFailure));
                    Assert.That(rawResult, Does.Not.Contain("store-id-sentinel"));
                    Assert.That(rawResult, Does.Not.Contain("entry-id-sentinel"));
                    var error = GetStructured(response).GetProperty("error");
                    Assert.That(error.GetProperty("code").GetString(), Is.EqualTo("INTERNAL"));
                    Assert.That(
                        error.GetProperty("message").GetString(),
                        Is.EqualTo("The Outlook read operation failed."));
                }
            }
        }

        [Test]
        public async Task OversizedPageIsTrimmedWithSafeContinuationCursor()
        {
            var oversizedPage = CreateOversizedMailboxPage();
            var call = 0;
            var gateway = new FakeOutlookGateway
            {
                ListMailboxesHandler = (_, __) =>
                {
                    call++;
                    return Task.FromResult(call == 1
                        ? oversizedPage
                        : new OutlookMailboxPage(
                            Array.Empty<OutlookMailboxSummary>(),
                            null));
                },
            };

            using (var codec = CreateCodec())
            {
                var adapter = CreateAdapter(gateway, codec);
                string cursor;
                string lastDisplayName;
                string lastStoreId;
                using (var first = await CallToolAsync(
                    adapter,
                    55,
                    ToolNames.OutlookListMailboxes,
                    "{\"pageSize\":50}"))
                {
                    var structured = GetStructured(first);
                    Assert.That(structured.GetProperty("ok").GetBoolean(), Is.True);
                    Assert.That(structured.GetProperty("partial").GetBoolean(), Is.True);
                    Assert.That(
                        Encoding.UTF8.GetByteCount(structured.GetRawText()),
                        Is.LessThanOrEqualTo(RequestLimits.MaximumToolResultBytes));
                    Assert.That(
                        structured.GetProperty("warnings")[0].GetString(),
                        Is.EqualTo(OutlookReadToolHandler.ResultTruncatedWarning));

                    var data = structured.GetProperty("data");
                    Assert.That(data.GetProperty("resultTruncated").GetBoolean(), Is.True);
                    var mailboxes = data.GetProperty("mailboxes");
                    Assert.That(mailboxes.GetArrayLength(), Is.GreaterThan(0));
                    Assert.That(
                        mailboxes.GetArrayLength(),
                        Is.LessThan(OutlookReadLimits.MaximumPageSize));
                    var last = mailboxes[mailboxes.GetArrayLength() - 1];
                    lastDisplayName = last.GetProperty("displayName").GetString()!;
                    lastStoreId = last.GetProperty("mailbox").GetProperty("storeId").GetString()!;
                    cursor = data.GetProperty("nextCursor").GetString()!;
                }

                using (var second = await CallToolAsync(
                    adapter,
                    56,
                    ToolNames.OutlookListMailboxes,
                    "{\"pageSize\":50,\"cursor\":\"" + cursor + "\"}"))
                {
                    Assert.That(GetStructured(second).GetProperty("ok").GetBoolean(), Is.True);
                }

                Assert.That(gateway.LastListMailboxesRequest!.Anchor, Is.Not.Null);
                Assert.That(
                    gateway.LastListMailboxesRequest.Anchor!.DisplayName,
                    Is.EqualTo(lastDisplayName));
                Assert.That(
                    gateway.LastListMailboxesRequest.Anchor.Mailbox.StoreId,
                    Is.EqualTo(lastStoreId));
            }
        }

        [Test]
        public async Task OversizedPartialSearchRetainsFailuresAndSuppressesCursor()
        {
            var gateway = new FakeOutlookGateway
            {
                SearchMessagesHandler = (_, __) =>
                    Task.FromResult(CreateOversizedPartialMessagePage()),
            };

            using (var codec = CreateCodec())
            {
                var adapter = CreateAdapter(gateway, codec);
                using (var response = await CallToolAsync(
                    adapter,
                    57,
                    ToolNames.OutlookSearchMessages,
                    "{\"scopes\":[" +
                    "{\"mailbox\":{\"storeId\":\"store-a\"}}," +
                    "{\"mailbox\":{\"storeId\":\"store-b\"}}]}"))
                {
                    var callToolResult = response.RootElement.GetProperty("result");
                    Assert.That(
                        Encoding.UTF8.GetByteCount(callToolResult.GetRawText()),
                        Is.LessThanOrEqualTo(RequestLimits.MaximumToolResultBytes));
                    var structured = callToolResult.GetProperty("structuredContent");
                    Assert.That(structured.GetProperty("ok").GetBoolean(), Is.True);
                    Assert.That(structured.GetProperty("partial").GetBoolean(), Is.True);
                    Assert.That(
                        structured.GetProperty("warnings")[0].GetString(),
                        Is.EqualTo(OutlookReadToolHandler.ResultTruncatedWarning));
                    var data = structured.GetProperty("data");
                    Assert.That(data.GetProperty("resultTruncated").GetBoolean(), Is.True);
                    Assert.That(data.GetProperty("nextCursor").ValueKind,
                        Is.EqualTo(JsonValueKind.Null));
                    Assert.That(data.GetProperty("messages").GetArrayLength(), Is.GreaterThan(0));
                    Assert.That(
                        data.GetProperty("messages").GetArrayLength(),
                        Is.LessThan(OutlookReadLimits.MaximumPageSize));
                    Assert.That(data.GetProperty("scopeFailures").GetArrayLength(), Is.EqualTo(1));
                    var failure = data.GetProperty("scopeFailures")[0];
                    Assert.That(failure.GetProperty("scopeIndex").GetInt32(), Is.EqualTo(1));
                    Assert.That(failure.GetProperty("code").GetString(), Is.EqualTo("ACCESS_DENIED"));
                }
            }
        }

        [Test]
        public async Task ReadDeadlineReturnsTypedTimeoutAndCancelsGatewayToken()
        {
            var pending = new TaskCompletionSource<OutlookMailboxPage>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            var calls = 0;
            var timedOutToken = CancellationToken.None;
            var gateway = new FakeOutlookGateway
            {
                ListMailboxesHandler = (_, __) => ++calls == 1
                    ? pending.Task
                    : Task.FromResult(new OutlookMailboxPage(
                        Array.Empty<OutlookMailboxSummary>(),
                        null)),
            };

            using (var codec = CreateCodec())
            {
                var adapter = new McpRequestAdapter(
                    () => new OutlookStatusSnapshot("Online", true, "1.0.0"),
                    gateway,
                    codec,
                    TimeSpan.FromMilliseconds(25));
                using (var response = await CallToolAsync(
                    adapter,
                    60,
                    ToolNames.OutlookListMailboxes,
                    "{}"))
                {
                    var error = GetStructured(response).GetProperty("error");
                    Assert.That(error.GetProperty("code").GetString(), Is.EqualTo("TIMEOUT"));
                    Assert.That(error.GetProperty("retryable").GetBoolean(), Is.True);
                    timedOutToken = gateway.LastCancellationToken;
                }

                pending.SetException(new InvalidOperationException("late failure sentinel"));
                using (var recovered = await CallToolAsync(
                    adapter,
                    61,
                    ToolNames.OutlookListMailboxes,
                    "{}"))
                {
                    Assert.That(GetStructured(recovered).GetProperty("ok").GetBoolean(), Is.True);
                }
            }

            Assert.That(timedOutToken.IsCancellationRequested, Is.True);
            Assert.That(gateway.ListMailboxesCallCount, Is.EqualTo(2));
        }

        private static FakeOutlookGateway CreateConfiguredGateway()
        {
            var message = CreateMessage();
            var gateway = new FakeOutlookGateway
            {
                ListMailboxesHandler = (_, __) => Task.FromResult(new OutlookMailboxPage(
                    new[] { CreateMailbox() },
                    null)),
                ListFoldersHandler = (_, __) => Task.FromResult(new OutlookFolderPage(
                    new[] { CreateFolder() },
                    null)),
                ListMessagesHandler = (_, __) => Task.FromResult(CreateMessagePage(message)),
                SearchMessagesHandler = (_, __) => Task.FromResult(CreateMessagePage(message)),
                GetMessageHandler = (_, __) => Task.FromResult(new OutlookMessageDetail(
                    message,
                    new[] { new OutlookMessageAddress("Recipient", "recipient@example.test") },
                    Array.Empty<OutlookMessageAddress>(),
                    Array.Empty<OutlookMessageAddress>(),
                    1,
                    0,
                    0,
                    new OutlookMessageBody(
                        OutlookBodyFormat.PlainText,
                        "body",
                        4,
                        false,
                        false))),
                GetConversationHandler = (_, __) => Task.FromResult(CreateMessagePage(message)),
                ListAttachmentsHandler = (_, __) => Task.FromResult(new OutlookAttachmentPage(
                    new[] { CreateAttachment(message.Item) },
                    null)),
            };
            return gateway;
        }

        private static OutlookMailboxSummary CreateMailbox()
        {
            var mailbox = new MailboxRef("store-a");
            return new OutlookMailboxSummary(
                mailbox,
                "Mailbox A",
                OutlookStoreType.PrimaryExchangeMailbox,
                new OutlookStoreCapabilities(true, false, true),
                new OutlookStandardFolderReferences(
                    new FolderRef("store-a", "inbox-a"),
                    null,
                    null,
                    null,
                    null));
        }

        private static OutlookMailboxPage CreateOversizedMailboxPage()
        {
            var storeId = "s" + new string('x', MailboxRef.MaximumStoreIdLength - 1);
            var entryId = "f" + new string('y', FolderRef.MaximumEntryIdLength - 1);
            var mailbox = new MailboxRef(storeId);
            var folder = new FolderRef(storeId, entryId);
            var standardFolders = new OutlookStandardFolderReferences(
                folder,
                folder,
                folder,
                folder,
                folder);
            var items = new OutlookMailboxSummary[OutlookReadLimits.MaximumPageSize];
            for (var index = 0; index < items.Length; index++)
            {
                items[index] = new OutlookMailboxSummary(
                    mailbox,
                    "Mailbox " + index.ToString("D2", System.Globalization.CultureInfo.InvariantCulture),
                    OutlookStoreType.PrimaryExchangeMailbox,
                    new OutlookStoreCapabilities(true, false, true),
                    standardFolders);
            }

            return new OutlookMailboxPage(items, null);
        }

        private static OutlookMessagePage CreateOversizedPartialMessagePage()
        {
            const string storeId = "store-a";
            var folderEntryId = new string('\u0100', FolderRef.MaximumEntryIdLength);
            var itemEntryId = new string('\u0101', ItemRef.MaximumEntryIdLength);
            var message = new OutlookMessageSummary(
                new ItemRef(
                    storeId,
                    itemEntryId,
                    new string('\u0102', ItemRef.MaximumItemClassLength)),
                new FolderRef(storeId, folderEntryId),
                new string('\u0103', OutlookMessageSummary.MaximumSubjectLength),
                new string('\u0104', OutlookMessageSummary.MaximumSenderLength),
                new string('\u0105', OutlookMessageSummary.MaximumSenderLength),
                new DateTime(2026, 7, 19, 12, 0, 0, DateTimeKind.Utc),
                null,
                null,
                isRead: false,
                attachmentCount: 0,
                new string('\u0106', OutlookMessageSummary.MaximumConversationIdLength));
            var items = new OutlookMessageSummary[OutlookReadLimits.MaximumPageSize];
            for (var index = 0; index < items.Length; index++)
            {
                items[index] = message;
            }

            var failedScope = new OutlookSearchScope(new MailboxRef("store-b"), null);
            return new OutlookMessagePage(
                items,
                null,
                2,
                new[]
                {
                    new OutlookScopeFailure(
                        failedScope,
                        OutlookGatewayFailure.AccessDenied),
                });
        }

        private static OutlookFolderSummary CreateFolder()
        {
            return new OutlookFolderSummary(
                new FolderRef("store-a", "inbox-a"),
                null,
                "Inbox",
                hasChildren: false);
        }

        private static OutlookMessageSummary CreateMessage()
        {
            return new OutlookMessageSummary(
                new ItemRef("store-a", "item-a", "IPM.Note"),
                new FolderRef("store-a", "inbox-a"),
                "Subject",
                "Sender",
                "sender@example.test",
                new DateTime(2026, 7, 19, 12, 0, 0, DateTimeKind.Utc),
                new DateTime(2026, 7, 19, 12, 0, 0, DateTimeKind.Utc),
                null,
                isRead: false,
                attachmentCount: 1,
                conversationId: "conversation-a");
        }

        private static OutlookMessagePage CreateMessagePage(OutlookMessageSummary message)
        {
            return new OutlookMessagePage(
                new[] { message },
                null,
                1,
                Array.Empty<OutlookScopeFailure>());
        }

        private static OutlookAttachmentSummary CreateAttachment(ItemRef item)
        {
            return new OutlookAttachmentSummary(
                new AttachmentRef(
                    item,
                    1,
                    "file.txt",
                    12,
                    sizeIsKnown: true,
                    Fingerprint),
                "text/plain");
        }

        private static McpRequestAdapter CreateAdapter(
            IOutlookGateway gateway,
            HmacCursorCodec codec)
        {
            return new McpRequestAdapter(
                () => new OutlookStatusSnapshot("Online", true, "1.0.0"),
                gateway,
                codec);
        }

        private static HmacCursorCodec CreateCodec()
        {
            Assert.That(BearerToken.TryCreate(TokenText, out var token), Is.True);
            using (token)
            {
                return token.CreateCursorCodec();
            }
        }

        private static async Task<JsonDocument> CallToolAsync(
            McpRequestAdapter adapter,
            int id,
            string toolName,
            string arguments)
        {
            return await InvokeAsync(
                adapter,
                "{\"jsonrpc\":\"2.0\",\"id\":" + id +
                ",\"method\":\"tools/call\",\"params\":{\"name\":\"" +
                toolName + "\",\"arguments\":" + arguments + "}}");
        }

        private static async Task<string> ReadNextCursorAsync(
            McpRequestAdapter adapter,
            int id,
            string toolName,
            string arguments)
        {
            using (var response = await CallToolAsync(adapter, id, toolName, arguments))
            {
                return GetStructured(response)
                    .GetProperty("data")
                    .GetProperty("nextCursor")
                    .GetString()!;
            }
        }

        private static string AddCursor(string arguments, string cursor)
        {
            return arguments.Substring(0, arguments.Length - 1) +
                ",\"cursor\":\"" + cursor + "\"}";
        }

        private static async Task<JsonDocument> InvokeAsync(
            McpRequestAdapter adapter,
            string requestJson)
        {
            using (var body = new MemoryStream(Encoding.UTF8.GetBytes(requestJson), writable: false))
            {
                var read = await McpRequestAdapter.ReadSingleMessageAsync(
                    body,
                    McpRequestAdapter.MaximumSupportedBodyBytes,
                    CancellationToken.None);
                Assert.That(read.Succeeded, Is.True);
                using (var response = new MemoryStream())
                {
                    Assert.That(await adapter.HandleAsync(
                        read.Message!,
                        response,
                        CancellationToken.None), Is.True);
                    var payload = Encoding.UTF8.GetString(response.ToArray());
                    var lines = payload.Split(LineSeparators, StringSplitOptions.None);
                    for (var index = 0; index < lines.Length; index++)
                    {
                        if (lines[index].StartsWith("data: ", StringComparison.Ordinal))
                        {
                            return JsonDocument.Parse(lines[index].Substring("data: ".Length));
                        }
                    }

                    Assert.Fail("The MCP response did not contain an SSE data record.");
                    throw new InvalidOperationException();
                }
            }
        }

        private static JsonElement GetStructured(JsonDocument response)
        {
            return response.RootElement
                .GetProperty("result")
                .GetProperty("structuredContent");
        }

        private static void AssertSuccessArray(
            JsonDocument response,
            string propertyName,
            int expectedCount)
        {
            var structured = GetStructured(response);
            Assert.That(structured.GetProperty("ok").GetBoolean(), Is.True);
            Assert.That(
                structured.GetProperty("data").GetProperty(propertyName).GetArrayLength(),
                Is.EqualTo(expectedCount));
        }
    }
}
