using System;
using System.Text.Json;
using System.Text.Json.Nodes;
using ModelContextProtocol.Protocol;
using OutlookClassicMcp.Core.Outlook;
using OutlookClassicMcp.Core.Policy;

namespace OutlookClassicMcp.Transport
{
    /// <summary>
    /// Creates the closed protocol descriptors for Phase 4 bounded reads.
    /// </summary>
    public static class OutlookReadCatalog
    {
        public static Tool CreateDescriptor(string toolName)
        {
            switch (toolName)
            {
                case ToolNames.OutlookListMailboxes:
                    return CreateTool(
                        toolName,
                        "List Outlook mailboxes",
                        "Lists bounded Outlook mailbox metadata and store-qualified standard-folder references.",
                        CreateListMailboxesInputSchema(),
                        CreatePageOutputSchema(
                            "mailboxes",
                            CreateMailboxSummarySchema(),
                            includeScopeFailures: false));
                case ToolNames.OutlookListFolders:
                    return CreateTool(
                        toolName,
                        "List Outlook folders",
                        "Lists one bounded page of folders in an explicitly selected Outlook mailbox.",
                        CreateListFoldersInputSchema(),
                        CreatePageOutputSchema(
                            "folders",
                            CreateFolderSummarySchema(),
                            includeScopeFailures: false));
                case ToolNames.OutlookListMessages:
                    return CreateTool(
                        toolName,
                        "List Outlook messages",
                        "Lists bounded message metadata in an explicitly selected Outlook folder without returning bodies.",
                        CreateListMessagesInputSchema(),
                        CreatePageOutputSchema(
                            "messages",
                            CreateMessageSummarySchema(),
                            includeScopeFailures: true));
                case ToolNames.OutlookSearchMessages:
                    return CreateTool(
                        toolName,
                        "Search Outlook messages",
                        "Searches explicit Outlook scopes with closed structured filters and returns bounded message metadata.",
                        CreateSearchMessagesInputSchema(),
                        CreatePageOutputSchema(
                            "messages",
                            CreateMessageSummarySchema(),
                            includeScopeFailures: true));
                case ToolNames.OutlookGetMessage:
                    return CreateTool(
                        toolName,
                        "Get Outlook message",
                        "Retrieves one store-qualified message with bounded recipients and protected or truncated body metadata.",
                        CreateGetMessageInputSchema(),
                        CreateMessageDetailOutputSchema());
                case ToolNames.OutlookGetConversation:
                    return CreateTool(
                        toolName,
                        "Get Outlook conversation",
                        "Lists bounded message metadata for the conversation containing an explicitly selected message.",
                        CreateGetConversationInputSchema(),
                        CreatePageOutputSchema(
                            "messages",
                            CreateMessageSummarySchema(),
                            includeScopeFailures: true));
                case ToolNames.OutlookListAttachments:
                    return CreateTool(
                        toolName,
                        "List Outlook attachments",
                        "Lists bounded attachment metadata for one store-qualified message without exporting attachment content.",
                        CreateListAttachmentsInputSchema(),
                        CreatePageOutputSchema(
                            "attachments",
                            CreateAttachmentSummarySchema(),
                            includeScopeFailures: false));
                default:
                    throw new ArgumentOutOfRangeException(nameof(toolName));
            }
        }

        public static bool IsReadTool(string? toolName)
        {
            switch (toolName)
            {
                case ToolNames.OutlookListMailboxes:
                case ToolNames.OutlookListFolders:
                case ToolNames.OutlookListMessages:
                case ToolNames.OutlookSearchMessages:
                case ToolNames.OutlookGetMessage:
                case ToolNames.OutlookGetConversation:
                case ToolNames.OutlookListAttachments:
                    return true;
                default:
                    return false;
            }
        }

        private static Tool CreateTool(
            string name,
            string title,
            string description,
            JsonObject inputSchema,
            JsonObject outputSchema)
        {
            return new Tool
            {
                Name = name,
                Title = title,
                Description = description,
                InputSchema = JsonSerializer.SerializeToElement(inputSchema),
                OutputSchema = JsonSerializer.SerializeToElement(outputSchema),
                Annotations = new ToolAnnotations
                {
                    ReadOnlyHint = true,
                    DestructiveHint = false,
                    IdempotentHint = true,
                    OpenWorldHint = false,
                },
            };
        }

        private static JsonObject CreateListMailboxesInputSchema()
        {
            return CreateClosedObjectSchema(new JsonObject
            {
                ["pageSize"] = CreatePageSizeSchema(),
                ["cursor"] = CreateCursorSchema(),
            });
        }

        private static JsonObject CreateListFoldersInputSchema()
        {
            return CreateClosedObjectSchema(
                new JsonObject
                {
                    ["mailbox"] = CreateMailboxRefSchema(),
                    ["parentFolder"] = CreateNullableSchema(CreateFolderRefSchema()),
                    ["pageSize"] = CreatePageSizeSchema(),
                    ["cursor"] = CreateCursorSchema(),
                },
                "mailbox");
        }

        private static JsonObject CreateListMessagesInputSchema()
        {
            return CreateClosedObjectSchema(
                new JsonObject
                {
                    ["folder"] = CreateFolderRefSchema(),
                    ["pageSize"] = CreatePageSizeSchema(),
                    ["cursor"] = CreateCursorSchema(),
                },
                "folder");
        }

        private static JsonObject CreateSearchMessagesInputSchema()
        {
            var scope = CreateClosedObjectSchema(
                new JsonObject
                {
                    ["mailbox"] = CreateMailboxRefSchema(),
                    ["folder"] = CreateNullableSchema(CreateFolderRefSchema()),
                },
                "mailbox");
            var filter = CreateClosedObjectSchema(new JsonObject
            {
                ["sender"] = CreateBoundedStringSchema(OutlookMessageSearchFilter.MaximumAddressLength),
                ["recipient"] = CreateBoundedStringSchema(OutlookMessageSearchFilter.MaximumAddressLength),
                ["subject"] = CreateBoundedStringSchema(OutlookMessageSearchFilter.MaximumSubjectLength),
                ["text"] = CreateBoundedStringSchema(OutlookMessageSearchFilter.MaximumTextLength),
                ["receivedFromUtc"] = CreateUtcTimestampSchema(),
                ["receivedToUtc"] = CreateUtcTimestampSchema(),
                ["isUnread"] = new JsonObject { ["type"] = "boolean" },
                ["category"] = CreateBoundedStringSchema(OutlookMessageSearchFilter.MaximumCategoryLength),
                ["hasAttachments"] = new JsonObject { ["type"] = "boolean" },
            });

            return CreateClosedObjectSchema(
                new JsonObject
                {
                    ["scopes"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["minItems"] = 1,
                        ["maxItems"] = OutlookReadLimits.MaximumSearchScopeCount,
                        ["items"] = scope,
                    },
                    ["filter"] = filter,
                    ["pageSize"] = CreatePageSizeSchema(),
                    ["cursor"] = CreateCursorSchema(),
                },
                "scopes");
        }

        private static JsonObject CreateGetMessageInputSchema()
        {
            return CreateClosedObjectSchema(
                new JsonObject
                {
                    ["item"] = CreateItemRefSchema(),
                    ["bodyFormat"] = CreateBodyFormatSchema(),
                    ["maximumBodyCharacters"] = new JsonObject
                    {
                        ["type"] = "integer",
                        ["minimum"] = 1,
                        ["maximum"] = OutlookReadLimits.MaximumBodyCharacters,
                        ["default"] = RequestLimits.DefaultMessageBodyCharacters,
                    },
                },
                "item");
        }

        private static JsonObject CreateGetConversationInputSchema()
        {
            return CreateClosedObjectSchema(
                new JsonObject
                {
                    ["item"] = CreateItemRefSchema(),
                    ["pageSize"] = CreatePageSizeSchema(),
                    ["cursor"] = CreateCursorSchema(),
                },
                "item");
        }

        private static JsonObject CreateListAttachmentsInputSchema()
        {
            return CreateGetConversationInputSchema();
        }

        private static JsonObject CreatePageOutputSchema(
            string arrayPropertyName,
            JsonObject itemSchema,
            bool includeScopeFailures)
        {
            var dataProperties = new JsonObject
            {
                [arrayPropertyName] = new JsonObject
                {
                    ["type"] = "array",
                    ["maxItems"] = OutlookReadLimits.MaximumPageSize,
                    ["items"] = itemSchema,
                },
                ["nextCursor"] = CreateNullableSchema(CreateCursorSchema()),
                ["resultTruncated"] = new JsonObject { ["type"] = "boolean" },
            };
            var required = includeScopeFailures
                ? new[]
                {
                    arrayPropertyName,
                    "nextCursor",
                    "resultTruncated",
                    "totalScopeCount",
                    "scopeFailures",
                }
                : new[] { arrayPropertyName, "nextCursor", "resultTruncated" };
            if (includeScopeFailures)
            {
                dataProperties["totalScopeCount"] = new JsonObject
                {
                    ["type"] = "integer",
                    ["minimum"] = 1,
                    ["maximum"] = OutlookReadLimits.MaximumSearchScopeCount,
                };
                dataProperties["scopeFailures"] = new JsonObject
                {
                    ["type"] = "array",
                    ["maxItems"] = OutlookReadLimits.MaximumSearchScopeCount - 1,
                    ["items"] = CreateScopeFailureSchema(),
                };
            }

            return CreateEnvelopeSchema(
                CreateClosedObjectSchema(dataProperties, required),
                partialMayBeTrue: true);
        }

        private static JsonObject CreateMessageDetailOutputSchema()
        {
            var body = CreateClosedObjectSchema(
                new JsonObject
                {
                    ["format"] = CreateStringEnum("plainText", "html"),
                    ["content"] = new JsonObject
                    {
                        ["type"] = "string",
                        ["maxLength"] = OutlookReadLimits.MaximumBodyCharacters,
                    },
                    ["originalCharacterCount"] = CreateNullableSchema(new JsonObject
                    {
                        ["type"] = "integer",
                        ["minimum"] = 0,
                    }),
                    ["isTruncated"] = new JsonObject { ["type"] = "boolean" },
                    ["isProtected"] = new JsonObject { ["type"] = "boolean" },
                },
                "format",
                "content",
                "originalCharacterCount",
                "isTruncated",
                "isProtected");
            var data = CreateClosedObjectSchema(
                new JsonObject
                {
                    ["message"] = CreateMessageSummarySchema(),
                    ["toRecipients"] = CreateArraySchema(
                        CreateAddressSchema(),
                        OutlookMessageDetail.MaximumRecipientCount),
                    ["ccRecipients"] = CreateArraySchema(
                        CreateAddressSchema(),
                        OutlookMessageDetail.MaximumRecipientCount),
                    ["bccRecipients"] = CreateArraySchema(
                        CreateAddressSchema(),
                        OutlookMessageDetail.MaximumRecipientCount),
                    ["totalToRecipientCount"] = CreateNonNegativeIntegerSchema(),
                    ["totalCcRecipientCount"] = CreateNonNegativeIntegerSchema(),
                    ["totalBccRecipientCount"] = CreateNonNegativeIntegerSchema(),
                    ["toRecipientsTruncated"] = new JsonObject { ["type"] = "boolean" },
                    ["ccRecipientsTruncated"] = new JsonObject { ["type"] = "boolean" },
                    ["bccRecipientsTruncated"] = new JsonObject { ["type"] = "boolean" },
                    ["body"] = body,
                },
                "message",
                "toRecipients",
                "ccRecipients",
                "bccRecipients",
                "totalToRecipientCount",
                "totalCcRecipientCount",
                "totalBccRecipientCount",
                "toRecipientsTruncated",
                "ccRecipientsTruncated",
                "bccRecipientsTruncated",
                "body");

            return CreateEnvelopeSchema(data, partialMayBeTrue: true);
        }

        private static JsonObject CreateEnvelopeSchema(JsonObject dataSchema, bool partialMayBeTrue)
        {
            var success = CreateClosedObjectSchema(
                new JsonObject
                {
                    ["ok"] = new JsonObject { ["const"] = true },
                    ["operationId"] = CreateOperationIdSchema(),
                    ["data"] = dataSchema,
                    ["partial"] = partialMayBeTrue
                        ? new JsonObject { ["type"] = "boolean" }
                        : new JsonObject { ["const"] = false },
                    ["warnings"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["maxItems"] = 3,
                        ["items"] = CreateStringEnum(
                            OutlookReadToolHandler.BodyProtectedWarning,
                            OutlookReadToolHandler.BodyTruncatedWarning,
                            OutlookReadToolHandler.RecipientsTruncatedWarning,
                            OutlookReadToolHandler.ResultTruncatedWarning),
                    },
                },
                "ok",
                "operationId",
                "data",
                "partial",
                "warnings");
            var error = CreateClosedObjectSchema(
                new JsonObject
                {
                    ["code"] = CreateStringEnum(
                        "INVALID_ARGUMENT",
                        "OUTLOOK_NOT_READY",
                        "HOST_DEGRADED",
                        "HOST_PAUSED",
                        "HOST_STOPPING",
                        "STORE_NOT_FOUND",
                        "FOLDER_NOT_FOUND",
                        "ITEM_NOT_FOUND",
                        "ITEM_MOVED_OR_DELETED",
                        "UNSUPPORTED_STORE",
                        "UNSUPPORTED_ITEM_TYPE",
                        "ACCESS_DENIED",
                        "OBJECT_MODEL_GUARD",
                        "QUEUE_FULL",
                        "TIMEOUT",
                        "COM_BUSY",
                        "STA_DISPATCH_FAILED",
                        "CURSOR_STALE",
                        "INTERNAL"),
                    ["message"] = CreateBoundedStringSchema(256),
                    ["retryable"] = new JsonObject { ["type"] = "boolean" },
                    ["details"] = CreateClosedObjectSchema(new JsonObject()),
                },
                "code",
                "message",
                "retryable",
                "details");
            var failure = CreateClosedObjectSchema(
                new JsonObject
                {
                    ["ok"] = new JsonObject { ["const"] = false },
                    ["operationId"] = CreateOperationIdSchema(),
                    ["error"] = error,
                },
                "ok",
                "operationId",
                "error");

            return new JsonObject
            {
                ["type"] = "object",
                ["oneOf"] = new JsonArray(success, failure),
            };
        }

        private static JsonObject CreateMailboxSummarySchema()
        {
            var standardFolders = CreateClosedObjectSchema(
                new JsonObject
                {
                    ["inbox"] = CreateNullableSchema(CreateFolderRefSchema()),
                    ["drafts"] = CreateNullableSchema(CreateFolderRefSchema()),
                    ["sent"] = CreateNullableSchema(CreateFolderRefSchema()),
                    ["deleted"] = CreateNullableSchema(CreateFolderRefSchema()),
                    ["archive"] = CreateNullableSchema(CreateFolderRefSchema()),
                },
                "inbox",
                "drafts",
                "sent",
                "deleted",
                "archive");
            return CreateClosedObjectSchema(
                new JsonObject
                {
                    ["mailbox"] = CreateMailboxRefSchema(),
                    ["displayName"] = CreateBoundedStringSchema(
                        OutlookMailboxSummary.MaximumDisplayNameLength,
                        minimumLength: 0),
                    ["storeType"] = CreateStringEnum(
                        "primaryExchangeMailbox",
                        "exchangeMailbox",
                        "exchangePublicFolder",
                        "additionalExchangeMailbox",
                        "nonExchange",
                        "unknown"),
                    ["capabilities"] = CreateCapabilitiesSchema(),
                    ["standardFolders"] = standardFolders,
                },
                "mailbox",
                "displayName",
                "storeType",
                "capabilities",
                "standardFolders");
        }

        private static JsonObject CreateFolderSummarySchema()
        {
            return CreateClosedObjectSchema(
                new JsonObject
                {
                    ["folder"] = CreateFolderRefSchema(),
                    ["parentFolder"] = CreateNullableSchema(CreateFolderRefSchema()),
                    ["displayName"] = CreateBoundedStringSchema(
                        OutlookFolderSummary.MaximumDisplayNameLength,
                        minimumLength: 0),
                    ["hasChildren"] = new JsonObject { ["type"] = "boolean" },
                },
                "folder",
                "parentFolder",
                "displayName",
                "hasChildren");
        }

        private static JsonObject CreateMessageSummarySchema()
        {
            return CreateClosedObjectSchema(
                new JsonObject
                {
                    ["item"] = CreateItemRefSchema(),
                    ["folder"] = CreateFolderRefSchema(),
                    ["subject"] = CreateBoundedStringSchema(
                        OutlookMessageSummary.MaximumSubjectLength,
                        minimumLength: 0),
                    ["senderDisplayName"] = CreateNullableSchema(
                        CreateBoundedStringSchema(
                            OutlookMessageSummary.MaximumSenderLength,
                            minimumLength: 0)),
                    ["senderAddress"] = CreateNullableSchema(
                        CreateBoundedStringSchema(OutlookMessageSummary.MaximumSenderLength)),
                    ["effectiveTimestampUtc"] = CreateUtcTimestampSchema(),
                    ["receivedUtc"] = CreateNullableSchema(CreateUtcTimestampSchema()),
                    ["sentUtc"] = CreateNullableSchema(CreateUtcTimestampSchema()),
                    ["isRead"] = new JsonObject { ["type"] = "boolean" },
                    ["attachmentCount"] = new JsonObject
                    {
                        ["type"] = "integer",
                        ["minimum"] = 0,
                    },
                    ["hasAttachments"] = new JsonObject { ["type"] = "boolean" },
                    ["conversationId"] = CreateNullableSchema(
                        CreateBoundedStringSchema(OutlookMessageSummary.MaximumConversationIdLength)),
                },
                "item",
                "folder",
                "subject",
                "senderDisplayName",
                "senderAddress",
                "effectiveTimestampUtc",
                "receivedUtc",
                "sentUtc",
                "isRead",
                "attachmentCount",
                "hasAttachments",
                "conversationId");
        }

        private static JsonObject CreateAttachmentSummarySchema()
        {
            var attachment = CreateClosedObjectSchema(
                new JsonObject
                {
                    ["item"] = CreateItemRefSchema(),
                    ["attachmentIndex"] = new JsonObject
                    {
                        ["type"] = "integer",
                        ["minimum"] = 1,
                    },
                    ["name"] = CreateBoundedStringSchema(
                        AttachmentRef.MaximumNameLength,
                        minimumLength: 0),
                    ["size"] = new JsonObject
                    {
                        ["type"] = "integer",
                        ["minimum"] = 0,
                    },
                    ["sizeIsKnown"] = new JsonObject { ["type"] = "boolean" },
                    ["metadataFingerprint"] = CreateSha256Schema(),
                },
                "item",
                "attachmentIndex",
                "name",
                "size",
                "sizeIsKnown",
                "metadataFingerprint");
            return CreateClosedObjectSchema(
                new JsonObject
                {
                    ["attachment"] = attachment,
                    ["contentType"] = CreateNullableSchema(
                        CreateBoundedStringSchema(OutlookAttachmentSummary.MaximumContentTypeLength)),
                },
                "attachment",
                "contentType");
        }

        private static JsonObject CreateScopeFailureSchema()
        {
            var scope = CreateClosedObjectSchema(
                new JsonObject
                {
                    ["mailbox"] = CreateMailboxRefSchema(),
                    ["folder"] = CreateNullableSchema(CreateFolderRefSchema()),
                },
                "mailbox",
                "folder");
            return CreateClosedObjectSchema(
                new JsonObject
                {
                    ["scopeIndex"] = new JsonObject
                    {
                        ["type"] = "integer",
                        ["minimum"] = 0,
                        ["maximum"] = OutlookReadLimits.MaximumSearchScopeCount - 1,
                    },
                    ["scope"] = scope,
                    ["code"] = CreateStringEnum(
                        "STORE_NOT_FOUND",
                        "FOLDER_NOT_FOUND",
                        "UNSUPPORTED_STORE",
                        "UNSUPPORTED_ITEM_TYPE",
                        "ACCESS_DENIED",
                        "OBJECT_MODEL_GUARD",
                        "TIMEOUT",
                        "COM_BUSY"),
                    ["message"] = CreateBoundedStringSchema(256),
                    ["retryable"] = new JsonObject { ["type"] = "boolean" },
                },
                "scopeIndex",
                "scope",
                "code",
                "message",
                "retryable");
        }

        private static JsonObject CreateMailboxRefSchema()
        {
            return CreateClosedObjectSchema(
                new JsonObject
                {
                    ["storeId"] = CreateBoundedStringSchema(MailboxRef.MaximumStoreIdLength),
                },
                "storeId");
        }

        private static JsonObject CreateFolderRefSchema()
        {
            return CreateClosedObjectSchema(
                new JsonObject
                {
                    ["storeId"] = CreateBoundedStringSchema(MailboxRef.MaximumStoreIdLength),
                    ["entryId"] = CreateBoundedStringSchema(FolderRef.MaximumEntryIdLength),
                },
                "storeId",
                "entryId");
        }

        private static JsonObject CreateItemRefSchema()
        {
            return CreateClosedObjectSchema(
                new JsonObject
                {
                    ["storeId"] = CreateBoundedStringSchema(MailboxRef.MaximumStoreIdLength),
                    ["entryId"] = CreateBoundedStringSchema(ItemRef.MaximumEntryIdLength),
                    ["itemClass"] = CreateBoundedStringSchema(ItemRef.MaximumItemClassLength),
                },
                "storeId",
                "entryId",
                "itemClass");
        }

        private static JsonObject CreateCapabilitiesSchema()
        {
            return CreateClosedObjectSchema(
                new JsonObject
                {
                    ["isExchangeStore"] = new JsonObject { ["type"] = "boolean" },
                    ["isDataFileStore"] = new JsonObject { ["type"] = "boolean" },
                    ["isCachedExchange"] = new JsonObject { ["type"] = "boolean" },
                },
                "isExchangeStore",
                "isDataFileStore",
                "isCachedExchange");
        }

        private static JsonObject CreateAddressSchema()
        {
            return CreateClosedObjectSchema(
                new JsonObject
                {
                    ["displayName"] = CreateNullableSchema(
                        CreateBoundedStringSchema(
                            OutlookMessageAddress.MaximumValueLength,
                            minimumLength: 0)),
                    ["address"] = CreateNullableSchema(
                        CreateBoundedStringSchema(OutlookMessageAddress.MaximumValueLength)),
                },
                "displayName",
                "address");
        }

        private static JsonObject CreateArraySchema(JsonObject itemSchema, int maximumItems)
        {
            return new JsonObject
            {
                ["type"] = "array",
                ["maxItems"] = maximumItems,
                ["items"] = itemSchema,
            };
        }

        private static JsonObject CreateNonNegativeIntegerSchema()
        {
            return new JsonObject
            {
                ["type"] = "integer",
                ["minimum"] = 0,
            };
        }

        private static JsonObject CreatePageSizeSchema()
        {
            return new JsonObject
            {
                ["type"] = "integer",
                ["minimum"] = OutlookReadLimits.MinimumPageSize,
                ["maximum"] = OutlookReadLimits.MaximumPageSize,
                ["default"] = RequestLimits.DefaultReadPageSize,
            };
        }

        private static JsonObject CreateBodyFormatSchema()
        {
            var schema = CreateStringEnum("plainText", "html");
            schema["default"] = "plainText";
            return schema;
        }

        private static JsonObject CreateCursorSchema()
        {
            return CreateBoundedStringSchema(HmacCursorCodec.MaximumCursorLength);
        }

        private static JsonObject CreateUtcTimestampSchema()
        {
            return new JsonObject
            {
                ["type"] = "string",
                ["format"] = "date-time",
                ["maxLength"] = 64,
            };
        }

        private static JsonObject CreateSha256Schema()
        {
            return new JsonObject
            {
                ["type"] = "string",
                ["pattern"] = "^[0-9a-f]{64}$",
            };
        }

        private static JsonObject CreateOperationIdSchema()
        {
            return new JsonObject
            {
                ["type"] = "string",
                ["pattern"] = "^[0-9a-f]{32}$",
            };
        }

        private static JsonObject CreateNullableSchema(JsonObject schema)
        {
            return new JsonObject
            {
                ["oneOf"] = new JsonArray(
                    schema,
                    new JsonObject { ["type"] = "null" }),
            };
        }

        private static JsonObject CreateBoundedStringSchema(
            int maximumLength,
            int minimumLength = 1)
        {
            return new JsonObject
            {
                ["type"] = "string",
                ["minLength"] = minimumLength,
                ["maxLength"] = maximumLength,
            };
        }

        private static JsonObject CreateStringEnum(params string[] values)
        {
            var allowedValues = new JsonArray();
            for (var index = 0; index < values.Length; index++)
            {
                allowedValues.Add(values[index]);
            }

            return new JsonObject
            {
                ["type"] = "string",
                ["enum"] = allowedValues,
            };
        }

        private static JsonObject CreateClosedObjectSchema(
            JsonObject properties,
            params string[] required)
        {
            var requiredProperties = new JsonArray();
            for (var index = 0; index < required.Length; index++)
            {
                requiredProperties.Add(required[index]);
            }

            return new JsonObject
            {
                ["type"] = "object",
                ["properties"] = properties,
                ["required"] = requiredProperties,
                ["additionalProperties"] = false,
            };
        }
    }
}
