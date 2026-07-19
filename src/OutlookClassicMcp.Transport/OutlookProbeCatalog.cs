using System.Text.Json;
using System.Text.Json.Nodes;
using ModelContextProtocol.Protocol;
using OutlookClassicMcp.Core.Outlook;
using OutlookClassicMcp.Core.Policy;

namespace OutlookClassicMcp.Transport
{
    /// <summary>
    /// Creates isolated protocol descriptors for the bounded Outlook probe.
    /// </summary>
    public static class OutlookProbeCatalog
    {
        internal const string ArchiveAvailabilityWarning =
            "Archive availability is not exposed by the Outlook Object Model.";
        internal const string StoreMetadataIncompleteWarning =
            "Some store metadata could not be determined.";
        internal const string StoreLimitReachedWarning =
            "Store results were limited to the supported maximum.";

        private static readonly JsonElement InputSchema = ParseSchema(
            "{\"type\":\"object\",\"properties\":{},\"additionalProperties\":false}");

        private static readonly JsonElement OutputSchema = CreateOutputSchema();

        public static Tool CreateDescriptor()
        {
            return new Tool
            {
                Name = ToolNames.OutlookProbe,
                Title = "Outlook probe",
                Description = "Reports bounded Outlook, verified UI STA dispatcher, store capability, and standard-folder availability metadata without reading messages or attachments.",
                InputSchema = InputSchema,
                OutputSchema = OutputSchema,
                Annotations = new ToolAnnotations
                {
                    ReadOnlyHint = true,
                    DestructiveHint = false,
                    IdempotentHint = true,
                    OpenWorldHint = false,
                },
            };
        }

        private static JsonElement CreateOutputSchema()
        {
            var dispatcher = CreateClosedObjectSchema(
                new JsonObject
                {
                    ["capturedManagedThreadId"] = CreateIntegerSchema(int.MaxValue, minimum: 1),
                    ["capturedNativeThreadId"] = CreateIntegerSchema(uint.MaxValue, minimum: 1),
                    ["executedManagedThreadId"] = CreateIntegerSchema(int.MaxValue, minimum: 1),
                    ["executedNativeThreadId"] = CreateIntegerSchema(uint.MaxValue, minimum: 1),
                    ["apartmentState"] = new JsonObject { ["const"] = "STA" },
                    ["matchesCapturedThread"] = new JsonObject { ["const"] = true },
                },
                "capturedManagedThreadId",
                "capturedNativeThreadId",
                "executedManagedThreadId",
                "executedNativeThreadId",
                "apartmentState",
                "matchesCapturedThread");
            var capabilities = CreateClosedObjectSchema(
                new JsonObject
                {
                    ["isExchangeStore"] = new JsonObject { ["type"] = "boolean" },
                    ["isDataFileStore"] = new JsonObject { ["type"] = "boolean" },
                    ["isCachedExchange"] = new JsonObject { ["type"] = "boolean" },
                },
                "isExchangeStore",
                "isDataFileStore",
                "isCachedExchange");
            var standardFolders = CreateClosedObjectSchema(
                new JsonObject
                {
                    ["inbox"] = CreateStringEnum("available", "missing", "unknown"),
                    ["drafts"] = CreateStringEnum("available", "missing", "unknown"),
                    ["sentItems"] = CreateStringEnum("available", "missing", "unknown"),
                    ["deletedItems"] = CreateStringEnum("available", "missing", "unknown"),
                    ["archive"] = CreateStringEnum("available", "missing", "unknown"),
                },
                "inbox",
                "drafts",
                "sentItems",
                "deletedItems",
                "archive");
            var store = CreateClosedObjectSchema(
                new JsonObject
                {
                    ["displayName"] = new JsonObject
                    {
                        ["type"] = "string",
                        ["maxLength"] = OutlookStoreProbe.MaximumDisplayNameLength,
                    },
                    ["storeType"] = CreateStringEnum(
                        "primaryExchangeMailbox",
                        "exchangeMailbox",
                        "exchangePublicFolder",
                        "additionalExchangeMailbox",
                        "nonExchange",
                        "unknown"),
                    ["capabilities"] = capabilities,
                    ["standardFolders"] = standardFolders,
                },
                "displayName",
                "storeType",
                "capabilities",
                "standardFolders");
            var data = CreateClosedObjectSchema(
                new JsonObject
                {
                    ["outlookVersion"] = new JsonObject
                    {
                        ["type"] = "string",
                        ["maxLength"] = OutlookProbeSnapshot.MaximumVersionLength,
                    },
                    ["outlookBitness"] = new JsonObject
                    {
                        ["type"] = "integer",
                        ["enum"] = new JsonArray(32, 64),
                    },
                    ["profileName"] = new JsonObject
                    {
                        ["type"] = "string",
                        ["maxLength"] = OutlookProbeSnapshot.MaximumProfileNameLength,
                    },
                    ["dispatcher"] = dispatcher,
                    ["configuredStoreCount"] = CreateIntegerSchema(int.MaxValue, minimum: 0),
                    ["stores"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["maxItems"] = OutlookProbeSnapshot.MaximumStoreCount,
                        ["items"] = store,
                    },
                },
                "outlookVersion",
                "outlookBitness",
                "profileName",
                "dispatcher",
                "configuredStoreCount",
                "stores");
            var warnings = new JsonObject
            {
                ["type"] = "array",
                ["maxItems"] = OutlookProbeSnapshot.MaximumWarningCount,
                ["items"] = CreateStringEnum(
                    ArchiveAvailabilityWarning,
                    StoreMetadataIncompleteWarning,
                    StoreLimitReachedWarning),
            };
            var success = CreateClosedObjectSchema(
                new JsonObject
                {
                    ["ok"] = new JsonObject { ["const"] = true },
                    ["operationId"] = CreateOperationIdSchema(),
                    ["data"] = data,
                    ["partial"] = new JsonObject { ["type"] = "boolean" },
                    ["warnings"] = warnings,
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
                        "HOST_STOPPING",
                        "QUEUE_FULL",
                        "TIMEOUT",
                        "COM_BUSY",
                        "ACCESS_DENIED",
                        "OBJECT_MODEL_GUARD",
                        "STA_DISPATCH_FAILED",
                        "INTERNAL"),
                    ["message"] = new JsonObject
                    {
                        ["type"] = "string",
                        ["maxLength"] = 256,
                    },
                    ["retryable"] = new JsonObject { ["type"] = "boolean" },
                    ["details"] = new JsonObject
                    {
                        ["type"] = "object",
                        ["properties"] = new JsonObject(),
                        ["additionalProperties"] = false,
                    },
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
            var root = new JsonObject
            {
                ["type"] = "object",
                ["oneOf"] = new JsonArray(success, failure),
            };

            return JsonSerializer.SerializeToElement(root);
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

        private static JsonObject CreateIntegerSchema(long maximum, int minimum)
        {
            return new JsonObject
            {
                ["type"] = "integer",
                ["minimum"] = minimum,
                ["maximum"] = maximum,
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

        private static JsonElement ParseSchema(string json)
        {
            using (var document = JsonDocument.Parse(json))
            {
                return document.RootElement.Clone();
            }
        }
    }
}
