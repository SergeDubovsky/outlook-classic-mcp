using System;
using System.Text.Json;
using System.Text.Json.Nodes;
using ModelContextProtocol.Protocol;
using OutlookClassicMcp.Core.Outlook;
using OutlookClassicMcp.Core.Policy;

namespace OutlookClassicMcp.Transport
{
    /// <summary>
    /// Contains the non-mailbox status values exposed by the Phase 2 MCP endpoint.
    /// </summary>
    public sealed class OutlookStatusSnapshot
    {
        public const int MaximumHostStateLength = 32;
        public const int MaximumVersionLength = 64;

        public OutlookStatusSnapshot(
            string hostState,
            bool listenerReady,
            string version,
            OutlookReadDiagnosticsSnapshot? readDiagnostics = null)
        {
            HostState = ValidateScalar(hostState, nameof(hostState), MaximumHostStateLength);
            ListenerReady = listenerReady;
            Version = ValidateScalar(version, nameof(version), MaximumVersionLength);
            ReadDiagnostics = readDiagnostics ?? new OutlookReadDiagnosticsSnapshot(0, 0, 0, 0, 0);
        }

        public string HostState { get; }

        public bool ListenerReady { get; }

        public string Version { get; }

        public OutlookReadDiagnosticsSnapshot ReadDiagnostics { get; }

        private static string ValidateScalar(string value, string parameterName, int maximumLength)
        {
            if (value == null)
            {
                throw new ArgumentNullException(parameterName);
            }

            if (string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentException("The value must not be empty or whitespace.", parameterName);
            }

            if (value.Length > maximumLength)
            {
                throw new ArgumentOutOfRangeException(
                    parameterName,
                    $"The value must not exceed {maximumLength} characters.");
            }

            for (var index = 0; index < value.Length; index++)
            {
                if (char.IsControl(value[index]))
                {
                    throw new ArgumentException("The value must not contain control characters.", parameterName);
                }
            }

            return value;
        }
    }

    /// <summary>
    /// Creates isolated protocol descriptors for the tools available in Phase 2.
    /// </summary>
    public static class OutlookStatusCatalog
    {
        private static readonly JsonElement InputSchema = ParseSchema(
            "{\"type\":\"object\",\"properties\":{},\"additionalProperties\":false}");

        private static readonly JsonElement OutputSchema = ParseSchema(
            "{\"type\":\"object\",\"oneOf\":[" +
            "{\"properties\":{" +
            "\"ok\":{\"const\":true}," +
            "\"operationId\":{\"type\":\"string\",\"pattern\":\"^[0-9a-f]{32}$\"}," +
            "\"data\":{\"type\":\"object\",\"properties\":{" +
            "\"hostState\":{\"type\":\"string\",\"maxLength\":32}," +
            "\"listenerReady\":{\"type\":\"boolean\"}," +
            "\"version\":{\"type\":\"string\",\"maxLength\":64}}," +
            "\"required\":[\"hostState\",\"listenerReady\",\"version\"],\"additionalProperties\":false}," +
            "\"partial\":{\"const\":false}," +
            "\"warnings\":{\"type\":\"array\",\"maxItems\":0}}," +
            "\"required\":[\"ok\",\"operationId\",\"data\",\"partial\",\"warnings\"]," +
            "\"additionalProperties\":false}," +
            "{\"properties\":{" +
            "\"ok\":{\"const\":false}," +
            "\"operationId\":{\"type\":\"string\",\"pattern\":\"^[0-9a-f]{32}$\"}," +
            "\"error\":{\"type\":\"object\",\"properties\":{" +
            "\"code\":{\"type\":\"string\",\"maxLength\":64}," +
            "\"message\":{\"type\":\"string\",\"maxLength\":256}," +
            "\"retryable\":{\"type\":\"boolean\"}," +
            "\"details\":{\"type\":\"object\",\"properties\":{},\"additionalProperties\":false}}," +
            "\"required\":[\"code\",\"message\",\"retryable\",\"details\"],\"additionalProperties\":false}}," +
            "\"required\":[\"ok\",\"operationId\",\"error\"]," +
            "\"additionalProperties\":false}]}");

        private static readonly JsonElement OutputSchemaWithReadDiagnostics =
            CreateOutputSchemaWithReadDiagnostics();

        public static Tool CreateDescriptor()
        {
            return CreateDescriptor(includeReadDiagnostics: false);
        }

        public static Tool CreateDescriptor(bool includeReadDiagnostics)
        {
            return new Tool
            {
                Name = ToolNames.OutlookStatus,
                Title = "Outlook status",
                Description = "Reports bounded add-in and listener readiness without reading Outlook stores, folders, messages, or attachments.",
                InputSchema = InputSchema,
                OutputSchema = includeReadDiagnostics
                    ? OutputSchemaWithReadDiagnostics
                    : OutputSchema,
                Annotations = new ToolAnnotations
                {
                    ReadOnlyHint = true,
                    DestructiveHint = false,
                    IdempotentHint = true,
                    OpenWorldHint = false,
                },
            };
        }

        private static JsonElement CreateOutputSchemaWithReadDiagnostics()
        {
            var schema = JsonNode.Parse(OutputSchema.GetRawText())!.AsObject();
            var data = schema["oneOf"]!.AsArray()[0]!["properties"]!["data"]!.AsObject();
            data["properties"]!.AsObject()["readDiagnostics"] = new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["comAcquired"] = CreateNonnegativeIntegerSchema(),
                    ["comReleased"] = CreateNonnegativeIntegerSchema(),
                    ["comOutstanding"] = CreateNonnegativeIntegerSchema(),
                    ["comPeak"] = CreateNonnegativeIntegerSchema(),
                    ["materializedItemHighWater"] = CreateNonnegativeIntegerSchema(),
                },
                ["required"] = new JsonArray(
                    "comAcquired",
                    "comReleased",
                    "comOutstanding",
                    "comPeak",
                    "materializedItemHighWater"),
                ["additionalProperties"] = false,
            };
            data["required"]!.AsArray().Add("readDiagnostics");
            return JsonSerializer.SerializeToElement(schema);
        }

        private static JsonObject CreateNonnegativeIntegerSchema()
        {
            return new JsonObject
            {
                ["type"] = "integer",
                ["minimum"] = 0,
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
