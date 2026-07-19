using System.Linq;
using System.Text.Json;
using NUnit.Framework;
using OutlookClassicMcp.Core.Outlook;
using OutlookClassicMcp.Core.Policy;

namespace OutlookClassicMcp.Transport.Tests
{
    [TestFixture]
    public sealed class OutlookProbeCatalogTests
    {
        private static readonly string[] SuccessRequired =
            { "ok", "operationId", "data", "partial", "warnings" };
        private static readonly string[] StoreRequired =
            { "displayName", "storeType", "capabilities", "standardFolders" };
        private static readonly string[] FolderAvailabilityValues =
            { "available", "missing", "unknown" };
        private static readonly string[] StoreTypeValues =
        {
            "primaryExchangeMailbox",
            "exchangeMailbox",
            "exchangePublicFolder",
            "additionalExchangeMailbox",
            "nonExchange",
            "unknown",
        };
        private static readonly string[] WarningValues =
        {
            OutlookProbeCatalog.ArchiveAvailabilityWarning,
            OutlookProbeCatalog.StoreMetadataIncompleteWarning,
            OutlookProbeCatalog.StoreLimitReachedWarning,
        };
        private static readonly string[] ErrorCodes =
        {
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
            "INTERNAL",
        };

        [Test]
        public void DescriptorIsClosedBoundedReadOnlyAndNonDestructive()
        {
            var descriptor = OutlookProbeCatalog.CreateDescriptor();

            Assert.That(descriptor.Name, Is.EqualTo(ToolNames.OutlookProbe));
            Assert.That(descriptor.Annotations, Is.Not.Null);
            Assert.That(descriptor.Annotations!.ReadOnlyHint, Is.True);
            Assert.That(descriptor.Annotations.DestructiveHint, Is.False);
            Assert.That(descriptor.Annotations.IdempotentHint, Is.True);
            Assert.That(descriptor.Annotations.OpenWorldHint, Is.False);
            Assert.That(descriptor.InputSchema.GetProperty("type").GetString(), Is.EqualTo("object"));
            Assert.That(descriptor.InputSchema.GetProperty("properties").EnumerateObject(), Is.Empty);
            Assert.That(descriptor.InputSchema.GetProperty("additionalProperties").GetBoolean(),
                Is.False);

            var output = descriptor.OutputSchema!.Value;
            Assert.That(output.GetProperty("type").GetString(), Is.EqualTo("object"));
            var success = output.GetProperty("oneOf")[0];
            Assert.That(success.GetProperty("additionalProperties").GetBoolean(), Is.False);
            Assert.That(
                success.GetProperty("required").EnumerateArray().Select(item => item.GetString()),
                Is.EquivalentTo(SuccessRequired));

            var properties = success.GetProperty("properties");
            var data = properties.GetProperty("data");
            Assert.That(data.GetProperty("additionalProperties").GetBoolean(), Is.False);
            Assert.That(
                data.GetProperty("properties").GetProperty("outlookVersion")
                    .GetProperty("maxLength").GetInt32(),
                Is.EqualTo(OutlookProbeSnapshot.MaximumVersionLength));
            Assert.That(
                data.GetProperty("properties").GetProperty("profileName")
                    .GetProperty("maxLength").GetInt32(),
                Is.EqualTo(OutlookProbeSnapshot.MaximumProfileNameLength));

            var stores = data.GetProperty("properties").GetProperty("stores");
            Assert.That(stores.GetProperty("maxItems").GetInt32(),
                Is.EqualTo(OutlookProbeSnapshot.MaximumStoreCount));
            var store = stores.GetProperty("items");
            Assert.That(store.GetProperty("additionalProperties").GetBoolean(), Is.False);
            Assert.That(
                store.GetProperty("properties").GetProperty("displayName")
                    .GetProperty("maxLength").GetInt32(),
                Is.EqualTo(OutlookStoreProbe.MaximumDisplayNameLength));
            Assert.That(
                store.GetProperty("required").EnumerateArray().Select(item => item.GetString()),
                Is.EquivalentTo(StoreRequired));
            Assert.That(
                store.GetProperty("properties").GetProperty("storeType").GetProperty("enum")
                    .EnumerateArray().Select(item => item.GetString()),
                Is.EqualTo(StoreTypeValues));

            var folderProperties = store.GetProperty("properties")
                .GetProperty("standardFolders").GetProperty("properties");
            Assert.That(
                folderProperties.GetProperty("archive").GetProperty("enum")
                    .EnumerateArray().Select(item => item.GetString()),
                Is.EqualTo(FolderAvailabilityValues));
            Assert.That(
                store.GetProperty("properties").GetProperty("standardFolders")
                    .GetProperty("additionalProperties").GetBoolean(),
                Is.False);

            var warnings = properties.GetProperty("warnings");
            Assert.That(warnings.GetProperty("maxItems").GetInt32(),
                Is.EqualTo(OutlookProbeSnapshot.MaximumWarningCount));
            Assert.That(
                warnings.GetProperty("items").GetProperty("enum")
                    .EnumerateArray().Select(item => item.GetString()),
                Is.EqualTo(WarningValues));
        }

        [Test]
        public void ErrorSchemaEnumeratesOnlyTheProbeErrorsAndClosesDetails()
        {
            var error = OutlookProbeCatalog.CreateDescriptor().OutputSchema!.Value
                .GetProperty("oneOf")[1]
                .GetProperty("properties")
                .GetProperty("error");
            Assert.That(error.GetProperty("additionalProperties").GetBoolean(), Is.False);
            var properties = error.GetProperty("properties");
            Assert.That(
                properties.GetProperty("code").GetProperty("enum")
                    .EnumerateArray().Select(item => item.GetString()),
                Is.EqualTo(ErrorCodes));
            Assert.That(
                properties.GetProperty("details").GetProperty("additionalProperties").GetBoolean(),
                Is.False);
            Assert.That(properties.GetProperty("details").GetProperty("properties").EnumerateObject(),
                Is.Empty);
        }

        [Test]
        public void DescriptorFactoryDoesNotExposeMutableCatalogState()
        {
            var first = OutlookProbeCatalog.CreateDescriptor();
            first.Name = "changed";
            first.InputSchema = ParseSchema("{\"type\":\"object\",\"properties\":{\"changed\":{}}}");

            var second = OutlookProbeCatalog.CreateDescriptor();

            Assert.That(second.Name, Is.EqualTo(ToolNames.OutlookProbe));
            Assert.That(second.InputSchema.GetProperty("properties").EnumerateObject(), Is.Empty);
            Assert.That(second.InputSchema.GetProperty("additionalProperties").GetBoolean(), Is.False);
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
