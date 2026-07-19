using System;
using System.Text.Json;
using NUnit.Framework;
using OutlookClassicMcp.Core.Policy;

namespace OutlookClassicMcp.Transport.Tests
{
    [TestFixture]
    public sealed class OutlookStatusCatalogTests
    {
        [Test]
        public void DescriptorExposesOnlyTheBoundedReadOnlyStatusTool()
        {
            var descriptor = OutlookStatusCatalog.CreateDescriptor();

            Assert.That(descriptor.Name, Is.EqualTo(ToolNames.OutlookStatus));
            Assert.That(descriptor.Annotations, Is.Not.Null);
            Assert.That(descriptor.Annotations!.ReadOnlyHint, Is.True);
            Assert.That(descriptor.Annotations.DestructiveHint, Is.False);
            Assert.That(descriptor.Annotations.IdempotentHint, Is.True);
            Assert.That(descriptor.Annotations.OpenWorldHint, Is.False);

            Assert.That(descriptor.InputSchema.GetProperty("type").GetString(), Is.EqualTo("object"));
            Assert.That(descriptor.InputSchema.GetProperty("additionalProperties").GetBoolean(), Is.False);

            var outputSchema = descriptor.OutputSchema;
            Assert.That(outputSchema.HasValue, Is.True);
            Assert.That(outputSchema!.Value.GetProperty("type").GetString(), Is.EqualTo("object"));
            var statusData = outputSchema.Value
                .GetProperty("oneOf")[0]
                .GetProperty("properties")
                .GetProperty("data")
                .GetProperty("properties");
            Assert.That(
                statusData.GetProperty("hostState").GetProperty("maxLength").GetInt32(),
                Is.EqualTo(OutlookStatusSnapshot.MaximumHostStateLength));
            Assert.That(
                statusData.GetProperty("version").GetProperty("maxLength").GetInt32(),
                Is.EqualTo(OutlookStatusSnapshot.MaximumVersionLength));
            var errorProperties = outputSchema.Value
                .GetProperty("oneOf")[1]
                .GetProperty("properties")
                .GetProperty("error")
                .GetProperty("properties");
            Assert.That(errorProperties.GetProperty("code").GetProperty("maxLength").GetInt32(),
                Is.EqualTo(64));
            Assert.That(errorProperties.GetProperty("message").GetProperty("maxLength").GetInt32(),
                Is.EqualTo(256));
            var details = errorProperties.GetProperty("details");
            Assert.That(details.GetProperty("properties").EnumerateObject(), Is.Empty);
            Assert.That(details.GetProperty("additionalProperties").GetBoolean(), Is.False);
        }

        [Test]
        public void DescriptorFactoryDoesNotExposeMutableCatalogState()
        {
            var first = OutlookStatusCatalog.CreateDescriptor();
            first.Name = "changed";
            first.InputSchema = ParseSchema("{\"type\":\"object\",\"properties\":{\"changed\":{}}}");

            var second = OutlookStatusCatalog.CreateDescriptor();

            Assert.That(second.Name, Is.EqualTo(ToolNames.OutlookStatus));
            Assert.That(
                second.InputSchema.GetProperty("additionalProperties").GetBoolean(),
                Is.False);
        }

        [Test]
        public void SnapshotPreservesValidatedSimpleScalars()
        {
            var snapshot = new OutlookStatusSnapshot("Online", listenerReady: true, "1.2.3");

            Assert.That(snapshot.HostState, Is.EqualTo("Online"));
            Assert.That(snapshot.ListenerReady, Is.True);
            Assert.That(snapshot.Version, Is.EqualTo("1.2.3"));
        }

        [Test]
        public void SnapshotRejectsUnboundedOrUnsafeScalars()
        {
            Assert.Throws<ArgumentNullException>((Action)(() =>
                _ = new OutlookStatusSnapshot(null!, true, "1.0")));
            Assert.Throws<ArgumentException>((Action)(() =>
                _ = new OutlookStatusSnapshot(" ", true, "1.0")));
            Assert.Throws<ArgumentException>((Action)(() =>
                _ = new OutlookStatusSnapshot("On\nline", true, "1.0")));
            Assert.Throws<ArgumentOutOfRangeException>((Action)(() =>
                _ = new OutlookStatusSnapshot(
                    new string('x', OutlookStatusSnapshot.MaximumHostStateLength + 1),
                    true,
                    "1.0")));
            Assert.Throws<ArgumentOutOfRangeException>((Action)(() =>
                _ = new OutlookStatusSnapshot(
                    "Online",
                    true,
                    new string('x', OutlookStatusSnapshot.MaximumVersionLength + 1))));
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
