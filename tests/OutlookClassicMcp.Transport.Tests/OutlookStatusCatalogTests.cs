using System;
using System.Linq;
using System.Text.Json;
using NUnit.Framework;
using OutlookClassicMcp.Core.Outlook;
using OutlookClassicMcp.Core.Policy;

namespace OutlookClassicMcp.Transport.Tests
{
    [TestFixture]
    public sealed class OutlookStatusCatalogTests
    {
        private static readonly string[] ReadDiagnosticPropertyNames =
        {
            "comAcquired",
            "comReleased",
            "comOutstanding",
            "comPeak",
            "materializedItemHighWater",
        };

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
            Assert.That(statusData.TryGetProperty("readDiagnostics", out _), Is.False);
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
        public void PhaseFourDescriptorRequiresClosedReadDiagnostics()
        {
            var data = OutlookStatusCatalog.CreateDescriptor(includeReadDiagnostics: true)
                .OutputSchema!.Value
                .GetProperty("oneOf")[0]
                .GetProperty("properties")
                .GetProperty("data");
            Assert.That(
                data.GetProperty("required").EnumerateArray()
                    .Select(value => value.GetString()),
                Does.Contain("readDiagnostics"));
            var diagnostics = data.GetProperty("properties").GetProperty("readDiagnostics");
            Assert.That(diagnostics.GetProperty("type").GetString(), Is.EqualTo("object"));
            Assert.That(diagnostics.GetProperty("additionalProperties").GetBoolean(), Is.False);
            Assert.That(
                diagnostics.GetProperty("properties").EnumerateObject()
                    .Select(property => property.Name),
                Is.EqualTo(ReadDiagnosticPropertyNames));
            Assert.That(
                diagnostics.GetProperty("required").EnumerateArray()
                    .Select(value => value.GetString()),
                Is.EqualTo(ReadDiagnosticPropertyNames));
            foreach (var property in diagnostics.GetProperty("properties").EnumerateObject())
            {
                Assert.That(property.Value.GetProperty("type").GetString(), Is.EqualTo("integer"));
                Assert.That(property.Value.GetProperty("minimum").GetInt64(), Is.Zero);
            }
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
            Assert.That(snapshot.ReadDiagnostics, Is.Not.Null);
            Assert.That(snapshot.ReadDiagnostics.ComAcquired, Is.Zero);
            Assert.That(snapshot.ReadDiagnostics.ComReleased, Is.Zero);
            Assert.That(snapshot.ReadDiagnostics.ComOutstanding, Is.Zero);
            Assert.That(snapshot.ReadDiagnostics.ComPeak, Is.Zero);
            Assert.That(snapshot.ReadDiagnostics.MaterializedItemHighWater, Is.Zero);
        }

        [Test]
        public void SnapshotPreservesReadDiagnostics()
        {
            var diagnostics = new OutlookReadDiagnosticsSnapshot(9, 7, 2, 4, 25);
            var snapshot = new OutlookStatusSnapshot(
                "Online",
                listenerReady: true,
                "1.2.3",
                diagnostics);

            Assert.That(snapshot.ReadDiagnostics, Is.SameAs(diagnostics));
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
