#if NET48
using System;
using System.Reflection;
using NUnit.Framework;
using OutlookClassicMcp.AddIn.Runtime;

namespace OutlookClassicMcp.Core.Tests
{
    [TestFixture]
    public sealed class MetadataDiagnosticsPrivacyTests
    {
        [Test]
        public void DiagnosticLineDoesNotIncludeExceptionMessageOrData()
        {
            const string privateMarker = "private-subject-private-store-private-entry";
            var exception = new InvalidOperationException(privateMarker);
            exception.Data["mailbox"] = privateMarker;
            using (var diagnostics = (MetadataDiagnostics?)Activator.CreateInstance(
                typeof(MetadataDiagnostics),
                nonPublic: true))
            {
                Assert.That(diagnostics, Is.Not.Null);
                var createLine = typeof(MetadataDiagnostics).GetMethod(
                    "CreateLine",
                    BindingFlags.Instance | BindingFlags.NonPublic);
                Assert.That(createLine, Is.Not.Null);

                var line = (string?)createLine!.Invoke(
                    diagnostics,
                    new object?[]
                    {
                        RuntimeDiagnosticEvent.StartupCompleted,
                        HostLifecycleState.Degraded,
                        false,
                        1L,
                        2,
                        3,
                        null,
                        null,
                        null,
                        null,
                        null,
                        null,
                        null,
                        null,
                        exception,
                    });

                Assert.Multiple((Action)(() =>
                {
                    Assert.That(line, Does.Contain(
                        "exception_type=System.InvalidOperationException"));
                    Assert.That(line, Does.Not.Contain(privateMarker));
                }));
            }
        }
    }
}
#endif
