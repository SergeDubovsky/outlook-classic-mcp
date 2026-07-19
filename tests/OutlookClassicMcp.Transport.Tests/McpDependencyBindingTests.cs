using System;
using System.Linq;
using System.Net.NetworkInformation;
using NUnit.Framework;

namespace OutlookClassicMcp.Transport.Tests
{
    [TestFixture]
    [NonParallelizable]
    public sealed class McpDependencyBindingTests
    {
        private static readonly string[] RequiredRuntimeAssemblies =
        {
            "Microsoft.Bcl.AsyncInterfaces",
            "Microsoft.Bcl.Memory",
            "Microsoft.Extensions.AI.Abstractions",
            "Microsoft.Extensions.DependencyInjection.Abstractions",
            "Microsoft.Extensions.Logging.Abstractions",
            "ModelContextProtocol.Core",
            "System.Buffers",
            "System.Diagnostics.DiagnosticSource",
            "System.IO.Pipelines",
            "System.Memory",
            "System.Net.ServerSentEvents",
            "System.Numerics.Vectors",
            "System.Runtime.CompilerServices.Unsafe",
            "System.Text.Encodings.Web",
            "System.Text.Json",
            "System.Threading.Channels",
            "System.Threading.Tasks.Extensions",
        };

        [Test]
        public void NetStandardMcpClosureAndTransportActivateWithoutListening()
        {
            Assert.That(IsMcpPortListening(), Is.False, "The MCP port must be free before activation.");

            var report = McpDependencyProbe.VerifyLoad();

            Assert.That(report.CoreVersion, Is.EqualTo(new Version(1, 4, 1, 0)));
            Assert.That(report.LoadedAssemblyIdentities, Has.Count.EqualTo(RequiredRuntimeAssemblies.Length));
            Assert.That(
                report.IdentitySha256,
                Is.EqualTo("65F5A0725589215872E39F708ABC09C1B5E9AAABB1D1563886EFAC1A3D1FC52F"));
            Assert.That(report.LoadedAssemblyNames, Is.SupersetOf(RequiredRuntimeAssemblies));
            Assert.That(
                report.LoadedAssemblyIdentities,
                Has.Some.EqualTo(
                    "ModelContextProtocol.Core, Version=1.4.1.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51"));
            Assert.That(IsMcpPortListening(), Is.False, "Transport activation must not open a listener.");
        }

        private static bool IsMcpPortListening()
        {
            return IPGlobalProperties.GetIPGlobalProperties()
                .GetActiveTcpListeners()
                .Any(endpoint => endpoint.Port == 8765);
        }
    }
}
