using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using ModelContextProtocol;
using ModelContextProtocol.Server;

namespace OutlookClassicMcp.Transport
{
    public static class McpDependencyProbe
    {
        private static readonly string[] RequiredRuntimeAssemblyIdentities =
        {
            "Microsoft.Bcl.AsyncInterfaces, Version=10.0.0.7, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51",
            "Microsoft.Bcl.Memory, Version=10.0.0.7, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51",
            "Microsoft.Extensions.AI.Abstractions, Version=10.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35",
            "Microsoft.Extensions.DependencyInjection.Abstractions, Version=10.0.0.7, Culture=neutral, PublicKeyToken=adb9793829ddae60",
            "Microsoft.Extensions.Logging.Abstractions, Version=10.0.0.7, Culture=neutral, PublicKeyToken=adb9793829ddae60",
            "ModelContextProtocol.Core, Version=1.4.1.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51",
            "System.Buffers, Version=4.0.5.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51",
            "System.Diagnostics.DiagnosticSource, Version=10.0.0.7, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51",
            "System.IO.Pipelines, Version=10.0.0.7, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51",
            "System.Memory, Version=4.0.5.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51",
            "System.Net.ServerSentEvents, Version=10.0.0.7, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51",
            "System.Numerics.Vectors, Version=4.1.6.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a",
            "System.Runtime.CompilerServices.Unsafe, Version=6.0.3.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a",
            "System.Text.Encodings.Web, Version=10.0.0.7, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51",
            "System.Text.Json, Version=10.0.0.7, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51",
            "System.Threading.Channels, Version=10.0.0.7, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51",
            "System.Threading.Tasks.Extensions, Version=4.2.4.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51",
        };

        public static McpDependencyLoadReport VerifyLoad()
        {
            var coreAssembly = typeof(McpJsonUtilities).Assembly;
            var coreName = coreAssembly.GetName();
            var coreVersion = coreName.Version
                ?? throw new InvalidOperationException("The MCP Core assembly has no version.");

            if (coreVersion != new Version(1, 4, 1, 0))
            {
                throw new InvalidOperationException($"Unexpected MCP Core version {coreVersion}.");
            }

            var token = coreName.GetPublicKeyToken();
            var tokenText = token == null
                ? string.Empty
                : string.Concat(token.Select(value => value.ToString("x2", CultureInfo.InvariantCulture)));
            if (!string.Equals(tokenText, "cc7b13ffcd2ddd51", StringComparison.Ordinal))
            {
                throw new InvalidOperationException("The MCP Core public key token is not the pinned identity.");
            }

            if (McpJsonUtilities.DefaultOptions == null)
            {
                throw new InvalidOperationException("The MCP JSON serializer options were not initialized.");
            }

            var loadedByName = new Dictionary<string, Assembly>(StringComparer.OrdinalIgnoreCase)
            {
                [coreName.Name] = coreAssembly,
            };

            foreach (var requiredIdentity in RequiredRuntimeAssemblyIdentities)
            {
                var requiredAssemblyName = new AssemblyName(requiredIdentity);
                var requiredName = requiredAssemblyName.Name
                    ?? throw new InvalidOperationException("A required runtime assembly has no name.");
                if (!loadedByName.TryGetValue(requiredName, out var assembly))
                {
                    assembly = Assembly.Load(requiredAssemblyName);
                    loadedByName.Add(requiredName, assembly);
                }

                if (!string.Equals(assembly.FullName, requiredIdentity, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException($"Unexpected runtime identity for {requiredName}.");
                }
                if (assembly.GlobalAssemblyCache)
                {
                    throw new InvalidOperationException($"The runtime assembly {requiredName} loaded from the GAC.");
                }

                _ = assembly.GetTypes();
            }

            var loadedAssemblies = loadedByName.Values
                .OrderBy(assembly => assembly.GetName().Name, StringComparer.Ordinal)
                .ToArray();
            var loadedNames = loadedAssemblies
                .Select(assembly => assembly.GetName().Name)
                .ToArray();
            var loadedIdentities = loadedAssemblies
                .Select(assembly => assembly.FullName ?? assembly.GetName().Name)
                .ToArray();
            var identitySha256 = ComputeIdentitySha256(loadedIdentities);
            if (!string.Equals(
                identitySha256,
                "65F5A0725589215872E39F708ABC09C1B5E9AAABB1D1563886EFAC1A3D1FC52F",
                StringComparison.Ordinal))
            {
                throw new InvalidOperationException("The MCP runtime identity fingerprint is not the pinned closure.");
            }

            VerifyRepresentativeTransportActivation();

            return new McpDependencyLoadReport(coreVersion, loadedNames, loadedIdentities, identitySha256);
        }

        private static void VerifyRepresentativeTransportActivation()
        {
            var transport = new StreamableHttpServerTransport(loggerFactory: null);
            try
            {
                if (transport.MessageReader == null)
                {
                    throw new InvalidOperationException("The MCP transport message reader was not initialized.");
                }

                _ = transport.SessionId;
                _ = transport.Stateless;
                _ = transport.FlowExecutionContextFromRequests;
            }
            finally
            {
                transport.DisposeAsync().AsTask().GetAwaiter().GetResult();
            }
        }

        private static string ComputeIdentitySha256(IEnumerable<string> identities)
        {
            var canonical = string.Join("\n", identities.OrderBy(identity => identity, StringComparer.Ordinal));
            using (var algorithm = SHA256.Create())
            {
                var hash = algorithm.ComputeHash(Encoding.UTF8.GetBytes(canonical));
                return string.Concat(
                    hash.Select(value => value.ToString("X2", CultureInfo.InvariantCulture)));
            }
        }
    }
}
