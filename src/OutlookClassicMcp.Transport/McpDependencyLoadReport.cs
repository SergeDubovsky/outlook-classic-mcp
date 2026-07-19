using System;
using System.Collections.Generic;

namespace OutlookClassicMcp.Transport
{
    public sealed class McpDependencyLoadReport
    {
        internal McpDependencyLoadReport(
            Version coreVersion,
            IReadOnlyList<string> loadedAssemblyNames,
            IReadOnlyList<string> loadedAssemblyIdentities,
            string identitySha256)
        {
            CoreVersion = coreVersion ?? throw new ArgumentNullException(nameof(coreVersion));
            LoadedAssemblyNames = loadedAssemblyNames ?? throw new ArgumentNullException(nameof(loadedAssemblyNames));
            LoadedAssemblyIdentities = loadedAssemblyIdentities
                ?? throw new ArgumentNullException(nameof(loadedAssemblyIdentities));
            IdentitySha256 = identitySha256 ?? throw new ArgumentNullException(nameof(identitySha256));
        }

        public Version CoreVersion { get; }

        public IReadOnlyList<string> LoadedAssemblyNames { get; }

        public IReadOnlyList<string> LoadedAssemblyIdentities { get; }

        public string IdentitySha256 { get; }
    }
}
