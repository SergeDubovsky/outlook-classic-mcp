# Dependency license inventory

Inventory date: 2026-07-15

The table is derived from the committed NuGet lock files and each restored package's `.nuspec` metadata. Phase 0 has no runtime MCP dependency; its external packages support compilation, analysis, and testing. Every reviewed package uses the MIT license.

| Package | Version | Relationship | Source |
|---|---:|---|---|
| Microsoft.ApplicationInsights | 2.23.0 | Transitive | https://github.com/microsoft/ApplicationInsights-dotnet |
| Microsoft.CodeAnalysis.NetAnalyzers | 10.0.302 | Direct analyzer | https://github.com/dotnet/dotnet |
| Microsoft.CodeCoverage | 18.8.1 | Transitive | https://github.com/microsoft/vstest |
| Microsoft.Extensions.DependencyModel | 8.0.2 | Transitive | https://github.com/dotnet/runtime |
| Microsoft.NET.Test.Sdk | 18.8.1 | Direct test dependency | https://github.com/microsoft/vstest |
| Microsoft.NETCore.Platforms | 1.1.0 | Transitive | https://github.com/dotnet/corefx |
| Microsoft.Testing.Extensions.Telemetry | 2.1.0 | Transitive | https://github.com/microsoft/testfx |
| Microsoft.Testing.Extensions.TrxReport.Abstractions | 2.1.0 | Transitive | https://github.com/microsoft/testfx |
| Microsoft.Testing.Extensions.VSTestBridge | 2.1.0 | Transitive | https://github.com/microsoft/testfx |
| Microsoft.Testing.Platform | 2.1.0 | Transitive | https://github.com/microsoft/testfx |
| Microsoft.Testing.Platform.MSBuild | 2.1.0 | Transitive | https://github.com/microsoft/testfx |
| Microsoft.TestPlatform.ObjectModel | 18.0.1, 18.8.1 | Transitive, target-dependent | https://github.com/microsoft/vstest |
| Microsoft.TestPlatform.TestHost | 18.8.1 | Transitive | https://github.com/microsoft/vstest |
| NETStandard.Library | 2.0.3 | Framework package | https://github.com/dotnet/standard |
| NUnit | 4.6.1 | Direct test dependency | https://github.com/nunit/nunit |
| NUnit.Analyzers | 4.14.0 | Direct test analyzer | https://github.com/nunit/nunit.analyzers |
| NUnit3TestAdapter | 6.2.0 | Direct test adapter | https://github.com/nunit/nunit3-vs-adapter |
| System.Buffers | 4.5.1 | Transitive | https://github.com/dotnet/corefx |
| System.Collections.Immutable | 8.0.0 | Transitive | https://github.com/dotnet/runtime |
| System.Diagnostics.DiagnosticSource | 6.0.0 | Transitive | https://github.com/dotnet/runtime |
| System.Memory | 4.5.5 | Transitive | https://github.com/dotnet/corefx |
| System.Numerics.Vectors | 4.5.0 | Transitive | https://github.com/dotnet/corefx |
| System.Reflection.Metadata | 8.0.0 | Transitive | https://github.com/dotnet/runtime |
| System.Runtime.CompilerServices.Unsafe | 6.0.0 | Transitive | https://github.com/dotnet/runtime |
| System.Threading.Tasks.Extensions | 4.5.4 | Transitive | https://github.com/dotnet/corefx |
| System.ValueTuple | 4.4.0 | Transitive | https://github.com/dotnet/corefx |

The source repositories above contain the applicable MIT license text. No Phase 0 package requires copyleft, attribution beyond its license, or redistribution-specific notices. Re-run this inventory and update `THIRD_PARTY_NOTICES.md` before every dependency change and release.
