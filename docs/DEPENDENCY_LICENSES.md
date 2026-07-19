# Dependency license inventory

Inventory date: 2026-07-15

The table is derived from the committed NuGet lock files and each restored package's `.nuspec` metadata. Each row represents one exact locked package version. The license and repository/project URL columns are verified against that restored metadata; older packages that predate NuGet license expressions retain reviewed MIT license URLs in their metadata.

Phase 1 introduces `ModelContextProtocol.Core` under Apache-2.0. Every other reviewed package uses the MIT license.

| Package | Version | Relationship | License | Repository/project URL |
|---|---:|---|---|---|
| Microsoft.ApplicationInsights | 2.23.0 | Transitive test dependency | MIT | https://github.com/Microsoft/ApplicationInsights-dotnet |
| Microsoft.Bcl.AsyncInterfaces | 10.0.7 | Transitive runtime | MIT | https://github.com/dotnet/dotnet |
| Microsoft.Bcl.Memory | 10.0.7 | Transitive runtime | MIT | https://github.com/dotnet/dotnet |
| Microsoft.CodeAnalysis.NetAnalyzers | 10.0.302 | Direct analyzer | MIT | https://github.com/dotnet/dotnet |
| Microsoft.CodeCoverage | 18.8.1 | Transitive test dependency | MIT | https://github.com/microsoft/vstest |
| Microsoft.Extensions.AI.Abstractions | 10.5.2 | Transitive runtime | MIT | https://github.com/dotnet/extensions |
| Microsoft.Extensions.DependencyInjection.Abstractions | 10.0.7 | Transitive runtime | MIT | https://github.com/dotnet/dotnet |
| Microsoft.Extensions.DependencyModel | 8.0.2 | Transitive test dependency | MIT | https://github.com/dotnet/runtime |
| Microsoft.Extensions.Logging.Abstractions | 10.0.7 | Transitive runtime | MIT | https://github.com/dotnet/dotnet |
| Microsoft.NET.Test.Sdk | 18.8.1 | Direct test dependency | MIT | https://github.com/microsoft/vstest |
| Microsoft.NETCore.Platforms | 1.1.0 | Transitive framework package | MIT | https://dot.net/ |
| Microsoft.Testing.Extensions.Telemetry | 2.1.0 | Transitive test dependency | MIT | https://github.com/microsoft/testfx |
| Microsoft.Testing.Extensions.TrxReport.Abstractions | 2.1.0 | Transitive test dependency | MIT | https://github.com/microsoft/testfx |
| Microsoft.Testing.Extensions.VSTestBridge | 2.1.0 | Transitive test dependency | MIT | https://github.com/microsoft/testfx |
| Microsoft.Testing.Platform | 2.1.0 | Transitive test dependency | MIT | https://github.com/microsoft/testfx |
| Microsoft.Testing.Platform.MSBuild | 2.1.0 | Transitive test dependency | MIT | https://github.com/microsoft/testfx |
| Microsoft.TestPlatform.ObjectModel | 18.0.1 | Transitive test dependency, target-dependent | MIT | https://github.com/microsoft/vstest |
| Microsoft.TestPlatform.ObjectModel | 18.8.1 | Transitive test dependency, target-dependent | MIT | https://github.com/microsoft/vstest |
| Microsoft.TestPlatform.TestHost | 18.8.1 | Transitive test dependency | MIT | https://github.com/microsoft/vstest |
| ModelContextProtocol.Core | 1.4.1 | Direct runtime | Apache-2.0 | https://github.com/modelcontextprotocol/csharp-sdk |
| NETStandard.Library | 2.0.3 | Framework package | MIT | https://dot.net/ |
| NUnit | 4.6.1 | Direct test dependency | MIT | https://github.com/nunit/nunit |
| NUnit.Analyzers | 4.14.0 | Direct test analyzer | MIT | https://github.com/nunit/nunit.analyzers |
| NUnit3TestAdapter | 6.2.0 | Direct test adapter | MIT | https://github.com/nunit/nunit3-vs-adapter |
| System.Buffers | 4.5.1 | Transitive test dependency | MIT | https://dot.net/ |
| System.Buffers | 4.6.1 | Transitive runtime and test dependency | MIT | https://github.com/dotnet/maintenance-packages |
| System.Collections.Immutable | 8.0.0 | Transitive test dependency | MIT | https://github.com/dotnet/runtime |
| System.Diagnostics.DiagnosticSource | 6.0.0 | Transitive test dependency | MIT | https://github.com/dotnet/runtime |
| System.Diagnostics.DiagnosticSource | 10.0.7 | Transitive runtime and test dependency | MIT | https://github.com/dotnet/dotnet |
| System.IO.Pipelines | 10.0.7 | Transitive runtime | MIT | https://github.com/dotnet/dotnet |
| System.Memory | 4.5.5 | Transitive test dependency | MIT | https://dot.net/ |
| System.Memory | 4.6.3 | Transitive runtime and test dependency | MIT | https://github.com/dotnet/maintenance-packages |
| System.Net.ServerSentEvents | 10.0.7 | Transitive runtime | MIT | https://github.com/dotnet/dotnet |
| System.Numerics.Vectors | 4.5.0 | Transitive test dependency | MIT | https://dot.net/ |
| System.Numerics.Vectors | 4.6.1 | Transitive runtime and test dependency | MIT | https://github.com/dotnet/maintenance-packages |
| System.Reflection.Metadata | 8.0.0 | Transitive test dependency | MIT | https://github.com/dotnet/runtime |
| System.Runtime.CompilerServices.Unsafe | 6.0.0 | Transitive test dependency | MIT | https://github.com/dotnet/runtime |
| System.Runtime.CompilerServices.Unsafe | 6.1.2 | Transitive runtime and test dependency | MIT | https://github.com/dotnet/maintenance-packages |
| System.Text.Encodings.Web | 10.0.7 | Transitive runtime | MIT | https://github.com/dotnet/dotnet |
| System.Text.Json | 10.0.7 | Transitive runtime | MIT | https://github.com/dotnet/dotnet |
| System.Threading.Channels | 10.0.7 | Transitive runtime | MIT | https://github.com/dotnet/dotnet |
| System.Threading.Tasks.Extensions | 4.5.4 | Transitive test dependency | MIT | https://dot.net/ |
| System.Threading.Tasks.Extensions | 4.6.3 | Transitive runtime and test dependency | MIT | https://github.com/dotnet/maintenance-packages |
| System.ValueTuple | 4.4.0 | Transitive test dependency | MIT | https://dot.net/ |
| System.ValueTuple | 4.6.2 | Transitive package; no net48 runtime DLL | MIT | https://github.com/dotnet/maintenance-packages |

No reviewed package requires copyleft. The runtime redistribution obligations and retained notices are recorded in `THIRD_PARTY_NOTICES.md`. Re-run `tools/verify-dependency-inventory.ps1` and review both notice files before every dependency change and release.
