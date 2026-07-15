# Third-party notices

The VSTO bootstrap files were produced by the installed Microsoft Visual Studio Outlook VSTO project template and wizard. Microsoft documents project templates as reusable project stubs intended for customization. The installed tooling is governed by the [Microsoft Visual Studio Community 2026 license terms](https://visualstudio.microsoft.com/license-terms/vs2026-ga-community/). This repository does not redistribute the template archive, wizard binaries, or Visual Studio components. Beyond that generated bootstrap, no third-party source code is vendored in the project.

Production dependencies will be inventoried before they are introduced in Phase 1. Phase 0 test and analyzer packages are centrally pinned and used under their respective licenses:

| Package | Purpose | License | Source |
|---|---|---|---|
| Microsoft.NET.Test.Sdk | Test host and protocol | MIT | https://github.com/microsoft/vstest |
| NUnit | Test framework | MIT | https://github.com/nunit/nunit |
| NUnit3TestAdapter | Visual Studio test adapter | MIT | https://github.com/nunit/nunit3-vs-adapter |
| NUnit.Analyzers | NUnit analyzers | MIT | https://github.com/nunit/nunit.analyzers |
| Microsoft.CodeAnalysis.NetAnalyzers | .NET code analyzers | MIT | https://github.com/dotnet/dotnet |

Exact versions and transitive dependencies are recorded in `Directory.Packages.props`, committed NuGet lock files, and `docs/DEPENDENCY_LICENSES.md`. Redistribution notices must be reviewed whenever dependencies change.

Behavioral references listed in `docs/IMPLEMENTATION_PLAN.md` are design references only unless a future change records deliberate provenance, the exact source revision, modifications, license, and required attribution here.
