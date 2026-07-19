# Third-party notices

The VSTO bootstrap files were produced by the installed Microsoft Visual Studio Outlook VSTO project template and wizard. Microsoft documents project templates as reusable project stubs intended for customization. The installed tooling is governed by the [Microsoft Visual Studio Community 2026 license terms](https://visualstudio.microsoft.com/license-terms/vs2026-ga-community/). This repository does not redistribute the template archive, wizard binaries, or Visual Studio components. Beyond that generated bootstrap, no third-party source code is vendored in the project.

## Runtime dependencies

`ModelContextProtocol.Core` 1.4.1 is the Phase 1 production dependency. Its NuGet metadata declares Apache-2.0 and identifies source commit `2b7fd35fbe58dfb9f00eae8b3393e1a7361b5e01` in the [Model Context Protocol C# SDK](https://github.com/modelcontextprotocol/csharp-sdk/tree/2b7fd35fbe58dfb9f00eae8b3393e1a7361b5e01). The upstream license is transitioning from MIT to Apache-2.0: new and consented code is Apache-2.0, while earlier contributions without relicensing consent remain MIT. No SDK source is vendored here.

The repository's root `LICENSE` contains the complete Apache License 2.0 text. The referenced SDK commit contains no `NOTICE` file. The SDK also licenses non-specification documentation under CC-BY-4.0; the current VSTO output does not redistribute that documentation. Any future distribution that includes it must preserve the required CC-BY-4.0 attribution.

The MCP package's runtime transitive dependencies are MIT-licensed `Microsoft.*` and `System.*` packages. Exact package versions, relationships, license expressions, and repository or project URLs are recorded in `docs/DEPENDENCY_LICENSES.md` and checked against restored NuGet metadata by `tools/verify-dependency-inventory.ps1`.

A binary release that redistributes these dependencies must include the repository's `LICENSE` and this notice file, or equivalent copies containing the same license and attribution material.

### Model Context Protocol retained MIT work

MIT License

Copyright (c) 2024-2026 Model Context Protocol a Series of LF Projects, LLC.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

### mcpdotnet retained MIT work

The Model Context Protocol C# SDK's upstream third-party notice identifies retained work from `mcpdotnet` under these terms:

MIT License

Copyright (c) 2024 Peder Holdgaard Pedersen

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

### .NET runtime libraries

The MIT-licensed runtime libraries sourced from `dotnet/dotnet` and `dotnet/maintenance-packages` retain this notice:

The MIT License (MIT)

Copyright (c) .NET Foundation and Contributors
All rights reserved.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

The MIT-licensed `Microsoft.Extensions.AI.Abstractions` package sourced from `dotnet/extensions` retains this notice:

The MIT License (MIT)

Copyright (c) .NET Foundation.
All rights reserved.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

## Development-only dependencies

Test, adapter, and analyzer packages are restored for development and are not redistributed in the VSTO runtime. Their exact MIT license and source metadata is included in `docs/DEPENDENCY_LICENSES.md`. If a future release redistributes any of them, its package-specific license and attribution must be added to the distributed notices.

Behavioral references listed in `docs/IMPLEMENTATION_PLAN.md` are design references only unless a future change records deliberate provenance, the exact source revision, modifications, license, and required attribution here.
