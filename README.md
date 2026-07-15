# Outlook Classic MCP

An open-source, brokerless VSTO add-in that exposes the stores in an active Classic Outlook profile through an authenticated local Model Context Protocol endpoint.

The project is currently in Phase 0: repository, toolchain, build, and runtime-feasibility proof. Mailbox-reading and write tools are not enabled yet.

## Security boundary

The server is designed for one interactive Windows user and binds only to:

```text
http://127.0.0.1:8765/mcp/
```

Every request will require a 256-bit bearer token supplied through `OUTLOOK_MCP_TOKEN`. Email content and attachments are untrusted data and must never be treated as user authority. Consequential tools remain unavailable until their implementation, idempotency, and approval gates pass.

Loopback authentication does not protect against malware, administrators, or injected code running as the same Windows user. Tool results may enter the MCP client's model context and are subject to that client's account, workspace, and retention policies.

## Architecture

The production system uses three assemblies in one `OUTLOOK.EXE` process:

- `OutlookClassicMcp.AddIn`: VSTO lifecycle, Outlook STA dispatch, and Outlook Object Model access.
- `OutlookClassicMcp.Transport`: authenticated loopback HTTP and MCP adaptation, testable without Outlook.
- `OutlookClassicMcp.Core`: Office-free contracts, policy, validation, and tool orchestration.

There is no service, background broker, Graph, EWS, IMAP, or connector fallback.

## Prerequisites

- Windows x64 with Classic Outlook x64
- Visual Studio 2026 with Microsoft 365 development and .NET desktop development
- VSTO 4 runtime 10.0.60910 or newer
- .NET Framework 4.8 SDK and targeting pack
- .NET 10 SDK for the modern Core test target

The committed `.vsconfig` and `tools/preflight.ps1` define and verify the development prerequisites.

## Build

Run from PowerShell:

```powershell
.\tools\preflight.ps1
.\tools\build.ps1
```

Save work and close Outlook gracefully before the build; never terminate it forcibly. The complete solution must be built with full Visual Studio MSBuild. Normal unit and loopback tests do not load Outlook or the VSTO host.

Routine builds are isolated: they refuse to overwrite an existing same-name VSTO development registration and remove their own temporary Outlook registration, certificate, and private-key container. Visual Studio registration and F5 debugging remain a separate interactive workflow.

## Codex development configuration

The reviewed repository-scoped configuration is loaded only when this checkout is trusted by Codex. Validate it without changing the environment:

```powershell
.\tools\configure-codex.ps1 -Action Validate
```

After reviewing the security boundary above, generate or retain a 256-bit token for the current Windows user with `-Action Install`. Outlook and Codex must be restarted when the token changes. Global user registration is intentionally deferred until it can be edited with a TOML-aware, structure-preserving workflow.

## Documentation

- [Implementation plan](docs/IMPLEMENTATION_PLAN.md)
- [Architecture](docs/ARCHITECTURE.md)
- [Tool contracts](docs/TOOL_CONTRACTS.md)
- [Testing](docs/TESTING.md)
- [Phase 0 evidence](docs/PHASE_0_EVIDENCE.md)
- [Dependency licenses](docs/DEPENDENCY_LICENSES.md)
- [Security policy](SECURITY.md)

## License

Licensed under the Apache License 2.0. See [LICENSE](LICENSE).
