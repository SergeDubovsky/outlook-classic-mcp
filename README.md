# Outlook Classic MCP

An open-source, brokerless VSTO add-in that exposes the stores in an active Classic Outlook profile through an authenticated local Model Context Protocol endpoint.

Phases 0 through 3 are complete. The authenticated MCP listener, mailbox-free status surface, and bounded Outlook store probe have passed automated, live Outlook, and live Codex acceptance. Message-reading and write tools remain unavailable.

## Security boundary

The server is designed for one interactive Windows user and binds only to:

```text
http://127.0.0.1:8765/mcp/
```

Every request requires a canonical 256-bit bearer token supplied through `OUTLOOK_MCP_TOKEN`. The setup tool stores it for the current Windows user; Outlook and Codex inherit that value when their processes start. The add-in snapshots the token from the Outlook process environment when it creates the listener, so installing or rotating the token requires both processes to restart. A missing or invalid token leaves the host degraded and the listener offline; there is no unauthenticated fallback.

The HTTP boundary accepts only authenticated `POST` requests at the exact endpoint above. It rejects non-loopback peers, alternate host or route values, browser `Origin` headers, malformed or duplicate authorization, unsupported media negotiation, stateful session headers, oversized requests, and excess concurrency before mailbox access is possible. Email content and attachments are untrusted data and must never be treated as user authority. Consequential tools remain unavailable until their implementation, idempotency, and approval gates pass.

Loopback authentication does not protect against malware, administrators, or injected code running as the same Windows user. Tool results may enter the MCP client's model context and are subject to that client's account, workspace, and retention policies.

## Architecture

The production system uses three assemblies in one `OUTLOOK.EXE` process:

- `OutlookClassicMcp.AddIn`: VSTO lifecycle, Outlook STA dispatch, and Outlook Object Model access.
- `OutlookClassicMcp.Transport`: authenticated loopback HTTP and MCP adaptation, testable without Outlook.
- `OutlookClassicMcp.Core`: Office-free contracts, policy, validation, and tool orchestration.

There is no service, background broker, Graph, EWS, IMAP, or connector fallback.

The pinned `ModelContextProtocol.Core` 1.4.1 SDK runs in stateless Streamable HTTP mode. The implemented protocol surface is initialization, the initialized notification, ping, tool discovery, and calls to exactly two tools: `outlook_status` and `outlook_probe`. Status returns bounded managed host state. Probe dispatches synchronously onto the captured Outlook UI STA and returns bounded Outlook version, profile, store capability, and standard-folder availability metadata as immutable managed data. It does not read messages, bodies, attachments, or store identifiers.

## Prerequisites

- Windows x64 with Classic Outlook x64
- Visual Studio 2026 with Microsoft 365 development and .NET desktop development
- Visual Studio-installed Office 15 primary interop assemblies
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

The completed Phase 1 lifecycle acceptance is preserved in [Phase 1 evidence](docs/PHASE_1_EVIDENCE.md). It is no longer a current runnable gate: the Phase 2 add-in always attempts authenticated listener startup, so Phase 1 mode in the historical lifecycle runner and verifier is explicitly retired.

Prepare an independent store inventory as described in [the Phase 3 smoke instructions](smoke/outlook/README.md), then run the current live gate against a dedicated Outlook profile with:

```powershell
.\tools\run-phase3-smoke.ps1 `
    -Profile '<dedicated test profile>' `
    -ExpectedStoreInventoryPath '<independently prepared inventory file>'
```

Provision the current-user token first, save work, and close Outlook gracefully before running it. The runner uses the installed Visual Studio toolchain only through `MSBuild.exe` and its OfficeTools targets; it does not launch the Visual Studio IDE. It creates a temporary Release VSTO registration, runs three normal Outlook start/close cycles, checks Outlook Event IDs 45 and 59 together with lifecycle metadata, and removes its temporary registration and certificate. It never force-terminates Outlook.

The gate checks the unauthenticated failure, initialize/initialized, ping, exact two-tool discovery, status, twenty sequential probes, four concurrent probes, one native Codex probe, post-restart probes, independent store-inventory agreement, UI-STA execution, Outlook responsiveness, and clean port release. Only the first probe after each listener start may retry an exact metadata-incomplete partial result, with at most five attempts and a 30-second total deadline. Every later probe is single-shot and fail-closed.

The final three-cycle Outlook and Codex acceptance run passed on 2026-07-19. See [Phase 3 evidence](docs/PHASE_3_EVIDENCE.md) for the recorded results.

## Codex development configuration

The reviewed repository-scoped configuration is loaded only when this checkout is trusted by Codex. Validate it without changing the environment:

```powershell
.\tools\configure-codex.ps1 -Action Validate
```

After reviewing the security boundary above, generate or retain a 256-bit token for the current Windows user with `-Action Install`:

```powershell
.\tools\configure-codex.ps1 -Action Install
```

Restart Outlook and Codex whenever the token is installed, rotated, or cleared, then rerun `-Action Validate`. `-Action Rotate` replaces the current-user token; only newly started Outlook and Codex processes inherit the replacement. The committed `.codex/config.toml` is repository-scoped, uses the exact endpoint and `OUTLOOK_MCP_TOKEN`, and is loaded only when this checkout is trusted. Global user registration remains intentionally deferred.

The live acceptance gate used `codex-cli` 0.144.6 to invoke `outlook_probe` exactly once against a running Outlook instance. A deterministic validator accepted its native structured result without retaining profile or store labels.

## Documentation

- [Implementation plan](docs/IMPLEMENTATION_PLAN.md)
- [Architecture](docs/ARCHITECTURE.md)
- [STA dispatcher wake-up decision](docs/ADR_0001_STA_DISPATCH_WAKEUP.md)
- [Tool contracts](docs/TOOL_CONTRACTS.md)
- [Testing](docs/TESTING.md)
- [Phase 0 evidence](docs/PHASE_0_EVIDENCE.md)
- [Phase 1 evidence](docs/PHASE_1_EVIDENCE.md)
- [Phase 2 evidence](docs/PHASE_2_EVIDENCE.md)
- [Phase 3 evidence](docs/PHASE_3_EVIDENCE.md)
- [Dependency licenses](docs/DEPENDENCY_LICENSES.md)
- [Security policy](SECURITY.md)

## License

Licensed under the Apache License 2.0. See [LICENSE](LICENSE).
