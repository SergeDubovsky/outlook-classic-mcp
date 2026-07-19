# Testing

## Automated tests

`OutlookClassicMcp.Core.Tests` targets both `net10.0` and `net48`. `OutlookClassicMcp.Transport.Tests` targets `net48` and uses the real Windows `HttpListener` with no Outlook dependency.

Run:

```powershell
.\tools\preflight.ps1
.\tools\build.ps1
```

The build script performs locked restore, builds Debug and Release/Any CPU with full Visual Studio MSBuild, and runs every test target. Normal test runners must never instantiate or load the VSTO add-in.

Phase 2 automated coverage includes:

- exact loopback listener binding and release;
- canonical token parsing, indistinguishable authorization failures, process-snapshotted rotation, and retired-token rejection after restart;
- exact peer, host, route, method, Origin, authorization, media-type, protocol-version, encoding, stateless-session, header, and body validation;
- malformed JSON, top-level batch, body-limit, four-handler concurrency, handler-deadline, recovery, and clean-shutdown behavior;
- the pinned SDK lifecycle, SSE/notification response contracts, tool descriptor and schemas, `outlook_status` success/error envelopes, and a real pinned .NET MCP client;
- twenty unauthorized requests followed by twenty valid requests without wedging the listener.

These tests use a real Windows `HttpListener` with managed status providers. They prove the Transport boundary without loading Outlook or the VSTO assembly.

## Endpoint CLI probe

With a Phase 2 listener running under Outlook, run:

```powershell
.\tools\test-phase2-endpoint.ps1
```

The probe uses only the exact production URL, disables proxying, reads the token from process or current-user scope, and checks an unauthenticated `401`, initialize, the empty initialized-notification `202`, ping, discovery of exactly `outlook_status`, and an online status result. It also checks the SSE cache and identity-encoding response contract. This probe does not enumerate or read mailbox data.

## Outlook smoke tests

Outlook smoke tests use a dedicated profile and disposable data on a logged-on interactive Windows desktop. Save work and close Outlook gracefully before registration or F5 tests; never terminate `OUTLOOK.EXE` forcibly.

Phase 0 required one stock-template F5 proof that reached `ThisAddIn_Startup`, plus confirmation that Outlook listed the add-in as active. Phase 1 added a repeatable lifecycle gate; its completed results are preserved in [Phase 1 evidence](PHASE_1_EVIDENCE.md). Phase 1 is not a current executable mode because the Phase 2 add-in always attempts authenticated listener startup. The historical `run-phase1-smoke.ps1` and `verify-phase1-smoke.ps1` filenames remain as shared lifecycle implementation used by the Phase 2 wrapper, but `-ExpectedPhase 1` now fails with an explicit retirement message.

Later phases add store probes, bounded reads, mutations, and approval-gated sending in the order defined by `IMPLEMENTATION_PLAN.md`.

Run the Phase 2 live gate from a non-elevated PowerShell process while Outlook is closed:

```powershell
.\tools\configure-codex.ps1 -Action Install
# Restart Codex after token provisioning, and keep Outlook closed.
.\tools\run-phase2-smoke.ps1 -Profile Outlook
```

The Phase 2 runner uses full Visual Studio only through `MSBuild.exe` and the OfficeTools targets; it does not launch the IDE. It creates a temporary Release VSTO registration and certificate, requires a canonical current-user token, and starts three separately owned Outlook processes. In every cycle it waits for listener readiness, runs the endpoint CLI probe, closes Outlook through its normal main-window path, verifies port 8765 is released in under three seconds, and checks the Phase 1 dependency, STA, diagnostics, Event ID 45/59, registration, and cleanup evidence. It then removes its temporary registration, certificate, and private-key container. It never force-terminates Outlook.

Checking **Active Application Add-ins** and **Disabled Items** in the Outlook UI is optional troubleshooting or release-audit evidence, not part of the routine gate. The verifier ignores well-formed records outside the bounded run but fails closed on malformed retained diagnostics. If a prior interrupted run left a malformed line, explicitly archive the metadata-only log files before recording a new start time.

The final Phase 2 run completed three Outlook cycles successfully on 2026-07-19. The endpoint passed in all three cycles, and `codex-cli` 0.144.6 invoked `outlook_status` in the first cycle and received an online, ready result. Full results are recorded in [Phase 2 evidence](PHASE_2_EVIDENCE.md).

## Codex configuration checks

The repository-scoped client configuration is validated with:

```powershell
.\tools\configure-codex.ps1 -Action Validate
```

Validation checks the committed fail-closed TOML, absence of a shadowing user-scoped registration, strict Codex configuration parsing, and the effective exact Streamable HTTP registration. `-Action Install` provisions a current-user token; `-Action Rotate` replaces it; `-Action ClearToken` removes it. Outlook and Codex must restart after any token change because both authenticate with values inherited into their processes.

Configuration validation alone is not a live endpoint acceptance result. Phase 2 acceptance therefore also ran the actual Codex CLI against the authenticated endpoint while Outlook was running; that separate live gate passed.

## Sensitive data

Tests and CI must not upload mailbox fixtures, message content, Outlook identifiers, bearer tokens, user configuration, runtime logs, or release signing keys.
