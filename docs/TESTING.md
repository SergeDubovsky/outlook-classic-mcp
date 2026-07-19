# Testing

## Automated tests

`OutlookClassicMcp.Core.Tests` targets both `net10.0` and `net48`. `OutlookClassicMcp.Transport.Tests` targets `net48` and uses the real Windows `HttpListener` with no Outlook dependency.

Run:

```powershell
.\tools\preflight.ps1
.\tools\build.ps1
```

The build script performs locked restore, builds Debug and Release/Any CPU with full Visual Studio MSBuild, and runs every test target. Normal test runners must never instantiate or load the VSTO add-in.

Phase 2 transport coverage remains in place and includes:

- exact loopback listener binding and release;
- canonical token parsing, indistinguishable authorization failures, process-snapshotted rotation, and retired-token rejection after restart;
- exact peer, host, route, method, Origin, authorization, media-type, protocol-version, encoding, stateless-session, header, and body validation;
- malformed JSON, top-level batch, body-limit, four-handler concurrency, handler-deadline, recovery, and clean-shutdown behavior;
- the pinned SDK lifecycle, SSE/notification response contracts, tool descriptor and schemas, `outlook_status` success/error envelopes, and a real pinned .NET MCP client;
- twenty unauthorized requests followed by twenty valid requests without wedging the listener.

These tests use a real Windows `HttpListener` with managed status providers. They prove the Transport boundary without loading Outlook or the VSTO assembly.

Phase 3 adds coverage for:

- immutable bounded probe, store, capability, folder-availability, warning, and UI-STA proof contracts;
- the exact `outlook_status` plus `outlook_probe` allowlist and closed input/output schemas;
- serialized bounded dispatcher work, queue saturation, cancellation before dispatch and while queued, wake-up failure, active-work shutdown quiescence, and one-time unavailable notification;
- mapped and sanitized Outlook failures, 14-second tool cancellation inside the 15-second HTTP handler deadline, recovery after timeout, and no COM or VSTO dependency in Core or Transport;
- twenty sequential and four concurrent real-listener probe calls with a fake gateway, plus pinned .NET client coverage.

The accepted Debug and Release builds each passed 75 Core tests on `net10.0`, 97 Core tests on `net48`, and 141 Transport tests on `net48`.

## Endpoint validation

The mailbox-free Phase 2 endpoint check remains available while a listener is running:

```powershell
.\tools\test-phase2-endpoint.ps1
```

The Phase 3 runner invokes `test-phase3-endpoint.ps1` with its privately tracked Outlook instance and independently prepared store inventory. It checks the same transport lifecycle, exact two-tool discovery, status, structured probe envelopes, unique operation identifiers, inventory agreement, dispatcher UI-STA proof, sequential and concurrent stability, and Outlook responsiveness. It never reads message or attachment content.

Only the first probe after each listener start has bounded readiness polling: at most five attempts, 500 ms between retries, and a 30-second total deadline. A retry is permitted only for the exact bounded partial result that reports temporarily incomplete store metadata. The first complete result is reused in the workload count. Every later probe is single-attempt and fail-closed.

## Outlook smoke tests

Outlook smoke tests use a dedicated profile and disposable data on a logged-on interactive Windows desktop. Save work and close Outlook gracefully before registration or F5 tests; never terminate `OUTLOOK.EXE` forcibly.

Phase 0 required one stock-template F5 proof that reached `ThisAddIn_Startup`, plus confirmation that Outlook listed the add-in as active. Phase 1 added a repeatable lifecycle gate; its completed results are preserved in [Phase 1 evidence](PHASE_1_EVIDENCE.md). Phase 1 is not a current executable mode because the add-in always attempts authenticated listener startup. The historical `run-phase1-smoke.ps1` and `verify-phase1-smoke.ps1` filenames remain as shared lifecycle implementation used by later wrappers, but `-ExpectedPhase 1` now fails with an explicit retirement message.

Later phases add bounded message reads, mutations, and approval-gated sending in the order defined by `IMPLEMENTATION_PLAN.md`.

Prepare the ignored independent inventory as described in [the smoke instructions](../smoke/outlook/README.md). Run the Phase 3 live gate from a non-elevated PowerShell process while Outlook is closed:

```powershell
.\tools\configure-codex.ps1 -Action Install
# Restart Codex after token provisioning, and keep Outlook closed.
.\tools\run-phase3-smoke.ps1 `
    -Profile '<dedicated test profile>' `
    -ExpectedStoreInventoryPath '<independently prepared inventory file>'
```

The Phase 3 runner uses full Visual Studio only through `MSBuild.exe` and the OfficeTools targets; it does not launch the IDE. It creates a temporary Release VSTO registration and certificate, requires a canonical current-user token, and owns three normal Outlook start/close cycles. The first cycle runs status, twenty sequential probes, four concurrent probes, and one native Codex probe; the next two cycles run one restart probe each. The runner closes Outlook through its normal main-window path, verifies port 8765 is released in under three seconds, and checks lifecycle, UI-STA, diagnostics, Event ID 45/59, registration, responsiveness, inventory, and cleanup evidence. It removes its temporary registration, certificate, and private-key container and never force-terminates Outlook.

Checking **Active Application Add-ins** and **Disabled Items** in the Outlook UI is optional troubleshooting or release-audit evidence, not part of the routine gate. The verifier ignores well-formed records outside the bounded run but fails closed on malformed retained diagnostics. If a prior interrupted run left a malformed line, explicitly archive the metadata-only log files before recording a new start time.

The final Phase 3 run completed three Outlook cycles successfully on 2026-07-19. Its 27 probes, 28 structured envelopes, independent seven-store inventory match, seven responsiveness checks, UI-STA proof, and cleanup checks passed without a readiness retry. `codex-cli` 0.144.6 made the one native probe call. Full non-content results are recorded in [Phase 3 evidence](PHASE_3_EVIDENCE.md).

## Codex configuration checks

The repository-scoped client configuration is validated with:

```powershell
.\tools\configure-codex.ps1 -Action Validate
```

Validation checks the committed fail-closed TOML, absence of a shadowing user-scoped registration, strict Codex configuration parsing, and the effective exact Streamable HTTP registration. `-Action Install` provisions a current-user token; `-Action Rotate` replaces it; `-Action ClearToken` removes it. Outlook and Codex must restart after any token change because both authenticate with values inherited into their processes.

Configuration validation alone is not a live endpoint acceptance result. Phase 3 acceptance therefore ran the actual Codex CLI against the authenticated endpoint while Outlook was running and deterministically validated its native structured probe result.

## Sensitive data

Tests and CI must not upload mailbox fixtures, message content, Outlook identifiers, bearer tokens, user configuration, runtime logs, or release signing keys.
