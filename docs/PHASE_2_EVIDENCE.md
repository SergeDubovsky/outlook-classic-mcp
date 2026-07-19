# Phase 2 evidence

Status: complete
Last updated: 2026-07-19

## Scope

Phase 2 adds the authenticated, stateless MCP transport without mailbox access. The listener binds only to `http://127.0.0.1:8765/mcp/`, uses the process-snapshotted current-user bearer token, and hosts the pinned `ModelContextProtocol.Core` 1.4.1 transport inside `OUTLOOK.EXE`.

The accepted protocol surface is initialize, `notifications/initialized`, ping, `tools/list`, and `tools/call` for the single `outlook_status` tool. Status returns only bounded host state, listener readiness, and version. No mailbox tools were exposed and no mailbox data was read.

## Build and automated tests

The final Debug and Release builds completed through full Visual Studio CLI MSBuild. Each configuration reported:

| Assembly target | Passed | Failed |
|---|---:|---:|
| Core Tests `net10.0` | 13 | 0 |
| Core Tests `net48` | 16 | 0 |
| Transport Tests `net48` | 114 | 0 |

Before the final protocol-version compatibility case was added, the 113-test Transport suite passed 10 of 10 isolated repetitions. The completed 114-test suite then passed in both Debug and Release. The maximum observed post-process port-release time was 1021 ms. Windows PowerShell 5.1 parsed all seven targeted configuration and smoke scripts successfully.

Raw endpoint validation covered:

- missing authentication returning `401` without capability disclosure;
- initialize, the empty initialized-notification `202`, and ping;
- `tools/list` returning only `outlook_status`;
- `outlook_status` returning online and listener-ready state.

## Token and Codex configuration

The current-user bearer token was installed and read back successfully without disclosing its value. The repository-scoped Codex configuration passed strict validation, and no user-scoped MCP registration with the same name was present to shadow it.

The live client proof used `codex-cli` 0.144.6, not configuration validation or a raw HTTP substitute. During the first Outlook cycle, Codex invoked `outlook_status` successfully and received an online, ready result. The probe rejected all other tool-bearing item types, required a successful terminal turn and native structured content, bounded its output, and verified cleanup of its complete process job.

## Live Outlook acceptance

The repeatable acceptance command is:

```powershell
.\tools\run-phase2-smoke.ps1 -Profile Outlook
```

The final Release gate ran from `2026-07-19 18:08:51 UTC` through `2026-07-19 18:09:56 UTC` against the Classic Outlook profile `Outlook`.

| Check | Result |
|---|---:|
| Distinct Outlook processes | 3 |
| Distinct diagnostics sessions | 3 |
| Endpoint verification successes | 3 |
| Live Codex CLI status calls | 1 |
| Target Event ID 45 records with `Load Behavior: 3` | 3 |
| Target Event ID 59 records | 0 |
| Maximum startup callback | 16.402 ms |
| Maximum port release | 1957.6899 ms |
| Resilience state changes | 0 |
| Residual Outlook processes/listeners | 0 |
| Residual VSTO registration/certificate state | 0 |

The pinned runtime identity fingerprint was `65F5A0725589215872E39F708ABC09C1B5E9AAABB1D1563886EFAC1A3D1FC52F`. Outlook closed normally, port 8765 was free after shutdown, resiliency state remained unchanged, and the temporary registration and signing certificate were removed.

## Phase 2 completion

All Phase 2 acceptance gates are complete. Phase 3 can add the first bounded Outlook/profile/store metadata probe through the established authenticated transport and Outlook STA boundary.
