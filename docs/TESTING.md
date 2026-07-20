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

The accepted Phase 3 Debug and Release builds each passed 75 Core tests on `net10.0`, 97 Core tests on `net48`, and 141 Transport tests on `net48`.

Phase 4 adds automated coverage for:

- exact nine-tool ordering, read-only annotations, closed schemas, page/body/search/result bounds, and absence of all write or attachment-export tools;
- store-qualified locator validation, deterministic keyset ordering, cursor HMAC derivation/authentication, canonical query binding, cross-tool/query rejection, anchor reacquisition, and `CURSOR_STALE` mapping;
- mailbox, folder, message, body, recipient, conversation, attachment, partial-scope, warning, and non-content read-diagnostic contracts;
- fixed internal DASL construction and escaping, canonical multi-scope search, bounded merge/deduplication, all-scope failure, and partial-result cursor suppression;
- Search Folder fail-closed detection, 1,024-folder bound, UI-STA assertions, explicit COM ownership telemetry, page-size-plus-one materialization, conversation depth/node limits, and Core/Transport assembly isolation;
- real-listener calls for every read tool, invalid/tampered cursors, result-size trimming, structured timeout recovery, sequential/concurrent workloads, and pinned-client behavior.

The final 2026-07-19 Release checkpoint passed 123 Core tests on `net10.0`, 170 Core tests on `net48`, and 211 Transport tests on `net48`; the AddIn Release build also succeeded after the latest fixture and harness changes. This is not final Phase 4 acceptance because the live evidence remains pending.

## Endpoint validation

The mailbox-free Phase 2 endpoint check remains available while a listener is running:

```powershell
.\tools\test-phase2-endpoint.ps1
```

The Phase 3 runner invokes `test-phase3-endpoint.ps1` with its privately tracked Outlook instance and independently prepared store inventory. It checks the same transport lifecycle, exact two-tool discovery, status, structured probe envelopes, unique operation identifiers, inventory agreement, dispatcher UI-STA proof, sequential and concurrent stability, and Outlook responsiveness. It never reads message or attachment content.

Only the first probe after each listener start has bounded readiness polling: at most five attempts, 500 ms between retries, and a 30-second total deadline. A retry is permitted only for the exact bounded partial result that reports temporarily incomplete store metadata. The first complete result is reused in the workload count. Every later probe is single-attempt and fail-closed.

The Phase 4 runner invokes `test-phase4-endpoint.ps1` once in full-workload mode and twice in restart mode. All cycles require exact nine-tool discovery, Phase 4 status counters, independent complete-inventory agreement, captured UI-STA proof, unique operation IDs, and Outlook responsiveness. The full cycle additionally exercises known-message discovery/reacquisition in at least two stores, three static pages of four items, tampered and wrong-kind cursor rejection, at least 1,001 large-folder items through bounded pagination, partial cross-store search, conversation retrieval, attachment metadata, bounded body truncation, optional protected-body behavior, client cancellation and recovery, four concurrent reads, 200 repeated reads, process-resource sampling, balanced COM acquisition/release, and a materialization high-water no greater than 51.

`test-phase4-codex.ps1` isolates Codex to the repository-scoped `outlook_classic` server and permits exactly one `outlook_list_mailboxes` call with page size 50. It validates the native structured result and returns only `MCP_PHASE4_CODEX_OK:<count>`; mailbox records and all returned strings are registered as privacy sentinels and suppressed from retained output.

## Outlook smoke tests

Outlook smoke tests use a dedicated profile and disposable data on a logged-on interactive Windows desktop. Save work and close Outlook gracefully before registration or F5 tests; never terminate `OUTLOOK.EXE` forcibly.

Phase 0 required one stock-template F5 proof that reached `ThisAddIn_Startup`, plus confirmation that Outlook listed the add-in as active. Phase 1 added a repeatable lifecycle gate; its completed results are preserved in [Phase 1 evidence](PHASE_1_EVIDENCE.md). Phase 1 is not a current executable mode because the add-in always attempts authenticated listener startup. The historical `run-phase1-smoke.ps1` and `verify-phase1-smoke.ps1` filenames remain as shared lifecycle implementation used by later wrappers, but `-ExpectedPhase 1` now fails with an explicit retirement message.

Phase 4 adds bounded message reads only. Mutations and approval-gated sending remain unavailable and continue in the order defined by `IMPLEMENTATION_PLAN.md`.

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

The current Phase 4 gate uses the synthetic fixture documented in [the Phase 4 smoke instructions](../smoke/outlook/PHASE_4_README.md). The conditional alternative is compiled only with `OutlookMcpSmokeSeeder=true` and runs in `OUTLOOK.EXE` on the verified Outlook UI STA without starting the normal MCP host. It requires an accountless dedicated profile and covers its complete three-store inventory:

- `bootstrap_store` maps to the profile's real bootstrap default store, has `expectsInbox: true`, and seeds/reacquires its known item through the standard Inbox;
- `fixture_store_a` and `fixture_store_b` are attached disposable PST stores with deterministic run-specific display names. Their `expectsInbox: false` values instruct the verifier to use each known message's bounded custom `folderPath`; they do not assert that Outlook failed to create an Inbox in those PSTs.

This conditional three-store fixture does not replace the checked-in manual template, which remains `source: classic-outlook-ui` with two stores whose `expectsInbox` values are both true.

The generated fixture also provides 12 static messages whose recorded inventory captures Outlook's actual ordering—`ReceivedTime` when usable, followed by a deterministic `EntryID` tie-break—for three pages of four, a separate 1,001-item folder, an unsent singleton whose `GetConversation()` result is null, two inert attachments, a body longer than 1,024 characters, and `protectedMessage: null`. Seeder mutations are confined to one synthetic item in the dedicated profile's bootstrap Inbox and synthetic items in the exact two secure run-directory PSTs. It emits ACL-protected local fixture/inventory/status files atomically and has a separate exact PST detach action. Detach removes the two fixture PSTs and run artifacts, but deliberately retains the dedicated profile, bootstrap store, and bootstrap-Inbox message. The regular live runner uses the fixture files only as expected values; all protocol locators are discovered at runtime and remain in memory.

The Phase 4 live run is paused before acceptance. Do not mark the phase complete or populate live-result values in [Phase 4 evidence](PHASE_4_EVIDENCE.md) until the runner and `verify-phase4-smoke.ps1` accept all three graceful Outlook cycles, the native Codex call, resource/COM checks, privacy scan, shutdown, port release, and temporary registration/signing cleanup. Fixture PST detachment remains the separate post-gate action documented in the Phase 4 smoke instructions.

## Codex configuration checks

The repository-scoped client configuration is validated with:

```powershell
.\tools\configure-codex.ps1 -Action Validate
```

Validation checks the committed fail-closed TOML, absence of a shadowing user-scoped registration, strict Codex configuration parsing, and the effective exact Streamable HTTP registration. `-Action Install` provisions a current-user token; `-Action Rotate` replaces it; `-Action ClearToken` removes it. Outlook and Codex must restart after any token change because both authenticate with values inherited into their processes.

Configuration validation alone is not a live endpoint acceptance result. Phase 3 acceptance therefore ran the actual Codex CLI against the authenticated endpoint while Outlook was running and deterministically validated its native structured probe result.

## Sensitive data

Tests and CI must not upload mailbox fixtures, message content, Outlook identifiers, bearer tokens, user configuration, runtime logs, or release signing keys. Phase 4 evidence retains only aggregate non-content numeric values and booleans from the final runner; it must not retain display names, profile values, folder paths, markers, subjects, bodies, recipients, attachment names, locators, cursors, operation IDs, raw protocol responses, or Codex event streams.
