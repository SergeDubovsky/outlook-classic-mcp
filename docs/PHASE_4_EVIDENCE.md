# Phase 4 evidence

Status: draft — acceptance pending
Last updated: 2026-07-19

## Scope

Phase 4 adds bounded read-only Outlook access without changing the accepted brokerless process boundary. The exact exposed tool order is:

1. `outlook_status`
2. `outlook_probe`
3. `outlook_list_mailboxes`
4. `outlook_list_folders`
5. `outlook_list_messages`
6. `outlook_search_messages`
7. `outlook_get_message`
8. `outlook_get_conversation`
9. `outlook_list_attachments`

All nine tools are read-only, non-destructive, idempotent, and closed-world. Message bodies are available only through the bounded `outlook_get_message` contract. Attachment content, drafts, mutations, delete, and send remain absent.

This document does not complete the Phase 4 acceptance gate. The final Release build/tests pass, but every live Outlook/Codex result is still pending.

## Final build and automated checkpoint

The final 2026-07-19 full Release checkpoint reported:

| Assembly target | Passed | Failed |
|---|---:|---:|
| Core Tests `net10.0` | 123 | 0 |
| Core Tests `net48` | 170 | 0 |
| Transport Tests `net48` | 211 | 0 |
| AddIn Release build | succeeded | — |

This checkpoint includes the final conditional-seeder and smoke-harness changes. It establishes the automated baseline but does not replace the live acceptance gate.

Automated coverage includes exact tool exposure and schemas, store-qualified locators, request/result bounds, HMAC cursor derivation and validation, canonical query binding, stale-anchor handling, deterministic keyset pagination, partial cross-scope behavior, fixed internal filters, Search Folder rejection, UI-STA assertions, explicit COM ownership telemetry, bounded materialization, body/recipient protection and truncation, conversation bounds, attachment metadata, error sanitization, real-listener recovery, and Core/Transport isolation from VSTO.

## Conditional live fixture boundary

The conditional fixture is compiled only when `OutlookMcpSmokeSeeder=true`. It runs in `OUTLOOK.EXE` from `ThisAddIn.Application` on the captured Outlook UI STA and does not start the normal MCP host in the seeding process. It requires a dedicated accountless profile and a secure per-run directory containing only synthetic disposable artifacts.

The generated fixture covers the complete three-store profile:

- `bootstrap_store` maps to the profile's real bootstrap default store. It has `expectsInbox: true`, and its known synthetic item is seeded and later resolved through the standard Inbox.
- `fixture_store_a` and `fixture_store_b` are disposable attached PST stores with run-specific display names. Each has `expectsInbox: false` in the fixture so its known message is resolved through the declared custom `folderPath`; this does not assert that Outlook created no Inbox in the PST.

The same fixture includes 12 static items whose recorded inventory captures Outlook's actual ordering—`ReceivedTime` when usable, followed by a deterministic `EntryID` tie-break—for three pages of four, a separate 1,001-item folder, a null-conversation singleton fallback, two inert attachment records, a body longer than 1,024 characters, and `protectedMessage: null`. Mutations are confined to one synthetic item in the dedicated profile's bootstrap Inbox and synthetic items in the exact two secure run-directory PSTs. The seeder applies and verifies current-user/SYSTEM-only ACLs on its files, writes local JSON atomically, and provides a separate exact PST detach action. Detach removes the two fixture PSTs and run artifacts but deliberately retains the dedicated profile, bootstrap store, and bootstrap-Inbox message. Fixture JSON, status, Outlook identifiers, and message markers remain ignored local artifacts and are not evidence payloads.

## Pending live acceptance

The repeatable live flow is:

```powershell
.\tools\seed-phase4-fixture.ps1 -Action Seed -Profile OutlookMcpTest
.\tools\run-phase4-smoke.ps1 `
    -Profile OutlookMcpTest `
    -ExpectedStoreInventoryPath '<generated store-inventory.local.json>' `
    -ReadFixturePath '<generated read-fixture.local.json>'
.\tools\seed-phase4-fixture.ps1 `
    -Action Detach `
    -Profile OutlookMcpTest `
    -RunDirectory '<same secure run directory>'
```

The gate must complete one full read workload followed by two restart workloads, all through normal Outlook start/close behavior. It must prove exact tool discovery, complete inventory agreement, known-message reacquisition in multiple stores, deterministic pagination, the large-folder bound, partial search, cursor rejection, conversation and attachment metadata, body truncation, optional protected-body semantics, cancellation recovery, concurrent and repeated reads, balanced COM telemetry, bounded materialization, process-resource stability, native Codex access, UI-STA execution, responsiveness, privacy scanning, clean shutdown, port release, and cleanup.

A live structured timeout is not required; `StructuredTimeoutObserved: 0` is valid because deterministic timeout behavior is covered by the real-`HttpListener` automated test. The conditional fixture intentionally has `protectedMessage: null`, so `ProtectedFixtureVerified: false` is also expected and does not fail this gate.

All retained live-result values are pending:

| Privacy-safe aggregate field | Value |
|---|---|
| `Phase` | pending |
| `VerifiedCycleCount` | pending |
| `DistinctOutlookProcessCount` | pending |
| `EndpointVerificationCount` | pending |
| `FullEndpointWorkloadCount` | pending |
| `ToolCount` | pending |
| `StoreCount` | pending |
| `StoreInventoryMatched` | pending |
| `KnownMessageStoreCount` | pending |
| `StaticPageCount` | pending |
| `StaticPageItemCount` | pending |
| `LargeFolderMinimumItemCount` | pending |
| `PartialSearchVerified` | pending |
| `CancellationRecoveryVerified` | pending |
| `StructuredTimeoutObserved` | pending |
| `RepeatedReadCount` | pending |
| `DirectComTelemetryBalanced` | pending |
| `MaterializedItemHighWater` | pending |
| `ResourceStable` | pending |
| `OutlookResponsive` | pending |
| `ProtectedFixtureVerified` | pending |
| `NativeCodexVerified` | pending |
| `UniqueOperationIdCount` | pending |
| `PrivacySentinelCount` | pending |
| `DiagnosticsLogFileCount` | pending |
| `MaximumPortReleaseMilliseconds` | pending |
| `OutlookStopped` | pending |
| `PortReleased` | pending |
| `CleanupVerified` | pending |

Only the final aggregate object emitted by `run-phase4-smoke.ps1` may populate this table. `CleanupVerified` covers the runner's temporary VSTO registration, certificate, and key cleanup; it does not mean the fixture PSTs were detached. Do not retain the verifier's intermediate endpoint or Codex objects, store-inventory fingerprint, operation identifiers, privacy sentinel values, profile or store labels, folder paths, message or attachment markers, locators, cursors, raw protocol responses, Codex event streams, or runtime logs.

## Completion rule

Phase 4 remains pending until `verify-phase4-smoke.ps1` accepts the complete live run. Only then may this status change to complete and the implementation-plan gate be updated separately.
