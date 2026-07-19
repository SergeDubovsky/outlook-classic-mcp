# Phase 3 evidence

Status: complete
Last updated: 2026-07-19

## Scope

Phase 3 adds the first Outlook Object Model tool, `outlook_probe`, through the
authenticated HTTP and MCP boundary, the bounded dispatcher, the captured
Outlook UI STA, and an immutable managed result. The probe reports bounded host,
dispatcher, store-type, store-capability, and standard-folder availability
metadata. It does not read messages, bodies, attachments, or store identifiers.

The exposed tool allowlist is exactly `outlook_status` and `outlook_probe`.
Consequential, draft, message-read, attachment-transfer, delete, and send tools
remain absent.

## Build and automated tests

Full Visual Studio CLI MSBuild completed the Debug and Release solution builds.
The Phase 3 suite reported:

| Assembly target | Passed | Failed |
|---|---:|---:|
| Core Tests `net10.0` | 75 | 0 |
| Core Tests `net48` | 97 | 0 |
| Transport Tests `net48` | 141 | 0 |

Coverage includes immutable bounded contracts, tool exposure, gateway success
and sanitized failure envelopes, dispatcher serialization, queue saturation,
cancellation and deadlines, shutdown quiescence, real-listener recovery, and
pinned-client calls. Core and Transport tests do not load the VSTO assembly.

## Live Outlook and Codex acceptance

The repeatable acceptance command is:

```powershell
.\tools\run-phase3-smoke.ps1 `
    -Profile '<dedicated test profile>' `
    -ExpectedStoreInventoryPath '<independently prepared inventory file>'
```

The final Release gate completed three normal Outlook start/close cycles. It
compared the probe with an independently prepared inventory, exercised the raw
endpoint sequentially and concurrently, and invoked the native MCP tool once
with `codex-cli` 0.144.6.

| Check | Result |
|---|---:|
| Outlook cycles | 3 |
| Distinct diagnostic sessions | 3 |
| Endpoint verifications | 3 |
| Discovered tools | 2 |
| Status calls | 1 |
| Sequential probes | 20 |
| Concurrent probes | 4 |
| Post-restart probes | 2 |
| Readiness retries | 0 |
| Raw endpoint probes | 26 |
| Native Codex probes | 1 |
| Total probes | 27 |
| Structured envelopes | 28 |
| Unique operation IDs | 28 |
| Configured stores | 7 |
| Independent inventory match | yes |
| UI STA proof | passed |
| Outlook responsiveness checks | 7 |
| Lifecycle events | 18 |
| Target Event ID 45 records | 3 |
| Target Event ID 59 records | 0 |
| Registration, certificate, listener, and process cleanup | passed |

The recorded non-content fingerprints were:

| Evidence set | SHA-256 |
|---|---|
| Store inventory | `dc4daed008c2729b72550a18f3d761ccfd4a4e2971fb40d5168e728299c90c19` |
| Store metadata | `6fd132c8dad644d37598d7dfc630ba4bf13a60b69d558ae39f3f45d944d70d62` |
| Environment | `6d9e1a992b7c5e2c5d7a3a3b89f3444c57a1d3f3aafa9e7d1040ae5baad3cbda` |
| Runtime identity | `65f5a0725589215872e39f708abc09c1b5e9aaabb1d1563886efac1a3d1fc52f` |

No store labels, Outlook profile value, operation identifiers, bearer value,
message content, or machine-specific account data are retained in this record.

## Bounded readiness semantics

Only the first probe after each listener start may use readiness polling. The
gate permits at most five attempts, spaces retries by 500 ms, and enforces a
30-second total deadline. A retry is allowed only for the exact bounded partial
shape that reports temporarily incomplete store metadata. Once complete, that
first result is reused as workload evidence; every later probe is single-shot
and fail-closed. The accepted run required no retries.

Archive-folder availability remains `unknown` with the fixed warning that the
Outlook Object Model does not expose it. That expected limitation makes the
successful probe structurally partial even when the complete store inventory
is present and matched.

## Phase 3 completion

Every Phase 3 acceptance gate passed: native Codex invocation, verified UI-STA
execution, independent store inventory agreement, sequential and concurrent
stability, Outlook responsiveness, clean shutdown, and three restart cycles.
Phase 4 can add bounded read-only mail access without changing the accepted
brokerless process boundary.
