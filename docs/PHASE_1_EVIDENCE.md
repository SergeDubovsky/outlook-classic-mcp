# Phase 1 evidence

Status: complete
Last updated: 2026-07-19

## Scope

Phase 1 establishes the add-in host lifecycle and Outlook STA dispatch boundary. It does not start the MCP listener, access mailbox data, or expose tools.

The production topology remains brokerless: AddIn, Transport, and Core load inside `OUTLOOK.EXE`. Startup uses the VSTO-provided `ThisAddIn.Application`; it never creates an `Outlook.Application` instance.

## Host implementation

The add-in now provides:

- An explicit lifecycle state machine covering startup, online, degraded, pausing, paused, stopping, and stopped states.
- A bounded dispatcher whose managed queue is woken only when work exists by a private message to a hidden UI-owned control on the captured Outlook STA.
- A whole-drain guard that prevents nested Windows message pumps from executing Outlook operations reentrantly.
- Fail-closed shutdown ownership, a bounded shutdown watchdog, and guaranteed terminal shutdown completion.
- Metadata-only diagnostics with protected current-user ACLs, bounded rotation, and no mailbox content or Outlook object identifiers.
- A runtime dependency probe that loads and activates the pinned MCP dependency closure without opening the listener.

The Core test assembly links the Office-free lifecycle and dispatcher sources directly. It does not load the VSTO assembly. Dispatcher regressions deliberately pump nested UI messages, prove that a second operation remains queued until the first completes on the same managed/native STA, and verify capacity, cancellation, idempotent shutdown, `HOST_STOPPING` completion of queued work, active-work settlement, terminal completion of every accepted task, and fail-closed native wake-up loss. [ADR 0001](ADR_0001_STA_DISPATCH_WAKEUP.md) records why live rapid-restart evidence replaced `Control.BeginInvoke` with the private-message wake-up.

## CLI lifecycle gate

The repeatable acceptance command is:

```powershell
.\tools\run-phase1-smoke.ps1 -Profile Outlook
```

The runner uses full Visual Studio only through `MSBuild.exe` and the OfficeTools targets. It does not launch the Visual Studio IDE or use Outlook COM automation. It:

1. Builds and tests Release with locked restore.
2. Creates a one-day, non-exportable current-user signing certificate.
3. Generates signed VSTO manifests, creates the local inclusion entry, and registers the exact Release manifest.
4. Starts Classic Outlook with the named profile for three separate cycles.
5. Correlates metadata diagnostics and Outlook Application Event ID 45 to each owned process.
6. Requests normal closure through `Process.CloseMainWindow()` and never force-terminates Outlook.
7. Rejects target Event ID 59 records and changes to `DisabledItems` or `CrashingAddinList`.
8. Runs the bounded metadata verifier while registration still exists.
9. Invokes `VSTOClean`, removes the certificate/private key, and verifies that no process, listener, registration, inclusion entry, certificate, or key residue remains.

The runner reads diagnostics with concurrent read/write sharing so its polling cannot block the next Outlook process from initializing the diagnostics file. Event Log queries use local `DateTime` boundaries together with global record-ID watermarks and exact process correlation, preventing stale evidence from satisfying a cycle.

## Acceptance result

The final three-cycle Release run completed successfully on 2026-07-19:

| Check | Result |
|---|---:|
| Distinct Outlook processes | 3 |
| Distinct diagnostics sessions | 3 |
| Target Event ID 45 records with `Load Behavior: 3` | 3 |
| Target Event ID 59 records | 0 |
| Maximum startup callback | 19.4868 ms |
| MCP runtime assemblies loaded | 17 |
| Same-STA dispatcher probes | 3 |
| Quiescent shutdowns | 3 |
| Resilience state changes | 0 |
| Residual Outlook processes/listeners | 0 |
| Residual VSTO/certificate/key state | 0 |

The pinned runtime identity fingerprint was `65F5A0725589215872E39F708ABC09C1B5E9AAABB1D1563886EFAC1A3D1FC52F` in every cycle. The listener remained inactive throughout Phase 1.

## Build and automated tests

Full Visual Studio MSBuild completed locked Debug and Release builds after the lifecycle gate. Results for each configuration:

| Assembly target | Passed | Failed |
|---|---:|---:|
| Core Tests `net10.0` | 12 | 0 |
| Core Tests `net48` | 15 | 0 |
| Transport Tests `net48` | 5 | 0 |

Generated VSTO files, dependency inventory, runtime asset closure, package locks, and license inventories pass their repository verifiers. Routine builds and the lifecycle runner both leave the machine free of development registration and ephemeral signing material.

## Phase 1 completion

All Phase 1 acceptance gates are complete. Phase 2 can add token provisioning and the authenticated loopback listener without weakening the established Outlook STA, shutdown, or cleanup boundaries.
