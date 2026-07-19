# Testing

## Automated tests

`OutlookClassicMcp.Core.Tests` targets both `net10.0` and `net48`. `OutlookClassicMcp.Transport.Tests` targets `net48` and uses the real Windows `HttpListener` with no Outlook dependency.

Run:

```powershell
.\tools\preflight.ps1
.\tools\build.ps1
```

The build script performs locked restore, builds Debug and Release/Any CPU with full Visual Studio MSBuild, and runs every test target. Normal test runners must never instantiate or load the VSTO add-in.

## Outlook smoke tests

Outlook smoke tests use a dedicated profile and disposable data on a logged-on interactive Windows desktop. Save work and close Outlook gracefully before registration or F5 tests; never terminate `OUTLOOK.EXE` forcibly.

Phase 0 requires one stock-template F5 proof that reaches `ThisAddIn_Startup`, plus confirmation that Outlook lists the add-in as active. Later phases add restart, listener, store-probe, bounded-read, mutation, and approval-gated send checks in the order defined by `IMPLEMENTATION_PLAN.md`.

Run the Phase 1 lifecycle gate from PowerShell while Outlook is closed:

```powershell
.\tools\run-phase1-smoke.ps1 -Profile Outlook
```

The runner uses full Visual Studio only through `MSBuild.exe` and the OfficeTools targets; it does not launch the IDE. It creates a temporary Release VSTO registration and certificate, performs three normal Outlook start/close cycles, and then cleans up the temporary registration, certificate, and private-key container. Outlook is closed through its normal main-window path and is never force-terminated.

For each cycle, the runner requires Outlook Event ID 45 to show `OutlookClassicMcp.AddIn` loaded with `LoadBehavior=3` and rejects Event ID 59 evidence that the add-in was disabled. It then invokes the metadata verifier, which requires three successful dependency-load and same-STA dispatcher probes, a sub-500 ms startup callback, quiescent shutdown with no residual listener, protected diagnostics ACLs, and the exact Release registration with `LoadBehavior=3`.

Checking **Active Application Add-ins** and **Disabled Items** in the Outlook UI is optional troubleshooting or release-audit evidence, not part of the routine Phase 1 gate.

The verifier ignores well-formed records outside the bounded run but fails closed on malformed retained diagnostics. If a prior interrupted run left a malformed line, explicitly archive the metadata-only log files before recording a new start time.

## Sensitive data

Tests and CI must not upload mailbox fixtures, message content, Outlook identifiers, bearer tokens, user configuration, runtime logs, or release signing keys.
