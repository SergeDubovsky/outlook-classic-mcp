# Testing

## Automated tests

`OutlookClassicMcp.Core.Tests` targets both `net10.0` and `net48`. `OutlookClassicMcp.Transport.Tests` targets `net48` and uses the real Windows `HttpListener` with no Outlook dependency.

Run:

```powershell
.\tools\preflight.ps1
.\tools\build.ps1
```

The build script performs locked restore, builds Debug and Release/Any CPU with full Visual Studio MSBuild, and runs every test target. Normal test runners must never instantiate or load the VSTO add-in.

## Interactive Outlook smoke tests

Interactive tests use a dedicated Outlook profile and disposable data. Save work and close Outlook gracefully before registration or F5 tests; never terminate `OUTLOOK.EXE` forcibly.

Phase 0 requires one stock-template F5 proof that reaches `ThisAddIn_Startup`, plus confirmation that Outlook lists the add-in as active. Later phases add restart, listener, store-probe, bounded-read, mutation, and approval-gated send checks in the order defined by `IMPLEMENTATION_PLAN.md`.

## Sensitive data

Tests and CI must not upload mailbox fixtures, message content, Outlook identifiers, bearer tokens, user configuration, runtime logs, or release signing keys.
