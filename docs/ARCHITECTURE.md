# Architecture

The authoritative architecture and its acceptance gates are in `IMPLEMENTATION_PLAN.md`. This document summarizes the implemented boundaries and will be expanded as each phase lands.

## Process model

All production assemblies load in the Classic Outlook process. There is no broker or background service.

```text
Codex MCP client
    -> authenticated http://127.0.0.1:8765/mcp/
    -> OutlookClassicMcp.Transport
    -> OutlookClassicMcp.Core policy and tools
    -> one bounded Outlook UI-STA dispatcher
    -> Outlook Object Model
    -> immutable managed DTOs
```

## Assembly ownership

- **AddIn** owns VSTO composition, lifecycle, STA dispatch, Outlook COM access, and metadata-only diagnostics.
- **Transport** owns `HttpListener`, authentication, HTTP validation, limits, and MCP adaptation. It contains no Office or VSTO references.
- **Core** owns immutable contracts, policy, validation, and tool orchestration. It contains no Office, VSTO, WinForms, registry, or COM references.

## Current phase

Phase 0 is proving the stock VSTO template, full-MSBuild toolchain, isolated Core/Transport projects, exact loopback namespace, locked restore, and tests. It does not expose mailbox data.
