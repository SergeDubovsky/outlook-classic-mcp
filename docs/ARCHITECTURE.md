# Architecture

The authoritative architecture and its acceptance gates are in `IMPLEMENTATION_PLAN.md`. This document summarizes the boundaries implemented and accepted through Phase 2. The live results are recorded in [Phase 2 evidence](PHASE_2_EVIDENCE.md).

## Process model

All production assemblies load in the Classic Outlook process. There is no broker or background service.

```text
Codex MCP client
    -> authenticated http://127.0.0.1:8765/mcp/
    -> OutlookClassicMcp.Transport
       -> initialize / initialized / ping / tools/list
       -> outlook_status (host state, listener readiness, version)

Future mailbox tools, not exposed in Phase 2:
    OutlookClassicMcp.Transport
    -> OutlookClassicMcp.Core policy and typed tools
    -> one bounded Outlook UI-STA dispatcher
    -> Outlook Object Model
    -> immutable managed DTOs
```

## Assembly ownership

- **AddIn** owns VSTO composition, lifecycle, STA dispatch, Outlook COM access, and metadata-only diagnostics.
- **Transport** owns `HttpListener`, authentication, HTTP validation, limits, and MCP adaptation. It contains no Office or VSTO references.
- **Core** owns immutable contracts, policy, validation, and tool orchestration. It contains no Office, VSTO, WinForms, registry, or COM references.

## Phase 2 transport

The add-in creates one `HttpListener` at the literal prefix `http://127.0.0.1:8765/mcp/`. It loads one canonical 32-byte base64url bearer token from the Outlook process environment before binding. `tools/configure-codex.ps1` provisions that variable at current-user scope, but an already-running Outlook process does not see a new value. The listener therefore authenticates against a process-snapshotted token until Outlook restarts. Missing or invalid configuration fails closed and leaves the add-in loaded in `Degraded` state without a listener.

Each validated HTTP POST creates an isolated, stateless server session using the pinned `ModelContextProtocol.Core` 1.4.1 `StreamableHttpServerTransport`. Session identifiers are neither issued nor accepted. The available protocol operations are initialize, `notifications/initialized`, ping, `tools/list`, and `tools/call` restricted to `outlook_status`. The status provider reads managed lifecycle state only; no Outlook Object Model call or STA dispatch occurs.

## HTTP policy

Validation occurs before MCP parsing and fails at the first rejected boundary:

- the remote address, `Host`, raw route, and method must be exactly `127.0.0.1`, `127.0.0.1:8765`, `/mcp/`, and uppercase `POST`;
- any `Origin` header is rejected, and exactly one canonical `Bearer` authorization value is required;
- `Content-Type` must be one `application/json` value; optional parameters are accepted, but a charset must be absent or UTF-8. `Accept` must include both `application/json` and `text/event-stream` at nonzero quality;
- when supplied, the protocol header must be exactly one of `2024-11-05`, `2025-03-26`, `2025-06-18`, or `2025-11-25`; content encoding is absent or `identity`, and `Mcp-Session-Id` is forbidden;
- requests are limited to 64 header fields, 16 KiB of aggregate header characters, 8 KiB per header value, and a 1 MiB body;
- at most four handlers run concurrently, excess work receives `503` with `Retry-After: 1`, and each handler has a 15-second deadline.

MCP responses use `text/event-stream`, `Cache-Control: no-cache, no-store`, and identity encoding. An accepted notification returns an empty `202` response. Boundary rejections use one fixed generic JSON error; malformed JSON, unsupported batches, and pre-response internal failures use separate fixed bounded errors. None disclose capabilities, tokens, request bodies, or mailbox data.

## Current phase

Phase 2 is complete. Automated transport tests, three live Outlook restart/endpoint cycles, clean port release, and an actual Codex CLI `outlook_status` call passed. Phase 3 will add the first Outlook/profile/store metadata probe through the STA dispatcher; no mailbox tool is currently exposed.
