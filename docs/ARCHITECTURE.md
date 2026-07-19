# Architecture

The authoritative architecture and its acceptance gates are in `IMPLEMENTATION_PLAN.md`. This document summarizes the boundaries implemented and accepted through Phase 3. The live results are recorded in [Phase 3 evidence](PHASE_3_EVIDENCE.md).

## Process model

All production assemblies load in the Classic Outlook process. There is no broker or background service.

```text
Codex MCP client
    -> authenticated http://127.0.0.1:8765/mcp/
    -> OutlookClassicMcp.Transport
       -> initialize / initialized / ping / tools/list
       -> outlook_status (host state, listener readiness, version)
       -> outlook_probe
          -> OutlookClassicMcp.Core bounded contracts and policy
          -> one bounded Outlook UI-STA dispatcher
          -> Outlook Object Model store metadata
          -> immutable managed DTOs

Future message and write tools are not exposed in Phase 3.
```

## Assembly ownership

- **AddIn** owns VSTO composition, lifecycle, STA dispatch, Outlook COM access, and metadata-only diagnostics.
- **Transport** owns `HttpListener`, authentication, HTTP validation, limits, and MCP adaptation. It contains no Office or VSTO references.
- **Core** owns immutable contracts, policy, validation, and tool orchestration. It contains no Office, VSTO, WinForms, registry, or COM references.

## Authenticated transport

The add-in creates one `HttpListener` at the literal prefix `http://127.0.0.1:8765/mcp/`. It loads one canonical 32-byte base64url bearer token from the Outlook process environment before binding. `tools/configure-codex.ps1` provisions that variable at current-user scope, but an already-running Outlook process does not see a new value. The listener therefore authenticates against a process-snapshotted token until Outlook restarts. Missing or invalid configuration fails closed and leaves the add-in loaded in `Degraded` state without a listener.

Each validated HTTP POST creates an isolated, stateless server session using the pinned `ModelContextProtocol.Core` 1.4.1 `StreamableHttpServerTransport`. Session identifiers are neither issued nor accepted. The available protocol operations are initialize, `notifications/initialized`, ping, `tools/list`, and `tools/call` restricted to the ordered allowlist `outlook_status`, `outlook_probe`. The status provider reads managed lifecycle state only; no Outlook Object Model call or STA dispatch occurs.

## Phase 3 Outlook boundary

`outlook_probe` is the first tool to cross the Outlook boundary. Transport applies a 14-second tool deadline, then calls the injected Core gateway interface. The AddIn gateway enqueues one static synchronous operation onto the capacity-16 dispatcher. A private window message wakes Outlook's captured UI thread; the operation verifies both managed and native thread identity and the `STA` apartment before accessing the Outlook Object Model.

The operation uses `ThisAddIn.Application` and its host-owned Session; it never creates an Outlook Application. Store collections, stores, and standard-folder objects are operation-owned and released in reverse scope after their scalar values are copied. No COM object enters Core, Transport, the queue payload, a continuation, a cache, or an MCP result.

The probe returns bounded immutable data for host version/bitness, the active profile label, dispatcher proof, configured store count, at most 64 store labels/types/capability flags, and tri-state standard-folder availability. Archive availability is always `unknown` with a fixed warning because the Outlook Object Model does not expose it. Message content, attachments, store identifiers, and folder objects are not read or returned.

Only one Outlook operation runs at a time. Cancellation can remove queued work but cannot abort COM work after it starts. Shutdown rejects new work, completes queued work with a bounded stopping failure, permits an active Outlook operation to finish, and defers final dispatcher disposal until quiescence. A dispatcher wake-up failure fails pending work and degrades the host instead of silently stranding requests.

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

Phase 3 is complete. Automated Core and Transport tests, independent store-inventory agreement, verified UI-STA execution, twenty sequential and four concurrent probes, one native Codex probe, three live Outlook cycles, responsiveness checks, and clean port release passed. Phase 4 will add bounded read-only message access; no message-read or write tool is currently exposed.
