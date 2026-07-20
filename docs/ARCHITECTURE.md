# Architecture

The authoritative architecture and its acceptance gates are in `IMPLEMENTATION_PLAN.md`. This document summarizes the boundaries accepted through Phase 3 and the Phase 4 bounded-read implementation awaiting final live acceptance. Accepted Phase 3 results are in [Phase 3 evidence](PHASE_3_EVIDENCE.md); the deliberately incomplete Phase 4 record is in [Phase 4 evidence](PHASE_4_EVIDENCE.md).

## Process model

All production assemblies load in the Classic Outlook process. There is no broker or background service.

```text
Codex MCP client
    -> authenticated http://127.0.0.1:8765/mcp/
    -> OutlookClassicMcp.Transport
       -> initialize / initialized / ping / tools/list
       -> outlook_status (host state, listener readiness, version, read counters)
       -> outlook_probe
       -> outlook_list_mailboxes / outlook_list_folders
       -> outlook_list_messages / outlook_search_messages
       -> outlook_get_message / outlook_get_conversation
       -> outlook_list_attachments
          -> OutlookClassicMcp.Core bounded contracts and policy
          -> one bounded Outlook UI-STA dispatcher
          -> Outlook Object Model metadata or explicitly requested bounded body
          -> immutable managed DTOs

Draft, mutation, attachment-content export, delete, and send tools are absent.
```

## Assembly ownership

- **AddIn** owns VSTO composition, lifecycle, STA dispatch, Outlook COM access, and metadata-only diagnostics.
- **Transport** owns `HttpListener`, authentication, HTTP validation, limits, and MCP adaptation. It contains no Office or VSTO references.
- **Core** owns immutable contracts, policy, validation, and tool orchestration. It contains no Office, VSTO, WinForms, registry, or COM references.

## Authenticated transport

The add-in creates one `HttpListener` at the literal prefix `http://127.0.0.1:8765/mcp/`. It loads one canonical 32-byte base64url bearer token from the Outlook process environment before binding. `tools/configure-codex.ps1` provisions that variable at current-user scope, but an already-running Outlook process does not see a new value. The listener therefore authenticates against a process-snapshotted token until Outlook restarts. Missing or invalid configuration fails closed and leaves the add-in loaded in `Degraded` state without a listener.

Each validated HTTP POST creates an isolated, stateless server session using the pinned `ModelContextProtocol.Core` 1.4.1 `StreamableHttpServerTransport`. Session identifiers are neither issued nor accepted. The available protocol operations are initialize, `notifications/initialized`, ping, `tools/list`, and `tools/call` restricted to this exact ordered allowlist: `outlook_status`, `outlook_probe`, `outlook_list_mailboxes`, `outlook_list_folders`, `outlook_list_messages`, `outlook_search_messages`, `outlook_get_message`, `outlook_get_conversation`, and `outlook_list_attachments`. The status provider reads managed lifecycle state and read counters only; no Outlook Object Model call or STA dispatch occurs.

## Outlook UI-STA and COM boundary

`outlook_probe` is the first tool to cross the Outlook boundary. Transport applies a 14-second tool deadline, then calls the injected Core gateway interface. The AddIn gateway enqueues one static synchronous operation onto the capacity-16 dispatcher. A private window message wakes Outlook's captured UI thread; the operation verifies both managed and native thread identity and the `STA` apartment before accessing the Outlook Object Model.

Every Outlook operation uses `ThisAddIn.Application` and its host-owned Session; it never creates an Outlook Application. Each delegate verifies the captured managed thread, native thread, and `STA` apartment before touching Outlook. Store, folder, item, recipient, attachment, property-accessor, and collection RCWs acquired by an operation are released in reverse ownership order after scalar data is copied. The host-owned Application and Session are never released. No COM object enters Core, Transport, the queue payload, a continuation, a cache, or an MCP result.

The probe returns bounded immutable data for host version/bitness, the active profile label, dispatcher proof, configured store count, at most 64 store labels/types/capability flags, and tri-state standard-folder availability. Archive availability is always `unknown` with a fixed warning because the Outlook Object Model does not expose it. Message content, attachments, store identifiers, and folder objects are not read or returned.

Only one Outlook operation runs at a time. Cancellation can remove queued work but cannot abort COM work after it starts. Shutdown rejects new work, completes queued work with a bounded stopping failure, permits an active Outlook operation to finish, and defers final dispatcher disposal until quiescence. A dispatcher wake-up failure fails pending work and degrades the host instead of silently stranding requests.

## Phase 4 bounded reads

All Phase 4 inputs are closed typed objects. Mailboxes are addressed by `StoreID`; folders by `StoreID + EntryID`; messages by `StoreID + EntryID + ItemClass`. Every call reacquires the referenced Outlook object and verifies its store, parent, and item class as applicable. Display names are output metadata and are never a lookup fallback.

Page sizes are 1 through 50, defaulting to 25. Message bodies are returned only by `outlook_get_message`, with a caller-selected plain-text or HTML representation and an absolute 50,000-character cap. Message summaries, conversations, and attachment listing never return body or attachment bytes. Recipients, warnings, search scopes, materialized candidates, and serialized tool results all have independent bounds; a tool result cannot exceed 1 MiB.

Continuation cursors are opaque, versioned payloads authenticated with HMAC-SHA-256 under a domain-separated key derived from the process-snapshotted bearer token. A cursor binds its tool kind, canonical query hash, and keyset anchor. Mailbox and folder pages anchor on their deterministic name/identifier order; message and conversation pages anchor on descending effective timestamp plus the complete store-qualified item identity; attachment pages anchor on attachment index plus a metadata fingerprint. Tampering, using a cursor with another tool or query, or exceeding the cursor bound fails before Outlook dispatch. If the authenticated anchor can no longer be reacquired exactly, the tool returns `CURSOR_STALE` rather than falling back by name.

Structured search accepts at most 64 explicit scopes and closed filters for sender, recipient, subject, text, received range, unread state, category, and attachment presence. Scopes are canonicalized by store and folder identifiers, executed sequentially through the single STA dispatcher, and merged outside COM into at most `pageSize + 1` candidates with exact store-qualified deduplication. If at least one scope succeeds, supported scope-local failures are returned with `partial=true`; a partial cross-scope page has no global continuation cursor. If every scope fails, the first failure in canonical scope order becomes the tool error.

Visible Outlook Search Folders are rejected for message listing and explicitly folder-scoped search. The gateway enumerates at most 1,024 entries from `Store.GetSearchFolders()` and compares entry identifiers with `NameSpace.CompareEntryIDs`. A match, unavailable search-folder metadata, or an unsupported provider fails closed; an over-limit enumeration returns a timeout. Default-Inbox search remains allowed because it does not accept a caller-selected Search Folder.

Phase 4 status adds only five non-content counters: acquired, released, outstanding, and peak operation-owned COM references, plus the materialized-item high-water mark. These counters support leak and page-bound evidence without exposing profile labels, store names, folder names, item identifiers, cursors, subjects, bodies, recipients, or attachment metadata.

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

Phase 3 remains complete. The Phase 4 implementation and final Release build/tests are present, but the three-cycle live Outlook/Codex run is paused. Phase 4 must not be marked complete until the live gate proves exact nine-tool discovery, fixture reads across the complete inventory, deterministic pagination, large-folder bounds, partial search, cursor rejection, cancellation recovery, balanced COM telemetry, resource stability, native Codex access, Outlook responsiveness, graceful shutdown, port release, privacy scanning, and temporary registration/signing cleanup. Fixture PST detachment is a separate post-gate action. See [Phase 4 evidence](PHASE_4_EVIDENCE.md).
