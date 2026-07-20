# Tool contracts

## Phase 4 surface

The current implementation exposes exactly nine tools, in this order:

1. `outlook_status`
2. `outlook_probe`
3. `outlook_list_mailboxes`
4. `outlook_list_folders`
5. `outlook_list_messages`
6. `outlook_search_messages`
7. `outlook_get_message`
8. `outlook_get_conversation`
9. `outlook_list_attachments`

The protocol surface supports initialize, `notifications/initialized`, ping, `tools/list`, and `tools/call` only for this allowlist. Every descriptor is read-only, non-destructive, idempotent, and closed-world. There are no resources, prompts, completions, subscriptions, draft, mutation, attachment-content transfer, delete, or send tools.

The first two tools passed Phase 3 raw endpoint and live Codex acceptance. The seven bounded-read tools are implemented and the final Release build/tests pass; live Phase 4 acceptance remains pending. See [Phase 3 evidence](PHASE_3_EVIDENCE.md) and [the Phase 4 draft evidence](PHASE_4_EVIDENCE.md).

### `outlook_status`

`outlook_status` takes no arguments: its arguments member may be omitted or supplied as an empty object, and any supplied key produces an error. Its descriptor is read-only, non-destructive, idempotent, and closed-world. It reads managed add-in lifecycle state only and never enumerates or accesses Outlook profiles, stores, folders, messages, bodies, or attachments.

A successful call returns bounded structured content in this shape:

```json
{
  "ok": true,
  "operationId": "32 lowercase hexadecimal characters",
  "data": {
    "hostState": "online",
    "listenerReady": true,
    "version": "1.0.0.0",
    "readDiagnostics": {
      "comAcquired": 0,
      "comReleased": 0,
      "comOutstanding": 0,
      "comPeak": 0,
      "materializedItemHighWater": 0
    }
  },
  "partial": false,
  "warnings": []
}
```

`hostState` is limited to 32 characters and `version` to 64 characters. In Phase 4, `readDiagnostics` contains only nonnegative process-lifetime counters for operation-owned COM acquisition/release/outstanding/peak and managed materialization high-water. It contains no profile, mailbox, folder, item, cursor, or content value. The text content contains only the bounded host-state summary. Arguments produce a typed `INVALID_ARGUMENT` tool error without calling the provider. An unavailable or failing provider produces a bounded, retryable `OUTLOOK_NOT_READY` tool error. Unknown tool names receive a protocol error and are not reflected in an unbounded response.

### `outlook_probe`

`outlook_probe` takes no arguments: its arguments member may be omitted or supplied as an empty object, and any supplied key produces `INVALID_ARGUMENT` without dispatching Outlook work. Its descriptor is read-only, non-destructive, idempotent, and closed-world.

The tool follows the complete production path:

```text
HTTP -> MCP adapter -> bounded dispatcher -> captured Outlook UI STA
     -> Outlook Object Model -> immutable managed DTO -> MCP envelope
```

A successful structured result contains:

- one bounded operation identifier and the standard `ok`, `partial`, and `warnings` fields;
- Outlook version up to 64 characters and bitness restricted to 32 or 64;
- a profile label up to 256 characters;
- captured and executed managed/native thread identifiers, `STA` apartment state, and a required captured-thread match;
- the configured store count and at most 64 returned stores;
- for each store, a display label up to 256 characters, one of `primaryExchangeMailbox`, `exchangeMailbox`, `exchangePublicFolder`, `additionalExchangeMailbox`, `nonExchange`, or `unknown`, three bounded capability flags, and tri-state `available`, `missing`, or `unknown` values for Inbox, Drafts, Sent Items, Deleted Items, and Archive.

Archive availability is always `unknown` because the Outlook Object Model does not expose it. The result therefore carries the fixed archive warning and is structurally partial even when every configured store is returned. Additional fixed warnings may report incomplete metadata or the 64-store result limit; warning text and count are closed and bounded.

Ordinary gateway failures return `CallToolResult.IsError = true` with the fixed failure envelope. Allowed codes are `INVALID_ARGUMENT`, `OUTLOOK_NOT_READY`, `HOST_DEGRADED`, `HOST_STOPPING`, `QUEUE_FULL`, `TIMEOUT`, `COM_BUSY`, `ACCESS_DENIED`, `OBJECT_MODEL_GUARD`, `STA_DISPATCH_FAILED`, and `INTERNAL`. Error messages are sanitized and bounded, and details are an empty closed object.

The transport cancels probe work at 14 seconds, inside the 15-second HTTP handler deadline. The live gate permits readiness retries only for the first probe after a listener start: at most five attempts and 30 seconds total, and only when the exact fixed incomplete-metadata partial shape is returned. Later calls are single-attempt and fail-closed. This polling rule belongs to acceptance tooling; it does not widen the MCP result contract.

## Common bounded-read contract

All seven Phase 4 read tools have closed input and output schemas. Unknown fields, malformed identifiers, out-of-range values, invalid cursor text, and cursor/query mismatches return `INVALID_ARGUMENT` without dispatching Outlook work. Ordinary successes use the standard closed envelope with `ok`, a unique operation ID, `data`, `partial`, and one of four fixed warnings for protected body, truncated body, truncated recipients, or truncated result. The serialized tool result is capped at 1 MiB; if necessary, the transport removes trailing page items, sets `partial=true` and `resultTruncated=true`, and returns a safe continuation cursor. Message text content is not duplicated into the MCP text block.

Store-qualified locators are mandatory:

- a mailbox reference contains `storeId`;
- a folder reference contains `storeId` and `entryId`;
- an item reference contains `storeId`, `entryId`, and `itemClass`.

The gateway reacquires the exact store, folder, or item on every call and validates store/parent/item-class identity. Display names are metadata only and are never used as a fallback. Missing or changed objects produce a typed not-found, moved/deleted, or stale-cursor result.

Page sizes range from 1 through 50 and default to 25. Outlook-side work independently caps a mailbox pass at 64 stores, one folder pass at 1,024 children, one message pass at 4,096 examined items with at most 1,024 equal-timestamp items, conversation traversal at 1,024 nodes and depth 64, and recipient examination at 1,024 entries. Managed page candidates retain at most `pageSize + 1` rows.

Continuation cursors are capped at 96 KiB and have the wire shape `v1.<canonical-base64url-JSON>.<HMAC-SHA-256>`. The payload is readable, not encrypted, and never embeds the bearer token. Its HMAC uses a 32-byte key domain-separated from the process-snapshotted bearer token with `outlook-classic-mcp/cursor-key/v1`. Each cursor binds a tool kind, a SHA-256 hash of the canonical query identity, and a typed keyset anchor:

- mailbox: display name plus store ID;
- folder: display name, store ID, and entry ID;
- message/search/conversation: effective UTC timestamp plus store ID, entry ID, and item class;
- attachment: one-based attachment index plus metadata fingerprint.

Cursors may contain Outlook identifiers and must not be logged or persisted as evidence. A malformed, tampered, wrong-kind, wrong-query, or prior-token cursor returns `INVALID_ARGUMENT` before gateway dispatch. Search scopes are sorted before the canonical query hash is computed. Authenticated anchors are reacquired before use, and a missing or changed anchor returns `CURSOR_STALE` rather than restarting by display name.

Message pages contain at most 50 summaries and include `nextCursor`, `resultTruncated`, `totalScopeCount`, and bounded `scopeFailures`. Cross-scope search executes canonical store/folder order sequentially through the Outlook STA, then merges and exact-deduplicates managed candidates by complete item identity. At least one successful scope can produce `partial=true` with sanitized per-scope failures; partial pages never carry a global continuation cursor. If all scopes fail, the first failure in canonical order becomes the tool error.

Explicit message folders are checked against the store's visible Search Folders before listing or folder-scoped searching. The check examines at most 1,024 search folders and uses Outlook's `CompareEntryIDs`; a match or unavailable/unsupported search-folder metadata returns `UNSUPPORTED_STORE`, while an over-limit enumeration returns `TIMEOUT`. A search scope with no folder selects that mailbox's standard Inbox and does not accept a caller-selected Search Folder.

### `outlook_list_mailboxes`

Input has optional `pageSize` and `cursor`. The output is a deterministic page of mailbox records containing the store-qualified mailbox reference, bounded display name, store type, capability flags, and nullable references for Inbox, Drafts, Sent, Deleted, and Archive. Unlike `outlook_probe`, this tool intentionally returns opaque store identifiers for subsequent read calls.

### `outlook_list_folders`

Input requires `mailbox` and optionally accepts `parentFolder`, `pageSize`, and `cursor`. A null or omitted parent lists direct children of the selected store root; otherwise the parent must belong to that store. The result contains folder and nullable parent references, bounded display names, and `hasChildren` only.

### `outlook_list_messages`

Input requires one `folder` and optionally accepts `pageSize` and `cursor`. The result contains metadata summaries only: store-qualified item and folder references, bounded subject and sender fields, effective/received/sent timestamps, read state, attachment count/presence, and a nullable conversation ID. No body, recipients, attachment bytes, or arbitrary MAPI properties are returned.

### `outlook_search_messages`

Input requires between 1 and 64 explicit scopes. Each scope contains a mailbox and optional folder; an omitted/null folder means that mailbox's standard Inbox. The optional closed filter supports sender, recipient, subject, text, received-from/to UTC, unread state, category, and attachment presence. Text matching and DASL construction remain internal and bounded; callers cannot supply DASL or a property name. Output uses the cross-scope partial-result rules above.

### `outlook_get_message`

Input requires one `item`; optional `bodyFormat` is `plainText` by default or `html`, and `maximumBodyCharacters` ranges from 1 through 50,000. Output contains the message summary, at most 128 entries in each To/Cc/Bcc recipient list with total/truncation metadata, and one body object. A successful body reports its format, returned content, original character count, truncation flag, and protection flag. An access-denied protected body is represented by empty content, unknown original count, `isProtected=true`, and a fixed warning; Object Model Guard denial remains a tool error.

### `outlook_get_conversation`

Input requires the seed `item` and optionally accepts `pageSize` and `cursor`. Conversation traversal is limited to depth 64 and 1,024 examined nodes, returns message metadata only, and validates the conversation identity on continuation. When Outlook returns no conversation for the exact unsent seed, the documented singleton fallback returns that seed as a one-item page.

### `outlook_list_attachments`

Input requires one `item` and optionally accepts `pageSize` and `cursor`. Output contains metadata only: item identity, one-based attachment index, bounded name, size plus `sizeIsKnown`, SHA-256 metadata fingerprint, and nullable bounded content type. It does not open, save, transfer, render, scan, or execute attachment content.

## Future tool requirements

Every future tool must:

- have an explicit typed input and bounded structured output;
- use store-qualified identifiers rather than display-name fallback;
- return message bodies only through the dedicated bounded read tool;
- map ordinary Outlook failures to MCP tool errors with a stable sanitized code;
- keep email content out of logs;
- remain absent when its implementation or policy gate is incomplete.

Send, permanent deletion, arbitrary COM/DASL/MAPI access, scripts, and unrestricted file paths are not available. Sending will accept only a saved draft after preview, fingerprint, idempotency, provider-reconciliation, and approval gates pass.
