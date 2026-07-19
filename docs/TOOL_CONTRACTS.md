# Tool contracts

## Phase 3 surface

Phase 3 exposes exactly two tools, in this order: `outlook_status` and `outlook_probe`. The protocol surface supports initialize, `notifications/initialized`, ping, `tools/list`, and `tools/call` only for those tools. There are no message-read, resource, prompt, completion, subscription, draft, mutation, attachment-transfer, delete, or send tools.

This contract passed raw endpoint tests and live Codex CLI acceptance. Probe output contains bounded profile and store metadata but no messages, bodies, attachments, store identifiers, or Outlook COM objects. See [Phase 3 evidence](PHASE_3_EVIDENCE.md).

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
    "version": "1.0.0.0"
  },
  "partial": false,
  "warnings": []
}
```

`hostState` is limited to 32 characters and `version` to 64 characters. The text content contains only the bounded host-state summary. Arguments produce a typed `INVALID_ARGUMENT` tool error without calling the provider. An unavailable or failing provider produces a bounded, retryable `OUTLOOK_NOT_READY` tool error. Unknown tool names receive a protocol error and are not reflected in an unbounded response.

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

## Future tool requirements

Every future tool must:

- have an explicit typed input and bounded structured output;
- use store-qualified identifiers rather than display-name fallback;
- return message bodies only through the dedicated bounded read tool;
- map ordinary Outlook failures to MCP tool errors with a stable sanitized code;
- keep email content out of logs;
- remain absent when its implementation or policy gate is incomplete.

Send, permanent deletion, arbitrary COM/DASL/MAPI access, scripts, and unrestricted file paths are not available. Sending will accept only a saved draft after preview, fingerprint, idempotency, provider-reconciliation, and approval gates pass.
