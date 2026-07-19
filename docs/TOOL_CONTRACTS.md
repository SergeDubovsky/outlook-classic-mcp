# Tool contracts

## Phase 2 surface

Phase 2 exposes exactly one tool, `outlook_status`. The protocol surface supports initialize, `notifications/initialized`, ping, `tools/list`, and `tools/call` only for that tool. There are no mailbox, resource, prompt, completion, subscription, draft, mutation, attachment-transfer, delete, or send tools.

This contract passed raw endpoint tests and live Codex CLI acceptance without exposing mailbox tools or mailbox data. See [Phase 2 evidence](PHASE_2_EVIDENCE.md).

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

`outlook_probe` is planned for Phase 3 and will be the first tool to read Outlook/profile/store metadata through the STA dispatcher. It is not currently exposed.

## Future tool requirements

Every future tool must:

- have an explicit typed input and bounded structured output;
- use store-qualified identifiers rather than display-name fallback;
- return message bodies only through the dedicated bounded read tool;
- map ordinary Outlook failures to MCP tool errors with a stable sanitized code;
- keep email content out of logs;
- remain absent when its implementation or policy gate is incomplete.

Send, permanent deletion, arbitrary COM/DASL/MAPI access, scripts, and unrestricted file paths are not available. Sending will accept only a saved draft after preview, fingerprint, idempotency, provider-reconciliation, and approval gates pass.
