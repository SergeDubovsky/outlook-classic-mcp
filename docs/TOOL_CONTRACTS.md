# Tool contracts

No mailbox-facing MCP tools are exposed in Phase 0.

The first tool introduced in Phase 2 will be `outlook_status`, implemented without enumerating stores or messages. `outlook_probe` follows in Phase 3 and reads only Outlook/profile/store metadata.

Every future tool must:

- have an explicit typed input and bounded structured output;
- use store-qualified identifiers rather than display-name fallback;
- return message bodies only through the dedicated bounded read tool;
- map ordinary Outlook failures to MCP tool errors with a stable sanitized code;
- keep email content out of logs;
- remain absent when its implementation or policy gate is incomplete.

Send, permanent deletion, arbitrary COM/DASL/MAPI access, scripts, and unrestricted file paths are not available. Sending will accept only a saved draft after preview, fingerprint, idempotency, provider-reconciliation, and approval gates pass.
