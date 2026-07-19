# Phase 3 local store inventory

The Phase 3 smoke gate requires an independently prepared Classic Outlook store
inventory. It must not be generated from `outlook_probe` or any MCP response.

1. In Classic Outlook, open **File > Account Settings > Account Settings**.
2. Use the **Email** and **Data Files** tabs, plus each Exchange account's
   **More Settings > Advanced** list, to record every configured store.
3. Copy `store-inventory.template.json` to
   `store-inventory.local.json` and replace the example entry. The
   `*.local.json` suffix is ignored by Git.
4. Preserve the exact display-name spelling, casing, whitespace, and duplicate
   entries shown in Outlook. Do not add Store IDs or other identifiers.

Use these store types:

- `primaryExchangeMailbox`: the profile's primary Exchange delivery store.
- `exchangeMailbox`: another independently configured Exchange mailbox.
- `additionalExchangeMailbox`: a mailbox listed under an Exchange account's
  **Open these additional mailboxes** setting.
- `exchangePublicFolder`: the Exchange Public Folders store.
- `nonExchange`: a PST or other non-Exchange store.
- `unknown`: only when the account settings do not identify the store type.

Run the gate with the local file's absolute path:

```powershell
.\tools\run-phase3-smoke.ps1 `
    -Profile OutlookMcpTest `
    -ExpectedStoreInventoryPath C:\path\to\store-inventory.local.json
```

The scripts read the local names only for an exact normalized multiset
comparison. The successful Phase 3 aggregate contains counts, booleans, and
SHA-256 digests, not profile names, store names, store IDs, process IDs, or
operation IDs.
