# Phase 4 read fixture

The Phase 4 live gate uses a dedicated Classic Outlook profile containing only
synthetic, disposable data. The checked-in template describes the manual
Classic Outlook UI source. The conditional VSTO seeder generates its own local
files; never relabel one source as the other.

## Safety boundary

- Use a dedicated Windows user or test-only Outlook profile such as
  `OutlookMcpTest`.
- Configure at least two distinct stores: one disposable mailbox and one
  disposable PST. Do not attach a personal or production mailbox.
- Use only invented senders, recipients, subjects, bodies, conversations, and
  attachment content. Do not send test mail to a real recipient.
- Keep remote-image loading, links, macros, and attachment execution disabled.
  The smoke gate reads attachment metadata only.
- Never place bearer credentials, Outlook object locators, cursors, mailbox
  data, or message content in source control, test output, screenshots, or bug
  reports.
- Close Outlook normally before the runner starts it. Never terminate Outlook
  forcibly.

The local fixture contains mailbox-derived display names, folder names, marker
text, and attachment metadata. It must remain local even though the data is
synthetic.

## Prepare the local file

From the repository root, copy the placeholder template to the ignored local
filename:

```powershell
Copy-Item .\smoke\outlook\read-fixture.template.json `
    .\smoke\outlook\read-fixture.local.json
```

The `*.local.json` suffix is ignored by Git. Replace every `REPLACE_...` value
before running the gate and keep `schema` set to `1`. Every subject or body
marker must be unique within the profile and at least 16 characters long. Use
exact, case-sensitive display names, folder names, attachment names, and marker
text.
Use at least eight characters for every store display name, folder name, and
attachment filename so each value can serve as an unambiguous privacy sentinel.

Do not add fields for object locators or credentials. The harness resolves
mailboxes, folders, messages, conversations, and attachments at runtime and
keeps the returned opaque values in memory only.

Validate the edited JSON before the live run:

```powershell
$null = Get-Content .\smoke\outlook\read-fixture.local.json -Raw |
    ConvertFrom-Json
```

Protect the local file so only the current Windows user and `SYSTEM` have
access. The harness refuses inherited or broader ACLs:

```powershell
$path = Resolve-Path .\smoke\outlook\read-fixture.local.json
$acl = Get-Acl -LiteralPath $path
$acl.SetAccessRuleProtection($true, $false)
$acl.Access | ForEach-Object { $acl.RemoveAccessRuleSpecific($_) }
$user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
foreach ($identity in @($user, 'SYSTEM')) {
    $rule = [System.Security.AccessControl.FileSystemAccessRule]::new(
        $identity,
        [System.Security.AccessControl.FileSystemRights]::FullControl,
        [System.Security.AccessControl.AccessControlType]::Allow)
    $acl.AddAccessRule($rule)
}
Set-Acl -LiteralPath $path -AclObject $acl
```

For each attachment, replace the numeric `0` sentinel with the exact integer
byte count. Keep it unquoted; for example, use `"size": 128`, not
`"size": "128"`.

## Seed the Outlook profile

The fixture paths start below the named store root; do not include the store
display name in `folderPath`.

1. Configure the two stores listed in `stores`. Each must have an Inbox, and
   each `knownMessage` must identify a different synthetic message in that
   store's Inbox. Put its subject and body markers in the corresponding message.
2. Create the pagination folder with exactly 12 messages and no other items.
   Give the messages distinct received timestamps, do not modify the folder
   during a run, and list `orderedSubjectMarkers` in the order Outlook returns
   them: newest first. A page size of four produces the required three static
   pages.
3. Create the large folder with at least 1,001 synthetic mail items. Keep it
   separate from the static pagination folder.
4. Choose one truthful conversation mode and use the same `source` in both
   local JSON files:
   - `classic-outlook-ui`, as shown by the template, requires a real reply chain
     and between two and fifty expected markers. Matching subjects alone are
     insufficient. Put the seed marker and first expected marker in the seed
     subject, then one distinct marker in every remaining message.
   - `conditional-vsto-seeder` requires exactly one expected subject marker.
     Use only the files produced by the seeder workflow below. The seeder
     creates one unsent item whose Outlook `GetConversation()` result is null;
     the live gate verifies the documented singleton fallback and exact seed
     locator.

   Unsent PST items cannot truthfully emulate a multi-message Outlook
   conversation. Multi-node traversal remains covered by the automated gateway
   and transport tests; the conditional live fixture exercises the real null
   conversation behavior.
5. Create the attachment message with the exact harmless filenames and byte
   sizes in `expectedAttachments`. Use at least two small inert files.
6. Create a message whose body begins with `bodyPrefixMarker` and contains at
   least `minimumCharacterCount` characters. Do not lower the template's
   128-character threshold.
7. If the disposable mailbox supports protected mail, create a synthetic
   protected item and complete `protectedMessage`. Otherwise set
   `protectedMessage` to `null`. A run without that item cannot be cited as live
   protected-body evidence.

Verify the store display names and types independently and prepare
`store-inventory.local.json` using [the Phase 3 instructions](README.md). The
read fixture must cover the complete inventory exactly once. Phase 3 accepts
only `classic-outlook-ui` by default; the Phase 4 runner explicitly allows
`conditional-vsto-seeder` and requires the inventory and read fixture sources
to match.

## Generate the conditional fixture

The conditional seeder is an alternative to the manual preparation above. It
runs only in the smoke build inside `OUTLOOK.EXE` on the Outlook UI STA. From a
non-elevated interactive PowerShell process, run:

```powershell
.\tools\seed-phase4-fixture.ps1 -Action Seed -Profile OutlookMcpTest
```

Optionally pass `-RunDirectory` with a dedicated secure leaf directory. The
driver creates the profile when needed, launches Outlook normally, attaches two
fixture PSTs, and writes `read-fixture.local.json`,
`store-inventory.local.json`, seed status, and driver state into that directory.
The generated read fixture covers the complete three-store accountless profile:
the dedicated bootstrap default store has `expectsInbox: true` and its known
message is resolved through the standard Inbox, while `fixture_store_a` and
`fixture_store_b` have `expectsInbox: false` and each known message declares a
bounded custom `folderPath`. It also contains a 12-item static folder, a
1,001-item large folder, the null-conversation singleton fallback, two inert
attachments, a 1,024-character long body, and `protectedMessage: null`.

Do not edit the generated source, markers, store names, or status. Use the two
generated local JSON paths for the live gate. After the gate and after Outlook
has closed normally, detach the two fixture PSTs with the same run directory:

```powershell
.\tools\seed-phase4-fixture.ps1 `
    -Action Detach `
    -Profile OutlookMcpTest `
    -RunDirectory <same-secure-run-directory>
```

## Run the live gate

Use a non-elevated PowerShell process on a logged-on interactive desktop. Save
work and close Classic Outlook gracefully, then run from the repository root:

```powershell
.\tools\run-phase4-smoke.ps1 `
    -Profile OutlookMcpTest `
    -ExpectedStoreInventoryPath "$PWD\smoke\outlook\store-inventory.local.json" `
    -ReadFixturePath "$PWD\smoke\outlook\read-fixture.local.json"
```

Replace `OutlookMcpTest` if the dedicated profile has a different name. The
runner must discover the fixture through the read-only tools; the local JSON is
not a source for protocol locators.

During the run, do not deliver new messages, edit fixture items, move items, or
allow rules to modify the static folder. After the run, confirm Outlook remains
responsive and close it normally. Preserve only the privacy-safe aggregate
result; do not preserve raw protocol responses or mailbox-derived logs.

The runner performs a full Release build and test pass, creates an ephemeral
current-user signing certificate, registers the exact Release add-in, and then
runs three graceful Outlook cycles: full read workload, restart, and recovery.
It never terminates Outlook forcibly. It verifies at least 1,001 live large-
folder items by bounded pagination, balanced direct COM telemetry, native Codex
discovery, and log privacy before removing the registration and signing state.
Every bounded concurrent call must return either a successful bounded page or
an exact structured `TIMEOUT`. A healthy machine may complete all four calls,
so the live timeout count may be zero. Deterministic structured-timeout and
cancellation behavior is proved by the real-`HttpListener` fake-gateway
transport tests run by the build; the live gate separately proves client
cancellation, endpoint recovery, and Outlook responsiveness.
