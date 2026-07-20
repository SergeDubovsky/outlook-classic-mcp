#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ExpectedStoreInventoryPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ReadFixturePath,

    [Parameter(Mandatory = $true)]
    [ValidateRange(1, 2147483647)]
    [int]$OutlookProcessId,

    [ValidateSet('Full', 'Restart')]
    [string]$Mode = 'Full',

    [Uri]$Endpoint = 'http://127.0.0.1:8765/mcp/',

    [ValidateRange(15, 60)]
    [int]$TimeoutSeconds = 30,

    [ValidateRange(60, 1000)]
    [int]$RepeatedReadCount = 200,

    [ValidateRange(1, 3)]
    [int]$CycleNumber = 1,

    [AllowNull()]
    [System.Collections.Generic.HashSet[string]]$SharedOperationIds,

    [AllowNull()]
    [System.Collections.Generic.HashSet[string]]$SharedSensitiveValues
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (($RepeatedReadCount % 20) -ne 0) {
    throw '-RepeatedReadCount must be divisible by twenty.'
}

$repositoryRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'phase3-smoke-common.ps1')
. (Join-Path $PSScriptRoot 'phase4-smoke-common.ps1')

$expectedInventory = Import-Phase3ExpectedStoreInventory `
    -Path $ExpectedStoreInventoryPath `
    -RepositoryRoot $repositoryRoot `
    -AllowedSources @('classic-outlook-ui', 'conditional-vsto-seeder')
$fixture = Import-Phase4ReadFixture `
    -Path $ReadFixturePath `
    -RepositoryRoot $repositoryRoot `
    -ExpectedInventory $expectedInventory

$token = [Environment]::GetEnvironmentVariable('OUTLOOK_MCP_TOKEN', 'Process')
if ($token -cnotmatch '^[A-Za-z0-9_-]{43}$') {
    $token = [Environment]::GetEnvironmentVariable('OUTLOOK_MCP_TOKEN', 'User')
}
if ($token -cnotmatch '^[A-Za-z0-9_-]{43}$') {
    throw 'OUTLOOK_MCP_TOKEN is not a canonical current-user or process token.'
}

$script:phase4Context = New-Phase4HttpContext `
    -Endpoint $Endpoint `
    -Token $token `
    -TimeoutSeconds $TimeoutSeconds
$script:phase4NextRequestId = 4000L
$script:phase4OperationIds = $null
if ($null -eq $SharedOperationIds) {
    $script:phase4OperationIds =
        [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
}
else {
    $script:phase4OperationIds = $SharedOperationIds
}
$operationCountBefore = $script:phase4OperationIds.Count
$script:phase4SensitiveValues = $null
if ($null -eq $SharedSensitiveValues) {
    $script:phase4SensitiveValues =
        [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
}
else {
    $script:phase4SensitiveValues = $SharedSensitiveValues
}
foreach ($sensitiveValue in $fixture.SensitiveValues) {
    $null = $script:phase4SensitiveValues.Add($sensitiveValue)
}
Add-Phase4SensitiveValue -Set $script:phase4SensitiveValues -Value $token
$script:phase4ResponsiveDurations = [System.Collections.Generic.List[double]]::new()

function Get-Phase4NextRequestId {
    $script:phase4NextRequestId++
    return $script:phase4NextRequestId
}

function Invoke-EndpointTool {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][object]$Arguments
    )

    return Invoke-Phase4ToolCall `
        -Context $script:phase4Context `
        -Id (Get-Phase4NextRequestId) `
        -ToolName $Name `
        -Arguments $Arguments `
        -SeenOperationIds $script:phase4OperationIds `
        -SensitiveValues $script:phase4SensitiveValues
}

function Assert-EndpointResponsive {
    $duration = Assert-Phase4OutlookResponsive -OutlookProcessId $OutlookProcessId
    $script:phase4ResponsiveDurations.Add($duration)
}

function Assert-Phase4ToolCatalog {
    $result = Invoke-Phase4McpRequest `
        -Context $script:phase4Context `
        -Id (Get-Phase4NextRequestId) `
        -Method 'tools/list' `
        -Params ([ordered]@{})
    Assert-Phase4ExactProperties `
        -InputObject $result `
        -Expected @('tools') `
        -Context 'tools/list result'
    $tools = @($result.tools)
    if ($result.tools -isnot [System.Array] -or
        $tools.Count -ne $script:Phase4ExpectedTools.Count) {
        throw 'The Phase 4 endpoint did not expose exactly nine tools.'
    }
    for ($index = 0; $index -lt $tools.Count; $index++) {
        $tool = $tools[$index]
        if ($tool.name -cne $script:Phase4ExpectedTools[$index] -or
            $null -eq $tool.inputSchema -or
            $null -eq $tool.outputSchema -or
            $null -eq $tool.annotations -or
            $tool.annotations.readOnlyHint -ne $true -or
            $tool.annotations.destructiveHint -ne $false -or
            $tool.annotations.idempotentHint -ne $true -or
            $tool.annotations.openWorldHint -ne $false) {
            throw 'The Phase 4 tool order, schemas, or read-only annotations are invalid.'
        }
        if ($tool.inputSchema.type -cne 'object' -or
            $tool.inputSchema.additionalProperties -ne $false -or
            $tool.outputSchema.type -cne 'object') {
            throw 'A Phase 4 tool descriptor is not closed and typed.'
        }
    }
}

function Get-Phase4StatusEnvelope {
    param(
        [Parameter(Mandatory = $true)][string]$Context,
        [bool]$RequireQuiescent = $true
    )

    $status = Invoke-EndpointTool -Name 'outlook_status' -Arguments ([ordered]@{})
    $envelope = Assert-Phase4SuccessEnvelope `
        -Result $status `
        -ExpectedDataFields @('hostState', 'listenerReady', 'version', 'readDiagnostics') `
        -ExpectedPartial $false `
        -Context $Context
    if ($envelope.data.hostState -cne 'online' -or
        $envelope.data.listenerReady -ne $true) {
        throw "$Context did not report a ready Phase 4 host."
    }

    $diagnostics = $envelope.data.readDiagnostics
    Assert-Phase4ExactProperties `
        -InputObject $diagnostics `
        -Expected @(
            'comAcquired', 'comReleased', 'comOutstanding', 'comPeak',
            'materializedItemHighWater'
        ) `
        -Context "$Context read diagnostics"
    $values = @{}
    foreach ($property in @(
            'comAcquired', 'comReleased', 'comOutstanding', 'comPeak',
            'materializedItemHighWater')) {
        $value = ConvertTo-Phase4Integer `
            -Value $diagnostics.$property `
            -Context "$Context $property"
        if ($value -lt 0) {
            throw "$Context returned a negative read diagnostic."
        }
        $values[$property] = $value
    }
    if ($values.comReleased -gt $values.comAcquired -or
        $values.comOutstanding -ne ($values.comAcquired - $values.comReleased) -or
        $values.comPeak -lt $values.comOutstanding -or
        $values.comPeak -gt $values.comAcquired -or
        ($RequireQuiescent -and $values.comOutstanding -ne 0)) {
        throw "$Context returned unbalanced Outlook COM diagnostics."
    }
    if ($values.materializedItemHighWater -gt 51) {
        throw "$Context exceeded the bounded page-size-plus-one materialization limit."
    }
    return $envelope
}

function Wait-Phase4ReadDiagnosticsQuiescent {
    param([Parameter(Mandatory = $true)][string]$Context)

    $deadline = [DateTime]::UtcNow.AddSeconds(15)
    do {
        $envelope = Get-Phase4StatusEnvelope `
            -Context $Context `
            -RequireQuiescent $false
        if ([int64]$envelope.data.readDiagnostics.comOutstanding -eq 0) {
            return $envelope
        }
        Assert-EndpointResponsive
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)

    throw 'Timed-out or canceled Phase 4 reads did not release their Outlook COM objects before the quiescence deadline.'
}

function Invoke-Phase4Handshake {
    $unauthorizedBody = New-Phase4JsonRpcBody `
        -Id (Get-Phase4NextRequestId) `
        -Method 'ping' `
        -Params ([ordered]@{})
    $unauthorizedPending = Start-Phase4McpPost `
        -Context $script:phase4Context `
        -Body $unauthorizedBody `
        -Authenticated $false
    $unauthorized = Complete-Phase4McpPost -Pending $unauthorizedPending
    if ($unauthorized.StatusCode -ne 401) {
        throw 'The Phase 4 endpoint did not fail closed without authentication.'
    }
    foreach ($toolName in $script:Phase4ExpectedTools) {
        if ($unauthorized.Body.Contains($toolName)) {
            throw 'The unauthenticated Phase 4 response disclosed capabilities.'
        }
    }

    $initializeId = Get-Phase4NextRequestId
    $initializeBody = New-Phase4JsonRpcBody `
        -Id $initializeId `
        -Method 'initialize' `
        -Params ([ordered]@{
            protocolVersion = $script:Phase4ProtocolVersion
            capabilities = [ordered]@{}
            clientInfo = [ordered]@{ name = 'phase4-smoke'; version = '1.0' }
        })
    $initializePending = Start-Phase4McpPost `
        -Context $script:phase4Context `
        -Body $initializeBody `
        -ProtocolVersion $null
    $initialize = ConvertFrom-Phase4SseResponse `
        -Response (Complete-Phase4McpPost -Pending $initializePending)
    if ((ConvertTo-Phase4Integer -Value $initialize.id -Context 'initialize ID') -ne $initializeId -or
        $initialize.result.protocolVersion -cne $script:Phase4ProtocolVersion -or
        $initialize.result.serverInfo.name -cne 'outlook-classic-mcp') {
        throw 'The Phase 4 initialize result is invalid.'
    }

    $notification = [ordered]@{
        jsonrpc = '2.0'
        method = 'notifications/initialized'
        params = [ordered]@{}
    } | ConvertTo-Json -Depth 8 -Compress
    $notificationPending = Start-Phase4McpPost `
        -Context $script:phase4Context `
        -Body $notification
    $notificationResponse = Complete-Phase4McpPost -Pending $notificationPending
    if ($notificationResponse.StatusCode -ne 202 -or
        $notificationResponse.Body.Length -ne 0 -or
        $null -ne $notificationResponse.ContentType -or
        $notificationResponse.ContentEncoding.Count -ne 0) {
        throw 'The initialized notification did not receive the empty HTTP 202 contract.'
    }

    $ping = Invoke-Phase4McpRequest `
        -Context $script:phase4Context `
        -Id (Get-Phase4NextRequestId) `
        -Method 'ping' `
        -Params ([ordered]@{})
    if ($null -eq $ping) {
        throw 'The Phase 4 ping result is missing.'
    }

    Assert-Phase4ToolCatalog

    $null = Get-Phase4StatusEnvelope -Context 'outlook_status'
}

function Invoke-Phase4Probe {
    $requestId = Get-Phase4NextRequestId
    $raw = Invoke-Phase4McpRequest `
        -Context $script:phase4Context `
        -Id $requestId `
        -Method 'tools/call' `
        -Params ([ordered]@{ name = 'outlook_probe'; arguments = [ordered]@{} })
    $localIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $proof = Assert-Phase3ProbeToolResult `
        -Result $raw `
        -ExpectedInventory $expectedInventory `
        -SeenOperationIds $localIds
    if ($localIds.Count -ne 1) {
        throw 'The Phase 4 probe did not return one operation ID.'
    }
    foreach ($operationId in $localIds) {
        if (-not $script:phase4OperationIds.Add($operationId)) {
            throw 'The Phase 4 probe reused an operation ID.'
        }
    }
    Add-Phase4SensitiveValuesFromObject `
        -Set $script:phase4SensitiveValues `
        -InputObject $raw.structuredContent
    return $proof
}

function Assert-Phase4MailboxShape {
    param([Parameter(Mandatory = $true)][object]$Mailbox)

    Assert-Phase4ExactProperties `
        -InputObject $Mailbox `
        -Expected @('mailbox', 'displayName', 'storeType', 'capabilities', 'standardFolders') `
        -Context 'mailbox record'
    Assert-Phase4ExactProperties `
        -InputObject $Mailbox.mailbox `
        -Expected @('storeId') `
        -Context 'mailbox reference'
    Assert-Phase4ExactProperties `
        -InputObject $Mailbox.standardFolders `
        -Expected @('inbox', 'drafts', 'sent', 'deleted', 'archive') `
        -Context 'standard-folder references'
    if ($Mailbox.mailbox.storeId -isnot [string] -or
        $Mailbox.mailbox.storeId.Length -lt 1 -or
        $Mailbox.storeType -cnotin $script:Phase4StoreTypes) {
        throw 'A Phase 4 mailbox record has invalid scalar fields.'
    }
    Assert-Phase4ExactProperties `
        -InputObject $Mailbox.capabilities `
        -Expected @('isExchangeStore', 'isDataFileStore', 'isCachedExchange') `
        -Context 'mailbox capabilities'
    foreach ($capability in @('isExchangeStore', 'isDataFileStore', 'isCachedExchange')) {
        if ($Mailbox.capabilities.$capability -isnot [bool]) {
            throw 'A Phase 4 mailbox capability is not Boolean.'
        }
    }
    foreach ($name in @('inbox', 'drafts', 'sent', 'deleted', 'archive')) {
        $folder = $Mailbox.standardFolders.$name
        if ($null -ne $folder) {
            Assert-Phase4ExactProperties `
                -InputObject $folder `
                -Expected @('storeId', 'entryId') `
                -Context 'standard-folder reference'
            if ($folder.storeId -cne $Mailbox.mailbox.storeId) {
                throw 'A standard-folder reference crossed a mailbox boundary.'
            }
        }
    }
}

function Get-Phase4Mailboxes {
    $mailboxes = [System.Collections.Generic.List[object]]::new()
    $cursor = $null
    for ($page = 0; $page -lt 64; $page++) {
        $arguments = [ordered]@{ pageSize = 50 }
        if ($null -ne $cursor) {
            $arguments.cursor = $cursor
        }
        $result = Invoke-EndpointTool -Name 'outlook_list_mailboxes' -Arguments $arguments
        $envelope = Assert-Phase4SuccessEnvelope `
            -Result $result `
            -ExpectedDataFields @('mailboxes', 'nextCursor', 'resultTruncated') `
            -ExpectedPartial $false `
            -Context 'outlook_list_mailboxes'
        if ($envelope.data.mailboxes -isnot [System.Array] -or
            $envelope.data.resultTruncated -isnot [bool] -or
            $envelope.data.resultTruncated -ne $false -or
            @($envelope.warnings).Count -ne 0 -or
            @($envelope.data.mailboxes).Count -gt 50) {
            throw 'The mailbox page is not bounded.'
        }
        foreach ($mailbox in @($envelope.data.mailboxes)) {
            Assert-Phase4MailboxShape -Mailbox $mailbox
            $mailboxes.Add($mailbox)
        }
        $cursor = $envelope.data.nextCursor
        if ($null -eq $cursor) {
            break
        }
        if ($cursor -isnot [string] -or $cursor.Length -gt 2048) {
            throw 'The mailbox continuation cursor is invalid.'
        }
    }
    if ($null -ne $cursor) {
        throw 'Mailbox enumeration exceeded its bounded page count.'
    }

    $inventoryEntries = [System.Collections.Generic.List[string]]::new()
    foreach ($mailbox in $mailboxes) {
        $inventoryEntries.Add((New-Phase3InventoryEntry `
            -DisplayName $mailbox.displayName `
            -StoreType $mailbox.storeType))
    }
    $actualInventory = Get-Phase3StringMultiset -Entries $inventoryEntries.ToArray()
    Assert-Phase3InventoryMatch -Expected $expectedInventory -Actual $actualInventory

    foreach ($storeFixture in @($fixture.Data.stores)) {
        $matches = @(
            $mailboxes |
                Where-Object {
                    $_.displayName -ceq $storeFixture.displayName -and
                    $_.storeType -ceq $storeFixture.storeType
                }
        )
        if ($matches.Count -ne 1) {
            throw 'A Phase 4 fixture store was not uniquely discoverable.'
        }
        if ($storeFixture.expectsInbox -and $null -eq $matches[0].standardFolders.inbox) {
            throw 'A fixture store expected to have an Inbox did not return one.'
        }
    }
    return $mailboxes.ToArray()
}

function Get-Phase4MailboxForAlias {
    param(
        [Parameter(Mandatory = $true)][object[]]$Mailboxes,
        [Parameter(Mandatory = $true)][string]$Alias
    )

    $storeFixture = @($fixture.Data.stores | Where-Object { $_.alias -ceq $Alias })
    if ($storeFixture.Count -ne 1) {
        throw 'The Phase 4 fixture references an unknown store alias.'
    }
    $matches = @(
        $Mailboxes |
            Where-Object {
                $_.displayName -ceq $storeFixture[0].displayName -and
                $_.storeType -ceq $storeFixture[0].storeType
            }
    )
    if ($matches.Count -ne 1) {
        throw 'The Phase 4 mailbox alias did not resolve uniquely.'
    }
    return $matches[0]
}

function Assert-Phase4MessagePage {
    param(
        [Parameter(Mandatory = $true)][object]$Result,
        [Parameter(Mandatory = $true)][string]$Context,
        [AllowNull()][bool]$ExpectedPartial
    )

    $arguments = @{
        Result = $Result
        ExpectedDataFields = @(
            'messages', 'nextCursor', 'resultTruncated',
            'totalScopeCount', 'scopeFailures'
        )
        Context = $Context
    }
    if ($PSBoundParameters.ContainsKey('ExpectedPartial')) {
        $arguments.ExpectedPartial = $ExpectedPartial
    }
    $envelope = Assert-Phase4SuccessEnvelope @arguments
    if ($envelope.data.messages -isnot [System.Array] -or
        $envelope.data.scopeFailures -isnot [System.Array] -or
        @($envelope.data.messages).Count -gt 50 -or
        $envelope.data.resultTruncated -isnot [bool] -or
        $envelope.data.resultTruncated -ne $false -or
        @($envelope.warnings).Count -ne 0 -or
        (ConvertTo-Phase4Integer -Value $envelope.data.totalScopeCount -Context 'scope count') -lt 1) {
        throw "$Context returned an invalid bounded message page."
    }
    foreach ($message in @($envelope.data.messages)) {
        Assert-Phase4ExactProperties `
            -InputObject $message `
            -Expected @(
                'item', 'folder', 'subject', 'senderDisplayName', 'senderAddress',
                'effectiveTimestampUtc', 'receivedUtc', 'sentUtc', 'isRead',
                'attachmentCount', 'hasAttachments', 'conversationId'
            ) `
            -Context 'message summary'
        Assert-Phase4ExactProperties `
            -InputObject $message.item `
            -Expected @('storeId', 'entryId', 'itemClass') `
            -Context 'message locator'
    }
    return $envelope
}

function Resolve-Phase4FolderPath {
    param(
        [Parameter(Mandatory = $true)][object]$Mailbox,
        [Parameter(Mandatory = $true)][object[]]$FolderPath
    )

    $parent = $null
    foreach ($component in $FolderPath) {
        $matches = [System.Collections.Generic.List[object]]::new()
        $cursor = $null
        for ($page = 0; $page -lt 64; $page++) {
            $arguments = [ordered]@{
                mailbox = [ordered]@{ storeId = $Mailbox.mailbox.storeId }
                pageSize = 50
            }
            if ($null -ne $parent) {
                $arguments.parentFolder = [ordered]@{
                    storeId = $parent.storeId
                    entryId = $parent.entryId
                }
            }
            if ($null -ne $cursor) {
                $arguments.cursor = $cursor
            }
            $result = Invoke-EndpointTool -Name 'outlook_list_folders' -Arguments $arguments
            $envelope = Assert-Phase4SuccessEnvelope `
                -Result $result `
                -ExpectedDataFields @('folders', 'nextCursor', 'resultTruncated') `
                -ExpectedPartial $false `
                -Context 'outlook_list_folders'
            if ($envelope.data.folders -isnot [System.Array] -or
                @($envelope.data.folders).Count -gt 50 -or
                $envelope.data.resultTruncated -ne $false -or
                @($envelope.warnings).Count -ne 0) {
                throw 'A folder page exceeded the Phase 4 page cap.'
            }
            foreach ($folder in @($envelope.data.folders)) {
                Assert-Phase4ExactProperties `
                    -InputObject $folder `
                    -Expected @('folder', 'parentFolder', 'displayName', 'hasChildren') `
                    -Context 'folder summary'
                if ($folder.displayName -ceq $component) {
                    $matches.Add($folder.folder)
                }
            }
            $cursor = $envelope.data.nextCursor
            if ($null -eq $cursor) {
                break
            }
        }
        if ($null -ne $cursor -or $matches.Count -ne 1) {
            throw 'A fixture folder path did not resolve uniquely within the bounded traversal.'
        }
        $parent = $matches[0]
    }
    return $parent
}

function Find-Phase4Message {
    param(
        [Parameter(Mandatory = $true)][object]$Folder,
        [Parameter(Mandatory = $true)][string]$SubjectMarker
    )

    $matches = [System.Collections.Generic.List[object]]::new()
    $cursor = $null
    for ($page = 0; $page -lt 64; $page++) {
        $arguments = [ordered]@{
            folder = [ordered]@{ storeId = $Folder.storeId; entryId = $Folder.entryId }
            pageSize = 50
        }
        if ($null -ne $cursor) {
            $arguments.cursor = $cursor
        }
        $result = Invoke-EndpointTool -Name 'outlook_list_messages' -Arguments $arguments
        $envelope = Assert-Phase4MessagePage `
            -Result $result `
            -Context 'outlook_list_messages' `
            -ExpectedPartial $false
        foreach ($message in @($envelope.data.messages)) {
            if ($message.subject -is [string] -and
                $message.subject.Contains($SubjectMarker)) {
                $matches.Add($message)
            }
        }
        $cursor = $envelope.data.nextCursor
        if ($null -eq $cursor) {
            break
        }
    }
    if ($null -ne $cursor -or $matches.Count -ne 1) {
        throw 'A seeded Phase 4 message was not uniquely listed within the bounded traversal.'
    }
    return $matches[0]
}

function Get-Phase4MessageDetail {
    param(
        [Parameter(Mandatory = $true)][object]$Message,
        [ValidateRange(1, 50000)][int]$MaximumBodyCharacters = 50000,
        [bool]$ExpectedPartial = $false
    )

    $result = Invoke-EndpointTool `
        -Name 'outlook_get_message' `
        -Arguments ([ordered]@{
            item = [ordered]@{
                storeId = $Message.item.storeId
                entryId = $Message.item.entryId
                itemClass = $Message.item.itemClass
            }
            bodyFormat = 'plainText'
            maximumBodyCharacters = $MaximumBodyCharacters
        })
    $envelope = Assert-Phase4SuccessEnvelope `
        -Result $result `
        -ExpectedDataFields @(
            'message', 'toRecipients', 'ccRecipients', 'bccRecipients',
            'totalToRecipientCount', 'totalCcRecipientCount', 'totalBccRecipientCount',
            'toRecipientsTruncated', 'ccRecipientsTruncated', 'bccRecipientsTruncated',
            'body'
        ) `
        -ExpectedPartial $ExpectedPartial `
        -Context 'outlook_get_message'
    Assert-Phase4ExactProperties `
        -InputObject $envelope.data.body `
        -Expected @('format', 'content', 'originalCharacterCount', 'isTruncated', 'isProtected') `
        -Context 'message body'
    foreach ($recipientList in @('toRecipients', 'ccRecipients', 'bccRecipients')) {
        if ($envelope.data.$recipientList -isnot [System.Array]) {
            throw 'outlook_get_message returned a non-array recipient list.'
        }
    }
    if ($envelope.data.message.item.storeId -cne $Message.item.storeId -or
        $envelope.data.message.item.entryId -cne $Message.item.entryId -or
        $envelope.data.message.item.itemClass -cne $Message.item.itemClass) {
        throw 'outlook_get_message did not reacquire the exact listed locator.'
    }
    if ($envelope.data.body.format -cne 'plainText' -or
        $envelope.data.body.content -isnot [string] -or
        $envelope.data.body.content.Length -gt $MaximumBodyCharacters -or
        $envelope.data.body.isTruncated -isnot [bool] -or
        $envelope.data.body.isProtected -isnot [bool]) {
        throw 'outlook_get_message returned an invalid bounded body projection.'
    }
    if (-not $ExpectedPartial -and @($envelope.warnings).Count -ne 0) {
        throw 'A complete message detail unexpectedly returned warnings.'
    }
    return $envelope
}

function Test-Phase4KnownMessages {
    param([Parameter(Mandatory = $true)][object[]]$Mailboxes)

    $reacquired = 0
    foreach ($storeFixture in @($fixture.Data.stores)) {
        $mailbox = Get-Phase4MailboxForAlias -Mailboxes $Mailboxes -Alias $storeFixture.alias
        $knownMessageFolder = $null
        if ($fixture.Data.source -cne 'conditional-vsto-seeder' -or
            $storeFixture.expectsInbox) {
            if ($null -eq $mailbox.standardFolders.inbox) {
                throw 'A known-message fixture store has no expected standard Inbox locator.'
            }
            $knownMessageFolder = $mailbox.standardFolders.inbox
        }
        else {
            $knownMessageFolder = Resolve-Phase4FolderPath `
                -Mailbox $mailbox `
                -FolderPath @($storeFixture.knownMessage.folderPath)
        }
        $message = Find-Phase4Message `
            -Folder $knownMessageFolder `
            -SubjectMarker $storeFixture.knownMessage.subjectMarker
        $detail = Get-Phase4MessageDetail -Message $message
        if ($detail.data.message.subject -isnot [string] -or
            -not $detail.data.message.subject.Contains($storeFixture.knownMessage.subjectMarker) -or
            $detail.data.body.content -isnot [string] -or
            -not $detail.data.body.content.Contains($storeFixture.knownMessage.bodyMarker)) {
            throw 'A known seeded message did not match its local fixture.'
        }
        $null = Get-Phase4MessageDetail -Message $message
        $reacquired++
    }
    return $reacquired
}

function Test-Phase4StaticPagination {
    param([Parameter(Mandatory = $true)][object[]]$Mailboxes)

    $mailbox = Get-Phase4MailboxForAlias `
        -Mailboxes $Mailboxes `
        -Alias $fixture.Data.pagination.storeAlias
    $folder = Resolve-Phase4FolderPath `
        -Mailbox $mailbox `
        -FolderPath @($fixture.Data.pagination.folderPath)
    $expected = @($fixture.Data.pagination.orderedSubjectMarkers)
    $seenLocators = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $cursor = $null
    $firstCursor = $null
    for ($page = 0; $page -lt 3; $page++) {
        $arguments = [ordered]@{
            folder = [ordered]@{ storeId = $folder.storeId; entryId = $folder.entryId }
            pageSize = 4
        }
        if ($null -ne $cursor) {
            $arguments.cursor = $cursor
        }
        $result = Invoke-EndpointTool -Name 'outlook_list_messages' -Arguments $arguments
        $envelope = Assert-Phase4MessagePage `
            -Result $result `
            -Context 'static pagination page' `
            -ExpectedPartial $false
        $messages = @($envelope.data.messages)
        if ($messages.Count -ne 4 -or @($envelope.data.scopeFailures).Count -ne 0) {
            throw 'A static pagination page did not contain exactly four messages.'
        }
        for ($offset = 0; $offset -lt 4; $offset++) {
            $expectedMarker = $expected[$page * 4 + $offset]
            if ($messages[$offset].subject -isnot [string] -or
                -not $messages[$offset].subject.Contains($expectedMarker) -or
                -not $seenLocators.Add((Get-Phase4LocatorKey -Item $messages[$offset].item))) {
                throw 'Static pagination contained an ordering gap or duplicate.'
            }
        }
        $cursor = $envelope.data.nextCursor
        if ($page -eq 0) {
            $firstCursor = $cursor
        }
        if ($page -lt 2 -and $null -eq $cursor) {
            throw 'Static pagination ended before three pages were returned.'
        }
    }
    if ($null -ne $cursor -or $seenLocators.Count -ne 12 -or $null -eq $firstCursor) {
        throw 'Static pagination did not terminate exactly after twelve fixture messages.'
    }

    $lastCharacter = $firstCursor.Substring($firstCursor.Length - 1, 1)
    $replacement = if ($lastCharacter -cne 'A') { 'A' } else { 'B' }
    $tampered = $firstCursor.Substring(0, $firstCursor.Length - 1) + $replacement
    $tamperedResult = Invoke-EndpointTool `
        -Name 'outlook_list_messages' `
        -Arguments ([ordered]@{
            folder = [ordered]@{ storeId = $folder.storeId; entryId = $folder.entryId }
            pageSize = 4
            cursor = $tampered
        })
    $null = Assert-Phase4ErrorEnvelope `
        -Result $tamperedResult `
        -ExpectedCode 'INVALID_ARGUMENT' `
        -Context 'tampered cursor'

    $wrongKind = Invoke-EndpointTool `
        -Name 'outlook_search_messages' `
        -Arguments ([ordered]@{
            scopes = @([ordered]@{
                mailbox = [ordered]@{ storeId = $mailbox.mailbox.storeId }
                folder = [ordered]@{ storeId = $folder.storeId; entryId = $folder.entryId }
            })
            pageSize = 4
            cursor = $firstCursor
        })
    $null = Assert-Phase4ErrorEnvelope `
        -Result $wrongKind `
        -ExpectedCode 'INVALID_ARGUMENT' `
        -Context 'wrong-kind cursor'

    return [pscustomobject]@{
        Folder = $folder
        PageCount = 3
        ItemCount = 12
        CursorErrorsVerified = 2
    }
}

function Test-Phase4LargeFolder {
    param([Parameter(Mandatory = $true)][object[]]$Mailboxes)

    $mailbox = Get-Phase4MailboxForAlias `
        -Mailboxes $Mailboxes `
        -Alias $fixture.Data.largeFolder.storeAlias
    $folder = Resolve-Phase4FolderPath `
        -Mailbox $mailbox `
        -FolderPath @($fixture.Data.largeFolder.folderPath)
    $cursor = $null
    $seenLocators = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $seenCursors = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    for ($page = 0; $page -lt 21; $page++) {
        $arguments = [ordered]@{
            folder = [ordered]@{ storeId = $folder.storeId; entryId = $folder.entryId }
            pageSize = 50
        }
        if ($null -ne $cursor) {
            $arguments.cursor = $cursor
        }
        $result = Invoke-EndpointTool `
            -Name 'outlook_list_messages' `
            -Arguments $arguments
        $envelope = Assert-Phase4MessagePage `
            -Result $result `
            -Context 'large-folder page' `
            -ExpectedPartial $false
        $pageMessages = @($envelope.data.messages)
        $pageCount = $pageMessages.Count
        foreach ($message in $pageMessages) {
            if (-not $seenLocators.Add((Get-Phase4LocatorKey -Item $message.item))) {
                throw 'The independently large folder repeated a message locator across continuation pages.'
            }
        }
        $cursor = $envelope.data.nextCursor
        if ($pageCount -lt 1 -or
            ($seenLocators.Count -lt 1001 -and $pageCount -ne 50) -or
            ($seenLocators.Count -lt 1001 -and $cursor -isnot [string]) -or
            ($null -ne $cursor -and -not $seenCursors.Add($cursor))) {
            throw 'The independently large folder did not return a bounded continuation page.'
        }
        if ($seenLocators.Count -ge 1001) {
            break
        }
    }
    if ($seenLocators.Count -lt 1001) {
        throw 'The live large folder did not expose more than one thousand distinct message locators.'
    }
    return [pscustomobject]@{
        Folder = $folder
        DistinctItemCount = $seenLocators.Count
    }
}

function Test-Phase4PartialSearch {
    param([Parameter(Mandatory = $true)][object[]]$Mailboxes)

    $storeFixture = @($fixture.Data.stores)[0]
    $mailbox = Get-Phase4MailboxForAlias -Mailboxes $Mailboxes -Alias $storeFixture.alias
    $result = Invoke-EndpointTool `
        -Name 'outlook_search_messages' `
        -Arguments ([ordered]@{
            scopes = @(
                [ordered]@{
                    mailbox = [ordered]@{ storeId = $mailbox.mailbox.storeId }
                    folder = [ordered]@{
                        storeId = $mailbox.standardFolders.inbox.storeId
                        entryId = $mailbox.standardFolders.inbox.entryId
                    }
                },
                [ordered]@{
                    mailbox = [ordered]@{ storeId = 'PHASE4_DETERMINISTIC_MISSING_STORE' }
                }
            )
            filter = [ordered]@{ subject = $storeFixture.knownMessage.subjectMarker }
            pageSize = 10
        })
    $envelope = Assert-Phase4MessagePage `
        -Result $result `
        -Context 'partial cross-store search' `
        -ExpectedPartial $true
    $failures = @($envelope.data.scopeFailures)
    if ($envelope.data.scopeFailures -isnot [System.Array] -or
        (ConvertTo-Phase4Integer -Value $envelope.data.totalScopeCount -Context 'partial scope count') -ne 2 -or
        $failures.Count -ne 1 -or
        $failures[0].scopeIndex -ne 1 -or
        $failures[0].code -cne 'STORE_NOT_FOUND' -or
        @($envelope.data.messages).Count -lt 1 -or
        $null -ne $envelope.data.nextCursor) {
        throw 'The cross-store failure did not return bounded partial success.'
    }
    $null = Get-Phase4StatusEnvelope -Context 'post-partial status'
}

function Test-Phase4Conversation {
    param([Parameter(Mandatory = $true)][object[]]$Mailboxes)

    $mailbox = Get-Phase4MailboxForAlias `
        -Mailboxes $Mailboxes `
        -Alias $fixture.Data.conversation.storeAlias
    $folder = Resolve-Phase4FolderPath `
        -Mailbox $mailbox `
        -FolderPath @($fixture.Data.conversation.folderPath)
    $seed = Find-Phase4Message `
        -Folder $folder `
        -SubjectMarker $fixture.Data.conversation.seedSubjectMarker
    $singletonMode = $fixture.Data.source -ceq 'conditional-vsto-seeder'
    if (($singletonMode -and $null -ne $seed.conversationId) -or
        (-not $singletonMode -and $null -eq $seed.conversationId)) {
        throw 'The conversation seed does not match the declared fixture source mode.'
    }
    $result = Invoke-EndpointTool `
        -Name 'outlook_get_conversation' `
        -Arguments ([ordered]@{
            item = [ordered]@{
                storeId = $seed.item.storeId
                entryId = $seed.item.entryId
                itemClass = $seed.item.itemClass
            }
            pageSize = 50
        })
    $envelope = Assert-Phase4MessagePage `
        -Result $result `
        -Context 'outlook_get_conversation' `
        -ExpectedPartial $false
    $subjects = @($envelope.data.messages | ForEach-Object { $_.subject })
    $expectedMarkers = @($fixture.Data.conversation.expectedSubjectMarkers)
    if ($null -ne $envelope.data.nextCursor -or $subjects.Count -ne $expectedMarkers.Count) {
        throw 'The seeded conversation did not fit its deterministic bounded page.'
    }
    if ($singletonMode -and
        ($subjects.Count -ne 1 -or
         $envelope.data.messages[0].conversationId -ne $null -or
         (Get-Phase4LocatorKey -Item $envelope.data.messages[0].item) -cne
            (Get-Phase4LocatorKey -Item $seed.item))) {
        throw 'The conditional VSTO fixture did not verify the null-conversation singleton fallback.'
    }
    foreach ($marker in $expectedMarkers) {
        if (@($subjects | Where-Object { $_ -is [string] -and $_.Contains($marker) }).Count -ne 1) {
            throw 'The seeded conversation did not contain every expected fixture message.'
        }
    }
    return @($envelope.data.messages).Count
}

function Test-Phase4Attachments {
    param([Parameter(Mandatory = $true)][object[]]$Mailboxes)

    $mailbox = Get-Phase4MailboxForAlias `
        -Mailboxes $Mailboxes `
        -Alias $fixture.Data.attachmentMessage.storeAlias
    $folder = Resolve-Phase4FolderPath `
        -Mailbox $mailbox `
        -FolderPath @($fixture.Data.attachmentMessage.folderPath)
    $message = Find-Phase4Message `
        -Folder $folder `
        -SubjectMarker $fixture.Data.attachmentMessage.subjectMarker
    $result = Invoke-EndpointTool `
        -Name 'outlook_list_attachments' `
        -Arguments ([ordered]@{
            item = [ordered]@{
                storeId = $message.item.storeId
                entryId = $message.item.entryId
                itemClass = $message.item.itemClass
            }
            pageSize = 50
        })
    $envelope = Assert-Phase4SuccessEnvelope `
        -Result $result `
        -ExpectedDataFields @('attachments', 'nextCursor', 'resultTruncated') `
        -ExpectedPartial $false `
        -Context 'outlook_list_attachments'
    $actual = @($envelope.data.attachments)
    $expected = @($fixture.Data.attachmentMessage.expectedAttachments)
    if ($envelope.data.attachments -isnot [System.Array] -or
        $actual.Count -ne $expected.Count -or $null -ne $envelope.data.nextCursor -or
        $envelope.data.resultTruncated -ne $false -or @($envelope.warnings).Count -ne 0) {
        throw 'Attachment metadata did not match the expected bounded page.'
    }
    for ($index = 0; $index -lt $expected.Count; $index++) {
        Assert-Phase4ExactProperties `
            -InputObject $actual[$index] `
            -Expected @('attachment', 'contentType') `
            -Context 'attachment summary'
        Assert-Phase4ExactProperties `
            -InputObject $actual[$index].attachment `
            -Expected @('item', 'attachmentIndex', 'name', 'size', 'sizeIsKnown', 'metadataFingerprint') `
            -Context 'attachment reference'
        if ($actual[$index].attachment.name -cne $expected[$index].name -or
            (ConvertTo-Phase4Integer -Value $actual[$index].attachment.size -Context 'attachment size') -ne
                (ConvertTo-Phase4Integer -Value $expected[$index].size -Context 'expected attachment size') -or
            $actual[$index].attachment.attachmentIndex -ne ($index + 1) -or
            $actual[$index].attachment.sizeIsKnown -ne $true -or
            $actual[$index].attachment.metadataFingerprint -cnotmatch '^[0-9a-f]{64}$') {
            throw 'Attachment metadata changed from the local fixture.'
        }
    }
    return $actual.Count
}

function Test-Phase4LongBody {
    param([Parameter(Mandatory = $true)][object[]]$Mailboxes)

    $mailbox = Get-Phase4MailboxForAlias `
        -Mailboxes $Mailboxes `
        -Alias $fixture.Data.longBodyMessage.storeAlias
    $folder = Resolve-Phase4FolderPath `
        -Mailbox $mailbox `
        -FolderPath @($fixture.Data.longBodyMessage.folderPath)
    $message = Find-Phase4Message `
        -Folder $folder `
        -SubjectMarker $fixture.Data.longBodyMessage.subjectMarker
    $detail = Get-Phase4MessageDetail `
        -Message $message `
        -MaximumBodyCharacters 64 `
        -ExpectedPartial $true
    if ($detail.partial -ne $true -or
        @($detail.warnings).Count -ne 1 -or
        $detail.warnings[0] -cne 'The message body was truncated to the requested character limit.' -or
        $detail.data.body.isTruncated -ne $true -or
        $detail.data.body.isProtected -ne $false -or
        $detail.data.body.content -isnot [string] -or
        -not $detail.data.body.content.StartsWith(
            $fixture.Data.longBodyMessage.bodyPrefixMarker,
            [StringComparison]::Ordinal) -or
        (ConvertTo-Phase4Integer -Value $detail.data.body.originalCharacterCount -Context 'body character count') -lt
            (ConvertTo-Phase4Integer -Value $fixture.Data.longBodyMessage.minimumCharacterCount -Context 'fixture body count')) {
        throw 'The long body did not return the expected bounded truncation metadata.'
    }
}

function Test-Phase4ProtectedMessage {
    param([Parameter(Mandatory = $true)][object[]]$Mailboxes)

    if ($null -eq $fixture.Data.protectedMessage) {
        return $false
    }
    $mailbox = Get-Phase4MailboxForAlias `
        -Mailboxes $Mailboxes `
        -Alias $fixture.Data.protectedMessage.storeAlias
    $folder = Resolve-Phase4FolderPath `
        -Mailbox $mailbox `
        -FolderPath @($fixture.Data.protectedMessage.folderPath)
    $message = Find-Phase4Message `
        -Folder $folder `
        -SubjectMarker $fixture.Data.protectedMessage.subjectMarker
    $detail = Get-Phase4MessageDetail `
        -Message $message `
        -MaximumBodyCharacters 64 `
        -ExpectedPartial $true
    if ($detail.partial -ne $true -or
        @($detail.warnings).Count -ne 1 -or
        $detail.warnings[0] -cne 'The message body is protected and was not returned.' -or
        $detail.data.body.isProtected -ne $true -or
        $detail.data.body.isTruncated -ne $false -or
        $null -ne $detail.data.body.originalCharacterCount -or
        $detail.data.body.content -cne '') {
        throw 'The protected message body was not withheld safely.'
    }
    return $true
}

function Test-Phase4CancellationRecovery {
    param([Parameter(Mandatory = $true)][object]$LargeFolder)

    $requestId = Get-Phase4NextRequestId
    $body = New-Phase4JsonRpcBody `
        -Id $requestId `
        -Method 'tools/call' `
        -Params ([ordered]@{
            name = 'outlook_list_messages'
            arguments = [ordered]@{
                folder = [ordered]@{
                    storeId = $LargeFolder.storeId
                    entryId = $LargeFolder.entryId
                }
                pageSize = 50
            }
        })
    $cancellation = [Threading.CancellationTokenSource]::new()
    $cancelled = $false
    try {
        $pending = Start-Phase4McpPost `
            -Context $script:phase4Context `
            -Body $body `
            -CancellationToken $cancellation.Token
        $cancellation.Cancel()
        try {
            $null = Complete-Phase4McpPost -Pending $pending
        }
        catch [Threading.Tasks.TaskCanceledException] {
            $cancelled = $true
        }
        catch [OperationCanceledException] {
            $cancelled = $true
        }
    }
    finally {
        $cancellation.Dispose()
    }
    if (-not $cancelled) {
        throw 'The explicit Phase 4 client cancellation did not terminate its response wait.'
    }
    $null = Wait-Phase4ReadDiagnosticsQuiescent -Context 'post-cancellation status'
}

function Test-Phase4ConcurrentReads {
    param([Parameter(Mandatory = $true)][object]$LargeFolder)

    $pendingCalls = [System.Collections.Generic.List[object]]::new()
    try {
        for ($index = 0; $index -lt 4; $index++) {
            $requestId = Get-Phase4NextRequestId
            $body = New-Phase4JsonRpcBody `
                -Id $requestId `
                -Method 'tools/call' `
                -Params ([ordered]@{
                    name = 'outlook_list_messages'
                    arguments = [ordered]@{
                        folder = [ordered]@{
                            storeId = $LargeFolder.storeId
                            entryId = $LargeFolder.entryId
                        }
                        pageSize = 50
                    }
                })
            $pendingCalls.Add([pscustomobject]@{
                Id = $requestId
                Pending = Start-Phase4McpPost -Context $script:phase4Context -Body $body
                Consumed = $false
            })
        }
        Assert-EndpointResponsive
        $timeoutCount = 0
        foreach ($call in $pendingCalls) {
            try {
                $response = Complete-Phase4McpPost -Pending $call.Pending
            }
            finally {
                $call.Consumed = $true
            }
            $message = ConvertFrom-Phase4SseResponse -Response $response
            if ((ConvertTo-Phase4Integer -Value $message.id -Context 'concurrent response ID') -ne $call.Id) {
                throw 'A concurrent Phase 4 read did not preserve its request ID.'
            }
            $raw = $message.result
            Assert-Phase4ExactProperties `
                -InputObject $raw `
                -Expected @('content', 'structuredContent', 'isError') `
                -Context 'concurrent read result'
            $structured = $raw.structuredContent
            if ($raw.isError -isnot [bool] -or
                $structured.operationId -isnot [string] -or
                $structured.operationId -cnotmatch '^[0-9a-f]{32}$' -or
                -not $script:phase4OperationIds.Add($structured.operationId)) {
                throw 'A concurrent Phase 4 read reused an operation ID.'
            }
            Add-Phase4SensitiveValuesFromObject `
                -Set $script:phase4SensitiveValues `
                -InputObject $structured
            Assert-Phase4SafeTextContent `
                -RawResult $raw `
                -SensitiveValues $script:phase4SensitiveValues `
                -Context 'concurrent read'
            if ($raw.isError) {
                $wrapper = [pscustomobject]@{ IsError = $true; Structured = $structured; Raw = $raw }
                $null = Assert-Phase4ErrorEnvelope `
                    -Result $wrapper `
                    -ExpectedCode 'TIMEOUT' `
                    -Context 'concurrent large-folder timeout'
                $timeoutCount++
            }
            else {
                $wrapper = [pscustomobject]@{ IsError = $false; Structured = $structured; Raw = $raw }
                $null = Assert-Phase4MessagePage `
                    -Result $wrapper `
                    -Context 'concurrent large-folder read' `
                    -ExpectedPartial $false
            }
        }
        $null = Wait-Phase4ReadDiagnosticsQuiescent -Context 'post-concurrent-read status'
        Assert-EndpointResponsive
        return $timeoutCount
    }
    finally {
        foreach ($call in $pendingCalls) {
            if (-not $call.Consumed) {
                try {
                    $null = Complete-Phase4McpPost -Pending $call.Pending
                }
                catch {
                    $call.Pending.Request.Dispose()
                }
            }
        }
    }
}

function Test-Phase4RepeatedReads {
    param([Parameter(Mandatory = $true)][object]$LargeFolder)

    $before = Get-Phase4StatusEnvelope -Context 'pre-repeated-read status'
    $arguments = [ordered]@{
        folder = [ordered]@{ storeId = $LargeFolder.storeId; entryId = $LargeFolder.entryId }
        pageSize = 1
    }
    for ($index = 0; $index -lt 20; $index++) {
        $result = Invoke-EndpointTool -Name 'outlook_list_messages' -Arguments $arguments
        $null = Assert-Phase4MessagePage `
            -Result $result `
            -Context 'resource warm-up read' `
            -ExpectedPartial $false
    }

    $baseline = [System.Collections.Generic.List[object]]::new()
    for ($index = 0; $index -lt 3; $index++) {
        Assert-EndpointResponsive
        $baseline.Add((Get-Phase4ProcessResourceSample -OutlookProcessId $OutlookProcessId))
    }

    $samples = [System.Collections.Generic.List[object]]::new()
    $batchCount = [int]($RepeatedReadCount / 20)
    for ($batch = 0; $batch -lt $batchCount; $batch++) {
        for ($index = 0; $index -lt 20; $index++) {
            $result = Invoke-EndpointTool -Name 'outlook_list_messages' -Arguments $arguments
            $null = Assert-Phase4MessagePage `
                -Result $result `
                -Context 'repeated Phase 4 read' `
                -ExpectedPartial $false
        }
        Assert-EndpointResponsive
        $samples.Add((Get-Phase4ProcessResourceSample -OutlookProcessId $OutlookProcessId))
    }
    $evidence = Assert-Phase4ResourceStability `
        -BaselineSamples $baseline.ToArray() `
        -WorkloadSamples $samples.ToArray()
    $after = Get-Phase4StatusEnvelope -Context 'post-repeated-read status'
    $acquiredDelta = [int64]$after.data.readDiagnostics.comAcquired -
        [int64]$before.data.readDiagnostics.comAcquired
    $releasedDelta = [int64]$after.data.readDiagnostics.comReleased -
        [int64]$before.data.readDiagnostics.comReleased
    if ($acquiredDelta -le 0 -or $releasedDelta -ne $acquiredDelta) {
        throw 'Repeated Phase 4 reads did not produce balanced direct COM telemetry.'
    }
    return [pscustomobject]@{
        ReadCount = $RepeatedReadCount
        SampleCount = $samples.Count + $baseline.Count
        ResourceEvidence = $evidence
        ComAcquiredDelta = $acquiredDelta
        ComReleasedDelta = $releasedDelta
        ComOutstanding = [int64]$after.data.readDiagnostics.comOutstanding
        MaterializedItemHighWater = [int64]$after.data.readDiagnostics.materializedItemHighWater
    }
}

try {
    Invoke-Phase4Handshake
    Assert-EndpointResponsive
    $probe = Invoke-Phase4Probe
    $mailboxes = @(Get-Phase4Mailboxes)
    $knownMessageCount = Test-Phase4KnownMessages -Mailboxes $mailboxes
    Assert-EndpointResponsive

    $pagination = $null
    $largeFolder = $null
    $partialVerified = $false
    $conversationCount = 0
    $attachmentCount = 0
    $bodyTruncationVerified = $false
    $protectedVerified = $false
    $cancellationVerified = $false
    $concurrentReadCount = 0
    $structuredTimeoutCount = 0
    $resource = $null

    if ($Mode -eq 'Full') {
        $pagination = Test-Phase4StaticPagination -Mailboxes $mailboxes
        $largeFolderEvidence = Test-Phase4LargeFolder -Mailboxes $mailboxes
        $largeFolder = $largeFolderEvidence.Folder
        Test-Phase4PartialSearch -Mailboxes $mailboxes
        $partialVerified = $true
        $conversationCount = Test-Phase4Conversation -Mailboxes $mailboxes
        $attachmentCount = Test-Phase4Attachments -Mailboxes $mailboxes
        Test-Phase4LongBody -Mailboxes $mailboxes
        $bodyTruncationVerified = $true
        $protectedVerified = Test-Phase4ProtectedMessage -Mailboxes $mailboxes
        Test-Phase4CancellationRecovery -LargeFolder $largeFolder
        $cancellationVerified = $true
        $structuredTimeoutCount = Test-Phase4ConcurrentReads -LargeFolder $largeFolder
        $concurrentReadCount = 4
        $resource = Test-Phase4RepeatedReads -LargeFolder $largeFolder
    }

    Assert-EndpointResponsive
    $maximumUiMilliseconds = if ($script:phase4ResponsiveDurations.Count -eq 0) {
        0
    }
    else {
        ($script:phase4ResponsiveDurations | Measure-Object -Maximum).Maximum
    }

    [pscustomobject]@{
        Phase = 4
        CycleNumber = $CycleNumber
        FixtureSource = $fixture.Data.source
        Mode = $Mode
        FullWorkload = $Mode -eq 'Full'
        ToolCount = $script:Phase4ExpectedTools.Count
        StoreCount = $mailboxes.Count
        StoreInventoryMatched = $true
        StoreInventorySha256 = $expectedInventory.Sha256
        InboxStoreCount = @($fixture.Data.stores | Where-Object expectsInbox).Count
        KnownMessageReacquireCount = $knownMessageCount
        StaticPageCount = if ($null -eq $pagination) { 0 } else { $pagination.PageCount }
        StaticPageItemCount = if ($null -eq $pagination) { 0 } else { $pagination.ItemCount }
        CursorErrorCount = if ($null -eq $pagination) { 0 } else { $pagination.CursorErrorsVerified }
        LargeFolderMinimumItemCount = if ($Mode -eq 'Full') {
            [int64]$largeFolderEvidence.DistinctItemCount
        }
        else { 0 }
        PartialSearchVerified = $partialVerified
        ConversationMessageCount = $conversationCount
        AttachmentMetadataCount = $attachmentCount
        BodyTruncationVerified = $bodyTruncationVerified
        ProtectedBodyVerified = $protectedVerified
        ProtectedFixtureConfigured = $fixture.ProtectedMessageConfigured
        CancellationRecoveryVerified = $cancellationVerified
        ConcurrentReadCount = $concurrentReadCount
        LiveStructuredTimeoutCount = $structuredTimeoutCount
        RepeatedReadCount = if ($null -eq $resource) { 0 } else { $resource.ReadCount }
        ResourceSampleCount = if ($null -eq $resource) { 0 } else { $resource.SampleCount }
        ResourceStable = $null -ne $resource -and $resource.ResourceEvidence.ResourceStable
        ComAcquiredDelta = if ($null -eq $resource) { 0 } else { $resource.ComAcquiredDelta }
        ComReleasedDelta = if ($null -eq $resource) { 0 } else { $resource.ComReleasedDelta }
        ComOutstanding = if ($null -eq $resource) { 0 } else { $resource.ComOutstanding }
        MaterializedItemHighWater = if ($null -eq $resource) {
            0
        }
        else { $resource.MaterializedItemHighWater }
        MaximumHandleCount = if ($null -eq $resource) { 0 } else {
            $resource.ResourceEvidence.MaximumHandleCount
        }
        MaximumPrivateBytes = if ($null -eq $resource) { 0 } else {
            $resource.ResourceEvidence.MaximumPrivateBytes
        }
        MaximumGdiObjects = if ($null -eq $resource) { 0 } else {
            $resource.ResourceEvidence.MaximumGdiObjects
        }
        MaximumUserObjects = if ($null -eq $resource) { 0 } else {
            $resource.ResourceEvidence.MaximumUserObjects
        }
        OutlookResponsiveCheckCount = $script:phase4ResponsiveDurations.Count
        MaximumUiResponseMilliseconds = [double]$maximumUiMilliseconds
        AllProbeCallsOnCapturedSta = $probe.StaVerified
        UniqueOperationIdCount = $script:phase4OperationIds.Count - $operationCountBefore
        UniqueOperationIds = $true
    }
}
finally {
    $script:phase4Context.Client.Dispose()
}
