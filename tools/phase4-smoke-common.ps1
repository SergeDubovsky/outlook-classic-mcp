#Requires -Version 5.1

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:Phase4ExpectedTools = @(
    'outlook_status'
    'outlook_probe'
    'outlook_list_mailboxes'
    'outlook_list_folders'
    'outlook_list_messages'
    'outlook_search_messages'
    'outlook_get_message'
    'outlook_get_conversation'
    'outlook_list_attachments'
)
$script:Phase4StoreTypes = @(
    'primaryExchangeMailbox'
    'exchangeMailbox'
    'exchangePublicFolder'
    'additionalExchangeMailbox'
    'nonExchange'
    'unknown'
)
$script:Phase4MaximumFixtureBytes = 256 * 1024
$script:Phase4MaximumResponseCharacters = 1024 * 1024 + 4096
$script:Phase4ProtocolVersion = '2025-11-25'
$script:Phase4AllowedWarnings = @(
    'The message body is protected and was not returned.'
    'The message body was truncated to the requested character limit.'
    'One or more recipient lists were truncated to the supported maximum.'
    'The result page was truncated to the supported response-size limit.'
)

function Assert-Phase4ExactProperties {
    param(
        [Parameter(Mandatory = $true)][object]$InputObject,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Expected,
        [Parameter(Mandatory = $true)][string]$Context
    )

    if ($null -eq $InputObject) {
        throw "$Context is missing."
    }

    $actual = @($InputObject.PSObject.Properties.Name | Sort-Object)
    $wanted = @($Expected | Sort-Object)
    if (($actual -join "`n") -cne ($wanted -join "`n")) {
        throw "$Context does not contain the exact expected fields."
    }
}

function Get-Phase4OptionalProperty {
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }
    return $property.Value
}

function Assert-Phase4BoundedString {
    param(
        [AllowNull()][object]$Value,
        [Parameter(Mandatory = $true)][string]$Context,
        [ValidateRange(0, 65536)][int]$MinimumLength = 1,
        [ValidateRange(1, 65536)][int]$MaximumLength = 4096
    )

    if ($Value -isnot [string] -or
        $Value.Length -lt $MinimumLength -or
        $Value.Length -gt $MaximumLength -or
        $Value -match '[\x00-\x1F\x7F]') {
        throw "$Context is not a valid bounded string."
    }
}

function ConvertTo-Phase4Integer {
    param(
        [AllowNull()][object]$Value,
        [Parameter(Mandatory = $true)][string]$Context
    )

    if ($Value -is [byte] -or
        $Value -is [int16] -or
        $Value -is [int32] -or
        $Value -is [int64] -or
        $Value -is [uint16] -or
        $Value -is [uint32]) {
        return [int64]$Value
    }
    throw "$Context is not an integer."
}

function Add-Phase4SensitiveValue {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.HashSet[string]]$Set,
        [AllowNull()][object]$Value
    )

    if ($Value -is [string] -and $Value.Length -ge 8) {
        $null = $Set.Add($Value)
    }
}

function Add-Phase4SensitiveValuesFromObject {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.HashSet[string]]$Set,
        [AllowNull()][object]$InputObject,
        [string]$PropertyName = '',
        [ValidateRange(0, 32)][int]$Depth = 0
    )

    if ($null -eq $InputObject -or $Depth -gt 24) {
        return
    }
    if ($InputObject -is [string]) {
        if ($PropertyName -match '^(storeId|entryId|conversationId|nextCursor|metadataFingerprint|subject|senderDisplayName|senderAddress|displayName|address|name|content)$') {
            Add-Phase4SensitiveValue -Set $Set -Value $InputObject
        }
        return
    }
    if ($InputObject -is [System.Collections.IEnumerable] -and
        $InputObject -isnot [System.Management.Automation.PSCustomObject]) {
        foreach ($element in $InputObject) {
            Add-Phase4SensitiveValuesFromObject `
                -Set $Set `
                -InputObject $element `
                -PropertyName $PropertyName `
                -Depth ($Depth + 1)
        }
        return
    }
    foreach ($property in @($InputObject.PSObject.Properties)) {
        Add-Phase4SensitiveValuesFromObject `
            -Set $Set `
            -InputObject $property.Value `
            -PropertyName $property.Name `
            -Depth ($Depth + 1)
    }
}

function Assert-Phase4RestrictedAcl {
    param([Parameter(Mandatory = $true)][string]$Path)

    $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
    if (-not $acl.AreAccessRulesProtected) {
        throw 'The Phase 4 read fixture must use a protected, non-inherited ACL.'
    }

    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $system = 'S-1-5-18'
    $allowed = @($currentUser, $system)
    $rules = @($acl.GetAccessRules($true, $false, [Security.Principal.SecurityIdentifier]))
    foreach ($rule in $rules) {
        if ($rule.AccessControlType -ne [Security.AccessControl.AccessControlType]::Allow -or
            $rule.IdentityReference.Value -cnotin $allowed) {
            throw 'The Phase 4 read fixture ACL grants an unexpected identity or access type.'
        }
    }
    foreach ($sid in $allowed) {
        $fullControl = @(
            $rules |
                Where-Object {
                    $_.IdentityReference.Value -eq $sid -and
                    ($_.FileSystemRights -band [Security.AccessControl.FileSystemRights]::FullControl) -eq
                        [Security.AccessControl.FileSystemRights]::FullControl
                }
        )
        if ($fullControl.Count -eq 0) {
            throw 'The Phase 4 read fixture ACL is missing a required FullControl entry.'
        }
    }
}

function Assert-Phase4FolderPath {
    param(
        [AllowNull()][object]$Value,
        [Parameter(Mandatory = $true)][string]$Context,
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.HashSet[string]]$SensitiveValues
    )

    $parts = @($Value)
    if ($Value -isnot [System.Array] -or $parts.Count -lt 1 -or $parts.Count -gt 16) {
        throw "$Context must contain between one and sixteen folder names."
    }
    foreach ($part in $parts) {
        Assert-Phase4BoundedString `
            -Value $part `
            -Context $Context `
            -MinimumLength 8 `
            -MaximumLength 256
        Add-Phase4SensitiveValue -Set $SensitiveValues -Value $part
    }
}

function Get-Phase4KnownMessageExpectedProperties {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('classic-outlook-ui', 'conditional-vsto-seeder')]
        [string]$Source,

        [Parameter(Mandatory = $true)]
        [bool]$ExpectsInbox
    )

    if ($Source -ceq 'conditional-vsto-seeder' -and -not $ExpectsInbox) {
        return @('folderPath', 'subjectMarker', 'bodyMarker')
    }
    return @('subjectMarker', 'bodyMarker')
}

function Assert-Phase4Marker {
    param(
        [AllowNull()][object]$Value,
        [Parameter(Mandatory = $true)][string]$Context,
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.HashSet[string]]$SensitiveValues,
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.HashSet[string]]$UniqueMarkers
    )

    Assert-Phase4BoundedString `
        -Value $Value `
        -Context $Context `
        -MinimumLength 16 `
        -MaximumLength 1024
    if (-not $UniqueMarkers.Add($Value)) {
        throw 'Every Phase 4 seed marker must be unique.'
    }
    Add-Phase4SensitiveValue -Set $SensitiveValues -Value $Value
}

function Import-Phase4ReadFixture {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$RepositoryRoot,
        [Parameter(Mandatory = $true)][object]$ExpectedInventory
    )

    $resolvedPath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
    if (-not $resolvedPath.EndsWith('.local.json', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'The Phase 4 read fixture must use the ignored .local.json suffix.'
    }
    $file = Get-Item -LiteralPath $resolvedPath -ErrorAction Stop
    if (-not $file.PSIsContainer -and $file.Length -gt $script:Phase4MaximumFixtureBytes) {
        throw 'The Phase 4 read fixture exceeds its 256-KiB limit.'
    }
    if ($file.PSIsContainer) {
        throw 'The Phase 4 read fixture path is not a file.'
    }

    $rootWithSeparator = [IO.Path]::GetFullPath($RepositoryRoot).TrimEnd('\') + '\'
    $fullFixturePath = [IO.Path]::GetFullPath($resolvedPath)
    if ($fullFixturePath.StartsWith($rootWithSeparator, [StringComparison]::OrdinalIgnoreCase)) {
        & git -C $RepositoryRoot check-ignore --quiet -- $fullFixturePath
        if ($LASTEXITCODE -ne 0) {
            $global:LASTEXITCODE = 0
            throw 'The repository-local Phase 4 read fixture is not ignored by Git.'
        }
    }
    Assert-Phase4RestrictedAcl -Path $resolvedPath

    $fixture = Get-Content -LiteralPath $resolvedPath -Raw -Encoding UTF8 |
        ConvertFrom-Json -ErrorAction Stop
    Assert-Phase4ExactProperties `
        -InputObject $fixture `
        -Expected @(
            'schema', 'source', 'stores', 'pagination', 'largeFolder',
            'conversation', 'attachmentMessage', 'longBodyMessage', 'protectedMessage'
        ) `
        -Context 'Phase 4 read fixture'
    if ((ConvertTo-Phase4Integer -Value $fixture.schema -Context 'fixture schema') -ne 1 -or
        $fixture.source -cnotin @('classic-outlook-ui', 'conditional-vsto-seeder')) {
        throw 'The Phase 4 read fixture schema or source is unsupported.'
    }
    if ($ExpectedInventory.Source -cne $fixture.source) {
        throw 'The Phase 4 read fixture and independent store inventory must use the same source.'
    }

    $sensitiveValues = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $uniqueMarkers = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $aliases = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $stores = @($fixture.stores)
    $conditionalSource = $fixture.source -ceq 'conditional-vsto-seeder'
    if ($stores.Count -lt 2 -or $stores.Count -gt 64) {
        throw 'The Phase 4 read fixture must describe between two and sixty-four stores.'
    }
    $conditionalStoreExpectations = $null
    if ($conditionalSource) {
        $conditionalStoreExpectations = @{
            bootstrap_store = $true
            fixture_store_a = $false
            fixture_store_b = $false
        }
        if ($stores.Count -ne $conditionalStoreExpectations.Count) {
            throw 'The conditional Phase 4 fixture must contain the bootstrap store and two run PST stores.'
        }
    }

    $expectedInventoryEntries = @{}
    $consumedInventoryKeys =
        [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($entryJson in @($ExpectedInventory.Entries)) {
        $entry = $entryJson | ConvertFrom-Json -ErrorAction Stop
        Assert-Phase4BoundedString `
            -Value $entry.displayName `
            -Context 'Phase 4 inventory store display name' `
            -MinimumLength 8 `
            -MaximumLength 256
        $key = "$($entry.displayName)$([char]31)$($entry.storeType)"
        Add-Phase4SensitiveValue -Set $sensitiveValues -Value $entry.displayName
        if (-not $expectedInventoryEntries.ContainsKey($key)) {
            $expectedInventoryEntries[$key] = 0
        }
        $expectedInventoryEntries[$key]++
    }

    foreach ($store in $stores) {
        Assert-Phase4ExactProperties `
            -InputObject $store `
            -Expected @('alias', 'displayName', 'storeType', 'expectsInbox', 'knownMessage') `
            -Context 'fixture store'
        Assert-Phase4BoundedString -Value $store.alias -Context 'store alias' -MaximumLength 64
        if ($store.alias -cnotmatch '^[a-z][a-z0-9_-]{0,63}$' -or -not $aliases.Add($store.alias)) {
            throw 'Fixture store aliases must be unique lower-case identifiers.'
        }
        Assert-Phase4BoundedString `
            -Value $store.displayName `
            -Context 'store display name' `
            -MinimumLength 8 `
            -MaximumLength 256
        if ($store.storeType -isnot [string] -or $store.storeType -cnotin $script:Phase4StoreTypes -or
            $store.expectsInbox -isnot [bool]) {
            throw 'A fixture store has an invalid type or expectsInbox value.'
        }
        if ($conditionalSource -and
            (-not $conditionalStoreExpectations.ContainsKey($store.alias) -or
             $store.expectsInbox -ne $conditionalStoreExpectations[$store.alias] -or
             $store.storeType -cne 'nonExchange')) {
            throw 'The conditional Phase 4 fixture has an invalid store alias, type, or Inbox expectation.'
        }
        $inventoryKey = "$($store.displayName)$([char]31)$($store.storeType)"
        if (-not $expectedInventoryEntries.ContainsKey($inventoryKey) -or
            $expectedInventoryEntries[$inventoryKey] -ne 1 -or
            -not $consumedInventoryKeys.Add($inventoryKey)) {
            throw 'Each Phase 4 fixture store must uniquely match the independent store inventory.'
        }
        Add-Phase4SensitiveValue -Set $sensitiveValues -Value $store.displayName

        $knownMessageProperties = @(
            Get-Phase4KnownMessageExpectedProperties `
                -Source $fixture.source `
                -ExpectsInbox $store.expectsInbox)
        Assert-Phase4ExactProperties `
            -InputObject $store.knownMessage `
            -Expected $knownMessageProperties `
            -Context 'known message fixture'
        if ($conditionalSource -and -not $store.expectsInbox) {
            Assert-Phase4FolderPath `
                -Value $store.knownMessage.folderPath `
                -Context 'known-message folder path' `
                -SensitiveValues $sensitiveValues
        }
        Assert-Phase4Marker `
            -Value $store.knownMessage.subjectMarker `
            -Context 'known message subject marker' `
            -SensitiveValues $sensitiveValues `
            -UniqueMarkers $uniqueMarkers
        Assert-Phase4Marker `
            -Value $store.knownMessage.bodyMarker `
            -Context 'known message body marker' `
            -SensitiveValues $sensitiveValues `
            -UniqueMarkers $uniqueMarkers
    }
    if ($stores.Count -ne $ExpectedInventory.Count -or
        $consumedInventoryKeys.Count -ne $expectedInventoryEntries.Count) {
        throw 'The Phase 4 fixture must cover the complete independent store inventory exactly once.'
    }
    $expectedInboxStoreCount = @($stores | Where-Object expectsInbox).Count
    if (($conditionalSource -and $expectedInboxStoreCount -lt 1) -or
        (-not $conditionalSource -and $expectedInboxStoreCount -lt 2)) {
        throw 'The Phase 4 fixture has insufficient independently expected standard Inboxes.'
    }
    if (@($stores | Where-Object { $null -ne $_.knownMessage }).Count -lt 2) {
        throw 'The Phase 4 fixture must seed known messages in at least two stores.'
    }

    Assert-Phase4ExactProperties `
        -InputObject $fixture.pagination `
        -Expected @('storeAlias', 'folderPath', 'pageSize', 'orderedSubjectMarkers') `
        -Context 'pagination fixture'
    if (-not $aliases.Contains($fixture.pagination.storeAlias) -or
        (ConvertTo-Phase4Integer -Value $fixture.pagination.pageSize -Context 'pagination page size') -ne 4) {
        throw 'The pagination fixture must select a known store and use pageSize 4.'
    }
    Assert-Phase4FolderPath `
        -Value $fixture.pagination.folderPath `
        -Context 'pagination folder path' `
        -SensitiveValues $sensitiveValues
    $pageMarkers = @($fixture.pagination.orderedSubjectMarkers)
    if ($fixture.pagination.orderedSubjectMarkers -isnot [System.Array] -or
        $pageMarkers.Count -ne 12) {
        throw 'The pagination fixture must contain exactly twelve ordered markers.'
    }
    foreach ($marker in $pageMarkers) {
        Assert-Phase4Marker `
            -Value $marker `
            -Context 'pagination subject marker' `
            -SensitiveValues $sensitiveValues `
            -UniqueMarkers $uniqueMarkers
    }

    Assert-Phase4ExactProperties `
        -InputObject $fixture.largeFolder `
        -Expected @('storeAlias', 'folderPath', 'minimumItemCount') `
        -Context 'large-folder fixture'
    if (-not $aliases.Contains($fixture.largeFolder.storeAlias) -or
        (ConvertTo-Phase4Integer -Value $fixture.largeFolder.minimumItemCount -Context 'large-folder count') -ne 1001) {
        throw 'The large-folder fixture must select a known store and declare the bounded 1,001-item live threshold.'
    }
    Assert-Phase4FolderPath `
        -Value $fixture.largeFolder.folderPath `
        -Context 'large-folder path' `
        -SensitiveValues $sensitiveValues

    Assert-Phase4ExactProperties `
        -InputObject $fixture.conversation `
        -Expected @('storeAlias', 'folderPath', 'seedSubjectMarker', 'expectedSubjectMarkers') `
        -Context 'conversation fixture'
    if (-not $aliases.Contains($fixture.conversation.storeAlias)) {
        throw 'The conversation fixture references an unknown store alias.'
    }
    Assert-Phase4FolderPath `
        -Value $fixture.conversation.folderPath `
        -Context 'conversation folder path' `
        -SensitiveValues $sensitiveValues
    Assert-Phase4Marker `
        -Value $fixture.conversation.seedSubjectMarker `
        -Context 'conversation seed marker' `
        -SensitiveValues $sensitiveValues `
        -UniqueMarkers $uniqueMarkers
    $conversationMarkers = @($fixture.conversation.expectedSubjectMarkers)
    if ($fixture.conversation.expectedSubjectMarkers -isnot [System.Array]) {
        throw 'The conversation expectedSubjectMarkers value must be an array.'
    }
    $minimumConversationMarkers = if ($fixture.source -ceq 'conditional-vsto-seeder') { 1 } else { 2 }
    $maximumConversationMarkers = if ($fixture.source -ceq 'conditional-vsto-seeder') { 1 } else { 50 }
    if ($conversationMarkers.Count -lt $minimumConversationMarkers -or
        $conversationMarkers.Count -gt $maximumConversationMarkers) {
        throw 'The conversation marker count does not match the declared fixture source mode.'
    }
    foreach ($marker in $conversationMarkers) {
        Assert-Phase4Marker `
            -Value $marker `
            -Context 'conversation message marker' `
            -SensitiveValues $sensitiveValues `
            -UniqueMarkers $uniqueMarkers
    }

    Assert-Phase4ExactProperties `
        -InputObject $fixture.attachmentMessage `
        -Expected @('storeAlias', 'folderPath', 'subjectMarker', 'expectedAttachments') `
        -Context 'attachment fixture'
    if (-not $aliases.Contains($fixture.attachmentMessage.storeAlias)) {
        throw 'The attachment fixture references an unknown store alias.'
    }
    Assert-Phase4FolderPath `
        -Value $fixture.attachmentMessage.folderPath `
        -Context 'attachment folder path' `
        -SensitiveValues $sensitiveValues
    Assert-Phase4Marker `
        -Value $fixture.attachmentMessage.subjectMarker `
        -Context 'attachment message marker' `
        -SensitiveValues $sensitiveValues `
        -UniqueMarkers $uniqueMarkers
    $attachments = @($fixture.attachmentMessage.expectedAttachments)
    if ($fixture.attachmentMessage.expectedAttachments -isnot [System.Array] -or
        $attachments.Count -lt 2 -or $attachments.Count -gt 50) {
        throw 'The attachment fixture must contain between two and fifty attachments.'
    }
    foreach ($attachment in $attachments) {
        Assert-Phase4ExactProperties `
            -InputObject $attachment `
            -Expected @('name', 'size') `
            -Context 'attachment expectation'
        Assert-Phase4BoundedString `
            -Value $attachment.name `
            -Context 'attachment name' `
            -MinimumLength 8 `
            -MaximumLength 255
        if ((ConvertTo-Phase4Integer -Value $attachment.size -Context 'attachment size') -lt 0) {
            throw 'An expected attachment size cannot be negative.'
        }
        Add-Phase4SensitiveValue -Set $sensitiveValues -Value $attachment.name
    }

    Assert-Phase4ExactProperties `
        -InputObject $fixture.longBodyMessage `
        -Expected @('storeAlias', 'folderPath', 'subjectMarker', 'bodyPrefixMarker', 'minimumCharacterCount') `
        -Context 'long-body fixture'
    if (-not $aliases.Contains($fixture.longBodyMessage.storeAlias) -or
        (ConvertTo-Phase4Integer -Value $fixture.longBodyMessage.minimumCharacterCount -Context 'long-body count') -le 64) {
        throw 'The long-body fixture is invalid.'
    }
    Assert-Phase4FolderPath `
        -Value $fixture.longBodyMessage.folderPath `
        -Context 'long-body folder path' `
        -SensitiveValues $sensitiveValues
    Assert-Phase4Marker `
        -Value $fixture.longBodyMessage.subjectMarker `
        -Context 'long-body subject marker' `
        -SensitiveValues $sensitiveValues `
        -UniqueMarkers $uniqueMarkers
    Assert-Phase4Marker `
        -Value $fixture.longBodyMessage.bodyPrefixMarker `
        -Context 'long-body prefix marker' `
        -SensitiveValues $sensitiveValues `
        -UniqueMarkers $uniqueMarkers
    if ($fixture.longBodyMessage.bodyPrefixMarker.Length -gt 64) {
        throw 'The long-body prefix marker must fit inside the 64-character truncation probe.'
    }

    if ($null -ne $fixture.protectedMessage) {
        Assert-Phase4ExactProperties `
            -InputObject $fixture.protectedMessage `
            -Expected @('storeAlias', 'folderPath', 'subjectMarker') `
            -Context 'protected-message fixture'
        if (-not $aliases.Contains($fixture.protectedMessage.storeAlias)) {
            throw 'The protected-message fixture references an unknown store alias.'
        }
        Assert-Phase4FolderPath `
            -Value $fixture.protectedMessage.folderPath `
            -Context 'protected-message folder path' `
            -SensitiveValues $sensitiveValues
        Assert-Phase4Marker `
            -Value $fixture.protectedMessage.subjectMarker `
            -Context 'protected-message subject marker' `
            -SensitiveValues $sensitiveValues `
            -UniqueMarkers $uniqueMarkers
    }

    return [pscustomobject]@{
        Path = $resolvedPath
        Data = $fixture
        SensitiveValues = $sensitiveValues
        StoreCount = $stores.Count
        ProtectedMessageConfigured = $null -ne $fixture.protectedMessage
    }
}

function New-Phase4HttpContext {
    param(
        [Parameter(Mandatory = $true)][Uri]$Endpoint,
        [Parameter(Mandatory = $true)][string]$Token,
        [ValidateRange(1, 180)][int]$TimeoutSeconds = 30
    )

    if ($Endpoint.AbsoluteUri -cne 'http://127.0.0.1:8765/mcp/') {
        throw 'The Phase 4 endpoint harness accepts only the canonical loopback endpoint.'
    }
    if ($Token -cnotmatch '^[A-Za-z0-9_-]{43}$') {
        throw 'The Phase 4 endpoint harness requires a canonical bearer token.'
    }
    Add-Type -AssemblyName System.Net.Http
    $handler = [Net.Http.HttpClientHandler]::new()
    $handler.UseProxy = $false
    $handler.AllowAutoRedirect = $false
    $handler.AutomaticDecompression = [Net.DecompressionMethods]::None
    $handler.MaxConnectionsPerServer = 4
    $client = [Net.Http.HttpClient]::new($handler, $true)
    $client.Timeout = [TimeSpan]::FromSeconds($TimeoutSeconds)
    return [pscustomobject]@{
        Endpoint = $Endpoint
        Token = $Token
        Client = $client
    }
}

function Start-Phase4McpPost {
    param(
        [Parameter(Mandatory = $true)][object]$Context,
        [Parameter(Mandatory = $true)][string]$Body,
        [bool]$Authenticated = $true,
        [AllowNull()][string]$ProtocolVersion = $script:Phase4ProtocolVersion,
        [Threading.CancellationToken]$CancellationToken = [Threading.CancellationToken]::None
    )

    $request = [Net.Http.HttpRequestMessage]::new(
        [Net.Http.HttpMethod]::Post,
        $Context.Endpoint)
    try {
        $request.Content = [Net.Http.StringContent]::new(
            $Body,
            [Text.Encoding]::UTF8,
            'application/json')
        $request.Headers.Host = '127.0.0.1:8765'
        if (-not $request.Headers.TryAddWithoutValidation(
                'Accept',
                'application/json, text/event-stream')) {
            throw 'Could not set the required MCP headers.'
        }
        if (-not [string]::IsNullOrEmpty($ProtocolVersion) -and
            -not $request.Headers.TryAddWithoutValidation(
                'MCP-Protocol-Version',
                $ProtocolVersion)) {
            throw 'Could not set the MCP protocol-version header.'
        }
        if ($Authenticated -and
            -not $request.Headers.TryAddWithoutValidation(
                'Authorization',
                "Bearer $($Context.Token)")) {
            throw 'Could not set the MCP authorization header.'
        }
        return [pscustomobject]@{
            Request = $request
            Task = $Context.Client.SendAsync($request, $CancellationToken)
        }
    }
    catch {
        $request.Dispose()
        throw
    }
}

function Complete-Phase4McpPost {
    param([Parameter(Mandatory = $true)][object]$Pending)

    $response = $null
    try {
        $response = $Pending.Task.GetAwaiter().GetResult()
        $body = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        if ($body.Length -gt $script:Phase4MaximumResponseCharacters) {
            throw 'An MCP response exceeded the Phase 4 validation limit.'
        }
        $contentType = if ($null -eq $response.Content.Headers.ContentType) {
            $null
        }
        else {
            $response.Content.Headers.ContentType.MediaType
        }
        $cacheControl = if ($null -eq $response.Headers.CacheControl) {
            @()
        }
        else {
            @(
                $response.Headers.CacheControl.ToString() -split ',' |
                    ForEach-Object { $_.Trim() }
            )
        }
        return [pscustomobject]@{
            StatusCode = [int]$response.StatusCode
            ContentType = $contentType
            CacheControl = $cacheControl
            ContentEncoding = @($response.Content.Headers.ContentEncoding)
            Body = $body
        }
    }
    finally {
        if ($null -ne $response) {
            $response.Dispose()
        }
        $Pending.Request.Dispose()
    }
}

function ConvertFrom-Phase4SseResponse {
    param([Parameter(Mandatory = $true)][object]$Response)

    if ($Response.StatusCode -ne 200 -or $Response.ContentType -cne 'text/event-stream') {
        throw 'The Phase 4 endpoint did not return the expected authenticated SSE response.'
    }
    $cache = @($Response.CacheControl | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($cache.Count -ne 2 -or $cache -cnotcontains 'no-cache' -or $cache -cnotcontains 'no-store' -or
        $Response.ContentEncoding.Count -ne 1 -or $Response.ContentEncoding[0] -cne 'identity') {
        throw 'The Phase 4 SSE response headers are not fail-closed.'
    }
    $dataLines = @(
        $Response.Body.Replace("`r`n", "`n").Split("`n") |
            Where-Object { $_.StartsWith('data: ', [StringComparison]::Ordinal) }
    )
    if ($dataLines.Count -ne 1) {
        throw 'The Phase 4 SSE response did not contain exactly one data event.'
    }
    try {
        return $dataLines[0].Substring(6) | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw 'The Phase 4 SSE data event was not valid JSON.'
    }
}

function New-Phase4JsonRpcBody {
    param(
        [Parameter(Mandatory = $true)][long]$Id,
        [Parameter(Mandatory = $true)][string]$Method,
        [AllowNull()][object]$Params
    )

    $request = [ordered]@{
        jsonrpc = '2.0'
        id = $Id
        method = $Method
    }
    if ($null -ne $Params) {
        $request.params = $Params
    }
    return $request | ConvertTo-Json -Depth 32 -Compress
}

function Invoke-Phase4McpRequest {
    param(
        [Parameter(Mandatory = $true)][object]$Context,
        [Parameter(Mandatory = $true)][long]$Id,
        [Parameter(Mandatory = $true)][string]$Method,
        [AllowNull()][object]$Params
    )

    $body = New-Phase4JsonRpcBody -Id $Id -Method $Method -Params $Params
    $pending = Start-Phase4McpPost -Context $Context -Body $body
    $response = Complete-Phase4McpPost -Pending $pending
    $message = ConvertFrom-Phase4SseResponse -Response $response
    if ($message.jsonrpc -cne '2.0' -or
        (ConvertTo-Phase4Integer -Value $message.id -Context 'JSON-RPC response ID') -ne $Id -or
        $null -eq $message.result) {
        throw 'The Phase 4 MCP response did not preserve its request identity.'
    }
    return $message.result
}

function Invoke-Phase4ToolCall {
    param(
        [Parameter(Mandatory = $true)][object]$Context,
        [Parameter(Mandatory = $true)][long]$Id,
        [Parameter(Mandatory = $true)][string]$ToolName,
        [Parameter(Mandatory = $true)][object]$Arguments,
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.HashSet[string]]$SeenOperationIds,
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.HashSet[string]]$SensitiveValues
    )

    $result = Invoke-Phase4McpRequest `
        -Context $Context `
        -Id $Id `
        -Method 'tools/call' `
        -Params ([ordered]@{ name = $ToolName; arguments = $Arguments })
    Assert-Phase4ExactProperties `
        -InputObject $result `
        -Expected @('content', 'structuredContent', 'isError') `
        -Context "$ToolName result"
    if ($result.isError -isnot [bool] -or $null -eq $result.structuredContent) {
        throw "$ToolName returned an invalid MCP tool result."
    }
    $structured = $result.structuredContent
    if ($structured.operationId -isnot [string] -or
        $structured.operationId -cnotmatch '^[0-9a-f]{32}$' -or
        -not $SeenOperationIds.Add($structured.operationId)) {
        throw 'A Phase 4 operation ID was missing, invalid, or reused.'
    }
    Add-Phase4SensitiveValuesFromObject -Set $SensitiveValues -InputObject $structured
    Assert-Phase4SafeTextContent `
        -RawResult $result `
        -SensitiveValues $SensitiveValues `
        -Context $ToolName
    return [pscustomobject]@{
        IsError = $result.isError
        Structured = $structured
        Raw = $result
    }
}

function Assert-Phase4SafeTextContent {
    param(
        [Parameter(Mandatory = $true)][object]$RawResult,
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.HashSet[string]]$SensitiveValues,
        [Parameter(Mandatory = $true)][string]$Context
    )

    $blocks = @($RawResult.content)
    if ($RawResult.content -isnot [System.Array] -or $blocks.Count -ne 1) {
        throw "$Context did not return exactly one bounded text content block."
    }
    Assert-Phase4ExactProperties `
        -InputObject $blocks[0] `
        -Expected @('type', 'text') `
        -Context "$Context text content block"
    $text = $blocks[0].text
    if ($blocks[0].type -cne 'text' -or $text -isnot [string] -or
        $text.Length -lt 1 -or $text.Length -gt 4096) {
        throw "$Context returned invalid bounded text content."
    }
    foreach ($sensitiveValue in $SensitiveValues) {
        if ($text.IndexOf($sensitiveValue, [StringComparison]::Ordinal) -ge 0) {
            throw "$Context exposed mailbox-derived data in its text content block."
        }
    }
}

function Assert-Phase4SuccessEnvelope {
    param(
        [Parameter(Mandatory = $true)][object]$Result,
        [Parameter(Mandatory = $true)][string[]]$ExpectedDataFields,
        [Parameter(Mandatory = $true)][string]$Context,
        [AllowNull()][bool]$ExpectedPartial
    )

    if ($Result.IsError) {
        throw "$Context returned a tool error."
    }
    $structured = $Result.Structured
    Assert-Phase4ExactProperties `
        -InputObject $structured `
        -Expected @('ok', 'operationId', 'data', 'partial', 'warnings') `
        -Context "$Context success envelope"
    Assert-Phase4ExactProperties `
        -InputObject $structured.data `
        -Expected $ExpectedDataFields `
        -Context "$Context data"
    if ($structured.ok -isnot [bool] -or -not $structured.ok -or
        $structured.partial -isnot [bool] -or
        $null -eq $structured.warnings) {
        throw "$Context returned an invalid success envelope."
    }
    $warnings = @($structured.warnings)
    if ($structured.warnings -isnot [System.Array] -or
        $warnings.Count -gt 3 -or
        @($warnings | Select-Object -Unique).Count -ne $warnings.Count) {
        throw "$Context returned an invalid warning set."
    }
    foreach ($warning in $warnings) {
        if ($warning -isnot [string] -or $warning -cnotin $script:Phase4AllowedWarnings) {
            throw "$Context returned an unsupported warning."
        }
    }
    if ($PSBoundParameters.ContainsKey('ExpectedPartial') -and
        $structured.partial -ne $ExpectedPartial) {
        throw "$Context returned an unexpected partial flag."
    }
    return $structured
}

function Assert-Phase4ErrorEnvelope {
    param(
        [Parameter(Mandatory = $true)][object]$Result,
        [Parameter(Mandatory = $true)][string]$ExpectedCode,
        [Parameter(Mandatory = $true)][string]$Context
    )

    if (-not $Result.IsError) {
        throw "$Context did not return a tool error."
    }
    $structured = $Result.Structured
    Assert-Phase4ExactProperties `
        -InputObject $structured `
        -Expected @('ok', 'operationId', 'error') `
        -Context "$Context error envelope"
    Assert-Phase4ExactProperties `
        -InputObject $structured.error `
        -Expected @('code', 'message', 'retryable', 'details') `
        -Context "$Context error"
    if ($structured.ok -isnot [bool] -or $structured.ok -or
        $structured.error.code -cne $ExpectedCode -or
        $structured.error.retryable -isnot [bool] -or
        @($structured.error.details.PSObject.Properties).Count -ne 0) {
        throw "$Context returned an unexpected structured error."
    }
    return $structured
}

function Get-Phase4LocatorKey {
    param([Parameter(Mandatory = $true)][object]$Item)

    return "$($Item.storeId)$([char]31)$($Item.entryId)$([char]31)$($Item.itemClass)"
}

function Initialize-Phase4NativeMethods {
    if ('OutlookClassicMcp.Tools.Phase4NativeMethods' -as [type]) {
        return
    }
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace OutlookClassicMcp.Tools
{
    public static class Phase4NativeMethods
    {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SendMessageTimeout(
            IntPtr window,
            uint message,
            UIntPtr wParam,
            IntPtr lParam,
            uint flags,
            uint timeoutMilliseconds,
            out UIntPtr result);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetGuiResources(IntPtr process, uint flags);
    }
}
'@
}

function Assert-Phase4OutlookResponsive {
    param(
        [Parameter(Mandatory = $true)][int]$OutlookProcessId,
        [ValidateRange(100, 5000)][int]$TimeoutMilliseconds = 2000
    )

    Initialize-Phase4NativeMethods
    $process = Get-Process -Id $OutlookProcessId -ErrorAction Stop
    try {
        $process.Refresh()
        if ($process.ProcessName -cne 'OUTLOOK' -or $process.HasExited -or
            $process.MainWindowHandle -eq [IntPtr]::Zero -or -not $process.Responding) {
            throw 'Outlook is not responsive during the Phase 4 read workload.'
        }
        $result = [UIntPtr]::Zero
        $stopwatch = [Diagnostics.Stopwatch]::StartNew()
        $sendResult = [OutlookClassicMcp.Tools.Phase4NativeMethods]::SendMessageTimeout(
            $process.MainWindowHandle,
            0,
            [UIntPtr]::Zero,
            [IntPtr]::Zero,
            2,
            [uint32]$TimeoutMilliseconds,
            [ref]$result)
        $stopwatch.Stop()
        if ($sendResult -eq [IntPtr]::Zero) {
            throw 'Outlook did not answer the Phase 4 UI responsiveness probe.'
        }
        return $stopwatch.Elapsed.TotalMilliseconds
    }
    finally {
        $process.Dispose()
    }
}

function Get-Phase4ProcessResourceSample {
    param([Parameter(Mandatory = $true)][int]$OutlookProcessId)

    Initialize-Phase4NativeMethods
    $process = Get-Process -Id $OutlookProcessId -ErrorAction Stop
    try {
        $process.Refresh()
        return [pscustomobject]@{
            HandleCount = [int64]$process.HandleCount
            PrivateBytes = [int64]$process.PrivateMemorySize64
            WorkingSet = [int64]$process.WorkingSet64
            GdiObjects = [int64][OutlookClassicMcp.Tools.Phase4NativeMethods]::GetGuiResources(
                $process.Handle,
                0)
            UserObjects = [int64][OutlookClassicMcp.Tools.Phase4NativeMethods]::GetGuiResources(
                $process.Handle,
                1)
        }
    }
    finally {
        $process.Dispose()
    }
}

function Test-Phase4StrictMonotonicGrowth {
    param(
        [Parameter(Mandatory = $true)][object[]]$Samples,
        [Parameter(Mandatory = $true)][string]$Property
    )

    if ($Samples.Count -lt 3) {
        throw 'Phase 4 resource analysis requires at least three samples.'
    }
    $increased = $false
    for ($index = 1; $index -lt $Samples.Count; $index++) {
        $previous = [int64]$Samples[$index - 1].$Property
        $current = [int64]$Samples[$index].$Property
        if ($current -lt $previous) {
            return $false
        }
        if ($current -gt $previous) {
            $increased = $true
        }
    }
    return $increased
}

function Get-Phase4Median {
    param([Parameter(Mandatory = $true)][int64[]]$Values)

    $ordered = @($Values | Sort-Object)
    if ($ordered.Count -eq 0) {
        throw 'Cannot calculate an empty median.'
    }
    $middle = [int]($ordered.Count / 2)
    if (($ordered.Count % 2) -eq 1) {
        return [double]$ordered[$middle]
    }
    return ([double]$ordered[$middle - 1] + [double]$ordered[$middle]) / 2.0
}

function Assert-Phase4ResourceStability {
    param(
        [Parameter(Mandatory = $true)][object[]]$BaselineSamples,
        [Parameter(Mandatory = $true)][object[]]$WorkloadSamples
    )

    if ($BaselineSamples.Count -ne 3 -or $WorkloadSamples.Count -lt 3) {
        throw 'Phase 4 resource stability evidence is incomplete.'
    }
    foreach ($property in @('HandleCount', 'GdiObjects', 'UserObjects')) {
        if (Test-Phase4StrictMonotonicGrowth -Samples $WorkloadSamples -Property $property) {
            throw "Outlook $property grew monotonically during repeated Phase 4 reads."
        }
    }

    $baselinePrivate = Get-Phase4Median -Values ([int64[]]@(
        $BaselineSamples | ForEach-Object { $_.PrivateBytes }))
    $tail = @($WorkloadSamples | Select-Object -Last 3)
    $tailPrivate = Get-Phase4Median -Values ([int64[]]@(
        $tail | ForEach-Object { $_.PrivateBytes }))
    $allowedGrowth = [Math]::Max(32MB, [int64]($baselinePrivate * 0.10))
    if (($tailPrivate - $baselinePrivate) -gt $allowedGrowth) {
        throw 'Outlook private bytes did not remain within the Phase 4 post-warm bound.'
    }

    return [pscustomobject]@{
        ResourceStable = $true
        MaximumHandleCount = [int64](
            $WorkloadSamples | Measure-Object -Property HandleCount -Maximum).Maximum
        MaximumPrivateBytes = [int64](
            $WorkloadSamples | Measure-Object -Property PrivateBytes -Maximum).Maximum
        MaximumGdiObjects = [int64](
            $WorkloadSamples | Measure-Object -Property GdiObjects -Maximum).Maximum
        MaximumUserObjects = [int64](
            $WorkloadSamples | Measure-Object -Property UserObjects -Maximum).Maximum
    }
}
