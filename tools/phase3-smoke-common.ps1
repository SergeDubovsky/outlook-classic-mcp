#Requires -Version 5.1

Set-StrictMode -Version Latest

$script:Phase3ArchiveWarning =
    'Archive availability is not exposed by the Outlook Object Model.'
$script:Phase3StoreMetadataIncompleteWarning =
    'Some store metadata could not be determined.'
$script:Phase3StoreTypes = @(
    'primaryExchangeMailbox'
    'exchangeMailbox'
    'exchangePublicFolder'
    'additionalExchangeMailbox'
    'nonExchange'
    'unknown'
)
$script:Phase3FolderStatuses = @('available', 'missing', 'unknown')

function Assert-Phase3ExactProperties {
    param(
        [Parameter(Mandatory = $true)][object]$InputObject,
        [Parameter(Mandatory = $true)][string[]]$Expected,
        [Parameter(Mandatory = $true)][string]$Context
    )

    if ($null -eq $InputObject -or $InputObject -is [string]) {
        throw "$Context was not an object."
    }

    $actual = @($InputObject.PSObject.Properties.Name | Sort-Object)
    $wanted = @($Expected | Sort-Object)
    if (($actual -join "`n") -cne ($wanted -join "`n")) {
        throw "$Context did not contain the exact expected fields."
    }
}

function Assert-Phase3BoundedText {
    param(
        [AllowNull()][object]$Value,
        [Parameter(Mandatory = $true)][ValidateRange(1, 4096)][int]$MaximumLength,
        [Parameter(Mandatory = $true)][string]$Context
    )

    if ($Value -isnot [string] -or
        [string]::IsNullOrWhiteSpace($Value) -or
        $Value.Length -gt $MaximumLength -or
        $Value -match '[\x00-\x1F\x7F]') {
        throw "$Context was not valid bounded text."
    }
}

function ConvertTo-Phase3Integer {
    param(
        [AllowNull()][object]$Value,
        [Parameter(Mandatory = $true)][long]$Minimum,
        [Parameter(Mandatory = $true)][long]$Maximum,
        [Parameter(Mandatory = $true)][string]$Context
    )

    $integerTypes = @(
        [sbyte], [byte], [int16], [uint16], [int32], [uint32],
        [int64], [uint64]
    )
    if ($null -eq $Value -or $Value.GetType() -notin $integerTypes) {
        throw "$Context was not an integer."
    }

    try {
        $converted = [Convert]::ToInt64($Value, [Globalization.CultureInfo]::InvariantCulture)
    }
    catch {
        throw "$Context was outside the supported integer range."
    }
    if ($converted -lt $Minimum -or $converted -gt $Maximum) {
        throw "$Context was outside the supported integer range."
    }
    return $converted
}

function Get-Phase3Sha256 {
    param([Parameter(Mandatory = $true)][string]$Text)

    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($Text)
        return ([BitConverter]::ToString($sha256.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha256.Dispose()
    }
}

function Get-Phase3StringMultiset {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Entries)

    $sorted = [System.Collections.Generic.List[string]]::new()
    foreach ($entry in $Entries) {
        if ($null -eq $entry) {
            throw 'A canonical Phase 3 entry was null.'
        }
        $sorted.Add($entry)
    }
    $sorted.Sort([StringComparer]::Ordinal)
    $canonicalJson = ConvertTo-Json -InputObject ([string[]]$sorted.ToArray()) -Compress
    return [pscustomobject]@{
        Count = $sorted.Count
        Entries = [string[]]$sorted.ToArray()
        Sha256 = Get-Phase3Sha256 -Text $canonicalJson
    }
}

function New-Phase3InventoryEntry {
    param(
        [Parameter(Mandatory = $true)][string]$DisplayName,
        [Parameter(Mandatory = $true)][string]$StoreType
    )

    $normalizedName = $DisplayName.Normalize([Text.NormalizationForm]::FormC)
    return ConvertTo-Json -InputObject ([ordered]@{
        displayName = $normalizedName
        storeType = $StoreType
    }) -Compress
}

function Import-Phase3ExpectedStoreInventory {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$RepositoryRoot,
        [ValidateSet('classic-outlook-ui', 'conditional-vsto-seeder')]
        [string[]]$AllowedSources = @('classic-outlook-ui')
    )

    $allowedSourceSet = @($AllowedSources | Select-Object -Unique)
    if ($allowedSourceSet.Count -lt 1 -or $allowedSourceSet.Count -gt 2) {
        throw 'The expected store inventory must allow one or two recognized sources.'
    }

    if (-not $Path.EndsWith('.local.json', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'The expected store inventory must use the ignored .local.json suffix.'
    }
    $resolvedPath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
    $inventoryFile = Get-Item -LiteralPath $resolvedPath -ErrorAction Stop
    if ($inventoryFile.PSIsContainer -or $inventoryFile.Length -gt 64KB) {
        throw 'The expected store inventory must be a file no larger than 64 KiB.'
    }

    $root = [IO.Path]::GetFullPath($RepositoryRoot).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    $fullPath = [IO.Path]::GetFullPath($resolvedPath)
    if ($fullPath.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
        & git.exe -C $RepositoryRoot check-ignore --quiet -- $fullPath
        if ($LASTEXITCODE -ne 0) {
            $global:LASTEXITCODE = 0
            throw 'An expected store inventory inside the repository must be ignored by Git.'
        }
        $global:LASTEXITCODE = 0
    }

    try {
        $document = Get-Content -Raw -LiteralPath $fullPath -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw 'The expected store inventory was not valid JSON.'
    }
    Assert-Phase3ExactProperties `
        -InputObject $document `
        -Expected @('schema', 'source', 'stores') `
        -Context 'The expected store inventory document'
    $schema = ConvertTo-Phase3Integer `
        -Value $document.schema `
        -Minimum 1 `
        -Maximum 1 `
        -Context 'The expected store inventory schema'
    if ($schema -ne 1 -or $document.source -isnot [string] -or
        $document.source -cnotin $allowedSourceSet) {
        throw 'The expected store inventory must declare schema 1 and an explicitly allowed source.'
    }

    if ($document.stores -isnot [System.Array]) {
        throw 'The expected store inventory stores field must be a JSON array.'
    }
    $stores = @($document.stores)
    if ($stores.Count -gt 64) {
        throw 'The expected store inventory exceeded the supported store limit.'
    }

    $entries = [System.Collections.Generic.List[string]]::new()
    foreach ($store in $stores) {
        Assert-Phase3ExactProperties `
            -InputObject $store `
            -Expected @('displayName', 'storeType') `
            -Context 'An expected store inventory entry'
        Assert-Phase3BoundedText `
            -Value $store.displayName `
            -MaximumLength 256 `
            -Context 'An expected store display name'
        if ($store.storeType -isnot [string] -or $store.storeType -cnotin $script:Phase3StoreTypes) {
            throw 'An expected store type was not supported.'
        }
        $entries.Add((New-Phase3InventoryEntry `
            -DisplayName $store.displayName `
            -StoreType $store.storeType))
    }

    $multiset = Get-Phase3StringMultiset -Entries $entries.ToArray()
    return [pscustomobject]@{
        Count = $multiset.Count
        Entries = $multiset.Entries
        Sha256 = $multiset.Sha256
        Source = $document.source
    }
}

function Assert-Phase3InventoryMatch {
    param(
        [Parameter(Mandatory = $true)][object]$Expected,
        [Parameter(Mandatory = $true)][object]$Actual
    )

    $matches = $Expected.Count -eq $Actual.Count -and
        $Expected.Entries.Count -eq $Actual.Entries.Count -and
        $Expected.Sha256 -ceq $Actual.Sha256
    if ($matches) {
        for ($index = 0; $index -lt $Expected.Entries.Count; $index++) {
            if ($Expected.Entries[$index] -cne $Actual.Entries[$index]) {
                $matches = $false
                break
            }
        }
    }
    if (-not $matches) {
        throw (
            'The probe store inventory did not match the independent Classic Outlook inventory. ' +
            "Expected count/digest: $($Expected.Count)/$($Expected.Sha256); " +
            "actual count/digest: $($Actual.Count)/$($Actual.Sha256).")
    }
}

function Add-Phase3OperationIdsToSharedSet {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$Source,
        [AllowNull()][AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$Destination
    )

    if ($null -eq $Destination) {
        return
    }
    foreach ($operationId in $Source) {
        if (-not $Destination.Add($operationId)) {
            throw 'A Phase 3 operation ID was reused across smoke workloads.'
        }
    }
}

function Assert-Phase3ProbeToolResult {
    param(
        [Parameter(Mandatory = $true)][object]$Result,
        [Parameter(Mandatory = $true)][object]$ExpectedInventory,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$SeenOperationIds,
        [switch]$AllowReadinessPartial,
        [switch]$CodexNormalized
    )

    if ($CodexNormalized) {
        Assert-Phase3ExactProperties `
            -InputObject $Result `
            -Expected @('content', 'structured_content') `
            -Context 'The Codex normalized MCP result'
        $structuredContent = $Result.structured_content
    }
    else {
        Assert-Phase3ExactProperties `
            -InputObject $Result `
            -Expected @('content', 'isError', 'structuredContent') `
            -Context 'The raw MCP tool result'
        if ($Result.isError -isnot [bool] -or $Result.isError -ne $false) {
            throw 'The outlook_probe MCP result reported a tool error.'
        }
        $structuredContent = $Result.structuredContent
    }

    if ($Result.content -isnot [System.Array] -or @($Result.content).Count -ne 1) {
        throw 'The outlook_probe result did not contain exactly one content block.'
    }
    Assert-Phase3ExactProperties `
        -InputObject $Result.content[0] `
        -Expected @('type', 'text') `
        -Context 'The outlook_probe content block'
    if ($Result.content[0].type -cne 'text') {
        throw 'The outlook_probe result contained a non-text content block.'
    }

    Assert-Phase3ExactProperties `
        -InputObject $structuredContent `
        -Expected @('ok', 'operationId', 'data', 'partial', 'warnings') `
        -Context 'The outlook_probe success envelope'
    if ($structuredContent.ok -isnot [bool] -or $structuredContent.ok -ne $true -or
        $structuredContent.partial -isnot [bool] -or $structuredContent.partial -ne $true -or
        $structuredContent.operationId -isnot [string] -or
        $structuredContent.operationId -cnotmatch '^[0-9a-f]{32}$') {
        throw 'The outlook_probe success envelope flags or operation ID were invalid.'
    }
    if (-not $SeenOperationIds.Add($structuredContent.operationId)) {
        throw 'The outlook_probe returned a duplicate operation ID.'
    }
    if ($structuredContent.warnings -isnot [System.Array]) {
        throw 'The outlook_probe warnings field was not an array.'
    }
    $warnings = @($structuredContent.warnings)
    $readinessComplete = $warnings.Count -eq 1 -and
        $warnings[0] -ceq $script:Phase3ArchiveWarning
    $readinessPartial = $AllowReadinessPartial -and
        $warnings.Count -eq 2 -and
        $warnings[0] -ceq $script:Phase3ArchiveWarning -and
        $warnings[1] -ceq $script:Phase3StoreMetadataIncompleteWarning
    if (-not $readinessComplete -and -not $readinessPartial) {
        throw 'The outlook_probe warnings did not match an allowed Phase 3 readiness state.'
    }

    $data = $structuredContent.data
    Assert-Phase3ExactProperties `
        -InputObject $data `
        -Expected @(
            'outlookVersion', 'outlookBitness', 'profileName',
            'dispatcher', 'configuredStoreCount', 'stores'
        ) `
        -Context 'The outlook_probe data object'
    Assert-Phase3BoundedText `
        -Value $data.outlookVersion `
        -MaximumLength 64 `
        -Context 'The Outlook version'
    Assert-Phase3BoundedText `
        -Value $data.profileName `
        -MaximumLength 256 `
        -Context 'The active Outlook profile name'
    $outlookBitness = ConvertTo-Phase3Integer `
        -Value $data.outlookBitness `
        -Minimum 32 `
        -Maximum 64 `
        -Context 'The Outlook bitness'
    if ($outlookBitness -notin @(32, 64)) {
        throw 'The Outlook bitness was neither 32 nor 64.'
    }

    $dispatcher = $data.dispatcher
    Assert-Phase3ExactProperties `
        -InputObject $dispatcher `
        -Expected @(
            'capturedManagedThreadId', 'capturedNativeThreadId',
            'executedManagedThreadId', 'executedNativeThreadId',
            'apartmentState', 'matchesCapturedThread'
        ) `
        -Context 'The outlook_probe dispatcher proof'
    $capturedManaged = ConvertTo-Phase3Integer `
        -Value $dispatcher.capturedManagedThreadId -Minimum 1 -Maximum 2147483647 `
        -Context 'The captured managed thread ID'
    $capturedNative = ConvertTo-Phase3Integer `
        -Value $dispatcher.capturedNativeThreadId -Minimum 1 -Maximum 4294967295 `
        -Context 'The captured native thread ID'
    $executedManaged = ConvertTo-Phase3Integer `
        -Value $dispatcher.executedManagedThreadId -Minimum 1 -Maximum 2147483647 `
        -Context 'The executed managed thread ID'
    $executedNative = ConvertTo-Phase3Integer `
        -Value $dispatcher.executedNativeThreadId -Minimum 1 -Maximum 4294967295 `
        -Context 'The executed native thread ID'
    if ($capturedManaged -ne $executedManaged -or
        $capturedNative -ne $executedNative -or
        $dispatcher.apartmentState -cne 'STA' -or
        $dispatcher.matchesCapturedThread -isnot [bool] -or
        $dispatcher.matchesCapturedThread -ne $true) {
        throw 'The outlook_probe did not execute on the captured Outlook UI STA.'
    }

    $configuredStoreCount = ConvertTo-Phase3Integer `
        -Value $data.configuredStoreCount `
        -Minimum 0 `
        -Maximum 2147483647 `
        -Context 'The configured Outlook store count'
    if ($data.stores -isnot [System.Array]) {
        throw 'The outlook_probe stores field was not a JSON array.'
    }
    $stores = @($data.stores)
    if ($stores.Count -gt 64 -or
        $configuredStoreCount -gt 64 -or
        $configuredStoreCount -lt $stores.Count) {
        throw 'The outlook_probe returned inconsistent bounded store counts.'
    }
    if ($readinessComplete -and $configuredStoreCount -ne $stores.Count) {
        throw 'The ready outlook_probe did not return the complete bounded store inventory.'
    }

    $inventoryEntries = [System.Collections.Generic.List[string]]::new()
    $metadataEntries = [System.Collections.Generic.List[string]]::new()
    foreach ($store in $stores) {
        Assert-Phase3ExactProperties `
            -InputObject $store `
            -Expected @('displayName', 'storeType', 'capabilities', 'standardFolders') `
            -Context 'An outlook_probe store entry'
        Assert-Phase3BoundedText `
            -Value $store.displayName `
            -MaximumLength 256 `
            -Context 'An outlook_probe store display name'
        if ($store.storeType -isnot [string] -or $store.storeType -cnotin $script:Phase3StoreTypes) {
            throw 'An outlook_probe store type was not supported.'
        }

        $capabilities = $store.capabilities
        Assert-Phase3ExactProperties `
            -InputObject $capabilities `
            -Expected @('isExchangeStore', 'isDataFileStore', 'isCachedExchange') `
            -Context 'An outlook_probe capability object'
        foreach ($name in @('isExchangeStore', 'isDataFileStore', 'isCachedExchange')) {
            if ($capabilities.$name -isnot [bool]) {
                throw 'An outlook_probe capability was not Boolean.'
            }
        }

        $folders = $store.standardFolders
        Assert-Phase3ExactProperties `
            -InputObject $folders `
            -Expected @('inbox', 'drafts', 'sentItems', 'deletedItems', 'archive') `
            -Context 'An outlook_probe standard-folder object'
        foreach ($name in @('inbox', 'drafts', 'sentItems', 'deletedItems', 'archive')) {
            if ($folders.$name -isnot [string] -or
                $folders.$name -cnotin $script:Phase3FolderStatuses) {
                throw 'An outlook_probe folder status was not available, missing, or unknown.'
            }
        }
        if ($folders.archive -cne 'unknown') {
            throw 'The archive folder status must remain unknown because Outlook OOM does not expose it.'
        }

        $inventoryEntries.Add((New-Phase3InventoryEntry `
            -DisplayName $store.displayName `
            -StoreType $store.storeType))
        $metadataEntries.Add((ConvertTo-Json -InputObject ([ordered]@{
            displayName = $store.displayName.Normalize([Text.NormalizationForm]::FormC)
            storeType = $store.storeType
            isExchangeStore = $capabilities.isExchangeStore
            isDataFileStore = $capabilities.isDataFileStore
            isCachedExchange = $capabilities.isCachedExchange
            inbox = $folders.inbox
            drafts = $folders.drafts
            sentItems = $folders.sentItems
            deletedItems = $folders.deletedItems
            archive = $folders.archive
        }) -Compress))
    }

    $inventory = Get-Phase3StringMultiset -Entries $inventoryEntries.ToArray()
    $metadata = Get-Phase3StringMultiset -Entries $metadataEntries.ToArray()
    if ($readinessComplete) {
        Assert-Phase3InventoryMatch -Expected $ExpectedInventory -Actual $inventory
    }

    $expectedContentText = if ($stores.Count -eq 1) {
        'Outlook probe verified STA execution and returned 1 store.'
    }
    else {
        "Outlook probe verified STA execution and returned $($stores.Count) stores."
    }
    if ($Result.content[0].text -isnot [string] -or
        $Result.content[0].text -cne $expectedContentText) {
        throw 'The outlook_probe content text was not the bounded count-only summary.'
    }

    $profileSha256 = Get-Phase3Sha256 -Text (
        $data.profileName.Normalize([Text.NormalizationForm]::FormC))
    $environmentSha256 = Get-Phase3Sha256 -Text (ConvertTo-Json -InputObject ([ordered]@{
        outlookVersion = $data.outlookVersion
        outlookBitness = $outlookBitness
        profileSha256 = $profileSha256
        storeMetadataSha256 = $metadata.Sha256
    }) -Compress)

    return [pscustomobject]@{
        OperationId = $structuredContent.operationId
        StoreCount = $stores.Count
        StoreInventorySha256 = $inventory.Sha256
        StoreMetadataSha256 = $metadata.Sha256
        ProfileSha256 = $profileSha256
        EnvironmentSha256 = $environmentSha256
        CapturedManagedThreadId = $capturedManaged
        CapturedNativeThreadId = $capturedNative
        StaVerified = $true
        InventoryMatched = $readinessComplete
        ReadinessComplete = $readinessComplete
    }
}
