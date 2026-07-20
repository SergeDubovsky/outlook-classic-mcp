#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [DateTimeOffset]$SinceUtc,

    [Parameter(Mandatory = $true)]
    [DateTimeOffset]$UntilUtc,

    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [object[]]$EndpointResults,

    [Parameter(Mandatory = $true)]
    [object]$CodexResult,

    [Parameter(Mandatory = $true)]
    [System.Collections.Generic.HashSet[string]]$SensitiveValues,

    [Parameter(Mandatory = $true)]
    [System.Collections.Generic.HashSet[string]]$OperationIds,

    [ValidateRange(60, 1000)]
    [int]$ExpectedRepeatedReadCount = 200,

    [ValidateRange(1, 10)]
    [int]$ExpectedCycles = 3,

    [string]$LogDirectory = (Join-Path $env:LOCALAPPDATA 'OutlookClassicMcp\logs')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($ExpectedCycles -ne 3) {
    throw 'The Phase 4 gate requires exactly three Outlook lifecycle cycles.'
}
if (($ExpectedRepeatedReadCount % 20) -ne 0) {
    throw '-ExpectedRepeatedReadCount must be divisible by twenty.'
}

$lifecycle = & (Join-Path $PSScriptRoot 'verify-phase3-smoke.ps1') `
    -SinceUtc $SinceUtc `
    -UntilUtc $UntilUtc `
    -ExpectedCycles $ExpectedCycles `
    -LogDirectory $LogDirectory

$expectedEndpointFields = @(
    'Phase', 'CycleNumber', 'FixtureSource', 'Mode', 'FullWorkload', 'ToolCount', 'StoreCount',
    'StoreInventoryMatched', 'StoreInventorySha256', 'InboxStoreCount',
    'KnownMessageReacquireCount', 'StaticPageCount', 'StaticPageItemCount',
    'CursorErrorCount', 'LargeFolderMinimumItemCount', 'PartialSearchVerified',
    'ConversationMessageCount', 'AttachmentMetadataCount',
    'BodyTruncationVerified', 'ProtectedBodyVerified',
    'ProtectedFixtureConfigured', 'CancellationRecoveryVerified',
    'ConcurrentReadCount', 'LiveStructuredTimeoutCount', 'RepeatedReadCount',
    'ResourceSampleCount', 'ResourceStable', 'ComAcquiredDelta',
    'ComReleasedDelta', 'ComOutstanding', 'MaterializedItemHighWater',
    'MaximumHandleCount', 'MaximumPrivateBytes', 'MaximumGdiObjects',
    'MaximumUserObjects', 'OutlookResponsiveCheckCount',
    'MaximumUiResponseMilliseconds', 'AllProbeCallsOnCapturedSta',
    'UniqueOperationIdCount', 'UniqueOperationIds'
)

function Assert-ExactFields {
    param(
        [Parameter(Mandatory = $true)][object]$Value,
        [Parameter(Mandatory = $true)][string[]]$Expected,
        [Parameter(Mandatory = $true)][string]$Context
    )

    $actual = @($Value.PSObject.Properties.Name | Sort-Object)
    $wanted = @($Expected | Sort-Object)
    if (($actual -join "`n") -cne ($wanted -join "`n")) {
        throw "$Context does not contain the exact expected evidence fields."
    }
}

function Assert-True {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

$results = @($EndpointResults | Sort-Object CycleNumber)
Assert-True ($results.Count -eq 3) 'Phase 4 did not produce exactly three endpoint results.'
$expectedModes = @('Full', 'Restart', 'Restart')
$inventoryHash = $null
$fixtureSource = $null
$endpointOperationCount = 0
foreach ($index in 0..2) {
    $result = $results[$index]
    Assert-ExactFields `
        -Value $result `
        -Expected $expectedEndpointFields `
        -Context "Phase 4 endpoint cycle $($index + 1)"
    Assert-True ($result.Phase -eq 4) 'An endpoint result has the wrong phase.'
    Assert-True ($result.CycleNumber -eq ($index + 1)) 'Endpoint cycle numbering is incomplete.'
    Assert-True ($result.FixtureSource -cin @('classic-outlook-ui', 'conditional-vsto-seeder')) 'The endpoint result has an unsupported fixture source.'
    Assert-True ($result.Mode -ceq $expectedModes[$index]) 'Endpoint cycle modes are not Full, Restart, Restart.'
    Assert-True ($result.FullWorkload -eq ($index -eq 0)) 'The full read workload did not run exactly once.'
    Assert-True ($result.ToolCount -eq 9) 'The endpoint did not expose exactly nine Phase 4 tools.'
    Assert-True ($result.StoreCount -ge 2) 'Fewer than two Outlook stores were discovered.'
    Assert-True ($result.StoreInventoryMatched -eq $true) 'The endpoint store inventory did not match the independent fixture.'
    if ($result.FixtureSource -ceq 'conditional-vsto-seeder') {
        Assert-True ($result.InboxStoreCount -ge 1) 'The conditional fixture did not verify its expected standard Inbox.'
    }
    else {
        Assert-True ($result.InboxStoreCount -ge 2) 'Fewer than two classic fixture stores with an Inbox were verified.'
    }
    Assert-True ($result.KnownMessageReacquireCount -ge 2) 'Known messages were not reacquired in at least two stores.'
    Assert-True ($result.AllProbeCallsOnCapturedSta -eq $true) 'A Phase 4 cycle did not prove Outlook UI-STA execution.'
    Assert-True ($result.UniqueOperationIds -eq $true -and $result.UniqueOperationIdCount -gt 0) 'A Phase 4 cycle has incomplete operation-ID evidence.'
    Assert-True ($result.OutlookResponsiveCheckCount -ge 2) 'A Phase 4 cycle has insufficient Outlook responsiveness checks.'
    Assert-True ($result.MaximumUiResponseMilliseconds -ge 0 -and $result.MaximumUiResponseMilliseconds -lt 2000) 'Outlook exceeded the Phase 4 UI response deadline.'
    $endpointOperationCount += [int]$result.UniqueOperationIdCount

    if ($null -eq $inventoryHash) {
        $inventoryHash = $result.StoreInventorySha256
        $fixtureSource = $result.FixtureSource
        Assert-True ($inventoryHash -is [string] -and $inventoryHash -cmatch '^[0-9a-f]{64}$') 'The store inventory fingerprint is invalid.'
    }
    else {
        Assert-True ($result.StoreInventorySha256 -ceq $inventoryHash) 'The independent store inventory changed between restart cycles.'
        Assert-True ($result.FixtureSource -ceq $fixtureSource) 'The fixture source changed between restart cycles.'
    }
}

$full = $results[0]
Assert-True ($full.StaticPageCount -eq 3 -and $full.StaticPageItemCount -eq 12) 'The static pagination fixture did not produce three complete pages.'
Assert-True ($full.CursorErrorCount -eq 2) 'The deterministic cursor error probes were incomplete.'
Assert-True ($full.LargeFolderMinimumItemCount -ge 1001) 'The independently seeded large-folder evidence is below 1,001 items.'
Assert-True ($full.PartialSearchVerified -eq $true) 'Cross-store partial success was not verified.'
Assert-True ($full.ConversationMessageCount -ge 1) 'Conversation retrieval did not return the seeded message or conversation.'
Assert-True ($full.AttachmentMetadataCount -ge 2) 'Attachment metadata was not verified for two inert attachments.'
Assert-True ($full.BodyTruncationVerified -eq $true) 'Message-body truncation metadata was not verified.'
if ($full.ProtectedFixtureConfigured) {
    Assert-True ($full.ProtectedBodyVerified -eq $true) 'The configured protected-item fixture was not verified.'
}
Assert-True ($full.CancellationRecoveryVerified -eq $true) 'Client cancellation and endpoint recovery were not verified.'
Assert-True ($full.ConcurrentReadCount -eq 4) 'The bounded concurrent-read workload was incomplete.'
Assert-True ($full.LiveStructuredTimeoutCount -ge 0 -and $full.LiveStructuredTimeoutCount -le 4) 'The bounded concurrent workload returned an invalid structured TIMEOUT count.'
Assert-True ($full.RepeatedReadCount -eq $ExpectedRepeatedReadCount) 'The repeated-read workload count is incomplete.'
Assert-True ($full.ResourceSampleCount -ge 6 -and $full.ResourceStable -eq $true) 'The repeated-read resource evidence is incomplete or unstable.'
Assert-True ($full.ComAcquiredDelta -gt 0 -and $full.ComReleasedDelta -eq $full.ComAcquiredDelta -and $full.ComOutstanding -eq 0) 'Direct Outlook COM telemetry was not balanced after repeated reads.'
Assert-True ($full.MaterializedItemHighWater -ge 1 -and $full.MaterializedItemHighWater -le 51) 'The large-folder workload exceeded the page-size-plus-one materialization bound.'
Assert-True ($full.MaximumHandleCount -gt 0 -and $full.MaximumPrivateBytes -gt 0 -and $full.MaximumGdiObjects -gt 0 -and $full.MaximumUserObjects -gt 0) 'Windows process resource evidence is incomplete.'

$expectedCodexFields = @(
    'Phase', 'Server', 'Tool', 'CompletionMarker', 'MailboxCount',
    'McpToolCallCount', 'StructuredEnvelopeCount', 'UniqueOperationIdCount',
    'UniqueOperationIds', 'ConfigurationValidated', 'ResultValidated',
    'ProcessTreeClosed'
)
Assert-ExactFields -Value $CodexResult -Expected $expectedCodexFields -Context 'native Codex result'
Assert-True ($CodexResult.Phase -eq 4) 'The native Codex result has the wrong phase.'
Assert-True ($CodexResult.Server -ceq 'outlook_classic' -and $CodexResult.Tool -ceq 'outlook_list_mailboxes') 'Codex used an unexpected MCP server or tool.'
Assert-True ($CodexResult.MailboxCount -eq $full.StoreCount -and $CodexResult.CompletionMarker -ceq 'MCP_PHASE4_CODEX_OK') 'The native Codex mailbox count did not match raw endpoint evidence.'
Assert-True ($CodexResult.McpToolCallCount -eq 1 -and $CodexResult.StructuredEnvelopeCount -eq 1 -and $CodexResult.UniqueOperationIdCount -eq 1) 'Codex did not perform exactly one native structured MCP call.'
Assert-True ($CodexResult.UniqueOperationIds -eq $true -and $CodexResult.ConfigurationValidated -eq $true -and $CodexResult.ResultValidated -eq $true -and $CodexResult.ProcessTreeClosed -eq $true) 'Native Codex validation or process cleanup is incomplete.'
Assert-True ($OperationIds.Count -eq ($endpointOperationCount + 1)) 'The shared operation-ID set does not exactly match endpoint plus native Codex evidence.'
Assert-True ($SensitiveValues.Count -ge 10) 'The privacy sentinel set is unexpectedly incomplete.'

$maximumLogBytes = 5L * 1024L * 1024L
$logFiles = @(Get-ChildItem -LiteralPath $LogDirectory -Filter 'outlook-classic-mcp*.log' -File -ErrorAction Stop)
Assert-True ($logFiles.Count -ge 1 -and $logFiles.Count -le 5) 'The privacy verifier did not find the bounded diagnostics log set.'
foreach ($logFile in $logFiles) {
    Assert-True ($logFile.Length -le $maximumLogBytes) 'A diagnostics log exceeds the five-megabyte privacy scan limit.'
    $content = [IO.File]::ReadAllText($logFile.FullName, [Text.Encoding]::UTF8)
    foreach ($sensitiveValue in $SensitiveValues) {
        if ($content.IndexOf($sensitiveValue, [StringComparison]::Ordinal) -ge 0) {
            throw 'A bearer token, mailbox-derived value, message marker, or opaque Outlook locator entered diagnostics.'
        }
    }
}

[pscustomobject]@{
    Phase = 4
    VerifiedCycleCount = $results.Count
    DistinctDiagnosticSessionCount = $lifecycle.VerifiedCycleCount
    EndpointVerificationCount = $results.Count
    FullEndpointWorkloadCount = 1
    ToolCount = 9
    StoreCount = $full.StoreCount
    StoreInventoryMatched = $true
    KnownMessageStoreCount = $full.KnownMessageReacquireCount
    StaticPageCount = $full.StaticPageCount
    StaticPageItemCount = $full.StaticPageItemCount
    LargeFolderMinimumItemCount = $full.LargeFolderMinimumItemCount
    PartialSearchVerified = $true
    CancellationRecoveryVerified = $true
    StructuredTimeoutObserved = $full.LiveStructuredTimeoutCount -gt 0
    RepeatedReadCount = $full.RepeatedReadCount
    DirectComTelemetryBalanced = $true
    MaterializedItemHighWater = $full.MaterializedItemHighWater
    ResourceStable = $true
    OutlookResponsive = $true
    ProtectedFixtureVerified = $full.ProtectedFixtureConfigured -and $full.ProtectedBodyVerified
    NativeCodexVerified = $true
    UniqueOperationIdCount = $OperationIds.Count
    PrivacySentinelCount = $SensitiveValues.Count
    DiagnosticsLogFileCount = $logFiles.Count
    OutlookStopped = $true
    PortReleased = $true
}
