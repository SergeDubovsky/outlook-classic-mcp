#Requires -Version 5.1
[CmdletBinding()]
param(
    [Uri]$Endpoint = 'http://127.0.0.1:8765/mcp/',

    [ValidateRange(1, 60)]
    [int]$TimeoutSeconds = 10,

    [ValidateSet(2, 3)]
    [int]$ExpectedPhase = 2,

    [ValidateSet('Full', 'Restart')]
    [string]$Phase3Mode = 'Full',

    [string]$ExpectedStoreInventoryPath,

    [int]$OutlookProcessId,

    [AllowNull()][AllowEmptyCollection()]
    [System.Collections.Generic.HashSet[string]]$SharedPhase3OperationIds
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
if ($ExpectedPhase -eq 3) {
    if ([string]::IsNullOrWhiteSpace($ExpectedStoreInventoryPath)) {
        throw 'Phase 3 requires -ExpectedStoreInventoryPath.'
    }
    if ($OutlookProcessId -le 0) {
        throw 'Phase 3 requires the owned Outlook process ID.'
    }
    . (Join-Path $PSScriptRoot 'phase3-smoke-common.ps1')
}

$tokenVariable = 'OUTLOOK_MCP_TOKEN'
$token = [Environment]::GetEnvironmentVariable($tokenVariable, 'Process')
if ($token -notmatch '^[A-Za-z0-9_-]{43}$') {
    $token = [Environment]::GetEnvironmentVariable($tokenVariable, 'User')
}
if ($token -notmatch '^[A-Za-z0-9_-]{43}$') {
    throw "$tokenVariable is not a canonical 32-byte base64url token in process or current-user scope."
}
if ($Endpoint.AbsoluteUri -ne 'http://127.0.0.1:8765/mcp/') {
    throw 'The production endpoint probe accepts only the canonical loopback endpoint.'
}

Add-Type -AssemblyName System.Net.Http
$handler = [Net.Http.HttpClientHandler]::new()
$handler.UseProxy = $false
$handler.AllowAutoRedirect = $false
$handler.AutomaticDecompression = [Net.DecompressionMethods]::None
$handler.MaxConnectionsPerServer = 4
$client = [Net.Http.HttpClient]::new($handler, $true)
$client.Timeout = [TimeSpan]::FromSeconds($TimeoutSeconds)

function Get-HeaderValues {
    param(
        [Parameter(Mandatory = $true)][Net.Http.HttpResponseMessage]$Response,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $values = [System.Collections.Generic.List[string]]::new()
    foreach ($header in $Response.Headers) {
        if ($header.Key -eq $Name) {
            foreach ($value in $header.Value) {
                $values.Add($value)
            }
        }
    }
    foreach ($header in $Response.Content.Headers) {
        if ($header.Key -eq $Name) {
            foreach ($value in $header.Value) {
                $values.Add($value)
            }
        }
    }
    return $values.ToArray()
}

function Start-McpPost {
    param(
        [Parameter(Mandatory = $true)][string]$Body,
        [bool]$Authenticated = $true,
        [AllowNull()][string]$ProtocolVersion = '2025-11-25',
        [Threading.CancellationToken]$CancellationToken =
            [Threading.CancellationToken]::None
    )

    $request = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::Post, $Endpoint)
    try {
        $request.Content = [Net.Http.StringContent]::new($Body, [Text.Encoding]::UTF8, 'application/json')
        $request.Headers.Host = '127.0.0.1:8765'
        if (-not $request.Headers.TryAddWithoutValidation('Accept', 'application/json, text/event-stream')) {
            throw 'Could not set the required Accept header.'
        }
        if ($Authenticated -and
            -not $request.Headers.TryAddWithoutValidation('Authorization', "Bearer $token")) {
            throw 'Could not set the Authorization header.'
        }
        if (-not [string]::IsNullOrEmpty($ProtocolVersion) -and
            -not $request.Headers.TryAddWithoutValidation('MCP-Protocol-Version', $ProtocolVersion)) {
            throw 'Could not set the MCP protocol-version header.'
        }

        return [pscustomobject]@{
            Request = $request
            Task = $client.SendAsync($request, $CancellationToken)
        }
    }
    catch {
        $request.Dispose()
        throw
    }
}

function Complete-McpPost {
    param([Parameter(Mandatory = $true)][object]$Pending)

    $response = $null
    try {
        $response = $Pending.Task.GetAwaiter().GetResult()
        $contentType = if ($null -eq $response.Content.Headers.ContentType) {
            $null
        }
        else {
            $response.Content.Headers.ContentType.MediaType
        }
        return [pscustomobject]@{
            StatusCode = [int]$response.StatusCode
            Body = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            ContentType = $contentType
            CacheControl = @(Get-HeaderValues -Response $response -Name 'Cache-Control')
            ContentEncoding = @(Get-HeaderValues -Response $response -Name 'Content-Encoding')
            WwwAuthenticate = @(Get-HeaderValues -Response $response -Name 'WWW-Authenticate')
        }
    }
    finally {
        if ($null -ne $response) {
            $response.Dispose()
        }
        $Pending.Request.Dispose()
    }
}

function Invoke-McpPost {
    param(
        [Parameter(Mandatory = $true)][string]$Body,
        [bool]$Authenticated = $true,
        [AllowNull()][string]$ProtocolVersion = '2025-11-25',
        [Threading.CancellationToken]$CancellationToken =
            [Threading.CancellationToken]::None
    )

    $pending = Start-McpPost `
        -Body $Body `
        -Authenticated $Authenticated `
        -ProtocolVersion $ProtocolVersion `
        -CancellationToken $CancellationToken
    return Complete-McpPost -Pending $pending
}

function ConvertFrom-SseResponse {
    param([Parameter(Mandatory = $true)]$Response)

    if ($Response.StatusCode -ne 200 -or $Response.ContentType -ne 'text/event-stream') {
        throw "Expected an SSE response with HTTP 200; received $($Response.StatusCode) and '$($Response.ContentType)'."
    }
    $cacheDirectives = @(
        $Response.CacheControl |
            ForEach-Object { $_ -split ',' } |
            ForEach-Object { $_.Trim() }
    )
    if (@($cacheDirectives | Where-Object { $_ -eq 'no-cache' }).Count -ne 1 -or
        @($cacheDirectives | Where-Object { $_ -eq 'no-store' }).Count -ne 1 -or
        $cacheDirectives.Count -ne 2) {
        throw 'The SSE response does not have exactly the no-cache and no-store directives.'
    }
    if ($Response.ContentEncoding.Count -ne 1 -or $Response.ContentEncoding[0] -ne 'identity') {
        throw 'The SSE response does not use identity content encoding.'
    }

    $dataLines = @(
        $Response.Body.Replace("`r`n", "`n").Split("`n") |
            Where-Object { $_.StartsWith('data: ', [StringComparison]::Ordinal) }
    )
    if ($dataLines.Count -ne 1) {
        throw "Expected one SSE data line; received $($dataLines.Count)."
    }
    return $dataLines[0].Substring(6) | ConvertFrom-Json
}

function Assert-OutlookResponding {
    if ($ExpectedPhase -ne 3) {
        return
    }

    $process = Get-Process -Id $OutlookProcessId -ErrorAction Stop
    try {
        if ($process.ProcessName -cne 'OUTLOOK') {
            throw 'The supplied Phase 3 process is not Outlook.'
        }
        $process.Refresh()
        if ($process.HasExited -or
            $process.MainWindowHandle -eq [IntPtr]::Zero -or
            -not $process.Responding) {
            throw 'Outlook was not responding during the Phase 3 concurrent probe gate.'
        }
    }
    finally {
        $process.Dispose()
    }
}

function Assert-Phase3StatusResult {
    param(
        [Parameter(Mandatory = $true)][object]$Result,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$SeenOperationIds
    )

    Assert-Phase3ExactProperties `
        -InputObject $Result `
        -Expected @('content', 'isError', 'structuredContent') `
        -Context 'The outlook_status MCP result'
    if ($Result.isError -isnot [bool] -or $Result.isError -ne $false) {
        throw 'The outlook_status result reported a tool error.'
    }
    $structured = $Result.structuredContent
    Assert-Phase3ExactProperties `
        -InputObject $structured `
        -Expected @('ok', 'operationId', 'data', 'partial', 'warnings') `
        -Context 'The outlook_status success envelope'
    if ($structured.ok -isnot [bool] -or $structured.ok -ne $true -or
        $structured.partial -isnot [bool] -or $structured.partial -ne $false -or
        $structured.operationId -isnot [string] -or
        $structured.operationId -cnotmatch '^[0-9a-f]{32}$' -or
        -not $SeenOperationIds.Add($structured.operationId) -or
        $structured.warnings -isnot [System.Array] -or
        @($structured.warnings).Count -ne 0) {
        throw 'The outlook_status structured success envelope was invalid.'
    }
    Assert-Phase3ExactProperties `
        -InputObject $structured.data `
        -Expected @('hostState', 'listenerReady', 'version') `
        -Context 'The outlook_status data object'
    if ($structured.data.hostState -cne 'online' -or
        $structured.data.listenerReady -isnot [bool] -or
        $structured.data.listenerReady -ne $true -or
        $structured.data.version -isnot [string] -or
        [string]::IsNullOrWhiteSpace($structured.data.version) -or
        $structured.data.version.Length -gt 64 -or
        $structured.data.version -match '[\x00-\x1F\x7F]') {
        throw 'The outlook_status result did not report an online ready listener.'
    }
}

function Assert-Phase3ProbeStability {
    param([Parameter(Mandatory = $true)][object[]]$Proofs)

    if ($Proofs.Count -eq 0) {
        throw 'No Phase 3 probe results were available for stability validation.'
    }
    $baseline = $Proofs[0]
    foreach ($proof in $Proofs) {
        if ($proof.StoreCount -ne $baseline.StoreCount -or
            $proof.StoreInventorySha256 -cne $baseline.StoreInventorySha256 -or
            $proof.StoreMetadataSha256 -cne $baseline.StoreMetadataSha256 -or
            $proof.ProfileSha256 -cne $baseline.ProfileSha256 -or
            $proof.EnvironmentSha256 -cne $baseline.EnvironmentSha256 -or
            $proof.CapturedManagedThreadId -ne $baseline.CapturedManagedThreadId -or
            $proof.CapturedNativeThreadId -ne $baseline.CapturedNativeThreadId -or
            -not $proof.StaVerified -or
            -not $proof.InventoryMatched) {
            throw 'Phase 3 probe metadata or captured Outlook STA identity changed within the cycle.'
        }
    }
}

function Invoke-Phase3ReadinessProbe {
    param(
        [Parameter(Mandatory = $true)][ValidateRange(1, 2147483643)]
        [int]$FirstRequestId,
        [Parameter(Mandatory = $true)][object]$ExpectedInventory,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$SeenOperationIds
    )

    $maximumAttempts = 5
    $retryDelayMilliseconds = 500
    $deadlineMilliseconds = 30000
    $timer = [Diagnostics.Stopwatch]::StartNew()

    for ($attempt = 1; $attempt -le $maximumAttempts; $attempt++) {
        if ($attempt -gt 1) {
            if ($timer.ElapsedMilliseconds -ge $deadlineMilliseconds) {
                throw 'The first outlook_probe did not become ready within the bounded readiness deadline.'
            }
            Start-Sleep -Milliseconds $retryDelayMilliseconds
            if ($timer.ElapsedMilliseconds -ge $deadlineMilliseconds) {
                throw 'The first outlook_probe did not become ready within the bounded readiness deadline.'
            }
        }

        $requestId = $FirstRequestId + $attempt - 1
        $body = '{"jsonrpc":"2.0","id":' + $requestId +
            ',"method":"tools/call","params":{"name":"outlook_probe","arguments":{}}}'
        $remainingMilliseconds = $deadlineMilliseconds - $timer.ElapsedMilliseconds
        if ($remainingMilliseconds -le 0) {
            throw 'The first outlook_probe did not become ready within the bounded readiness deadline.'
        }

        $deadlineCancellation = [Threading.CancellationTokenSource]::new()
        $deadlineCancellation.CancelAfter([int]$remainingMilliseconds)
        try {
            try {
                $probe = ConvertFrom-SseResponse (Invoke-McpPost `
                    -Body $body `
                    -CancellationToken $deadlineCancellation.Token)
            }
            catch {
                if ($deadlineCancellation.IsCancellationRequested) {
                    throw 'The first outlook_probe did not become ready within the bounded readiness deadline.'
                }
                throw
            }
        }
        finally {
            $deadlineCancellation.Dispose()
        }

        if ($probe.id -ne $requestId) {
            throw 'A readiness outlook_probe response did not preserve its request ID.'
        }
        $proof = Assert-Phase3ProbeToolResult `
            -Result $probe.result `
            -ExpectedInventory $ExpectedInventory `
            -SeenOperationIds $SeenOperationIds `
            -AllowReadinessPartial
        if ($timer.ElapsedMilliseconds -ge $deadlineMilliseconds) {
            throw 'The first outlook_probe did not become ready within the bounded readiness deadline.'
        }
        if ($proof.ReadinessComplete) {
            return [pscustomobject]@{
                Proof = $proof
                RetryCount = $attempt - 1
                AttemptCount = $attempt
            }
        }
        if ($attempt -eq $maximumAttempts) {
            throw 'The first outlook_probe did not become ready within the bounded attempt limit.'
        }
        if ($timer.ElapsedMilliseconds -ge $deadlineMilliseconds) {
            throw 'The first outlook_probe did not become ready within the bounded readiness deadline.'
        }
    }
}

try {
    if ($ExpectedPhase -eq 3) {
        $expectedInventory = Import-Phase3ExpectedStoreInventory `
            -Path $ExpectedStoreInventoryPath `
            -RepositoryRoot $repositoryRoot
        $seenOperationIds = [System.Collections.Generic.HashSet[string]]::new(
            [StringComparer]::Ordinal)
        $probeProofs = [System.Collections.Generic.List[object]]::new()

        if ($Phase3Mode -eq 'Restart') {
            Assert-OutlookResponding
            $readiness = Invoke-Phase3ReadinessProbe `
                -FirstRequestId 300 `
                -ExpectedInventory $expectedInventory `
                -SeenOperationIds $seenOperationIds
            $probeProofs.Add($readiness.Proof)
            Assert-OutlookResponding
            Assert-Phase3ProbeStability -Proofs $probeProofs.ToArray()
            Add-Phase3OperationIdsToSharedSet `
                -Source $seenOperationIds `
                -Destination $SharedPhase3OperationIds
            $operationSet = Get-Phase3StringMultiset -Entries ([string[]]@($seenOperationIds))
            $baseline = $probeProofs[0]

            [pscustomobject]@{
                Phase = 3
                FullWorkload = $false
                AuthenticationVerified = $false
                InitializeVerified = $false
                InitializedNotificationVerified = $false
                PingVerified = $false
                ToolCount = 0
                StatusCallCount = 0
                SequentialProbeCount = 0
                ConcurrentProbeCount = 0
                RestartProbeCount = 1
                ReadinessRetryCount = $readiness.RetryCount
                ProbeResultCount = $readiness.AttemptCount
                StructuredEnvelopeCount = $seenOperationIds.Count
                UniqueOperationIdCount = $seenOperationIds.Count
                UniqueOperationIds =
                    $seenOperationIds.Count -eq (1 + $readiness.RetryCount)
                OperationIdSetSha256 = $operationSet.Sha256
                StoreCount = $baseline.StoreCount
                StoreInventoryMatched = $true
                StoreInventorySha256 = $baseline.StoreInventorySha256
                ExpectedStoreInventorySha256 = $expectedInventory.Sha256
                StoreMetadataStable = $true
                StoreMetadataSha256 = $baseline.StoreMetadataSha256
                EnvironmentSha256 = $baseline.EnvironmentSha256
                DispatcherIdentityStable = $true
                AllProbeCallsOnCapturedSta = $true
                OutlookResponsiveCheckCount = 2
                OutlookResponsive = $true
            }
            return
        }

        $readinessRetryCount = 0

        $unauthorized = Invoke-McpPost `
            -Body '{"jsonrpc":"2.0","id":1,"method":"ping","params":{}}' `
            -Authenticated $false
        if ($unauthorized.StatusCode -ne 401 -or
            $unauthorized.WwwAuthenticate.Count -ne 1 -or
            $unauthorized.WwwAuthenticate[0] -ne 'Bearer' -or
            $unauthorized.Body.Contains('outlook_status') -or
            $unauthorized.Body.Contains('outlook_probe')) {
            throw 'The unauthenticated request did not receive the fail-closed Phase 3 response.'
        }

        $initialize = ConvertFrom-SseResponse (Invoke-McpPost `
            -Body '{"jsonrpc":"2.0","id":2,"method":"initialize","params":{"protocolVersion":"2025-11-25","capabilities":{},"clientInfo":{"name":"phase3-smoke","version":"1.0"}}}' `
            -ProtocolVersion $null)
        if ($initialize.id -ne 2 -or
            $initialize.result.protocolVersion -ne '2025-11-25' -or
            $initialize.result.serverInfo.name -ne 'outlook-classic-mcp') {
            throw 'The initialize response did not match the Phase 3 server contract.'
        }

        $initialized = Invoke-McpPost `
            -Body '{"jsonrpc":"2.0","method":"notifications/initialized","params":{}}'
        if ($initialized.StatusCode -ne 202 -or
            $initialized.Body.Length -ne 0 -or
            $null -ne $initialized.ContentType -or
            $initialized.ContentEncoding.Count -ne 0) {
            throw 'The initialized notification did not receive the empty HTTP 202 contract.'
        }

        $ping = ConvertFrom-SseResponse (Invoke-McpPost `
            -Body '{"jsonrpc":"2.0","id":3,"method":"ping","params":{}}')
        if ($ping.id -ne 3 -or $null -eq $ping.result) {
            throw 'The ping response did not preserve its request ID.'
        }

        $tools = ConvertFrom-SseResponse (Invoke-McpPost `
            -Body '{"jsonrpc":"2.0","id":4,"method":"tools/list","params":{}}')
        $listedTools = @($tools.result.tools)
        if ($tools.id -ne 4 -or
            $listedTools.Count -ne 2 -or
            $listedTools[0].name -cne 'outlook_status' -or
            $listedTools[1].name -cne 'outlook_probe') {
            throw 'The Phase 3 endpoint did not expose exactly outlook_status followed by outlook_probe.'
        }
        Assert-Phase3ExactProperties `
            -InputObject $listedTools[1].inputSchema `
            -Expected @('type', 'properties', 'additionalProperties') `
            -Context 'The outlook_probe input schema'
        if ($listedTools[1].inputSchema.type -cne 'object' -or
            @($listedTools[1].inputSchema.properties.PSObject.Properties).Count -ne 0 -or
            $listedTools[1].inputSchema.additionalProperties -isnot [bool] -or
            $listedTools[1].inputSchema.additionalProperties -ne $false) {
            throw 'The outlook_probe input schema was not a closed empty object.'
        }

        $status = ConvertFrom-SseResponse (Invoke-McpPost `
            -Body '{"jsonrpc":"2.0","id":5,"method":"tools/call","params":{"name":"outlook_status","arguments":{}}}')
        if ($status.id -ne 5) {
            throw 'The outlook_status response did not preserve its request ID.'
        }
        Assert-Phase3StatusResult -Result $status.result -SeenOperationIds $seenOperationIds

        $readiness = Invoke-Phase3ReadinessProbe `
            -FirstRequestId 100 `
            -ExpectedInventory $expectedInventory `
            -SeenOperationIds $seenOperationIds
        $readinessRetryCount = $readiness.RetryCount
        $probeProofs.Add($readiness.Proof)

        for ($index = 1; $index -lt 20; $index++) {
            $requestId = 104 + $index
            $body = '{"jsonrpc":"2.0","id":' + $requestId +
                ',"method":"tools/call","params":{"name":"outlook_probe","arguments":{}}}'
            $probe = ConvertFrom-SseResponse (Invoke-McpPost -Body $body)
            if ($probe.id -ne $requestId) {
                throw 'A sequential outlook_probe response did not preserve its request ID.'
            }
            $probeProofs.Add((Assert-Phase3ProbeToolResult `
                -Result $probe.result `
                -ExpectedInventory $expectedInventory `
                -SeenOperationIds $seenOperationIds))
        }

        Assert-OutlookResponding
        $pendingProbes = [System.Collections.Generic.List[object]]::new()
        try {
            for ($index = 0; $index -lt 4; $index++) {
                $requestId = 200 + $index
                $body = '{"jsonrpc":"2.0","id":' + $requestId +
                    ',"method":"tools/call","params":{"name":"outlook_probe","arguments":{}}}'
                $pendingProbes.Add((Start-McpPost -Body $body))
            }
            Assert-OutlookResponding
            for ($index = 0; $index -lt $pendingProbes.Count; $index++) {
                $requestId = 200 + $index
                $probe = ConvertFrom-SseResponse (Complete-McpPost -Pending $pendingProbes[$index])
                if ($probe.id -ne $requestId) {
                    throw 'A concurrent outlook_probe response did not preserve its request ID.'
                }
                $probeProofs.Add((Assert-Phase3ProbeToolResult `
                    -Result $probe.result `
                    -ExpectedInventory $expectedInventory `
                    -SeenOperationIds $seenOperationIds))
            }
        }
        finally {
            foreach ($pendingProbe in $pendingProbes) {
                $pendingProbe.Request.Dispose()
            }
        }
        Assert-OutlookResponding

        if ($probeProofs.Count -ne 24 -or
            $seenOperationIds.Count -ne (25 + $readinessRetryCount)) {
            throw 'The Phase 3 first-cycle endpoint workload returned an unexpected result count.'
        }
        Assert-Phase3ProbeStability -Proofs $probeProofs.ToArray()
        Add-Phase3OperationIdsToSharedSet `
            -Source $seenOperationIds `
            -Destination $SharedPhase3OperationIds
        $operationSet = Get-Phase3StringMultiset -Entries ([string[]]@($seenOperationIds))
        $baseline = $probeProofs[0]

        [pscustomobject]@{
            Phase = 3
            FullWorkload = $true
            AuthenticationVerified = $true
            InitializeVerified = $true
            InitializedNotificationVerified = $true
            PingVerified = $true
            ToolCount = 2
            StatusCallCount = 1
            SequentialProbeCount = 20
            ConcurrentProbeCount = 4
            RestartProbeCount = 0
            ReadinessRetryCount = $readinessRetryCount
            ProbeResultCount = $probeProofs.Count + $readinessRetryCount
            StructuredEnvelopeCount = $seenOperationIds.Count
            UniqueOperationIdCount = $seenOperationIds.Count
            UniqueOperationIds =
                $seenOperationIds.Count -eq (25 + $readinessRetryCount)
            OperationIdSetSha256 = $operationSet.Sha256
            StoreCount = $baseline.StoreCount
            StoreInventoryMatched = $true
            StoreInventorySha256 = $baseline.StoreInventorySha256
            ExpectedStoreInventorySha256 = $expectedInventory.Sha256
            StoreMetadataStable = $true
            StoreMetadataSha256 = $baseline.StoreMetadataSha256
            EnvironmentSha256 = $baseline.EnvironmentSha256
            DispatcherIdentityStable = $true
            AllProbeCallsOnCapturedSta = $true
            OutlookResponsiveCheckCount = 3
            OutlookResponsive = $true
        }
        return
    }

    $unauthorized = Invoke-McpPost `
        -Body '{"jsonrpc":"2.0","id":1,"method":"ping","params":{}}' `
        -Authenticated $false
    if ($unauthorized.StatusCode -ne 401 -or
        $unauthorized.WwwAuthenticate.Count -ne 1 -or
        $unauthorized.WwwAuthenticate[0] -ne 'Bearer' -or
        $unauthorized.Body.Contains('outlook_status')) {
        throw 'The unauthenticated request did not receive the fail-closed Phase 2 response.'
    }

    $initialize = ConvertFrom-SseResponse (Invoke-McpPost `
        -Body '{"jsonrpc":"2.0","id":2,"method":"initialize","params":{"protocolVersion":"2025-11-25","capabilities":{},"clientInfo":{"name":"phase2-smoke","version":"1.0"}}}' `
        -ProtocolVersion $null)
    if ($initialize.id -ne 2 -or
        $initialize.result.protocolVersion -ne '2025-11-25' -or
        $initialize.result.serverInfo.name -ne 'outlook-classic-mcp') {
        throw 'The initialize response did not match the Phase 2 server contract.'
    }

    $initialized = Invoke-McpPost `
        -Body '{"jsonrpc":"2.0","method":"notifications/initialized","params":{}}'
    if ($initialized.StatusCode -ne 202 -or
        $initialized.Body.Length -ne 0 -or
        $null -ne $initialized.ContentType -or
        $initialized.ContentEncoding.Count -ne 0) {
        throw 'The initialized notification did not receive the empty HTTP 202 contract.'
    }

    $ping = ConvertFrom-SseResponse (Invoke-McpPost `
        -Body '{"jsonrpc":"2.0","id":3,"method":"ping","params":{}}')
    if ($ping.id -ne 3 -or $null -eq $ping.result) {
        throw 'The ping response did not preserve its request ID.'
    }

    $tools = ConvertFrom-SseResponse (Invoke-McpPost `
        -Body '{"jsonrpc":"2.0","id":4,"method":"tools/list","params":{}}')
    if ($tools.result.tools.Count -ne 1 -or $tools.result.tools[0].name -ne 'outlook_status') {
        throw 'The Phase 2 endpoint did not expose exactly outlook_status.'
    }

    $status = ConvertFrom-SseResponse (Invoke-McpPost `
        -Body '{"jsonrpc":"2.0","id":5,"method":"tools/call","params":{"name":"outlook_status","arguments":{}}}')
    if ($status.id -ne 5 -or
        $status.result.isError -ne $false -or
        $status.result.structuredContent.ok -ne $true -or
        $status.result.structuredContent.data.hostState -ne 'online' -or
        $status.result.structuredContent.data.listenerReady -ne $true) {
        throw 'The outlook_status result did not report an online ready listener.'
    }

    [pscustomobject]@{
        Endpoint = $Endpoint.AbsoluteUri
        AuthenticationEnforced = $true
        Initialize = $true
        InitializedNotification = $true
        Ping = $true
        Tools = @($tools.result.tools | ForEach-Object { $_.name })
        StatusOnline = $true
        OnlyStatusToolExposed = $true
    }
}
finally {
    $client.Dispose()
}
