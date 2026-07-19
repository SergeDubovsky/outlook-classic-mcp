#Requires -Version 5.1
[CmdletBinding()]
param(
    [Uri]$Endpoint = 'http://127.0.0.1:8765/mcp/',

    [ValidateRange(1, 60)]
    [int]$TimeoutSeconds = 10
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$tokenVariable = 'OUTLOOK_MCP_TOKEN'
$token = [Environment]::GetEnvironmentVariable($tokenVariable, 'Process')
if ($token -notmatch '^[A-Za-z0-9_-]{43}$') {
    $token = [Environment]::GetEnvironmentVariable($tokenVariable, 'User')
}
if ($token -notmatch '^[A-Za-z0-9_-]{43}$') {
    throw "$tokenVariable is not a canonical 32-byte base64url token in process or current-user scope."
}
if ($Endpoint.AbsoluteUri -ne 'http://127.0.0.1:8765/mcp/') {
    throw 'The Phase 2 production probe accepts only the canonical loopback endpoint.'
}

Add-Type -AssemblyName System.Net.Http
$handler = [Net.Http.HttpClientHandler]::new()
$handler.UseProxy = $false
$handler.AllowAutoRedirect = $false
$handler.AutomaticDecompression = [Net.DecompressionMethods]::None
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

function Invoke-McpPost {
    param(
        [Parameter(Mandatory = $true)][string]$Body,
        [bool]$Authenticated = $true,
        [AllowNull()][string]$ProtocolVersion = '2025-11-25'
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

        $response = $client.SendAsync($request).GetAwaiter().GetResult()
        try {
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
            $response.Dispose()
        }
    }
    finally {
        $request.Dispose()
    }
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

try {
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
