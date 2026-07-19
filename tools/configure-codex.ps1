#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('Validate', 'Install', 'Rotate', 'ClearToken')]
    [string]$Action = 'Validate',

    [string]$Token
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$configPath = Join-Path $repositoryRoot '.codex\config.toml'
$serverName = 'outlook_classic'
$tokenVariable = 'OUTLOOK_MCP_TOKEN'
$expectedUrl = 'http://127.0.0.1:8765/mcp/'

function New-BearerToken {
    $bytes = [byte[]]::new(32)
    $generator = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $generator.GetBytes($bytes)
    }
    finally {
        $generator.Dispose()
    }

    return [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function Test-BearerToken([string]$Value) {
    if ($Value -notmatch '^[A-Za-z0-9_-]{43}$') {
        return $false
    }

    $base64 = $Value.Replace('-', '+').Replace('_', '/') + '='
    $bytes = $null
    try {
        $bytes = [Convert]::FromBase64String($base64)
        if ($bytes.Length -ne 32) {
            return $false
        }

        $canonical = [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        return [string]::Equals($Value, $canonical, [StringComparison]::Ordinal)
    }
    catch {
        return $false
    }
    finally {
        if ($null -ne $bytes) {
            [Array]::Clear($bytes, 0, $bytes.Length)
        }
    }
}

function Invoke-CodexCommand {
    param(
        [Parameter(Mandatory)]
        [string]$WorkingDirectory,

        [Parameter(Mandatory)]
        [string[]]$CommandArguments,

        [switch]$CloseStandardInput
    )

    Push-Location $WorkingDirectory
    try {
        $previousErrorActionPreference = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            if ($CloseStandardInput) {
                $output = @($null | & codex @CommandArguments 2>&1)
            }
            else {
                $output = @(& codex @CommandArguments 2>&1)
            }
            $exitCode = $LASTEXITCODE
        }
        finally {
            $ErrorActionPreference = $previousErrorActionPreference
        }
    }
    finally {
        Pop-Location
    }

    return [pscustomobject]@{
        ExitCode = $exitCode
        Output = [string[]]@($output | ForEach-Object { "$_" })
    }
}

function Test-ProjectConfiguration {
    if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
        throw "The repository-owned Codex configuration is missing: $configPath"
    }

    $expectedConfiguration = @'
[mcp_servers.outlook_classic]
url = "http://127.0.0.1:8765/mcp/"
bearer_token_env_var = "OUTLOOK_MCP_TOKEN"
required = false
default_tools_approval_mode = "writes"
tool_timeout_sec = 30

[mcp_servers.outlook_classic.tools.outlook_send_draft]
approval_mode = "prompt"

[mcp_servers.outlook_classic.tools.outlook_delete_messages]
approval_mode = "prompt"
'@

    $actual = [IO.File]::ReadAllText($configPath).Replace("`r`n", "`n").Trim()
    $expected = $expectedConfiguration.Replace("`r`n", "`n").Trim()
    if (-not [string]::Equals($actual, $expected, [StringComparison]::Ordinal)) {
        throw 'The committed project configuration differs from the reviewed fail-closed configuration. Restore .codex\config.toml before continuing.'
    }

    $neutralDirectory = [Environment]::GetFolderPath([Environment+SpecialFolder]::Windows)
    $globalListResult = Invoke-CodexCommand -WorkingDirectory $neutralDirectory -CommandArguments @('mcp', 'list', '--json')
    if ($globalListResult.ExitCode -ne 0) {
        throw 'Codex could not enumerate user-scoped MCP registrations from a neutral directory.'
    }
    try {
        $globalServers = @((($globalListResult.Output -join [Environment]::NewLine) | ConvertFrom-Json))
    }
    catch {
        throw 'Codex returned an unreadable user-scoped MCP registration list.'
    }
    $globalNameCollision = @(
        $globalServers |
            Where-Object { $_.PSObject.Properties.Name -contains 'name' -and $_.name -eq $serverName }
    )
    if ($globalNameCollision.Count -gt 0) {
        throw "A user-scoped MCP registration named '$serverName' shadows the repository entry. Remove that global entry before validating project scope."
    }

    $strictResult = Invoke-CodexCommand `
        -WorkingDirectory $repositoryRoot `
        -CommandArguments @('app-server', '--strict-config', '--stdio') `
        -CloseStandardInput
    if ($strictResult.ExitCode -ne 0) {
        throw "Codex strict configuration validation failed: $($strictResult.Output -join [Environment]::NewLine)"
    }

    $serverResult = Invoke-CodexCommand `
        -WorkingDirectory $repositoryRoot `
        -CommandArguments @('mcp', 'get', $serverName, '--json')
    if ($serverResult.ExitCode -ne 0) {
        throw "Codex did not load the repository-scoped server. Trust this repository in Codex, then retry. Details: $($serverResult.Output -join [Environment]::NewLine)"
    }

    try {
        $server = $serverResult.Output -join [Environment]::NewLine | ConvertFrom-Json
    }
    catch {
        throw "Codex returned an unreadable MCP configuration: $($_.Exception.Message)"
    }

    if ($server.name -ne $serverName -or
        $server.transport.type -ne 'streamable_http' -or
        $server.transport.url -ne $expectedUrl -or
        $server.transport.bearer_token_env_var -ne $tokenVariable -or
        [double]$server.tool_timeout_sec -ne 30) {
        throw 'The effective Codex MCP registration does not match the repository configuration.'
    }
}

if (-not (Get-Command codex -ErrorAction SilentlyContinue)) {
    throw 'The Codex CLI is not available on PATH.'
}
if ($Action -in @('Validate', 'ClearToken') -and -not [string]::IsNullOrWhiteSpace($Token)) {
    throw "-Token is not valid with -Action $Action."
}

Test-ProjectConfiguration

$previousToken = [Environment]::GetEnvironmentVariable($tokenVariable, 'User')
$tokenChanged = $false
if ($Action -eq 'Install') {
    if ([string]::IsNullOrWhiteSpace($Token)) {
        $Token = if (Test-BearerToken $previousToken) { $previousToken } else { New-BearerToken }
    }
    if (-not (Test-BearerToken $Token)) {
        throw 'Token must be a 32-byte value encoded as 43-character base64url without padding.'
    }

    if (-not [string]::Equals($previousToken, $Token, [StringComparison]::Ordinal)) {
        [Environment]::SetEnvironmentVariable($tokenVariable, $Token, 'User')
        $tokenChanged = $true
    }
}
elseif ($Action -eq 'Rotate') {
    Write-Warning 'Token rotation takes effect after Outlook and Codex restart and invalidates outstanding authenticated cursors.'
    if ([string]::IsNullOrWhiteSpace($Token)) {
        $Token = New-BearerToken
    }
    if (-not (Test-BearerToken $Token)) {
        throw 'Token must be a 32-byte value encoded as 43-character base64url without padding.'
    }

    [Environment]::SetEnvironmentVariable($tokenVariable, $Token, 'User')
    $tokenChanged = -not [string]::Equals($previousToken, $Token, [StringComparison]::Ordinal)
}
elseif ($Action -eq 'ClearToken' -and -not [string]::IsNullOrWhiteSpace($previousToken)) {
    [Environment]::SetEnvironmentVariable($tokenVariable, $null, 'User')
    $tokenChanged = $true
}

$storedToken = [Environment]::GetEnvironmentVariable($tokenVariable, 'User')
if ($Action -in @('Install', 'Rotate') -and -not (Test-BearerToken $storedToken)) {
    throw "The current-user $tokenVariable value did not pass readback validation."
}
if ($Action -eq 'ClearToken' -and -not [string]::IsNullOrWhiteSpace($storedToken)) {
    throw "The current-user $tokenVariable value was not removed."
}

[pscustomobject]@{
    Action = $Action
    Scope = 'Project'
    ConfigPath = $configPath
    ConfigurationValidated = $true
    GlobalNameCollisionAbsent = $true
    RepositoryConfigurationLoadedByCodex = $true
    TokenVariable = $tokenVariable
    TokenChanged = $tokenChanged
    TokenStoredForCurrentUser = -not [string]::IsNullOrWhiteSpace($storedToken)
    RestartRequired = $tokenChanged
    ConsequentialToolsRemainImplementationGated = $true
}
