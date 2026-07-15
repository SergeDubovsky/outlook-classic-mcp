#Requires -Version 5.1
[CmdletBinding()]
param(
    [switch]$SkipListenerProbe
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$prefix = 'http://127.0.0.1:8765/mcp/'
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$requiredComponents = @(
    'Microsoft.VisualStudio.Workload.Office',
    'Microsoft.VisualStudio.Component.TeamOffice',
    'Microsoft.VisualStudio.Workload.ManagedDesktop',
    'Microsoft.Net.Component.4.8.TargetingPack',
    'Microsoft.Net.Component.4.8.SDK'
)

$failures = [System.Collections.Generic.List[string]]::new()
$warnings = [System.Collections.Generic.List[string]]::new()
$windowsIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$windowsPrincipal = [Security.Principal.WindowsPrincipal]::new($windowsIdentity)
$isAdministrator = $windowsPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not (Test-Path -LiteralPath $vswhere -PathType Leaf)) {
    $failures.Add("vswhere was not found at $vswhere")
    $visualStudio = $null
}
else {
    $arguments = @(
        '-products', '*',
        '-version', '[18.0,19.0)',
        '-requires'
    ) + $requiredComponents + @('-format', 'json', '-utf8')

    $instanceJson = & $vswhere @arguments
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace(($instanceJson -join [Environment]::NewLine))) {
        $failures.Add('A full Visual Studio 2026 IDE is missing one or more required Office/VSTO, desktop, or .NET Framework 4.8 components.')
        $visualStudio = $null
    }
    else {
        $visualStudio = @((($instanceJson -join [Environment]::NewLine) | ConvertFrom-Json)) |
            Where-Object { $_.PSObject.Properties.Name -contains 'productId' -and $_.productId -ne 'Microsoft.VisualStudio.Product.BuildTools' } |
            Select-Object -First 1
        if ($null -eq $visualStudio) {
            $failures.Add('The required components were found only in Build Tools; Phase 0 requires a full Visual Studio 2026 IDE.')
        }
    }
}

$msbuildPath = $null
$officeTargetsPath = $null
$templatePath = $null
if ($null -ne $visualStudio) {
    $msbuildPath = Join-Path $visualStudio.installationPath 'MSBuild\Current\Bin\MSBuild.exe'
    $officeTargetsPath = Join-Path $visualStudio.installationPath 'MSBuild\Microsoft\VisualStudio\v18.0\OfficeTools\Microsoft.VisualStudio.Tools.Office.targets'
    $templatePath = Join-Path $visualStudio.installationPath 'Common7\IDE\ProjectTemplates\CSharp\Office\Addins\1033\VSTOOutlook15AddInV4\OutlookAddIn.vstemplate'

    foreach ($requiredPath in @($msbuildPath, $officeTargetsPath, $templatePath)) {
        if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
            $failures.Add("Required Visual Studio file is missing: $requiredPath")
        }
    }
}

$net48ReferencePath = Join-Path ${env:ProgramFiles(x86)} 'Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8\mscorlib.dll'
if (-not (Test-Path -LiteralPath $net48ReferencePath -PathType Leaf)) {
    $failures.Add(".NET Framework 4.8 reference assemblies are missing: $net48ReferencePath")
}

$outlookPath = Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\OUTLOOK.EXE'
$outlookVersion = $null
if (-not (Test-Path -LiteralPath $outlookPath -PathType Leaf)) {
    $failures.Add("Classic Outlook was not found at $outlookPath")
}
else {
    $outlookVersion = (Get-Item -LiteralPath $outlookPath).VersionInfo.FileVersion
}

$officePlatform = $null
$officeConfiguration = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
if ($null -ne $officeConfiguration) {
    $officePlatform = $officeConfiguration.Platform
}
if ($officePlatform -ne 'x64') {
    $failures.Add("Classic Outlook x64 is required; detected platform: $officePlatform")
}

$vstoRuntime = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R' -ErrorAction SilentlyContinue
if ($null -eq $vstoRuntime -or $vstoRuntime.VSTORFeature_CLR40 -ne 1) {
    $failures.Add('The VSTO 4 runtime is not registered as installed.')
}

$dotnetSdks = @()
$resolvedDotNetSdk = $null
if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    $failures.Add('The .NET SDK host is not available on PATH.')
}
else {
    Push-Location $repositoryRoot
    try {
        $previousErrorActionPreference = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            $resolvedSdkOutput = @(& dotnet --version 2>&1)
            $resolvedSdkExitCode = $LASTEXITCODE
            $dotnetSdks = @(& dotnet --list-sdks 2>&1)
            $listSdksExitCode = $LASTEXITCODE
        }
        finally {
            $ErrorActionPreference = $previousErrorActionPreference
        }
    }
    finally {
        Pop-Location
    }

    if ($resolvedSdkExitCode -eq 0 -and $resolvedSdkOutput.Count -gt 0) {
        $resolvedDotNetSdk = "$($resolvedSdkOutput[0])".Trim()
    }

    [Version]$parsedSdkVersion = $null
    $pinnedFeatureBandResolved =
        $listSdksExitCode -eq 0 -and
        [Version]::TryParse($resolvedDotNetSdk, [ref]$parsedSdkVersion) -and
        $parsedSdkVersion.Major -eq 10 -and
        $parsedSdkVersion.Minor -eq 0 -and
        $parsedSdkVersion.Build -ge 302 -and
        $parsedSdkVersion.Build -lt 400
    if (-not $pinnedFeatureBandResolved) {
        $failures.Add("global.json did not resolve the required .NET SDK 10.0.3xx feature band at patch 302 or newer. Resolved: $resolvedDotNetSdk")
    }
}

$nodeVersion = $null
if (Get-Command node -ErrorAction SilentlyContinue) {
    $nodeVersion = (& node --version).Trim()
}
if ([string]::IsNullOrWhiteSpace($nodeVersion) -or $nodeVersion -notmatch '^v2[2-9]\.') {
    $warnings.Add("Node 22 or newer is not installed; MCP Inspector remains deferred. Detected: $nodeVersion")
}

$outlookProcesses = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object Id, SessionId, StartTime)
if ($outlookProcesses.Count -gt 0) {
    $warnings.Add('Outlook is running. Save work and close it gracefully before registration or F5 tests; never terminate it forcibly.')
}

$portOwnersBefore = @(
    Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue |
        Select-Object LocalAddress, LocalPort, OwningProcess, State
)
if ($portOwnersBefore.Count -gt 0) {
    $failures.Add('Port 8765 is already in use.')
}

$listenerProbeSucceeded = $null
$listenerProbeAttempted = $false
$remainingListeners = $portOwnersBefore
if (-not $SkipListenerProbe -and $isAdministrator) {
    $failures.Add('The exact-prefix listener proof must run from a non-elevated PowerShell process; this process is elevated.')
}
elseif (-not $SkipListenerProbe -and $portOwnersBefore.Count -eq 0) {
    $listenerProbeAttempted = $true
    $listener = [System.Net.HttpListener]::new()
    try {
        $listener.Prefixes.Add($prefix)
        $listener.Start()
        $listenerProbeSucceeded = $listener.IsListening
        if (-not $listenerProbeSucceeded) {
            $failures.Add("HttpListener did not enter the listening state at $prefix")
        }
    }
    catch {
        $failures.Add("Non-elevated HttpListener proof failed at ${prefix}: $($_.Exception.Message)")
    }
    finally {
        if ($listener.IsListening) {
            $listener.Stop()
        }
        $listener.Close()
    }

    $deadline = [DateTime]::UtcNow.AddSeconds(3)
    do {
        $remainingListeners = @(
            Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue |
                Select-Object LocalAddress, LocalPort, OwningProcess, State
        )
        if ($remainingListeners.Count -eq 0) {
            break
        }
        Start-Sleep -Milliseconds 50
    } while ([DateTime]::UtcNow -lt $deadline)

    if ($remainingListeners.Count -ne 0) {
        $failures.Add('The Phase 0 listener probe did not release port 8765 within three seconds.')
    }
}

$urlAclOutput = @(& netsh http show urlacl 2>&1)
$urlAclQueryExitCode = $LASTEXITCODE
if ($urlAclQueryExitCode -ne 0) {
    $failures.Add("HTTP.sys URL ACL inspection failed with exit code $urlAclQueryExitCode.")
}
$urlAclReservationBlocks = [System.Collections.Generic.List[string]]::new()
for ($lineIndex = 0; $lineIndex -lt $urlAclOutput.Count; $lineIndex++) {
    if ($urlAclOutput[$lineIndex] -notmatch [regex]::Escape($prefix)) {
        continue
    }

    $block = [System.Collections.Generic.List[string]]::new()
    $block.Add("$($urlAclOutput[$lineIndex])")
    for ($followingIndex = $lineIndex + 1; $followingIndex -lt $urlAclOutput.Count; $followingIndex++) {
        if ([string]::IsNullOrWhiteSpace("$($urlAclOutput[$followingIndex])")) {
            break
        }
        $block.Add("$($urlAclOutput[$followingIndex])")
    }
    $urlAclReservationBlocks.Add(($block -join [Environment]::NewLine))
}
$urlAclPresent = $urlAclReservationBlocks.Count -gt 0

$codexVersion = $null
if (Get-Command codex -ErrorAction SilentlyContinue) {
    $codexVersion = (& codex --version).Trim()
}
else {
    $failures.Add('The Codex CLI is not available on PATH.')
}

$result = [pscustomobject]@{
    RepositoryRoot = $repositoryRoot
    VisualStudio = if ($null -eq $visualStudio) { $null } else { $visualStudio.displayName }
    VisualStudioVersion = if ($null -eq $visualStudio) { $null } else { $visualStudio.installationVersion }
    MSBuildPath = $msbuildPath
    OfficeTargetsPath = $officeTargetsPath
    OutlookTemplatePath = $templatePath
    Net48ReferenceAssemblies = Test-Path -LiteralPath $net48ReferencePath
    DotNetSdks = $dotnetSdks
    ResolvedDotNetSdk = $resolvedDotNetSdk
    VstoRuntimeVersion = if ($null -eq $vstoRuntime) { $null } else { $vstoRuntime.Version }
    OutlookPath = $outlookPath
    OutlookVersion = $outlookVersion
    OutlookPlatform = $officePlatform
    OutlookProcesses = $outlookProcesses
    NodeVersion = $nodeVersion
    CodexVersion = $codexVersion
    Endpoint = $prefix
    IsAdministrator = $isAdministrator
    PortWasFree = $portOwnersBefore.Count -eq 0
    PortOwnersBefore = $portOwnersBefore
    ListenerProbeAttempted = $listenerProbeAttempted
    ListenerProbeSkipped = [bool]$SkipListenerProbe
    ListenerProbeSucceeded = $listenerProbeSucceeded
    PortOwnersAfter = $remainingListeners
    UrlAclQueryExitCode = $urlAclQueryExitCode
    UrlAclReservationBlocks = $urlAclReservationBlocks.ToArray()
    UrlAclPresent = $urlAclPresent
    Warnings = $warnings.ToArray()
    Failures = $failures.ToArray()
}

$result | ConvertTo-Json -Depth 6

if ($failures.Count -gt 0) {
    throw "Preflight failed with $($failures.Count) blocking issue(s)."
}
