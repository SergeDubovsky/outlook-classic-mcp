#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Profile,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ExpectedStoreInventoryPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ReadFixturePath,

    [ValidateRange(15, 180)]
    [int]$StartupTimeoutSeconds = 60,

    [ValidateRange(15, 180)]
    [int]$ShutdownTimeoutSeconds = 60,

    [ValidateRange(60, 1000)]
    [int]$RepeatedReadCount = 200
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (($RepeatedReadCount % 20) -ne 0) {
    throw '-RepeatedReadCount must be divisible by twenty.'
}
if ($Profile.IndexOf('"') -ge 0 -or $Profile.IndexOf('\') -ge 0 -or
    $Profile -match '[\x00-\x1F]') {
    throw 'The Outlook profile name must not contain quotation marks, backslashes, or control characters.'
}
if (-not [Environment]::UserInteractive) {
    throw 'The Phase 4 smoke gate requires an interactive logged-on Windows desktop.'
}

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]::new($identity)
if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Run the Phase 4 smoke gate from a non-elevated PowerShell process.'
}

$repositoryRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'phase3-smoke-common.ps1')
. (Join-Path $PSScriptRoot 'phase4-smoke-common.ps1')

$ExpectedStoreInventoryPath = (
    Resolve-Path -LiteralPath $ExpectedStoreInventoryPath -ErrorAction Stop).Path
$ReadFixturePath = (Resolve-Path -LiteralPath $ReadFixturePath -ErrorAction Stop).Path
$expectedInventory = Import-Phase3ExpectedStoreInventory `
    -Path $ExpectedStoreInventoryPath `
    -RepositoryRoot $repositoryRoot `
    -AllowedSources @('classic-outlook-ui', 'conditional-vsto-seeder')
$fixture = Import-Phase4ReadFixture `
    -Path $ReadFixturePath `
    -RepositoryRoot $repositoryRoot `
    -ExpectedInventory $expectedInventory

$sharedSensitiveValues =
    [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
foreach ($sensitiveValue in $fixture.SensitiveValues) {
    $null = $sharedSensitiveValues.Add($sensitiveValue)
}
$sharedOperationIds =
    [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$endpointResults = [System.Collections.Generic.List[object]]::new()
$cycleResults = [System.Collections.Generic.List[object]]::new()

$outlookPath = Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\OUTLOOK.EXE'
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$addInProjectPath = Join-Path $repositoryRoot 'src\OutlookClassicMcp.AddIn\OutlookClassicMcp.AddIn.csproj'
$releaseManifestPath = Join-Path $repositoryRoot 'src\OutlookClassicMcp.AddIn\bin\Release\OutlookClassicMcp.AddIn.vsto'
$expectedManifest = ([Uri]::new($releaseManifestPath)).AbsoluteUri + '|vstolocal'
$addInSubKey = 'Software\Microsoft\Office\Outlook\Addins\OutlookClassicMcp.AddIn'
$profileSubKey = "Software\Microsoft\Office\16.0\Outlook\Profiles\$Profile"

function Open-Hkcu64Key {
    param([Parameter(Mandatory = $true)][string]$SubKey)

    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::CurrentUser,
        [Microsoft.Win32.RegistryView]::Registry64)
    try {
        return $base.OpenSubKey($SubKey)
    }
    finally {
        $base.Dispose()
    }
}

function Assert-ExactPhase4Registration {
    $key = Open-Hkcu64Key -SubKey $addInSubKey
    if ($null -eq $key) {
        throw "The 64-bit Outlook development registration is missing: HKCU:\$addInSubKey"
    }
    try {
        $names = @($key.GetValueNames())
        if ($names -notcontains 'LoadBehavior' -or
            $key.GetValueKind('LoadBehavior') -ne [Microsoft.Win32.RegistryValueKind]::DWord -or
            $key.GetValue('LoadBehavior') -ne 3) {
            throw 'The Phase 4 add-in registration is not LoadBehavior=3 as a DWORD.'
        }
        if ($names -notcontains 'Manifest' -or
            $key.GetValueKind('Manifest') -ne [Microsoft.Win32.RegistryValueKind]::String -or
            -not [string]::Equals(
                "$($key.GetValue('Manifest'))",
                $expectedManifest,
                [StringComparison]::OrdinalIgnoreCase)) {
            throw 'The Phase 4 add-in registration does not reference the exact Release manifest.'
        }
    }
    finally {
        $key.Dispose()
    }
}

function Assert-OutlookStoppedAndPortFree {
    $outlook = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
    if ($outlook.Count -ne 0) {
        throw "Outlook is running. Close it gracefully before the Phase 4 gate; it will never be terminated forcibly. Running process IDs: $($outlook.Id -join ', ')"
    }
    $listeners = @(Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue)
    if ($listeners.Count -ne 0) {
        throw "TCP port 8765 is already listening. Owning process IDs: $($listeners.OwningProcess -join ', ')"
    }
}

function Wait-ForOutlookReady {
    param(
        [Parameter(Mandatory = $true)][Diagnostics.Process]$Process,
        [Parameter(Mandatory = $true)][DateTime]$DeadlineUtc
    )

    do {
        if ($Process.HasExited) {
            throw "Outlook process $($Process.Id) exited before Phase 4 readiness."
        }
        $Process.Refresh()
        $soleOutlook = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
        $listener = @(Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue)
        if ($Process.MainWindowHandle -ne [IntPtr]::Zero -and $Process.Responding -and
            $soleOutlook.Count -eq 1 -and $soleOutlook[0].Id -eq $Process.Id -and
            $listener.Count -eq 1) {
            return
        }
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $DeadlineUtc)

    throw "Outlook process $($Process.Id) did not expose one responsive window and listener before the startup deadline."
}

function Wait-ForPortRelease {
    param([Parameter(Mandatory = $true)][Diagnostics.Stopwatch]$Stopwatch)

    do {
        $listeners = @(Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue)
        if ($listeners.Count -eq 0) {
            if ($Stopwatch.Elapsed -ge [TimeSpan]::FromSeconds(3)) {
                break
            }
            return $Stopwatch.Elapsed.TotalMilliseconds
        }
        Start-Sleep -Milliseconds 100
    } while ($Stopwatch.Elapsed -lt [TimeSpan]::FromSeconds(3))

    throw 'TCP port 8765 remained bound for at least three seconds after normal Outlook closure was requested.'
}

function Request-NormalOutlookCloseAfterFailure {
    param([AllowNull()][Diagnostics.Process]$Process)

    if ($null -eq $Process) {
        return
    }
    try {
        if ($Process.HasExited) {
            return
        }
        $Process.Refresh()
        if ($Process.MainWindowHandle -ne [IntPtr]::Zero -and $Process.CloseMainWindow()) {
            $null = $Process.WaitForExit($ShutdownTimeoutSeconds * 1000)
        }
    }
    catch {
        Write-Warning "Outlook did not accept the best-effort normal close request: $($_.Exception.Message)"
    }
}

function Invoke-Phase4VstoTarget {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('Build', 'VSTOClean')][string]$Target,
        [Parameter(Mandatory = $true)][string]$CertificateThumbprint
    )

    $arguments = @(
        $addInProjectPath,
        "/t:$Target",
        '/p:Configuration=Release',
        '/p:Platform=AnyCPU',
        '/p:VisualStudioVersion=18.0',
        '/p:SignManifests=true',
        "/p:ManifestCertificateThumbprint=$CertificateThumbprint",
        '/p:RestoreLockedMode=true',
        '/v:minimal',
        '/nologo',
        '/nr:false'
    )
    & $script:msbuild @arguments | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Release VSTO $Target failed with exit code $LASTEXITCODE."
    }
}

if (-not (Test-Path -LiteralPath $outlookPath -PathType Leaf)) {
    throw "Classic Outlook was not found at $outlookPath"
}
if (-not (Test-Path -LiteralPath $vswhere -PathType Leaf)) {
    throw "vswhere was not found at $vswhere"
}
$profileKey = Open-Hkcu64Key -SubKey $profileSubKey
if ($null -eq $profileKey) {
    throw "The Outlook profile '$Profile' does not exist in the 64-bit current-user profile registry."
}
$profileKey.Dispose()
Assert-OutlookStoppedAndPortFree
$preexistingAddInKey = Open-Hkcu64Key -SubKey $addInSubKey
if ($null -ne $preexistingAddInKey) {
    $preexistingAddInKey.Dispose()
    throw 'A pre-existing OutlookClassicMcp.AddIn registration is present. Clean or preserve it explicitly before the isolated Phase 4 gate.'
}

$requiredComponents = @(
    'Microsoft.VisualStudio.Workload.Office',
    'Microsoft.VisualStudio.Component.TeamOffice',
    'Microsoft.VisualStudio.Workload.ManagedDesktop',
    'Microsoft.Net.Component.4.8.TargetingPack',
    'Microsoft.Net.Component.4.8.SDK'
)
$vswhereArguments = @('-products', '*', '-version', '[18.0,19.0)', '-requires') +
    $requiredComponents + @('-format', 'json', '-utf8')
$instances = @(((@(& $vswhere @vswhereArguments) -join [Environment]::NewLine) |
    ConvertFrom-Json -ErrorAction Stop))
$visualStudio = $instances |
    Where-Object {
        $_.PSObject.Properties.Name -contains 'productId' -and
        $_.productId -ne 'Microsoft.VisualStudio.Product.BuildTools'
    } |
    Select-Object -First 1
if ($null -eq $visualStudio) {
    throw 'A full Visual Studio 2026 installation with the required Office/VSTO components was not found.'
}
$script:msbuild = Join-Path $visualStudio.installationPath 'MSBuild\Current\Bin\MSBuild.exe'
if (-not (Test-Path -LiteralPath $script:msbuild -PathType Leaf)) {
    throw "MSBuild was not found at $script:msbuild"
}

$tokenVariable = 'OUTLOOK_MCP_TOKEN'
$token = [Environment]::GetEnvironmentVariable($tokenVariable, 'User')
if ($token -cnotmatch '^[A-Za-z0-9_-]{43}$') {
    throw 'Phase 4 requires a canonical current-user OUTLOOK_MCP_TOKEN. Run tools\configure-codex.ps1 -Action Install first.'
}
Add-Phase4SensitiveValue -Set $sharedSensitiveValues -Value $token
$previousProcessToken = [Environment]::GetEnvironmentVariable($tokenVariable, 'Process')
$previousVstoLogAlerts = [Environment]::GetEnvironmentVariable('VSTO_LOGALERTS', 'Process')
$previousVstoSuppressAlerts = [Environment]::GetEnvironmentVariable('VSTO_SUPPRESSDISPLAYALERTS', 'Process')

$certificate = $null
$certificateThumbprint = $null
$certificateSubject = "CN=OutlookClassicMcp Phase4 Smoke $([Guid]::NewGuid().ToString('N'))"
$keyContainerName = "OutlookClassicMcp-Phase4-$([Guid]::NewGuid().ToString('N'))"
$keyProvider = [System.Security.Cryptography.CngProvider]::MicrosoftSoftwareKeyStorageProvider
$registrationAttempted = $false
$ownedProcess = $null
$codexResult = $null
$smokeStartedUtc = $null
$smokeFinishedUtc = $null
$verification = $null
$primaryFailure = $null
$cleanupFailures = [System.Collections.Generic.List[string]]::new()

Push-Location $repositoryRoot
try {
    try {
        & (Join-Path $PSScriptRoot 'build.ps1') -Configuration Release | Out-Host

        $certificate = New-SelfSignedCertificate `
            -Type CodeSigningCert `
            -Subject $certificateSubject `
            -CertStoreLocation 'Cert:\CurrentUser\My' `
            -NotAfter (Get-Date).AddDays(1) `
            -KeyAlgorithm RSA `
            -KeyLength 2048 `
            -HashAlgorithm SHA256 `
            -Provider $keyProvider.Provider `
            -Container $keyContainerName `
            -KeyExportPolicy NonExportable
        $certificateThumbprint = $certificate.Thumbprint
        $registrationAttempted = $true
        Invoke-Phase4VstoTarget -Target Build -CertificateThumbprint $certificateThumbprint
        & (Join-Path $PSScriptRoot 'verify-mcp-runtime-assets.ps1') -Configuration Release | Out-Null
        Assert-ExactPhase4Registration

        [Environment]::SetEnvironmentVariable($tokenVariable, $token, 'Process')
        [Environment]::SetEnvironmentVariable('VSTO_LOGALERTS', '1', 'Process')
        [Environment]::SetEnvironmentVariable('VSTO_SUPPRESSDISPLAYALERTS', '1', 'Process')
        $smokeStartedUtc = [DateTimeOffset]::UtcNow

        for ($cycle = 1; $cycle -le 3; $cycle++) {
            Assert-OutlookStoppedAndPortFree
            Assert-ExactPhase4Registration
            $ownedProcess = $null
            try {
                $startInfo = [Diagnostics.ProcessStartInfo]::new()
                $startInfo.FileName = $outlookPath
                $startInfo.Arguments = "/profile `"$Profile`""
                $startInfo.WorkingDirectory = Split-Path -Parent $outlookPath
                $startInfo.UseShellExecute = $false
                $ownedProcess = [Diagnostics.Process]::Start($startInfo)
                if ($null -eq $ownedProcess) {
                    throw "Outlook did not return a process for Phase 4 cycle $cycle."
                }
                Wait-ForOutlookReady `
                    -Process $ownedProcess `
                    -DeadlineUtc ([DateTime]::UtcNow.AddSeconds($StartupTimeoutSeconds))

                $mode = if ($cycle -eq 1) { 'Full' } else { 'Restart' }
                $endpointResult = & (Join-Path $PSScriptRoot 'test-phase4-endpoint.ps1') `
                    -ExpectedStoreInventoryPath $ExpectedStoreInventoryPath `
                    -ReadFixturePath $ReadFixturePath `
                    -OutlookProcessId $ownedProcess.Id `
                    -Mode $mode `
                    -TimeoutSeconds ([Math]::Min($StartupTimeoutSeconds, 60)) `
                    -RepeatedReadCount $RepeatedReadCount `
                    -CycleNumber $cycle `
                    -SharedOperationIds $sharedOperationIds `
                    -SharedSensitiveValues $sharedSensitiveValues
                $endpointResults.Add($endpointResult)

                if ($cycle -eq 1) {
                    $codexResult = & (Join-Path $PSScriptRoot 'test-phase4-codex.ps1') `
                        -TimeoutSeconds ([Math]::Min($StartupTimeoutSeconds + 30, 180)) `
                        -SharedOperationIds $sharedOperationIds `
                        -SharedSensitiveValues $sharedSensitiveValues
                }

                $closeStopwatch = [Diagnostics.Stopwatch]::StartNew()
                if (-not $ownedProcess.CloseMainWindow()) {
                    throw "Outlook process $($ownedProcess.Id) rejected the normal close request; it was not terminated."
                }
                $portReleaseMilliseconds = Wait-ForPortRelease -Stopwatch $closeStopwatch
                if (-not $ownedProcess.WaitForExit($ShutdownTimeoutSeconds * 1000)) {
                    throw "Outlook process $($ownedProcess.Id) did not exit normally before the shutdown deadline; it was not terminated."
                }
                $cycleResults.Add([pscustomobject]@{
                    Cycle = $cycle
                    ProcessId = $ownedProcess.Id
                    PortReleaseMilliseconds = $portReleaseMilliseconds
                })
                Assert-OutlookStoppedAndPortFree
                Write-Host "Phase 4 smoke cycle $cycle/3 passed."
            }
            catch {
                Request-NormalOutlookCloseAfterFailure -Process $ownedProcess
                throw
            }
            finally {
                if ($null -ne $ownedProcess) {
                    $ownedProcess.Dispose()
                    $ownedProcess = $null
                }
            }
        }

        $smokeFinishedUtc = [DateTimeOffset]::UtcNow
        $verification = & (Join-Path $PSScriptRoot 'verify-phase4-smoke.ps1') `
            -SinceUtc $smokeStartedUtc `
            -UntilUtc $smokeFinishedUtc `
            -EndpointResults $endpointResults.ToArray() `
            -CodexResult $codexResult `
            -SensitiveValues $sharedSensitiveValues `
            -OperationIds $sharedOperationIds `
            -ExpectedRepeatedReadCount $RepeatedReadCount `
            -ExpectedCycles 3
    }
    catch {
        $primaryFailure = $_
    }
    finally {
        [Environment]::SetEnvironmentVariable($tokenVariable, $previousProcessToken, 'Process')
        [Environment]::SetEnvironmentVariable('VSTO_LOGALERTS', $previousVstoLogAlerts, 'Process')
        [Environment]::SetEnvironmentVariable('VSTO_SUPPRESSDISPLAYALERTS', $previousVstoSuppressAlerts, 'Process')

        Request-NormalOutlookCloseAfterFailure -Process $ownedProcess
        $runningOutlook = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
        if ($runningOutlook.Count -ne 0 -and $registrationAttempted) {
            $cleanupFailures.Add(
                'Outlook is still running. The runner left the loaded registration and ephemeral certificate intact; close Outlook normally before cleanup.')
        }
        else {
            if ($registrationAttempted) {
                try {
                    Invoke-Phase4VstoTarget `
                        -Target VSTOClean `
                        -CertificateThumbprint $certificateThumbprint
                }
                catch {
                    $cleanupFailures.Add("Release VSTOClean failed: $($_.Exception.Message)")
                }
            }
            if ($null -ne $certificate) {
                $certificate.Dispose()
                $certificate = $null
            }
            try {
                foreach ($storedCertificate in @(
                        Get-ChildItem 'Cert:\CurrentUser\My' |
                            Where-Object Subject -eq $certificateSubject)) {
                    Remove-Item `
                        -LiteralPath $storedCertificate.PSPath `
                        -DeleteKey `
                        -Force `
                        -ErrorAction Stop
                }
                foreach ($authorityCertificate in @(
                        Get-ChildItem 'Cert:\CurrentUser\CA' |
                            Where-Object Subject -eq $certificateSubject)) {
                    Remove-Item -LiteralPath $authorityCertificate.PSPath -Force -ErrorAction Stop
                }
                if ([System.Security.Cryptography.CngKey]::Exists($keyContainerName, $keyProvider)) {
                    $key = [System.Security.Cryptography.CngKey]::Open($keyContainerName, $keyProvider)
                    try { $key.Delete() } finally { $key.Dispose() }
                }
            }
            catch {
                $cleanupFailures.Add("Ephemeral certificate cleanup failed: $($_.Exception.Message)")
            }
            $residue = @(
                Get-ChildItem 'Cert:\CurrentUser\My', 'Cert:\CurrentUser\CA' |
                    Where-Object Subject -eq $certificateSubject
            )
            if ($residue.Count -ne 0 -or
                [System.Security.Cryptography.CngKey]::Exists($keyContainerName, $keyProvider)) {
                $cleanupFailures.Add('Phase 4 ephemeral signing state remains after cleanup.')
            }
            if ($registrationAttempted) {
                $remainingAddInKey = Open-Hkcu64Key -SubKey $addInSubKey
                if ($null -ne $remainingAddInKey) {
                    $remainingAddInKey.Dispose()
                    $cleanupFailures.Add('The Phase 4 Outlook add-in registration remains after VSTOClean.')
                }
            }
        }
        if ($null -ne $certificate) {
            $certificate.Dispose()
        }
    }

    if ($null -ne $primaryFailure) {
        foreach ($cleanupFailure in $cleanupFailures) {
            Write-Warning $cleanupFailure
        }
        $PSCmdlet.ThrowTerminatingError($primaryFailure)
    }
    if ($cleanupFailures.Count -ne 0) {
        throw "Phase 4 smoke cleanup failed: $($cleanupFailures -join [Environment]::NewLine)"
    }

    $maximumPortReleaseMilliseconds = (
        $cycleResults | Measure-Object -Property PortReleaseMilliseconds -Maximum).Maximum
    if ($cycleResults.Count -ne 3 -or
        @($cycleResults | Select-Object -ExpandProperty ProcessId -Unique).Count -ne 3 -or
        $maximumPortReleaseMilliseconds -ge 3000) {
        throw 'The Phase 4 lifecycle aggregate is incomplete.'
    }

    [pscustomobject]@{
        Phase = 4
        VerifiedCycleCount = $verification.VerifiedCycleCount
        DistinctOutlookProcessCount = 3
        EndpointVerificationCount = $verification.EndpointVerificationCount
        FullEndpointWorkloadCount = $verification.FullEndpointWorkloadCount
        ToolCount = $verification.ToolCount
        StoreCount = $verification.StoreCount
        StoreInventoryMatched = $verification.StoreInventoryMatched
        KnownMessageStoreCount = $verification.KnownMessageStoreCount
        StaticPageCount = $verification.StaticPageCount
        StaticPageItemCount = $verification.StaticPageItemCount
        LargeFolderMinimumItemCount = $verification.LargeFolderMinimumItemCount
        PartialSearchVerified = $verification.PartialSearchVerified
        CancellationRecoveryVerified = $verification.CancellationRecoveryVerified
        StructuredTimeoutObserved = $verification.StructuredTimeoutObserved
        RepeatedReadCount = $verification.RepeatedReadCount
        DirectComTelemetryBalanced = $verification.DirectComTelemetryBalanced
        MaterializedItemHighWater = $verification.MaterializedItemHighWater
        ResourceStable = $verification.ResourceStable
        OutlookResponsive = $verification.OutlookResponsive
        ProtectedFixtureVerified = $verification.ProtectedFixtureVerified
        NativeCodexVerified = $verification.NativeCodexVerified
        UniqueOperationIdCount = $verification.UniqueOperationIdCount
        PrivacySentinelCount = $verification.PrivacySentinelCount
        DiagnosticsLogFileCount = $verification.DiagnosticsLogFileCount
        MaximumPortReleaseMilliseconds = [double]$maximumPortReleaseMilliseconds
        OutlookStopped = $verification.OutlookStopped
        PortReleased = $verification.PortReleased
        CleanupVerified = $true
    }
}
finally {
    Pop-Location
}
