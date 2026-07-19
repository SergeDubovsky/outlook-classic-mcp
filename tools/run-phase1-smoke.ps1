#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Profile,

    [ValidateRange(15, 180)]
    [int]$StartupTimeoutSeconds = 60,

    [ValidateRange(15, 180)]
    [int]$ShutdownTimeoutSeconds = 60,

    [ValidateSet(1, 2)]
    [int]$ExpectedPhase = 1
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($ExpectedPhase -eq 1) {
    throw (
        'Phase 1 smoke mode is retired because the current add-in always starts the Phase 2 listener. ' +
        'Run tools\run-phase2-smoke.ps1 instead; historical results remain in docs\PHASE_1_EVIDENCE.md.')
}

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$addInProjectPath = Join-Path $repositoryRoot 'src\OutlookClassicMcp.AddIn\OutlookClassicMcp.AddIn.csproj'
$releaseDirectory = Join-Path $repositoryRoot 'src\OutlookClassicMcp.AddIn\bin\Release'
$releaseManifestPath = Join-Path $releaseDirectory 'OutlookClassicMcp.AddIn.vsto'
$expectedManifest = ([Uri]::new($releaseManifestPath)).AbsoluteUri + '|vstolocal'
$outlookPath = Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\OUTLOOK.EXE'
$logDirectory = Join-Path $env:LOCALAPPDATA 'OutlookClassicMcp\logs'
$addInProgId = 'OutlookClassicMcp.AddIn'
$addInSubKey = "Software\Microsoft\Office\Outlook\Addins\$addInProgId"
$profileSubKey = "Software\Microsoft\Office\16.0\Outlook\Profiles\$Profile"
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$requiredComponents = @(
    'Microsoft.VisualStudio.Workload.Office'
    'Microsoft.VisualStudio.Component.TeamOffice'
    'Microsoft.VisualStudio.Workload.ManagedDesktop'
    'Microsoft.Net.Component.4.8.TargetingPack'
    'Microsoft.Net.Component.4.8.SDK'
)
$resiliencySubKeys = @(
    'Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems'
    'Software\Microsoft\Office\16.0\Outlook\Resiliency\CrashingAddinList'
)

function Get-RegistryQueryMatches {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SubKey,

        [string]$Search
    )

    $matches = [System.Collections.Generic.List[string]]::new()
    foreach ($view in @('32', '64')) {
        $arguments = @('query', "HKCU\$SubKey")
        if (-not [string]::IsNullOrWhiteSpace($Search)) {
            $arguments += @('/s', '/f', $Search)
        }
        $arguments += "/reg:$view"

        $previousErrorActionPreference = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            $output = @(& reg.exe @arguments 2>&1)
            $exitCode = $LASTEXITCODE
            $global:LASTEXITCODE = 0
        }
        finally {
            $ErrorActionPreference = $previousErrorActionPreference
        }

        if ($exitCode -eq 0) {
            foreach ($line in $output) {
                $matches.Add("registry-$view`: $line")
            }
        }
        elseif ($exitCode -ne 1) {
            throw "Registry query failed for HKCU\$SubKey in the $view-bit view with exit code $exitCode."
        }
    }

    return $matches.ToArray()
}

function Get-VstoDevelopmentState {
    $repositoryUri = ([Uri]::new(($repositoryRoot.TrimEnd('\') + '\'))).AbsoluteUri
    $state = [System.Collections.Generic.List[string]]::new()

    foreach ($line in (Get-RegistryQueryMatches -SubKey $addInSubKey)) {
        $state.Add("Outlook add-in: $line")
    }
    foreach ($line in (Get-RegistryQueryMatches -SubKey 'Software\Microsoft\Office\Outlook\FormRegions' -Search $addInProgId)) {
        $state.Add("Outlook form region: $line")
    }
    foreach ($line in (Get-RegistryQueryMatches -SubKey 'Software\Microsoft\VSTO\Security\Inclusion' -Search $repositoryUri)) {
        $state.Add("VSTO inclusion: $line")
    }
    foreach ($line in (Get-RegistryQueryMatches -SubKey 'Software\Microsoft\VSTO\SolutionMetadata' -Search $repositoryUri)) {
        $state.Add("VSTO metadata: $line")
    }

    return $state.ToArray()
}

function Open-RegistryKey {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.Win32.RegistryHive]$Hive,

        [Parameter(Mandatory = $true)]
        [string]$SubKey
    )

    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        $Hive,
        [Microsoft.Win32.RegistryView]::Registry64)
    try {
        return $baseKey.OpenSubKey($SubKey)
    }
    finally {
        $baseKey.Dispose()
    }
}

function Test-RegistryValueName {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.Win32.RegistryHive]$Hive,

        [Parameter(Mandatory = $true)]
        [string]$SubKey,

        [Parameter(Mandatory = $true)]
        [string]$ValueName
    )

    $key = Open-RegistryKey -Hive $Hive -SubKey $SubKey
    if ($null -eq $key) {
        return $false
    }
    try {
        return @($key.GetValueNames() | Where-Object {
            [string]::Equals($_, $ValueName, [StringComparison]::OrdinalIgnoreCase)
        }).Count -gt 0
    }
    finally {
        $key.Dispose()
    }
}

function Assert-AddInLoggingEnabled {
    $loggingSubKey = 'Software\Policies\Microsoft\Office\16.0\Outlook\Options\Logging'
    foreach ($hive in @(
        [Microsoft.Win32.RegistryHive]::CurrentUser,
        [Microsoft.Win32.RegistryHive]::LocalMachine)) {
        $key = Open-RegistryKey -Hive $hive -SubKey $loggingSubKey
        if ($null -eq $key) {
            continue
        }
        try {
            $matchingNames = @($key.GetValueNames() | Where-Object {
                [string]::Equals($_, 'DisableAddinLogging', [StringComparison]::OrdinalIgnoreCase)
            })
            if ($matchingNames.Count -eq 0) {
                continue
            }

            $name = $matchingNames[0]
            $kind = $key.GetValueKind($name)
            $value = $key.GetValue($name)
            if ($kind -ne [Microsoft.Win32.RegistryValueKind]::DWord -or $value -ne 0) {
                throw 'Outlook add-in event logging is disabled or has an invalid policy value. Event ID 45 is required for the lifecycle smoke gate.'
            }
        }
        finally {
            $key.Dispose()
        }
    }
}

function Assert-AddInNotPolicyManaged {
    $locations = @(
        [pscustomobject]@{
            Hive = [Microsoft.Win32.RegistryHive]::CurrentUser
            SubKey = 'Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList'
        }
        [pscustomobject]@{
            Hive = [Microsoft.Win32.RegistryHive]::CurrentUser
            SubKey = 'Software\Microsoft\Office\16.0\Outlook\Resiliency\AddinList'
        }
        [pscustomobject]@{
            Hive = [Microsoft.Win32.RegistryHive]::CurrentUser
            SubKey = 'Software\Policies\Microsoft\Office\16.0\Outlook\Resiliency\AddinList'
        }
        [pscustomobject]@{
            Hive = [Microsoft.Win32.RegistryHive]::LocalMachine
            SubKey = 'Software\Policies\Microsoft\Office\16.0\Outlook\Resiliency\AddinList'
        }
    )

    foreach ($location in $locations) {
        if (Test-RegistryValueName -Hive $location.Hive -SubKey $location.SubKey -ValueName $addInProgId) {
            throw "The target add-in is managed by Outlook resiliency policy at $($location.Hive)\$($location.SubKey). Remove that test exception before running the lifecycle smoke gate."
        }
    }
}

function ConvertTo-RegistrySnapshotValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [Microsoft.Win32.RegistryValueKind]$Kind
    )

    if ($Kind -eq [Microsoft.Win32.RegistryValueKind]::Binary) {
        return [Convert]::ToBase64String([byte[]]$Value)
    }
    if ($Kind -eq [Microsoft.Win32.RegistryValueKind]::MultiString) {
        return [Convert]::ToBase64String(
            [Text.Encoding]::UTF8.GetBytes((([string[]]$Value) -join "`0")))
    }

    return [Convert]::ToBase64String(
        [Text.Encoding]::UTF8.GetBytes([Convert]::ToString(
            $Value,
            [Globalization.CultureInfo]::InvariantCulture)))
}

function Get-ResiliencySnapshot {
    $snapshot = [ordered]@{}
    foreach ($subKey in $resiliencySubKeys) {
        $key = Open-RegistryKey -Hive ([Microsoft.Win32.RegistryHive]::CurrentUser) -SubKey $subKey
        if ($null -eq $key) {
            $snapshot[$subKey] = 'absent'
            continue
        }
        try {
            $entries = [System.Collections.Generic.List[string]]::new()
            foreach ($name in @($key.GetValueNames() | Sort-Object)) {
                $kind = $key.GetValueKind($name)
                $value = $key.GetValue(
                    $name,
                    $null,
                    [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                $encodedName = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($name))
                $encodedValue = ConvertTo-RegistrySnapshotValue -Value $value -Kind $kind
                $entries.Add("$encodedName|$kind|$encodedValue")
            }
            $snapshot[$subKey] = 'present:' + ($entries -join ';')
        }
        finally {
            $key.Dispose()
        }
    }

    return ($snapshot | ConvertTo-Json -Compress)
}

function Assert-ExactRegistration {
    $key = Open-RegistryKey -Hive ([Microsoft.Win32.RegistryHive]::CurrentUser) -SubKey $addInSubKey
    if ($null -eq $key) {
        throw "The 64-bit Outlook development registration is missing: HKCU:\$addInSubKey"
    }
    try {
        $valueNames = @($key.GetValueNames())
        if ($valueNames -notcontains 'LoadBehavior' -or
            $key.GetValueKind('LoadBehavior') -ne [Microsoft.Win32.RegistryValueKind]::DWord -or
            $key.GetValue('LoadBehavior') -ne 3) {
            throw 'The target add-in is not registered with LoadBehavior=3 as a DWORD.'
        }
        if ($valueNames -notcontains 'Manifest' -or
            $key.GetValueKind('Manifest') -ne [Microsoft.Win32.RegistryValueKind]::String -or
            -not [string]::Equals(
                "$($key.GetValue('Manifest'))",
                $expectedManifest,
                [StringComparison]::OrdinalIgnoreCase)) {
            throw "The target add-in does not reference the exact Release manifest: $expectedManifest"
        }
    }
    finally {
        $key.Dispose()
    }
}

function Assert-OutlookStoppedAndPortFree {
    $outlookProcesses = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
    if ($outlookProcesses.Count -gt 0) {
        throw "Outlook is running. Save open work and close Outlook gracefully before the smoke gate; never terminate it forcibly. Running process IDs: $($outlookProcesses.Id -join ', ')"
    }

    $listeners = @(Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue)
    if ($listeners.Count -gt 0) {
        throw "TCP port 8765 is already listening. Owning process IDs: $($listeners.OwningProcess -join ', ')"
    }
}

function Invoke-VstoTarget {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Build', 'VSTOClean')]
        [string]$Target,

        [Parameter(Mandatory = $true)]
        [string]$CertificateThumbprint
    )

    $arguments = @(
        $addInProjectPath
        "/t:$Target"
        '/p:Configuration=Release'
        '/p:Platform=AnyCPU'
        '/p:VisualStudioVersion=18.0'
        '/p:SignManifests=true'
        "/p:ManifestCertificateThumbprint=$CertificateThumbprint"
        '/p:RestoreLockedMode=true'
        '/v:minimal'
        '/nologo'
        '/nr:false'
    )

    & $script:msbuild @arguments | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Release VSTO $Target failed with exit code $LASTEXITCODE."
    }
}

function Get-ApplicationEventWatermark {
    $latestEvent = Get-WinEvent -LogName Application -MaxEvents 1 -ErrorAction Stop
    return [long]$latestEvent.RecordId
}

function Get-EventProcessId {
    param([Parameter(Mandatory = $true)][object]$EventRecord)

    [xml]$eventXml = $EventRecord.ToXml()
    return [int]$eventXml.Event.System.Execution.ProcessID
}

function Get-TargetLoadBehaviors {
    param([Parameter(Mandatory = $true)][string]$Message)

    $behaviors = [System.Collections.Generic.List[int]]::new()
    $lines = @($Message -split "\r?\n")
    for ($index = 0; $index -lt $lines.Count; $index++) {
        if (-not [string]::Equals(
            $lines[$index].Trim(),
            "ProgID: $addInProgId",
            [StringComparison]::Ordinal)) {
            continue
        }

        $foundLoadBehavior = $false
        for ($following = $index + 1; $following -lt $lines.Count; $following++) {
            $line = $lines[$following].Trim()
            if ($line.StartsWith('Name:', [StringComparison]::Ordinal) -or
                $line.StartsWith('ProgID:', [StringComparison]::Ordinal)) {
                break
            }
            if ($line -match '^Load Behavior:\s*(\d+)$') {
                $behaviors.Add([int]$Matches[1])
                $foundLoadBehavior = $true
                break
            }
        }
        if (-not $foundLoadBehavior) {
            $behaviors.Add(-1)
        }
    }

    return $behaviors.ToArray()
}

function Get-TargetOutlookEvents {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet(45, 59)]
        [int]$EventId,

        [Parameter(Mandatory = $true)]
        [long]$AfterRecordId,

        [Parameter(Mandatory = $true)]
        [DateTimeOffset]$SinceUtc,

        [int]$ProcessId
    )

    $events = @(Get-WinEvent -FilterHashtable @{
        LogName = 'Application'
        ProviderName = 'Outlook'
        Id = $EventId
        StartTime = $SinceUtc.LocalDateTime.AddSeconds(-2)
    } -ErrorAction SilentlyContinue)
    $matches = [System.Collections.Generic.List[object]]::new()
    foreach ($eventRecord in $events) {
        if ([long]$eventRecord.RecordId -le $AfterRecordId) {
            continue
        }
        $eventProcessId = Get-EventProcessId -EventRecord $eventRecord
        if ($PSBoundParameters.ContainsKey('ProcessId') -and $eventProcessId -ne $ProcessId) {
            continue
        }

        if ($EventId -eq 45) {
            $loadBehaviors = @(Get-TargetLoadBehaviors -Message "$($eventRecord.Message)")
            if ($loadBehaviors.Count -eq 0) {
                continue
            }
            $matches.Add([pscustomobject]@{
                RecordId = [long]$eventRecord.RecordId
                ProcessId = $eventProcessId
                TimeCreated = $eventRecord.TimeCreated
                LoadBehaviors = $loadBehaviors
            })
        }
        elseif ("$($eventRecord.Message)" -match "(?im)^\s*ProgID:\s*$([regex]::Escape($addInProgId))\s*$") {
            $matches.Add([pscustomobject]@{
                RecordId = [long]$eventRecord.RecordId
                ProcessId = $eventProcessId
                TimeCreated = $eventRecord.TimeCreated
            })
        }
    }

    return $matches.ToArray()
}

function ConvertFrom-SmokeDiagnosticLine {
    param([Parameter(Mandatory = $true)][string]$Line)

    $fields = [Collections.Generic.Dictionary[string, string]]::new(
        [StringComparer]::Ordinal)
    foreach ($part in $Line.Split("`t")) {
        $separator = $part.IndexOf('=')
        if ($separator -le 0) {
            return $null
        }
        $key = $part.Substring(0, $separator)
        if ($fields.ContainsKey($key)) {
            return $null
        }
        $fields.Add($key, $part.Substring($separator + 1))
    }

    foreach ($requiredField in @('timestamp', 'session', 'event', 'state', 'result')) {
        if (-not $fields.ContainsKey($requiredField)) {
            return $null
        }
    }

    [DateTimeOffset]$timestamp = [DateTimeOffset]::MinValue
    if (-not [DateTimeOffset]::TryParse(
        $fields['timestamp'],
        [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::RoundtripKind,
        [ref]$timestamp)) {
        return $null
    }

    [Guid]$sessionId = [Guid]::Empty
    if (-not [Guid]::TryParseExact($fields['session'], 'N', [ref]$sessionId)) {
        return $null
    }

    return [pscustomobject]@{
        Timestamp = $timestamp
        Session = $fields['session']
        Event = $fields['event']
        State = $fields['state']
        Result = $fields['result']
    }
}

function Get-SmokeDiagnosticRecords {
    param([Parameter(Mandatory = $true)][DateTimeOffset]$SinceUtc)

    if (-not (Test-Path -LiteralPath $logDirectory -PathType Container)) {
        return @()
    }

    $records = [System.Collections.Generic.List[object]]::new()
    foreach ($logFile in @(Get-ChildItem -LiteralPath $logDirectory -Filter 'outlook-classic-mcp*.log' -File)) {
        try {
            $stream = [IO.FileStream]::new(
                $logFile.FullName,
                [IO.FileMode]::Open,
                [IO.FileAccess]::Read,
                [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
            $reader = [IO.StreamReader]::new(
                $stream,
                [Text.Encoding]::UTF8,
                $true,
                4096,
                $false)
            try {
                while (($line = $reader.ReadLine()) -ne $null) {
                    if ([string]::IsNullOrWhiteSpace($line)) {
                        continue
                    }
                    $record = ConvertFrom-SmokeDiagnosticLine -Line $line
                    if ($null -ne $record -and $record.Timestamp -ge $SinceUtc.ToUniversalTime()) {
                        $records.Add($record)
                    }
                }
            }
            finally {
                $reader.Dispose()
            }
        }
        catch [IO.IOException] {
            continue
        }
    }

    return $records.ToArray()
}

function Wait-ForHostQuiescent {
    param(
        [Parameter(Mandatory = $true)][DateTimeOffset]$SinceUtc,
        [Parameter(Mandatory = $true)][DateTime]$DeadlineUtc
    )

    do {
        $records = @(Get-SmokeDiagnosticRecords -SinceUtc $SinceUtc)
        $matches = @($records | Where-Object Event -eq 'host_quiescent')
        if ($matches.Count -gt 1) {
            throw "More than one diagnostic session reached host_quiescent after $SinceUtc."
        }
        if ($matches.Count -eq 1) {
            $record = $matches[0]
            if ($record.Result -ne 'success' -or $record.State -ne 'online') {
                throw "The Outlook host reached '$($record.State)' with result '$($record.Result)'."
            }
            return $record.Session
        }
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $DeadlineUtc)

    throw "Outlook did not record a successful host_quiescent event within $StartupTimeoutSeconds seconds."
}

function Wait-ForShutdownCompleted {
    param(
        [Parameter(Mandatory = $true)][string]$SessionId,
        [Parameter(Mandatory = $true)][DateTimeOffset]$SinceUtc,
        [Parameter(Mandatory = $true)][DateTime]$DeadlineUtc
    )

    do {
        $matches = @(Get-SmokeDiagnosticRecords -SinceUtc $SinceUtc | Where-Object {
            $_.Session -eq $SessionId -and $_.Event -eq 'shutdown_completed'
        })
        if ($matches.Count -gt 1) {
            throw "Session $SessionId recorded more than one shutdown_completed event."
        }
        if ($matches.Count -eq 1) {
            $record = $matches[0]
            if ($record.Result -ne 'success' -or $record.State -ne 'stopped') {
                throw "Session $SessionId shutdown reached '$($record.State)' with result '$($record.Result)'."
            }
            return
        }
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $DeadlineUtc)

    throw "Session $SessionId did not record a successful shutdown_completed event within $ShutdownTimeoutSeconds seconds."
}

function Wait-ForOutlookEvent45 {
    param(
        [Parameter(Mandatory = $true)][int]$ProcessId,
        [Parameter(Mandatory = $true)][long]$AfterRecordId,
        [Parameter(Mandatory = $true)][DateTimeOffset]$SinceUtc,
        [Parameter(Mandatory = $true)][DateTime]$DeadlineUtc
    )

    do {
        $matches = @(Get-TargetOutlookEvents `
            -EventId 45 `
            -AfterRecordId $AfterRecordId `
            -SinceUtc $SinceUtc `
            -ProcessId $ProcessId)
        if ($matches.Count -gt 0) {
            return
        }
        Start-Sleep -Milliseconds 250
    } while ([DateTime]::UtcNow -lt $DeadlineUtc)

    throw "Outlook process $ProcessId did not emit Event ID 45 for $addInProgId within $StartupTimeoutSeconds seconds."
}

function Wait-ForOutlookMainWindow {
    param(
        [Parameter(Mandatory = $true)][Diagnostics.Process]$Process,
        [Parameter(Mandatory = $true)][DateTime]$DeadlineUtc
    )

    do {
        if ($Process.HasExited) {
            throw "Outlook process $($Process.Id) exited before its main window became available."
        }
        $Process.Refresh()
        if ($Process.MainWindowHandle -ne [IntPtr]::Zero -and $Process.Responding) {
            return
        }
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $DeadlineUtc)

    throw "Outlook process $($Process.Id) did not expose a responsive main window within $StartupTimeoutSeconds seconds."
}

function Get-VstoAlertSnapshot {
    $snapshot = @{}
    foreach ($directory in @($releaseDirectory, $env:TEMP) | Select-Object -Unique) {
        if ([string]::IsNullOrWhiteSpace($directory) -or
            -not (Test-Path -LiteralPath $directory -PathType Container)) {
            continue
        }
        foreach ($file in @(Get-ChildItem -LiteralPath $directory -Filter '*.vsto.log' -File -ErrorAction SilentlyContinue)) {
            $snapshot[$file.FullName] = [pscustomobject]@{
                Length = $file.Length
                LastWriteTimeUtc = $file.LastWriteTimeUtc
            }
        }
    }
    return $snapshot
}

function Get-ChangedVstoAlertLogs {
    param([Parameter(Mandatory = $true)][hashtable]$Before)

    $after = Get-VstoAlertSnapshot
    $changes = [System.Collections.Generic.List[string]]::new()
    foreach ($path in $after.Keys) {
        $current = $after[$path]
        if ($current.Length -eq 0) {
            continue
        }
        if (-not $Before.ContainsKey($path) -or
            $Before[$path].Length -ne $current.Length -or
            $Before[$path].LastWriteTimeUtc -ne $current.LastWriteTimeUtc) {
            $changes.Add($path)
        }
    }
    return $changes.ToArray()
}

function Close-OwnedOutlookAfterFailure {
    param([AllowNull()][Diagnostics.Process]$Process)

    if ($null -eq $Process) {
        return
    }
    try {
        if ($Process.HasExited) {
            return
        }
        $Process.Refresh()
        if ($Process.MainWindowHandle -eq [IntPtr]::Zero) {
            return
        }
        if ($Process.CloseMainWindow()) {
            $null = $Process.WaitForExit($ShutdownTimeoutSeconds * 1000)
        }
    }
    catch {
        Write-Warning "The runner could not request normal closure for Outlook process $($Process.Id): $($_.Exception.Message)"
    }
}

function Wait-ForLoopbackPortRelease {
    param([Parameter(Mandatory = $true)][Diagnostics.Stopwatch]$Stopwatch)

    do {
        $listeners = @(Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue)
        if ($listeners.Count -eq 0) {
            $elapsed = $Stopwatch.Elapsed
            if ($elapsed -lt [TimeSpan]::FromSeconds(3)) {
                return $elapsed.TotalMilliseconds
            }

            break
        }
        Start-Sleep -Milliseconds 100
    } while ($Stopwatch.Elapsed -lt [TimeSpan]::FromSeconds(3))

    throw "TCP port 8765 remained bound for at least three seconds after Outlook close was requested."
}

if ($Profile.IndexOf('"') -ge 0 -or $Profile.IndexOf('\') -ge 0 -or $Profile -match '[\x00-\x1F]') {
    throw 'The Outlook profile name must not contain quotation marks, backslashes, or control characters.'
}
if (-not [Environment]::UserInteractive) {
    throw 'The Outlook lifecycle smoke gate requires an interactive logged-on Windows desktop.'
}

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]::new($identity)
if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Run the Outlook lifecycle smoke gate from a non-elevated PowerShell process.'
}
if (-not (Test-Path -LiteralPath $outlookPath -PathType Leaf)) {
    throw "Classic Outlook was not found at $outlookPath"
}
if (-not (Test-Path -LiteralPath $vswhere -PathType Leaf)) {
    throw "vswhere was not found at $vswhere"
}

$profileKey = Open-RegistryKey -Hive ([Microsoft.Win32.RegistryHive]::CurrentUser) -SubKey $profileSubKey
if ($null -eq $profileKey) {
    throw "The Outlook profile '$Profile' does not exist under HKCU:\$profileSubKey"
}
$profileKey.Dispose()

$vswhereArguments = @(
    '-products', '*',
    '-version', '[18.0,19.0)',
    '-requires'
) + $requiredComponents + @('-format', 'json', '-utf8')
$instanceJson = @(& $vswhere @vswhereArguments)
$instances = @((($instanceJson -join [Environment]::NewLine) | ConvertFrom-Json))
$visualStudio = $instances |
    Where-Object { $_.PSObject.Properties.Name -contains 'productId' -and $_.productId -ne 'Microsoft.VisualStudio.Product.BuildTools' } |
    Select-Object -First 1
if ($null -eq $visualStudio) {
    throw 'A full Visual Studio 2026 installation with the required Office/VSTO components was not found.'
}
$script:msbuild = Join-Path $visualStudio.installationPath 'MSBuild\Current\Bin\MSBuild.exe'
if (-not (Test-Path -LiteralPath $script:msbuild -PathType Leaf)) {
    throw "MSBuild was not found at $script:msbuild"
}

Assert-OutlookStoppedAndPortFree
Assert-AddInLoggingEnabled
Assert-AddInNotPolicyManaged
$null = Get-WinEvent -ListLog Application -ErrorAction Stop
$initialVstoState = @(Get-VstoDevelopmentState)
if ($initialVstoState.Count -gt 0) {
    throw "A VSTO registration for this checkout or add-in already exists. Clean it before running the smoke gate. Detected state: $($initialVstoState -join [Environment]::NewLine)"
}
$smokeCertificateResidue = @(
    Get-ChildItem 'Cert:\CurrentUser\My', 'Cert:\CurrentUser\CA' |
        Where-Object Subject -Like 'CN=OutlookClassicMcp Ephemeral Smoke *'
)
if ($smokeCertificateResidue.Count -gt 0) {
    throw 'A certificate from an interrupted lifecycle smoke run remains in a current-user certificate store.'
}

$primaryFailure = $null
$cleanupFailures = [System.Collections.Generic.List[string]]::new()
$certificate = $null
$certificateThumbprint = $null
$certificateSubject = "CN=OutlookClassicMcp Ephemeral Smoke $([Guid]::NewGuid().ToString('N'))"
$keyContainerName = "OutlookClassicMcp-Smoke-$([Guid]::NewGuid().ToString('N'))"
$keyProvider = [System.Security.Cryptography.CngProvider]::MicrosoftSoftwareKeyStorageProvider
$keyUniqueName = $null
$keyFilePath = $null
$registrationAttempted = $false
$vstoLogAlertsBefore = $null
$previousVstoLogAlerts = [Environment]::GetEnvironmentVariable('VSTO_LOGALERTS', 'Process')
$previousVstoSuppressAlerts = [Environment]::GetEnvironmentVariable('VSTO_SUPPRESSDISPLAYALERTS', 'Process')
$cycleResults = [System.Collections.Generic.List[object]]::new()
$smokeStartedUtc = $null
$smokeFinishedUtc = $null
$verificationResult = $null
$previousOutlookMcpToken = [Environment]::GetEnvironmentVariable('OUTLOOK_MCP_TOKEN', 'Process')
if ($ExpectedPhase -ge 2) {
    $phaseToken = [Environment]::GetEnvironmentVariable('OUTLOOK_MCP_TOKEN', 'User')
    if ($phaseToken -notmatch '^[A-Za-z0-9_-]{43}$') {
        throw 'Phase 2 requires a canonical current-user OUTLOOK_MCP_TOKEN. Run tools\configure-codex.ps1 -Action Install first.'
    }
    [Environment]::SetEnvironmentVariable('OUTLOOK_MCP_TOKEN', $phaseToken, 'Process')
}

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

        $createdKey = [System.Security.Cryptography.CngKey]::Open($keyContainerName, $keyProvider)
        try {
            $keyUniqueName = $createdKey.UniqueName
        }
        finally {
            $createdKey.Dispose()
        }
        $keyFilePath = Join-Path (
            [Environment]::GetFolderPath([Environment+SpecialFolder]::ApplicationData)) (
            "Microsoft\Crypto\Keys\$keyUniqueName")

        $registrationAttempted = $true
        Invoke-VstoTarget -Target Build -CertificateThumbprint $certificateThumbprint
        & (Join-Path $PSScriptRoot 'verify-mcp-runtime-assets.ps1') -Configuration Release | Out-Null
        Assert-ExactRegistration

        [Environment]::SetEnvironmentVariable('VSTO_LOGALERTS', '1', 'Process')
        [Environment]::SetEnvironmentVariable('VSTO_SUPPRESSDISPLAYALERTS', '1', 'Process')
        $vstoLogAlertsBefore = Get-VstoAlertSnapshot
        $initialResiliencySnapshot = Get-ResiliencySnapshot
        $smokeStartedUtc = [DateTimeOffset]::UtcNow

        for ($cycle = 1; $cycle -le 3; $cycle++) {
            Assert-OutlookStoppedAndPortFree
            Assert-ExactRegistration
            $cycleStartedUtc = [DateTimeOffset]::UtcNow
            $eventWatermark = Get-ApplicationEventWatermark
            $ownedProcess = $null
            $normalCloseRequested = $false
            try {
                $startInfo = [Diagnostics.ProcessStartInfo]::new()
                $startInfo.FileName = $outlookPath
                $startInfo.Arguments = "/profile `"$Profile`""
                $startInfo.WorkingDirectory = Split-Path -Parent $outlookPath
                $startInfo.UseShellExecute = $false
                $ownedProcess = [Diagnostics.Process]::Start($startInfo)
                if ($null -eq $ownedProcess) {
                    throw "Outlook did not return a process for smoke cycle $cycle."
                }

                $startupDeadline = [DateTime]::UtcNow.AddSeconds($StartupTimeoutSeconds)
                $sessionId = Wait-ForHostQuiescent -SinceUtc $cycleStartedUtc -DeadlineUtc $startupDeadline
                if ($ownedProcess.HasExited) {
                    throw "Outlook process $($ownedProcess.Id) exited during smoke cycle $cycle startup."
                }
                $outlookProcesses = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
                if ($outlookProcesses.Count -ne 1 -or $outlookProcesses[0].Id -ne $ownedProcess.Id) {
                    throw "Smoke cycle $cycle does not own the sole Outlook process. Expected PID $($ownedProcess.Id); found: $($outlookProcesses.Id -join ', ')"
                }

                Wait-ForOutlookEvent45 `
                    -ProcessId $ownedProcess.Id `
                    -AfterRecordId $eventWatermark `
                    -SinceUtc $cycleStartedUtc `
                    -DeadlineUtc $startupDeadline
                Wait-ForOutlookMainWindow -Process $ownedProcess -DeadlineUtc $startupDeadline

                $endpointResult = $null
                $codexResult = $null
                if ($ExpectedPhase -ge 2) {
                    $endpointResult = & (Join-Path $PSScriptRoot 'test-phase2-endpoint.ps1') `
                        -TimeoutSeconds ([Math]::Min($StartupTimeoutSeconds, 60))
                    if ($cycle -eq 1) {
                        $codexResult = & (Join-Path $PSScriptRoot 'test-phase2-codex.ps1') `
                            -TimeoutSeconds ([Math]::Min($StartupTimeoutSeconds + 30, 180))
                    }
                }

                $normalCloseRequested = $true
                $portReleaseStopwatch = [Diagnostics.Stopwatch]::StartNew()
                if (-not $ownedProcess.CloseMainWindow()) {
                    throw "Outlook process $($ownedProcess.Id) rejected the normal CloseMainWindow request. It was not terminated."
                }
                $portReleaseMilliseconds = Wait-ForLoopbackPortRelease -Stopwatch $portReleaseStopwatch
                if (-not $ownedProcess.WaitForExit($ShutdownTimeoutSeconds * 1000)) {
                    throw "Outlook process $($ownedProcess.Id) did not exit normally within $ShutdownTimeoutSeconds seconds. It was not terminated."
                }

                $shutdownDeadline = [DateTime]::UtcNow.AddSeconds($ShutdownTimeoutSeconds)
                Wait-ForShutdownCompleted `
                    -SessionId $sessionId `
                    -SinceUtc $cycleStartedUtc `
                    -DeadlineUtc $shutdownDeadline
                Assert-OutlookStoppedAndPortFree
                Assert-ExactRegistration

                $event45 = @(Get-TargetOutlookEvents `
                    -EventId 45 `
                    -AfterRecordId $eventWatermark `
                    -SinceUtc $cycleStartedUtc `
                    -ProcessId $ownedProcess.Id)
                if ($event45.Count -ne 1) {
                    throw "Smoke cycle $cycle produced $($event45.Count) PID-correlated Event ID 45 records for $addInProgId; expected one."
                }
                if ($event45[0].LoadBehaviors.Count -ne 1 -or $event45[0].LoadBehaviors[0] -ne 3) {
                    throw "Smoke cycle $cycle Event ID 45 did not report exactly one Load Behavior 3 entry for $addInProgId."
                }

                $event59 = @(Get-TargetOutlookEvents `
                    -EventId 59 `
                    -AfterRecordId $eventWatermark `
                    -SinceUtc $cycleStartedUtc)
                if ($event59.Count -ne 0) {
                    throw "Smoke cycle $cycle recorded Outlook Event ID 59 for $addInProgId."
                }
                if ((Get-ResiliencySnapshot) -ne $initialResiliencySnapshot) {
                    throw "Outlook resilience state changed during smoke cycle $cycle. DisabledItems and CrashingAddinList were not modified by the runner."
                }

                $cycleResults.Add([pscustomobject]@{
                    Cycle = $cycle
                    ProcessId = $ownedProcess.Id
                    SessionId = $sessionId
                    Event45RecordId = $event45[0].RecordId
                    EventWatermark = $eventWatermark
                    StartedUtc = $cycleStartedUtc
                    EndpointVerified = $null -ne $endpointResult
                    CodexVerified = $null -ne $codexResult
                    PortReleaseMilliseconds = $portReleaseMilliseconds
                })
                Write-Host "Phase $ExpectedPhase smoke cycle $cycle/3 passed (PID $($ownedProcess.Id), session $sessionId)."
            }
            catch {
                if (-not $normalCloseRequested) {
                    Close-OwnedOutlookAfterFailure -Process $ownedProcess
                }
                throw
            }
            finally {
                if ($null -ne $ownedProcess) {
                    $ownedProcess.Dispose()
                }
            }
        }

        $smokeFinishedUtc = [DateTimeOffset]::UtcNow
        if (@($cycleResults | Select-Object -ExpandProperty ProcessId -Unique).Count -ne 3 -or
            @($cycleResults | Select-Object -ExpandProperty SessionId -Unique).Count -ne 3) {
            throw "The Phase $ExpectedPhase gate did not produce three distinct Outlook processes and diagnostic sessions."
        }
        if ($ExpectedPhase -ge 2 -and
            @($cycleResults | Where-Object CodexVerified).Count -ne 1) {
            throw 'The Phase 2 gate did not complete exactly one Codex outlook_status probe.'
        }

        Start-Sleep -Seconds 2
        foreach ($cycleResult in $cycleResults) {
            $delayedEvent59 = @(Get-TargetOutlookEvents `
                -EventId 59 `
                -AfterRecordId $cycleResult.EventWatermark `
                -SinceUtc $cycleResult.StartedUtc)
            if ($delayedEvent59.Count -ne 0) {
                throw "Outlook recorded a delayed Event ID 59 for $addInProgId after smoke cycle $($cycleResult.Cycle)."
            }
        }

        $changedVstoLogs = @(Get-ChangedVstoAlertLogs -Before $vstoLogAlertsBefore)
        if ($changedVstoLogs.Count -gt 0) {
            throw "VSTO loader diagnostics were written during the smoke gate: $($changedVstoLogs -join ', ')"
        }

        $verificationResult = & (Join-Path $PSScriptRoot 'verify-phase1-smoke.ps1') `
            -SinceUtc $smokeStartedUtc `
            -UntilUtc $smokeFinishedUtc `
            -ExpectedCycles 3 `
            -ExpectedPhase $ExpectedPhase
    }
    catch {
        $primaryFailure = $_
    }
    finally {
        [Environment]::SetEnvironmentVariable('VSTO_LOGALERTS', $previousVstoLogAlerts, 'Process')
        [Environment]::SetEnvironmentVariable('VSTO_SUPPRESSDISPLAYALERTS', $previousVstoSuppressAlerts, 'Process')

        $runningOutlook = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
        if ($runningOutlook.Count -gt 0 -and $registrationAttempted) {
            if ($null -ne $certificate) {
                $certificate.Dispose()
            }
            $cleanupFailures.Add(
                "Outlook is still running as PID(s) $($runningOutlook.Id -join ', '). " +
                'The runner did not unregister the loaded add-in or remove its temporary certificate; close Outlook normally before cleanup.')
        }
        else {
            if ($registrationAttempted) {
                try {
                    Invoke-VstoTarget -Target VSTOClean -CertificateThumbprint $certificateThumbprint
                }
                catch {
                    $cleanupFailures.Add("Release VSTOClean failed: $($_.Exception.Message)")
                }
            }

            if ($null -ne $certificate) {
                $certificate.Dispose()
            }
            try {
                foreach ($storedCertificate in @(
                    Get-ChildItem 'Cert:\CurrentUser\My' |
                        Where-Object Subject -eq $certificateSubject)) {
                    Remove-Item -LiteralPath $storedCertificate.PSPath -DeleteKey -Force -ErrorAction Stop
                }
            }
            catch {
                $cleanupFailures.Add("Certificate/private-key removal failed: $($_.Exception.Message)")
            }
            try {
                foreach ($authorityCertificate in @(
                    Get-ChildItem 'Cert:\CurrentUser\CA' |
                        Where-Object Subject -eq $certificateSubject)) {
                    Remove-Item -LiteralPath $authorityCertificate.PSPath -Force -ErrorAction Stop
                }
            }
            catch {
                $cleanupFailures.Add("Certificate authority-store cleanup failed: $($_.Exception.Message)")
            }
            try {
                if ([System.Security.Cryptography.CngKey]::Exists($keyContainerName, $keyProvider)) {
                    $orphanedKey = [System.Security.Cryptography.CngKey]::Open($keyContainerName, $keyProvider)
                    try {
                        if ([string]::IsNullOrWhiteSpace($keyUniqueName)) {
                            $keyUniqueName = $orphanedKey.UniqueName
                            $keyFilePath = Join-Path (
                                [Environment]::GetFolderPath([Environment+SpecialFolder]::ApplicationData)) (
                                "Microsoft\Crypto\Keys\$keyUniqueName")
                        }
                        $orphanedKey.Delete()
                    }
                    finally {
                        $orphanedKey.Dispose()
                    }
                }
            }
            catch {
                $cleanupFailures.Add("CNG key-container cleanup failed: $($_.Exception.Message)")
            }

            try {
                $certificateResidue = @(
                    Get-ChildItem 'Cert:\CurrentUser\My', 'Cert:\CurrentUser\CA' |
                        Where-Object Subject -eq $certificateSubject
                )
                if ($certificateResidue.Count -gt 0) {
                    $cleanupFailures.Add('The ephemeral smoke signing certificate remains in a current-user certificate store.')
                }
                if ([System.Security.Cryptography.CngKey]::Exists($keyContainerName, $keyProvider)) {
                    $cleanupFailures.Add("The ephemeral smoke CNG container remains: $keyContainerName")
                }
                if (-not [string]::IsNullOrWhiteSpace($keyFilePath) -and
                    (Test-Path -LiteralPath $keyFilePath)) {
                    $cleanupFailures.Add("The ephemeral smoke CNG key file remains: $keyFilePath")
                }
                $remainingVstoState = @(Get-VstoDevelopmentState)
                if ($remainingVstoState.Count -gt 0) {
                    $cleanupFailures.Add("VSTO development registration remains: $($remainingVstoState -join [Environment]::NewLine)")
                }
            }
            catch {
                $cleanupFailures.Add("Cleanup verification failed: $($_.Exception.Message)")
            }
        }
    }

    if ($null -ne $primaryFailure) {
        foreach ($cleanupFailure in $cleanupFailures) {
            Write-Warning $cleanupFailure
        }
        $PSCmdlet.ThrowTerminatingError($primaryFailure)
    }
    if ($cleanupFailures.Count -gt 0) {
        throw "Phase $ExpectedPhase smoke cleanup failed: $($cleanupFailures -join [Environment]::NewLine)"
    }

    $maximumPortReleaseMilliseconds =
        ($cycleResults | Measure-Object -Property PortReleaseMilliseconds -Maximum).Maximum
    if ($maximumPortReleaseMilliseconds -ge 3000) {
        throw "The maximum loopback port-release time was $maximumPortReleaseMilliseconds ms; expected less than 3000 ms."
    }

    [pscustomobject]@{
        Profile = $Profile
        Phase = $ExpectedPhase
        VerifiedCycles = $cycleResults.Count
        Cycles = $cycleResults.ToArray()
        SinceUtc = $smokeStartedUtc
        UntilUtc = $smokeFinishedUtc
        MaximumStartupMilliseconds = $verificationResult.MaximumStartupMilliseconds
        MaximumPortReleaseMilliseconds = $maximumPortReleaseMilliseconds
        RuntimeIdentityFingerprint = $verificationResult.RuntimeIdentityFingerprint
        Event45Records = $cycleResults.Count
        Event59Records = 0
        ResiliencyStateUnchanged = $true
        OutlookStopped = $true
        PortReleased = $true
        CodexVerified = $ExpectedPhase -ge 2
        RegistrationRemoved = $true
        TemporaryCertificateRemoved = $true
    }
}
finally {
    [Environment]::SetEnvironmentVariable('OUTLOOK_MCP_TOKEN', $previousOutlookMcpToken, 'Process')
    Pop-Location
}
