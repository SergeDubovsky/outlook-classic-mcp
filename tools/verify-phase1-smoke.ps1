#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [DateTimeOffset]$SinceUtc,

    [Parameter(Mandatory = $true)]
    [DateTimeOffset]$UntilUtc,

    [ValidateRange(1, 10)]
    [int]$ExpectedCycles = 3,

    [ValidateSet(1, 2, 3)]
    [int]$ExpectedPhase = 1,

    [string]$LogDirectory = (Join-Path $env:LOCALAPPDATA 'OutlookClassicMcp\logs')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($ExpectedPhase -eq 1) {
    throw (
        'Phase 1 verification mode is retired because the current add-in always starts the Phase 2 listener. ' +
        'Use the Phase 2 smoke wrapper; historical results remain in docs\PHASE_1_EVIDENCE.md.')
}

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$releaseManifestPath = Join-Path $repositoryRoot 'src\OutlookClassicMcp.AddIn\bin\Release\OutlookClassicMcp.AddIn.vsto'
$expectedManifest = ([Uri]::new($releaseManifestPath)).AbsoluteUri + '|vstolocal'
$verificationStartedUtc = [DateTimeOffset]::UtcNow
if ($UntilUtc -lt $SinceUtc) {
    throw 'UntilUtc must not precede SinceUtc.'
}
if ($UntilUtc -gt $verificationStartedUtc.AddSeconds(5)) {
    throw 'UntilUtc must not be in the future.'
}
if (-not (Test-Path -LiteralPath $releaseManifestPath -PathType Leaf)) {
    throw "The Release VSTO manifest does not exist: $releaseManifestPath"
}
& (Join-Path $PSScriptRoot 'verify-mcp-runtime-assets.ps1') -Configuration Release | Out-Null

$maximumLogBytes = 5L * 1024L * 1024L
$expectedFingerprint = '65F5A0725589215872E39F708ABC09C1B5E9AAABB1D1563886EFAC1A3D1FC52F'
$expectedEvents = @(
    'startup_completed'
    'dependency_binding_completed'
    'dispatcher_probe_completed'
    if ($ExpectedPhase -ge 2) {
        'listener_binding_completed'
    }
    'host_quiescent'
    'shutdown_completed'
)
$expectedFields = @(
    'schema'
    'timestamp'
    'session'
    'event'
    'state'
    'result'
    'duration_ticks'
    'stopwatch_frequency'
    'build'
    'debugger_attached'
    'queue_depth'
    'tracked_tasks'
    'listener_active'
    'mcp_core_version'
    'dependency_count'
    'dependency_identity_sha256'
    'captured_managed_thread'
    'captured_native_thread'
    'captured_apartment'
    'executed_managed_thread'
    'executed_native_thread'
    'executed_apartment'
    'exception_type'
    'hresult'
)
$failures = [System.Collections.Generic.List[string]]::new()

function Add-Failure {
    param([Parameter(Mandatory = $true)][string]$Message)

    $failures.Add($Message)
}

function Test-RestrictedAcl {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][bool]$RequireProtectedAcl,
        [Parameter(Mandatory = $true)][string[]]$AllowedSids
    )

    $acl = Get-Acl -LiteralPath $Path
    if ($RequireProtectedAcl -and -not $acl.AreAccessRulesProtected) {
        Add-Failure "$Path inherits access rules instead of using the protected diagnostics ACL."
    }

    try {
        $ownerSid = $acl.GetOwner([Security.Principal.SecurityIdentifier]).Value
    }
    catch {
        Add-Failure "The owner SID for $Path could not be resolved."
        return
    }
    if ($AllowedSids -notcontains $ownerSid) {
        Add-Failure "$Path has an unexpected owner SID: $ownerSid"
    }

    $rules = @($acl.GetAccessRules(
        $true,
        $true,
        [Security.Principal.SecurityIdentifier]))
    foreach ($rule in $rules) {
        $sid = $rule.IdentityReference.Value
        if ($rule.AccessControlType -ne [Security.AccessControl.AccessControlType]::Allow) {
            Add-Failure "$Path has an unexpected deny access rule for $sid."
        }
        elseif ($AllowedSids -notcontains $sid) {
            Add-Failure "$Path grants access to an unexpected SID: $sid"
        }
    }

    foreach ($allowedSid in $AllowedSids) {
        $fullControlRule = @($rules | Where-Object {
            $_.AccessControlType -eq [Security.AccessControl.AccessControlType]::Allow -and
            $_.IdentityReference.Value -eq $allowedSid -and
            ($_.PropagationFlags -band [Security.AccessControl.PropagationFlags]::InheritOnly) -eq 0 -and
            ($_.FileSystemRights -band [Security.AccessControl.FileSystemRights]::FullControl) -eq
                [Security.AccessControl.FileSystemRights]::FullControl
        })
        if ($fullControlRule.Count -eq 0) {
            Add-Failure "$Path does not grant FullControl to required SID $allowedSid."
        }
    }
}

function ConvertFrom-DiagnosticLine {
    param(
        [Parameter(Mandatory = $true)][string]$Line,
        [Parameter(Mandatory = $true)][string]$Source
    )

    $fields = [Collections.Generic.Dictionary[string, string]]::new(
        [StringComparer]::Ordinal)
    foreach ($part in $Line.Split("`t")) {
        $separator = $part.IndexOf('=')
        if ($separator -le 0) {
            Add-Failure "$Source contains a malformed diagnostics field."
            return $null
        }

        $key = $part.Substring(0, $separator)
        $value = $part.Substring($separator + 1)
        if ($fields.ContainsKey($key)) {
            Add-Failure "$Source contains duplicate field '$key'."
            return $null
        }
        $fields.Add($key, $value)
    }

    if ($fields.Count -ne $expectedFields.Count) {
        Add-Failure "$Source has $($fields.Count) fields; expected $($expectedFields.Count)."
        return $null
    }
    foreach ($expectedField in $expectedFields) {
        if (-not $fields.ContainsKey($expectedField)) {
            Add-Failure "$Source is missing field '$expectedField'."
            return $null
        }
    }

    [DateTimeOffset]$timestamp = [DateTimeOffset]::MinValue
    if (-not [DateTimeOffset]::TryParse(
        $fields['timestamp'],
        [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::RoundtripKind,
        [ref]$timestamp)) {
        Add-Failure "$Source has an invalid timestamp."
        return $null
    }

    [Guid]$sessionId = [Guid]::Empty
    if (-not [Guid]::TryParseExact($fields['session'], 'N', [ref]$sessionId)) {
        Add-Failure "$Source has an invalid session identifier."
        return $null
    }

    return [pscustomobject]@{
        Fields = $fields
        Timestamp = $timestamp
        Source = $Source
    }
}

function Get-SingleEvent {
    param(
        [Parameter(Mandatory = $true)][object[]]$Records,
        [Parameter(Mandatory = $true)][string]$EventName,
        [Parameter(Mandatory = $true)][string]$SessionId
    )

    $matches = @($Records | Where-Object { $_.Fields['event'] -eq $EventName })
    if ($matches.Count -ne 1) {
        Add-Failure "Session $SessionId has $($matches.Count) '$EventName' events; expected one."
        return $null
    }
    return $matches[0]
}

function Assert-Field {
    param(
        [AllowNull()][object]$Record,
        [Parameter(Mandatory = $true)][string]$Field,
        [Parameter(Mandatory = $true)][string]$Expected,
        [Parameter(Mandatory = $true)][string]$Context
    )

    if ($null -ne $Record -and $Record.Fields[$Field] -ne $Expected) {
        Add-Failure "$Context has $Field='$($Record.Fields[$Field])'; expected '$Expected'."
    }
}

if (-not (Test-Path -LiteralPath $LogDirectory -PathType Container)) {
    throw "The lifecycle diagnostics directory does not exist: $LogDirectory"
}

$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent().User
if ($null -eq $currentUser) {
    throw 'The current Windows user has no security identifier.'
}
$systemSid = [Security.Principal.SecurityIdentifier]::new(
    [Security.Principal.WellKnownSidType]::LocalSystemSid,
    $null)
$allowedSids = @($currentUser.Value, $systemSid.Value)
Test-RestrictedAcl -Path $LogDirectory -RequireProtectedAcl $true -AllowedSids $allowedSids

$logFiles = @(Get-ChildItem -LiteralPath $LogDirectory -Filter 'outlook-classic-mcp*.log' -File)
if ($logFiles.Count -eq 0 -or $logFiles.Count -gt 5) {
    Add-Failure "The diagnostics directory contains $($logFiles.Count) log files; expected one through five."
}

$records = [System.Collections.Generic.List[object]]::new()
foreach ($logFile in $logFiles) {
    if ($logFile.Length -gt $maximumLogBytes) {
        Add-Failure "$($logFile.FullName) exceeds 5 MiB."
    }
    if ($logFile.LastWriteTimeUtc -lt [DateTime]::UtcNow.AddDays(-30)) {
        Add-Failure "$($logFile.FullName) is older than the 30-day retention limit."
    }
    Test-RestrictedAcl -Path $logFile.FullName -RequireProtectedAcl $true -AllowedSids $allowedSids

    $lineNumber = 0
    foreach ($line in [IO.File]::ReadLines($logFile.FullName)) {
        $lineNumber++
        if ([string]::IsNullOrWhiteSpace($line)) {
            Add-Failure "$($logFile.FullName):$lineNumber is blank."
            continue
        }
        $source = "$($logFile.FullName):$lineNumber"
        $record = ConvertFrom-DiagnosticLine `
            -Line $line `
            -Source $source
        if ($null -eq $record) {
            continue
        }
        if ($record.Timestamp -gt $verificationStartedUtc.AddSeconds(5)) {
            Add-Failure "$source contains a future diagnostics timestamp."
            continue
        }
        if ($record.Timestamp -lt $SinceUtc.ToUniversalTime() -or
            $record.Timestamp -gt $UntilUtc.ToUniversalTime()) {
            continue
        }
        $records.Add($record)
    }
}

$smokeRecords = @($records)
$sessions = @($smokeRecords | Group-Object { $_.Fields['session'] })
if ($sessions.Count -ne $ExpectedCycles) {
    Add-Failure "Found $($sessions.Count) diagnostic sessions since $SinceUtc; expected $ExpectedCycles."
}

$startupDurations = [System.Collections.Generic.List[double]]::new()
foreach ($session in $sessions) {
    $sessionId = $session.Name
    $sessionRecords = @($session.Group | Sort-Object Timestamp)
    if ($sessionRecords.Count -ne $expectedEvents.Count) {
        Add-Failure "Session $sessionId has $($sessionRecords.Count) events; expected $($expectedEvents.Count)."
    }
    foreach ($record in $sessionRecords) {
        if ($expectedEvents -notcontains $record.Fields['event']) {
            Add-Failure "Session $sessionId contains unexpected event '$($record.Fields['event'])'."
        }
        Assert-Field $record 'schema' '1' "Session $sessionId event $($record.Fields['event'])"
        Assert-Field $record 'result' 'success' "Session $sessionId event $($record.Fields['event'])"
        $expectedListener = if ($ExpectedPhase -ge 2 -and
            $record.Fields['event'] -in @('listener_binding_completed', 'host_quiescent')) {
            'true'
        }
        else {
            'false'
        }
        Assert-Field $record 'listener_active' $expectedListener "Session $sessionId event $($record.Fields['event'])"
        Assert-Field $record 'exception_type' '-' "Session $sessionId event $($record.Fields['event'])"
        Assert-Field $record 'hresult' '-' "Session $sessionId event $($record.Fields['event'])"
    }

    $startup = Get-SingleEvent $sessionRecords 'startup_completed' $sessionId
    $dependency = Get-SingleEvent $sessionRecords 'dependency_binding_completed' $sessionId
    $dispatcher = Get-SingleEvent $sessionRecords 'dispatcher_probe_completed' $sessionId
    $listener = if ($ExpectedPhase -ge 2) {
        Get-SingleEvent $sessionRecords 'listener_binding_completed' $sessionId
    }
    else {
        $null
    }
    $quiescent = Get-SingleEvent $sessionRecords 'host_quiescent' $sessionId
    $shutdown = Get-SingleEvent $sessionRecords 'shutdown_completed' $sessionId

    Assert-Field $startup 'state' 'starting' "Session $sessionId startup"
    Assert-Field $startup 'build' 'Release' "Session $sessionId startup"
    Assert-Field $startup 'debugger_attached' 'false' "Session $sessionId startup"
    Assert-Field $dependency 'state' 'starting' "Session $sessionId dependency binding"
    Assert-Field $dependency 'mcp_core_version' '1.4.1.0' "Session $sessionId dependency binding"
    Assert-Field $dependency 'dependency_count' '17' "Session $sessionId dependency binding"
    Assert-Field $dependency 'dependency_identity_sha256' $expectedFingerprint "Session $sessionId dependency binding"
    Assert-Field $dispatcher 'state' 'starting' "Session $sessionId dispatcher probe"
    Assert-Field $dispatcher 'queue_depth' '0' "Session $sessionId dispatcher probe"
    Assert-Field $dispatcher 'captured_apartment' 'STA' "Session $sessionId dispatcher probe"
    Assert-Field $dispatcher 'executed_apartment' 'STA' "Session $sessionId dispatcher probe"
    if ($ExpectedPhase -ge 2) {
        Assert-Field $listener 'state' 'online' "Session $sessionId listener binding"
        Assert-Field $listener 'queue_depth' '0' "Session $sessionId listener binding"
        Assert-Field $listener 'listener_active' 'true' "Session $sessionId listener binding"
    }
    if ($null -ne $dispatcher) {
        [int]$capturedManagedThread = 0
        [int]$executedManagedThread = 0
        [uint32]$capturedNativeThread = 0
        [uint32]$executedNativeThread = 0
        $validThreadMetadata =
            [int]::TryParse($dispatcher.Fields['captured_managed_thread'], [ref]$capturedManagedThread) -and
            [int]::TryParse($dispatcher.Fields['executed_managed_thread'], [ref]$executedManagedThread) -and
            [uint32]::TryParse($dispatcher.Fields['captured_native_thread'], [ref]$capturedNativeThread) -and
            [uint32]::TryParse($dispatcher.Fields['executed_native_thread'], [ref]$executedNativeThread) -and
            $capturedManagedThread -gt 0 -and $executedManagedThread -gt 0 -and
            $capturedNativeThread -gt 0 -and $executedNativeThread -gt 0
        if (-not $validThreadMetadata) {
            Add-Failure "Session $sessionId dispatcher probe has invalid thread identifiers."
        }
        elseif ($capturedManagedThread -ne $executedManagedThread -or
            $capturedNativeThread -ne $executedNativeThread) {
            Add-Failure "Session $sessionId dispatcher probe did not return to the captured Outlook thread."
        }
    }
    Assert-Field $quiescent 'state' 'online' "Session $sessionId quiescent host"
    Assert-Field $quiescent 'queue_depth' '0' "Session $sessionId quiescent host"
    Assert-Field $quiescent 'tracked_tasks' '0' "Session $sessionId quiescent host"
    Assert-Field $shutdown 'state' 'stopped' "Session $sessionId shutdown"
    Assert-Field $shutdown 'queue_depth' '0' "Session $sessionId shutdown"
    Assert-Field $shutdown 'tracked_tasks' '0' "Session $sessionId shutdown"
    Assert-Field $shutdown 'listener_active' 'false' "Session $sessionId shutdown"

    if ($null -ne $startup) {
        [long]$durationTicks = 0
        [long]$frequency = 0
        if (-not [long]::TryParse($startup.Fields['duration_ticks'], [ref]$durationTicks) -or
            -not [long]::TryParse($startup.Fields['stopwatch_frequency'], [ref]$frequency) -or
            $durationTicks -lt 0 -or $frequency -le 0) {
            Add-Failure "Session $sessionId has invalid startup timing metadata."
        }
        else {
            $durationMilliseconds = 1000.0 * $durationTicks / $frequency
            $startupDurations.Add($durationMilliseconds)
            if ($durationMilliseconds -ge 500.0) {
                Add-Failure "Session $sessionId startup callback took $durationMilliseconds ms."
            }
        }
    }

}

$outlookProcesses = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
if ($outlookProcesses.Count -ne 0) {
    Add-Failure "Outlook is still running after the smoke cycles: $($outlookProcesses.Id -join ', ')"
}
$listeners = @(Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue)
if ($listeners.Count -ne 0) {
    Add-Failure "TCP port 8765 is still listening after the smoke cycles."
}

$addInSubKey = 'Software\Microsoft\Office\Outlook\Addins\OutlookClassicMcp.AddIn'
$registryBaseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
    [Microsoft.Win32.RegistryHive]::CurrentUser,
    [Microsoft.Win32.RegistryView]::Registry64)
try {
    $addInKey = $registryBaseKey.OpenSubKey($addInSubKey)
    if ($null -eq $addInKey) {
        Add-Failure "The 64-bit Outlook development registration is missing: HKCU:\$addInSubKey"
    }
    else {
        try {
            $valueNames = @($addInKey.GetValueNames())
            if ($valueNames -notcontains 'LoadBehavior') {
                Add-Failure 'The add-in registration has no LoadBehavior value.'
            }
            else {
                $loadBehaviorKind = $addInKey.GetValueKind('LoadBehavior')
                $loadBehavior = $addInKey.GetValue('LoadBehavior')
                if ($loadBehaviorKind -ne [Microsoft.Win32.RegistryValueKind]::DWord -or
                    $loadBehavior -ne 3) {
                    Add-Failure (
                        "The add-in LoadBehavior is '$loadBehavior' ($loadBehaviorKind); " +
                        'expected 3 (DWord).')
                }
            }

            if ($valueNames -notcontains 'Manifest') {
                Add-Failure 'The add-in registration has no Manifest value.'
            }
            else {
                $manifestKind = $addInKey.GetValueKind('Manifest')
                $registeredManifest = "$($addInKey.GetValue('Manifest'))"
                if ($manifestKind -ne [Microsoft.Win32.RegistryValueKind]::String -or
                    -not [string]::Equals(
                        $registeredManifest,
                        $expectedManifest,
                        [StringComparison]::OrdinalIgnoreCase)) {
                    Add-Failure (
                        "The add-in registration references '$registeredManifest' ($manifestKind); " +
                        "expected '$expectedManifest' (String).")
                }
            }
        }
        finally {
            $addInKey.Dispose()
        }
    }
}
finally {
    $registryBaseKey.Dispose()
}

if ($failures.Count -gt 0) {
    throw "Phase $ExpectedPhase smoke verification failed:`n - $($failures -join "`n - ")"
}

if ($ExpectedPhase -eq 3) {
    [pscustomobject]@{
        Phase = 3
        VerifiedCycleCount = $sessions.Count
        ExpectedEventCountPerCycle = $expectedEvents.Count
        VerifiedEventCount = $smokeRecords.Count
        StartupCallbackUnderLimitCount = $startupDurations.Count
        DispatcherThreadMatchCount = $sessions.Count
        RuntimeIdentitySha256 = $expectedFingerprint.ToLowerInvariant()
        DiagnosticsLogFileCount = $logFiles.Count
        DiagnosticsAclVerified = $true
        OutlookStopped = $true
        PortReleased = $true
        LoadBehaviorVerified = $true
    }
    return
}

[pscustomobject]@{
    SinceUtc = $SinceUtc.ToUniversalTime()
    UntilUtc = $UntilUtc.ToUniversalTime()
    VerifiedCycles = $sessions.Count
    ExpectedPhase = $ExpectedPhase
    MaximumStartupMilliseconds = ($startupDurations | Measure-Object -Maximum).Maximum
    RuntimeIdentityFingerprint = $expectedFingerprint
    LogFiles = $logFiles.Count
    DiagnosticsAclVerified = $true
    OutlookStopped = $true
    PortReleased = $true
    LoadBehavior = 3
}
