#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('Seed', 'Detach')]
    [string]$Action = 'Seed',

    [ValidatePattern('^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$')]
    [string]$Profile = 'OutlookMcpTest',

    [string]$RunDirectory,

    [ValidateRange(60, 3600)]
    [int]$SeederTimeoutSeconds = 900,

    [ValidateRange(15, 180)]
    [int]$StartupTimeoutSeconds = 60,

    [ValidateRange(15, 180)]
    [int]$ShutdownTimeoutSeconds = 60
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$addInProjectPath = Join-Path $repositoryRoot 'src\OutlookClassicMcp.AddIn\OutlookClassicMcp.AddIn.csproj'
$addInStartupPath = Join-Path $repositoryRoot 'src\OutlookClassicMcp.AddIn\ThisAddIn.cs'
$seederSourcePath = Join-Path $repositoryRoot 'src\OutlookClassicMcp.AddIn\Smoke\Phase4FixtureSeeder.cs'
$releaseManifestPath = Join-Path $repositoryRoot 'src\OutlookClassicMcp.AddIn\bin\Release\OutlookClassicMcp.AddIn.vsto'
$expectedManifest = ([Uri]::new($releaseManifestPath)).AbsoluteUri + '|vstolocal'
$outlookPath = Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\OUTLOOK.EXE'
$vswherePath = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$addInProgId = 'OutlookClassicMcp.AddIn'
$addInSubKey = "Software\Microsoft\Office\Outlook\Addins\$addInProgId"
$stateFileName = 'seeder-state.local.json'
$fixtureFileName = 'read-fixture.local.json'
$inventoryFileName = 'store-inventory.local.json'
$seedStatusFileName = 'seed-status.local.json'
$detachStatusFileName = 'detach-status.local.json'
$progressFileName = 'seeder-progress.local.json'
$pstAFileName = 'phase4-fixture-a.pst'
$pstBFileName = 'phase4-fixture-b.pst'
$requiredComponents = @(
    'Microsoft.VisualStudio.Workload.Office'
    'Microsoft.VisualStudio.Component.TeamOffice'
    'Microsoft.VisualStudio.Workload.ManagedDesktop'
    'Microsoft.Net.Component.4.8.TargetingPack'
    'Microsoft.Net.Component.4.8.SDK'
)
$seederEnvironmentNames = @(
    'OUTLOOK_MCP_SMOKE_SEEDER_ACTION'
    'OUTLOOK_MCP_SMOKE_SEEDER_RUN_ID'
    'OUTLOOK_MCP_SMOKE_SEEDER_RUN_DIRECTORY'
    'OUTLOOK_MCP_SMOKE_SEEDER_EXPECTED_PROFILE'
    'OUTLOOK_MCP_SMOKE_SEEDER_PST_A'
    'OUTLOOK_MCP_SMOKE_SEEDER_PST_B'
    'OUTLOOK_MCP_SMOKE_SEEDER_STATUS_PATH'
)

function Get-CurrentUserSid {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    try {
        return $identity.User
    }
    finally {
        $identity.Dispose()
    }
}

function Set-SecureDirectoryAcl {
    param([Parameter(Mandatory = $true)][string]$Path)

    $userSid = Get-CurrentUserSid
    $systemSid = [Security.Principal.SecurityIdentifier]::new(
        [Security.Principal.WellKnownSidType]::LocalSystemSid,
        $null)
    $security = [Security.AccessControl.DirectorySecurity]::new()
    $security.SetOwner($userSid)
    $security.SetAccessRuleProtection($true, $false)
    $inheritance = [Security.AccessControl.InheritanceFlags]::ContainerInherit -bor
        [Security.AccessControl.InheritanceFlags]::ObjectInherit
    foreach ($sid in @($userSid, $systemSid)) {
        $security.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new(
            $sid,
            [Security.AccessControl.FileSystemRights]::FullControl,
            $inheritance,
            [Security.AccessControl.PropagationFlags]::None,
            [Security.AccessControl.AccessControlType]::Allow))
    }
    Set-Acl -LiteralPath $Path -AclObject $security
}

function Set-SecureFileAcl {
    param([Parameter(Mandatory = $true)][string]$Path)

    $userSid = Get-CurrentUserSid
    $systemSid = [Security.Principal.SecurityIdentifier]::new(
        [Security.Principal.WellKnownSidType]::LocalSystemSid,
        $null)
    $security = [Security.AccessControl.FileSecurity]::new()
    $security.SetOwner($userSid)
    $security.SetAccessRuleProtection($true, $false)
    foreach ($sid in @($userSid, $systemSid)) {
        $security.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new(
            $sid,
            [Security.AccessControl.FileSystemRights]::FullControl,
            [Security.AccessControl.AccessControlType]::Allow))
    }
    $aclExtensions = 'System.IO.FileSystemAclExtensions' -as [type]
    if ($null -ne $aclExtensions) {
        [IO.FileSystemAclExtensions]::SetAccessControl([IO.FileInfo]::new($Path), $security)
    }
    else {
        [IO.File]::SetAccessControl($Path, $security)
    }
}

function Assert-SecurePathAcl {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][ValidateSet('Directory', 'File')][string]$PathType
    )

    if ($PathType -eq 'Directory') {
        $security = Get-Acl -LiteralPath $Path
    }
    else {
        $security = Get-Acl -LiteralPath $Path
    }

    if (-not $security.AreAccessRulesProtected) {
        throw "The secure $PathType path inherits ACL entries."
    }

    $userSid = Get-CurrentUserSid
    $systemSid = [Security.Principal.SecurityIdentifier]::new(
        [Security.Principal.WellKnownSidType]::LocalSystemSid,
        $null)
    $owner = $security.GetOwner([Security.Principal.SecurityIdentifier])
    if (-not $owner.Equals($userSid)) {
        throw "The secure $PathType path is not owned by the current user."
    }

    $seenUser = $false
    $seenSystem = $false
    $rules = $security.GetAccessRules(
        $true,
        $true,
        [Security.Principal.SecurityIdentifier])
    foreach ($rule in $rules) {
        if ($rule.IsInherited -or
            $rule.AccessControlType -ne [Security.AccessControl.AccessControlType]::Allow -or
            (($rule.FileSystemRights -band [Security.AccessControl.FileSystemRights]::FullControl) -ne
                [Security.AccessControl.FileSystemRights]::FullControl)) {
            throw "The secure $PathType path has an unsupported ACL entry."
        }

        if ($rule.IdentityReference.Equals($userSid)) {
            $seenUser = $true
        }
        elseif ($rule.IdentityReference.Equals($systemSid)) {
            $seenSystem = $true
        }
        else {
            throw "The secure $PathType path grants access to an unsupported identity."
        }
    }

    if (-not $seenUser -or -not $seenSystem) {
        throw "The secure $PathType path is missing a required ACL entry."
    }
}

function Write-AtomicSecureJson {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object]$Value
    )

    $parent = Split-Path -Parent $Path
    $temporaryPath = Join-Path $parent ('.' + [IO.Path]::GetFileName($Path) + '.' +
        [Guid]::NewGuid().ToString('N') + '.tmp')
    try {
        $json = $Value | ConvertTo-Json -Depth 8
        [IO.File]::WriteAllText(
            $temporaryPath,
            $json,
            [Text.UTF8Encoding]::new($false))
        Set-SecureFileAcl -Path $temporaryPath
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            [IO.File]::Replace($temporaryPath, $Path, $null, $true)
        }
        else {
            [IO.File]::Move($temporaryPath, $Path)
        }
        Set-SecureFileAcl -Path $Path
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
            Remove-Item -LiteralPath $temporaryPath -Force
        }
    }
}

function Assert-LocalNonReparsePath {
    param([Parameter(Mandatory = $true)][string]$CanonicalPath)

    if ($CanonicalPath -match '^(\\\\|//)') {
        throw 'The Phase 4 run directory must be on a local fixed drive, not a UNC or device path.'
    }

    $driveRoot = [IO.Path]::GetPathRoot($CanonicalPath)
    if ([string]::IsNullOrWhiteSpace($driveRoot) -or
        $driveRoot -notmatch '^[A-Za-z]:\\$') {
        throw 'The Phase 4 run directory must be on a local fixed drive.'
    }
    try {
        $drive = [IO.DriveInfo]::new($driveRoot)
    }
    catch {
        throw 'The Phase 4 run directory drive could not be validated.'
    }
    if (-not $drive.IsReady -or $drive.DriveType -ne [IO.DriveType]::Fixed) {
        throw 'The Phase 4 run directory must be on a ready local fixed drive.'
    }

    $relativePath = $CanonicalPath.Substring($driveRoot.Length).TrimStart('\')
    $current = $driveRoot
    foreach ($segment in @($relativePath -split '\\' | Where-Object { $_.Length -gt 0 })) {
        $current = Join-Path $current $segment
        if (-not (Test-Path -LiteralPath $current)) {
            break
        }
        $item = Get-Item -LiteralPath $current -Force
        $actual = [IO.Path]::GetFullPath($item.FullName).TrimEnd('\')
        $expected = [IO.Path]::GetFullPath($current).TrimEnd('\')
        if (-not [string]::Equals($actual, $expected, [StringComparison]::OrdinalIgnoreCase)) {
            throw 'The Phase 4 run directory path did not retain its expected local identity.'
        }
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw 'The Phase 4 run directory cannot contain a reparse point.'
        }
    }
}

function Get-CanonicalRunDirectory {
    param(
        [AllowNull()][string]$RequestedPath,
        [Parameter(Mandatory = $true)][bool]$Create
    )

    if ([string]::IsNullOrWhiteSpace($RequestedPath)) {
        if (-not $Create) {
            throw '-RunDirectory is required for Detach.'
        }
        $basePath = Join-Path $env:LOCALAPPDATA 'OutlookClassicMcp\Smoke\Phase4'
        $RequestedPath = Join-Path $basePath ([Guid]::NewGuid().ToString('N'))
    }

    if (-not [IO.Path]::IsPathRooted($RequestedPath) -or $RequestedPath.IndexOf([char]0) -ge 0) {
        throw 'The Phase 4 run directory must be an absolute path.'
    }
    $canonical = [IO.Path]::GetFullPath($RequestedPath).TrimEnd('\')
    Assert-LocalNonReparsePath -CanonicalPath $canonical
    $root = [IO.Path]::GetPathRoot($canonical).TrimEnd('\')
    if ([string]::Equals($canonical, $root, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'The Phase 4 run directory cannot be a filesystem root.'
    }

    if ($Create -and -not (Test-Path -LiteralPath $canonical)) {
        $null = New-Item -ItemType Directory -Path $canonical -Force
        Set-SecureDirectoryAcl -Path $canonical
    }
    if (-not (Test-Path -LiteralPath $canonical -PathType Container)) {
        throw 'The Phase 4 run directory does not exist.'
    }

    Assert-LocalNonReparsePath -CanonicalPath $canonical
    $item = Get-Item -LiteralPath $canonical -Force
    if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw 'The Phase 4 run directory cannot be a reparse point.'
    }
    Assert-SecurePathAcl -Path $canonical -PathType Directory
    return $canonical
}

function Get-CanonicalChildPath {
    param(
        [Parameter(Mandatory = $true)][string]$Parent,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $candidate = [IO.Path]::GetFullPath((Join-Path $Parent $Name))
    $prefix = $Parent.TrimEnd('\') + '\'
    if (-not $candidate.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'A generated child path escaped the secure run directory.'
    }
    return $candidate
}

function Open-RegistryKey {
    param(
        [Parameter(Mandatory = $true)][Microsoft.Win32.RegistryHive]$Hive,
        [Parameter(Mandatory = $true)][Microsoft.Win32.RegistryView]$View,
        [Parameter(Mandatory = $true)][string]$SubKey
    )

    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, $View)
    try {
        return $baseKey.OpenSubKey($SubKey)
    }
    finally {
        $baseKey.Dispose()
    }
}

function Get-ProfileSubKey {
    param([Parameter(Mandatory = $true)][string]$ProfileName)

    return "Software\Microsoft\Office\16.0\Outlook\Profiles\$ProfileName"
}

function Test-ProfileExists {
    param([Parameter(Mandatory = $true)][string]$ProfileName)

    $key = Open-RegistryKey `
        -Hive ([Microsoft.Win32.RegistryHive]::CurrentUser) `
        -View ([Microsoft.Win32.RegistryView]::Registry64) `
        -SubKey (Get-ProfileSubKey -ProfileName $ProfileName)
    if ($null -eq $key) {
        return $false
    }
    $key.Dispose()
    return $true
}

function Get-RegistrationState {
    $state = [Collections.Generic.List[string]]::new()
    foreach ($view in @(
        [Microsoft.Win32.RegistryView]::Registry32,
        [Microsoft.Win32.RegistryView]::Registry64)) {
        $key = Open-RegistryKey `
            -Hive ([Microsoft.Win32.RegistryHive]::CurrentUser) `
            -View $view `
            -SubKey $addInSubKey
        if ($null -ne $key) {
            try {
                $state.Add("$view Outlook add-in registration")
            }
            finally {
                $key.Dispose()
            }
        }
    }
    return $state.ToArray()
}

function Assert-ExactRegistration {
    $key = Open-RegistryKey `
        -Hive ([Microsoft.Win32.RegistryHive]::CurrentUser) `
        -View ([Microsoft.Win32.RegistryView]::Registry64) `
        -SubKey $addInSubKey
    if ($null -eq $key) {
        throw 'The smoke VSTO build did not create the expected Outlook add-in registration.'
    }
    try {
        $manifest = [string]$key.GetValue('Manifest', $null)
        $loadBehavior = $key.GetValue('LoadBehavior', $null)
        if (-not [string]::Equals($manifest, $expectedManifest, [StringComparison]::Ordinal) -or
            $null -eq $loadBehavior -or [int]$loadBehavior -ne 3) {
            throw 'The smoke VSTO registration does not match the exact Release manifest.'
        }
    }
    finally {
        $key.Dispose()
    }
}

function Assert-OutlookStopped {
    $processes = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
    if ($processes.Count -gt 0) {
        throw "Outlook is running. Close it normally first; process IDs: $($processes.Id -join ', ')."
    }
}

function Close-OutlookNormally {
    param(
        [Parameter(Mandatory = $true)][Diagnostics.Process]$Process,
        [Parameter(Mandatory = $true)][string]$Context
    )

    if ($Process.HasExited) {
        throw "$Context exited before the runner issued a normal close request."
    }

    $deadline = [DateTime]::UtcNow.AddSeconds($StartupTimeoutSeconds)
    while (-not $Process.HasExited -and
        $Process.MainWindowHandle -eq [IntPtr]::Zero -and
        [DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 250
        $Process.Refresh()
    }
    if ($Process.HasExited) {
        throw "$Context exited before the runner issued a normal close request."
    }
    if ($Process.MainWindowHandle -eq [IntPtr]::Zero -or -not $Process.CloseMainWindow()) {
        throw "$Context did not expose or accept a normal CloseMainWindow request. It was not terminated."
    }
    if (-not $Process.WaitForExit($ShutdownTimeoutSeconds * 1000)) {
        throw "$Context did not exit after a normal close request. It was not terminated."
    }
}

function Start-ProfileBootstrap {
    param([Parameter(Mandatory = $true)][string]$ProfileName)

    Assert-OutlookStopped
    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $outlookPath
    $startInfo.Arguments = "/PIM `"$ProfileName`""
    $startInfo.WorkingDirectory = Split-Path -Parent $outlookPath
    $startInfo.UseShellExecute = $false
    $process = [Diagnostics.Process]::Start($startInfo)
    if ($null -eq $process) {
        throw 'Classic Outlook did not start for accountless profile creation.'
    }
    $bootstrapFailure = $null
    $normalCloseFailure = $null
    try {
        $deadline = [DateTime]::UtcNow.AddSeconds($StartupTimeoutSeconds)
        while (-not (Test-ProfileExists -ProfileName $ProfileName) -and
            -not $process.HasExited -and
            [DateTime]::UtcNow -lt $deadline) {
            Start-Sleep -Milliseconds 250
            $process.Refresh()
        }
        if (-not (Test-ProfileExists -ProfileName $ProfileName)) {
            throw 'Classic Outlook did not create the requested accountless /PIM profile.'
        }
    }
    catch {
        $bootstrapFailure = $_
    }
    finally {
        try {
            Close-OutlookNormally -Process $process -Context 'Accountless profile bootstrap Outlook'
        }
        catch {
            $normalCloseFailure = $_
        }
        finally {
            $process.Dispose()
        }
    }
    if ($null -ne $bootstrapFailure) {
        if ($null -ne $normalCloseFailure) {
            $bootstrapFailure.Exception.Data['NormalCloseFailure'] =
                $normalCloseFailure.Exception.Message
        }
        throw $bootstrapFailure
    }
    if ($null -ne $normalCloseFailure) {
        throw $normalCloseFailure
    }
}

function Get-FullVisualStudioMsBuild {
    if (-not (Test-Path -LiteralPath $vswherePath -PathType Leaf)) {
        throw "vswhere was not found at $vswherePath"
    }
    $arguments = @('-products', '*', '-version', '[18.0,19.0)', '-requires') +
        $requiredComponents + @('-format', 'json', '-utf8')
    $json = @(& $vswherePath @arguments)
    $instances = @((($json -join [Environment]::NewLine) | ConvertFrom-Json))
    $instance = $instances |
        Where-Object {
            $_.PSObject.Properties.Name -contains 'productId' -and
            $_.productId -ne 'Microsoft.VisualStudio.Product.BuildTools'
        } |
        Select-Object -First 1
    if ($null -eq $instance) {
        throw 'A full Visual Studio 2026 installation with Office/VSTO and .NET Framework 4.8 components was not found.'
    }
    $msbuild = Join-Path $instance.installationPath 'MSBuild\Current\Bin\MSBuild.exe'
    if (-not (Test-Path -LiteralPath $msbuild -PathType Leaf)) {
        throw "Full Visual Studio MSBuild was not found at $msbuild"
    }
    return $msbuild
}

function Invoke-VstoTarget {
    param(
        [Parameter(Mandatory = $true)][string]$MsBuild,
        [Parameter(Mandatory = $true)][ValidateSet('Build', 'VSTOClean')][string]$Target,
        [Parameter(Mandatory = $true)][string]$CertificateThumbprint
    )

    $arguments = @(
        $addInProjectPath
        "/t:$Target"
        '/p:Configuration=Release'
        '/p:Platform=AnyCPU'
        '/p:VisualStudioVersion=18.0'
        '/p:OutlookMcpSmokeSeeder=true'
        '/p:SignManifests=true'
        "/p:ManifestCertificateThumbprint=$CertificateThumbprint"
        '/p:RestoreLockedMode=true'
        '/v:minimal'
        '/nologo'
        '/nr:false'
    )
    & $MsBuild @arguments | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Release VSTO $Target failed with exit code $LASTEXITCODE."
    }
}

function Assert-SeederIntegration {
    foreach ($path in @($addInProjectPath, $addInStartupPath, $seederSourcePath)) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Required seeder integration file is missing: $path"
        }
    }
    $project = Get-Content -LiteralPath $addInProjectPath -Raw
    $startup = Get-Content -LiteralPath $addInStartupPath -Raw
    if ($project -notmatch 'OutlookMcpSmokeSeeder' -or
        $project -notmatch 'OUTLOOK_MCP_SMOKE_SEEDER' -or
        $project -notmatch 'Smoke\\Phase4FixtureSeeder\.cs' -or
        $startup -notmatch 'Phase4FixtureSeeder\.TryStart') {
        throw 'The compile-conditional VSTO seeder integration has not been added to the project and ThisAddIn startup.'
    }
}

function Test-TransientAtomicFileError {
    param([Parameter(Mandatory = $true)][Exception]$Exception)

    $current = $Exception
    while ($null -ne $current) {
        if ($current -is [IO.FileNotFoundException] -or
            $current -is [IO.DirectoryNotFoundException] -or
            $current -is [Management.Automation.ItemNotFoundException]) {
            return $true
        }
        if ($current -is [IO.IOException]) {
            $nativeCode = $current.HResult -band 0xffff
            if ($nativeCode -in @(2, 3, 32, 33)) {
                return $true
            }
        }
        $current = $current.InnerException
    }
    return $false
}

function Read-SecureAtomicJson {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $null
    }
    $stream = $null
    $reader = $null
    $json = $null
    try {
        Assert-SecurePathAcl -Path $Path -PathType File
        $stream = [IO.FileStream]::new(
            $Path,
            [IO.FileMode]::Open,
            [IO.FileAccess]::Read,
            ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
        $reader = [IO.StreamReader]::new(
            $stream,
            [Text.UTF8Encoding]::new($false),
            $true)
        $json = $reader.ReadToEnd()
    }
    catch {
        if (Test-TransientAtomicFileError -Exception $_.Exception) {
            return $null
        }
        throw
    }
    finally {
        if ($null -ne $reader) {
            $reader.Dispose()
        }
        elseif ($null -ne $stream) {
            $stream.Dispose()
        }
    }
    $value = $json | ConvertFrom-Json -ErrorAction Stop
    if ($null -eq $value) {
        throw "The atomic JSON document is invalid: $([IO.Path]::GetFileName($Path))"
    }
    return $value
}

function Get-CurrentRunSeederProgress {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ExpectedAction,
        [Parameter(Mandatory = $true)][string]$ExpectedRunId
    )

    $progress = Read-SecureAtomicJson -Path $Path
    if ($null -eq $progress) {
        return $null
    }
    try {
        $completedCount = [Convert]::ToInt32($progress.completedCount)
        $targetCount = [Convert]::ToInt32($progress.targetCount)
    }
    catch {
        throw 'The VSTO seeder returned an invalid progress document.'
    }
    if ($progress.schema -ne 1 -or
        $progress.runId -cne $ExpectedRunId -or
        $progress.action -cne $ExpectedAction.ToLowerInvariant() -or
        $progress.stage -isnot [string] -or
        $completedCount -lt 0 -or
        $targetCount -lt 0) {
        throw 'The VSTO seeder returned an invalid progress document.'
    }
    return [pscustomobject]@{
        Stage = $progress.stage
        CompletedCount = $completedCount
        TargetCount = $targetCount
    }
}

function Wait-SeederStatus {
    param(
        [Parameter(Mandatory = $true)][Diagnostics.Process]$Process,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ProgressPath,
        [Parameter(Mandatory = $true)][string]$ExpectedAction,
        [Parameter(Mandatory = $true)][string]$ExpectedRunId
    )

    $startupDeadline = [DateTime]::UtcNow.AddSeconds($StartupTimeoutSeconds)
    $deadline = [DateTime]::UtcNow.AddSeconds($SeederTimeoutSeconds)
    $sawCurrentRunSignal = $false
    while ([DateTime]::UtcNow -lt $deadline) {
        $status = Read-SecureAtomicJson -Path $Path
        if ($null -ne $status) {
            if ($status.schema -ne 1 -or
                $status.runId -cne $ExpectedRunId -or
                $status.action -cne $ExpectedAction.ToLowerInvariant() -or
                $status.success -isnot [bool]) {
                throw 'The VSTO seeder returned an invalid status document.'
            }
            if (-not $status.success) {
                $stage = if ($status.PSObject.Properties.Name -contains 'stage') {
                    " at stage '$($status.stage)'"
                }
                else {
                    ''
                }
                throw "The VSTO seeder failed with aggregate code '$($status.errorCode)'$stage."
            }
            return $status
        }
        $progress = Get-CurrentRunSeederProgress -Path $ProgressPath -ExpectedAction $ExpectedAction -ExpectedRunId $ExpectedRunId
        if ($null -ne $progress) {
            $sawCurrentRunSignal = $true
        }
        if ($Process.HasExited) {
            throw 'Classic Outlook exited before the VSTO seeder wrote status.'
        }
        if (-not $sawCurrentRunSignal -and [DateTime]::UtcNow -ge $startupDeadline) {
            throw "The VSTO seeder did not write a current-run progress or status signal within $StartupTimeoutSeconds seconds."
        }
        Start-Sleep -Milliseconds 250
        $Process.Refresh()
    }
    $progressSummary = ''
    $progress = Get-CurrentRunSeederProgress -Path $ProgressPath -ExpectedAction $ExpectedAction -ExpectedRunId $ExpectedRunId
    if ($null -ne $progress) {
        $progressSummary = " Last progress: stage '$($progress.Stage)', $($progress.CompletedCount)/$($progress.TargetCount)."
    }
    throw "The VSTO seeder did not finish within $SeederTimeoutSeconds seconds.$progressSummary"
}

function Invoke-SeederOutlook {
    param(
        [Parameter(Mandatory = $true)][string]$RunId,
        [Parameter(Mandatory = $true)][string]$SecureRunDirectory,
        [Parameter(Mandatory = $true)][string]$PstAPath,
        [Parameter(Mandatory = $true)][string]$PstBPath,
        [Parameter(Mandatory = $true)][string]$StatusPath,
        [Parameter(Mandatory = $true)][string]$ProgressPath
    )

    Assert-OutlookStopped
    if (Test-Path -LiteralPath $StatusPath -PathType Leaf) {
        Remove-Item -LiteralPath $StatusPath -Force
    }
    if (Test-Path -LiteralPath $ProgressPath -PathType Leaf) {
        Remove-Item -LiteralPath $ProgressPath -Force
    }
    $values = @{
        OUTLOOK_MCP_SMOKE_SEEDER_ACTION = $Action.ToLowerInvariant()
        OUTLOOK_MCP_SMOKE_SEEDER_RUN_ID = $RunId
        OUTLOOK_MCP_SMOKE_SEEDER_RUN_DIRECTORY = $SecureRunDirectory
        OUTLOOK_MCP_SMOKE_SEEDER_EXPECTED_PROFILE = $Profile
        OUTLOOK_MCP_SMOKE_SEEDER_PST_A = $PstAPath
        OUTLOOK_MCP_SMOKE_SEEDER_PST_B = $PstBPath
        OUTLOOK_MCP_SMOKE_SEEDER_STATUS_PATH = $StatusPath
    }
    foreach ($name in $values.Keys) {
        [Environment]::SetEnvironmentVariable($name, $values[$name], 'Process')
    }

    $process = $null
    $seederFailure = $null
    $normalCloseFailure = $null
    $status = $null
    try {
        $startInfo = [Diagnostics.ProcessStartInfo]::new()
        $startInfo.FileName = $outlookPath
        $startInfo.Arguments = "/profile `"$Profile`""
        $startInfo.WorkingDirectory = Split-Path -Parent $outlookPath
        $startInfo.UseShellExecute = $false
        $process = [Diagnostics.Process]::Start($startInfo)
        if ($null -eq $process) {
            throw 'Classic Outlook did not start for the VSTO seeder.'
        }
        $status = Wait-SeederStatus `
            -Process $process `
            -Path $StatusPath `
            -ProgressPath $ProgressPath `
            -ExpectedAction $Action `
            -ExpectedRunId $RunId
    }
    catch {
        $seederFailure = $_
    }
    finally {
        foreach ($name in $seederEnvironmentNames) {
            [Environment]::SetEnvironmentVariable($name, $null, 'Process')
        }
        if ($null -ne $process) {
            try {
                Close-OutlookNormally -Process $process -Context 'Seeder Outlook'
            }
            catch {
                $normalCloseFailure = $_
            }
            finally {
                $process.Dispose()
            }
        }
    }
    if ($null -ne $seederFailure) {
        if ($null -ne $normalCloseFailure) {
            $seederFailure.Exception.Data['NormalCloseFailure'] =
                $normalCloseFailure.Exception.Message
        }
        throw $seederFailure
    }
    if ($null -ne $normalCloseFailure) {
        throw $normalCloseFailure
    }
    return $status
}

function Remove-ExactRunArtifacts {
    param(
        [Parameter(Mandatory = $true)][string]$SecureRunDirectory,
        [Parameter(Mandatory = $true)][string[]]$Paths,
        [Parameter(Mandatory = $true)][string]$StatePath
    )

    $validatedDirectory = Get-CanonicalRunDirectory `
        -RequestedPath $SecureRunDirectory `
        -Create $false
    if (-not [string]::Equals(
            $validatedDirectory,
            $SecureRunDirectory,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Cleanup refused a run directory whose exact local identity changed.'
    }

    $prefix = $validatedDirectory.TrimEnd('\') + '\'
    $canonicalPaths = [Collections.Generic.List[string]]::new()
    $expectedPaths = [Collections.Generic.HashSet[string]]::new(
        [StringComparer]::OrdinalIgnoreCase)
    foreach ($path in $Paths) {
        $canonical = [IO.Path]::GetFullPath($path)
        if (-not $canonical.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Cleanup refused a path outside the secure run directory.'
        }
        $canonicalPaths.Add($canonical)
        if (-not $expectedPaths.Add($canonical)) {
            throw 'Cleanup received a duplicate exact artifact path.'
        }
        if (Test-Path -LiteralPath $canonical -PathType Container) {
            throw "Cleanup refused a directory where an exact artifact file was expected: $([IO.Path]::GetFileName($canonical))"
        }
        if (Test-Path -LiteralPath $canonical -PathType Leaf) {
            $item = Get-Item -LiteralPath $canonical -Force
            $actual = [IO.Path]::GetFullPath($item.FullName)
            if (-not [string]::Equals($actual, $canonical, [StringComparison]::OrdinalIgnoreCase) -or
                ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "Cleanup refused an artifact whose exact local identity changed: $([IO.Path]::GetFileName($canonical))"
            }
        }
    }

    $canonicalStatePath = [IO.Path]::GetFullPath($StatePath)
    if (-not $expectedPaths.Contains($canonicalStatePath) -or
        -not (Test-Path -LiteralPath $canonicalStatePath -PathType Leaf)) {
        throw 'Cleanup cannot preserve the exact seeder state needed for retry.'
    }
    Assert-SecurePathAcl -Path $canonicalStatePath -PathType File
    $stateBytes = [IO.File]::ReadAllBytes($canonicalStatePath)

    foreach ($member in @(Get-ChildItem -LiteralPath $validatedDirectory -Force)) {
        $memberPath = [IO.Path]::GetFullPath($member.FullName)
        if ($member.PSIsContainer -or -not $expectedPaths.Contains($memberPath)) {
            throw "Exact detach retained unexpected artifact before cleanup: $($member.Name)"
        }
    }

    foreach ($canonical in @($canonicalPaths | Where-Object {
            -not [string]::Equals(
                $_,
                $canonicalStatePath,
                [StringComparison]::OrdinalIgnoreCase)
        })) {
        $revalidatedDirectory = Get-CanonicalRunDirectory `
            -RequestedPath $validatedDirectory `
            -Create $false
        if (-not [string]::Equals(
                $revalidatedDirectory,
                $validatedDirectory,
                [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Cleanup refused a run directory whose exact local identity changed.'
        }
        if (Test-Path -LiteralPath $canonical -PathType Leaf) {
            $item = Get-Item -LiteralPath $canonical -Force
            $actual = [IO.Path]::GetFullPath($item.FullName)
            if (-not [string]::Equals($actual, $canonical, [StringComparison]::OrdinalIgnoreCase) -or
                ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "Cleanup refused an artifact whose exact local identity changed: $([IO.Path]::GetFileName($canonical))"
            }
            Remove-Item -LiteralPath $canonical -Force
        }
    }

    $finalDirectory = Get-CanonicalRunDirectory `
        -RequestedPath $validatedDirectory `
        -Create $false
    $remaining = @(Get-ChildItem -LiteralPath $finalDirectory -Force)
    if ($remaining.Count -eq 1 -and
        -not $remaining[0].PSIsContainer -and
        [string]::Equals(
            [IO.Path]::GetFullPath($remaining[0].FullName),
            $canonicalStatePath,
            [StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $canonicalStatePath -Force
        $directoryRemovalFailure = $null
        try {
            Remove-Item -LiteralPath $finalDirectory -Force
        }
        catch {
            $directoryRemovalFailure = $_
        }
        if ($null -ne $directoryRemovalFailure -or
            (Test-Path -LiteralPath $finalDirectory)) {
            if ((Test-Path -LiteralPath $finalDirectory -PathType Container) -and
                -not (Test-Path -LiteralPath $canonicalStatePath)) {
                try {
                    $stream = [IO.FileStream]::new(
                        $canonicalStatePath,
                        [IO.FileMode]::CreateNew,
                        [IO.FileAccess]::Write,
                        [IO.FileShare]::None)
                    try {
                        $stream.Write($stateBytes, 0, $stateBytes.Length)
                        $stream.Flush($true)
                    }
                    finally {
                        $stream.Dispose()
                    }
                    Set-SecureFileAcl -Path $canonicalStatePath
                    Assert-SecurePathAcl -Path $canonicalStatePath -PathType File
                }
                catch {
                    throw "Cleanup could not remove the exact run directory or restore retry state: $($_.Exception.Message)"
                }
            }
            if ($null -ne $directoryRemovalFailure) {
                throw $directoryRemovalFailure
            }
            throw 'Cleanup did not remove the exact empty secure run directory; retry state was restored.'
        }
        return
    }
    $retained = @($remaining | ForEach-Object { $_.Name }) -join ', '
    throw "Exact detach retained artifact(s): $retained"
}

$principal = [Security.Principal.WindowsPrincipal]::new(
    [Security.Principal.WindowsIdentity]::GetCurrent())
if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Run the Phase 4 fixture seeder from a non-elevated PowerShell process.'
}
if (-not (Test-Path -LiteralPath $outlookPath -PathType Leaf)) {
    throw "Classic Outlook was not found at $outlookPath"
}
Assert-SeederIntegration
Assert-OutlookStopped
$initialRegistration = @(Get-RegistrationState)
if ($initialRegistration.Count -gt 0) {
    throw "An OutlookClassicMcp VSTO registration already exists: $($initialRegistration -join ', '). Clean it before seeding."
}

$profileCreated = $false
if (-not (Test-ProfileExists -ProfileName $Profile)) {
    if ($Action -ne 'Seed') {
        throw "The exact Outlook profile '$Profile' does not exist."
    }
    Start-ProfileBootstrap -ProfileName $Profile
    $profileCreated = $true
}

$secureRunDirectory = Get-CanonicalRunDirectory `
    -RequestedPath $RunDirectory `
    -Create ($Action -eq 'Seed')
$statePath = Get-CanonicalChildPath -Parent $secureRunDirectory -Name $stateFileName
$pstAPath = Get-CanonicalChildPath -Parent $secureRunDirectory -Name $pstAFileName
$pstBPath = Get-CanonicalChildPath -Parent $secureRunDirectory -Name $pstBFileName
$fixturePath = Get-CanonicalChildPath -Parent $secureRunDirectory -Name $fixtureFileName
$inventoryPath = Get-CanonicalChildPath -Parent $secureRunDirectory -Name $inventoryFileName
$seedStatusPath = Get-CanonicalChildPath -Parent $secureRunDirectory -Name $seedStatusFileName
$detachStatusPath = Get-CanonicalChildPath -Parent $secureRunDirectory -Name $detachStatusFileName
$progressPath = Get-CanonicalChildPath -Parent $secureRunDirectory -Name $progressFileName

if (Test-Path -LiteralPath $statePath -PathType Leaf) {
    Assert-SecurePathAcl -Path $statePath -PathType File
    $state = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 |
        ConvertFrom-Json -ErrorAction Stop
    if ($state.schema -ne 1 -or
        $state.profile -cne $Profile -or
        $state.runId -cnotmatch '^[0-9a-f]{32}$' -or
        $state.pstAFileName -cne $pstAFileName -or
        $state.pstBFileName -cne $pstBFileName) {
        throw 'The secure run directory contains incompatible seeder state.'
    }
    $runId = [string]$state.runId
    $profileCreated = [bool]$state.profileCreatedByDriver
}
else {
    if ($Action -ne 'Seed') {
        throw 'The secure run directory has no seeder state for exact detach.'
    }
    $unexpected = @(Get-ChildItem -LiteralPath $secureRunDirectory -Force)
    if ($unexpected.Count -gt 0) {
        throw 'A new seed run requires an empty secure run directory.'
    }
    $runId = [Guid]::NewGuid().ToString('N')
    $state = [ordered]@{
        schema = 1
        runId = $runId
        profile = $Profile
        profileCreatedByDriver = $profileCreated
        pstAFileName = $pstAFileName
        pstBFileName = $pstBFileName
    }
    Write-AtomicSecureJson -Path $statePath -Value $state
}

$msbuild = Get-FullVisualStudioMsBuild
$certificate = $null
$certificateThumbprint = $null
$certificateSubject = "CN=OutlookClassicMcp Ephemeral Seeder $([Guid]::NewGuid().ToString('N'))"
$keyContainerName = "OutlookClassicMcp-Seeder-$([Guid]::NewGuid().ToString('N'))"
$keyProvider = [Security.Cryptography.CngProvider]::MicrosoftSoftwareKeyStorageProvider
$registrationAttempted = $false
$primaryFailure = $null
$cleanupFailures = [Collections.Generic.List[string]]::new()
$status = $null

Push-Location $repositoryRoot
try {
    try {
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
        Invoke-VstoTarget `
            -MsBuild $msbuild `
            -Target Build `
            -CertificateThumbprint $certificateThumbprint
        Assert-ExactRegistration

        $statusPath = if ($Action -eq 'Seed') { $seedStatusPath } else { $detachStatusPath }
        $status = Invoke-SeederOutlook `
            -RunId $runId `
            -SecureRunDirectory $secureRunDirectory `
            -PstAPath $pstAPath `
            -PstBPath $pstBPath `
            -StatusPath $statusPath `
            -ProgressPath $progressPath

        if ($Action -eq 'Seed') {
            foreach ($path in @(
                $pstAPath,
                $pstBPath,
                $fixturePath,
                $inventoryPath,
                $seedStatusPath,
                $progressPath)) {
                if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                    throw 'The successful seed run did not leave every required local artifact.'
                }
                Set-SecureFileAcl -Path $path
                Assert-SecurePathAcl -Path $path -PathType File
            }
        }
        else {
            Remove-ExactRunArtifacts `
                -SecureRunDirectory $secureRunDirectory `
                -StatePath $statePath `
                -Paths @(
                    $pstAPath,
                    $pstBPath,
                    $fixturePath,
                    $inventoryPath,
                    $seedStatusPath,
                    $detachStatusPath,
                    $progressPath,
                    $statePath,
                    (Get-CanonicalChildPath -Parent $secureRunDirectory -Name ("OCM-P4-$runId-Attachment-A.txt")),
                    (Get-CanonicalChildPath -Parent $secureRunDirectory -Name ("OCM-P4-$runId-Attachment-B.bin"))
                )
        }
    }
    catch {
        $primaryFailure = $_
        foreach ($path in @($pstAPath, $pstBPath)) {
            if (Test-Path -LiteralPath $path -PathType Leaf) {
                try {
                    Set-SecureFileAcl -Path $path
                    Assert-SecurePathAcl -Path $path -PathType File
                }
                catch {
                    $cleanupFailures.Add("Failed to secure a partial fixture PST: $($_.Exception.Message)")
                }
            }
        }
    }
    finally {
        $runningOutlook = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
        if ($runningOutlook.Count -gt 0 -and $registrationAttempted) {
            $cleanupFailures.Add(
                "Outlook is still running as PID(s) $($runningOutlook.Id -join ', '); the loaded registration and certificate were retained.")
        }
        else {
            if ($registrationAttempted -and $null -ne $certificateThumbprint) {
                try {
                    Invoke-VstoTarget `
                        -MsBuild $msbuild `
                        -Target VSTOClean `
                        -CertificateThumbprint $certificateThumbprint
                }
                catch {
                    $cleanupFailures.Add("VSTOClean failed: $($_.Exception.Message)")
                }
            }
            if ($null -ne $certificate) {
                $certificate.Dispose()
            }
            try {
                foreach ($stored in @(Get-ChildItem 'Cert:\CurrentUser\My' |
                    Where-Object Subject -eq $certificateSubject)) {
                    Remove-Item -LiteralPath $stored.PSPath -DeleteKey -Force
                }
                foreach ($stored in @(Get-ChildItem 'Cert:\CurrentUser\CA' |
                    Where-Object Subject -eq $certificateSubject)) {
                    Remove-Item -LiteralPath $stored.PSPath -Force
                }
                if ([Security.Cryptography.CngKey]::Exists($keyContainerName, $keyProvider)) {
                    $key = [Security.Cryptography.CngKey]::Open($keyContainerName, $keyProvider)
                    try {
                        $key.Delete()
                    }
                    finally {
                        $key.Dispose()
                    }
                }
            }
            catch {
                $cleanupFailures.Add("Ephemeral certificate cleanup failed: $($_.Exception.Message)")
            }
            if (@(Get-RegistrationState).Count -gt 0) {
                $cleanupFailures.Add('The temporary Outlook add-in registration remains after VSTOClean.')
            }
        }
    }
}
finally {
    Pop-Location
}

if ($null -ne $primaryFailure) {
    if ($cleanupFailures.Count -gt 0) {
        $primaryFailure.Exception.Data['CleanupFailures'] = $cleanupFailures -join ' '
    }
    throw $primaryFailure
}
if ($cleanupFailures.Count -gt 0) {
    throw "Seeder cleanup failed: $($cleanupFailures -join ' ')"
}

if ($Action -eq 'Seed') {
    [pscustomobject]@{
        Action = 'Seed'
        Success = $true
        RunDirectory = $secureRunDirectory
        ReadFixturePath = $fixturePath
        StoreInventoryPath = $inventoryPath
        FixtureStoreCount = [int]$status.fixtureStoreCount
        StaticPaginationItemCount = [int]$status.staticPaginationItemCount
        LargeFolderItemCount = [int]$status.largeFolderItemCount
        ConversationMode = [string]$status.conversationMode
        AttachmentCount = [int]$status.attachmentCount
        ProfileRetained = $true
        ProfileCreatedByDriver = $profileCreated
    }
}
else {
    [pscustomobject]@{
        Action = 'Detach'
        Success = $true
        DetachedStoreCount = [int]$status.detachedStoreCount
        ProfileRetained = $true
        RunDirectoryRemoved = -not (Test-Path -LiteralPath $secureRunDirectory)
    }
}
