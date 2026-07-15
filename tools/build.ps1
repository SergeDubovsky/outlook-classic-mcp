#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release', 'All')]
    [string]$Configuration = 'All',

    [switch]$SkipTests,

    [switch]$UpdateLocks
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$env:DOTNET_NOLOGO = '1'
$env:DOTNET_SKIP_FIRST_TIME_EXPERIENCE = '1'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$solutionPath = Join-Path $repositoryRoot 'OutlookClassicMcp.sln'
$addInProjectPath = Join-Path $repositoryRoot 'src\OutlookClassicMcp.AddIn\OutlookClassicMcp.AddIn.csproj'
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$requiredComponents = @(
    'Microsoft.VisualStudio.Workload.Office',
    'Microsoft.VisualStudio.Component.TeamOffice',
    'Microsoft.VisualStudio.Workload.ManagedDesktop',
    'Microsoft.Net.Component.4.8.TargetingPack',
    'Microsoft.Net.Component.4.8.SDK'
)

function Get-RegistryQueryMatches {
    param(
        [Parameter(Mandatory)]
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

    foreach ($line in (Get-RegistryQueryMatches -SubKey 'Software\Microsoft\Office\Outlook\Addins\OutlookClassicMcp.AddIn')) {
        $state.Add("Outlook add-in: $line")
    }
    foreach ($line in (Get-RegistryQueryMatches -SubKey 'Software\Microsoft\Office\Outlook\FormRegions' -Search 'OutlookClassicMcp.AddIn')) {
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

function Invoke-VstoClean {
    param(
        [Parameter(Mandatory)]
        [string]$BuildConfiguration,

        [string]$CertificateThumbprint
    )

    $arguments = @(
        $addInProjectPath,
        '/t:VSTOClean',
        "/p:Configuration=$BuildConfiguration",
        '/p:Platform=AnyCPU',
        '/p:VisualStudioVersion=18.0',
        '/p:SignManifests=true',
        "/p:ManifestCertificateThumbprint=$CertificateThumbprint",
        '/v:minimal',
        '/nologo',
        '/nr:false'
    )

    $output = @(& $script:msbuild @arguments 2>&1)
    if ($LASTEXITCODE -eq 0) {
        return $null
    }

    return "$BuildConfiguration VSTOClean failed with exit code $LASTEXITCODE`: $($output -join [Environment]::NewLine)"
}

if (-not (Test-Path -LiteralPath $vswhere -PathType Leaf)) {
    throw "vswhere was not found at $vswhere"
}

$arguments = @(
    '-products', '*',
    '-version', '[18.0,19.0)',
    '-requires'
) + $requiredComponents + @('-format', 'json', '-utf8')

$instanceJson = @(& $vswhere @arguments)
$instances = @((($instanceJson -join [Environment]::NewLine) | ConvertFrom-Json))
$visualStudio = $instances |
    Where-Object { $_.PSObject.Properties.Name -contains 'productId' -and $_.productId -ne 'Microsoft.VisualStudio.Product.BuildTools' } |
    Select-Object -First 1
if ($null -eq $visualStudio) {
    throw 'A full Visual Studio 2026 IDE with the required Office/VSTO components was not found. Run tools\preflight.ps1.'
}

$script:msbuild = Join-Path $visualStudio.installationPath 'MSBuild\Current\Bin\MSBuild.exe'
if (-not (Test-Path -LiteralPath $script:msbuild -PathType Leaf)) {
    throw "MSBuild was not found at $script:msbuild"
}

$outlookProcesses = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object Id, StartTime)
if ($outlookProcesses.Count -gt 0) {
    throw "Outlook is running. Save open work and close Outlook gracefully before this VSTO build; never terminate it forcibly. Running process IDs: $($outlookProcesses.Id -join ', ')"
}

$initialVstoState = @(Get-VstoDevelopmentState)
if ($initialVstoState.Count -gt 0) {
    throw "A VSTO registration for this checkout or add-in already exists. Close Outlook gracefully and clean the intentional Visual Studio registration before running the isolated build. Detected state: $($initialVstoState -join [Environment]::NewLine)"
}

$configurations = if ($Configuration -eq 'All') { @('Debug', 'Release') } else { @($Configuration) }
$attemptedConfigurations = [System.Collections.Generic.List[string]]::new()
$cleanupFailures = [System.Collections.Generic.List[string]]::new()
$primaryFailure = $null
$certificate = $null
$certificateThumbprint = $null
$certificateSubject = "CN=OutlookClassicMcp Ephemeral Build $([Guid]::NewGuid().ToString('N'))"
$keyContainerName = "OutlookClassicMcp-Ephemeral-$([Guid]::NewGuid().ToString('N'))"
$keyProvider = [System.Security.Cryptography.CngProvider]::MicrosoftSoftwareKeyStorageProvider
$keyUniqueName = $null
$keyFilePath = $null

Push-Location $repositoryRoot
try {
    try {
        $restoreArguments = @(
            $solutionPath,
            '/t:Restore',
            '/p:RestorePackagesWithLockFile=true',
            '/p:VisualStudioVersion=18.0',
            '/v:minimal',
            '/nologo',
            '/nr:false'
        )
        if (-not $UpdateLocks) {
            $restoreArguments += '/p:RestoreLockedMode=true'
        }

        & $script:msbuild @restoreArguments
        if ($LASTEXITCODE -ne 0) {
            throw "Locked restore failed with exit code $LASTEXITCODE."
        }

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
        $keyFilePath = Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::ApplicationData)) "Microsoft\Crypto\Keys\$keyUniqueName"

        foreach ($buildConfiguration in $configurations) {
            $attemptedConfigurations.Add($buildConfiguration)
            $configurationFailure = $null
            try {
                $buildArguments = @(
                    $solutionPath,
                    '/t:Rebuild',
                    '/m',
                    "/p:Configuration=$buildConfiguration",
                    '/p:Platform=Any CPU',
                    '/p:VisualStudioVersion=18.0',
                    '/p:SignManifests=true',
                    "/p:ManifestCertificateThumbprint=$certificateThumbprint",
                    '/p:RestoreLockedMode=true',
                    '/v:minimal',
                    '/nologo',
                    '/nr:false'
                )

                & $script:msbuild @buildArguments
                if ($LASTEXITCODE -ne 0) {
                    throw "$buildConfiguration build failed with exit code $LASTEXITCODE."
                }

                if (-not $SkipTests) {
                    & dotnet test 'tests\OutlookClassicMcp.Core.Tests\OutlookClassicMcp.Core.Tests.csproj' `
                        --configuration $buildConfiguration --no-build --no-restore --logger 'console;verbosity=minimal'
                    if ($LASTEXITCODE -ne 0) {
                        throw "$buildConfiguration Core tests failed with exit code $LASTEXITCODE."
                    }

                    & dotnet test 'tests\OutlookClassicMcp.Transport.Tests\OutlookClassicMcp.Transport.Tests.csproj' `
                        --configuration $buildConfiguration --no-build --no-restore --logger 'console;verbosity=minimal'
                    if ($LASTEXITCODE -ne 0) {
                        throw "$buildConfiguration Transport tests failed with exit code $LASTEXITCODE."
                    }
                }
            }
            catch {
                $configurationFailure = $_
            }

            $vstoCleanFailure = Invoke-VstoClean -BuildConfiguration $buildConfiguration -CertificateThumbprint $certificateThumbprint
            if (-not [string]::IsNullOrWhiteSpace($vstoCleanFailure)) {
                $cleanupFailures.Add($vstoCleanFailure)
            }

            if ($null -ne $configurationFailure) {
                $primaryFailure = $configurationFailure
                break
            }
            if (-not [string]::IsNullOrWhiteSpace($vstoCleanFailure)) {
                break
            }
        }
    }
    catch {
        if ($null -eq $primaryFailure) {
            $primaryFailure = $_
        }
    }
    finally {
        foreach ($buildConfiguration in ($attemptedConfigurations | Select-Object -Unique)) {
            try {
                $vstoCleanFailure = Invoke-VstoClean -BuildConfiguration $buildConfiguration -CertificateThumbprint $certificateThumbprint
                if (-not [string]::IsNullOrWhiteSpace($vstoCleanFailure)) {
                    $cleanupFailures.Add($vstoCleanFailure)
                }
            }
            catch {
                $cleanupFailures.Add("$buildConfiguration VSTOClean raised an exception: $($_.Exception.Message)")
            }
        }

        if ($null -ne $certificate) {
            $certificate.Dispose()
        }

        try {
            $myCertificates = @(Get-ChildItem 'Cert:\CurrentUser\My' | Where-Object Subject -eq $certificateSubject)
            foreach ($myCertificate in $myCertificates) {
                Remove-Item -LiteralPath $myCertificate.PSPath -DeleteKey -Force -ErrorAction Stop
            }
        }
        catch {
            $cleanupFailures.Add("Certificate/private-key removal failed: $($_.Exception.Message)")
        }

        try {
            $authorityCertificates = @(Get-ChildItem 'Cert:\CurrentUser\CA' | Where-Object Subject -eq $certificateSubject)
            foreach ($authorityCertificate in $authorityCertificates) {
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
                        $keyFilePath = Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::ApplicationData)) "Microsoft\Crypto\Keys\$keyUniqueName"
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
                $cleanupFailures.Add('The ephemeral signing certificate remains in a current-user certificate store.')
            }
            if ([System.Security.Cryptography.CngKey]::Exists($keyContainerName, $keyProvider)) {
                $cleanupFailures.Add("The ephemeral CNG container remains: $keyContainerName")
            }
            if (-not [string]::IsNullOrWhiteSpace($keyFilePath) -and (Test-Path -LiteralPath $keyFilePath)) {
                $cleanupFailures.Add("The ephemeral CNG key file remains: $keyFilePath")
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

    if ($null -ne $primaryFailure) {
        foreach ($cleanupFailure in $cleanupFailures) {
            Write-Warning $cleanupFailure
        }
        $PSCmdlet.ThrowTerminatingError($primaryFailure)
    }
    if ($cleanupFailures.Count -gt 0) {
        throw "Build cleanup failed: $($cleanupFailures -join [Environment]::NewLine)"
    }
}
finally {
    Pop-Location
}
