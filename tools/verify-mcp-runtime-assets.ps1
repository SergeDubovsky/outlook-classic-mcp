#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-PublicKeyTokenText {
    param(
        [Parameter(Mandatory = $true)]
        [Reflection.AssemblyName]$AssemblyName
    )

    $token = $AssemblyName.GetPublicKeyToken()
    if ($null -eq $token -or $token.Length -eq 0) {
        return ''
    }

    return [BitConverter]::ToString($token).Replace('-', '').ToLowerInvariant()
}

function Get-AssemblyCulture {
    param(
        [Parameter(Mandatory = $true)]
        [Reflection.AssemblyName]$AssemblyName
    )

    if ([string]::IsNullOrWhiteSpace($AssemblyName.CultureName)) {
        return 'neutral'
    }

    return $AssemblyName.CultureName.ToLowerInvariant()
}

function Get-ManifestCulture {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$Identity
    )

    $culture = $Identity.GetAttribute('culture')
    if ([string]::IsNullOrWhiteSpace($culture)) {
        $culture = $Identity.GetAttribute('language')
    }
    if ([string]::IsNullOrWhiteSpace($culture)) {
        return 'neutral'
    }

    return $culture.ToLowerInvariant()
}

function Assert-ManifestIdentityMatchesAssembly {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$ManifestIdentity,

        [Parameter(Mandatory = $true)]
        [Reflection.AssemblyName]$AssemblyIdentity,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    $manifestName = $ManifestIdentity.GetAttribute('name')
    $manifestVersionText = $ManifestIdentity.GetAttribute('version')
    $manifestToken = $ManifestIdentity.GetAttribute('publicKeyToken').ToLowerInvariant()
    $manifestCulture = Get-ManifestCulture -Identity $ManifestIdentity
    $manifestArchitecture = $ManifestIdentity.GetAttribute('processorArchitecture').ToLowerInvariant()

    [Version]$manifestVersion = $null
    if (-not [Version]::TryParse($manifestVersionText, [ref]$manifestVersion)) {
        throw "$Context has an invalid manifest version: $manifestVersionText"
    }

    $assemblyToken = Get-PublicKeyTokenText -AssemblyName $AssemblyIdentity
    $assemblyCulture = Get-AssemblyCulture -AssemblyName $AssemblyIdentity
    $assemblyArchitecture = $AssemblyIdentity.ProcessorArchitecture.ToString().ToLowerInvariant()

    if (
        -not [string]::Equals($manifestName, $AssemblyIdentity.Name, [StringComparison]::Ordinal) -or
        $manifestVersion -ne $AssemblyIdentity.Version -or
        -not [string]::Equals($manifestToken, $assemblyToken, [StringComparison]::OrdinalIgnoreCase) -or
        -not [string]::Equals($manifestCulture, $assemblyCulture, [StringComparison]::OrdinalIgnoreCase) -or
        (
            -not [string]::IsNullOrWhiteSpace($manifestArchitecture) -and
            -not [string]::Equals(
                $assemblyArchitecture,
                'none',
                [StringComparison]::OrdinalIgnoreCase) -and
            -not [string]::Equals(
                $manifestArchitecture,
                $assemblyArchitecture,
                [StringComparison]::OrdinalIgnoreCase)
        )
    ) {
        throw (
            "$Context identity mismatch. Manifest: " +
            "$manifestName, Version=$manifestVersionText, Culture=$manifestCulture, " +
            "PublicKeyToken=$manifestToken, ProcessorArchitecture=$manifestArchitecture. " +
            "Assembly: $($AssemblyIdentity.FullName), ProcessorArchitecture=$assemblyArchitecture."
        )
    }
}

function Assert-ManifestIdentitiesEqual {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$Expected,

        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$Actual,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    foreach ($attributeName in @(
        'name',
        'version',
        'publicKeyToken',
        'language',
        'processorArchitecture',
        'type'
    )) {
        $expectedValue = $Expected.GetAttribute($attributeName)
        $actualValue = $Actual.GetAttribute($attributeName)
        if (-not [string]::Equals($expectedValue, $actualValue, [StringComparison]::OrdinalIgnoreCase)) {
            throw (
                "$Context identity attribute '$attributeName' does not match. " +
                "Expected '$expectedValue'; found '$actualValue'."
            )
        }
    }
}

function Get-VersionRange {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    $parts = @($Text.Split(@('-'), 2, [StringSplitOptions]::None))
    [Version]$minimum = $null
    [Version]$maximum = $null
    if ($parts.Count -eq 1) {
        if (-not [Version]::TryParse($parts[0], [ref]$minimum)) {
            throw "$Context has an invalid version: $Text"
        }
        $maximum = $minimum
    }
    elseif (
        -not [Version]::TryParse($parts[0], [ref]$minimum) -or
        -not [Version]::TryParse($parts[1], [ref]$maximum)
    ) {
        throw "$Context has an invalid version range: $Text"
    }

    if ($minimum -gt $maximum) {
        throw "$Context has a descending version range: $Text"
    }

    return [pscustomobject]@{
        Minimum = $minimum
        Maximum = $maximum
    }
}

function Test-VersionInRange {
    param(
        [Parameter(Mandatory = $true)]
        [Version]$Version,

        [Parameter(Mandatory = $true)]
        [Version]$Minimum,

        [Parameter(Mandatory = $true)]
        [Version]$Maximum
    )

    return $Version -ge $Minimum -and $Version -le $Maximum
}

function Get-Sha256Base64 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $algorithm = [Security.Cryptography.SHA256]::Create()
    try {
        return [Convert]::ToBase64String(
            $algorithm.ComputeHash([IO.File]::ReadAllBytes($Path)))
    }
    finally {
        $algorithm.Dispose()
    }
}

function Get-Sha256Hex {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $algorithm = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($Text)
        return [BitConverter]::ToString($algorithm.ComputeHash($bytes)).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $algorithm.Dispose()
    }
}

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$outputDirectory = Join-Path $repositoryRoot "src\OutlookClassicMcp.AddIn\bin\$Configuration"
$applicationManifestPath = Join-Path $outputDirectory 'OutlookClassicMcp.AddIn.dll.manifest'
$deploymentManifestPath = Join-Path $outputDirectory 'OutlookClassicMcp.AddIn.vsto'
$configurationPath = Join-Path $outputDirectory 'OutlookClassicMcp.AddIn.dll.config'
$runtimeProbePath = Join-Path $repositoryRoot 'src\OutlookClassicMcp.Transport\McpDependencyProbe.cs'
$runtimeAssemblies = @(
    'Microsoft.Bcl.AsyncInterfaces'
    'Microsoft.Bcl.Memory'
    'Microsoft.Extensions.AI.Abstractions'
    'Microsoft.Extensions.DependencyInjection.Abstractions'
    'Microsoft.Extensions.Logging.Abstractions'
    'ModelContextProtocol.Core'
    'System.Buffers'
    'System.Diagnostics.DiagnosticSource'
    'System.IO.Pipelines'
    'System.Memory'
    'System.Net.ServerSentEvents'
    'System.Numerics.Vectors'
    'System.Runtime.CompilerServices.Unsafe'
    'System.Text.Encodings.Web'
    'System.Text.Json'
    'System.Threading.Channels'
    'System.Threading.Tasks.Extensions'
)

foreach ($requiredPath in @(
    $applicationManifestPath,
    $deploymentManifestPath,
    $configurationPath,
    $runtimeProbePath
)) {
    if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
        throw "Required Phase 1 deployment file is missing: $requiredPath"
    }
}

[xml]$applicationManifest = Get-Content -LiteralPath $applicationManifestPath -Raw
$applicationIdentity = $applicationManifest.SelectSingleNode(
    "/*[local-name()='assembly']/*[local-name()='assemblyIdentity']")
if ($null -eq $applicationIdentity) {
    throw 'The VSTO application manifest has no root assembly identity.'
}

$deployedByName = @{}
$deployedByCodebase = @{}
$manifestDependencies = @(
    $applicationManifest.SelectNodes(
        "//*[local-name()='dependentAssembly' and @codebase]")
)
foreach ($dependency in $manifestDependencies) {
    $codebase = $dependency.GetAttribute('codebase')
    if (-not $codebase.EndsWith('.dll', [StringComparison]::OrdinalIgnoreCase)) {
        continue
    }

    if ($deployedByCodebase.ContainsKey($codebase)) {
        throw "The VSTO application manifest contains duplicate codebase '$codebase'."
    }

    $manifestIdentity = $dependency.SelectSingleNode("*[local-name()='assemblyIdentity']")
    if ($null -eq $manifestIdentity) {
        throw "The VSTO application manifest codebase '$codebase' has no assembly identity."
    }

    $assemblyPath = Join-Path $outputDirectory $codebase
    if (-not (Test-Path -LiteralPath $assemblyPath -PathType Leaf)) {
        throw "The VSTO application manifest references a missing assembly: $assemblyPath"
    }

    $assemblyIdentity = [Reflection.AssemblyName]::GetAssemblyName($assemblyPath)
    Assert-ManifestIdentityMatchesAssembly `
        -ManifestIdentity $manifestIdentity `
        -AssemblyIdentity $assemblyIdentity `
        -Context "Application manifest codebase '$codebase'"

    if ($deployedByName.ContainsKey($assemblyIdentity.Name)) {
        throw "The VSTO application manifest deploys assembly '$($assemblyIdentity.Name)' more than once."
    }

    $entry = [pscustomobject]@{
        Codebase = $codebase
        Path = [IO.Path]::GetFullPath($assemblyPath)
        Identity = $assemblyIdentity
        ManifestIdentity = $manifestIdentity
    }
    $deployedByCodebase[$codebase] = $entry
    $deployedByName[$assemblyIdentity.Name] = $entry
}

foreach ($assemblyFile in Get-ChildItem -LiteralPath $outputDirectory -Filter '*.dll' -File) {
    if (-not $deployedByCodebase.ContainsKey($assemblyFile.Name)) {
        throw "Built assembly is absent from the VSTO application manifest: $($assemblyFile.Name)"
    }
}

foreach ($assemblyName in $runtimeAssemblies) {
    $fileName = "$assemblyName.dll"
    if (-not $deployedByCodebase.ContainsKey($fileName)) {
        throw "The VSTO application manifest does not install $fileName."
    }
}

[xml]$deploymentManifest = Get-Content -LiteralPath $deploymentManifestPath -Raw
$deploymentIdentity = $deploymentManifest.SelectSingleNode(
    "/*[local-name()='assembly']/*[local-name()='assemblyIdentity']")
if ($null -eq $deploymentIdentity -or
    -not [string]::Equals(
        $deploymentIdentity.GetAttribute('name'),
        'OutlookClassicMcp.AddIn.vsto',
        [StringComparison]::Ordinal)) {
    throw 'The VSTO deployment manifest has an unexpected root assembly identity.'
}

$applicationLinks = @(
    $deploymentManifest.SelectNodes(
        "//*[local-name()='dependentAssembly' and @dependencyType='install']") |
        Where-Object {
            [string]::Equals(
                $_.GetAttribute('codebase'),
                'OutlookClassicMcp.AddIn.dll.manifest',
                [StringComparison]::Ordinal)
        }
)
if ($applicationLinks.Count -ne 1) {
    throw 'The VSTO deployment manifest must contain exactly one link to OutlookClassicMcp.AddIn.dll.manifest.'
}

$applicationLink = $applicationLinks[0]
$linkedIdentity = $applicationLink.SelectSingleNode("*[local-name()='assemblyIdentity']")
if ($null -eq $linkedIdentity) {
    throw 'The VSTO deployment manifest application link has no assembly identity.'
}
Assert-ManifestIdentitiesEqual `
    -Expected $applicationIdentity `
    -Actual $linkedIdentity `
    -Context 'VSTO deployment-to-application manifest link'

[long]$declaredApplicationManifestSize = 0
if (-not [long]::TryParse(
    $applicationLink.GetAttribute('size'),
    [Globalization.NumberStyles]::None,
    [Globalization.CultureInfo]::InvariantCulture,
    [ref]$declaredApplicationManifestSize)) {
    throw 'The VSTO deployment manifest application link has an invalid size.'
}
$actualApplicationManifestSize = (Get-Item -LiteralPath $applicationManifestPath).Length
if ($declaredApplicationManifestSize -ne $actualApplicationManifestSize) {
    throw (
        'The VSTO deployment manifest application-manifest size is stale. ' +
        "Declared $declaredApplicationManifestSize; actual $actualApplicationManifestSize."
    )
}

$digestMethod = $applicationLink.SelectSingleNode(".//*[local-name()='DigestMethod']")
$digestValue = $applicationLink.SelectSingleNode(".//*[local-name()='DigestValue']")
$identityTransform = $applicationLink.SelectSingleNode(
    ".//*[local-name()='Transform' and @Algorithm='urn:schemas-microsoft-com:HashTransforms.Identity']")
if ($null -eq $digestMethod -or
    $digestMethod.GetAttribute('Algorithm') -ne 'http://www.w3.org/2000/09/xmldsig#sha256' -or
    $null -eq $identityTransform -or
    $null -eq $digestValue) {
    throw 'The VSTO deployment manifest application link does not use the expected SHA-256 identity digest.'
}
$actualApplicationManifestDigest = Get-Sha256Base64 -Path $applicationManifestPath
if (-not [string]::Equals(
    $digestValue.InnerText.Trim(),
    $actualApplicationManifestDigest,
    [StringComparison]::Ordinal)) {
    throw 'The VSTO deployment manifest application-manifest digest is stale.'
}

$observedMismatches = New-Object 'System.Collections.Generic.List[object]'
foreach ($entry in $deployedByName.Values) {
    try {
        $assembly = [Reflection.Assembly]::LoadFile($entry.Path)
        $references = $assembly.GetReferencedAssemblies()
    }
    catch {
        throw "Unable to inspect assembly references for '$($entry.Codebase)': $($_.Exception.Message)"
    }

    foreach ($reference in $references) {
        if (-not $deployedByName.ContainsKey($reference.Name)) {
            continue
        }

        $deployedEntry = $deployedByName[$reference.Name]
        $deployedIdentity = $deployedEntry.Identity
        $referenceToken = Get-PublicKeyTokenText -AssemblyName $reference
        $deployedToken = Get-PublicKeyTokenText -AssemblyName $deployedIdentity
        $referenceCulture = Get-AssemblyCulture -AssemblyName $reference
        $deployedCulture = Get-AssemblyCulture -AssemblyName $deployedIdentity
        if (
            -not [string]::Equals($referenceToken, $deployedToken, [StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals($referenceCulture, $deployedCulture, [StringComparison]::OrdinalIgnoreCase)
        ) {
            throw (
                "Assembly '$($entry.Codebase)' references incompatible identity '$($reference.FullName)', " +
                "but the application manifest deploys '$($deployedIdentity.FullName)'."
            )
        }

        if ($reference.Version -ne $deployedIdentity.Version) {
            [void]$observedMismatches.Add([pscustomobject]@{
                Source = $entry.Codebase
                Name = $reference.Name
                RequestedVersion = $reference.Version
                DeployedVersion = $deployedIdentity.Version
                PublicKeyToken = $deployedToken
                Culture = $deployedCulture
            })
        }
    }
}

[xml]$configurationXml = Get-Content -LiteralPath $configurationPath -Raw
$redirects = New-Object 'System.Collections.Generic.List[object]'
foreach ($dependentAssembly in $configurationXml.SelectNodes("//*[local-name()='dependentAssembly']")) {
    $identity = $dependentAssembly.SelectSingleNode("*[local-name()='assemblyIdentity']")
    $bindingRedirect = $dependentAssembly.SelectSingleNode("*[local-name()='bindingRedirect']")
    if ($null -eq $identity -or $null -eq $bindingRedirect) {
        continue
    }

    $name = $identity.GetAttribute('name')
    if (-not $deployedByName.ContainsKey($name)) {
        continue
    }

    $deployedIdentity = $deployedByName[$name].Identity
    $expectedToken = Get-PublicKeyTokenText -AssemblyName $deployedIdentity
    $configuredToken = $identity.GetAttribute('publicKeyToken').ToLowerInvariant()
    $expectedCulture = Get-AssemblyCulture -AssemblyName $deployedIdentity
    $configuredCulture = $identity.GetAttribute('culture').ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($configuredCulture)) {
        $configuredCulture = 'neutral'
    }

    if (-not [string]::Equals($configuredToken, $expectedToken, [StringComparison]::OrdinalIgnoreCase) -or
        -not [string]::Equals($configuredCulture, $expectedCulture, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Binding redirect '$name' has an incorrect assembly identity."
    }

    [Version]$newVersion = $null
    $newVersionText = $bindingRedirect.GetAttribute('newVersion')
    if (-not [Version]::TryParse($newVersionText, [ref]$newVersion)) {
        throw "Binding redirect '$name' has an invalid newVersion: $newVersionText"
    }
    if ($newVersion -ne $deployedIdentity.Version) {
        throw (
            "Binding redirect '$name' targets $newVersion, " +
            "but the application manifest deploys $($deployedIdentity.Version)."
        )
    }

    $range = Get-VersionRange `
        -Text $bindingRedirect.GetAttribute('oldVersion') `
        -Context "Binding redirect '$name'"
    [void]$redirects.Add([pscustomobject]@{
        Name = $name
        PublicKeyToken = $configuredToken
        Culture = $configuredCulture
        MinimumVersion = $range.Minimum
        MaximumVersion = $range.Maximum
        NewVersion = $newVersion
    })
}

foreach ($mismatch in $observedMismatches) {
    $matchingRedirect = $null
    foreach ($redirect in $redirects) {
        if (
            [string]::Equals($redirect.Name, $mismatch.Name, [StringComparison]::OrdinalIgnoreCase) -and
            [string]::Equals(
                $redirect.PublicKeyToken,
                $mismatch.PublicKeyToken,
                [StringComparison]::OrdinalIgnoreCase) -and
            [string]::Equals($redirect.Culture, $mismatch.Culture, [StringComparison]::OrdinalIgnoreCase) -and
            $redirect.NewVersion -eq $mismatch.DeployedVersion -and
            (Test-VersionInRange `
                -Version $mismatch.RequestedVersion `
                -Minimum $redirect.MinimumVersion `
                -Maximum $redirect.MaximumVersion)
        ) {
            $matchingRedirect = $redirect
            break
        }
    }

    if ($null -eq $matchingRedirect) {
        throw (
            "Assembly '$($mismatch.Source)' requests $($mismatch.Name) $($mismatch.RequestedVersion), " +
            "but no binding redirect covers it to deployed version $($mismatch.DeployedVersion)."
        )
    }
}

if (Test-Path -LiteralPath (Join-Path $outputDirectory 'System.ValueTuple.dll')) {
    throw 'System.ValueTuple.dll should not be deployed for net48; the package supplies a framework placeholder.'
}

$runtimeProbeSource = Get-Content -LiteralPath $runtimeProbePath -Raw
$versionMatch = [regex]::Match(
    $runtimeProbeSource,
    'coreVersion\s*!=\s*new\s+Version\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)')
$tokenMatch = [regex]::Match(
    $runtimeProbeSource,
    'string\.Equals\(\s*tokenText\s*,\s*"([0-9a-fA-F]+)"')
$identityBlockMatch = [regex]::Match(
    $runtimeProbeSource,
    'RequiredRuntimeAssemblyIdentities\s*=\s*\{(?<body>.*?)\};',
    [Text.RegularExpressions.RegexOptions]::Singleline)
$pinnedFingerprintMatch = [regex]::Match(
    $runtimeProbeSource,
    'identitySha256\s*,\s*"([0-9a-fA-F]{64})"')
if (-not $versionMatch.Success -or
    -not $tokenMatch.Success -or
    -not $identityBlockMatch.Success -or
    -not $pinnedFingerprintMatch.Success) {
    throw 'Unable to read the pinned MCP assembly identity contract from McpDependencyProbe.cs.'
}

$probeVersion = [Version](
    "$($versionMatch.Groups[1].Value).$($versionMatch.Groups[2].Value)." +
    "$($versionMatch.Groups[3].Value).$($versionMatch.Groups[4].Value)")
$probeToken = $tokenMatch.Groups[1].Value.ToLowerInvariant()
$mcpIdentity = $deployedByName['ModelContextProtocol.Core'].Identity
$mcpToken = Get-PublicKeyTokenText -AssemblyName $mcpIdentity
if ($mcpIdentity.Version -ne $probeVersion -or
    -not [string]::Equals($mcpToken, $probeToken, [StringComparison]::OrdinalIgnoreCase)) {
    throw (
        "The deployed MCP identity '$($mcpIdentity.FullName)' does not match the runtime probe contract " +
        "Version=$probeVersion, PublicKeyToken=$probeToken."
    )
}

$requiredIdentityMatches = @(
    [regex]::Matches(
        $identityBlockMatch.Groups['body'].Value,
        '"([^"\r\n]+)"')
)
if ($requiredIdentityMatches.Count -ne $runtimeAssemblies.Count) {
    throw (
        'The runtime probe identity closure and deployment verifier assembly list differ in size. ' +
        "Probe=$($requiredIdentityMatches.Count); verifier=$($runtimeAssemblies.Count)."
    )
}

$deployedRuntimeIdentities = New-Object 'System.Collections.Generic.List[string]'
$requiredRuntimeNames = @{}
foreach ($identityMatch in $requiredIdentityMatches) {
    $requiredIdentityText = $identityMatch.Groups[1].Value
    try {
        $requiredIdentity = New-Object Reflection.AssemblyName -ArgumentList $requiredIdentityText
    }
    catch {
        throw "The runtime probe contains an invalid assembly identity '$requiredIdentityText'."
    }

    if ([string]::IsNullOrWhiteSpace($requiredIdentity.Name)) {
        throw "The runtime probe contains a nameless assembly identity '$requiredIdentityText'."
    }
    if ($requiredRuntimeNames.ContainsKey($requiredIdentity.Name)) {
        throw "The runtime probe contains duplicate assembly identity '$($requiredIdentity.Name)'."
    }
    $requiredRuntimeNames[$requiredIdentity.Name] = $requiredIdentityText

    if (-not $deployedByName.ContainsKey($requiredIdentity.Name)) {
        throw "The runtime probe requires an assembly that is not deployed: $($requiredIdentity.Name)"
    }

    $deployedIdentityText = $deployedByName[$requiredIdentity.Name].Identity.FullName
    if (-not [string]::Equals(
        $deployedIdentityText,
        $requiredIdentityText,
        [StringComparison]::Ordinal)) {
        throw (
            "The deployed identity for '$($requiredIdentity.Name)' does not match the runtime probe contract. " +
            "Expected '$requiredIdentityText'; found '$deployedIdentityText'."
        )
    }
    [void]$deployedRuntimeIdentities.Add($deployedIdentityText)
}

foreach ($runtimeAssemblyName in $runtimeAssemblies) {
    if (-not $requiredRuntimeNames.ContainsKey($runtimeAssemblyName)) {
        throw "The deployment verifier assembly '$runtimeAssemblyName' is absent from the runtime probe contract."
    }
}

$sortedRuntimeIdentities = $deployedRuntimeIdentities.ToArray()
[Array]::Sort($sortedRuntimeIdentities, [StringComparer]::Ordinal)
$runtimeIdentityCanonical = [string]::Join("`n", $sortedRuntimeIdentities)
$runtimeIdentityFingerprint = (Get-Sha256Hex -Text $runtimeIdentityCanonical).ToUpperInvariant()
$pinnedRuntimeIdentityFingerprint = $pinnedFingerprintMatch.Groups[1].Value.ToUpperInvariant()
if (-not [string]::Equals(
    $runtimeIdentityFingerprint,
    $pinnedRuntimeIdentityFingerprint,
    [StringComparison]::Ordinal)) {
    throw (
        "The deployed runtime identity fingerprint '$runtimeIdentityFingerprint' does not match " +
        "the probe contract '$pinnedRuntimeIdentityFingerprint'."
    )
}

[pscustomobject]@{
    Configuration = $Configuration
    RuntimeAssemblies = $runtimeAssemblies.Count
    DeployedAssemblies = $deployedByName.Count
    ObservedVersionMismatches = $observedMismatches.Count
    BindingRedirects = $redirects.Count
    McpCoreIdentity = $mcpIdentity.FullName
    RuntimeIdentityFingerprint = $runtimeIdentityFingerprint
    ApplicationManifestVerified = $true
    DeploymentManifestVerified = $true
}
