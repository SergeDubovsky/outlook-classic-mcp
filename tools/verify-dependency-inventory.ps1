#Requires -Version 5.1
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$sourceRoots = @(
    (Join-Path $repositoryRoot 'src')
    (Join-Path $repositoryRoot 'tests')
)
$lockFiles = Get-ChildItem -LiteralPath $sourceRoots -Filter 'packages.lock.json' -File -Recurse
$locked = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

foreach ($lockFile in $lockFiles) {
    $lock = Get-Content -LiteralPath $lockFile.FullName -Raw | ConvertFrom-Json
    foreach ($framework in $lock.dependencies.PSObject.Properties) {
        foreach ($dependency in $framework.Value.PSObject.Properties) {
            $resolved = $dependency.Value.PSObject.Properties['resolved']
            if ($null -ne $resolved -and $resolved.Value) {
                [void]$locked.Add("$($dependency.Name)|$($resolved.Value)")
            }
        }
    }
}

$inventoryPath = Join-Path $repositoryRoot 'docs\DEPENDENCY_LICENSES.md'
$inventory = [System.Collections.Generic.Dictionary[string, object]]::new(
    [StringComparer]::OrdinalIgnoreCase)
$rowPattern = '^\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|$'

foreach ($line in (Get-Content -LiteralPath $inventoryPath)) {
    $match = [regex]::Match($line, $rowPattern)
    if (-not $match.Success) {
        continue
    }

    $package = $match.Groups[1].Value.Trim()
    if ($package -in @('Package', '---')) {
        continue
    }

    $version = $match.Groups[2].Value.Trim()
    if ($version.Contains(',')) {
        throw "Dependency inventory rows must contain exactly one package version: $line"
    }

    $key = "$package|$version"
    if ($inventory.ContainsKey($key)) {
        throw "Dependency inventory contains a duplicate row for $key."
    }

    $inventory.Add(
        $key,
        [pscustomobject]@{
            Package = $package
            Version = $version
            Relationship = $match.Groups[3].Value.Trim()
            License = $match.Groups[4].Value.Trim()
            Source = $match.Groups[5].Value.Trim()
        })
}

$inventoried = [System.Collections.Generic.HashSet[string]]::new(
    $inventory.Keys,
    [StringComparer]::OrdinalIgnoreCase)
$missing = @($locked | Where-Object { -not $inventoried.Contains($_) } | Sort-Object)
$stale = @($inventoried | Where-Object { -not $locked.Contains($_) } | Sort-Object)
if ($missing.Count -gt 0 -or $stale.Count -gt 0) {
    throw "Dependency license inventory is stale. Missing: $($missing -join ', '). Stale: $($stale -join ', ')."
}

$packageRoots = [System.Collections.Generic.HashSet[string]]::new(
    [StringComparer]::OrdinalIgnoreCase)
if (-not [string]::IsNullOrWhiteSpace($env:NUGET_PACKAGES)) {
    [void]$packageRoots.Add($env:NUGET_PACKAGES)
}

foreach ($assetsFile in (Get-ChildItem -LiteralPath $sourceRoots -Filter 'project.assets.json' -File -Recurse)) {
    $assets = Get-Content -LiteralPath $assetsFile.FullName -Raw | ConvertFrom-Json
    $packageFolders = $assets.PSObject.Properties['packageFolders']
    if ($null -eq $packageFolders) {
        continue
    }

    foreach ($folder in $packageFolders.Value.PSObject.Properties) {
        [void]$packageRoots.Add($folder.Name)
    }
}

$defaultPackageRoot = Join-Path (
    [Environment]::GetFolderPath([Environment+SpecialFolder]::UserProfile)) '.nuget\packages'
[void]$packageRoots.Add($defaultPackageRoot)

$legacyMitLicenseUrls = [System.Collections.Generic.HashSet[string]]::new(
    [StringComparer]::OrdinalIgnoreCase)
foreach ($licenseUrl in @(
    'http://go.microsoft.com/fwlink/?LinkId=329770'
    'https://github.com/dotnet/corefx/blob/master/LICENSE.TXT'
    'https://github.com/dotnet/standard/blob/master/LICENSE.TXT'
)) {
    [void]$legacyMitLicenseUrls.Add($licenseUrl)
}

function Get-NormalizedUrl {
    param(
        [Parameter(Mandatory)]
        [string]$Url
    )

    return $Url.Trim().TrimEnd('/')
}

function Get-RestoredNuspecPath {
    param(
        [Parameter(Mandatory)]
        [string]$Package,

        [Parameter(Mandatory)]
        [string]$Version
    )

    foreach ($packageRoot in $packageRoots) {
        $packageDirectory = Join-Path (
            (Join-Path $packageRoot $Package.ToLowerInvariant())) $Version.ToLowerInvariant()
        if (-not (Test-Path -LiteralPath $packageDirectory -PathType Container)) {
            continue
        }

        $nuspec = Get-ChildItem -LiteralPath $packageDirectory -Filter '*.nuspec' -File |
            Select-Object -First 1
        if ($null -ne $nuspec) {
            return $nuspec.FullName
        }
    }

    throw "Restored metadata for $Package $Version was not found. Run locked restore before verifying dependencies."
}

$metadataMismatches = [System.Collections.Generic.List[string]]::new()
foreach ($key in ($locked | Sort-Object)) {
    $record = $inventory[$key]
    $nuspecPath = Get-RestoredNuspecPath -Package $record.Package -Version $record.Version
    [xml]$nuspec = Get-Content -LiteralPath $nuspecPath -Raw
    $metadata = $nuspec.SelectSingleNode("/*[local-name()='package']/*[local-name()='metadata']")
    if ($null -eq $metadata) {
        throw "NuGet metadata is missing from $nuspecPath."
    }

    $licenseNode = $metadata.SelectSingleNode("*[local-name()='license']")
    if ($null -ne $licenseNode) {
        $licenseType = $licenseNode.GetAttribute('type')
        if (-not [string]::Equals($licenseType, 'expression', [StringComparison]::OrdinalIgnoreCase)) {
            throw "Unsupported NuGet license metadata type '$licenseType' for $key. Review it manually."
        }

        $metadataLicense = $licenseNode.InnerText.Trim()
    }
    else {
        $licenseUrlNode = $metadata.SelectSingleNode("*[local-name()='licenseUrl']")
        $licenseUrl = if ($null -eq $licenseUrlNode) { '' } else { $licenseUrlNode.InnerText.Trim() }
        if (-not $legacyMitLicenseUrls.Contains($licenseUrl)) {
            throw "No reviewed license expression is available for $key. Legacy license URL: $licenseUrl"
        }

        $metadataLicense = 'MIT'
    }

    $repositoryNode = $metadata.SelectSingleNode("*[local-name()='repository']")
    $metadataSource = if ($null -ne $repositoryNode) {
        $repositoryNode.GetAttribute('url').Trim()
    }
    else {
        $projectUrlNode = $metadata.SelectSingleNode("*[local-name()='projectUrl']")
        if ($null -eq $projectUrlNode) { '' } else { $projectUrlNode.InnerText.Trim() }
    }

    if ([string]::IsNullOrWhiteSpace($metadataSource)) {
        throw "No repository or project URL is available in restored metadata for $key."
    }

    if (-not [string]::Equals(
        $record.License,
        $metadataLicense,
        [StringComparison]::OrdinalIgnoreCase)) {
        $metadataMismatches.Add(
            "$key license: inventory '$($record.License)', metadata '$metadataLicense'")
    }

    if (-not [string]::Equals(
        (Get-NormalizedUrl -Url $record.Source),
        (Get-NormalizedUrl -Url $metadataSource),
        [StringComparison]::OrdinalIgnoreCase)) {
        $metadataMismatches.Add(
            "$key source: inventory '$($record.Source)', metadata '$metadataSource'")
    }
}

if ($metadataMismatches.Count -gt 0) {
    throw "Dependency metadata inventory differs from restored NuGet metadata: $($metadataMismatches -join '; ')"
}

[pscustomobject]@{
    LockFiles = $lockFiles.Count
    Packages = $locked.Count
    InventoryCurrent = $true
    MetadataCurrent = $true
}
