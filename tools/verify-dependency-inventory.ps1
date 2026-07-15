#Requires -Version 5.1
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$lockFiles = Get-ChildItem -LiteralPath (Join-Path $repositoryRoot 'src'), (Join-Path $repositoryRoot 'tests') -Filter 'packages.lock.json' -File -Recurse
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
$inventoried = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($line in (Get-Content -LiteralPath $inventoryPath)) {
    if ($line -notmatch '^\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|') {
        continue
    }

    $package = $matches[1].Trim()
    if ($package -in @('Package', '---')) {
        continue
    }

    foreach ($version in $matches[2].Split(',')) {
        [void]$inventoried.Add("$package|$($version.Trim())")
    }
}

$missing = @($locked | Where-Object { -not $inventoried.Contains($_) } | Sort-Object)
$stale = @($inventoried | Where-Object { -not $locked.Contains($_) } | Sort-Object)
if ($missing.Count -gt 0 -or $stale.Count -gt 0) {
    throw "Dependency license inventory is stale. Missing: $($missing -join ', '). Stale: $($stale -join ', ')."
}

[pscustomobject]@{
    LockFiles = $lockFiles.Count
    Packages = $locked.Count
    InventoryCurrent = $true
}
