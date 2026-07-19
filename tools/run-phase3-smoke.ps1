#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Profile,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ExpectedStoreInventoryPath,

    [ValidateRange(15, 180)]
    [int]$StartupTimeoutSeconds = 60,

    [ValidateRange(15, 180)]
    [int]$ShutdownTimeoutSeconds = 60
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

& (Join-Path $PSScriptRoot 'run-phase1-smoke.ps1') `
    -Profile $Profile `
    -ExpectedStoreInventoryPath $ExpectedStoreInventoryPath `
    -StartupTimeoutSeconds $StartupTimeoutSeconds `
    -ShutdownTimeoutSeconds $ShutdownTimeoutSeconds `
    -ExpectedPhase 3
