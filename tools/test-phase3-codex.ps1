#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ExpectedStoreInventoryPath,

    [ValidateRange(15, 180)]
    [int]$TimeoutSeconds = 90,

    [AllowNull()][AllowEmptyCollection()]
    [System.Collections.Generic.HashSet[string]]$SharedOperationIds
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

& (Join-Path $PSScriptRoot 'test-phase2-codex.ps1') `
    -TimeoutSeconds $TimeoutSeconds `
    -ExpectedPhase 3 `
    -ExpectedStoreInventoryPath $ExpectedStoreInventoryPath `
    -SharedPhase3OperationIds $SharedOperationIds
