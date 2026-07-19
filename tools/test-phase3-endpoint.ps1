#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ExpectedStoreInventoryPath,

    [Parameter(Mandatory = $true)]
    [ValidateRange(1, 2147483647)]
    [int]$OutlookProcessId,

    [ValidateSet('Full', 'Restart')]
    [string]$Mode = 'Full',

    [Uri]$Endpoint = 'http://127.0.0.1:8765/mcp/',

    [ValidateRange(1, 60)]
    [int]$TimeoutSeconds = 10,

    [AllowNull()][AllowEmptyCollection()]
    [System.Collections.Generic.HashSet[string]]$SharedOperationIds
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

& (Join-Path $PSScriptRoot 'test-phase2-endpoint.ps1') `
    -Endpoint $Endpoint `
    -TimeoutSeconds $TimeoutSeconds `
    -ExpectedPhase 3 `
    -Phase3Mode $Mode `
    -ExpectedStoreInventoryPath $ExpectedStoreInventoryPath `
    -OutlookProcessId $OutlookProcessId `
    -SharedPhase3OperationIds $SharedOperationIds
