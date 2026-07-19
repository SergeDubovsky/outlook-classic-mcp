#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [DateTimeOffset]$SinceUtc,

    [Parameter(Mandatory = $true)]
    [DateTimeOffset]$UntilUtc,

    [ValidateRange(1, 10)]
    [int]$ExpectedCycles = 3,

    [string]$LogDirectory = (Join-Path $env:LOCALAPPDATA 'OutlookClassicMcp\logs')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

& (Join-Path $PSScriptRoot 'verify-phase1-smoke.ps1') `
    -SinceUtc $SinceUtc `
    -UntilUtc $UntilUtc `
    -ExpectedCycles $ExpectedCycles `
    -ExpectedPhase 2 `
    -LogDirectory $LogDirectory
