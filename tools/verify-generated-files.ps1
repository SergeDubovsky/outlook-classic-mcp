#Requires -Version 5.1
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$expected = [ordered]@{
    'src\OutlookClassicMcp.AddIn\ThisAddIn.Designer.cs' = '259419964B326D447D5194495265F93C4A08600D1760CA1724614051168DC436'
    'src\OutlookClassicMcp.AddIn\ThisAddIn.Designer.xml' = 'CE48FE49C83A3F7B16E4E72EDEA95A59C0143EC40B321BBE269C811ADFD3920B'
    'src\OutlookClassicMcp.AddIn\Properties\Resources.Designer.cs' = '6FB6715D3EE01AA9263A2119109A9CB8499600192880C70D2CE3132433892855'
    'src\OutlookClassicMcp.AddIn\Properties\Settings.Designer.cs' = '671815A8F70DB2530DC85758C6ECB036A09AFD2B3ACA5E17D221E915FF8133BB'
}

$results = foreach ($relativePath in $expected.Keys) {
    $path = Join-Path $repositoryRoot $relativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Generated VSTO file is missing: $relativePath"
    }

    $actual = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    if ($actual -ne $expected[$relativePath]) {
        throw "Generated VSTO file changed without an updated reviewed baseline: $relativePath"
    }

    [pscustomobject]@{
        File = $relativePath
        Sha256 = $actual
    }
}

$results
