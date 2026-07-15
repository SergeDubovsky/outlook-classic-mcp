#Requires -Version 5.1
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$runtimeSubKey = 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R'
$registryBase = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
    [Microsoft.Win32.RegistryHive]::LocalMachine,
    [Microsoft.Win32.RegistryView]::Registry32
)
$runtimeKey = $null
try {
    $runtimeKey = $registryBase.OpenSubKey($runtimeSubKey)
    if ($null -eq $runtimeKey) {
        throw "The VSTO 4 runtime is not registered in the 32-bit registry view at HKLM\$runtimeSubKey"
    }

    $runtimeFeature = $runtimeKey.GetValue('VSTORFeature_CLR40')
    $runtimeVersionText = "$($runtimeKey.GetValue('Version'))"
}
finally {
    if ($null -ne $runtimeKey) {
        $runtimeKey.Dispose()
    }
    $registryBase.Dispose()
}

[Version]$runtimeVersion = $null
if (
    $runtimeFeature -ne 1 -or
    -not [Version]::TryParse($runtimeVersionText, [ref]$runtimeVersion) -or
    $runtimeVersion -lt [Version]'10.0.60910'
) {
    throw 'The VSTO 4 runtime 10.0.60910 or newer is not registered as installed.'
}

$assemblyRoot = Join-Path $env:windir 'Microsoft.NET\assembly\GAC_MSIL\Microsoft.VisualStudio.Tools.Office.Runtime'
$assembly = Get-ChildItem `
    -LiteralPath $assemblyRoot `
    -Filter 'Microsoft.VisualStudio.Tools.Office.Runtime.dll' `
    -Recurse `
    -ErrorAction SilentlyContinue |
        Select-Object -First 1
if ($null -eq $assembly) {
    throw "The VSTO runtime assembly is missing from the .NET Framework GAC: $assemblyRoot"
}

$assemblyName = [Reflection.AssemblyName]::GetAssemblyName($assembly.FullName)
$publicKeyToken = [BitConverter]::ToString($assemblyName.GetPublicKeyToken()).Replace('-', '').ToLowerInvariant()
if (
    $assemblyName.Name -ne 'Microsoft.VisualStudio.Tools.Office.Runtime' -or
    $assemblyName.Version -ne [Version]'10.0.0.0' -or
    $publicKeyToken -ne 'b03f5f7f11d50a3a'
) {
    throw "Unexpected VSTO runtime assembly identity: $($assemblyName.FullName)"
}

[pscustomobject]@{
    Version = $runtimeVersion.ToString()
    AssemblyPath = $assembly.FullName
    AssemblyIdentity = $assemblyName.FullName
}
