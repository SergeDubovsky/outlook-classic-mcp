#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$PiaRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($PiaRoot)) {
    $programFilesX86 = [Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFilesX86)
    $PiaRoot = Join-Path $programFilesX86 'Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15'
}

$expectedAssemblies = @(
    [pscustomobject]@{
        FileName = 'Office.dll'
        AssemblyName = 'office'
        Version = [Version]'15.0.0.0'
        PublicKeyToken = '71e9bce111e9429c'
    },
    [pscustomobject]@{
        FileName = 'Microsoft.Office.Interop.Outlook.dll'
        AssemblyName = 'Microsoft.Office.Interop.Outlook'
        Version = [Version]'15.0.0.0'
        PublicKeyToken = '71e9bce111e9429c'
    }
)

foreach ($expectedAssembly in $expectedAssemblies) {
    $assemblyPath = Join-Path $PiaRoot $expectedAssembly.FileName
    if (-not (Test-Path -LiteralPath $assemblyPath -PathType Leaf)) {
        throw "The Visual Studio Office PIA is missing: $assemblyPath"
    }

    $assemblyName = [Reflection.AssemblyName]::GetAssemblyName($assemblyPath)
    $publicKeyToken = [BitConverter]::ToString($assemblyName.GetPublicKeyToken()).Replace('-', '').ToLowerInvariant()
    if (
        $assemblyName.Name -ne $expectedAssembly.AssemblyName -or
        $assemblyName.Version -ne $expectedAssembly.Version -or
        $publicKeyToken -ne $expectedAssembly.PublicKeyToken
    ) {
        throw "Unexpected Office PIA identity: $($assemblyName.FullName)"
    }

    $assemblyFile = Get-Item -LiteralPath $assemblyPath
    [pscustomobject]@{
        FileName = $expectedAssembly.FileName
        Path = $assemblyFile.FullName
        AssemblyIdentity = $assemblyName.FullName
        FileVersion = $assemblyFile.VersionInfo.FileVersion
    }
}
