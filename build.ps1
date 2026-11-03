# Builds the ready-to-copy folder dist\Fanak with the C# compiler included in Windows.
# No SDK or other installation is required.
$ErrorActionPreference = 'Stop'

$fw = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$csc = Join-Path $fw 'csc.exe'
if (-not (Test-Path $csc)) { throw "C# compiler not found: $csc" }

$root = $PSScriptRoot
$out = Join-Path $root 'dist\Fanak'
New-Item -ItemType Directory -Force -Path $out | Out-Null

$refs = @(
    "$fw\WPF\PresentationCore.dll",
    "$fw\WPF\PresentationFramework.dll",
    "$fw\WPF\WindowsBase.dll",
    "$fw\System.Xaml.dll",
    "$fw\System.Windows.Forms.dll",
    "$fw\System.Drawing.dll",
    "$fw\System.IO.Compression.dll",
    "$fw\System.Management.dll",
    "$fw\System.Xml.dll",
    "$fw\System.Xml.Linq.dll",
    "$fw\System.Runtime.WindowsRuntime.dll",
    "$env:WINDIR\Microsoft.NET\assembly\GAC_MSIL\System.Runtime\v4.0_4.0.0.0__b03f5f7f11d50a3a\System.Runtime.dll",
    "$env:WINDIR\System32\WinMetadata\Windows.Data.winmd",
    "$env:WINDIR\System32\WinMetadata\Windows.Foundation.winmd",
    "$env:WINDIR\System32\WinMetadata\Windows.Globalization.winmd",
    "$env:WINDIR\System32\WinMetadata\Windows.Graphics.winmd",
    "$env:WINDIR\System32\WinMetadata\Windows.Media.winmd",
    "$env:WINDIR\System32\WinMetadata\Windows.Storage.winmd",
    "$env:WINDIR\System32\WinMetadata\Windows.UI.winmd"
) | ForEach-Object { "/r:$_" }

$sources = Get-ChildItem -Path (Join-Path $root 'src') -Filter *.cs | ForEach-Object { $_.FullName }
$icon = Join-Path $root 'src\Fanak.ico'

& $csc /nologo /target:winexe /optimize+ /codepage:65001 "/win32icon:$icon" "/resource:$icon,Fanak.ico" "/out:$out\Fanak.exe" $refs $sources
if ($LASTEXITCODE -ne 0) { throw "Build failed (exit code $LASTEXITCODE)" }

Copy-Item -Path (Join-Path $root 'package\*') -Destination $out -Force
Write-Output "Built: $out"
