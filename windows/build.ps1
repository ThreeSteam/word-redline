$ErrorActionPreference = 'Stop'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (!(Test-Path -LiteralPath $compiler)) {
    $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
}
if (!(Test-Path -LiteralPath $compiler)) { throw '.NET Framework C# compiler was not found.' }
$output = Join-Path $PSScriptRoot 'portable\Word Redline.exe'
$framework = Split-Path $compiler
& $compiler /nologo /target:winexe "/out:$output" /reference:System.Web.Extensions.dll "/reference:$framework\WPF\PresentationFramework.dll" "/reference:$framework\WPF\PresentationCore.dll" "/reference:$framework\WPF\WindowsBase.dll" /reference:System.Xaml.dll "/resource:$PSScriptRoot\src\MainWindow.xaml,MainWindow.xaml" "/resource:$PSScriptRoot\assets\logo.png,logo.png" "/resource:$PSScriptRoot\assets\app.ico,app.ico" "/win32icon:$PSScriptRoot\assets\app.ico" (Join-Path $PSScriptRoot 'src\FluentApp.cs') (Join-Path $PSScriptRoot 'src\Matching.cs')
if ($LASTEXITCODE -ne 0) { throw 'Build failed. Close Word Redline before rebuilding.' }
Write-Output "Built: $output"
