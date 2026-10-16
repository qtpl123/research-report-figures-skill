<#
把 Graphviz 的 .dot 源文件渲染成 PNG/SVG。

特点：
- 使用项目本地便携版 Graphviz，不改系统 PATH
- 默认注入中文字体（Microsoft YaHei），避免中文乱码/方框

用法:
  pwsh -NoProfile -File tools/bin/graphviz_render.ps1 -InFile fig.dot -OutFile fig.png
  ... [-Format png|svg] [-Dpi 200] [-FontName "Microsoft YaHei"]
#>
# 注意：参数名故意避开 $Input / $Output 这类 PowerShell 自动变量，避免参数绑定异常。
param(
    [Parameter(Mandatory = $true)][string]$InFile,
    [Parameter(Mandatory = $true)][string]$OutFile,
    [ValidateSet("png", "svg")][string]$Format = "png",
    [int]$Dpi = 220,
    [string]$FontName = "Microsoft YaHei"
)

$ErrorActionPreference = "Stop"
$dotExe = Join-Path $PSScriptRoot "..\graphviz\Graphviz-16.1.0-win64\bin\dot.exe"
if (-not (Test-Path -LiteralPath $dotExe)) {
    throw "找不到 dot.exe: $dotExe"
}

$inPath = (Resolve-Path -LiteralPath $InFile).Path
$outPath = [System.IO.Path]::GetFullPath($OutFile)
$outDir = Split-Path -Parent $outPath
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
}

$dotArgs = @(
    "-T$Format",
    "-Gdpi=$Dpi",
    "-Gfontname=$FontName",
    "-Nfontname=$FontName",
    "-Efontname=$FontName",
    "-o", $outPath,
    $inPath
)
& $dotExe @dotArgs
if ($LASTEXITCODE -ne 0) {
    throw "dot 渲染失败，退出码 $LASTEXITCODE"
}

if ($Format -eq "png") {
    Add-Type -AssemblyName System.Drawing
    $img = [System.Drawing.Image]::FromFile($outPath)
    Write-Output ("PNG: {0}  ({1}x{2} px, {3:N0} KB)" -f $outPath, $img.Width, $img.Height, ((Get-Item $outPath).Length / 1KB))
    $img.Dispose()
}
else {
    Write-Output ("SVG: {0}" -f $outPath)
}
