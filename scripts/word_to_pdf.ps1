<#
用本机 Microsoft Word 把 docx 导出为 PDF（用于版面检查）。

为什么用它：官方 documents 技能的渲染脚本依赖 LibreOffice，本机未安装；
本机已装 Office，用 Word 导出可以得到与用户看到的一致排版。

用法:
  pwsh -NoProfile -File tools/bin/word_to_pdf.ps1 -Docx report.docx -Pdf report.pdf
#>
param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [Parameter(Mandatory = $true)][string]$Pdf
)

$ErrorActionPreference = "Stop"
$docxPath = (Resolve-Path -LiteralPath $Docx).Path
$pdfPath = [System.IO.Path]::GetFullPath($Pdf)
$pdfDir = Split-Path -Parent $pdfPath
if ($pdfDir -and -not (Test-Path -LiteralPath $pdfDir)) {
    New-Item -ItemType Directory -Force -Path $pdfDir | Out-Null
}

$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.ScreenUpdating = $false

    # ReadOnly:=true, AddToRecentFiles:=false —— 只读打开，不改动原文件
    $doc = $word.Documents.Open($docxPath, $false, $true, $false)

    # 先更新目录、图号、页码等域，再导出
    foreach ($hf in $doc.Headers) { $hf.Range.Fields.Update() | Out-Null }
    foreach ($hf in $doc.Footers) { $hf.Range.Fields.Update() | Out-Null }
    foreach ($sf in $doc.StoryRanges) { $sf.Fields.Update() | Out-Null }
    $doc.Fields.Update() | Out-Null
    foreach ($toc in $doc.TablesOfContents) { $toc.Update() | Out-Null }

    # wdExportFormatPDF = 17
    $doc.ExportAsFixedFormat($pdfPath, 17)

    $pages = $doc.ComputeStatistics(2)  # wdStatisticPages
    Write-Output ("PDF: {0}  页数: {1}  大小: {2:N0} KB" -f $pdfPath, $pages, ((Get-Item $pdfPath).Length / 1KB))
}
finally {
    if ($doc) { $doc.Close(0) | Out-Null }
    if ($word) {
        $word.Quit() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
