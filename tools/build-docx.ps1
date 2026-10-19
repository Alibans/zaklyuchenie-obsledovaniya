# Собирает .docx из .html через Microsoft Word (отдельный скрытый экземпляр).
# Использование: powershell -ExecutionPolicy Bypass -File tools\build-docx.ps1 "<вход.html>" "<выход.docx>"
param(
    [Parameter(Mandatory = $true)][string]$HtmlPath,
    [Parameter(Mandatory = $true)][string]$DocxPath
)
$ErrorActionPreference = 'Stop'
$HtmlPath = (Resolve-Path $HtmlPath).Path
$DocxPath = [IO.Path]::GetFullPath($DocxPath)

try { $word = New-Object -ComObject Word.Application }
catch {
    Write-Output "Word не найден. Откройте $HtmlPath в Word и сохраните как .docx."
    exit 2
}
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($HtmlPath, $false, $true)
    $doc.PageSetup.Orientation = 1          # альбомная: реестр широкий
    $doc.PageSetup.TopMargin = 42.5
    $doc.PageSetup.BottomMargin = 42.5
    $doc.PageSetup.LeftMargin = 42.5
    $doc.PageSetup.RightMargin = 42.5
    foreach ($t in $doc.Tables) {
        $t.AllowAutoFit = $true
        $t.PreferredWidthType = 2           # проценты
        $t.PreferredWidth = 100
        $t.Rows.Item(1).HeadingFormat = -1  # шапка повторяется на каждой странице
    }
    $doc.SaveAs2($DocxPath, 16)             # wdFormatDocumentDefault
    $doc.Close(0)
    Write-Output "Готово: $DocxPath"
}
finally {
    $word.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
