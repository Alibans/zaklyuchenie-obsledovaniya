# Достаёт текст из .docx (расшифровка Teams, требования клиента) без Word.
# Использование: powershell -ExecutionPolicy Bypass -File tools\docx-to-text.ps1 "<файл.docx>"
# Результат: Результат\_тексты\<имя>.txt (таблицы — ячейки через " | ").
param([Parameter(Mandatory = $true)][string]$DocxPath)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$DocxPath = (Resolve-Path $DocxPath).Path
$outDir = Join-Path $PSScriptRoot '..\Результат\_тексты'
New-Item -ItemType Directory -Force $outDir | Out-Null
$zip = [IO.Compression.ZipFile]::OpenRead($DocxPath)
try {
    $reader = New-Object IO.StreamReader($zip.GetEntry('word/document.xml').Open())
    $xml = $reader.ReadToEnd()
    $reader.Close()
}
finally { $zip.Dispose() }
$text = $xml -replace '<w:tab/>', "`t" -replace '<w:br/>', "`n" -replace '</w:p>', "`n" -replace '</w:tc>', ' | ' -replace '<[^>]+>', ''
$text = [Net.WebUtility]::HtmlDecode($text)
$out = Join-Path $outDir ([IO.Path]::GetFileNameWithoutExtension($DocxPath) + '.txt')
[IO.File]::WriteAllText($out, $text, (New-Object Text.UTF8Encoding($true)))
Write-Output "Готово: $((Resolve-Path $out).Path)"
