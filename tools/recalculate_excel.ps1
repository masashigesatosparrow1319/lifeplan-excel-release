# Windows desktop Excel required. Run only after generating current cases.
[CmdletBinding()]
param(
    [string]$InputDirectory,
    [string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
# Resolve defaults only after parameter binding (Windows PowerShell 5.1).
$scriptDirectory = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptDirectory)) {
    $scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if ([string]::IsNullOrWhiteSpace($InputDirectory)) {
    $InputDirectory = Join-Path $scriptDirectory '..\output\recalc_input'
}
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $scriptDirectory ('..\output\recalc_verified-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
}
$InputDirectory = (Resolve-Path $InputDirectory).Path
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
if ($InputDirectory -eq $OutputDirectory) { throw '入出力フォルダーを分けてください。' }
$manifest = Get-Content (Join-Path $InputDirectory 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
foreach ($entry in $manifest) {
    $path = Join-Path $InputDirectory $entry.file
    if (!(Test-Path $path)) { throw "ケースがありません: $path" }
    if ((Get-FileHash $path -Algorithm SHA256).Hash.ToLowerInvariant() -ne $entry.sha256) {
        throw "生成後に入力が変更されています: $path"
    }
    if (Test-Path (Join-Path $OutputDirectory $entry.file)) { throw "既存の結果を上書きしません: $($entry.file)" }
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$excel = $null
$book = $null
$results = @()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    foreach ($entry in $manifest) {
        try {
            $book = $excel.Workbooks.Open((Join-Path $InputDirectory $entry.file), 0, $false)
            $book.ForceFullCalculation = $true
            $excel.CalculateFullRebuild()
            $destination = Join-Path $OutputDirectory $entry.file
            $book.SaveAs($destination, 51)
            $book.Close($false)
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) | Out-Null
            $book = $null
            $results += [pscustomobject]@{file=$entry.file; status='recalculated_and_saved'; excel_version=$excel.Version; excel_build=$excel.Build}
        } catch {
            $results += [pscustomobject]@{file=$entry.file; status='failed'; error=$_.Exception.Message}
            throw
        }
    }
} finally {
    if ($null -ne $book) {
        $book.Close($false)
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) | Out-Null
    }
    if ($null -ne $excel) {
        $excel.Quit()
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    }
    $results | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $OutputDirectory 'recalculation-log.json') -Encoding UTF8
}
Write-Output "再計算保存完了（合格判定はPython比較テストで行う）: $OutputDirectory"
