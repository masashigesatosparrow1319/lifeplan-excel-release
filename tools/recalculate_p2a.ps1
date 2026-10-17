# Windows PowerShell 5.1 + desktop Excel. This file must retain UTF-8 BOM.
[CmdletBinding()]
param([string]$InputDirectory, [string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($root)) { $root = Split-Path -Parent $MyInvocation.MyCommand.Path }
if ([string]::IsNullOrWhiteSpace($InputDirectory)) { $InputDirectory = Join-Path $root '..\output\p2a_recalc_input' }
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { $OutputDirectory = Join-Path $root ('..\output\p2a_recalc_verified-' + (Get-Date -Format 'yyyyMMdd-HHmmss')) }
$InputDirectory = (Resolve-Path $InputDirectory).Path
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
if ($InputDirectory -eq $OutputDirectory) { throw '入出力を分けてください' }
$manifest = Get-Content (Join-Path $InputDirectory 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
foreach ($entry in $manifest) {
    $inputPath = Join-Path $InputDirectory $entry.file
    if ((Get-FileHash $inputPath -Algorithm SHA256).Hash.ToLowerInvariant() -ne $entry.sha256) { throw "入力更新あり: $($entry.file)" }
    if (Test-Path (Join-Path $OutputDirectory $entry.file)) { throw "既存結果を上書きしません: $($entry.file)" }
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
            $openWatch = [Diagnostics.Stopwatch]::StartNew()
            $book = $excel.Workbooks.Open((Join-Path $InputDirectory $entry.file),0,$false)
            $openWatch.Stop()
            $book.ForceFullCalculation = $true
            $watch = [Diagnostics.Stopwatch]::StartNew()
            $excel.CalculateFullRebuild()
            while ($excel.CalculationState -ne 0) { Start-Sleep -Milliseconds 100 }
            $watch.Stop()
            $saveWatch = [Diagnostics.Stopwatch]::StartNew()
            $book.SaveAs((Join-Path $OutputDirectory $entry.file),51)
            $saveWatch.Stop()
            $book.Close($false)
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) | Out-Null
            $book = $null
            $exceeded = $watch.Elapsed.TotalSeconds -gt $entry.monitor_seconds
            $results += [pscustomobject]@{
                file=$entry.file; status='recalculated_and_saved'
                excel_version=$excel.Version; excel_build=$excel.Build; operating_system=$excel.OperatingSystem
                open_seconds=$openWatch.Elapsed.TotalSeconds; full_recalc_seconds=$watch.Elapsed.TotalSeconds
                save_seconds=$saveWatch.Elapsed.TotalSeconds; monitor_seconds=$entry.monitor_seconds; monitor_exceeded=$exceeded
            }
            if ($exceeded) { Write-Warning "監視目標超過（数値合否は別判定）: $($entry.file) $($watch.Elapsed.TotalSeconds)秒" }
        } catch {
            $results += [pscustomobject]@{file=$entry.file; status='failed'; error=$_.Exception.Message}
            throw
        }
    }
} finally {
    try {
        if ($null -ne $book) {
            try { $book.Close($false) } finally { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) | Out-Null }
        }
    } finally {
        try {
            if ($null -ne $excel) {
                try { $excel.Quit() } finally { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null }
            }
        } finally {
            $results | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $OutputDirectory 'p2a-recalculation-log.json') -Encoding UTF8
        }
    }
}
Write-Output "保存完了。Python照合と性能ログを確認してください: $OutputDirectory"
