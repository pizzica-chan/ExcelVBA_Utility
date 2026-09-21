# UTF-8 BOM で保存すること。Windows PowerShell 5.1 は BOM がないと日本語を誤って読みます。
param(
    [string]$OutputPath = "",
    [switch]$SkipTest,
    [switch]$Visible
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$sourceDir = Join-Path $repoRoot "src"
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $repoRoot "dist\VBA_Lib.xlsm"
}
$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)

function Read-VbaSource([string]$Path) {
    $utf8 = New-Object System.Text.UTF8Encoding $false
    $text = [System.IO.File]::ReadAllText($Path, $utf8)
    if ($text.Length -gt 0 -and [int][char]$text[0] -eq 0xFEFF) {
        $text = $text.Substring(1)
    }
    return $text
}

function Get-VbaBody([string]$Text) {
    $lines = @($Text -split "\r\n|\n|\r" | Where-Object { $_ -notmatch '^\s*Attribute\s' })
    if ($lines.Count -eq 0) { return "" }
    return (($lines -join "`r`n").Trim() + "`r`n")
}

$excel = $null
$wb = $null
$saved = $false
$exitCode = 0
try {
    if (-not (Test-Path $sourceDir)) {
        throw "src フォルダが見つかりません: $sourceDir"
    }
    $files = @(Get-ChildItem -Path $sourceDir -Filter "*.bas" | Sort-Object Name)
    if ($files.Count -eq 0) {
        throw "src に .bas ファイルがありません。"
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$Visible
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1

    $wb = $excel.Workbooks.Add()
    $project = $null
    try {
        $project = $wb.VBProject
    } catch {
        $project = $null
    }
    if ($null -eq $project) {
        throw @"
VBA プロジェクトへアクセスできません。
Excel の [ファイル] → [オプション] → [トラスト センター] → [トラスト センターの設定] → [マクロの設定] で
「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」をオンにしてから、このスクリプトを再実行してください。
"@
    }

    foreach ($file in $files) {
        $text = Read-VbaSource $file.FullName
        if ($text -notmatch '(?m)^Attribute VB_Name = "([^"]+)"') {
            throw "Attribute VB_Name がありません: $($file.Name)"
        }
        $moduleName = $Matches[1]
        $body = Get-VbaBody $text
        $component = $wb.VBProject.VBComponents.Add(1)
        $component.Name = $moduleName
        if ($component.CodeModule.CountOfLines -gt 0) {
            $component.CodeModule.DeleteLines(1, $component.CodeModule.CountOfLines)
        }
        $component.CodeModule.AddFromString($body)
        Write-Host "取り込み: $moduleName"
    }

    $guide = $wb.Worksheets.Item(1)
    $guide.Name = "使い方"
    $guide.Range("A1").Value = "VBA_Lib"
    $guide.Range("A1").Font.Size = 18
    $guide.Range("A1").Font.Bold = $true
    $guide.Range("A3").Value = "汎用の Excel VBA です。標準モジュールの関数を、いつものマクロから呼び出せます。"
    $guide.Range("A4").Value = "ボタンやショートカットに割り当てるのは、Run_ で始まるマクロだけにしてください。"
    $guide.Range("A5").Value = "関数の一覧と使い方は、このファイルと同じ場所にある README.md に書いてあります。"
    $guide.Range("A7").Value = "呼び出し例"
    $guide.Range("A7").Font.Bold = $true
    $guide.Range("A8").Value = "LastRow ActiveSheet, ""B"""
    $guide.Range("A9").Value = "GetOrCreateSheet ""集計"""
    $guide.Range("A10").Value = "JapaneseHolidayName Date"
    $guide.Range("A11").Value = "ExportRangeCsv Selection, filePath"
    $guide.Columns.Item("A").ColumnWidth = 88

    if (-not $SkipTest) {
        Write-Host "自己テストを実行しています..."
        $excel.Run("SelfTest", $true)
        Write-Host "自己テストに成功しました。"
    }

    $outDir = Split-Path -Parent $OutputPath
    if (-not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir | Out-Null
    }
    if (Test-Path $OutputPath) {
        Remove-Item -LiteralPath $OutputPath -Force
    }
    $wb.SaveAs($OutputPath, 52)
    $saved = $true
    Write-Host "保存しました: $OutputPath"
} catch {
    $exitCode = 1
    $logPath = Join-Path $repoRoot "dist\import-log.txt"
    $logDir = Split-Path -Parent $logPath
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
    $utf8 = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($logPath, ($_.Exception.Message + "`r`n" + $_.ScriptStackTrace), $utf8)
    Write-Host $_.Exception.Message
    if (-not $saved) {
        Write-Host "ブックは保存していません。詳細: $logPath"
    }
} finally {
    if ($null -ne $wb) {
        $wb.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
    if ($null -ne $excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

exit $exitCode
