<#
    監査勤務形態一覧 ６ブック統合スクリプト

    \\192.168.10.100\iwakubo\指導監査用\▲監査勤務形態一覧▲ にある
    「２０２６０１～のシフト管理表(事業所).xlsx」６ファイルを
    １つのブックに統合します。

    順番：西九条 → 九条 → 酉島 → 新郡山 → 春日出 → 出来島

    Excel 本体のシートコピー機能を使うため、書式・罫線・セル色（夜勤の黄色塗り）・
    数式・列幅・印刷設定はすべて元のまま引き継がれます。

    使い方（通常）：同じフォルダの「統合実行.bat」をダブルクリック
    使い方（手動）：
        powershell -NoProfile -ExecutionPolicy Bypass -File .\Merge-ShiftWorkbooks.ps1
        powershell ... -Root "D:\作業用\監査" -OutFile "D:\作業用\統合.xlsx"
#>
[CmdletBinding()]
param(
    # 元ファイルが置いてあるフォルダ
    [string]$Root = '\\192.168.10.100\iwakubo\指導監査用\▲監査勤務形態一覧▲',

    # 出力先。省略時は $Root に 統合_監査勤務形態一覧_yyyyMMdd_HHmm.xlsx を作成
    [string]$OutFile,

    # 統合する順番（事業所名）
    [string[]]$Order = @('西九条', '九条', '酉島', '新郡山', '春日出', '出来島'),

    # 取り込まないシート名（Sheet1 は他事業所名が残った索引シートのため既定で除外）
    [string[]]$SkipSheets = @('Sheet1'),

    # 目次シートを作らない
    [switch]$NoIndexSheet,

    # 統合後に Excel を開いたままにする（保存はされます）
    [switch]$KeepExcelOpen,

    # 出力先が既にある場合に上書きする
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# Excel 定数
$xlCalculationManual    = -4135
$xlCalculationAutomatic = -4105
$xlOpenXMLWorkbook      = 51
$xlLinkTypeExcelLinks   = 1

function Write-Step { param([string]$Message) Write-Host "  $Message" }

function Resolve-SiteWorkbook {
    <#
        事業所名からファイルを 1 つ特定する。
        「九条」が「西九条」に部分一致してしまうため、括弧で囲まれた形
        （例：シフト管理表(九条).xlsx）でのみ一致させる。
    #>
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$Site
    )

    $bracketed = '[(（]' + [regex]::Escape($Site) + '[)）]'
    $matched = @(
        Get-ChildItem -LiteralPath $Root -Filter '*.xlsx' -File |
            Where-Object { $_.Name -notlike '~$*' -and $_.BaseName -match $bracketed }
    )

    if ($matched.Count -eq 0) {
        throw "「$Site」のファイルが見つかりません。フォルダ：$Root"
    }
    if ($matched.Count -gt 1) {
        $names = ($matched | ForEach-Object { $_.Name }) -join ' / '
        throw "「$Site」に一致するファイルが複数あります：$names"
    }

    $matched[0]
}

function Release-Com {
    param($ComObject)
    if ($null -ne $ComObject) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) } catch { }
    }
}

# ---------------------------------------------------------------- 事前チェック

if (-not (Test-Path -LiteralPath $Root)) {
    throw "フォルダにアクセスできません：$Root`n共有フォルダに接続できているか確認してください。"
}

Write-Host ''
Write-Host '=== 監査勤務形態一覧 ６ブック統合 ==='
Write-Host "元フォルダ：$Root"
Write-Host "統合順　　：$($Order -join ' → ')"
Write-Host ''

Write-Host '[1/4] 元ファイルを確認しています...'
$sources = foreach ($site in $Order) {
    $file = Resolve-SiteWorkbook -Root $Root -Site $site
    Write-Step "$site : $($file.Name)"
    [pscustomobject]@{ Site = $site; File = $file }
}

if (-not $OutFile) {
    $stamp   = Get-Date -Format 'yyyyMMdd_HHmm'
    $OutFile = Join-Path $Root "統合_監査勤務形態一覧_$stamp.xlsx"
}
$outDir = Split-Path -Parent $OutFile
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
    throw "出力先フォルダがありません：$outDir"
}
if ((Test-Path -LiteralPath $OutFile) -and -not $Force) {
    throw "出力先が既に存在します：$OutFile`n上書きする場合は -Force を付けて実行してください。"
}

# ---------------------------------------------------------------- 統合処理

Write-Host ''
Write-Host '[2/4] Excel を起動しています...'

$excel = New-Object -ComObject Excel.Application
$excel.Visible          = $false
$excel.DisplayAlerts    = $false
$excel.AskToUpdateLinks = $false
$excel.EnableEvents     = $false
$excel.ScreenUpdating   = $false

$dest    = $null
$copied  = New-Object System.Collections.Generic.List[object]
$skipped = New-Object System.Collections.Generic.List[string]

try {
    $dest = $excel.Workbooks.Add()

    # Application.Calculation はブックが開いた後でないと設定できない
    $excel.Calculation = $xlCalculationManual

    # 追加直後の既定シートは最後にまとめて削除する
    $placeholders = @($dest.Worksheets | ForEach-Object { $_.Name })

    Write-Host ''
    Write-Host '[3/4] シートをコピーしています...'

    foreach ($entry in $sources) {
        $site = $entry.Site
        $path = $entry.File.FullName
        Write-Step "$site を開いています..."

        $src = $excel.Workbooks.Open($path, 0, $true)   # UpdateLinks=0, ReadOnly=True
        try {
            foreach ($ws in $src.Worksheets) {
                if ($SkipSheets -contains $ws.Name) {
                    $skipped.Add("$site / $($ws.Name)")
                    continue
                }

                $ws.Copy([System.Type]::Missing, $dest.Sheets($dest.Sheets.Count))

                # コピー先は常に末尾。名前が衝突した場合 Excel が自動で改名するので実名を拾う
                $new = $dest.Sheets($dest.Sheets.Count)
                $copied.Add([pscustomobject]@{
                    Site      = $site
                    SheetName = $new.Name
                    Source    = $entry.File.Name
                })
            }
            Write-Step "  → $(@($copied | Where-Object Site -eq $site).Count) シート"
        }
        finally {
            $src.Close($false)
            Release-Com $src
        }
    }

    if ($copied.Count -eq 0) {
        throw 'コピーできるシートが 1 枚もありませんでした。'
    }

    # 目次シート
    if (-not $NoIndexSheet) {
        Write-Step '目次シートを作成しています...'
        $idx = $dest.Worksheets.Add($dest.Sheets(1))

        # 取り込んだシートに「目次」があると衝突するので空き名を探す
        $idxName = '目次'
        $n = 2
        while ($copied.SheetName -contains $idxName) {
            $idxName = "目次$n"
            $n++
        }
        $idx.Name = $idxName

        $idx.Range('A1').Value2 = '監査勤務形態一覧　統合ブック'
        $idx.Range('A1').Font.Bold = $true
        $idx.Range('A1').Font.Size = 14
        $idx.Range('A2').Value2 = "作成日時：$(Get-Date -Format 'yyyy/MM/dd HH:mm')"

        $header = @('No', '事業所', 'シート名', '元ファイル')
        for ($c = 0; $c -lt $header.Count; $c++) {
            $cell = $idx.Cells(4, $c + 1)
            $cell.Value2    = $header[$c]
            $cell.Font.Bold = $true
        }

        $row = 5
        for ($i = 0; $i -lt $copied.Count; $i++) {
            $item = $copied[$i]
            $idx.Cells($row, 1).Value2 = $i + 1
            $idx.Cells($row, 2).Value2 = $item.Site
            $idx.Cells($row, 4).Value2 = $item.Source

            $link = $idx.Cells($row, 3)
            [void]$idx.Hyperlinks.Add(
                $link,
                '',
                "'$($item.SheetName)'!A1",
                $item.SheetName,
                $item.SheetName
            )
            $row++
        }

        $idx.Columns('A:D').AutoFit() | Out-Null
    }

    # 既定の空シートを削除
    foreach ($name in $placeholders) {
        try { $dest.Worksheets($name).Delete() } catch { }
    }

    $dest.Sheets(1).Activate()

    Write-Host ''
    Write-Host '[4/4] 保存しています...'
    $dest.SaveAs($OutFile, $xlOpenXMLWorkbook)

    # 他ブックへのリンクが残っていないか確認
    $links = $dest.LinkSources($xlLinkTypeExcelLinks)
    $externalLinks = @()
    if ($links) { $externalLinks = @($links) }

    Write-Host ''
    Write-Host '=== 完了 ==='
    Write-Host "出力ファイル：$OutFile"
    Write-Host "統合シート数：$($copied.Count)"
    foreach ($site in $Order) {
        $n = @($copied | Where-Object Site -eq $site).Count
        Write-Host "  $site : $n シート"
    }
    if ($skipped.Count -gt 0) {
        Write-Host "除外したシート（$($skipped.Count) 枚）：$($skipped -join ', ')"
    }
    if ($externalLinks.Count -gt 0) {
        Write-Host ''
        Write-Warning '統合ブックに元ファイルへの外部リンクが残っています：'
        foreach ($l in $externalLinks) { Write-Host "  $l" }
        Write-Host '（元ファイルを移動・削除すると値が更新できなくなります）'
    }
}
finally {
    # Calculation はブックが開いているうちに戻す
    try { $excel.Calculation = $xlCalculationAutomatic } catch { }

    if ($dest -and -not $KeepExcelOpen) {
        try { $dest.Close($false) } catch { }
    }
    Release-Com $dest

    $excel.ScreenUpdating = $true
    $excel.EnableEvents   = $true
    $excel.DisplayAlerts  = $true

    if ($KeepExcelOpen) {
        $excel.Visible = $true
    }
    else {
        try { $excel.Quit() } catch { }
    }
    Release-Com $excel

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
