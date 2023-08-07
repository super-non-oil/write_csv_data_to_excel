
<#
.SYNOPSIS

sample

.NOTES

require
 - powershell 7.x

.EXAMPLE

PS> .\xxx.ps1

#>

Set-StrictMode -Version Latest

[scriptblock] $now          = { (Get-Date -f "[ yyyy-MM-dd HH:mm:ss.fff ]") }
[scriptblock] $currentDate  = { (Get-Date -f "yyyyMMdd") }
[scriptblock] $currentTime  = { (Get-Date -f "HHmmss") }
[string] $logPath           = "$PSScriptRoot\log"
[string] $processLogFile    = "$logPath\$(& $currentDate)process.log"
[string] $errorLogFile      = "$logPath\$(& $currentDate)error.log"

[string] $LOG_TYPE_INFO = "info"
[string] $LOG_TYPE_DEBUG = "debug"
[string] $LOG_TYPE_ERROR = "error"

[string] $csvPath = "$PSScriptRoot\csv"
[string] $currExeDate = Get-Content "$PSScriptRoot\next_exe_date.txt"
[string] $csvName = "list$currExeDate.csv"
[string] $excelName = "sample.xlsx"

function AppendLog {
    [OutputType([System.Void])]
    param(
        [Parameter(Position=0, Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $type,
        [Parameter(Position=1, Mandatory)]
        [string] $str
    )
    Write-Output "$(& $now) : [$type] $str" | Tee-Object $processLogFile -Append
}

AppendLog $LOG_TYPE_INFO "処理開始"

try {
    $ErrorActionPreference = "stop"

    <#
        バッチのキック等
    #>

    # ログファイル用のディレクトリ無ければ作成
    if (-not(Test-Path $logPath)) { [void](New-Item $logPath -ItemType Directory -Force) }
    # 対象CSVが無ければ終了
    if (-not(Test-Path "$csvPath\$csvName")) { Write-Error "$csvPath\$csvName が存在しません。" }

    # CSV読み込み
    [Object[]] $writeData = Import-Csv -Path "$csvPath\$csvName" -Encoding UTF8
    # カラム数 = ヘッダー数
    [int] $colnum = ($writeData | Get-Member -MemberType NoteProperty).count

    [System.__ComObject] $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false


    # Excelがあるか確認
    if (-not(Test-Path "$PSScriptRoot\$excelName")) {
        # Excelを作成、ヘッダーのみ
        AppendLog $LOG_TYPE_INFO "Excelを新規作成します"
        [__ComObject] $book = $excel.Workbooks.Add()
        [__ComObject] $sheet = $book.WorkSheets(1)
        # ヘッダー書き込み
        $sheet.Name = "list"
        $sheet.Range($sheet.Cells(1, 1), $sheet.Cells(1, $colnum)) = $writeData[0].PSObject.Properties.Name
        # テーブル作成
        $tableName = "table"
        $sheet.ListObjects.Add(1, $sheet.Range($sheet.Cells(1, 1), $sheet.Cells(1, $colnum)), "", 1).Name = $tableName
        # クエリ追加
        $queryName = "query"
        $queryData = "Excel.CurrentWorkbook(){[Name=`"$tableName`"]}[Content]"
        [void] $book.Queries.Add($queryName, $queryData)
        # 保存
        $book.saveAs("$PSScriptRoot\$excelName")
        AppendLog $LOG_TYPE_INFO "Excelを新規作成しました。path : $PSScriptRoot\$excelName"
    }
    [__ComObject] $book = $excel.Workbooks.Open("$PSScriptRoot\$excelName")
    [__ComObject] $sheet = $book.Sheets(1)
    [int] $rownum = $sheet.UsedRange.Rows.Count

    # データ書き込み
    AppendLog $LOG_TYPE_INFO "Excelへの書き込み開始"
    $writeData | & { process {
        $row = $_.PSObject.Properties.value
        $rownum++
        $sheet.Range($sheet.Cells($rownum, 1), $sheet.Cells($rownum, $colnum)) = $row
        AppendLog $LOG_TYPE_DEBUG "行 $rownum $($row -join ",")"
    }}
    AppendLog $LOG_TYPE_INFO "Excelへの書き込み完了"

    # Excel保存
    $book.save()
    $excel.quit()

    # 次回実行日の更新
    $d = [DateTime]::ParseExact($currExeDate,"yyyyMMdd", $null)
    if ($d.Day -eq 1) {
        # 実行日 : 1日 ⇒ 15日
        $d = (Get-Date $d -Day 15)
    } elseif ($d.Day -eq 15) {
        # 実行日 : 15日 ⇒ 翌月1日
        $d = (Get-Date $d.AddMonths(1) -Day 1)
    }
    AppendLog $LOG_TYPE_INFO "次回実行日は $(Get-Date $d -f d)"
    Get-Date $d -f "yyyyMMdd" | Set-Content -Path "$PSScriptRoot\next_exe_date.txt"
} catch {
    Write-Output "$(& $now) : [$LOG_TYPE_ERROR] $($error)" | Tee-Object $processLogFile -Append
} finally {
    # プロセスを残さない
    [Void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    [Void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($book)
    [Void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet)
    [GC]::Collect()
}

AppendLog $LOG_TYPE_INFO "処理終了"