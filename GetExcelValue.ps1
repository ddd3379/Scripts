# 読み取るExcelの情報を変数に設定する
$excelname = "C:\temp\test.xlsx"
$sheetname = "Table"

# 変数を定義する
$excel = $null
$workbook = $null
$worksheet = $null

$excelSub = $null
$workbookSub = $null
$worksheeSub = $null

try
{
    # ExcelのCOMオブジェクトを取得する
    $excel = New-Object -ComObject Excel.Application
    # Excelの動作の設定をする
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    # Excelファイルの指定したシートのセルの値を取得する
    $workbook = $excel.Workbooks.Open($excelname)
    $worksheet = $workbook.Sheets($sheetname)

    $line = 1
    while($true){
        # データ取得
        $tcd = $worksheet.Cells.Item($line, 2).Text

        ##################
        ## Get orignal data
        $value01 = $worksheet.Cells.Item($line, 3).Text
        $value02 = $worksheet.Cells.Item($line, 4).Text
        ##################

        # データが無くなれば終了
        if ($tcd -eq "") { break }

        $folderList = Get-ChildItem "C:\temp\$tcd*" | Where-Object { $_.PSIsContainer }
        if ($folderList.Count -ne 1){
            Write-Host "folderList size is not 1."
            $line++
            continue
        }
        $folderName = "C:\temp\"+$folderList[0].Name+"\temp"
        Write-Host "=== $tcd($folderName) ==="


        # ファイル名の取得
        $filePath = "$folderName\$tcd*.xlsx"
        $fileList = Get-ChildItem $filePath | Where-Object { ! $_.PSIsContainer }
        if ($fileList.Count -ne 1) { 
            Write-Host "fileList size is not 1."
            $line++
            continue
        }
        Write-Host $fileList[0].Name

        # ExcelのCOMオブジェクトを取得する
        $excelSub = New-Object -ComObject Excel.Application
        # Excelの動作の設定をする
        $excelSub.Visible = $false
        $excelSub.DisplayAlerts = $false
        # Excelファイルの指定したシートのセルの値を取得する
        $workbookSub = $excelSub.Workbooks.Open("$folderName\"+$fileList[0].Name)
        $worksheetSub = $workbookSub.Sheets("Table")

        ##################
        ## Get sub data
        $value11 = $worksheetSub.Range("A2").Text
        $value12 = $worksheetSub.Range("B3").Text
        ##################

        ##################
        ## Output both data
        Write-Host "Original`t|`tSub"
        Write-Host "$value01`t|`t$value11"
        Write-Host "$value02`t|`t$value12"
        ##################

        $line++
    }
}
catch
{
    # エラーメッセージを表示する
    Write-Error("Error"+$_.Exception)
}
finally
{
    # COMオブジェクトを開放する
    # ※ReleaseComObjectは戻り値0を返してくるんだけど
    #   コンソールに0が表示されるので変数で受け取って表示しないようにする
    if($worksheet -ne $null)
    {
        # ワークシートを破棄する
        $result = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet)
    }
    if($workbook -ne $null)
    {
        # Excelファイルを閉じる
        $workbook.Close($false)
        $result = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if($excel -ne $null)
    {
        # Excelを閉じる
        $excel.Quit()
        $result = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }


    #   コンソールに0が表示されるので変数で受け取って表示しないようにする
    if($worksheetSub -ne $null)
    {
        # ワークシートを破棄する
        $result = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheetSub)
    }
    if($workbookSub -ne $null)
    {
        # Excelファイルを閉じる
        $workbookSub.Close($false)
        $result = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbookSub)
    }
    if($excelSub -ne $null)
    {
        # Excelを閉じる
        $excelSub.Quit()
        $result = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelSub)
    }
}
# 実行はおしまい
Write-Host "finished!"
