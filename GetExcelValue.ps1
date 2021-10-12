# �ǂݎ��Excel�̏���ϐ��ɐݒ肷��
$excelname = "C:\temp\test.xlsx"
$sheetname = "Table"

# �ϐ����`����
$excel = $null
$workbook = $null
$worksheet = $null

$excelSub = $null
$workbookSub = $null
$worksheeSub = $null

try
{
    # Excel��COM�I�u�W�F�N�g���擾����
    $excel = New-Object -ComObject Excel.Application
    # Excel�̓���̐ݒ������
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    # Excel�t�@�C���̎w�肵���V�[�g�̃Z���̒l���擾����
    $workbook = $excel.Workbooks.Open($excelname)
    $worksheet = $workbook.Sheets($sheetname)

    $line = 1
    while($true){
        # �f�[�^�擾
        $tcd = $worksheet.Cells.Item($line, 2).Text

        ##################
        ## Get orignal data
        $value01 = $worksheet.Cells.Item($line, 3).Text
        $value02 = $worksheet.Cells.Item($line, 4).Text
        ##################

        # �f�[�^�������Ȃ�ΏI��
        if ($tcd -eq "") { break }

        $folderList = Get-ChildItem "C:\temp\$tcd*" | Where-Object { $_.PSIsContainer }
        if ($folderList.Count -ne 1){
            Write-Host "folderList size is not 1."
            $line++
            continue
        }
        $folderName = "C:\temp\"+$folderList[0].Name+"\temp"
        Write-Host "=== $tcd($folderName) ==="


        # �t�@�C�����̎擾
        $filePath = "$folderName\$tcd*.xlsx"
        $fileList = Get-ChildItem $filePath | Where-Object { ! $_.PSIsContainer }
        if ($fileList.Count -ne 1) { 
            Write-Host "fileList size is not 1."
            $line++
            continue
        }
        Write-Host $fileList[0].Name

        # Excel��COM�I�u�W�F�N�g���擾����
        $excelSub = New-Object -ComObject Excel.Application
        # Excel�̓���̐ݒ������
        $excelSub.Visible = $false
        $excelSub.DisplayAlerts = $false
        # Excel�t�@�C���̎w�肵���V�[�g�̃Z���̒l���擾����
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
    # �G���[���b�Z�[�W��\������
    Write-Error("Error"+$_.Exception)
}
finally
{
    # COM�I�u�W�F�N�g���J������
    # ��ReleaseComObject�͖߂�l0��Ԃ��Ă���񂾂���
    #   �R���\�[����0���\�������̂ŕϐ��Ŏ󂯎���ĕ\�����Ȃ��悤�ɂ���
    if($worksheet -ne $null)
    {
        # ���[�N�V�[�g��j������
        $result = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet)
    }
    if($workbook -ne $null)
    {
        # Excel�t�@�C�������
        $workbook.Close($false)
        $result = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if($excel -ne $null)
    {
        # Excel�����
        $excel.Quit()
        $result = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }


    #   �R���\�[����0���\�������̂ŕϐ��Ŏ󂯎���ĕ\�����Ȃ��悤�ɂ���
    if($worksheetSub -ne $null)
    {
        # ���[�N�V�[�g��j������
        $result = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheetSub)
    }
    if($workbookSub -ne $null)
    {
        # Excel�t�@�C�������
        $workbookSub.Close($false)
        $result = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbookSub)
    }
    if($excelSub -ne $null)
    {
        # Excel�����
        $excelSub.Quit()
        $result = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelSub)
    }
}
# ���s�͂����܂�
Write-Host "finished!"
