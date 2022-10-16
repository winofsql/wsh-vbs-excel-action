# powershell .\load-data.ps1

Add-Type -Assembly System.Windows.Forms

$ExcelApp = New-Object -ComObject Excel.Application

# �\������
$ExcelApp.Visible = $true
# �x�����o���Ȃ��悤�ɂ���
$ExcelApp.DisplayAlerts = $false

# �J��
$workbook = $ExcelApp.Workbooks.Open( (Convert-Path .) + "\syain.xlsx" )
# �ő剻
$ExcelApp.ActiveWindow.WindowState = -4137

[System.Windows.Forms.MessageBox]::Show("STOP", "�^�C�g��")

# �ۑ��������ɂ���
$workbook.Saved = $true

# �A�v�����I��
$ExcelApp.Quit()

# �I�u�W�F�N�g������
$ExcelApp = $null

