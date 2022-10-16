# powershell .\load-data.ps1

Add-Type -Assembly System.Windows.Forms

$ExcelApp = New-Object -ComObject Excel.Application

# 表示する
$ExcelApp.Visible = $true
# 警告を出さないようにする
$ExcelApp.DisplayAlerts = $false

# 開く
$workbook = $ExcelApp.Workbooks.Open( (Convert-Path .) + "\syain.xlsx" )
# 最大化
$ExcelApp.ActiveWindow.WindowState = -4137

[System.Windows.Forms.MessageBox]::Show("STOP", "タイトル")

# 保存した事にする
$workbook.Saved = $true

# アプリを終了
$ExcelApp.Quit()

# オブジェクト初期化
$ExcelApp = $null

