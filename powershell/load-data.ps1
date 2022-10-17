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

# オブジェクトを解放
#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelApp)

# C# ではほぼ完全解放無理なので強制終了させる
foreach ($p in [System.Diagnostics.Process]::GetProcessesByName("EXCEL")) {
	if ($p.MainWindowTitle -eq "") {
		$p.Kill()
	}
}

# オブジェクト初期化
$ExcelApp = $null

