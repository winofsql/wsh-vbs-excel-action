# wsh-vbs-excel-action

### [オブジェクト モデル (Excel)](https://learn.microsoft.com/ja-jp/office/vba/api/overview/excel/object-model)

### [PHP com クラス](https://www.php.net/manual/ja/class.com.php)

### load-data.wsf (UTF-8)
```vbscript
<?xml version="1.0" encoding="utf-8" ?>
<job>
<script language="vbscript">

Dim WshShell : Set WshShell = CreateObject("WScript.Shell")
Dim ExcelApp : Set ExcelApp = CreateObject("Excel.Application")

Dim workbook

' 表示する
ExcelApp.Visible = True
' 警告を出さないようにする
ExcelApp.DisplayAlerts = False

Dim CurDir
CurDir = WshShell.CurrentDirectory

' 開く( utf8のせいみたいで、ソース内に ＆(これは全角) を使うとエラーになる )
Set workbook = ExcelApp.Workbooks.Open( CurDir + "\syain.xlsx" )
' 最大化
ExcelApp.ActiveWindow.WindowState = -4137

' 保存した事にする
workbook.Saved = True

' アプリを終了
ExcelApp.Quit()

' オブジェクト初期化( Windows10+Excel365で無くても終了している )
Set ExcelApp = Nothing
ExcelApp = ""

</script>
</job>
```

### load-data.php (SHIFT_JIS) : ※ PHP の終了( Excel の解放 には少々時間がかかります )
```php
<?php
// php.ini
// extension=php_com_dotnet.dll
$WshShell = new com("WScript.Shell");

$ExcelApp = new com("Excel.Application") or die("Unable to instantiate Excel");

// 表示する
$ExcelApp->Visible = true;
// 警告を出さないようにする
$ExcelApp->DisplayAlerts = false;

// 開く
$workbook = $ExcelApp->Workbooks->Open( getcwd() ."\\syain.xlsx" );
// 最大化
$ExcelApp->ActiveWindow->WindowState = -4137;

// 保存した事にする
$workbook->Saved = true;

// アプリを終了
$ExcelApp->Quit();

// オブジェクト初期化
$ExcelApp = null;

$WshShell = null;
?>
```

### load-data.ps1 (SHIFT_JIS)
```powershell
## powershell .\load-data.ps1

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
```

### HTA ( vbscript )
```vbscript
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=10">
<meta charset="utf-8">
<title>Excel Action</title>
<script language="VBScript">

Dim ExcelApp

Function Excel_Action()

	Dim CurDir : CurDir = location.pathname
	CurDir = Replace(CurDir, "/", "")
	Dim aData : aData = Split( CurDir, "\" )
	CurDir = Replace(CurDir, "\" & aData( Ubound(aData) ), "" )

	Set ExcelApp = CreateObject("Excel.Application")

	Dim workbook

	' 表示する
	ExcelApp.Visible = True
	' 警告を出さないようにする
	ExcelApp.DisplayAlerts = False

	' 開く
	Set workbook = ExcelApp.Workbooks.Open( CurDir & "\syain.xlsx" )
	' 最大化
	ExcelApp.ActiveWindow.WindowState = -4137

	' 保存した事にする
	workbook.Saved = True
	
	' アプリを終了
	ExcelApp.Quit()

	' オブジェクト初期化
	Set ExcelApp = Nothing
	ExcelApp = ""

End Function
</script>
</head>
<body>

<input type="button" value="実行" onclick="Call Excel_Action()" language="VBScript">

</body>
</html>
```

### HTA ( javascript )
```javascript
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="ie=edge">
<meta charset="utf-8">
<title>Excel Action</title>
<script>

var ExcelApp;

function Excel_Action() {

	var CurDir = location.pathname;
	CurDir = CurDir.replace("/", "");
	var aData = CurDir.split( "\\" );
	CurDir = CurDir.replace( "\\" + aData[ aData.length - 1 ], "" );
	
	ExcelApp = new ActiveXObject("Excel.Application");

	var workbook;

	// 表示する
	ExcelApp.Visible = true;
	// 警告を出さないようにする
	ExcelApp.DisplayAlerts = false;

	// 開く
	workbook = ExcelApp.Workbooks.Open( CurDir + "\\syain.xlsx" );
	// 最大化
	ExcelApp.ActiveWindow.WindowState = -4137;

	// 保存した事にする
	workbook.Saved = true;
	
	// アプリを終了
	ExcelQuit();

}

// ****************************
// JavaScript による
// Excel の完全終了処理
// ****************************
var idTmr = "";
function ExcelQuit() {
	ExcelApp.Quit();
	ExcelApp = null;
	idTmr = window.setTimeout("Cleanup();",1);
}
function Cleanup() {
	window.clearInterval(idTmr);
	CollectGarbage();
}
</script>
</head>
<body>

<input type="button" value="実行" onclick="Excel_Action()">

</body>
</html>
```

### Python
```python
# python -m pip install --upgrade pip
# pip install pywin32

import win32com.client
import os
from tkinter import messagebox

curdir = os.getcwd()

ExcelApp = win32com.client.Dispatch("Excel.Application")

ExcelApp.Visible = True

ExcelApp.DisplayAlerts = False

workbook = ExcelApp.Workbooks.Open( curdir + "\\syain.xlsx" )

ExcelApp.ActiveWindow.WindowState = -4137

workbook.Saved = True

ExcelApp.Quit()
```

## C# Dotnet 6.0
```cs
using System;
using System.Windows.Forms;
using System.Diagnostics;

namespace load_data
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var path = $@"{Environment.CurrentDirectory}\syain.xlsx";

            // Excel アプリケーション
            dynamic excelApp =
                Activator
                    .CreateInstance(Type
                        .GetTypeFromProgID("Excel.Application"));

            // Excel を表示( 完成したらコメント化 )
            excelApp.Visible = true;

            // 警告を出さない
            excelApp.DisplayAlerts = false;

            // Excel ブック( 既存 )
            dynamic workBook = excelApp.Workbooks.Open(path);

            // 最大化
            excelApp.ActiveWindow.WindowState = -4137;

            MessageBox.Show("STOP");

            // 保存した事にする
            workBook.Saved = true;

            // 終了
            excelApp.Quit();

            // 解放
            System.Runtime.InteropServices.Marshal.ReleaseComObject (excelApp);

            // C# ではほぼ完全解放無理なので強制終了させる
            foreach (var p in Process.GetProcessesByName("EXCEL"))
            {
                if (p.MainWindowTitle == "")
                {
                    p.Kill();
                }
            }
            // dotnet.exe が常駐すると他のコマンドアプリの excel 解放処理を邪魔する
            foreach (var p in Process.GetProcessesByName("dotnet"))
            {
                if (p.MainWindowTitle == "")
                {
                    p.Kill();
                }
            }

        }
    }
}
```
