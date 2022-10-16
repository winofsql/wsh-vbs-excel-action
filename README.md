# wsh-vbs-excel-action

### [オブジェクト モデル (Excel)](https://learn.microsoft.com/ja-jp/office/vba/api/overview/excel/object-model)

### [PHP com クラス](https://www.php.net/manual/ja/class.com.php)

### load-data.wsf
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

MsgBox("STOP")

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
