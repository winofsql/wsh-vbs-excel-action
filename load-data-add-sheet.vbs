Dim WshShell : Set WshShell = CreateObject("WScript.Shell")
Dim ExcelApp : Set ExcelApp = CreateObject("Excel.Application")

Dim workbook
Dim worksheet1
Dim worksheet2

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

' 現在のシート
Set worksheet1 = workbook.Sheets(1)

' シートを追加
Set worksheet2 = workbook.Worksheets.Add()

worksheet2.Name = "あたらしいシート" + CStr( workbook.Worksheets.Count )

For I = 1  to 11
	For J = 1 to 51
		worksheet2.Cells(J, I) = worksheet1.Cells(J, I)
	Next
Next

' 名前を付けて別ファイルに保存
workbook.SaveAs( CurDir + "\syain2.xlsx" )

' 閉じる
workbook.Close()

' アプリを終了
ExcelApp.Quit()

' オブジェクト初期化( 環境によってはプロセスが残るかもしれない )
Set ExcelApp = Nothing
ExcelApp = ""

' 実行中の Excel を全て強制終了
WshShell.Run("taskkill /F /IM excel.exe")
