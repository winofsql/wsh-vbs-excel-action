Dim WshShell : Set WshShell = CreateObject("WScript.Shell")
Dim ExcelApp : Set ExcelApp = CreateObject("Excel.Application")

Dim workbook

' �\������
ExcelApp.Visible = True
' �x�����o���Ȃ��悤�ɂ���
ExcelApp.DisplayAlerts = False

Dim CurDir
CurDir = WshShell.CurrentDirectory

' �J��( utf8�̂����݂����ŁA�\�[�X���� ��(����͑S�p) ���g���ƃG���[�ɂȂ� )
Set workbook = ExcelApp.Workbooks.Open( CurDir + "\syain.xlsx" )
' �ő剻
ExcelApp.ActiveWindow.WindowState = -4137

' �ۑ��������ɂ���
workbook.Saved = True

' ����
workbook.Close()

' �A�v�����I��
ExcelApp.Quit()

' �I�u�W�F�N�g������( ���ɂ���Ă̓v���Z�X���c�邩������Ȃ� )
Set ExcelApp = Nothing
ExcelApp = ""

' ���s���� Excel ��S�ċ����I��
WshShell.Run("taskkill /F /IM excel.exe")
