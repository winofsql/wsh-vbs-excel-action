Dim WshShell : Set WshShell = CreateObject("WScript.Shell")
Dim ExcelApp : Set ExcelApp = CreateObject("Excel.Application")

Dim workbook
Dim worksheet1
Dim worksheet2

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

' ���݂̃V�[�g
Set worksheet1 = workbook.Sheets(1)

' �V�[�g��ǉ�
Set worksheet2 = workbook.Worksheets.Add()

worksheet2.Name = "�����炵���V�[�g" + CStr( workbook.Worksheets.Count )

For I = 1  to 11
	For J = 1 to 51
		worksheet2.Cells(J, I) = worksheet1.Cells(J, I)
	Next
Next

' ���O��t���ĕʃt�@�C���ɕۑ�
workbook.SaveAs( CurDir + "\syain2.xlsx" )

' ����
workbook.Close()

' �A�v�����I��
ExcelApp.Quit()

' �I�u�W�F�N�g������( ���ɂ���Ă̓v���Z�X���c�邩������Ȃ� )
Set ExcelApp = Nothing
ExcelApp = ""

' ���s���� Excel ��S�ċ����I��
WshShell.Run("taskkill /F /IM excel.exe")
