<?php
// php.ini
// extension=php_com_dotnet.dll
$WshShell = new com("WScript.Shell");

$ExcelApp = new com("Excel.Application") or die("Unable to instantiate Excel");

// �\������
$ExcelApp->Visible = true;
// �x�����o���Ȃ��悤�ɂ���
$ExcelApp->DisplayAlerts = false;

// �J��
$workbook = $ExcelApp->Workbooks->Open( getcwd() ."\\syain.xlsx" );
// �ő剻
$ExcelApp->ActiveWindow->WindowState = -4137;

$WshShell->Popup("STOP");

// �ۑ��������ɂ���
// $workook->Saved = true;

// �A�v�����I��
$ExcelApp->Quit();

// �I�u�W�F�N�g������
$ExcelApp = null;

?>
