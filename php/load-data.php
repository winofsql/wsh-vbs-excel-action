<?php

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
$workook->Saved = true;

// アプリを終了
$ExcelApp->Quit();

// オブジェクト初期化
$ExcelApp = null;

?>
