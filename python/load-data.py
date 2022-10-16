import win32com.client
import os
from tkinter import messagebox

curdir = os.getcwd()

ExcelApp = win32com.client.Dispatch("Excel.Application")

ExcelApp.Visible = True

ExcelApp.DisplayAlerts = False

workbook = ExcelApp.Workbooks.Open( curdir + "\\syain.xlsx" )

ExcelApp.ActiveWindow.WindowState = -4137

messagebox.showinfo("タイトル", "STOP")

workbook.Saved = True

ExcelApp.Quit()

