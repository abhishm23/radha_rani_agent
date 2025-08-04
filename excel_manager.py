import os
import win32com.client
from openpyxl import Workbook

def create_xlsm(path: str):
    folder = os.path.dirname(path)
    os.makedirs(folder, exist_ok=True)

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Add()

    # Save as .xlsm (52 = xlOpenXMLWorkbookMacroEnabled)
    wb.SaveAs(Filename=path, FileFormat=52)
    wb.Close(False)
    excel.Quit()

    return path

def inject_macro(path: str, vba_code: str):
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(path)
    module = wb.VBProject.VBComponents.Add(1)
    module.CodeModule.AddFromString(vba_code)
    wb.Save()
    wb.Close(False)
    excel.Quit()
