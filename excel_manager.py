import os
import win32com.client
from openpyxl import Workbook

def create_xlsm(path: str):
    folder = os.path.dirname(path)
    os.makedirs(folder, exist_ok=True)
    wb = Workbook()
    wb.save(path)
    if not path.lower().endswith(".xlsm"):
        base, _ = os.path.splitext(path)
        os.rename(path, base + ".xlsm")
        path = base + ".xlsm"
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
