import os
import win32com.client
from openpyxl import Workbook
import pythoncom
import time

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
    pythoncom.CoInitialize()  # Important for COM threading
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False

    try:
        wb = excel.Workbooks.Open(path)
        module = wb.VBProject.VBComponents.Add(1)
        module.CodeModule.AddFromString(vba_code)
        wb.Save()
        wb.Close(SaveChanges=False)  # ❗ keep this False unless you want to prompt Save dialog
    except Exception as e:
        print(f"[ERROR] Failed to inject macro: {e}")
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()
        time.sleep(1)  # gives Excel time to fully shut down
