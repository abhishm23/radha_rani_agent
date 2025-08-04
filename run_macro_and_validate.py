import win32com.client
import re
import traceback
from typing import Optional

def extract_macro_name(vba_code: str) -> Optional[str]:
    match = re.search(r"Sub\s+(\w+)\s*\(", vba_code, re.IGNORECASE)
    if match:
        return match.group(1)
    return None

def run_macro(file_path: str, vba_code: str):
    try:
        macro_name = extract_macro_name(vba_code)
        if not macro_name:
            return False, "Could not extract macro name from VBA code."

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False

        wb = excel.Workbooks.Open(file_path)
        try:
            excel.Run(f"'{wb.Name}'!{macro_name}")
        except Exception as run_err:
            return False, f"Macro execution failed: {run_err}"

        wb.Close(SaveChanges=False)
        excel.Quit()
        return True, None

    except Exception as e:
        return False, f"Unexpected error:\n{traceback.format_exc()}"