from model_interface import generate_vba_macro
from excel_manager import create_xlsm, inject_macro

def main():
    task = input("Describe your macro task:\n").strip()
    file_path = input("Enter full path for .xlsm file to create/modify:\n").strip()
    
    print("\nGenerating VBA macro…")
    vba = generate_vba_macro(task)
    
    print("Creating or opening workbook…")
    xlsm = create_xlsm(file_path)
    
    print("Injecting macro into workbook…")
    inject_macro(xlsm, vba)
    
    print("Done. Macro injected at:", xlsm)

if __name__ == "__main__":
    main()
