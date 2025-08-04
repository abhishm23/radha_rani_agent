from model_interface import generate_vba_macro
from excel_manager import create_xlsm, inject_macro
from run_macro_and_validate import run_macro
from error_feedback_loop import request_macro_fix

def main():
    task = input("Describe your macro task:\n").strip()
    file_path = input("Enter full path for .xlsm file to create/modify:\n").strip()

    xlsm = create_xlsm(file_path)

    max_attempts = 5
    attempt = 1
    current_macro = generate_vba_macro(task)

    while attempt <= max_attempts:
        print(f"\nAttempt {attempt} of {max_attempts}...")

        print("Injecting macro...")
        inject_macro(xlsm, current_macro)

        print("Running macro...")
        success, error = run_macro(xlsm, current_macro)

        if success:
            print("Macro executed successfully.")
            break
        else:
            print("Macro failed with error:")
            print(error)

            if attempt == max_attempts:
                print("Max attempts reached. Macro is still failing.")
                break

            print("Asking model to fix the macro based on error...")
            current_macro = request_macro_fix(task, current_macro, error or "")

        attempt += 1

if __name__ == "__main__":
    main()
