from model_interface import generate_vba_macro

def request_macro_fix(original_task: str, failed_code: str, error_message: str) -> str:
    """
    Ask the Gemini model to regenerate a corrected macro based on error feedback.
    """
    prompt = f"""
You are a senior VBA developer. The following macro failed to run due to this error:

--- ERROR MESSAGE ---
{error_message}

--- ORIGINAL MACRO ---
{failed_code}

Please regenerate the complete corrected VBA macro to fulfill the task:
{original_task}

Only return the full corrected VBA macro code, no explanation or instructions.
"""
    return generate_vba_macro(prompt)