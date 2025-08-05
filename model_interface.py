import google.generativeai as genai
from dotenv import load_dotenv
import os
from utils.cleaner import clean_vba_response

load_dotenv()  # Load variables from .env file
API_KEY = os.getenv("GOOGLE_API_KEY")

genai.configure(api_key=API_KEY)  # configure API key

def generate_vba_macro(task_description: str) -> str:
    prompt = (
        "You are a VBA expert. Write ONLY the complete VBA macro (no explanation).\n"
        "Make sure the macro runs fully in the background, handling all pop-ups silently.\n"
        "Include clear, simple, and understandable comments for every line of code.\n"
        "Always save and close all workbooks involved in the process at the end.\n"
        "Add a generic 'Task completed' message box at the end.\n"
        "Keep the code clean and easy to understand.\n\n"

        "Strict file handling rules (Do not break these):\n"
        "1. The file where this macro is stored (e.g., 'output.xlsm') is ONLY used to host and run the macro.\n"
        "2. Do NOT use it as a data source or target for any copy-paste or sheet operations.\n"
        "3. If only one file is mentioned, assume it is both source and destination.\n"
        "4. If two files are mentioned, treat them strictly as separate: source vs destination.\n"
        "5. Always hardcode the file names unless instructed otherwise\n"
        "6. Do NOT use ThisWorkbook or ActiveWorkbook unless explicitly stated.\n"
        "7. if only one file is being used save and close it once only at the end.\n\n"

        "Now write the complete macro to perform the following task:\n\n" + task_description
    )

    model = genai.GenerativeModel('gemini-1.5-flash')  # specify Gemini model
    chat = model.start_chat()  # start chat session
    response = chat.send_message(prompt)  # send user prompt and get response
    return clean_vba_response(response.text)  # get generated output text

