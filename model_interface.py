import google.generativeai as genai
from dotenv import load_dotenv
import os

load_dotenv()  # Load variables from .env file
API_KEY = os.getenv("GOOGLE_API_KEY")

genai.configure(api_key=API_KEY)  # configure API key

def generate_vba_macro(task_description: str) -> str:
    prompt = (
        "You are a VBA expert. Write ONLY the complete VBA macro (no explanation) "
        "to perform this task:\n\n" + task_description
    )
    model = genai.GenerativeModel('gemini-1.5-flash')  # specify Gemini model
    chat = model.start_chat()  # start chat session
    response = chat.send_message(prompt)  # send user prompt and get response
    return response.text  # get generated output text

