import re

def clean_vba_response(response_text: str) -> str:
    """
    Cleans Gemini response by removing Markdown code fences like ```vba and ``` 
    and returns only the macro code.
    """
    lines = response_text.strip().splitlines()
    cleaned_lines = [
        line for line in lines
        if not re.match(r"^```(?:vba)?$", line.strip(), re.IGNORECASE)
    ]
    return "\n".join(cleaned_lines).strip()