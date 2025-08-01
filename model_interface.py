import google.generativeai as palm

API_KEY = "AIzaSyCoANaKvTypu-avUqHleLpqm-sJRkb5mVQ"

palm.configure(api_key=API_KEY)

def generate_vba_macro(task_description: str) -> str:
    prompt = (
        "You are a VBA expert. Write ONLY the complete VBA macro (no explanation) "
        "to perform this task:\n\n" + task_description
    )
    response = palm.chat(messages=[{"author": "user", "content": prompt}])
    return response.last