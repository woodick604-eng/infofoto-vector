import os
import google.generativeai as genai

api_key = os.environ.get("GEMINI_API_KEY", "AIzaSyAPYtwi9oS91U9P0fniY4jS1CS2vtKAI7U")
genai.configure(api_key=api_key)

try:
    for m in genai.list_models():
        print(f"Model: {m.name} - Capabilities: {m.supported_generation_methods}")
except Exception as e:
    print(f"Error listing models: {e}")
