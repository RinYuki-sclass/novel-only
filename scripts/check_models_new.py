import os
from dotenv import load_dotenv
from google import genai
from google.genai.errors import APIError

load_dotenv()

keys = [v for k, v in os.environ.items() if k.startswith('GEMINI_API_KEY_')]

with open('scripts/available_models.txt', 'w', encoding='utf-8') as f:
    for i, api_key in enumerate(keys):
        print(f"\n--- Checking API Key {i+1} ---")
        try:
            client = genai.Client(api_key=api_key)
            models = list(client.models.list())
            print(f"API Key {i+1} works! Models found: {len(models)}")
            f.write(f"\n=== Models for API Key {i+1} ===\n")
            for m in models:
                line = f"Name: {m.name} | Display Name: {m.display_name}\n"
                f.write(line)
        except APIError as e:
            print(f"API Key {i+1} failed with APIError: {e}")
        except Exception as e:
            print(f"API Key {i+1} failed with Exception: {e}")
