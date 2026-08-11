import datetime
import json
import os
import re
import smtplib

# 1. Environment variables
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")

# 2. INITIALIZE script_content HERE (Before any provider blocks)
script_content = None

# 3. Provider execution flow
if GEMINI_API_KEY:
  try:
    print("Attempting Primary Provider: Google Gemini...")
    # ... execution logic ...
    script_content = response.text
  except Exception as e:
    print(f"Gemini API failed: {e}")

if not script_content and OPENAI_API_KEY:
  try:
    print("Switching to Fallback Provider 1: OpenAI...")
    # ... execution logic ...
    script_content = response.choices[0].message.content
  except Exception as e:
    print(f"OpenAI API failed: {e}")

if not script_content and ANTHROPIC_API_KEY:
  try:
    print("Switching to Fallback Provider 2: Anthropic Claude...")
    # ... execution logic ...
    script_content = response.content[0].text
  except Exception as e:
    print(f"Anthropic API failed: {e}")

if not script_content:
  print("CRITICAL ERROR: All AI providers failed or credentials missing.")
  exit(1)