# --- ATTEMPT 1: GOOGLE GEMINI ---
if GEMINI_API_KEY:
  try:
    print("Attempting Primary Provider: Google Gemini...")
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=GEMINI_API_KEY)

    # Updated model strings compatible with google-genai SDK
    for g_model in ["gemini-2.0-flash", "gemini-1.5-flash-latest"]:
      try:
        response = client.models.generate_content(
            model=g_model,
            contents=system_prompt,
            config=types.GenerateContentConfig(tools=[{"google_search": {}}]),
        )
        if response and hasattr(response, "text") and response.text:
          script_content = response.text
          print(f"SUCCESS: Generated via Google Gemini ({g_model})!")
          break
      except Exception as g_err:
        print(f"Gemini {g_model} failed: {g_err}")

  except Exception as e:
    print(f"Gemini API initialization failed: {e}")