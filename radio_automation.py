# --- ATTEMPT 1: GOOGLE GEMINI ---
if GEMINI_API_KEY:
  try:
    print("Attempting Primary Provider: Google Gemini...")
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=GEMINI_API_KEY)

    # Try 2.0 Flash first, fallback to 1.5 Flash if 2.0 quota is 0
    for g_model in ["gemini-2.0-flash", "gemini-1.5-flash"]:
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

# --- ATTEMPT 2: OPENAI CHATGPT (FALLBACK) ---
if not script_content and OPENAI_API_KEY:
  try:
    print("Switching to Fallback Provider 1: OpenAI ChatGPT...")
    from openai import OpenAI

    client = OpenAI(api_key=OPENAI_API_KEY)
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": system_prompt}],
    )
    if (
        response
        and response.choices
        and response.choices[0].message
        and response.choices[0].message.content
    ):
      script_content = response.choices[0].message.content
      print("SUCCESS: Generated via OpenAI ChatGPT!")
  except Exception as e:
    print(f"OpenAI API failed: {e}")

# --- ATTEMPT 3: ANTHROPIC CLAUDE (FALLBACK) ---
if not script_content and ANTHROPIC_API_KEY:
  try:
    print("Switching to Fallback Provider 2: Anthropic Claude...")
    import anthropic

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    for model_name in [
        "claude-3-haiku-20240307",
        "claude-3-5-sonnet-20241022",
        "claude-3-sonnet-20240229",
    ]:
      try:
        response = client.messages.create(
            model=model_name,
            max_tokens=4000,
            messages=[{"role": "user", "content": system_prompt}],
        )
        if response and response.content and response.content[0].text:
          script_content = response.content[0].text
          print(f"SUCCESS: Generated via Anthropic Claude ({model_name})!")
          break
      except Exception as inner_e:
        print(f"Anthropic model {model_name} failed: {inner_e}")

  except Exception as e:
    print(f"Anthropic API failed: {e}")