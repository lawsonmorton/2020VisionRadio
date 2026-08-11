# --- ATTEMPT 3: ANTHROPIC CLAUDE (FALLBACK) ---
if not script_content and ANTHROPIC_API_KEY:
  try:
    print("Switching to Fallback Provider 2: Anthropic Claude...")
    import anthropic

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # Expanded list of model strings across all API tiers
    anthropic_models = [
        "claude-3-5-sonnet-latest",
        "claude-3-5-sonnet-20241022",
        "claude-3-haiku-20240307",
        "claude-3-5-haiku-latest",
        "claude-3-opus-20240229",
    ]

    for model_name in anthropic_models:
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