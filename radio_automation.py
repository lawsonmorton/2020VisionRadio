# --- ATTEMPT 3: ANTHROPIC CLAUDE (FALLBACK) ---
if not script_content and ANTHROPIC_API_KEY:
  try:
    print("Switching to Fallback Provider 2: Anthropic Claude...")
    import anthropic

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # Try standard models compatible with basic API tiers
    for model_name in [
        "claude-3-haiku-20240307",
        "claude-3-5-sonnet-20241022",
    ]:
      try:
        response = client.messages.create(
            model=model_name,
            max_tokens=4000,
            messages=[{"role": "user", "content": system_prompt}],
        )
        script_content = response.content[0].text
        print(f"SUCCESS: Generated via Anthropic Claude ({model_name})!")
        break
      except Exception as inner_e:
        print(f"Anthropic model {model_name} failed: {inner_e}")

  except Exception as e:
    print(f"Anthropic API failed: {e}")