# --- ATTEMPT 3: ANTHROPIC CLAUDE (FALLBACK) ---
if not script_content and ANTHROPIC_API_KEY:
  try:
    print("Switching to Fallback Provider 2: Anthropic Claude...")
    import anthropic

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # Try 3.5 Haiku first, then 3 Haiku for maximum compatibility across all API tiers
    try:
      response = client.messages.create(
          model="claude-3-5-haiku-latest",
          max_tokens=8000,
          messages=[{"role": "user", "content": system_prompt}],
      )
    except Exception:
      response = client.messages.create(
          model="claude-3-haiku-20240307",
          max_tokens=4000,
          messages=[{"role": "user", "content": system_prompt}],
      )

    script_content = response.content[0].text
    print("SUCCESS: Generated via Anthropic Claude!")
  except Exception as e:
    print(f"Anthropic API failed: {e}")