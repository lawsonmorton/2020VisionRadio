# --- ATTEMPT 3: ANTHROPIC CLAUDE (FALLBACK) ---
if not script_content and ANTHROPIC_API_KEY:
  try:
    print("Switching to Fallback Provider 2: Anthropic Claude...")
    import anthropic

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # Try updated models to avoid 404 / EOL errors
    try:
      response = client.messages.create(
          model="claude-3-7-sonnet-20250219",
          max_tokens=8000,
          messages=[{"role": "user", "content": system_prompt}],
      )
    except Exception:
      response = client.messages.create(
          model="claude-3-5-haiku-20241022",
          max_tokens=8000,
          messages=[{"role": "user", "content": system_prompt}],
      )

    script_content = response.content[0].text
    print("SUCCESS: Generated via Anthropic Claude!")
  except Exception as e:
    print(f"Anthropic API failed: {e}")