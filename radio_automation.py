import datetime
import json
import os
import re
import smtplib
import time
import traceback
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

# Environment variables
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "").strip()
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "").strip()

SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "").strip()
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD", "").strip()
PRESENTER_EMAIL_RAW = os.environ.get("PRESENTER_EMAIL", "")

recipients = [
    email.strip() for email in PRESENTER_EMAIL_RAW.split(",") if email.strip()
]


def add_hyperlink(paragraph, url, text, color="0000FF", underline=True):
  """Adds a native, fully functional XML hyperlink to a Word document."""
  part = paragraph.part
  r_id = part.relate_to(
      url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True
  )

  hyperlink = parse_xml(
      f'<w:hyperlink'
      f' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
      f' r:id="{r_id}"/>'
  )

  new_run = OxmlElement("w:r")
  rPr = OxmlElement("w:rPr")

  if color:
    c = OxmlElement("w:color")
    c.set(qn("w:val"), color)
    rPr.append(c)

  if underline:
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)

  rPr.append(
      parse_xml(r'<w:rFonts %s w:ascii="Arial" w:hAnsi="Arial"/>' % nsdecls("w"))
  )
  rPr.append(parse_xml(r'<w:sz %s w:val="18"/>' % nsdecls("w")))  # 9pt
  rPr.append(parse_xml(r'<w:i %s/>' % nsdecls("w")))

  new_run.append(rPr)

  text_elem = OxmlElement("w:t")
  text_elem.text = text
  new_run.append(text_elem)

  hyperlink.append(new_run)
  paragraph._p.append(hyperlink)


# Load history
HISTORY_FILE = "seen_stories.json"
recent_topics = []
if os.path.exists(HISTORY_FILE):
  try:
    with open(HISTORY_FILE, "r") as f:
      recent_topics = json.load(f).get("recent_topics", [])[-15:]
  except Exception:
    pass

history_context = ""
if recent_topics:
  topics_formatted = "\n- ".join(recent_topics)
  history_context = (
      f"\nRECENTLY COVERED TOPICS (DO NOT REPEAT UNLESS THERE IS A MAJOR NEW"
      f" UPDATE):\n- {topics_formatted}\n"
  )

today = datetime.date.today()
tomorrow = today + datetime.timedelta(days=1)
date_str = tomorrow.strftime("%A, %B %d, %Y").upper()

# 52-MINUTE TOTAL BROADCAST SYSTEM PROMPT
system_prompt = (
    f"You are an expert Caribbean radio news presenter writing a complete"
    f" 52-minute broadcast script for 2020 Vision Radio.\nTomorrow's broadcast"
    f" date is {date_str}. Scheduled Broadcast Window: 8:00 AM - 8:52 AM.\n\nSTRICT"
    " TIMING & WORD COUNT REQUIREMENTS:\n1. NEWS SEGMENT (30 MINUTES):"
    " Must contain ~4,000 words. Provide comprehensive, multi-paragraph"
    " reporting for St. Kitts & Nevis, Antigua & Barbuda, Anguilla, Sint"
    " Maarten, Statia, Saba, Montserrat, and Wider Caribbean. Include historical"
    " context, official statements, and station break placeholders.\n2. WEATHER"
    " SEGMENT (7 MINUTES): Must contain ~900 to 1,000 words. Focus EXCLUSIVELY"
    " on St. Kitts and Nevis. Provide an in-depth breakdown including synoptic"
    " analysis, island-by-island breakdown (Basseterre, Sandy Point, Nevis"
    " Peak, Charlestown), marine swell advisory, tide schedules, and 5-day"
    " outlook.\n3. SPORTS SEGMENT (15 MINUTES): Must contain ~2,000 words."
    " Lead with St. Kitts & Nevis local athletics and community cricket/football"
    " leagues, followed by Leeward Islands Cricket, CWI, CONCACAF, and regional"
    " sports updates.\n\nFORMAT & STYLE RULES:\n- Write in clear, spoken"
    " radio-presenter English.\n- Include stage cues in brackets like [pause],"
    " [station intro sting], [transition sting], and commercial break"
    " placeholders.\n- MUST include direct source URLs in brackets next to each"
    " story (e.g., [Source: https://...]).\n- MEDIA LINKS REQUIREMENT: If an"
    " item includes associated audio or video, include direct access links"
    " inside the source brackets (e.g. [Source: https://... | Audio: https://..."
    " | Video: https://...]).\n\nBROADCAST STRUCTURE:\n=== SCRIPT 1: NEWS"
    " SEGMENT (30 MINUTES) ===\n=== SCRIPT 2: WEATHER SEGMENT (7 MINUTES)"
    f" ===\n=== SCRIPT 3: SPORTS SEGMENT (15 MINUTES) ===\n{history_context}\nAt"
    " the very end of your response, output a single JSON"
    ' block:\n```json\n{"topics": ["Topic 1", "Topic 2"]}\n```'
)

script_content = None

# --- ATTEMPT 1: GOOGLE GEMINI ---
if GEMINI_API_KEY:
  try:
    print("Attempting Primary Provider: Google Gemini...")
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=GEMINI_API_KEY)
    for g_model in [
        "gemini-2.0-flash",
        "gemini-2.0-flash-lite",
        "gemini-1.5-pro",
    ]:
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

    client = OpenAI(api_key=OPENAI_API_KEY, timeout=30.0, max_retries=2)
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
        "claude-3-5-sonnet-20241022",
        "claude-3-haiku-20240307",
        "claude-3-5-haiku-20241022",
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

if not script_content:
  print("CRITICAL ERROR: All AI providers failed or credentials missing.")
  exit(1)

# Extract history topics
clean_script = script_content
if "```json" in script_content:
  try:
    parts = script_content.split("```json")
    clean_script = parts[0].strip()
    json_str = parts[1].split("```")[0].strip()
    extracted_data = json.loads(json_str)
    recent_topics.extend(extracted_data.get("topics", []))
    with open(HISTORY_FILE, "w") as f:
      json.dump({"recent_topics": recent_topics[-25:]}, f, indent=2)
  except Exception:
    pass

# Build Word Document
doc = docx.Document()
for section in doc.sections:
  section.top_margin = Inches(1)
  section.bottom_margin = Inches(1)
  section.left_margin = Inches(1)
  section.right_margin = Inches(1)

tbl = doc.add_table(rows=1, cols=1)
tbl.autofit = False
tbl.columns[0].width = Inches(6.5)
cell = tbl.cell(0, 0)
shading = parse_xml(r'<w:shd {} w:fill="1B365D"/>'.format(nsdecls("w")))
cell._tc.get_or_add_tcPr().append(shading)

p = cell.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r1 = p.add_run("2020 VISION RADIO - FULL 52-MINUTE BROADCAST SCRIPT")
r1.font.name = "Arial"
r1.font.size = Pt(16)
r1.font.bold = True
r1.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

p2 = cell.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
r2 = p2.add_run(
    f"FOR BROADCAST: {date_str} (30m News | 7m Weather | 15m Sports)"
)
r2.font.name = "Arial"
r2.font.size = Pt(11)
r2.font.bold = True
r2.font.color.rgb = RGBColor(0xD0, 0xE1, 0xF9)

# Regex to identify URLs inside source brackets
url_pattern = re.compile(r"https?://[^\s\]\|]+")

for line in clean_script.split("\n"):
  if line.strip():
    p_line = doc.add_paragraph()
    parts = line.split("[")
    p_line.add_run(parts[0])
    for part in parts[1:]:
      if "]" in part:
        cue_text, rest = part.split("]", 1)
        if cue_text.startswith("Source:"):
          r_bracket = p_line.add_run("[")
          r_bracket.font.italic = True
          r_bracket.font.bold = True
          r_bracket.font.size = Pt(9)
          r_bracket.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

          last_idx = 0
          for match in url_pattern.finditer(cue_text):
            start, end = match.span()
            if start > last_idx:
              r_prefix = p_line.add_run(cue_text[last_idx:start])
              r_prefix.font.italic = True
              r_prefix.font.bold = True
              r_prefix.font.size = Pt(9)
              r_prefix.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

            url = match.group(0)
            add_hyperlink(p_line, url, url, color="0000FF", underline=True)
            last_idx = end

          if last_idx < len(cue_text):
            r_suffix = p_line.add_run(cue_text[last_idx:])
            r_suffix.font.italic = True
            r_suffix.font.bold = True
            r_suffix.font.size = Pt(9)
            r_suffix.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

          r_bracket_close = p_line.add_run("]")
          r_bracket_close.font.italic = True
          r_bracket_close.font.bold = True
          r_bracket_close.font.size = Pt(9)
          r_bracket_close.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        else:
          r_cue = p_line.add_run("[" + cue_text + "]")
          r_cue.font.bold = True
          r_cue.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)

        p_line.add_run(rest)
      else:
        p_line.add_run("[" + part)

doc_filename = f"2020_Vision_Radio_Script_{tomorrow.strftime('%Y_%m_%d')}.docx"
doc.save(doc_filename)

# Send Email with Attachment
msg = MIMEMultipart()
msg["From"] = SENDER_EMAIL
msg["To"] = ", ".join(recipients)
msg["Subject"] = (
    f"2020 Vision Radio Script (52-Min Block) - For"
    f" {tomorrow.strftime('%A, %b %d')}"
)

body = (
    f"Hello Team,\n\nAttached is tomorrow's automated news (30m), weather (7m),"
    f" and sports (15m) script for 2020 Vision Radio ({date_str}).\n\nScheduled"
    " Broadcast: Tomorrow starting at 8:00 AM.\n"
)
msg.attach(MIMEText(body, "plain"))

with open(doc_filename, "rb") as attachment:
  part = MIMEBase("application", "octet-stream")
  part.set_payload(attachment.read())
  encoders.encode_base64(part)
  part.add_header(
      "Content-Disposition", f"attachment; filename= {doc_filename}"
  )
  msg.attach(part)

server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(SENDER_EMAIL, SENDER_PASSWORD)
server.sendmail(SENDER_EMAIL, recipients, msg.as_string())
server.quit()
print("SUCCESS: Full script generated and emailed to all recipients!")