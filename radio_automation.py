import datetime
import json
import os
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
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")

SENDER_EMAIL = os.environ.get("SENDER_EMAIL")
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD")
PRESENTER_EMAIL_RAW = os.environ.get("PRESENTER_EMAIL", "")

recipients = [
    email.strip() for email in PRESENTER_EMAIL_RAW.split(",") if email.strip()
]

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

system_prompt = (
    f"You are an expert Caribbean radio news presenter writing a broadcast"
    f" script for 2020 Vision Radio.\nTomorrow's broadcast date is"
    f" {date_str}. Broadcast time is 8:00 AM.\n\nFORMAT & STYLE"
    " RULES:\n- Script formatted for an on-air DJ/Presenter to read.\n- Spoken"
    " English style for all on-air text.\n- Include stage cues in brackets"
    " like [pause] and [transition sting].\n- MUST include direct source URLs in"
    " brackets next to each news story (e.g. [Source:"
    " https://...]).\n\nNEWS COVERAGE AREA & STRICT ORDER:\n1. St. Kitts and"
    " Nevis (2-4 real stories)\n2. Antigua and Barbuda (2-4 real stories)\n3."
    " Anguilla (2-3 real stories)\n4. Sint Maarten / St. Martin (2-3 real"
    " stories)\n5. Sint Eustatius (Statia) (1-2 real stories)\n6. Saba (1-2"
    " real stories)\n7. Montserrat (1-2 real stories)\n8. Wider Caribbean News"
    " (Key current stories from Dominica, Jamaica, St."
    f" Vincent)\n{history_context}\nWEATHER SEGMENT:\n- Focus EXCLUSIVELY on"
    " St. Kitts and Nevis weather outlook for tomorrow. Do not include weather"
    " for other islands.\n\nSPORTS SEGMENT:\n- ALWAYS lead with St. Kitts and"
    " Nevis sports news first.\n- Follow with regional cricket, CONCACAF"
    " football, and Leeward sports.\n\nAt the very end of your response, output"
    ' a single JSON block:\n```json\n{"topics": ["Topic 1",'
    ' "Topic 2"]}\n```\n\nFormat the main script into two clear sections:\n==='
    " SCRIPT 1: NEWS + WEATHER ===\n=== SCRIPT 2: SPORTS ==="
)

script_content = None

# --- ATTEMPT 1: GOOGLE GEMINI ---
if GEMINI_API_KEY:
  try:
    print("Attempting Primary Provider: Google Gemini...")
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=GEMINI_API_KEY)
    response = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=system_prompt,
        config=types.GenerateContentConfig(tools=[{"google_search": {}}]),
    )
    script_content = response.text
    print("SUCCESS: Generated via Google Gemini!")
  except Exception as e:
    print(f"Gemini API failed or rate limited: {e}")

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
    response = client.messages.create(
        model="claude-3-5-sonnet-20241022",
        max_tokens=4000,
        messages=[{"role": "user", "content": system_prompt}],
    )
    script_content = response.content[0].text
    print("SUCCESS: Generated via Anthropic Claude!")
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
r1 = p.add_run("2020 VISION RADIO - DAILY NEWS & WEATHER SCRIPT")
r1.font.name = "Arial"
r1.font.size = Pt(16)
r1.font.bold = True
r1.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

p2 = cell.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
r2 = p2.add_run(f"FOR BROADCAST: {date_str} AT 8:00 AM")
r2.font.name = "Arial"
r2.font.size = Pt(11)
r2.font.bold = True
r2.font.color.rgb = RGBColor(0xD0, 0xE1, 0xF9)

for line in clean_script.split("\n"):
  if line.strip():
    p_line = doc.add_paragraph()
    parts = line.split("[")
    p_line.add_run(parts[0])
    for part in parts[1:]:
      if "]" in part:
        cue_text, rest = part.split("]", 1)
        r_cue = p_line.add_run("[" + cue_text + "]")
        r_cue.font.bold = True
        if cue_text.startswith("Source:"):
          r_cue.font.italic = True
          r_cue.font.size = Pt(9)
          r_cue.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        else:
          r_cue.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)
        p_line.add_run(rest)
      else:
        p_line.add_run("[" + part)

doc_filename = f"2020_Vision_Radio_Script_{tomorrow.strftime('%Y_%m_%d')}.docx"
doc.save(doc_filename)

# Send Email
msg = MIMEMultipart()
msg["From"] = SENDER_EMAIL
msg["To"] = ", ".join(recipients)
msg["Subject"] = (
    f"2020 Vision Radio Script - For {tomorrow.strftime('%A, %b %d')}"
)

body = (
    f"Hello Team,\n\nAttached is tomorrow's automated news, sports, and weather"
    f" script for 2020 Vision Radio ({date_str}).\n\nScheduled Broadcast:"
    " Tomorrow at 8:00 AM.\n"
)
msg.attach(MIMEText(body, "plain"))

server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(SENDER_EMAIL, SENDER_PASSWORD)
server.sendmail(SENDER_EMAIL, recipients, msg.as_string())
server.quit()
print("SUCCESS: Script generated and emailed to all recipients!")