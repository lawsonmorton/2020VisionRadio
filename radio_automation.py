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

# 1. VERIFY ENVIRONMENT VARIABLES
required_secrets = [
    "GEMINI_API_KEY",
    "SENDER_EMAIL",
    "SENDER_PASSWORD",
    "PRESENTER_EMAIL",
]
missing = [sec for sec in required_secrets if not os.environ.get(sec)]

if missing:
  print(f"ERROR: Missing required GitHub secrets: {', '.join(missing)}")
  print(
      "Please add them under Repository Settings > Secrets and variables >"
      " Actions."
  )
  exit(1)

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
SENDER_EMAIL = os.environ.get("SENDER_EMAIL")
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD")
PRESENTER_EMAIL_RAW = os.environ.get("PRESENTER_EMAIL")

# Parse comma-separated list of recipient email addresses
recipients = [
    email.strip() for email in PRESENTER_EMAIL_RAW.split(",") if email.strip()
]

# 2. LOAD STORY HISTORY TO PREVENT REPETITION
HISTORY_FILE = "seen_stories.json"
recent_topics = []

if os.path.exists(HISTORY_FILE):
  try:
    with open(HISTORY_FILE, "r") as f:
      history_data = json.load(f)
      recent_topics = history_data.get("recent_topics", [])[-15:]
  except Exception as e:
    print(f"Notice: Could not load history file ({e}). Starting fresh.")

history_context = ""
if recent_topics:
  history_context = (
      "\nRECENTLY COVERED TOPICS (DO NOT REPEAT UNLESS THERE IS A SIGNIFICANT"
      " NEW UPDATE):\n- "
      + "\n- ".join(recent_topics)
      + "\n"
  )

try:
  from google import genai
  from google.genai import types

  client = genai.Client(api_key=GEMINI_API_KEY)

  # Calculate target broadcast date (tomorrow at 8:00 AM)
  today = datetime.date.today()
  tomorrow = today + datetime.timedelta(days=1)
  date_str = tomorrow.strftime("%A, %B %d, %Y").upper()

  system_prompt = f"""
    You are an expert Caribbean radio news presenter writing a broadcast script for 2020 Vision Radio.
    Tomorrow's broadcast date is {date_str}. The reading broadcast time is 8:00 AM.

    FORMAT & STYLE RULES:
    - Script formatted for an on-air DJ/Presenter to read.
    - Spoken English style for all on-air text.
    - Include stage cues in brackets like [pause] and [transition sting].
    - MUST include direct source URLs in brackets next to each news story so the presenter can verify credibility (e.g. [Source: https://...]).

    NEWS COVERAGE AREA & STRICT ORDER:
    1. St. Kitts and Nevis (2-4 real stories)
    2. Antigua and Barbuda (2-4 real stories)
    3. Anguilla (2-3 real stories)
    4. Sint Maarten / St. Martin (2-3 real stories)
    5. Sint Eustatius (Statia) (1-2 real stories)
    6. Saba (1-2 real stories)
    7. Montserrat (1-2 real stories)
    8. Wider Caribbean News (Key current stories from Dominica, Jamaica, St. Vincent and the Grenadines, etc.)
    {history_context}
    STRICT DEDUPLICATION RULE:
    - Do NOT repeat previously covered news items listed above unless there is a fresh, substantial development or breaking update to report today. Skip stale filler if no new development exists for an island.

    WEATHER SEGMENT:
    - Focus EXCLUSIVELY on St. Kitts and Nevis weather outlook for tomorrow (High/Low, Wind speed/direction, Rain %, Hurricane/tropical advisories). Do not include weather for other islands.

    SPORTS SEGMENT:
    - ALWAYS lead with St. Kitts and Nevis sports news first (SKNFA, SKN cricket, athletics, netball, or athletes with SKN ties).
    - Follow with regional cricket (CPL, West Indies), CONCACAF/regional football, and other Leeward sports.

    At the very end of your response, output a single JSON block on a new line containing a list of 3-5 short topic phrases representing today's main news stories, structured like this:
    ```json
    {{"topics": ["Topic 1", "Topic 2", "Topic 3"]}}
    ```

    Format the main script into two clear sections:
    === SCRIPT 1: NEWS + WEATHER ===
    === SCRIPT 2: SPORTS ===
    """

  # Valid Gemini models supporting Google Search Grounding in google-genai SDK
  candidate_models = ["gemini-2.0-flash", "gemini-2.0-flash-lite"]
  script_content = None

  for model_name in candidate_models:
    print(
        f"Attempting API call with model: {model_name} for date: {date_str}..."
    )
    for attempt in range(1, 3):
      try:
        response = client.models.generate_content(
            model=model_name,
            contents=system_prompt,
            config=types.GenerateContentConfig(
                tools=[{"google_search": {}}],
            ),
        )
        script_content = response.text
        print(f"SUCCESS: Script generated using model '{model_name}'.")
        break
      except Exception as api_err:
        err_str = str(api_err)
        if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
          print(
              f"Rate limited on {model_name} (Attempt {attempt}). Waiting 70"
              " seconds for rate limit window reset..."
          )
          time.sleep(70)
        else:
          print(f"Error on {model_name}: {api_err}")
          break
    if script_content:
      break

  if not script_content:
    print(
        "CRITICAL ERROR: Daily search grounding rate limit reached. Please wait"
        " 10-15 minutes before re-testing."
    )
    exit(1)

  # Extract JSON topics block from AI output to update history
  clean_script = script_content
  if "```json" in script_content:
    try:
      parts = script_content.split("```json")
      clean_script = parts[0].strip()
      json_str = parts[1].split("```")[0].strip()
      extracted_data = json.loads(json_str)
      new_topics = extracted_data.get("topics", [])
      recent_topics.extend(new_topics)

      # Save updated history
      with open(HISTORY_FILE, "w") as f:
        json.dump({"recent_topics": recent_topics[-25:]}, f, indent=2)
      print(f"Updated story history with {len(new_topics)} new topic(s).")
    except Exception as parse_err:
      print(f"Notice: Could not parse story history JSON ({parse_err}).")

  # 3. BUILD WORD DOCUMENT (.DOCX)
  doc = docx.Document()
  for section in doc.sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

  # Header Banner
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

  # Line-by-line formatting for bracketed cues & sources
  lines = clean_script.split("\n")
  for line in lines:
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

  doc_filename = (
      f"2020_Vision_Radio_Script_{tomorrow.strftime('%Y_%m_%d')}.docx"
  )
  doc.save(doc_filename)
  print(f"Document created: {doc_filename}")

  # 4. EMAIL ATTACHMENT TO ALL RECIPIENTS
  msg = MIMEMultipart()
  msg["From"] = SENDER_EMAIL
  msg["To"] = ", ".join(recipients)
  msg["Subject"] = (
      f"2020 Vision Radio Script - For {tomorrow.strftime('%A, %b %d')}"
  )

  body = f"""Hello Team,

Attached is tomorrow's automated news, sports, and weather script for 2020 Vision Radio ({date_str}).

Coverage Summary:
- Core Leeward Islands: St. Kitts & Nevis, Antigua & Barbuda, Anguilla, Sint Maarten/St. Martin, Statia, Saba, Montserrat
- Wider Caribbean: Dominica, Jamaica, St. Vincent
- Weather: St. Kitts & Nevis exclusively
- Verified source links included in brackets for presenter reference.

Scheduled Broadcast: Tomorrow at 8:00 AM.
"""
  msg.attach(MIMEText(body, "plain"))

  with open(doc_filename, "rb") as f:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition", f"attachment; filename={doc_filename}"
    )
    msg.attach(part)

  print(
      f"Connecting to SMTP server to email {len(recipients)} recipient(s):"
      f" {', '.join(recipients)}..."
  )
  server = smtplib.SMTP("smtp.gmail.com", 587)
  server.starttls()
  server.login(SENDER_EMAIL, SENDER_PASSWORD)
  server.sendmail(SENDER_EMAIL, recipients, msg.as_string())
  server.quit()
  print("SUCCESS: Script generated and emailed to all recipients!")

except Exception as e:
  print("\n--- CRITICAL ERROR ENCOUNTERED ---")
  traceback.print_exc()
  exit(1)