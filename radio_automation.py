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
PRESENTER_EMAIL_RAW = os.environ.get("PRESENTER_EMAIL")

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
  history_context = (
      "\nRECENTLY COVERED TOPICS (DO NOT REPEAT UNLESS THERE IS A MAJOR NEW"
      " UPDATE):\n- "
      + "\n- ".join(recent_topics)
      + "\n"
  )

today = datetime.date.today()
tomorrow = today + datetime.timedelta(days=1)
date_str = tomorrow.strftime("%A, %B %d, %Y").upper()

system_prompt = f"""
You are an expert Caribbean radio news presenter writing a broadcast script for 2020 Vision Radio.
Tomorrow's broadcast date is {date_str}. Broadcast time is 8:00 AM.

FORMAT & STYLE RULES:
- Script formatted for an on-air DJ/Presenter to read.
- Spoken English style for all on-air text.
- Include stage cues in brackets like [pause] and [transition sting].
- MUST include direct source URLs in brackets next to each news story (e.g. [Source: https://...]).

NEWS COVERAGE AREA & STRICT ORDER:
1. St. Kitts and Nevis (2-4 real stories)
2. Antigua and Barbuda (2-4 real stories)
3. Anguilla (2-3 real stories)
4. Sint Maarten / St. Martin (2-3 real stories)
5. Sint Eustatius (Statia) (1-2 real stories)
6. Saba (1-2 real stories)
7. Montserrat (1-2 real stories)
8. Wider Caribbean News (Key current stories from Dominica, Jamaica, St. Vincent)
{history_context}
WEATHER SEGMENT:
- Focus EXCLUSIVELY on St. Kitts and Nevis weather outlook for tomorrow. Do not include weather for other islands.

SPORTS SEGMENT:
- ALWAYS lead with St. Kitts and Nevis sports news first.
- Follow with regional cricket, CONCACAF football, and Leeward sports.

At the very end of your response, output a single JSON block:
```json
{{"topics": ["Topic 1", "Topic 2"]}}