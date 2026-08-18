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
import feedparser
import requests

# --- 1. ENVIRONMENT VARIABLES ---
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "").strip()
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "").strip()

SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "").strip()
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD", "").strip()
PRESENTER_EMAIL_RAW = os.environ.get("PRESENTER_EMAIL", "")

# RSS Feed Links
SKN_POLICE_FB_RSS = os.environ.get("SKN_POLICE_FB_RSS", "").strip()
NEVIS_NEWSCAST_RSS = os.environ.get("NEVIS_NEWSCAST_RSS", "").strip()
TV_CARIBBEAN_RSS = os.environ.get("TV_CARIBBEAN_RSS", "").strip()

recipients = [
    email.strip() for email in PRESENTER_EMAIL_RAW.split(",") if email.strip()
]


def add_hyperlink(paragraph, url, text, color="0000FF", underline=True):
    """Adds a native XML hyperlink to a Word document."""
    part = paragraph.part
    r_id = part.relate_to(
        url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True
    )

    hyperlink = parse_xml(
        '<w:hyperlink'
        ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
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


def verify_and_clean_links(script_text):
    """
    Scans the generated script for URLs in [Source: ...] tags, tests each HTTP status,
    and replaces dead or unreachable links with '[Link Unavailable]' to guarantee working links.
    """
    url_pattern = re.compile(r"https?://[^\s\]\|]+")
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
    }

    validated_urls = {}

    def check_url(match):
        raw_url = match.group(0)
        url = raw_url.rstrip(".,;)")  # Clean trailing punctuation

        if url in validated_urls:
            return validated_urls[url]

        print(f"Validating link: {url}")
        try:
            # 1. Try HEAD request first for fast checking
            response = requests.head(
                url, timeout=4, headers=headers, allow_redirects=True
            )
            if response.status_code < 400:
                validated_urls[url] = url
                return url

            # 2. Fallback to stream GET request if HEAD is rejected by host
            response = requests.get(
                url, timeout=4, headers=headers, stream=True
            )
            if response.status_code < 400:
                validated_urls[url] = url
                return url
        except Exception as err:
            print(f"   Link check failed ({err}) for: {url}")

        print(f"--> REMOVED BROKEN LINK: {url}")
        validated_urls[url] = "[Link Unavailable]"
        return "[Link Unavailable]"

    return url_pattern.sub(check_url, script_text)


# --- 2. FETCH OFFICIAL RSS & MEDIA FEEDS ---
media_context_items = []
rss_sources = [
    ("SKN Police Force", SKN_POLICE_FB_RSS),
    ("Nevis Newscast", NEVIS_NEWSCAST_RSS),
    ("Television Caribbean", TV_CARIBBEAN_RSS),
]

for source_name, rss_url in rss_sources:
    if rss_url:
        try:
            print(f"Fetching {source_name} feed: {rss_url}")
            feed = feedparser.parse(rss_url)
            for entry in feed.entries[:5]:  # Fetch top 5 items
                title = getattr(entry, "title", "")
                summary = getattr(entry, "summary", "") or getattr(
                    entry, "description", ""
                )
                link = getattr(entry, "link", "")
                clean_summary = re.sub("<[^<]+?>", "", summary)[:300]
                media_context_items.append(
                    f"- [{source_name}] {title}: {clean_summary} (Source: {link})"
                )
        except Exception as rss_err:
            print(
                f"Notice: Could not parse {source_name} feed ({rss_url}): {rss_err}"
            )

fb_context_str = ""
if media_context_items:
    fb_context_str = (
        "\nLATEST OFFICIAL BULLETINS, POLICE REPORTS & MEDIA FEEDS:\n"
        + "\n".join(media_context_items)
        + "\n"
    )

# --- 3. LOAD STORY HISTORY ---
HISTORY_FILE = "seen_stories.json"
recent_topics = []
if os.path.exists(HISTORY_FILE):
    try:
        with open(HISTORY_FILE, "r") as f:
            recent_topics = json.load(f).get("recent_topics", [])[-20:]
    except Exception:
        pass

history_context = ""
if recent_topics:
    topics_formatted = "\n- ".join(recent_topics)
    history_context = (
        f"\nRECENTLY COVERED TOPICS (STRICT DEDUPLICATION - DO NOT REPEAT UNLESS"
        f" THERE IS A MAJOR NEW DEVELOPMENT):\n- {topics_formatted}\n"
    )

today = datetime.date.today()
tomorrow = today + datetime.timedelta(days=1)
date_str = tomorrow.strftime("%A, %B %d, %Y").upper()

# --- 4. SYSTEM PROMPT ---
system_prompt = (
    "You are an expert Caribbean radio news presenter writing a broadcast"
    " script for 2020 Vision Radio 102.1 FM.\nTomorrow's broadcast date is"
    f" {date_str}. Scheduled Broadcast Window starts at 8:00 AM.\n\nSTRICT"
    " REGIONAL NEWS COVERAGE QUOTAS:\n1. ST. KITTS AND NEVIS (AT LEAST 6"
    " ITEMS TOTAL):\n   - You MUST include at least 3 distinct news stories for"
    " St. Kitts.\n   - You MUST include at least 3 distinct news stories for"
    " Nevis.\n2. LEEWARD & NEIGHBORING ISLANDS: Include stories for Antigua &"
    " Barbuda, Anguilla, Sint Maarten/St. Martin, Statia, Saba, and"
    " Montserrat.\n3. MANDATORY WIDER CARIBBEAN COVERAGE (AT LEAST 1 ITEM"
    " EACH):\n   You MUST attempt to include at least 1 real news item for"
    " each of the following territories:\n   - Trinidad and Tobago\n   - St."
    " Vincent and the Grenadines\n   - Dominica\n   - Jamaica\n   - Grenada\n  "
    " - St. Lucia\n   - Guyana\n   - Belize\n   - United States Virgin Islands"
    " (USVI)\n   - British Virgin Islands (BVI)\n\nSTRICT WEATHER RULES:\n-"
    " Focus EXCLUSIVELY on St. Kitts and Nevis.\n- DAILY WEATHER RULE: Provide"
    f" the forecast strictly for tomorrow's date ({date_str}) ONLY. Do NOT"
    " include a 5-day forecast or general multi-day outlook.\n- THURSDAY"
    " SPECIAL RULE: If tomorrow's date is Friday (script generated on Thursday),"
    " you MUST create a dedicated section immediately after Friday's forecast"
    " titled 'Weekend Outlook', providing a detailed weather forecast"
    " specifically for Saturday and Sunday.\n- Forecast details must include"
    " synoptic analysis, island-by-island breakdown (Basseterre, Sandy Point,"
    " Nevis Peak, Charlestown), marine swell advisories, and tide schedules.\n\n4."
    " SPORTS SEGMENT: Lead with St. Kitts & Nevis local athletics and community"
    " cricket/football leagues, followed by Leeward Islands Cricket, CWI,"
    " CONCACAF, and regional sports updates.\n\nACCURACY & FORMATTING"
    " RULES:\n- STRICT DEDUPLICATION: Do NOT repeat news items covered in"
    " previous days unless there is a fresh, breaking update.\n- DO NOT mention"
    " segment durations, timing lengths, or word counts inside the generated"
    " script text (e.g. do NOT write '30 Minutes', '7 Minutes', or '4000"
    " words').\n- DO NOT invent fictional persons or names (e.g., do NOT use"
    " hallucinated names like 'Marcus St. Clair'). Report ONLY real, verifiable"
    " public officials and real citizens.\n- Write in clear, spoken"
    " radio-presenter English for 2020 Vision Radio 102.1 FM.\n- Include stage"
    " cues in brackets like [pause], [station intro sting], [transition sting],"
    " and commercial break placeholders.\n- LINK ACCURACY MANDATE: Include direct source URLs"
    " ONLY if retrieved directly from live RSS context feeds or verified live search results."
    " NEVER fabricate, guess, or synthesize URL strings.\n- MEDIA LINKS REQUIREMENT: If an item"
    " includes verified associated audio or video, include direct access links inside the source brackets"
    f" (e.g. [Source: https://... | Audio: https://...]).\n{fb_context_str}{history_context}\nBROADCAST"
    " STRUCTURE:\n=== SCRIPT 1: NEWS SEGMENT ===\n=== SCRIPT 2: WEATHER SEGMENT"
    " ===\n=== SCRIPT 3: SPORTS SEGMENT ===\n\nAt the very end of your response,"
    ' output a single JSON block:\n
http://googleusercontent.com/immersive_entry_chip/0