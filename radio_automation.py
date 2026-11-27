#!/usr/bin/env python3
"""
2020 Vision Radio 102.1 FM - daily broadcast script generator.

HOW NEWS NOW WORKS (the key change from the old version)
---------------------------------------------------------
Old: the AI was asked to "find" news and write the source links itself. Only the
Gemini path had web search; the OpenAI/Claude fallbacks had none, so they could
only make news up. Links were only checked for "does the page load", never
"is this the right story", and no story had a date.

New: Python collects real, dated news items from RSS feeds (direct outlet feeds
plus Google News searches per island), keeps only items from the last
NEWS_MAX_AGE_HOURS, then downloads each full article. The AI writes a finished,
flowing on-air news item from that article text only. Source links, audio/video
links and dates are attached by the code from the real data, so the AI cannot
invent a story, a date or a URL. Each story gets a footnote: source link with the
publication date in [brackets].

Run normally:          python radio_automation.py
Test your news feeds:  python radio_automation.py --check-feeds
"""

import calendar
import concurrent.futures
import datetime
import difflib
import hashlib
import html
import json
import os
import re
import smtplib
import sys
import traceback
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from urllib.parse import quote

import docx
import feedparser
import requests
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

# --- 1. SETTINGS / ENVIRONMENT VARIABLES -------------------------------------
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "").strip()
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "").strip()

SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "").strip()
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD", "").strip()
PRESENTER_EMAIL_RAW = os.environ.get("PRESENTER_EMAIL", "")
recipients = [e.strip() for e in PRESENTER_EMAIL_RAW.split(",") if e.strip()]

# Your existing optional feeds (still supported)
SKN_POLICE_FB_RSS = os.environ.get("SKN_POLICE_FB_RSS", "").strip()
NEVIS_NEWSCAST_RSS = os.environ.get("NEVIS_NEWSCAST_RSS", "").strip()
TV_CARIBBEAN_RSS = os.environ.get("TV_CARIBBEAN_RSS", "").strip()
# Extra feeds without editing code. Format: "territory_key|Outlet name|URL;..."
EXTRA_FEEDS = os.environ.get("EXTRA_FEEDS", "").strip()

NEWS_MAX_AGE_HOURS = int(os.environ.get("NEWS_MAX_AGE_HOURS", "48"))
USE_GOOGLE_NEWS = os.environ.get("USE_GOOGLE_NEWS", "1") != "0"
ARTICLE_CHARS = int(os.environ.get("ARTICLE_CHARS", "6000"))  # article text given to the AI per story

OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o")
ANTHROPIC_MODELS = [
    m.strip()
    for m in os.environ.get(
        "ANTHROPIC_MODELS", "claude-sonnet-5-5,claude-haiku-5-5"
    ).split(",")
    if m.strip()
]

HISTORY_FILE = "seen_stories.json"
HISTORY_DAYS = 14

# St. Kitts & Nevis is UTC-4 all year (no daylight saving). GitHub runners use UTC,
# so we compute "today" in local time to avoid the date being a day off in the evening.
AST = datetime.timezone(datetime.timedelta(hours=-4), "AST")

UA = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}

# --- 2. COVERAGE: territories, keywords, Google News queries -------------------
# key, label, group, max stories, keywords (item must mention one), Google News query
TERRITORIES = [
    ("stkitts", "ST. KITTS", "local", 3,
     ["st. kitts", "st kitts", "saint kitts", "basseterre", "kittitian", "frigate bay",
      "sandy point", "cayon", "old road", "tabernacle", "dieppe bay"],
     '"St Kitts" OR "St. Kitts" OR Basseterre OR Nevis OR Charlestown Nevis'),
    ("nevis", "NEVIS", "local", 3,
     ["nevis", "nevisian", "charlestown", "gingerland", "newcastle nevis", "nia "],
     None),  # shares the St. Kitts & Nevis query above
    ("antigua", "ANTIGUA & BARBUDA", "neighbours", 2,
     ["antigua", "barbuda", "st. john's", "st john's"], "Antigua OR Barbuda"),
    ("anguilla", "ANGUILLA", "neighbours", 1, ["anguilla"], "Anguilla"),
    ("sxm", "SINT MAARTEN / ST. MARTIN", "neighbours", 2,
     ["sint maarten", "st. maarten", "st maarten", "saint martin", "st. martin",
      "st martin", "philipsburg", "marigot"],
     '"Sint Maarten" OR "St. Maarten" OR "Saint Martin"'),
    ("statia_saba", "STATIA & SABA", "neighbours", 1,
     ["statia", "sint eustatius", "st. eustatius", "saba"],
     '"Sint Eustatius" OR Statia OR "Saba island"'),
    ("montserrat", "MONTSERRAT", "neighbours", 1, ["montserrat", "plymouth"],
     "Montserrat Caribbean"),
    ("tt", "TRINIDAD & TOBAGO", "wider", 1,
     ["trinidad", "tobago", "port of spain", "t&t"], '"Trinidad and Tobago"'),
    ("svg", "ST. VINCENT & THE GRENADINES", "wider", 1,
     ["st. vincent", "st vincent", "saint vincent", "grenadines", "kingstown"],
     '"St Vincent and the Grenadines" OR "St. Vincent and the Grenadines"'),
    ("dominica", "DOMINICA", "wider", 1,
     ["dominica", "roseau"], 'Dominica Roseau -"Dominican Republic"'),
    ("jamaica", "JAMAICA", "wider", 1,
     ["jamaica", "kingston"], 'Jamaica Kingston OR "Jamaica Gleaner" OR "Jamaica Observer"'),
    ("grenada", "GRENADA", "wider", 1,
     ["grenada", "st. george's", "carriacou"], "Grenada Carriacou OR \"St. George's\""),
    ("stlucia", "ST. LUCIA", "wider", 1,
     ["st. lucia", "st lucia", "saint lucia", "castries"], '"Saint Lucia" OR "St Lucia"'),
    ("guyana", "GUYANA", "wider", 1, ["guyana", "georgetown"], "Guyana"),
    ("belize", "BELIZE", "wider", 1, ["belize"], "Belize"),
    ("barbados", "BARBADOS", "wider", 1, ["barbados", "bridgetown"], "Barbados"),
    ("usvi", "U.S. VIRGIN ISLANDS", "wider", 1,
     ["u.s. virgin islands", "us virgin islands", "usvi", "st. thomas", "st. croix",
      "st. john, usvi", "virgin islands daily news"],
     '"U.S. Virgin Islands" OR USVI OR "St. Thomas" OR "St. Croix"'),
    ("bvi", "BRITISH VIRGIN ISLANDS", "wider", 1,
     ["british virgin islands", "bvi", "tortola", "road town", "virgin gorda"],
     '"British Virgin Islands" OR BVI OR Tortola'),
]
T_BY_KEY = {t[0]: t for t in TERRITORIES}

# Direct outlet feeds: (territory key or None, outlet, website or feed URL).
# You can give just the outlet's homepage: the script auto-discovers the real RSS
# feed (from the page's <link rel="alternate"> tag, then common paths like /feed/).
# None = Caribbean-wide feed; each item is routed to an island by keywords.
# Domains for SKNIS, WINN FM and ZIZ were confirmed by web search; the rest are
# best-effort. Run `--check-feeds` to see which work and the exact feed URL found.
DIRECT_FEEDS = [
    ("stkitts", "SKNIS", "https://www.sknis.gov.kn/"),
    ("stkitts", "WINN FM", "https://winnmediaskn.com/"),
    ("stkitts", "ZIZ Broadcasting", "https://zizonline.com/"),
    ("nevis", "Nevis Pages", "https://www.nevispages.com/"),
    ("antigua", "Antigua Observer", "https://antiguaobserver.com/"),
    ("antigua", "Antigua News Room", "https://antiguanewsroom.com/"),
    ("sxm", "The Daily Herald (St. Maarten)", "https://www.thedailyherald.sx/"),
    ("tt", "Newsday (T&T)", "https://newsday.co.tt/"),
    ("jamaica", "Jamaica Observer", "https://www.jamaicaobserver.com/"),
    ("guyana", "Guyana Chronicle", "https://guyanachronicle.com/"),
    ("guyana", "Stabroek News", "https://www.stabroeknews.com/"),
    ("dominica", "Dominica News Online", "https://dominicanewsonline.com/"),
    ("svg", "iWitness News", "https://www.iwnsvg.com/"),
    ("grenada", "NOW Grenada", "https://www.nowgrenada.com/"),
    ("stlucia", "St. Lucia Times", "https://stluciatimes.com/"),
    ("belize", "Breaking Belize News", "https://www.breakingbelizenews.com/"),
    ("barbados", "Barbados Today", "https://barbadostoday.bb/"),
    ("usvi", "St. Thomas Source", "https://stthomassource.com/"),
    ("bvi", "BVI Beacon", "https://bvibeacon.com/"),
]
if SKN_POLICE_FB_RSS:
    DIRECT_FEEDS.append(("stkitts", "SKN Police Force", SKN_POLICE_FB_RSS))
if NEVIS_NEWSCAST_RSS:
    DIRECT_FEEDS.append(("nevis", "Nevis Newscast", NEVIS_NEWSCAST_RSS))
if TV_CARIBBEAN_RSS:
    DIRECT_FEEDS.append((None, "Television Caribbean", TV_CARIBBEAN_RSS))
for _entry in filter(None, (x.strip() for x in EXTRA_FEEDS.split(";"))):
    _bits = [b.strip() for b in _entry.split("|")]
    if len(_bits) == 3:
        DIRECT_FEEDS.append((_bits[0] if _bits[0] in T_BY_KEY else None, _bits[1], _bits[2]))

SKN_BOTH_PHRASES = [
    "st. kitts and nevis", "st kitts and nevis", "st. kitts & nevis",
    "st kitts & nevis", "saint kitts and nevis", "st. kitts-nevis", "st kitts-nevis",
]


# --- 3. DOCX HELPERS ------------------------------------------------------------
def add_hyperlink(paragraph, url, text, color="0000FF", underline=True):
    """Adds a native clickable hyperlink to a Word paragraph."""
    r_id = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = parse_xml(
        '<w:hyperlink'
        ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        f' r:id="{r_id}"/>'
    )
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    # Child order follows the Word schema (rFonts, i, color, sz, u)
    rPr.append(parse_xml(r'<w:rFonts %s w:ascii="Arial" w:hAnsi="Arial"/>' % nsdecls("w")))
    rPr.append(parse_xml(r'<w:i %s/>' % nsdecls("w")))
    if color:
        c = OxmlElement("w:color")
        c.set(qn("w:val"), color)
        rPr.append(c)
    rPr.append(parse_xml(r'<w:sz %s w:val="18"/>' % nsdecls("w")))  # 9pt
    if underline:
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        rPr.append(u)
    new_run.append(rPr)
    text_elem = OxmlElement("w:t")
    text_elem.text = text
    new_run.append(text_elem)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


GREY = RGBColor(0x55, 0x55, 0x55)
NAVY = RGBColor(0x1B, 0x36, 0x5D)
URL_PATTERN = re.compile(r"https?://[^\s\]\|]+")


def _small_run(paragraph, text):
    r = paragraph.add_run(text)
    r.font.italic = True
    r.font.bold = True
    r.font.size = Pt(9)
    r.font.color.rgb = GREY
    return r


def render_footnote(doc, text):
    """Small grey source footnote with clickable links; the [date] sits next to the link."""
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_after = Pt(14)
    last = 0
    for m in URL_PATTERN.finditer(text):
        if m.start() > last:
            _small_run(p, text[last:m.start()])
        add_hyperlink(p, m.group(0), m.group(0))
        last = m.end()
    if last < len(text):
        _small_run(p, text[last:])


def render_line(doc, line):
    """
    Renders one script line.
      '=== X ===' segment heading | '## X' island heading | '# X' story headline
      anything else = spoken text; [bracketed] = cues; [Source: ...] = clickable links
    """
    s = line.rstrip()
    if not s.strip():
        return
    if s.startswith("^ "):
        render_footnote(doc, s[2:])
        return
    if s.strip() == "###":
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run("###").bold = True
        return
    if s.startswith("==="):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(18)
        r = p.add_run(s.strip("= ").strip())
        r.font.bold = True
        r.font.size = Pt(15)
        r.font.color.rgb = NAVY
        return
    if s.startswith("## "):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(s[3:].strip())
        r.font.bold = True
        r.font.size = Pt(13)
        r.font.color.rgb = RGBColor(0x2E, 0x6D, 0xA4)
        return
    if s.startswith("# "):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(s[2:].strip())
        r.font.bold = True
        r.font.size = Pt(13)
        return

    p_line = doc.add_paragraph()
    parts = s.split("[")
    p_line.add_run(parts[0])
    for part in parts[1:]:
        if "]" in part:
            cue_text, rest = part.split("]", 1)
            if cue_text.startswith("Source:"):
                _small_run(p_line, "[")
                last_idx = 0
                for match in URL_PATTERN.finditer(cue_text):
                    start, end = match.span()
                    if start > last_idx:
                        _small_run(p_line, cue_text[last_idx:start])
                    add_hyperlink(p_line, match.group(0), match.group(0))
                    last_idx = end
                if last_idx < len(cue_text):
                    _small_run(p_line, cue_text[last_idx:])
                _small_run(p_line, "]")
            else:
                r_cue = p_line.add_run(f"[{cue_text}]")
                r_cue.font.bold = True
            p_line.add_run(rest)
        else:
            p_line.add_run(f"[{part}")


# --- 4. NEWS COLLECTION -----------------------------------------------------------
def clean_html(text):
    text = html.unescape(re.sub(r"<[^>]+>", " ", text or ""))
    return re.sub(r"\s+", " ", text).strip()


def norm_title(title):
    return re.sub(r"[^a-z0-9]+", " ", (title or "").lower()).strip()


def story_key(title):
    return hashlib.sha1(norm_title(title)[:90].encode()).hexdigest()[:16]


FEED_PATHS = ["feed/", "feed", "rss", "rss.xml", "feed.xml", "?feed=rss2", "index.xml"]


def _entries_from(content):
    return list(feedparser.parse(content).entries)


def fetch_feed(url, timeout=15):
    """
    Returns (entries, error, feed_url_used).
    Accepts a real feed URL OR a plain website address: if the page is not a feed,
    the feed is discovered from the page's <link rel="alternate"> tag or common paths.
    """
    from urllib.parse import urljoin, urlparse

    first_err, page_html = None, ""
    try:
        resp = requests.get(url, headers=UA, timeout=timeout)
        resp.raise_for_status()
        entries = _entries_from(resp.content)
        if entries:
            return entries, None, url
        page_html = resp.text
    except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as err:
        return [], f"cannot connect ({str(err)[:80]})", None  # site down: don't guess paths
    except Exception as err:  # HTTP error such as 404: the feed path may just be wrong
        first_err = str(err)[:120]

    candidates = []
    for m in re.finditer(r"<link[^>]+>", page_html, flags=re.I):
        tag = m.group(0)
        if re.search(r"type=[\"']application/(rss|atom)\+xml", tag, flags=re.I):
            h = re.search(r"href=[\"']([^\"']+)", tag, flags=re.I)
            if h:
                candidates.append(urljoin(url, html.unescape(h.group(1))))
    parsed = urlparse(url)
    base = f"{parsed.scheme}://{parsed.netloc}/"
    candidates += [urljoin(base, p) for p in FEED_PATHS]
    tried = set()
    for c in candidates:
        if c in tried or c == url:
            continue
        tried.add(c)
        try:
            r = requests.get(c, headers=UA, timeout=timeout)
            if r.ok:
                entries = _entries_from(r.content)
                if entries:
                    return entries, None, c
        except Exception:
            continue
    return [], first_err or "no RSS feed found on this site", None


def entry_datetime(entry):
    for attr in ("published_parsed", "updated_parsed"):
        t = entry.get(attr)
        if t:
            try:
                return datetime.datetime.fromtimestamp(
                    calendar.timegm(t), datetime.timezone.utc
                )
            except Exception:
                pass
    return None


AUDIO_EXT = (".mp3", ".m4a", ".aac", ".ogg", ".wav")
VIDEO_EXT = (".mp4", ".m4v", ".mov", ".webm")
VIDEO_HOSTS = ("youtube.com", "youtu.be", "vimeo.com")


def extract_media(entry, link, raw_html):
    """Only returns media links the feed itself provides. Never guessed."""
    audio, video = [], []

    def classify(href, typ=""):
        if not href or not href.startswith("http"):
            return
        h, t = href.lower().split("?")[0], (typ or "").lower()
        if t.startswith("audio") or h.endswith(AUDIO_EXT):
            audio.append(href)
        elif t.startswith("video") or h.endswith(VIDEO_EXT) or any(v in h for v in VIDEO_HOSTS):
            video.append(href)

    for enc in entry.get("enclosures", []) or []:
        classify(enc.get("href"), enc.get("type"))
    for mc in entry.get("media_content", []) or []:
        classify(mc.get("url"), mc.get("type") or mc.get("medium"))
    classify(link)
    for u in re.findall(r"https?://[^\s\"'<>]+", raw_html or ""):
        if any(v in u.lower() for v in VIDEO_HOSTS) or u.lower().split("?")[0].endswith(AUDIO_EXT + VIDEO_EXT):
            classify(u)
    return (audio[0] if audio else None), (video[0] if video else None)


def route_territory(text, fixed=None, allowed=None):
    """Decide which island an item belongs to. Returns a territory key or None."""
    t = (text or "").lower()
    if fixed and fixed not in ("stkitts", "nevis"):
        return fixed
    # St. Kitts vs Nevis: ignore the phrase 'St. Kitts and Nevis' when deciding
    skn_text = t
    for ph in SKN_BOTH_PHRASES:
        skn_text = skn_text.replace(ph, " ")
    is_both = skn_text != t
    nevis_hit = any(k in skn_text for k in T_BY_KEY["nevis"][4])
    kitts_hit = any(k in skn_text for k in T_BY_KEY["stkitts"][4])
    skn_guess = None
    if nevis_hit and not kitts_hit:
        skn_guess = "nevis"
    elif kitts_hit:
        skn_guess = "stkitts"
    elif is_both:
        skn_guess = "stkitts"
    if fixed in ("stkitts", "nevis"):
        # Outlet already known to be local; only decide which island
        return skn_guess if skn_guess else fixed
    if skn_guess and (allowed is None or skn_guess in allowed):
        return skn_guess
    for key, _l, _g, _m, kws, _q in TERRITORIES:
        if key in ("stkitts", "nevis"):
            continue
        if (allowed is None or key in allowed) and any(k in t for k in kws):
            return key
    return None


def google_news_url(query):
    q = f"{query} when:{max(1, NEWS_MAX_AGE_HOURS // 24)}d"
    return f"https://news.google.com/rss/search?q={quote(q)}&hl=en-US&gl=US&ceid=US:en"


def build_source_list():
    """List of (territory_key_or_None, outlet, url, kind)."""
    sources = [(t, o, u, "direct") for t, o, u in DIRECT_FEEDS]
    if USE_GOOGLE_NEWS:
        seen_q = set()
        for key, label, _g, _m, _k, query in TERRITORIES:
            if not query or query in seen_q:
                continue
            seen_q.add(query)
            sources.append((key, f"Google News: {label}", google_news_url(query), "google"))
    return sources


def make_items(entries, fixed_territory, outlet, kind, now, source_ok_keys):
    items = []
    for e in entries:
        title = clean_html(e.get("title", ""))
        link = e.get("link", "") or ""
        raw = e.get("summary", "") or e.get("description", "") or ""
        summary = clean_html(raw)[:500]
        published = entry_datetime(e)
        if not title or not link or published is None:
            continue  # undated items cannot be verified as 'latest'
        if published > now + datetime.timedelta(hours=1):
            continue
        if now - published > datetime.timedelta(hours=NEWS_MAX_AGE_HOURS):
            continue
        item_outlet = outlet
        if kind == "google":
            src = e.get("source")
            src_title = (src.get("title") if isinstance(src, dict) else getattr(src, "title", None)) if src else None
            if src_title:
                item_outlet = src_title
                if title.endswith(" - " + src_title):
                    title = title[: -len(" - " + src_title)].strip()
            summary = "" if norm_title(summary).startswith(norm_title(title)[:40]) else summary
        text = f"{title} {summary}"
        if kind == "google":
            # Google results must actually mention the island (keyword check)
            territory = route_territory(text, allowed=source_ok_keys)
        else:
            territory = route_territory(text, fixed=fixed_territory)
        if territory is None:
            continue
        audio, video = extract_media(e, link, raw)
        items.append({
            "territory": territory, "outlet": item_outlet, "title": title,
            "summary": summary, "link": link.replace("]", "%5D").replace("|", "%7C"),
            "published": published, "audio": audio, "video": video, "kind": kind,
        })
    return items


def gather_news(seen):
    """Fetch all sources in parallel, filter, dedupe, and pick the best items per island."""
    now = datetime.datetime.now(datetime.timezone.utc)
    sources = build_source_list()
    report = {"failed": [], "ok": 0, "empty": []}
    all_items = []

    def work(src):
        key, outlet, url, kind = src
        entries, err, used = fetch_feed(url)
        return src, entries, err

    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as pool:
        for src, entries, err in pool.map(work, sources):
            key, outlet, url, kind = src
            if err:
                print(f"  FEED FAILED: {outlet} ({err})")
                report["failed"].append(f"{outlet}: {err}")
                continue
            report["ok"] += 1
            allowed = {"stkitts", "nevis"} if key in ("stkitts", "nevis") and kind == "google" else (
                {key} if key else None)
            items = make_items(entries, key, outlet, kind, now, allowed)
            print(f"  {outlet}: {len(entries)} entries, {len(items)} fresh+relevant")
            all_items.extend(items)

    # Direct feeds first, then newest first, so duplicates keep the direct (better) link
    all_items.sort(key=lambda i: (i["kind"] != "direct", -i["published"].timestamp()))
    unique, titles = [], []
    for it in all_items:
        k = story_key(it["title"])
        if k in seen:
            continue
        nt = norm_title(it["title"])
        if any(difflib.SequenceMatcher(None, nt, t).ratio() > 0.85 for t in titles):
            continue
        titles.append(nt)
        it["key"] = k
        unique.append(it)

    selected = []
    for key, label, _g, max_items, _k, _q in TERRITORIES:
        pool = sorted((i for i in unique if i["territory"] == key),
                      key=lambda i: -i["published"].timestamp())[:max_items]
        if not pool:
            report["empty"].append(label)
        selected.extend(pool)
    for n, it in enumerate(selected, 1):
        it["id"] = f"N{n}"
    return selected, report


# --- 5. LINK CHECK ------------------------------------------------------------------
def check_link(url):
    """True if reachable, False if not. Google News redirect links are trusted as-is."""
    if "news.google.com" in url:
        return True
    for method in (requests.head, requests.get):
        try:
            r = method(url, timeout=10, headers=UA, allow_redirects=True,
                       **({"stream": True} if method is requests.get else {}))
            ok = r.status_code < 400
            if method is requests.get:
                r.close()
            if ok:
                return True
        except Exception:
            pass
    return False


def check_links(items):
    urls = {i["link"] for i in items} | {i[k] for i in items for k in ("audio", "video") if i[k]}
    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as pool:
        return dict(zip(urls, pool.map(check_link, urls)))


# --- 6. AI TEXT GENERATION (provider chain) ----------------------------------------
BAD_MODEL_WORDS = ("image", "tts", "live", "embed", "audio", "vision", "robotics", "computer-use")


def gemini_candidates(client):
    env = [m.strip() for m in os.environ.get("GEMINI_MODELS", "").split(",") if m.strip()]
    defaults = ["gemini-2.5-flash", "gemini-2.5-pro", "gemini-2.5-flash-lite"]
    found = []
    try:
        for m in client.models.list():
            n = getattr(m, "name", "").replace("models/", "")
            low = n.lower()
            if ("flash" in low or "pro" in low) and not any(b in low for b in BAD_MODEL_WORDS):
                found.append(n)
    except Exception as err:
        print(f"Could not list Gemini models: {err}")
    out = []
    for n in env + defaults + sorted(found):
        if n not in out:
            out.append(n)
    return out[:6]


def gemini_sources(resp):
    out = []
    try:
        for ch in resp.candidates[0].grounding_metadata.grounding_chunks or []:
            w = ch.web
            if w and w.uri and w.uri not in [u for _t, u in out]:
                out.append((w.title or "Source", w.uri))
    except Exception:
        pass
    return out


def generate(prompt, search=False, json_mode=False, max_tokens=8000):
    """Returns dict(text, provider, sources) or None. Search only exists on Gemini."""
    if GEMINI_API_KEY:
        try:
            from google import genai
            from google.genai import types

            client = genai.Client(api_key=GEMINI_API_KEY)
            for g_model in gemini_candidates(client):
                try:
                    kwargs = {}
                    if search:
                        kwargs["tools"] = [{"google_search": {}}]
                    if json_mode:
                        kwargs["response_mime_type"] = "application/json"
                    resp = client.models.generate_content(
                        model=g_model, contents=prompt,
                        config=types.GenerateContentConfig(**kwargs))
                    if getattr(resp, "text", None):
                        print(f"SUCCESS: Gemini ({g_model})")
                        return {"text": resp.text, "provider": f"Gemini ({g_model})",
                                "sources": gemini_sources(resp) if search else []}
                except Exception as err:
                    print(f"Gemini model {g_model} failed: {err}")
        except Exception as err:
            print(f"Gemini initialisation failed: {err}")

    if OPENAI_API_KEY:
        try:
            from openai import OpenAI

            client = OpenAI(api_key=OPENAI_API_KEY, timeout=60.0, max_retries=2)
            kwargs = {"response_format": {"type": "json_object"}} if json_mode else {}
            resp = client.chat.completions.create(
                model=OPENAI_MODEL, messages=[{"role": "user", "content": prompt}], **kwargs)
            text = resp.choices[0].message.content
            if text:
                print("SUCCESS: OpenAI")
                return {"text": text, "provider": f"OpenAI ({OPENAI_MODEL})", "sources": []}
        except Exception as err:
            print(f"OpenAI failed: {err}")

    if ANTHROPIC_API_KEY:
        try:
            import anthropic

            client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
            for model_name in ANTHROPIC_MODELS:
                try:
                    resp = client.messages.create(
                        model=model_name, max_tokens=max_tokens,
                        messages=[{"role": "user", "content": prompt}])
                    if resp.content and resp.content[0].text:
                        print(f"SUCCESS: Anthropic ({model_name})")
                        return {"text": resp.content[0].text,
                                "provider": f"Anthropic ({model_name})", "sources": []}
                except Exception as err:
                    print(f"Anthropic model {model_name} failed: {err}")
        except Exception as err:
            print(f"Anthropic failed: {err}")
    return None


def parse_json_block(text):
    text = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.M).strip()
    try:
        return json.loads(text)
    except Exception:
        m = re.search(r"\{.*\}", text, flags=re.S)
        return json.loads(m.group(0)) if m else {}


def spoken_date(dt_utc):
    local = dt_utc.astimezone(AST)
    return f"{local.strftime('%A, %B')} {local.day}"


def footnote_date(dt_utc):
    local = dt_utc.astimezone(AST)
    return f"{local.strftime('%A, %B')} {local.day}, {local.year}"


def clean_link(url):
    return (url or "").replace("]", "%5D").replace("|", "%7C")


# ---- Full article text (so stories can be written in full, from facts only) ----
def resolve_google_link(url):
    """Google News links are redirects; decode them to the real article address."""
    if "news.google.com" not in url:
        return url
    try:
        from googlenewsdecoder import gnewsdecoder

        r = gnewsdecoder(url)
        if r.get("status") and r.get("decoded_url"):
            return r["decoded_url"]
    except Exception as err:
        print(f"  could not decode Google link: {str(err)[:80]}")
    return url


def extract_article_text(page_html):
    text = ""
    try:
        import trafilatura

        text = trafilatura.extract(page_html, include_comments=False, include_tables=False,
                                   favor_precision=True) or ""
    except Exception:
        pass
    if len(text) < 300:  # fallback: collect the page's longer paragraphs
        paras = [clean_html(p) for p in re.findall(r"<p[^>]*>(.*?)</p>", page_html, flags=re.S | re.I)]
        text = "\n".join(p for p in paras if len(p) > 60)
    return text[:ARTICLE_CHARS]


def media_from_html(page_html):
    """Audio/video embedded inside the article itself (never from sidebars)."""
    m = re.search(r"<article.*?</article>", page_html, flags=re.S | re.I)
    if not m:
        return None, None
    body = m.group(0)
    video = audio = None
    y = re.search(r"youtube(?:-nocookie)?\.com/embed/([\w-]{11})", body)
    if y:
        video = f"https://www.youtube.com/watch?v={y.group(1)}"
    else:
        v = re.search(r"https?://[^\s\"'<>]+\.(?:mp4|m4v|webm)", body)
        video = v.group(0) if v else None
    a_ = re.search(r"https?://[^\s\"'<>]+\.(?:mp3|m4a|aac)", body)
    audio = a_.group(0) if a_ else None
    return audio, video


def fetch_article(it):
    it["link"] = clean_link(resolve_google_link(it["link"]))
    it["article"] = ""
    if "news.google.com" in it["link"]:
        return  # could not decode: we only have the headline/summary for this one
    try:
        r = requests.get(it["link"], headers=UA, timeout=15)
        r.raise_for_status()
        it["article"] = extract_article_text(r.text)
        audio, video = media_from_html(r.text)
        it["audio"] = it["audio"] or (clean_link(audio) if audio else None)
        it["video"] = it["video"] or (clean_link(video) if video else None)
    except Exception as err:
        print(f"  could not read article {it['link'][:70]}: {str(err)[:60]}")


def enrich_articles(selected):
    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as pool:
        list(pool.map(fetch_article, selected))


# ---- Writing the broadcast copy ----
def clean_ai_text(text):
    """Strips URLs and brackets but keeps paragraph breaks."""
    out = []
    for para in re.split(r"\n\s*\n", text or ""):
        para = re.sub(r"https?://\S+", "", para).replace("[", "(").replace("]", ")")
        para = re.sub(r"\s+", " ", para).strip()
        if para:
            out.append(para)
    return "\n".join(out)


def word_count(text):
    return len(re.findall(r"\b[\w'’-]+\b", text))


def trt(words):
    secs = int(round(words / 2.5))  # about 150 words per minute
    return f"{secs // 60}:{secs % 60:02d}"


MONTHS_DAYS = set("""monday tuesday wednesday thursday friday saturday sunday january february march
april may june july august september october november december""".split())


def unverified_terms(script, source_text, allow):
    """Capitalised names/terms in the script that do not appear in the source article."""
    src = source_text.lower()
    flagged = []
    for sent in re.split(r"(?<=[.!?])\s+", script.replace("\n", " ")):
        words = re.findall(r"[A-Za-z][A-Za-z'’\-]*", sent)
        for w in words[1:]:
            lw = w.lower()
            if w[0].isupper() and len(w) >= 4 and lw not in MONTHS_DAYS and lw not in allow:
                if lw not in src and lw.rstrip("s") not in src:
                    flagged.append(w)
    return sorted(set(flagged))[:8]


def write_news_copy(selected):
    """
    AI writes finished on-air news items from the REAL article text.
    Returns ({id: {headline, script, full}}, providers).
    """
    copy, providers = {}, []
    for start in range(0, len(selected), 5):
        batch = selected[start:start + 5]
        payload = []
        for i in batch:
            full = len(i.get("article", "")) >= 400
            local = i["territory"] in ("stkitts", "nevis")
            payload.append({
                "id": i["id"], "island": T_BY_KEY[i["territory"]][1], "outlet": i["outlet"],
                "published_day": spoken_date(i["published"]), "headline": i["title"],
                "source_text": i["article"] if full else (i["summary"] or "(headline only)"),
                "full_article_available": full,
                "target_words": ("110 to 190" if local else "70 to 130") if full else "35 to 55",
            })
        prompt = (
            "You are the senior news writer at 2020 Vision Radio 102.1 FM in St. Kitts and Nevis. "
            "Write each story below as a FINISHED, broadcast-ready news item that the newscaster "
            "will read on air exactly as written, with no editing.\n\n"
            "ACCURACY (most important):\n"
            "- Use ONLY facts found in that story's source_text and headline. Never add names, "
            "titles, numbers, dates, places, causes, background or quotes that are not there. "
            "Do not guess. If the source says 'alleged', 'police say' or 'reportedly', keep that "
            "wording.\n"
            "- If full_article_available is false you only have a headline or short summary: write "
            "a brief item of the target length from those facts alone.\n\n"
            "STYLE (write for the ear, in professional broadcast style):\n"
            "- Lead with the most important news in the first sentence, then the key details, then "
            "any context or what happens next. The item must flow naturally from start to end as "
            "one continuous piece, not a list of facts or a summary.\n"
            "- Short, clear, conversational sentences (usually under 20 words). Active voice. "
            "Use short paragraphs of two to four sentences.\n"
            "- Give the attribution naturally in the first or second sentence, e.g. 'the St. Kitts "
            "and Nevis Information Service says'. Refer to the day using published_day (for example "
            "'on Thursday'). Never say 'yesterday', 'today' or 'tonight'.\n"
            "- For any island other than St. Kitts or Nevis, name the island in the first sentence.\n"
            "- Do not use direct quotations; report what was said in your own words and name who "
            "said it.\n"
            "- Spell out numbers, money and abbreviations the way they are spoken ('five thousand "
            "dollars', 'Prime Minister', 'percent'). Write names exactly as in the source.\n"
            "- No URLs, square brackets, parentheses, bullet points, stage directions or sign-offs. "
            "Do not address the newscaster. Do not end with a teaser.\n"
            "- Aim for the item's target_words.\n"
            "- Also write a short headline slug of at most 8 words in capital letters.\n\n"
            "Return ONLY JSON: {\"items\": [{\"id\": \"N1\", \"headline\": \"...\", "
            "\"script\": \"paragraph one\\n\\nparagraph two\"}]}\n\n"
            f"STORIES:\n{json.dumps(payload, ensure_ascii=False, indent=1)}"
        )
        result = generate(prompt, json_mode=True, max_tokens=8000)
        if not result:
            continue
        try:
            data = parse_json_block(result["text"])
            ids = {i["id"]: i for i in batch}
            for row in data.get("items", []):
                it = ids.get(row.get("id"))
                if it and row.get("script"):
                    copy[it["id"]] = {
                        "headline": clean_ai_text(row.get("headline", "")).upper() or it["title"].upper(),
                        "script": clean_ai_text(row["script"]),
                        "full": len(it.get("article", "")) >= 400,
                    }
            providers.append(result["provider"])
        except Exception as err:
            print(f"Could not parse news copy JSON: {err}")
    return copy, providers


def write_weather_sports(date_str, friday_next):
    prompt = f"""You are an expert Caribbean radio presenter writing for 2020 Vision Radio 102.1 FM.
Tomorrow's broadcast date is {date_str}. The broadcast starts at 8:00 AM.

Write TWO segments, using exactly these headings:
=== SCRIPT 2: WEATHER SEGMENT ===
=== SCRIPT 3: SPORTS SEGMENT ===

WEATHER RULES:
- Focus exclusively on St. Kitts and Nevis, forecast for tomorrow ({date_str}) ONLY. No multi-day outlook.
{"- Tomorrow is Friday, so add a dedicated 'Weekend Outlook' section right after Friday's forecast covering Saturday and Sunday." if friday_next else ""}
- Include an island summary (Basseterre, Sandy Point, Nevis Peak, Charlestown) and marine/swell conditions. Use real data from your live search only. If you cannot find a real figure (for example tide times), say it is not available instead of guessing.

SPORTS RULES:
- Lead with St. Kitts and Nevis local sport, then Leeward Islands cricket, CWI, CONCACAF and regional sport. Report only real, recent, verifiable results. If nothing recent is found for an item, leave it out.

GENERAL RULES:
- Plain, spoken radio English. Do not invent people, scores or numbers. No URLs. Do not mention timings or word counts.
- Use bracketed cues like [pause] and [transition sting] where helpful.
"""
    return generate(prompt, search=True)


# --- 7. HISTORY (so the same story is not repeated tomorrow) ------------------------
def load_seen():
    if not os.path.exists(HISTORY_FILE):
        return {}
    try:
        with open(HISTORY_FILE, "r") as f:
            data = json.load(f).get("seen", {})
        cutoff = (datetime.date.today() - datetime.timedelta(days=HISTORY_DAYS)).isoformat()
        return {k: v for k, v in data.items() if v >= cutoff}
    except Exception:
        return {}


def save_seen(seen, selected):
    today_iso = datetime.datetime.now(AST).date().isoformat()
    for it in selected:
        seen[it["key"]] = today_iso
    try:
        with open(HISTORY_FILE, "w") as f:
            json.dump({"seen": seen}, f, indent=1)
    except Exception as err:
        print(f"Could not save history: {err}")


# --- 8. FEED CHECKER ------------------------------------------------------------------
def check_feeds():
    now = datetime.datetime.now(datetime.timezone.utc)
    print(f"Checking {len(build_source_list())} sources...\n")
    for key, outlet, url, kind in build_source_list():
        entries, err, used = fetch_feed(url)
        if err:
            print(f"[FAIL ] {outlet}\n        {url}\n        {err}")
            continue
        if kind == "direct" and used != url:
            print(f"        (feed found for {outlet}: {used})")
        dated = [entry_datetime(e) for e in entries]
        dated = [d for d in dated if d]
        newest = max(dated).astimezone(AST).strftime("%b %d %H:%M") if dated else "no dates"
        fresh = sum(1 for d in dated if now - d <= datetime.timedelta(hours=NEWS_MAX_AGE_HOURS))
        flag = "OK   " if fresh else "STALE"
        print(f"[{flag}] {outlet}: {len(entries)} entries, newest {newest}, "
              f"{fresh} within {NEWS_MAX_AGE_HOURS}h")
    print("\nRemove or fix any FAIL/STALE direct feeds in DIRECT_FEEDS, or add new ones with "
          "the EXTRA_FEEDS setting.")


# --- 9. MAIN --------------------------------------------------------------------------
def footnote_line(it, link_ok):
    """Footnote under each story: source link with the story's date in brackets."""
    text = f"^ Source: {it['outlet']}: {it['link']} [{footnote_date(it['published'])}]"
    if it["audio"]:
        text += f"   Audio: {it['audio']}"
    if it["video"]:
        text += f"   Video: {it['video']}"
    if not link_ok.get(it["link"], True):
        text += "   (link not confirmed - verify before airing)"
    return text


def main():
    if "--check-feeds" in sys.argv:
        check_feeds()
        return

    now_ast = datetime.datetime.now(AST)
    tomorrow = now_ast.date() + datetime.timedelta(days=1)
    date_str = tomorrow.strftime("%A, %B %d, %Y").upper()
    friday_next = tomorrow.weekday() == 4
    notes = []

    # --- News ---
    print("Collecting news...")
    seen = load_seen()
    selected, report = gather_news(seen)
    print(f"Selected {len(selected)} news items. Reading full articles...")
    enrich_articles(selected)
    link_ok = check_links(selected) if selected else {}
    bad_links = [u for u, ok in link_ok.items() if not ok]

    copy, providers = ({}, [])
    if selected:
        print("Writing news items...")
        copy, providers = write_news_copy(selected)
    missing = [i for i in selected if i["id"] not in copy]
    if missing:
        notes.append(f"{len(missing)} story(ies) could NOT be written by the AI and appear as the "
                     "raw headline and summary. Edit before airing: "
                     + "; ".join(i["title"] for i in missing))
    short = [i for i in selected if i["id"] in copy and not copy[i["id"]]["full"]]
    if short:
        notes.append("Short items (the full article could not be read, so only the headline/summary "
                     "was available): " + "; ".join(f"{i['title']} ({i['outlet']})" for i in short))
    for i in selected:
        c = copy.get(i["id"])
        if not c:
            continue
        allow = set(re.findall(r"[a-z']+", f"{i['outlet']} {T_BY_KEY[i['territory']][1]}".lower()))
        allow |= {"caribbean", "kitts", "nevis", "saint", "federation", "information", "service"}
        odd = unverified_terms(c["script"], f"{i['title']} {i['summary']} {i.get('article', '')}", allow)
        if odd:
            notes.append(f"CHECK '{c['headline']}': these names/terms are not in the source "
                         f"article: {', '.join(odd)}.")

    LEAD_IN = {
        "stkitts": "We begin close to home, here in St. Kitts.",
        "nevis": "Now to our sister island, Nevis.",
    }
    GROUP_LEAD = {
        "neighbours": "Turning now to our neighbours around the region.",
        "wider": "And now, news from across the wider Caribbean.",
    }
    intro = (f"Good morning, and welcome to the news on 2020 Vision Radio, one-oh-two point one FM. "
             f"It is {tomorrow.strftime('%A, %B')} {tomorrow.day}. Here are the stories we are following.")
    outro = ("That is the latest news from St. Kitts, Nevis and across the Caribbean. "
             "Stay with us for the weather and sports.")
    body_lines, total_words = [], word_count(intro) + word_count(outro)
    groups_done, has_kitts = set(), any(i["territory"] == "stkitts" for i in selected)
    for key, label, group, _m, _k, _q in TERRITORIES:
        items = [i for i in selected if i["territory"] == key]
        if not items:
            continue
        body_lines.append(f"## {label}")
        if key == "stkitts":
            lead = LEAD_IN["stkitts"]
        elif key == "nevis":
            lead = LEAD_IN["nevis"] if has_kitts else "We begin close to home, with news from Nevis."
        elif group not in groups_done:
            lead = GROUP_LEAD[group]
        else:
            lead = None
        groups_done.add(group)
        if lead:
            body_lines.append(lead)
            total_words += word_count(lead)
        for it in items:
            c = copy.get(it["id"])
            if c:
                headline, script = c["headline"], c["script"]
            else:
                headline = it["title"].upper()
                script = (f"{it['title']}. {it['summary']}").strip()
            words = word_count(script)
            total_words += words
            body_lines += [f"# {headline}  (TRT {trt(words)})"] + script.split("\n") + [footnote_line(it, link_ok)]
    lines = [f"=== SCRIPT 1: NEWS SEGMENT (approx. {trt(total_words)}) ===", "[station intro sting]",
             intro, "[pause]"]
    if selected:
        lines += body_lines + [outro]
    else:
        lines.append("No fresh, verifiable news items were found today. Please check the sources manually.")
        notes.append("NO news items passed the freshness/relevance filters. Run --check-feeds.")
    lines += ["###", "[transition sting]"]

    # --- Weather & sports ---
    print("Generating weather and sports...")
    ws = write_weather_sports(date_str, friday_next)
    if ws:
        lines += ws["text"].splitlines()
        if ws["sources"]:
            lines.append("[Source: Weather/sports reference links (from the AI's live search) | "
                         + " | ".join(f"{t}: {u}" for t, u in ws["sources"][:8]) + "]")
        else:
            notes.append(f"Weather and sports came from {ws['provider']} WITHOUT live web search "
                         "sources. Treat as UNVERIFIED and check before airing.")
    else:
        lines += ["=== SCRIPT 2: WEATHER SEGMENT ===",
                  "Weather could not be generated today. Please add manually.",
                  "=== SCRIPT 3: SPORTS SEGMENT ===",
                  "Sports could not be generated today. Please add manually."]
        notes.append("Weather and sports generation FAILED on all AI providers.")

    if not selected and not ws:
        print("CRITICAL ERROR: no news and no AI output.")
        sys.exit(1)

    # --- Producer notes (not for air) ---
    if report["empty"]:
        notes.append("No fresh items found for: " + ", ".join(report["empty"]) + ".")
    if report["failed"]:
        names = ", ".join(f.split(":")[0] for f in report["failed"])
        notes.append(f"{len(report['failed'])} news feed(s) failed to load: {names}.")
    if bad_links:
        notes.append(f"{len(bad_links)} link(s) did not respond and are flagged in the script.")
    notes.append(f"News items are from the last {NEWS_MAX_AGE_HOURS} hours. Footnote dates are "
                 "the publication date of each source. Stories are written only from the article "
                 "text; still click the source link to confirm anything important before airing.")

    # --- Build Word document ---
    doc = docx.Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(13)  # larger text is easier to read on air
    normal.paragraph_format.line_spacing = 1.5  # script-style spacing
    normal.paragraph_format.space_after = Pt(6)
    for section in doc.sections:
        section.top_margin = section.bottom_margin = Inches(1)
        section.left_margin = section.right_margin = Inches(1)

    tbl = doc.add_table(rows=1, cols=1)
    tbl.autofit = False
    tbl.columns[0].width = Inches(6.5)
    cell = tbl.cell(0, 0)
    cell._tc.get_or_add_tcPr().append(parse_xml(r'<w:shd {} w:fill="1B365D"/>'.format(nsdecls("w"))))
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p.add_run("2020 VISION RADIO 102.1 FM - DAILY BROADCAST SCRIPT")
    r1.font.size, r1.font.bold, r1.font.color.rgb = Pt(16), True, RGBColor(0xFF, 0xFF, 0xFF)
    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(f"FOR BROADCAST: {date_str} (News | Weather | Sports)")
    r2.font.size, r2.font.bold, r2.font.color.rgb = Pt(11), True, RGBColor(0xD0, 0xE1, 0xF9)

    for line in lines:
        render_line(doc, line)

    doc.add_page_break()
    doc.add_paragraph().add_run("PRODUCER NOTES (NOT FOR AIR)").bold = True
    for n in notes:
        doc.add_paragraph("- " + n)

    output_filename = f"Broadcast_Script_{tomorrow.strftime('%Y%m%d')}.docx"
    doc.save(output_filename)
    print(f"Successfully generated script: {output_filename}")
    save_seen(seen, selected)

    # --- Email ---
    if SENDER_EMAIL and SENDER_PASSWORD and recipients:
        print(f"Sending email to: {', '.join(recipients)}...")
        try:
            msg = MIMEMultipart()
            msg["From"] = SENDER_EMAIL
            msg["To"] = ", ".join(recipients)
            msg["Subject"] = f"2020 Vision Radio Script - {date_str}"
            body = (f"Hello,\n\nPlease find attached the daily broadcast script for {date_str}.\n"
                    f"It contains {len(selected)} dated news items with clickable source links.\n\n"
                    "Producer notes:\n" + "\n".join("- " + n for n in notes) +
                    "\n\nBest regards,\n2020 Vision Radio Automation")
            msg.attach(MIMEText(body, "plain"))
            with open(output_filename, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={output_filename}")
            msg.attach(part)
            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login(SENDER_EMAIL, SENDER_PASSWORD)
                server.sendmail(SENDER_EMAIL, recipients, msg.as_string())
            print("Email sent successfully!")
        except Exception as email_err:
            print(f"ERROR: Failed to send email: {email_err}")
            traceback.print_exc()
    else:
        print("Notice: Email skipped. Check SENDER_EMAIL, SENDER_PASSWORD and PRESENTER_EMAIL secrets.")


if __name__ == "__main__":
    main()
