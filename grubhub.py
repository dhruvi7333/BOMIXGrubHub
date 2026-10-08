# Grubhub scraper v3
#  - FIX: timezone glitch (zipcodes 'timezone' is an IANA string, not an int offset)
#  - NEW: closed stores are no longer skipped -> details/hours are scraped
#  - NEW: stores that show CLOSED but are inside their posted hours are FLAGGED
#  - NEW (v3): duplicate-store protection
#       * Pre-scrape:  dedup input.xlsx rows by Grubhub store ID (from the URL)
#       * Post-scrape: hard-dedup by store ID or physical address (auto-removed)
#       * Post-scrape: soft-check by name+zipcode (logged only, never auto-removed -
#         chains can legitimately have two branches in the same zip)
import pandas as pd
import re
import time
import sys
import os
from pathlib import Path
from datetime import datetime, time as dtime   # ── CHANGE: also import time class
import pytz
import openpyxl
import zipcodes

import smtplib
import platform            # ── NEW (v5): machine name for run visibility
import socket
import getpass
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import logging
from logging.handlers import RotatingFileHandler

# ── CHANGE: file logging so a client's unattended runs leave a trace you can inspect ──
log = logging.getLogger("grubhub")

CONSOLE_VERBOSE = True     # ── NEW (v5): False = quiet console for unattended
                           # runs. The .log file is written either way.


def setup_logging():
    """Console + rotating file logging. The .log sits next to the .exe / input.xlsx.
       Rotates at ~2 MB, keeps 5 backups, so it never grows without bound."""
    if log.handlers:                       # already configured -> safe to call again
        return
    log_path = app_dir() / "grubhub_scraper.log"
    log.setLevel(logging.DEBUG)      # ── v5.1: file keeps everything...
    fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%Y-%m-%d %H:%M:%S")
    try:
        fh = RotatingFileHandler(log_path, maxBytes=2_000_000, backupCount=5, encoding="utf-8")
        fh.setFormatter(fmt)
        fh.setLevel(logging.DEBUG)   # ...so the .log is unchanged, still full detail
        log.addHandler(fh)
    except Exception as e:
        print(f"WARNING: could not open log file {log_path}: {e}")
    ch = logging.StreamHandler(sys.stdout)
    # ── NEW (v5): back to INFO with a bare formatter, so the run narrates on
    #    screen the way it used to. Full timestamps still go to the .log file.
    ch.setFormatter(logging.Formatter("%(message)s") if CONSOLE_VERBOSE else fmt)
    ch.setLevel(logging.INFO if CONSOLE_VERBOSE else logging.WARNING)
    log.addHandler(ch)
    log.info("==== Grubhub scraper started ==== (log file: %s)", log_path)
    log.info("Running on: %s", machine_name())


# ============================================================
# ── CHANGE (v3): duplicate-detection helpers ──
# ============================================================
# Grubhub URLs always end in a numeric restaurant ID, with or without a slug
# in front of it, e.g.:
#   http://grubhub.com/restaurant/12276808
#   https://www.grubhub.com/restaurant/bomi--flowers-plants--gifts-.../12276824
_STORE_ID_RE = re.compile(r'/restaurant/(?:[^/]+-)?(\d+)(?:[/?].*)?$')


def extract_store_id(url):
    """Pull the trailing numeric Grubhub restaurant ID out of a URL, so two
       different URL formats pointing at the same store dedup correctly.
       Falls back to the normalized full URL if no ID pattern is found."""
    m = _STORE_ID_RE.search(str(url).strip())
    if m:
        return m.group(1)
    return str(url).strip().rstrip('/').lower()


def clean_zip_5(zip_val):
    """Same 5-digit zip cleanup as get_location_info(), pulled out so it can
       also be used for building dedup keys."""
    zip_str = str(zip_val).split('.')[0]
    z = re.sub(r'[^0-9]', '', zip_str).strip()
    if len(z) > 5:
        z = z[:5]
    elif len(z) == 4:
        z = z.zfill(5)
    return z if len(z) == 5 else ""


def _norm(s):
    """Lowercase, strip to alphanumerics only, so formatting differences like
       'BOMI- Flowers, Plants, & Gifts' vs 'bomi flowers plants gifts' compare equal."""
    return re.sub(r'[^a-z0-9]', '', str(s).lower())
# ============================================================


# ============================================================
# ── CHANGE: discrepancy-detection config + helpers ──
# ============================================================
NO_MENU_LABEL = "NO MENU"       # ── NEW (v5.3): the storefront has no menu to sell
                                # from. Detected ONLY when BOTH hold: the About
                                # tab is absent AND the page says the menu is not
                                # available. Either one alone can be a transient
                                # page-load hiccup; together they are the real thing.
PAUSE_LABEL = "PAUSED"          # ── NEW (v5): the store is switched OFF on Grubhub.
                                # Set ONLY when the storefront is taking preorders
                                # only / shows Closed AND the clock says it is inside
                                # the store's own posted open hours.
                                # Everything else below is untouched.
FLAG_LABEL = "FLAGGED"          # the word placed in the Status column for a mismatch.
                                # rename to taste, e.g. "REVIEW" / "SHOULD BE OPEN" /
                                # "MISMATCH". Keep it WITHOUT the word "closed".
STRICT_WINDOW_CHECK = True      # True  -> flag a closed store ONLY if it is inside its
                                #          posted hours right now (no false alarms for
                                #          stores that are simply before opening / after
                                #          closing at scrape time). RECOMMENDED.
                                # False -> flag ANY closed store that still lists
                                #          pickup/delivery hours (your "basic" version).

# IANA timezone -> your display bucket
IANA_TO_GROUP = {
    "America/New_York": "EST", "America/Detroit": "EST", "America/Indiana/Indianapolis": "EST",
    "America/Chicago": "CT", "America/Indiana/Knox": "CT", "America/Menominee": "CT",
    "America/Denver": "MT", "America/Boise": "MT", "America/Phoenix": "MT",
    "America/Los_Angeles": "PT",
    "America/Anchorage": "AKT", "Pacific/Honolulu": "HT",
}
# display bucket -> a representative zone (used only for the section-header timestamp)
GROUP_TO_IANA = {
    "EST": "US/Eastern", "CT": "US/Central", "MT": "US/Mountain",
    "PT": "US/Pacific", "AKT": "US/Alaska", "HT": "US/Hawaii",
}
GROUP_ORDER = {"EST": 0, "CT": 1, "MT": 2, "PT": 3, "AKT": 4, "HT": 5}


def _group_from_iana(iana):
    """Fallback bucket for any IANA zone not in the table above (by current UTC offset)."""
    try:
        off = datetime.now(pytz.timezone(iana)).utcoffset().total_seconds() / 3600
    except Exception:
        return "EST"
    return {-5: "EST", -6: "CT", -7: "MT", -8: "PT", -9: "AKT", -10: "HT"}.get(round(off), "EST")


_TIME_RE = re.compile(r"(\d{1,2})(?::(\d{2}))?\s*([ap])\.?m", re.I)


def _parse_clock(token):
    m = _TIME_RE.search(token)
    if not m:
        return None
    h = int(m.group(1)); mn = int(m.group(2) or 0); ap = m.group(3).lower()
    if ap == "p" and h != 12:
        h += 12
    if ap == "a" and h == 12:
        h = 0
    return dtime(h, mn)


def parse_windows_from_timings(text):
    """Scraped Timings cell -> list of (open, close) windows.
       Handles 'Pickup:'/'Delivery:' labels, en/em dashes, split shifts, and 'Closed'."""
    if not text:
        return []
    raw = str(text).replace("\u2013", "-").replace("\u2014", "-")
    out = []
    # break on newlines AND on the service labels, so each service line is handled
    for seg in re.split(r"(?i)pickup\s*:|delivery\s*:|\n", raw):
        seg = seg.strip()
        if not seg or "closed" in seg.lower():
            continue
        for chunk in seg.split(","):                       # split shifts (10-2, 5-9)
            if "-" not in chunk:
                continue
            a, b = chunk.split("-", 1)
            ta, tb = _parse_clock(a), _parse_clock(b)
            if ta and tb:
                out.append((ta, tb))
    return out


def is_within_hours(windows, iana_tz, now_utc):
    if not windows:
        return False
    cur = now_utc.astimezone(pytz.timezone(iana_tz)).time()
    for o, c in windows:
        if c > o:                                          # normal same-day window
            if o <= cur < c:
                return True
        else:                                              # window crosses midnight (7am-1:45am)
            if cur >= o or cur < c:
                return True
    return False


def classify_store(live_closed, timings_text, iana_tz, now_utc, paused_evidence="",
                   no_menu_evidence=""):
    """Returns (status_label, flag_reason).

    ── NEW (v5): the PAUSED branch, and nothing else. ──
    paused_evidence is non-empty when the storefront itself says it is not
    taking orders right now - the "Preorder for 4:00pm." banner, or the word
    Closed in that same spot, with neither Delivery nor Pickup live. If that
    is true WHILE the store is inside its own posted open hours, somebody has
    switched the store off: PAUSED.

    Outside posted hours the same banner is normal (a shut shop offering a
    future slot), so it falls through to the original logic untouched.
    """
    windows = parse_windows_from_timings(timings_text)
    within_hours = is_within_hours(windows, iana_tz, now_utc) if windows else False

    # ── NEW (v5.3): no menu at all. Checked first: a storefront with nothing
    #    to sell cannot trade whatever the clock says, and with no About tab
    #    there are no posted hours to judge open/closed against anyway.
    if no_menu_evidence:
        return NO_MENU_LABEL, f"No About tab and the page says: {no_menu_evidence}"

    if paused_evidence and within_hours:
        return PAUSE_LABEL, f"Store is switched off on Grubhub ({paused_evidence}) " \
                            f"while inside its posted open hours"

    # ---- everything from here down is the original logic, unchanged ----
    if not live_closed:
        return "Open", "Live status: open"
    if not windows:                                        # no hours OR 'Closed' in hours
        return "Closed", "Storefront closed and no posted hours today - expected"
    if STRICT_WINDOW_CHECK:
        if within_hours:
            return FLAG_LABEL, "Storefront shows CLOSED during posted open hours -> owner should check"
        return "Closed", "Closed now, but it is outside today's posted hours - expected"
    return FLAG_LABEL, "Storefront closed but still lists pickup/delivery hours -> check"


# ============================================================
# ── NEW (v5): reading the two things that prove a store is PAUSED ──
# ------------------------------------------------------------
# Nothing here keys off position or CSS class. The absolute path
#   /html/body/ghs-site-container/span/div/div[4]/.../div[2]/span/div
# breaks the moment Grubhub adds a wrapper, and the classes are
# styled-components hashes that change on every deploy:
#   <span class="sc-fzpans sc-fzplWN hKcEsM">Preorder for 4:00pm.</span>
# So each read anchors on TEXT Grubhub cannot change without changing what
# the page says, then walks the ancestor axis to the node it needs.
# ============================================================

# 1. The banner. Same spot shows "Preorder for 4:00pm." or just "Closed".
# 1. THE PREORDER ASK. This, and only this, is what makes a store PAUSED.
#    A store that merely says "Closed" is NOT paused - it is FLAGGED when it
#    should be open, and Closed when it should be closed. Matching the bare
#    word "Closed" here is what made flagged stores show up as paused.
_PAUSE_BANNER_XPATHS = [
    "//span[contains(text(),'Preorder for')]",
    "//*[contains(text(),'Preorder for')]",
    "//*[@id='navSection-categories']//span[contains(text(),'Preorder for')]",
]

# 2. The Delivery/Pickup tab pair: find the 'Delivery' label, then climb to
#    the NEAREST ancestor that also contains 'Pickup'. That is the lowest
#    common ancestor of the two tabs - the container the absolute path ends
#    at - reached without counting a single div. ancestor:: is a reverse
#    axis, so [1] is the closest one, not the outermost.
_TABS_XPATHS = [
    "//*[normalize-space(text())='Delivery']/ancestor::*[contains(.,'Pickup')][1]",
    "//*[normalize-space(text())='Pickup']/ancestor::*[contains(.,'Delivery')][1]",
    "//*[contains(text(),'Delivery')]/ancestor::*[contains(.,'Pickup')][1]",
]

# "Preorder for 4:00pm." / "Preorder for 1:30am tomorrow." / "...on Sep 14."
# The TIME is required: a bare "Preorder" button sits on plenty of pages.
_PREORDER_TEXT_RE = re.compile(r"pre-?order\s+for\s+\d{1,2}(?::\d{2})?\s*[ap]\.?m", re.I)

# "Out of range" / "doesn't deliver to your address" describe the MISSING
# LOCATION PIN, not the store. The scraper never sets an address, so healthy
# open stores show these. They must never read as a channel being live OR shut.
_ADDRESS_NOISE_RE = re.compile(
    r"out of range|too far|doesn'?t deliver|does not deliver|update address|enter your address",
    re.I)

# A channel that is actually serving: "25-35 min", "30 min", "ASAP", "Available".
_CHANNEL_LIVE_RE = re.compile(r"\d+\s*-\s*\d+\s*min|\b\d+\s*min\b|\basap\b|\bavailable\b", re.I)


def _el_text(el, cap=220):
    try:
        t = (el.get_attribute("innerText") or el.text or "").strip()
    except Exception:
        return ""
    return re.sub(r"\s*\n\s*", " ", t)[:cap]


# "This menu isn't available right now" - the wording under the hero when a
# storefront has nothing to sell. Curly and straight apostrophes both appear.
_NO_MENU_RE = re.compile(
    r"menu\s+(isn'?t|is\s+not|not)\s+available"
    r"|menu\s+unavailable"
    r"|no\s+menu\s+items?\s+available",
    re.I)


def read_menu_unavailable(driver):
    """Return the on-page wording that says the menu is missing, or "".

    Read from the page text, NOT from a class or position. The caller only
    trusts it when the About tab was also absent.
    """
    try:
        body = driver.find_element(By.TAG_NAME, "body")
        txt = (body.get_attribute("innerText") or body.text or "")
    except Exception:
        return ""
    txt = txt.replace("\u2019", "'")
    m = _NO_MENU_RE.search(txt)
    if not m:
        return ""
    # hand back the whole sentence it sits in, for the log and the reason column
    line = next((l.strip() for l in txt.split("\n") if _NO_MENU_RE.search(l)), m.group(0))
    return line[:80]


# ── NEW (v5.4): rating / review count ──
# The old block used class and data-testid guesses (".rating-value",
# "[data-testid*='rating']"). Those are styled-components hashes on Grubhub,
# so they matched nothing and every store came back 0.0 / 0. This reads the
# header text instead, the same way the banner and the tabs are read.
#
# Grubhub prints "New" in the rating slot for a store with no ratings yet -
# that is a REAL answer, not a failure, so it is reported as "New" rather
# than being silently written out as 0.0.

# "4.4 (14 ratings)" / "4.5 • 200+ ratings" / "4.2 stars, 1,203 reviews"
# The rating MUST carry its decimal point. Without that, "37 reviews" was
# being split into a rating of 3 and a count of 7.
_RATING_PAIR_RE = re.compile(
    r"(?<![\d.])([0-5]\.\d)\s*(?:stars?)?\s*[\(•,\-\|]?\s*([\d,]+)\s*\+?\s*(?:ratings?|reviews?)",
    re.I)
_RATING_ONLY_RE = re.compile(r"(?<![\d.])([0-5]\.\d)(?!\d)\s*(?:stars?|out of 5)?", re.I)
_COUNT_ONLY_RE = re.compile(r"([\d,]+)\s*\+?\s*(?:ratings?|reviews?)", re.I)
# "New" where a rating would otherwise be: "New | 1080 Creekview Ridge Ct".
# Anchored on the separator that FOLLOWS it, not on a newline: the header
# text is flattened to a single line before this runs, so \n anchors never
# match. Requiring the separator also keeps a shop called "New York
# Flowers\" from being mistaken for an unrated store.
_NEW_STORE_RE = re.compile(r"(?<![A-Za-z])New\s*(?:[\|•·]|$)")


def read_rating(driver):
    """(rating, review_count) as strings. ("New", "0") for an unrated store,
       ("0.0", "0") only when nothing could be read at all."""
    # the header block: walk up from <h1> while the text is still small enough
    # to BE the header rather than half the page
    text = ""
    try:
        node = driver.find_element(By.XPATH, "//h1")
        for _ in range(5):
            try:
                node = node.find_element(By.XPATH, "..")
            except Exception:
                break
            t = _el_text(node, 1200)
            if len(t) > 1000:
                break
            text = t
    except Exception:
        text = ""

    if not text:                       # fall back to aria-labels anywhere on the page
        try:
            for el in driver.find_elements(By.CSS_SELECTOR, "[aria-label]")[:80]:
                al = el.get_attribute("aria-label") or ""
                if re.search(r"rating|review|star", al, re.I):
                    text += " | " + al
        except Exception:
            pass
    if not text:
        return "0.0", "0"

    flat = text.replace(",", "")
    m = _RATING_PAIR_RE.search(flat)
    if m:
        return m.group(1), str(int(m.group(2)))

    rating, count = "0.0", "0"
    mc = _COUNT_ONLY_RE.search(flat)
    if mc:
        count = str(int(mc.group(1)))
    mr = _RATING_ONLY_RE.search(flat)
    if mr:
        rating = mr.group(1)

    if rating == "0.0" and count == "0" and _NEW_STORE_RE.search(text):
        return "New", "0"              # genuinely unrated, not a failed read
    return rating, count


def read_pause_evidence(driver):
    """Return why the store is PAUSED, or "" if it is not paused.

    PAUSED means exactly one thing: the storefront is asking for a PREORDER
    ("Preorder for 4:00pm.") while not serving on either channel. The caller
    then checks the clock - inside posted open hours that is a pause, outside
    them it is an ordinary closed shop offering a future slot.

    Deliberately NOT pause evidence:
      - the word "Closed" on its own            -> that is a FLAGGED store
      - Delivery and Pickup both closed, alone  -> that is a FLAGGED store
      - either channel actually serving         -> the store is trading
    Returning "" for those leaves the original Open/Closed/FLAGGED logic to
    decide, exactly as it did before PAUSED existed.
    """
    # -- the preorder ask
    banner = ""
    for xp in _PAUSE_BANNER_XPATHS:
        try:
            for el in driver.find_elements(By.XPATH, xp)[:4]:
                t = _el_text(el, 120)
                if _PREORDER_TEXT_RE.search(t):
                    banner = t
                    break
        except Exception:
            continue
        if banner:
            break

    if not banner:
        return ""                    # no preorder ask -> never PAUSED

    # -- is either channel still serving? If so the store is trading, and a
    #    preorder option alongside it does not make it paused.
    tabs = ""
    for xp in _TABS_XPATHS:
        try:
            for el in driver.find_elements(By.XPATH, xp)[:3]:
                t = _el_text(el, 220)
                low = t.lower()
                if "delivery" in low and "pickup" in low:
                    tabs = t
                    break
        except Exception:
            continue
        if tabs:
            break

    if tabs:
        probe = _ADDRESS_NOISE_RE.sub(" ", tabs.lower())
        if _CHANNEL_LIVE_RE.search(probe):
            return ""                # trading on a real channel -> not paused
        if probe.count("closed") >= 2:
            return f"'{banner}' with Delivery and Pickup both closed"

    return f"'{banner}'"

# ============================================================


def machine_name():
    """── NEW (v5): who ran this. Goes in the subject AND the body. Because the
       send uses your own Gmail, a copy of every run lands in your Sent folder
       automatically - a run log the client can neither see nor delete."""
    try:
        host = platform.node() or socket.gethostname()
    except Exception:
        host = "unknown-host"
    try:
        user = getpass.getuser()
    except Exception:
        user = "unknown-user"
    return f"{host}/{user}"


def send_email_report(file_path, recipient_email, store_count=0):
    sender_email = "t2001e01s01t@gmail.com"
    sender_password = "gdez pcox datp rdnf"   # ── CHANGE: redacted; put your app password back here

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    _machine = machine_name()
    _stamp = datetime.now().strftime('%d-%m-%Y %I:%M %p')
    msg['Subject'] = (f"[Grubhub Scraper] {_machine} | {store_count} stores | "
                      f"to {recipient_email} | {_stamp}")          # ── NEW (v5)
    msg.attach(MIMEText(f"Grubhub store status report\n"
                        f"Machine   : {_machine}\n"
                        f"Recipient : {recipient_email}\n"
                        f"Stores    : {store_count}\n"
                        f"Run time  : {_stamp}\n", "plain"))

    try:
        with open(file_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(file_path)}")
            msg.attach(part)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        log.info("Email sent to %s", recipient_email)
    except Exception as e:
        log.error("Email failed: %s", e, exc_info=True)


def get_location_info(zip_val):
    zip_str = str(zip_val).split('.')[0]
    clean_zip = re.sub(r'[^0-9]', '', zip_str).strip()

    if len(clean_zip) > 5:
        clean_zip = clean_zip[:5]
    elif len(clean_zip) == 4:
        clean_zip = clean_zip.zfill(5)

    # ── CHANGE: return a 3rd value (precise IANA tz) and fix the offset bug ──
    if len(clean_zip) != 5:
        return "Unknown State", "EST", "America/New_York"

    state_names = {
        'AK': 'Alaska', 'AL': 'Alabama', 'AR': 'Arkansas', 'AZ': 'Arizona', 'CA': 'California', 'CO': 'Colorado',
        'CT': 'Connecticut', 'DC': 'District of Columbia', 'DE': 'Delaware', 'FL': 'Florida', 'GA': 'Georgia',
        'HI': 'Hawaii', 'IA': 'Iowa', 'ID': 'Idaho', 'IL': 'Illinois', 'IN': 'Indiana', 'KS': 'Kansas',
        'KY': 'Kentucky', 'LA': 'Louisiana', 'MA': 'Massachusetts', 'MD': 'Maryland', 'ME': 'Maine',
        'MI': 'Michigan', 'MN': 'Minnesota', 'MO': 'Missouri', 'MS': 'Mississippi', 'MT': 'Montana',
        'NC': 'North Carolina', 'ND': 'North Dakota', 'NE': 'Nebraska', 'NH': 'New Hampshire', 'NJ': 'New Jersey',
        'NM': 'New Mexico', 'NV': 'Nevada', 'NY': 'New York', 'OH': 'Ohio', 'OK': 'Oklahoma', 'OR': 'Oregon',
        'PA': 'Pennsylvania', 'RI': 'Rhode Island', 'SC': 'South Carolina', 'SD': 'South Dakota', 'TN': 'Tennessee',
        'TX': 'Texas', 'UT': 'Utah', 'VA': 'Virginia', 'VT': 'Vermont', 'WA': 'Washington', 'WI': 'Wisconsin',
        'WV': 'West Virginia', 'WY': 'Wyoming'
    }

    try:
        data = zipcodes.matching(clean_zip)
        if not data:
            return "Unknown State", "EST", "America/New_York"

        z_info = data[0]
        abbr = z_info.get('state', 'Unknown')
        iana_tz = z_info.get('timezone') or "America/New_York"   # ← the real value is an IANA string
        full_state = state_names.get(abbr, abbr)
        tz_group = IANA_TO_GROUP.get(iana_tz, _group_from_iana(iana_tz))
        return full_state, tz_group, iana_tz

    except Exception:
        return "Unknown State", "EST", "America/New_York"


def app_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path.cwd()


def load_automation_config():
    file_path = app_dir() / "input.xlsx"
    if not file_path.exists():
        log.critical("input.xlsx not found at %s", file_path)
        sys.exit(1)

    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip().str.lower()

    interval_seconds = int(df['interval'][0]) * (60 if 'minute' in str(df['unit'][0]).lower() else 3600)
    total_runs = df['runs'][0] if not pd.isna(df['runs'][0]) else None
    target_email = str(df['email'][0])

    input_records = df[["website_url", "zipcode"]].dropna(subset=["website_url"]).copy()

    # ── CHANGE (v3): dedup input rows by Grubhub store ID before we ever scrape.
    #    Catches exact-duplicate rows AND same store listed under different URL
    #    formatting (slug present/absent, http vs https, trailing slash, etc). ──
    input_records["_dedup_key"] = input_records["website_url"].apply(extract_store_id)

    dupes = input_records[input_records.duplicated(subset="_dedup_key", keep=False)]
    if not dupes.empty:
        log.warning(
            "Found %d duplicate row(s) in input.xlsx by store ID - keeping only the "
            "first occurrence of each. Duplicated URLs: %s",
            len(dupes), ", ".join(sorted(dupes["website_url"].astype(str).unique()))
        )

    before = len(input_records)
    input_records = input_records.drop_duplicates(subset="_dedup_key", keep="first")
    input_records = input_records.drop(columns="_dedup_key")
    after = len(input_records)

    log.info("Loaded %d store(s) from input.xlsx (%d duplicate row(s) removed)", after, before - after)

    return input_records.to_dict('records'), interval_seconds, total_runs, target_email


def run_single_scrape(input_data, recipient_email):
    timeout_seconds = 60
    eastern = pytz.timezone('US/Eastern')
    now_et = datetime.now(eastern)
    date_str = now_et.strftime('%d-%m-%Y')
    time_str = now_et.strftime('%d-%m-%Y_%I-%M-%S_%p')

    output_folder = app_dir() / f"grubhub_extracted_{date_str}"
    output_folder.mkdir(parents=True, exist_ok=True)
    output_excel = str(output_folder / f"Extracted_GRUBHUB_{time_str}.xlsx")

    options = Options()
    options.add_argument("--headless=new")
    try:
        driver = webdriver.Chrome(options=options)
    except Exception:
        log.critical("Failed to start Chrome WebDriver. Is Chrome installed and the "
                     "chromedriver version compatible?", exc_info=True)
        return

    scraped_data = []

    # ── CHANGE (v3): running dedup state for this run ──
    seen_store_ids = set()
    seen_addresses = set()
    seen_name_zip = set()
    skipped_dupes = []

    for record in input_data:
        url = record['website_url'].strip()
        zip_val = record['zipcode']
        state_full, tz_group, tz_iana = get_location_info(zip_val)   # ── CHANGE: 3 values

        log.debug("Scraping: %s", url)                       # ── v5.1: file only
        driver.get(url)

        # defaults so every path can reach a single append at the end
        store_name = ""
        live_closed = False
        paused_evidence = ""      # ── NEW (v5)
        no_menu_evidence = ""     # ── NEW (v5.3)
        rating_no = "0.0"
        review_count = "0"
        tags = []
        address = ""
        mobile = ""
        timings = ""

        try:
            # 1. Store Name
            try:
                store_name = WebDriverWait(driver, timeout_seconds).until(
                    EC.presence_of_element_located((By.XPATH, "//h1"))
                ).text.strip()
            except:
                store_name = ""

            # 2. Live status detection (just a boolean now; we DON'T skip on closed) ── CHANGE
            try:
                closed_el = driver.find_element(
                    By.XPATH, "//*[contains(translate(text(),'CLOSED','closed'),'closed')]")
                if closed_el.is_displayed():
                    live_closed = True
            except NoSuchElementException:
                pass

            # 2b. ── NEW (v5): read the two PAUSED signals off the hero, while it
            #     is still visible. Must happen BEFORE the About panel opens,
            #     because the modal covers the banner and the tab pair.
            paused_evidence = read_pause_evidence(driver)
            if paused_evidence:
                # ── v5.1: file only - the same evidence already appears in
                #    the one-line result printed for this store below.
                log.debug("   pause signal: %s", paused_evidence)

            # 3. Reviews & Ratings ── CHANGE (v5.4): runs for EVERY store.
            #    It used to sit behind "if not live_closed:", so any store that
            #    read as closed - most of them - never had its rating read and
            #    was written out as 0.0 / 0.
            rating_no, review_count = read_rating(driver)
            log.debug("Extracted: %s stars | %s reviews", rating_no, review_count)

            # 4. Open About Tab  ── CHANGE: now runs for CLOSED stores too (we need their hours)
            clicked = False
            possible_locators = [
                (By.XPATH, "//*[normalize-space()='About']"),
                (By.XPATH, "//*[contains(text(),'About')]"),
                (By.XPATH, "//button[contains(.,'About')]"),
                (By.XPATH, "//button[.//text()[contains(.,'About')]]"),
                (By.XPATH, "//a[contains(text(),'About')]"),
                (By.XPATH, "//*[@aria-label='About']"),
                (By.CSS_SELECTOR, "[data-anchor-id='about']")
            ]
            for locator in possible_locators:
                try:
                    element = WebDriverWait(driver, 15).until(EC.element_to_be_clickable(locator))
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
                    time.sleep(1)
                    try:
                        element.click()
                    except:
                        driver.execute_script("arguments[0].click();", element)
                    clicked = True
                    break
                except:
                    continue

            if clicked:
                time.sleep(4)

                # 5. Tags
                try:
                    tags_elements = driver.find_elements(
                        By.XPATH,
                        "//*[contains(@class,'cuisine') or contains(@class,'tag') or contains(@data-testid,'cuisine')]//a | "
                        "//*[contains(@class,'tag')]//span"
                    )
                    for el in tags_elements:
                        clean_val = el.text.strip().replace("$", "").strip(",")
                        if clean_val and clean_val not in tags:
                            tags.append(clean_val)
                except Exception as e:
                    log.warning("Tags extraction issue: %s", e)

                # 6. Address
                address_strategies = [
                    (By.XPATH, "//a[contains(@href, 'google.com/maps') or contains(@href, 'apple.com/maps')]"),
                    (By.XPATH, "//*[@data-testid='restaurant-address']"),
                    (By.XPATH, "//div[contains(@class, 'address')]//span"),
                    (By.TAG_NAME, "address")
                ]
                for locator in address_strategies:
                    try:
                        element = WebDriverWait(driver, 8).until(EC.presence_of_element_located(locator))
                        inner_html = element.get_attribute("innerHTML")
                        clean_address = re.sub(r'<br\s*/?>', '\n', inner_html, flags=re.IGNORECASE)
                        clean_address = re.sub(r'<.*?>', '', clean_address).strip()
                        if clean_address and len(clean_address) > 5:
                            address = clean_address
                            break
                    except:
                        continue
                address = address.replace('\n', ', ').strip()

                # 7. Phone
                try:
                    element = driver.find_element(
                        By.XPATH, "//a[contains(@href,'tel')] | //button[contains(@aria-label,'phone')] | //*[contains(@class,'phone')]")
                    raw_data = element.get_attribute("href") or element.text
                    mobile = "".join(char for char in raw_data if char.isdigit() or char == "+")
                    if len(mobile) < 10:
                        mobile = ""
                except:
                    mobile = ""

                # 8. Timings
                try:
                    timings_elements = driver.find_elements(
                        By.XPATH, "//*[contains(text(),'Delivery') or contains(text(),'Pickup')]")
                    valid_lines = []
                    for el in timings_elements:
                        text = el.text.strip()
                        if re.search(r"\d{1,2}:\d{2}", text) and ":" in text:
                            valid_lines.append(text)
                    timings = "\n".join(valid_lines)
                except:
                    timings = ""
            else:
                # ── NEW (v5.3): no About tab. Confirm with the page wording
                #    before calling it a menu-less store.
                no_menu_evidence = read_menu_unavailable(driver)
                if no_menu_evidence:
                    log.debug("   no-menu wording: %s", no_menu_evidence)
                else:
                    log.warning("About tab not found for %s (continuing; will classify "
                                "on whatever was found)", url)

            while len(tags) < 5:
                tags.append("")

            # ── CHANGE: classify status (Open / Closed / FLAGGED) using the scraped hours ──
            now_utc = datetime.now(pytz.utc)
            final_status, flag_reason = classify_store(
                live_closed, timings, tz_iana, now_utc,
                paused_evidence, no_menu_evidence)                         # ── NEW (v5.3)
            store_local = now_utc.astimezone(pytz.timezone(tz_iana)).strftime('%d-%m %I:%M %p')

            # ============================================================
            # ── CHANGE (v3): duplicate check before appending this row ──
            #   HARD dedup (auto-removed): same Grubhub store ID, or same
            #     physical address - these can't legitimately be two
            #     different real stores.
            #   SOFT check (logged only, never removed): same store name +
            #     same zipcode but different ID/address - chains can
            #     legitimately have two branches in one zip, so this is
            #     just flagged for manual review, not dropped.
            # ============================================================
            store_id_key = extract_store_id(url)
            addr_key = _norm(address)
            zip_key = clean_zip_5(zip_val)
            name_zip_key = f"{_norm(store_name)}|{zip_key}" if store_name and zip_key else ""

            is_hard_dupe = (
                store_id_key in seen_store_ids
                or (addr_key and addr_key in seen_addresses)
            )

            if not is_hard_dupe and name_zip_key and name_zip_key in seen_name_zip:
                log.warning(
                    "POSSIBLE duplicate (same name + zipcode, but different ID/address - "
                    "review manually): %s | %s | zip %s",
                    store_name, url, zip_key
                )

            if is_hard_dupe:
                reason = "same store ID" if store_id_key in seen_store_ids else "same address"
                skipped_dupes.append((store_name or url, url, reason))
                log.warning("Skipping duplicate (%s): %s | %s", reason, store_name, url)
                continue

            seen_store_ids.add(store_id_key)
            if addr_key:
                seen_addresses.add(addr_key)
            if name_zip_key:
                seen_name_zip.add(name_zip_key)
            # ============================================================

            scraped_data.append({
                "Timezone": tz_group,
                "State": state_full,
                "Store Name": store_name,
                "Status": final_status,            # Open / Closed / FLAGGED
                "Store Local Time": store_local,   # ── CHANGE: new column (local time at check)
                "Rating": rating_no,
                "Review Count": review_count,
                "Tag 1": tags[0], "Tag 2": tags[1], "Tag 3": tags[2], "Tag 4": tags[3], "Tag 5": tags[4],
                "All Tags": ", ".join([t for t in tags if t]),
                "Address": address, "Mobile": mobile, "Timings": timings, "Website_link": url
            })
            # ── v5.1: ONE line per store. The bare print() that used to sit
            #    here wrote the same text to the screen a second time.
            log.info("[%d/%d] %-38s %-8s %s", len(scraped_data), len(input_data),
                     (store_name or url)[:38], final_status, flag_reason)
            if not timings:
                log.debug("   (no posted hours captured for %s)", store_name or url)

        except Exception as e:
            log.error("Error on URL %s: %s", url, e, exc_info=True)

    driver.quit()

    # ── CHANGE (v3): run-level duplicate summary ──
    if skipped_dupes:
        log.warning("Removed %d duplicate row(s) this run:", len(skipped_dupes))
        for name, dupe_url, reason in skipped_dupes:
            log.warning("   %s | %s | matched on: %s", name, dupe_url, reason)

    if not scraped_data:
        log.warning("No data available to format.")
        return

    df = pd.DataFrame(scraped_data)

    # ── CHANGE: per-run summary in the log (counts + the flagged store names) ──
    counts = {}
    for row in scraped_data:
        counts[row["Status"]] = counts.get(row["Status"], 0) + 1
    log.info("Run summary - %d stores: %s", len(scraped_data),
             ", ".join(f"{k}={v}" for k, v in sorted(counts.items())))
    for _lbl in (PAUSE_LABEL, FLAG_LABEL, NO_MENU_LABEL):   # ── v5.3: NO MENU too
        _names = [r["Store Name"] or "(unnamed)" for r in scraped_data
                  if str(r["Status"]) == _lbl]
        if _names:
            log.warning("%s (%d): %s", _lbl, len(_names), "; ".join(_names))

    # ── CHANGE: tz grouping now generic over whatever zones are present (EST/CT/MT/PT/AKT/HT) ──
    present_tz = list(df['Timezone'].unique())
    tz_status_summary = {}
    for tz_code in present_tz:
        tz_rows = df[df['Timezone'] == tz_code]
        all_closed = tz_rows['Status'].apply(lambda x: str(x).strip().lower() == "closed").all()
        tz_status_summary[tz_code] = 1 if all_closed else 0

    sorted_tz_order = sorted(present_tz, key=lambda tz: (tz_status_summary[tz], GROUP_ORDER.get(tz, 9)))
    tz_dynamic_rank = {tz: i for i, tz in enumerate(sorted_tz_order)}

    # ── CHANGE: FLAGGED rows float to the top of each zone, then Closed, then Open ──
    def _status_rank(x):
        s = str(x).strip().lower()
        if s == PAUSE_LABEL.lower():               # ── NEW (v5): paused first
            return 0
        if s == FLAG_LABEL.lower():
            return 1
        if s == NO_MENU_LABEL.lower():             # ── NEW (v5.3)
            return 2
        if s == "closed":
            return 3
        return 4

    df['tz_rank'] = df['Timezone'].map(tz_dynamic_rank).fillna(9)
    df['status_rank'] = df['Status'].apply(_status_rank)
    df = df.sort_values(by=["tz_rank", "status_rank", "State", "Store Name"])
    df.drop(columns=['tz_rank', 'status_rank'], inplace=True)
    df.insert(0, "Serial No", range(1, len(df) + 1))
    df.to_excel(output_excel, index=False)

    wb = load_workbook(output_excel)
    ws = wb.active
    num_columns = ws.max_column

    zone_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    closed_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    flag_fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")   # ── CHANGE: gold for FLAGGED
    pause_fill = PatternFill(start_color="8E44AD", end_color="8E44AD", fill_type="solid")  # ── NEW (v5): violet for PAUSED
    header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    low_rating_fill = PatternFill(start_color="FF8C00", end_color="FF8C00", fill_type="solid")
    open_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")   # ── v5.2: legend swatch only
    no_menu_fill = PatternFill(start_color="2E86C1", end_color="2E86C1", fill_type="solid")  # ── v5.3: blue for NO MENU

    white_bold = Font(color="FFFFFF", bold=True)
    white_reg = Font(color="FFFFFF")
    black_bold = Font(bold=True)
    flag_font = Font(bold=True, color="000000")   # ── CHANGE
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    col_map = {str(ws.cell(row=1, column=c).value).strip(): c for c in range(1, num_columns + 1)}
    status_idx = col_map.get("Status")
    tz_idx = col_map.get("Timezone")
    rating_idx = col_map.get("Rating")

    for c in range(1, num_columns + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = black_bold
        cell.border = thin_border
        cell.alignment = center_align

    current_row = 2
    for tz_code in sorted_tz_order:
        rows_for_this_tz = []
        for r in range(current_row, ws.max_row + 1):
            if ws.cell(row=r, column=tz_idx).value == tz_code:
                rows_for_this_tz.append(r)

        if rows_for_this_tz:
            ws.insert_rows(current_row)
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=num_columns)

            tz_now = datetime.now(pytz.timezone(GROUP_TO_IANA.get(tz_code, 'US/Eastern'))).strftime('%d-%m-%Y %I:%M %p')
            header_cell = ws.cell(row=current_row, column=1)

            closed_flag_string = " [ALL STORES CLOSED]" if tz_status_summary[tz_code] == 1 else ""
            header_cell.value = f"--- {tz_code} TIME ZONE STORES ---{closed_flag_string} (Reported at: {tz_now} {tz_code})"

            header_cell.fill = zone_fill
            header_cell.font = white_bold
            header_cell.alignment = center_align
            ws.row_dimensions[current_row].height = 25

            for c in range(1, num_columns + 1):
                ws.cell(row=current_row, column=c).border = thin_border

            for r in range(current_row + 1, current_row + 1 + len(rows_for_this_tz)):
                status_val = str(ws.cell(row=r, column=status_idx).value).strip().lower()
                is_flag = (status_val == FLAG_LABEL.lower())   # ── CHANGE
                is_paused = (status_val == PAUSE_LABEL.lower())  # ── NEW (v5)
                is_no_menu = (status_val == NO_MENU_LABEL.lower())  # ── NEW (v5.3)
                max_lines = 1

                for c in range(1, num_columns + 1):
                    cell = ws.cell(row=r, column=c)
                    cell.alignment = center_align
                    cell.border = thin_border

                    if status_val == "closed":
                        cell.fill = closed_fill
                        cell.font = white_reg
                    elif is_paused:                            # ── NEW (v5): violet row
                        cell.fill = pause_fill
                        cell.font = white_bold
                    elif is_no_menu:                           # ── NEW (v5.3): blue row
                        cell.fill = no_menu_fill
                        cell.font = white_reg
                    elif is_flag:                              # ── CHANGE: gold row for mismatches
                        cell.fill = flag_fill
                        cell.font = flag_font

                    if cell.value:
                        line_count = str(cell.value).count('\n') + 1
                        if line_count > max_lines:
                            max_lines = line_count

                if rating_idx and status_val not in ("closed", FLAG_LABEL.lower(),
                                                     PAUSE_LABEL.lower(),
                                                     NO_MENU_LABEL.lower()):
                    rating_cell = ws.cell(row=r, column=rating_idx)
                    try:
                        rating_val = float(str(rating_cell.value).strip())
                        if 0.1 <= rating_val < 3.5:
                            rating_cell.fill = low_rating_fill
                            rating_cell.font = Font(bold=True, color="FFFFFF")
                    except (ValueError, TypeError):
                        pass

                ws.row_dimensions[r].height = max_lines * 15

            current_row += len(rows_for_this_tz) + 1
            if current_row <= ws.max_row:
                ws.insert_rows(current_row)
                ws.row_dimensions[current_row].height = 10
                current_row += 1

    ws.freeze_panes = "A2"
    for col in ws.columns:
        max_length = 0
        column_letter = col[0].column_letter
        if column_letter in ['A', 'B']:
            ws.column_dimensions[column_letter].width = 10
            continue
        for cell in col:
            if cell.value and not isinstance(cell.value, datetime):
                lines = str(cell.value).split('\n')
                current_max = max(len(line) for line in lines)
                if current_max > max_length:
                    max_length = current_max
        adjusted_width = min(max_length + 4, 60)
        ws.column_dimensions[column_letter].width = adjusted_width

    # ── CHANGE: Add a Legend at the bottom of the worksheet ──
    legend_row = ws.max_row + 3
    ws.cell(row=legend_row, column=1, value="LEGEND / COLOR KEY:").font = black_bold

    # ── v5.2: every colour the sheet can show, each with what it means and
    #    what the owner should do about it. Open has no fill, so it is listed
    #    with a plain white swatch and an outline rather than left unexplained.
    legend_items = [
        (header_fill,     "Column Header row - the titles at the top of the sheet"),
        (zone_fill,       "Time Zone Section Header - the stores below this banner "
                          "are all in this time zone"),
        (pause_fill,      f"{PAUSE_LABEL} - store is INSIDE its posted open hours but Grubhub "
                          f"is only taking preorders. Somebody switched it off: chase it."),
        (flag_fill,       f"{FLAG_LABEL} - storefront reads CLOSED while its posted hours say it "
                          f"should be OPEN. No preorder shown, so the cause is unknown: check it."),
        (no_menu_fill,    f"{NO_MENU_LABEL} - the storefront has no menu at all (no About tab "
                          f"and the page says the menu is not available). Nothing can be "
                          f"ordered from it: fix the listing."),
        (closed_fill,     "Closed - shut, and its posted hours agree it should be shut. "
                          "Expected, no action needed."),
        (open_fill,       "Open - trading, and its posted hours agree it should be. "
                          "Expected, no action needed (rows are left unfilled)."),
        (low_rating_fill, "Low Rating - the Rating cell only, when the store scores "
                          "below 3.5 stars. The row's own status is unaffected."),
    ]
    for _i, (_fill, _text) in enumerate(legend_items, start=1):
        _sw = ws.cell(row=legend_row + _i, column=1)
        _sw.fill = _fill
        _sw.border = thin_border
        _lbl = ws.cell(row=legend_row + _i, column=2, value=_text)
        _lbl.font = black_bold
        _lbl.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)

    wb.save(output_excel)
    log.info("File saved and formatted: %s", output_excel)
    send_email_report(output_excel, recipient_email, store_count=len(scraped_data))


def main():
    setup_logging()
    recs, interval, runs, email = load_automation_config()
    count = 0
    while True:
        count += 1
        log.info("--- RUN %d ---", count)
        run_single_scrape(recs, email)
        if runs and count >= int(runs):
            break
        time.sleep(interval)


if __name__ == "__main__":
    setup_logging()
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        log.critical("Fatal error - scraper stopped unexpectedly", exc_info=True)
        raise