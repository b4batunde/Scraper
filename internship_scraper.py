"""
2026 Tech Internship Scraper
============================
Sources:
  1. github.com/SimplifyJobs/Summer2026-Internships          (summer)
  2. github.com/SimplifyJobs/Summer2026-Internships Off-Season (fall/winter/spring)
  3. github.com/speedyapply/2026-SWE-College-Jobs
  4. github.com/vanshb03/Summer2027-Internships
  5. workatastartup.com/companies (Engineering internships, sorted newest first)

Output: 2026_Tech_Internships_USA.xlsx

Install:
  pip install requests beautifulsoup4 openpyxl lxml playwright
  playwright install chromium

Run:
  python3 internship_scraper.py
"""

import re
import math
import json
import datetime
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────────────────────────────────────
# Config
# ─────────────────────────────────────────────────────────────────────────────

TODAY       = datetime.date.today()
OUTPUT_FILE = "2026_Tech_Internships_USA.xlsx"

GITHUB_SOURCES = {
    "Simplify":    "https://raw.githubusercontent.com/SimplifyJobs/Summer2026-Internships/dev/README.md",
    "Simplify-OS": "https://raw.githubusercontent.com/SimplifyJobs/Summer2026-Internships/dev/README-Off-Season.md",
    "SpeedyApply": "https://raw.githubusercontent.com/speedyapply/2026-SWE-College-Jobs/main/README.md",
    "Speedy2027":  "https://raw.githubusercontent.com/speedyapply/2027-SWE-College-Jobs/main/README.md",
    "Vansh":       "https://raw.githubusercontent.com/vanshb03/Summer2027-Internships/main/README.md",
    "Sndsh":       "https://raw.githubusercontent.com/sndsh404/summer-2027-internships/main/README.md",
}


# ─────────────────────────────────────────────────────────────────────────────
# Filters
# ─────────────────────────────────────────────────────────────────────────────

TECH_KW = re.compile(
    r"\b(software|engineer|developer|dev\b|ML|AI\b|machine.?learning|data|cloud|"
    r"backend|front.?end|full.?stack|SWE|SDE|tech|platform|infrastructure|"
    r"computer|systems|network|security|cyber|embedded|firmware|robotics|"
    r"autonomous|NLP|deep.?learning|algorithm|API|mobile|android|iOS|web|"
    r"devops|site.?reliability|quant|simulation|applied.?science|"
    r"research.?engineer|compiler|GPU|hardware|FPGA|agent|LLM|generative|"
    r"analytics|intelligence|inference|middleware|supercomputing|forward.?deployed)\b",
    re.IGNORECASE,
)
USA_BLACKLIST = re.compile(
    r"\b(Canada|CAN\b|UK\b|Germany|France|Singapore|India|China|"
    r"Australia|Netherlands|Sweden|Ireland|Israel|Japan|Brazil|"
    r"Switzerland|Poland|Ontario|British Columbia|Quebec|Toronto|Vancouver|"
    r"London|Berlin|Paris|Dublin|Amsterdam|Tel Aviv|Sydney|Melbourne|"
    r"Newmarket|Ottawa|Cayman)\b",
    re.IGNORECASE,
)
USA_KW = re.compile(
    r"\b(US\b|USA\b|United States|America|Remote|"
    r"AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|"
    r"MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|"
    r"VT|VA|WA|WV|WI|WY|DC|"
    r"New York|California|Texas|Washington|Seattle|San Francisco|San Jose|"
    r"Boston|Chicago|Austin|Atlanta|Denver|Pittsburgh|Philadelphia|"
    r"Santa Clara|Menlo Park|Palo Alto|Mountain View|Los Angeles|"
    r"Sunnyvale|Cupertino|Redmond|Bellevue|Cambridge|Raleigh|Portland|"
    r"Dallas|Herndon|Milpitas|Albuquerque|San Diego|Grand Rapids|Marietta|"
    r"San Mateo|Scottsdale|Chantilly|Oakland|Salt Lake City|Bensenville|"
    r"Hammond|Irvine|Foster City|Houston|Miami|Tampa|Orlando|Charlotte|"
    r"Nashville|Memphis|Indianapolis|Columbus|Cincinnati|Milwaukee|"
    r"Minneapolis|Kansas City|Baltimore|Louisville|Bala Cynwyd|Kent|Lehi|"
    r"Jupiter|Redwood City|Coral Gables|Alpharetta|Springfield|Burlington|"
    r"Fort Lauderdale|Jacksonville|West Palm Beach|New Orleans)\b",
    re.IGNORECASE,
)
MASTERS_RE   = re.compile(r"\b(masters?|master's|phd|ph\.d|graduate|grad\b|mba)\b", re.IGNORECASE)
UNDERGRAD_OK = re.compile(r"\b(undergrad|undergraduate|BS\b|B\.S\.|junior|senior|sophomore|freshman|bachelor)\b", re.IGNORECASE)

NO_SPONSOR = "🛂"
US_CITIZEN = "🇺🇸"
GRAD_EMOJI = "🎓"
CLOSED     = "🔒"


def is_usa(loc: str) -> bool:
    if not loc: return False
    if USA_BLACKLIST.search(loc): return False
    return bool(USA_KW.search(loc))

def is_tech(title: str) -> bool:
    return bool(TECH_KW.search(title))

def is_grad_only(title: str) -> bool:
    if GRAD_EMOJI in title and not UNDERGRAD_OK.search(title): return True
    if MASTERS_RE.search(title) and not UNDERGRAD_OK.search(title): return True
    return False


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def strip_html(text: str) -> str:
    return BeautifulSoup(text, "lxml").get_text(" ", strip=True)

def extract_href(html: str) -> str:
    m = re.search(r'href=["\']([^"\']+)["\']', html)
    return m.group(1) if m else ""

def fetch(url: str) -> str:
    resp = requests.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    return resp.text

def age_str_to_days(s: str) -> int:
    s = s.strip()
    if not s or s in ("Age", "N/A", "—", ""): return 999
    m = re.match(r"^(\d+)d$", s)
    if m: return int(m.group(1))
    m = re.match(r"^(\d+)mo$", s)
    if m: return int(m.group(1)) * 30
    # pass year explicitly to avoid deprecation warning
    for fmt in ("%b %d %Y", "%B %d %Y"):
        try:
            parsed = datetime.datetime.strptime(f"{s} {TODAY.year}", fmt).date()
            if parsed > TODAY:
                parsed = parsed.replace(year=TODAY.year - 1)
            return (TODAY - parsed).days
        except ValueError:
            pass
    m = re.match(r"^(\d+)$", s)
    if m: return int(m.group(1))
    return 999

def unix_to_days(ts) -> int:
    if not ts: return 999
    try:
        return max(0, (TODAY - datetime.date.fromtimestamp(int(ts))).days)
    except Exception:
        return 999


# ─────────────────────────────────────────────────────────────────────────────
# Scoring
# ─────────────────────────────────────────────────────────────────────────────

BIG_TECH  = re.compile(r"\b(Google|Meta|Microsoft|Apple|Amazon|Netflix|Salesforce|Adobe|Intel|"
    r"Qualcomm|AMD|Oracle|IBM|Cisco|VMware|Uber|Lyft|Twitter|LinkedIn|"
    r"TikTok|ByteDance|Tencent|Baidu|Samsung)\b", re.IGNORECASE)
STRONG_CO = re.compile(r"\b(NVIDIA|Anthropic|OpenAI|DeepMind|Waymo|Tesla|Stripe|Airbnb|"
    r"Databricks|Snowflake|Figma|Ramp|Robinhood|Palantir|SpaceX|Neuralink|"
    r"CrowdStrike|Cloudflare|Datadog|MongoDB|Coinbase|Anduril|Groq|Cerebras|"
    r"Etched|Zoox|Skydio|Aquatic|Cubist|Susquehanna|D\.E\. Shaw|Jane Street|"
    r"Two Sigma|Citadel|Renaissance|Point72|Jump Trading|Stoke Space)\b", re.IGNORECASE)
AI_ML     = re.compile(r"\b(ML|machine.?learning|AI\b|deep.?learning|NLP|LLM|generative|agent|"
    r"reinforcement|computer.?vision|applied.?science|foundation.?model|"
    r"inference|post.?training|research.?scientist|intelligence)\b", re.IGNORECASE)
SWE       = re.compile(r"\b(software|SWE|SDE|developer|full.?stack|backend|frontend|platform|"
    r"infrastructure|cloud|devops|firmware|embedded|robotics|systems)\b", re.IGNORECASE)
TOP_HUB   = re.compile(r"\b(San Francisco|Palo Alto|Mountain View|Menlo Park|Santa Clara|"
    r"Sunnyvale|Cupertino|New York|NYC|Seattle|Boston|Cambridge|San Jose|"
    r"Foster City|Redwood City)\b", re.IGNORECASE)
MID_HUB   = re.compile(r"\b(Austin|Chicago|Los Angeles|Denver|Atlanta|Raleigh|Portland|"
    r"San Diego|Pittsburgh|Miami|Houston|Minneapolis)\b", re.IGNORECASE)


def recency_points(days: int) -> int:
    """
    Exponential decay: score = 50 × e^(-0.085 × days)
    day 0→50 | 1→46 | 2→42 | 3→39 | 5→33 | 7→28 | 14→15 | 21→8 | 30→4 | 45+→1
    """
    if days == 0: return 50
    return max(1, round(50 * math.exp(-0.085 * days)))


def score_entry(company, title, location, flags, days):
    reasons = []
    rp = recency_points(days)
    reasons.append(f"Posted {days}d ago (+{rp})")

    if STRONG_CO.search(company):   cp = 20; reasons.append("Strong/AI company (+20)")
    elif BIG_TECH.search(company):  cp = 10; reasons.append("Big tech (+10)")
    else:                           cp = 0

    if AI_ML.search(title):   rtp = 25; reasons.append("AI/ML/Automation role (+25)")
    elif SWE.search(title):   rtp = 12; reasons.append("SWE role (+12)")
    else:                     rtp = 0

    if TOP_HUB.search(location):          lp = 8; reasons.append("Top tech hub (+8)")
    elif "remote" in location.lower():    lp = 7; reasons.append("Remote (+7)")
    elif MID_HUB.search(location):        lp = 4; reasons.append("Tech city (+4)")
    else:                                 lp = 0

    pen = 0
    if NO_SPONSOR in flags: pen -= 20; reasons.append("No sponsorship 🛂 (−20)")
    if US_CITIZEN in flags: pen -= 30; reasons.append("US citizen only 🇺🇸 (−30)")

    return max(0, min(100, rp + cp + rtp + lp + pen)), " | ".join(reasons)


# ─────────────────────────────────────────────────────────────────────────────
# Parsers — GitHub
# ─────────────────────────────────────────────────────────────────────────────

def parse_simplify(content, source_label="Simplify"):
    """
    Parses Simplify HTML tables. Handles both formats:
      - Summer (5 cols): company | title | location | apply | age
      - Off-Season (6 cols): company | title | location | term | apply | age
    """
    rows = re.findall(r"<tr>(.*?)</tr>", content, re.DOTALL)
    out = []; last_co = ""
    for r in rows:
        cells = re.findall(r"<td>(.*?)</td>", r, re.DOTALL)
        if len(cells) < 4: continue

        raw_co    = cells[0]
        raw_title = cells[1]
        raw_loc   = cells[2]

        # 6-col off-season: [company, title, location, term, apply, age]
        # 5-col summer:     [company, title, location, apply, age]
        if len(cells) >= 6:
            raw_apply = cells[4]
            age_str   = cells[5].strip()
        else:
            raw_apply = cells[3]
            age_str   = cells[4].strip() if len(cells) > 4 else ""

        # Skip closed listings
        if CLOSED in raw_apply: continue

        if "↳" in raw_co:
            company = last_co
        else:
            company = re.sub(r"[🔥🛂🇺🇸]", "", strip_html(raw_co)).strip()
            last_co = company

        title    = strip_html(raw_title).strip()
        location = strip_html(raw_loc).replace("\n", ", ").strip()
        apply    = extract_href(raw_apply)
        flags    = raw_co + raw_title

        if not is_usa(location): continue
        if not is_tech(title):   continue
        if is_grad_only(title):  continue

        days = age_str_to_days(age_str)
        sc, why = score_entry(company, title, location, flags, days)
        out.append(dict(company=company, title=title, location=location,
            apply_url=apply, source=source_label, age_str=age_str, days=days,
            salary="", score=sc, score_reason=why,
            no_sponsor=NO_SPONSOR in flags, us_citizen=US_CITIZEN in flags))
    return out


def parse_speedyapply(content, source_label="SpeedyApply"):
    """Handles both speedyapply/2026-SWE-College-Jobs and speedyapply/2027-SWE-College-Jobs."""
    out = []
    for line in content.split("\n"):
        line = line.strip()
        if not line.startswith("|") or line.startswith("|---"): continue
        cols = [c.strip() for c in line.strip("|").split("|")]
        if len(cols) < 3: continue
        raw_co, raw_title, raw_loc = cols[0], cols[1], cols[2]
        if raw_co in ("Company", "**Company**", ""): continue
        salary  = strip_html(cols[3]).strip() if len(cols) > 3 else ""
        apply   = extract_href(cols[4]) if len(cols) > 4 else ""
        age_str = cols[5].strip() if len(cols) > 5 else ""
        if age_str in ("Age", "Posting"): continue
        m = re.search(r"<strong>(.*?)</strong>", raw_co, re.DOTALL)
        company  = strip_html(m.group(1)) if m else strip_html(raw_co)
        title    = strip_html(raw_title).strip()
        location = strip_html(raw_loc).strip()
        # Speedy2027 has "+N" location suffix like "Irvine, CA +1" — strip it
        location = re.sub(r"\s*\+\d+$", "", location).strip()
        if not is_usa(location): continue
        if not is_tech(title):   continue
        if is_grad_only(title):  continue
        if MASTERS_RE.search(title) and not UNDERGRAD_OK.search(title): continue
        days = age_str_to_days(age_str)
        sc, why = score_entry(company, title, location, "", days)
        out.append(dict(company=company, title=title, location=location,
            apply_url=apply, source=source_label, age_str=age_str, days=days,
            salary=salary, score=sc, score_reason=why,
            no_sponsor=False, us_citizen=False))
    return out


def parse_vansh(content):
    out = []; last_co = ""
    for line in content.split("\n"):
        line = line.strip()
        if not line.startswith("|") or line.startswith("|---"): continue
        cols = [c.strip() for c in line.strip("|").split("|")]
        if len(cols) < 4: continue
        raw_co, raw_title, raw_loc, raw_apply = cols[0], cols[1], cols[2], cols[3]
        date_str = cols[4].strip() if len(cols) > 4 else ""
        if raw_co in ("Company", ""): continue
        if CLOSED in raw_apply: continue
        if "↳" in raw_co:
            company = last_co
        else:
            company = re.sub(r"[🛂🇺🇸🔒🔥]", "", strip_html(raw_co)).strip()
            last_co = company
        title     = re.sub(r"[🔒]", "", strip_html(raw_title)).strip()
        loc_clean = re.sub(r"<details.*?</details>", "", raw_loc, flags=re.DOTALL)
        location  = strip_html(loc_clean).strip()
        apply     = extract_href(raw_apply)
        flags     = raw_co + raw_title
        if not is_usa(location): continue
        if not is_tech(title):   continue
        if is_grad_only(title):  continue
        if MASTERS_RE.search(title) and not UNDERGRAD_OK.search(title): continue
        days = age_str_to_days(date_str)
        sc, why = score_entry(company, title, location, flags, days)
        out.append(dict(company=company, title=title, location=location,
            apply_url=apply, source="Vansh/Ouckah", age_str=date_str, days=days,
            salary="", score=sc, score_reason=why,
            no_sponsor=NO_SPONSOR in flags, us_citizen=US_CITIZEN in flags))
    return out




def parse_sndsh(content):
    """
    Pipe table: Company | Role | Location | Apply | Added
    Added is ISO date (2026-07-09) or dash.
    """
    out = []
    for line in content.split("\n"):
        line = line.strip()
        if not line.startswith("|") or line.startswith("|---"): continue
        cols = [c.strip() for c in line.strip("|").split("|")]
        if len(cols) < 4: continue
        raw_co, raw_title, raw_loc, raw_apply = cols[0], cols[1], cols[2], cols[3]
        date_str = cols[4].strip() if len(cols) > 4 else ""
        if raw_co in ("Company", "---", ""): continue

        company  = re.sub(r"[\U0001F6C2\U0001F1FA\U0001F512\U0001F525]", "", raw_co).strip()
        title    = re.sub(r"[\U0001F6C2\U0001F1FA\U0001F512\U0001F525]", "", raw_title).strip()
        location = raw_loc.strip()
        # Handle [apply](url) markdown links
        apply_m  = re.search(r"\[.*?\]\((.+?)\)", raw_apply)
        apply    = apply_m.group(1) if apply_m else extract_href(raw_apply)
        flags    = raw_co + raw_title

        if not company or not title:  continue
        if not is_usa(location):      continue
        if not is_tech(title):        continue
        if is_grad_only(title):       continue
        if MASTERS_RE.search(title) and not UNDERGRAD_OK.search(title): continue

        # ISO date: "2026-07-09"
        days = 999
        if date_str and date_str not in ("—", "-", ""):
            try:
                posted = datetime.date.fromisoformat(date_str)
                days   = max(0, (TODAY - posted).days)
            except ValueError:
                days = age_str_to_days(date_str)

        age_disp = f"{days}d" if days < 999 else "—"
        sc, why  = score_entry(company, title, location, flags, days)
        out.append(dict(company=company, title=title, location=location,
            apply_url=apply, source="Sndsh", age_str=age_disp, days=days,
            salary="", score=sc, score_reason=why,
            no_sponsor=NO_SPONSOR in flags, us_citizen=US_CITIZEN in flags))
    return out



# ─────────────────────────────────────────────────────────────────────────────
# Parser — YC via Playwright (JS-rendered page)
# ─────────────────────────────────────────────────────────────────────────────

def _process_co_data(co_data: dict, out: list):
    """Parse one company blob from /companies/fetch and append matching jobs to out."""
    company = (co_data.get("name") or "").strip()
    jobs    = co_data.get("jobs") or []
    for job in jobs:
        # job type filter — keep intern / internship
        jtype = (job.get("type") or job.get("jobType") or "").lower()
        if jtype and "intern" not in jtype:
            continue

        title = (job.get("title") or "").strip()
        if not title or not company:
            continue

        # Location — locations[] can be list of dicts or strings
        locs     = job.get("locations") or []
        loc_strs = []
        for l in locs:
            if isinstance(l, dict):
                loc_strs.append(l.get("text") or l.get("name") or l.get("city") or "")
            else:
                loc_strs.append(str(l))
        location = ", ".join(s for s in loc_strs if s).strip()
        if not location and job.get("remote"):
            location = "Remote"

        # Apply URL
        job_id = job.get("id") or ""
        apply  = f"https://www.workatastartup.com/jobs/{job_id}" if job_id else ""

        # Date — try several field names; value may be unix int, ISO string, or Rails datetime
        created_raw = (job.get("created_at") or job.get("createdAt")
                       or job.get("posted_at") or job.get("postedAt"))
        if isinstance(created_raw, (int, float)):
            days = unix_to_days(int(created_raw))
        elif isinstance(created_raw, str) and created_raw:
            # ISO: "2026-05-27T18:00:00.000Z"
            try:
                dt   = datetime.datetime.fromisoformat(created_raw.replace("Z", "+00:00"))
                days = max(0, (TODAY - dt.date()).days)
            except Exception:
                days = 999
        else:
            days = 999

        age_str = f"{days}d" if days < 999 else "—"

        # Salary
        sal_raw = job.get("salary") or job.get("comp") or ""
        salary  = str(sal_raw).strip() if sal_raw else ""

        # Visa — several possible field names
        visa_ok = (job.get("visaSponsorship") or job.get("willSponsor")
                   or job.get("usVisaNotRequired") or job.get("visa_sponsorship"))
        # "US Citizen/Visa Only" means they WON'T sponsor
        us_citizen_only = "citizen" in str(job.get("visa") or "").lower()

        if not is_usa(location):  return
        if not is_tech(title):    return
        if is_grad_only(title):   return

        sc, why = score_entry(company, title, location,
                              US_CITIZEN if us_citizen_only else "", days)
        out.append(dict(
            company=company, title=title, location=location,
            apply_url=apply, source="YC", age_str=age_str, days=days,
            salary=salary, score=sc, score_reason=why,
            no_sponsor=not bool(visa_ok), us_citizen=us_citizen_only,
        ))


def parse_yc_playwright() -> list:
    """
    Uses Playwright to:
      1. Open ycombinator.com/internships and pause for you to log in (30 sec)
      2. Navigate to workatastartup.com/companies with Engineering+Intern filters
      3. Intercept the /companies/fetch POST response (fired automatically by the page)
      4. Parse all company+job data from the JSON

    Install once:
        pip install playwright
        playwright install chromium
    """
    from playwright.sync_api import sync_playwright

    YC_LOGIN_URL = "https://www.ycombinator.com/internships"
    YC_URL = (
        "https://www.workatastartup.com/companies"
        "?demographic=any&hasEquity=any&hasSalary=any&industry=any"
        "&interviewProcess=any&jobType=intern&layout=list-compact"
        "&role=eng&sortBy=created_desc&tab=any&usVisaNotRequired=any"
    )

    print("    Launching browser (NOT headless — you need to log in)...")
    print()
    print("    " + "─" * 58)
    print("    A browser window will open to ycombinator.com/internships.")
    print("    Log in with your YC account, then come back here.")
    print("    The script will continue automatically after 45 seconds.")
    print("    " + "─" * 58)
    print()

    out      = []
    api_data = []

    def handle_response(response):
        if "companies/fetch" in response.url and response.status == 200:
            try:
                body = response.json()
                if isinstance(body, list):
                    api_data.extend(body)
                elif isinstance(body, dict):
                    # response shape: {"companies": [...], "recommendedJobIds": [...]}
                    cos = body.get("companies") or []
                    if cos:
                        api_data.extend(cos)
            except Exception:
                pass

    with sync_playwright() as p:
        # headless=False so the browser window is visible for login
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            user_agent=("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) "
                        "Chrome/124.0.0.0 Safari/537.36"),
        )
        page = context.new_page()

        # Step 1: open login page, wait for user to sign in
        page.goto(YC_LOGIN_URL, wait_until="domcontentloaded", timeout=30000)
        print("    Waiting 45s for you to log in...")
        page.wait_for_timeout(45000)

        # Step 2: navigate to the filtered internships page — this fires /companies/fetch
        print("    Navigating to internships listing...")
        page.on("response", handle_response)
        page.goto(YC_URL, wait_until="networkidle", timeout=45000)
        page.wait_for_timeout(3000)

        # Step 3: scroll to load all results (page uses infinite scroll)
        print("    Scrolling to load all results...")
        prev_count = 0
        for _ in range(15):
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            page.wait_for_timeout(1500)
            if len(api_data) == prev_count:
                break
            prev_count = len(api_data)

        browser.close()

    print(f"    Got {len(api_data)} companies from /companies/fetch")

    if not api_data:
        print("    ⚠️  No data captured — were you able to log in?")
        return []

    for co_data in api_data:
        _process_co_data(co_data, out)

    return out


def parse_yc() -> list:
    """Entry point for YC scraping — tries Playwright, skips gracefully if not installed."""
    try:
        import playwright  # noqa: F401
    except ImportError:
        print("    ⚠️  Playwright not installed.")
        print("    Run: pip install playwright && playwright install chromium")
        return []
    try:
        return parse_yc_playwright()
    except Exception as e:
        print(f"    ⚠️  Playwright failed: {e}")
        return []


# ─────────────────────────────────────────────────────────────────────────────
# Deduplication
# ─────────────────────────────────────────────────────────────────────────────

def normalize(s: str) -> str:
    s = s.lower().strip()
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def dedup_key(job: dict) -> tuple:
    """Company + first 6 title words — catches minor wording differences across sources."""
    co    = normalize(job["company"])
    words = normalize(job["title"]).split()[:6]
    return (co, " ".join(words))


# ─────────────────────────────────────────────────────────────────────────────
# Excel builder
# ─────────────────────────────────────────────────────────────────────────────

NAVY   = "0D1B2A"; TEAL   = "1B998B"; WHITE  = "FFFFFF"; LGRAY  = "F7F8FA"
LINK   = "1A56DB"
GRN_BG = "D1FAE5"; GRN_FG = "065F46"; YLW_BG = "FEF9C3"; YLW_FG = "92400E"
RED_BG = "FFE4E6"; RED_FG = "9F1239"; ORG_BG = "FFF7ED"; ORG_FG = "9A3412"
AMB_BG = "FFFBEB"; AMB_FG = "B45309"

SOURCE_BG = {"Simplify":"EDE9FE","Simplify-OS":"F3E8FF","SpeedyApply":"DBEAFE","Speedy2027":"BFDBFE",
             "Vansh/Ouckah":"DCFCE7","Sndsh":"D1FAE5","YC":"FEF9C3"}
SOURCE_FG = {"Simplify":"4C1D95","Simplify-OS":"6B21A8","SpeedyApply":"1E3A8A","Speedy2027":"1E40AF",
             "Vansh/Ouckah":"14532D","Sndsh":"065F46","YC":"713F12"}

HEADERS = ["Score","Company","Role / Title","Location","Salary","Source","Posted","Visa","Apply","Score Breakdown"]
WIDTHS  = [8, 22, 46, 26, 9, 13, 8, 13, 10, 68]

def _border():
    s = Side(style="thin", color="D5D8DC")
    return Border(left=s, right=s, top=s, bottom=s)

def _sfill(sc):
    if sc >= 70: return PatternFill("solid", fgColor=GRN_BG)
    if sc >= 50: return PatternFill("solid", fgColor=AMB_BG)
    if sc >= 30: return PatternFill("solid", fgColor=YLW_BG)
    return PatternFill("solid", fgColor=RED_BG)

def _sfont(sc):
    if sc >= 70: return Font(name="Arial", bold=True, size=11, color=GRN_FG)
    if sc >= 50: return Font(name="Arial", bold=True, size=11, color=AMB_FG)
    if sc >= 30: return Font(name="Arial", bold=True, size=11, color=YLW_FG)
    return Font(name="Arial", bold=True, size=11, color=RED_FG)


def build_excel(jobs: list, output_path: str):
    wb = Workbook(); ws = wb.active; ws.title = "Internships"
    bdr = _border()

    ws.merge_cells("A1:J1")
    c = ws["A1"]
    c.value = "⚡  2026 Tech Internships  |  USA Only  |  Undergrad Eligible  |  Recency-First Scoring"
    c.font  = Font(name="Arial", bold=True, size=12, color=WHITE)
    c.fill  = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:J2")
    sub = ws["A2"]
    sub.value = ("📅  Recency (max 50 pts, exponential decay):  "
        "Today→50  |  1d→46  |  2d→42  |  3d→39  |  5d→33  |  "
        "7d→28  |  14d→15  |  21d→8  |  30d→4  |  45d+→1")
    sub.font = Font(name="Arial", size=9, italic=True, color="374151")
    sub.fill = PatternFill("solid", fgColor="F0FDF4")
    sub.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 18

    for ci, (h, w) in enumerate(zip(HEADERS, WIDTHS), 1):
        cell = ws.cell(row=3, column=ci, value=h)
        cell.font      = Font(name="Arial", bold=True, size=10, color=WHITE)
        cell.fill      = PatternFill("solid", fgColor=TEAL)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[3].height = 26

    for ri, job in enumerate(jobs, start=4):
        bg  = LGRAY if ri % 2 == 0 else WHITE
        bf  = PatternFill("solid", fgColor=bg)
        ws.row_dimensions[ri].height = 17
        sc  = job["score"]; days = job["days"]

        def cell(col, val, fnt=None, fil=None, aln=None):
            c = ws.cell(row=ri, column=col, value=val)
            c.border = bdr
            c.font   = fnt or Font(name="Arial", size=10)
            c.fill   = fil or bf
            c.alignment = aln or Alignment(vertical="center")
            return c

        c1 = ws.cell(row=ri, column=1, value=sc)
        c1.font = _sfont(sc); c1.fill = _sfill(sc); c1.border = bdr
        c1.alignment = Alignment(horizontal="center", vertical="center")

        cell(2, job["company"], fnt=Font(name="Arial", bold=True, size=10))
        cell(3, re.sub(r"[🎓🛂🇺🇸🔥🔒]", "", job["title"]).strip())
        cell(4, job["location"])

        sal = job.get("salary", "")
        cell(5, sal, fnt=Font(name="Arial", size=10, color="166534" if sal else "9CA3AF"),
             aln=Alignment(horizontal="center", vertical="center"))

        src = job["source"]
        cell(6, src, fnt=Font(name="Arial", size=9, bold=True, color=SOURCE_FG.get(src,"374151")),
             fil=PatternFill("solid", fgColor=SOURCE_BG.get(src,"F3F4F6")),
             aln=Alignment(horizontal="center", vertical="center"))

        if days <= 2:    p_clr, p_bold = "065F46", True
        elif days <= 5:  p_clr, p_bold = "166534", True
        elif days <= 10: p_clr, p_bold = "B45309", True
        elif days <= 21: p_clr, p_bold = "92400E", False
        else:            p_clr, p_bold = "9CA3AF", False
        cell(7, job["age_str"] or "—",
             fnt=Font(name="Arial", size=9, color=p_clr, bold=p_bold),
             aln=Alignment(horizontal="center", vertical="center"))

        if job["us_citizen"]:
            vv,vfil,vfnt = "❌ Citizen Only", PatternFill("solid",fgColor=RED_BG), Font(name="Arial",size=9,color=RED_FG,bold=True)
        elif job["no_sponsor"]:
            vv,vfil,vfnt = "⚠️ No Sponsor",  PatternFill("solid",fgColor=ORG_BG), Font(name="Arial",size=9,color=ORG_FG,bold=True)
        else:
            vv,vfil,vfnt = "✅ Likely OK",    PatternFill("solid",fgColor=GRN_BG), Font(name="Arial",size=9,color=GRN_FG)
        cell(8, vv, fnt=vfnt, fil=vfil, aln=Alignment(horizontal="center", vertical="center"))

        url = job["apply_url"]
        c9  = ws.cell(row=ri, column=9); c9.border = bdr; c9.fill = bf
        if url:
            c9.value = "Apply →"; c9.hyperlink = url
            c9.font  = Font(name="Arial", size=10, color=LINK, underline="single")
        else:
            c9.value = "—"; c9.font = Font(name="Arial", size=10, color="9CA3AF")
        c9.alignment = Alignment(horizontal="center", vertical="center")

        cell(10, job["score_reason"],
             fnt=Font(name="Arial", size=9, color="4B5563", italic=True),
             aln=Alignment(vertical="center", wrap_text=True))

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:J{len(jobs)+3}"

    ls = wb.create_sheet("Scoring Guide")
    legend = [
        ("RECENCY (max 50 pts — exponential decay)", "score = 50 × e^(−0.085 × days)"),
        ("Today","50 pts"),("1 day ago","46 pts"),("2 days ago","42 pts"),
        ("3 days ago","39 pts"),("5 days ago","33 pts"),("7 days ago","28 pts"),
        ("10 days ago","21 pts"),("14 days ago","15 pts"),("21 days ago","8 pts"),
        ("30 days ago","4 pts"),("45+ days","1 pt"),("",""),
        ("OTHER FACTORS",""),
        ("AI lab / hot startup (NVIDIA, Anthropic, etc.)","+20"),
        ("Big tech (Google, Meta, Amazon, etc.)","+10  ← reduced"),
        ("AI/ML/Automation role title","+25  ← highest role weight"),("SWE/Dev/Systems role","+12"),
        ("Top tech hub (SF, NYC, Seattle, etc.)","+8"),
        ("Remote role","+7"),("Mid-tier tech city (Austin, Chicago, etc.)","+4"),
        ("",""),("PENALTIES",""),
        ("No visa sponsorship 🛂","−20"),("US citizen only 🇺🇸","−30"),
        ("Grad-only role 🎓","EXCLUDED — filtered out entirely"),
        ("",""),("VISA COLUMN",""),
        ("✅ Likely OK","No flag in source — verify on company site"),
        ("⚠️ No Sponsor","🛂 flag — does NOT sponsor visas"),
        ("❌ Citizen Only","🇺🇸 flag — requires US citizenship"),
        ("",""),("SOURCES",""),
        ("Simplify","github.com/SimplifyJobs/Summer2026-Internships  (summer)"),
        ("Simplify-OS","github.com/SimplifyJobs/Summer2026-Internships  (off-season: fall/winter/spring)"),
        ("SpeedyApply","github.com/speedyapply/2026-SWE-College-Jobs"),
        ("Speedy2027","github.com/speedyapply/2027-SWE-College-Jobs"),
        ("Vansh/Ouckah","github.com/vanshb03/Summer2027-Internships"),
        ("Sndsh","github.com/sndsh404/summer-2027-internships"),
        ("YC","workatastartup.com/companies?jobType=intern&role=eng&sortBy=created_desc"),
        ("",""),("TIP","Filter column H (Visa) to hide ⚠️ and ❌ rows instantly"),
    ]
    sections = {"RECENCY (max 50 pts — exponential decay)","OTHER FACTORS","PENALTIES","VISA COLUMN","SOURCES","TIP"}
    for ri2, (k, v) in enumerate(legend, 1):
        ck = ls.cell(row=ri2, column=1, value=k)
        cv = ls.cell(row=ri2, column=2, value=v)
        if k in sections:
            for cc in (ck, cv):
                cc.font = Font(name="Arial", bold=True, size=11, color=WHITE)
                cc.fill = PatternFill("solid", fgColor=NAVY)
        else:
            ck.font = Font(name="Arial", size=10)
            cv.font = Font(name="Arial", size=10,
                color="166534" if ("pts" in v or "pt " in v)
                else ("9F1239" if ("−" in v or "EXCL" in v) else "374151"))
    ls.column_dimensions["A"].width = 52
    ls.column_dimensions["B"].width = 46

    wb.save(output_path)
    print(f"✅  Saved → {output_path}  ({len(jobs)} internships)")


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print(f"📅  Running for {TODAY}\n")
    all_jobs = []

    for name, url in GITHUB_SOURCES.items():
        print(f"  [{name}]  fetching...")
        try:
            content = fetch(url)
        except Exception as e:
            print(f"    ⚠️  Failed: {e}"); continue

        if name == "Simplify":       jobs = parse_simplify(content, "Simplify")
        elif name == "Simplify-OS":  jobs = parse_simplify(content, "Simplify-OS")
        elif name == "SpeedyApply":  jobs = parse_speedyapply(content)
        elif name == "Speedy2027":   jobs = parse_speedyapply(content, "Speedy2027")
        elif name == "Vansh":        jobs = parse_vansh(content)
        elif name == "Sndsh":        jobs = parse_sndsh(content)
        else:                        jobs = []

        print(f"    → {len(jobs)} matching roles")
        all_jobs.extend(jobs)

    print(f"\n  [YC]  fetching from workatastartup.com...")
    try:
        yc_jobs = parse_yc()
        print(f"    → {len(yc_jobs)} matching roles")
        all_jobs.extend(yc_jobs)
    except Exception as e:
        print(f"    ⚠️  Failed: {e}")

    # ── Deduplication ──────────────────────────────────────────────────────────
    # Key: (company_normalized, first_6_title_words_normalized)
    # Catches minor wording differences like "Software Engineer Intern" vs
    # "Software Engineering Intern" from different sources.
    # When a duplicate is found, keep whichever has the LOWER days (fresher).
    seen    = {}   # key -> index in unique
    unique  = []
    for j in all_jobs:
        key = dedup_key(j)
        if key not in seen:
            seen[key] = len(unique)
            unique.append(j)
        else:
            # Replace if this copy is fresher
            idx = seen[key]
            if j["days"] < unique[idx]["days"]:
                unique[idx] = j

    unique.sort(key=lambda x: (-x["score"], x["days"]))

    print(f"\n  Total unique internships: {len(unique)}")
    print("  Score distribution:")
    for lo, hi in [(80,101),(60,80),(40,60),(20,40),(0,20)]:
        n = sum(1 for j in unique if lo <= j["score"] < hi)
        print(f"    {lo:>3}–{hi-1:<3}: {n}")
    print()
    build_excel(unique, OUTPUT_FILE)


if __name__ == "__main__":
    main()