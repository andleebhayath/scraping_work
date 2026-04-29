#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════════════════╗
║         WHED Pakistan Universities Scraper — Production Level               ║
║  Features: Pagination, Detail Scraping, Retry, Resume, Anti-Bot, CSV+JSON  ║
╚══════════════════════════════════════════════════════════════════════════════╝

Usage:
    1. Update COOKIES section with your browser cookies (cf_clearance + PHPSESSID)
    2. pip install requests beautifulsoup4 lxml
    3. python whed_scraper.py

Auto-resumes from last checkpoint if interrupted.
"""

import requests
from bs4 import BeautifulSoup
import json
import csv
import time
import random
import logging
import re
import sys
import os
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

# ══════════════════════════════════════════════════════════
#  CONFIGURATION  — Edit this section before running
# ══════════════════════════════════════════════════════════

# ⚠️  Paste your fresh cookies here (they expire — update when blocked)
COOKIES = {
    "cf_clearance": "E5eqvFiHCnEg8dWqNqW0isSPsXIGrscP2cOZpMy7xlI-1775453571-1.2.1.1-3j1qQGxPlNpI5zJbwZU1rDaNQrKBE51pFd699wDKPlLA4S63qkMc2imB5jukpgyCoVKx_qiTQlbj07hg0p97psgQvoddIXGtOGawlK7SGifrhSNP0mVKjQqNJCRpSrFpVj7atsbzPwvGNdtQFCshA2Pma5SGJ8.hByCpDZpyV8rdRilWJuLjdHeUBQa2seOc3q101blp.3zj19U22Slwlrhb6vjR2tuuh6m.yjWfiXMa1wouypU1wkMwMuml5ASJgHmGkRBs2H.ODYQmKTOnS4VY2JZmIbL3toP5FiAQocR9t0g4DCHkaOCEbC4ARebB3n01BVuSB7BHF6sru4wefQ",
    "PHPSESSID": "k4o6ut440nithdd7pjbkquohjg",
    "accepte_cookie": "1",
}

# Country to scrape
COUNTRY = "Pakistan"

# Delay between requests (seconds) — tune to avoid blocks
DELAY_MIN = 2.5
DELAY_MAX = 5.5

# Retry config
MAX_RETRIES = 4
RETRY_BACKOFF_BASE = 8   # seconds; doubles each retry
RETRY_JITTER       = 3   # random jitter added to backoff

# Output
OUTPUT_DIR     = Path("whed_output")
LOG_FILE       = OUTPUT_DIR / "scraper.log"
PROGRESS_FILE  = OUTPUT_DIR / "progress.json"
JSON_OUT       = OUTPUT_DIR / "institutions.json"
CSV_OUT        = OUTPUT_DIR / "institutions.csv"

# ══════════════════════════════════════════════════════════
#  CONSTANTS
# ══════════════════════════════════════════════════════════

BASE_URL   = "https://whed.net"
LIST_URL   = f"{BASE_URL}/results_institutions.php"
DETAIL_URL = f"{BASE_URL}/detail_institution.php"

# Rotate user-agents to appear more human-like
USER_AGENTS = [
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/145.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
    "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:124.0) Gecko/20100101 Firefox/124.0",
]

# ══════════════════════════════════════════════════════════
#  SETUP: Logging + Output dirs
# ══════════════════════════════════════════════════════════

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("whed")


# ══════════════════════════════════════════════════════════
#  SESSION FACTORY
# ══════════════════════════════════════════════════════════

def make_session() -> requests.Session:
    """Create a requests Session that looks like a real browser."""
    session = requests.Session()
    session.cookies.update(COOKIES)
    return session


def get_headers(referer: str = LIST_URL, is_iframe: bool = False) -> dict:
    """Return randomised but realistic browser headers."""
    ua = random.choice(USER_AGENTS)
    h = {
        "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "accept-language": "en-US,en;q=0.9",
        "cache-control": "no-cache" if random.random() > 0.5 else "max-age=0",
        "dnt": "1",
        "pragma": "no-cache",
        "referer": referer,
        "sec-ch-ua": '"Chromium";v="146", "Not-A.Brand";v="24", "Google Chrome";v="146"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"Linux"',
        "sec-fetch-dest": "iframe" if is_iframe else "document",
        "sec-fetch-mode": "navigate",
        "sec-fetch-site": "same-origin",
        "sec-fetch-user": "?1",
        "upgrade-insecure-requests": "1",
        "user-agent": ua,
    }
    return h


# ══════════════════════════════════════════════════════════
#  HUMAN-LIKE DELAY
# ══════════════════════════════════════════════════════════

def human_delay(min_s: float = DELAY_MIN, max_s: float = DELAY_MAX):
    """Sleep for a random amount simulating human reading time."""
    t = random.uniform(min_s, max_s)
    time.sleep(t)


# ══════════════════════════════════════════════════════════
#  SAFE HTTP FETCH  (retry + backoff + bot-detection check)
# ══════════════════════════════════════════════════════════

def safe_get(session: requests.Session, url: str, **kwargs) -> requests.Response | None:
    """
    GET with exponential-backoff retry.
    Returns Response on success, None if all retries exhausted.
    """
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.get(url, timeout=30, allow_redirects=True, **kwargs)

            # ── Bot / Cloudflare detection ────────────────
            if resp.status_code == 403:
                log.warning(f"403 Forbidden — likely Cloudflare block. Update cookies! (attempt {attempt})")
                _backoff(attempt)
                continue

            if resp.status_code == 429:
                wait = int(resp.headers.get("Retry-After", RETRY_BACKOFF_BASE * (2 ** attempt)))
                log.warning(f"429 Rate-limit. Sleeping {wait}s …")
                time.sleep(wait + random.uniform(0, RETRY_JITTER))
                continue

            if resp.status_code == 503:
                log.warning(f"503 Service Unavailable (attempt {attempt}). Backing off …")
                _backoff(attempt)
                continue

            if resp.status_code != 200:
                log.warning(f"HTTP {resp.status_code} for {url} (attempt {attempt})")
                _backoff(attempt)
                continue

            # Content-based bot detection
            text_low = resp.text.lower()
            if "just a moment" in text_low and "cloudflare" in text_low:
                log.error("Cloudflare JS challenge detected — cookies may have expired!")
                _backoff(attempt, long=True)
                continue

            if "access denied" in text_low and len(resp.text) < 1000:
                log.warning(f"Access denied page (attempt {attempt}). Backing off …")
                _backoff(attempt)
                continue

            return resp

        except requests.exceptions.SSLError as e:
            log.error(f"SSL error on {url}: {e}")
            _backoff(attempt)
        except requests.exceptions.ConnectionError as e:
            log.warning(f"Connection error (attempt {attempt}): {e}")
            _backoff(attempt)
        except requests.exceptions.Timeout:
            log.warning(f"Timeout on {url} (attempt {attempt})")
            _backoff(attempt)
        except Exception as e:
            log.error(f"Unexpected error on {url}: {e}")
            _backoff(attempt)

    log.error(f"All {MAX_RETRIES} retries failed for: {url}")
    return None


def safe_post(session: requests.Session, url: str, data: dict, **kwargs) -> requests.Response | None:
    """POST with same retry logic."""
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.post(url, data=data, timeout=30, allow_redirects=True, **kwargs)

            if resp.status_code == 403:
                log.warning(f"403 — Cloudflare block on POST (attempt {attempt}). Update cookies!")
                _backoff(attempt)
                continue

            if resp.status_code == 429:
                wait = int(resp.headers.get("Retry-After", RETRY_BACKOFF_BASE * (2 ** attempt)))
                log.warning(f"429 Rate-limit. Sleeping {wait}s …")
                time.sleep(wait + random.uniform(0, RETRY_JITTER))
                continue

            if resp.status_code != 200:
                log.warning(f"HTTP {resp.status_code} on POST (attempt {attempt})")
                _backoff(attempt)
                continue

            text_low = resp.text.lower()
            if "just a moment" in text_low and "cloudflare" in text_low:
                log.error("Cloudflare JS challenge on POST — update cookies!")
                _backoff(attempt, long=True)
                continue

            return resp

        except requests.exceptions.ConnectionError as e:
            log.warning(f"Connection error POST (attempt {attempt}): {e}")
            _backoff(attempt)
        except requests.exceptions.Timeout:
            log.warning(f"Timeout on POST (attempt {attempt})")
            _backoff(attempt)
        except Exception as e:
            log.error(f"Unexpected POST error: {e}")
            _backoff(attempt)

    log.error("All retries failed for POST.")
    return None


def _backoff(attempt: int, long: bool = False):
    """Exponential backoff with jitter."""
    base = RETRY_BACKOFF_BASE * (2 ** (attempt - 1))
    if long:
        base *= 3
    sleep_time = base + random.uniform(0, RETRY_JITTER)
    log.info(f"  ↳ Backoff: sleeping {sleep_time:.1f}s …")
    time.sleep(sleep_time)


# ══════════════════════════════════════════════════════════
#  PROGRESS TRACKER  (resume support)
# ══════════════════════════════════════════════════════════

def load_progress() -> dict:
    if PROGRESS_FILE.exists():
        try:
            data = json.loads(PROGRESS_FILE.read_text(encoding="utf-8"))
            log.info(f"Resuming from checkpoint: {data.get('scraped_count', 0)} already done")
            return data
        except Exception:
            pass
    return {"scraped_ids": [], "scraped_count": 0, "institutions_index": []}


def save_progress(progress: dict):
    try:
        PROGRESS_FILE.write_text(json.dumps(progress, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception as e:
        log.error(f"Could not save progress: {e}")


# ══════════════════════════════════════════════════════════
#  HTML PARSERS
# ══════════════════════════════════════════════════════════

def parse_list_page(html: str) -> tuple[list[dict], int]:
    """
    Parse the results list page.
    Returns (list of institution stubs, total count).
    Each stub: {iau_id, name, abbreviation, detail_url}
    """
    soup = BeautifulSoup(html, "lxml")
    institutions = []

    # Extract total count
    total = 0
    info_p = soup.find("p", class_="infos")
    if info_p:
        m = re.search(r"(\d+)\s+results", info_p.text)
        if m:
            total = int(m.group(1))

    # Each institution is inside <li class="iaumember ..."> or <li class="even ...">
    # The IAU ID appears in a <span class="gui"> right before each <li>
    results_ul = soup.find("ul", id="results")
    if not results_ul:
        log.warning("Could not find <ul id='results'> — page structure may have changed")
        return [], total

    # Walk all children of the ul
    iau_id_pattern = re.compile(r"IAU-(\d+)")
    current_iau_id = None

    for element in results_ul.children:
        if not hasattr(element, "name"):
            continue

        if element.name == "span" and "gui" in element.get("class", []):
            m = iau_id_pattern.search(element.text)
            if m:
                current_iau_id = f"IAU-{m.group(1)}"

        elif element.name == "li":
            h3 = element.find("h3")
            if not h3:
                continue

            link = h3.find("a")
            if not link:
                continue

            name = link.get_text(strip=True)
            detail_href = link.get("href", "")

            # Abbreviation
            abbr_p = element.find("p", class_="i_name")
            abbreviation = ""
            if abbr_p:
                abbr_text = abbr_p.get_text(strip=True)
                m = re.search(r"\((.+?)\)", abbr_text)
                if m:
                    abbreviation = m.group(1)

            # Build full detail URL
            if detail_href and not detail_href.startswith("http"):
                detail_href = urljoin(BASE_URL + "/", detail_href)

            institutions.append({
                "iau_id": current_iau_id or "",
                "name": name,
                "abbreviation": abbreviation,
                "detail_url": detail_href,
                "is_iau_member": "iaumember" in element.get("class", []),
            })

    return institutions, total


def parse_detail_page(html: str, iau_id: str) -> dict:
    """
    Parse a detail page into a structured dict.
    Handles missing sections gracefully.
    """
    soup = BeautifulSoup(html, "lxml")
    data: dict = {"iau_id": iau_id}

    # ── Name & abbreviation ──────────────────────────────
    name_div = soup.find("div", class_="detail_right")
    if name_div:
        divs = name_div.find_all("div")
        if divs:
            data["name"] = divs[0].get_text(strip=True)
        detail_name = name_div.find("div", class_="detail_name")
        if detail_name:
            abbr = detail_name.get_text(strip=True).strip("()")
            data["abbreviation"] = abbr

    # ── Country ──────────────────────────────────────────
    country_p = soup.find("p", class_="country")
    if country_p:
        data["country"] = country_p.get_text(strip=True)

    # ── IAU Permalink ────────────────────────────────────
    permalink_a = soup.find("a", href=lambda h: h and "whed.net/institutions/" in h)
    if permalink_a:
        data["permalink"] = permalink_a["href"]

    # ── Parse all dl blocks (key-value sections) ─────────
    # The page uses <div class="dl"><span class="dt">Key</span><div class="dd">…</div></div>
    general = {}
    officers = []
    divisions = []
    degrees = []
    periodicals = []
    students = {}

    dl_blocks = soup.find_all("div", class_="dl")

    for dl in dl_blocks:
        dt = dl.find("span", class_="dt")
        dd = dl.find("div", class_="dd")
        if not dt or not dd:
            continue

        key = dt.get_text(strip=True).rstrip(":")
        if not key:
            # No label — check for section by what h3 precedes this
            # This usually means it's inside Officers / Divisions / Degrees
            pass

        # ── Detect section by preceding h3 ──────────────
        h3 = dl.find_previous("h3")
        section_text = h3.get_text(strip=True).lower() if h3 else ""

        if "general" in section_text or not h3:
            # General info fields
            _parse_general_dl(key, dd, general)

        elif "officer" in section_text:
            _parse_officers_dl(dd, officers)

        elif "division" in section_text or "department" in section_text:
            _parse_divisions_dl(dd, divisions)

        elif "degree" in section_text:
            _parse_degrees_dl(dd, degrees)

        elif "periodical" in section_text:
            _parse_periodicals_dl(key, dd, periodicals)

        elif "student" in section_text or "staff" in section_text:
            _parse_students_dl(key, dd, students)

    data["general_info"] = general
    data["officers"] = officers
    data["divisions"] = divisions
    data["degrees"] = degrees
    data["periodicals"] = periodicals
    data["students_staff"] = students

    # ── Last updated ─────────────────────────────────────
    updated_p = soup.find("p", class_="right italik")
    if updated_p:
        data["last_updated"] = updated_p.get_text(strip=True).replace("Updated on", "").strip()

    return data


def _parse_general_dl(key: str, dd, out: dict):
    """Extract general info fields from a dd block."""
    if not key:
        return

    # Address is special — multiple sub-fields
    if key.lower() == "address":
        address = {}
        for p in dd.find_all("p"):
            libelle = p.find("span", class_="libelle")
            contenu = p.find("span", class_="contenu")
            if libelle and contenu:
                sub_key = libelle.get_text(strip=True).rstrip(":")
                val = contenu.get_text(" ", strip=True)
                address[sub_key.lower().replace(" ", "_")] = val
            elif not libelle:
                # Could be WWW link
                a = p.find("a")
                if a:
                    address["www"] = a.get_text(strip=True)
        out["address"] = address
        return

    # Tuition fees — multiple sub-fields
    if "tuition" in key.lower():
        fees = {}
        for p in dd.find_all("p"):
            libelle = p.find("span", class_="libelle")
            contenu = p.find("span", class_="contenu")
            if libelle and contenu:
                sub_key = libelle.get_text(strip=True).rstrip(":")
                fees[sub_key.lower()] = contenu.get_text(strip=True)
        out["tuition_fees"] = fees if fees else dd.get_text(strip=True)
        return

    # Standard single-value fields
    val = dd.get_text(" ", strip=True)
    val = re.sub(r"\s+", " ", val).strip()
    if val:
        safe_key = key.lower().replace(" ", "_").replace("/", "_").replace("(", "").replace(")", "")
        out[safe_key] = val


def _parse_officers_dl(dd, out: list):
    """Extract officer records separated by <hr>."""
    current = {}
    for child in dd.children:
        if not hasattr(child, "name"):
            continue
        if child.name == "hr":
            if current:
                out.append(current)
                current = {}
        elif child.name == "p":
            cls = child.get("class", [])
            if "principal" in cls:
                # e.g. "Head : Saima Hamid"
                txt = child.get_text(strip=True)
                if ":" in txt:
                    role, name = txt.split(":", 1)
                    current["role"] = role.strip()
                    current["name"] = name.strip()
                else:
                    current["name"] = txt
            else:
                libelle = child.find("span", class_="libelle")
                contenu = child.find("span", class_="contenu")
                if libelle and contenu:
                    k = libelle.get_text(strip=True).rstrip(":")
                    v = contenu.get_text(strip=True)
                    current[k.lower().replace(" ", "_")] = v
    if current:
        out.append(current)


def _parse_divisions_dl(dd, out: list):
    """Extract department/division records."""
    current = {}
    for child in dd.children:
        if not hasattr(child, "name"):
            continue
        if child.name == "hr":
            if current:
                out.append(current)
                current = {}
        elif child.name == "p":
            cls = child.get("class", [])
            if "principal" in cls:
                txt = child.get_text(strip=True)
                if ":" in txt:
                    dtype, name = txt.split(":", 1)
                    current["type"] = dtype.strip()
                    current["name"] = name.strip()
                else:
                    current["name"] = txt
            else:
                libelle = child.find("span", class_="libelle")
                contenu = child.find("span", class_="contenu")
                if libelle and contenu:
                    k = libelle.get_text(strip=True).rstrip(":")
                    v = contenu.get_text(strip=True)
                    current[k.lower().replace(" ", "_")] = v
    if current:
        out.append(current)


def _parse_degrees_dl(dd, out: list):
    """Extract degree types and their fields."""
    current = {}
    for child in dd.children:
        if not hasattr(child, "name"):
            continue
        if child.name == "hr":
            if current:
                out.append(current)
                current = {}
        elif child.name == "p":
            cls = child.get("class", [])
            if "principal" in cls:
                current["degree"] = child.get_text(strip=True)
            else:
                libelle = child.find("span", class_="libelle")
                contenu = child.find("span", class_="contenu")
                if libelle and contenu:
                    current["fields_of_study"] = contenu.get_text(strip=True)
    if current:
        out.append(current)


def _parse_periodicals_dl(key: str, dd, out: list):
    for p in dd.find_all("p", class_="principal"):
        out.append(p.get_text(strip=True))


def _parse_students_dl(key: str, dd, out: dict):
    for p in dd.find_all("p"):
        libelle = p.find("span", class_="libelle")
        contenu = p.find("span", class_="contenu")
        if libelle and contenu:
            k = libelle.get_text(strip=True).rstrip(":")
            out[k.lower().replace(" ", "_")] = contenu.get_text(strip=True)
        elif not libelle:
            val = p.get_text(strip=True)
            if val:
                out[key.lower().replace(" ", "_")] = val


# ══════════════════════════════════════════════════════════
#  PHASE 1 — Scrape institution list (all pages)
# ══════════════════════════════════════════════════════════

def scrape_all_list_pages(session: requests.Session) -> list[dict]:
    """
    Fetch all paginated list pages and return combined institution stubs.
    POST is used (as the site requires form submission).
    """
    log.info("=" * 60)
    log.info(f"Phase 1: Scraping institution list for country='{COUNTRY}'")
    log.info("=" * 60)

    all_institutions = []
    offset = 0
    total = None

    while True:
        log.info(f"  Fetching list page offset={offset} …")

        form_data = {
            "Chp1": COUNTRY,
            "Chp0": "",
            "Chp2": "",
            "Chp4": "",
            "debut": str(offset),
            "nbr_ref_pge": "10",
            "sort": "InstNameEnglish,iBranchName",
            "afftri": "yes",
        }

        resp = safe_post(
            session,
            LIST_URL,
            data=form_data,
            headers=get_headers(referer=LIST_URL),
        )

        if resp is None:
            log.error(f"  Failed to fetch list page at offset={offset}. Stopping list scrape.")
            break

        page_institutions, page_total = parse_list_page(resp.text)

        if total is None and page_total:
            total = page_total
            log.info(f"  Total institutions found: {total}")

        if not page_institutions:
            log.warning(f"  No institutions parsed at offset={offset} — may have reached end.")
            break

        log.info(f"  Parsed {len(page_institutions)} institutions from this page")
        all_institutions.extend(page_institutions)

        offset += len(page_institutions)

        if total and offset >= total:
            log.info("  All list pages fetched.")
            break

        # Be polite between list pages
        human_delay(DELAY_MIN, DELAY_MAX)

    log.info(f"Phase 1 complete: {len(all_institutions)} institutions collected")
    return all_institutions


# ══════════════════════════════════════════════════════════
#  PHASE 2 — Scrape detail for each institution
# ══════════════════════════════════════════════════════════

def scrape_all_details(
    session: requests.Session,
    institutions_index: list[dict],
    progress: dict,
) -> list[dict]:
    """
    For each institution stub, fetch its detail page and parse it.
    Skips already-scraped IDs (resume support).
    """
    log.info("=" * 60)
    log.info("Phase 2: Scraping detail pages")
    log.info("=" * 60)

    scraped_ids = set(progress.get("scraped_ids", []))
    all_details = []

    # Load previously saved data to avoid re-writing
    if JSON_OUT.exists():
        try:
            all_details = json.loads(JSON_OUT.read_text(encoding="utf-8"))
            log.info(f"  Loaded {len(all_details)} previously saved records")
        except Exception:
            all_details = []

    total = len(institutions_index)

    for idx, stub in enumerate(institutions_index, 1):
        iau_id = stub.get("iau_id", "")
        name   = stub.get("name", "?")
        url    = stub.get("detail_url", "")

        if iau_id in scraped_ids:
            log.info(f"  [{idx}/{total}] SKIP (already done): {iau_id} — {name}")
            continue

        if not url:
            log.warning(f"  [{idx}/{total}] No detail URL for {iau_id} — {name}, skipping")
            continue

        log.info(f"  [{idx}/{total}] Fetching: {iau_id} — {name}")

        resp = safe_get(
            session,
            url,
            headers=get_headers(referer=LIST_URL, is_iframe=True),
        )

        if resp is None:
            log.error(f"    ✗ Failed to fetch detail for {iau_id}. Recording partial record.")
            record = {**stub, "error": "fetch_failed", "scraped_at": datetime.now().isoformat()}
        else:
            try:
                detail = parse_detail_page(resp.text, iau_id)
                record = {**stub, **detail, "scraped_at": datetime.now().isoformat()}
                log.info(f"    ✓ OK — {len(detail.get('divisions', []))} divisions, "
                         f"{len(detail.get('degrees', []))} degree types")
            except Exception as e:
                log.error(f"    ✗ Parse error for {iau_id}: {e}")
                record = {**stub, "error": f"parse_error:{e}", "scraped_at": datetime.now().isoformat()}

        all_details.append(record)
        scraped_ids.add(iau_id)

        # Save after every institution (self-healing checkpoint)
        progress["scraped_ids"] = list(scraped_ids)
        progress["scraped_count"] = len(scraped_ids)
        save_progress(progress)
        _save_json(all_details)
        _save_csv(all_details)

        if idx < total:
            human_delay(DELAY_MIN, DELAY_MAX)

    log.info(f"Phase 2 complete: {len(all_details)} total records")
    return all_details


# ══════════════════════════════════════════════════════════
#  OUTPUT WRITERS
# ══════════════════════════════════════════════════════════

def _save_json(records: list[dict]):
    try:
        JSON_OUT.write_text(
            json.dumps(records, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
    except Exception as e:
        log.error(f"JSON save failed: {e}")


def _save_csv(records: list[dict]):
    """
    Flatten nested structures for CSV export.
    Complex nested fields (divisions, degrees, officers) are JSON-stringified.
    """
    if not records:
        return

    flat_records = []
    for r in records:
        gi = r.get("general_info", {})
        addr = gi.get("address", {})
        fees = gi.get("tuition_fees", {})
        students = r.get("students_staff", {})

        flat = {
            "iau_id": r.get("iau_id", ""),
            "name": r.get("name", ""),
            "abbreviation": r.get("abbreviation", ""),
            "is_iau_member": r.get("is_iau_member", ""),
            "country": r.get("country", ""),
            "permalink": r.get("permalink", ""),
            "detail_url": r.get("detail_url", ""),
            # Address
            "city": addr.get("city", ""),
            "post_code": addr.get("post_code", ""),
            "street": addr.get("street", ""),
            "www": addr.get("www", ""),
            # General
            "institution_funding": gi.get("institution_funding", ""),
            "history": gi.get("history", ""),
            "academic_year": gi.get("academic_year", ""),
            "admission_requirements": gi.get("admission_requirements", ""),
            "tuition_national": fees.get("national", "") if isinstance(fees, dict) else fees,
            "language": gi.get("language_s_", "") or gi.get("languages", ""),
            "accrediting_agency": gi.get("accrediting_agency", ""),
            "religious_affiliation": gi.get("religious_affiliation", ""),
            "student_body": gi.get("student_body", ""),
            # Counts
            "total_students": students.get("total", ""),
            "statistics_year": students.get("statistics_year", ""),
            # JSON blobs for complex fields
            "officers_json": json.dumps(r.get("officers", []), ensure_ascii=False),
            "divisions_json": json.dumps(r.get("divisions", []), ensure_ascii=False),
            "degrees_json": json.dumps(r.get("degrees", []), ensure_ascii=False),
            "periodicals_json": json.dumps(r.get("periodicals", []), ensure_ascii=False),
            # Meta
            "last_updated": r.get("last_updated", ""),
            "scraped_at": r.get("scraped_at", ""),
            "error": r.get("error", ""),
        }
        flat_records.append(flat)

    try:
        with open(CSV_OUT, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=list(flat_records[0].keys()))
            writer.writeheader()
            writer.writerows(flat_records)
    except Exception as e:
        log.error(f"CSV save failed: {e}")


# ══════════════════════════════════════════════════════════
#  MAIN ENTRYPOINT
# ══════════════════════════════════════════════════════════

def main():
    log.info("╔════════════════════════════════════════════════════╗")
    log.info("║       WHED Scraper  — Starting                     ║")
    log.info(f"║  Country : {COUNTRY:<40}║")
    log.info(f"║  Output  : {str(OUTPUT_DIR):<40}║")
    log.info("╚════════════════════════════════════════════════════╝")

    session  = make_session()
    progress = load_progress()

    # ── Phase 1: Get institution index (unless already saved) ──
    if progress.get("institutions_index"):
        institutions_index = progress["institutions_index"]
        log.info(f"Using cached institution index: {len(institutions_index)} entries")
    else:
        institutions_index = scrape_all_list_pages(session)
        if not institutions_index:
            log.error("No institutions found. Check cookies and try again.")
            sys.exit(1)
        progress["institutions_index"] = institutions_index
        save_progress(progress)

    # ── Phase 2: Get details for each institution ──────────────
    all_records = scrape_all_details(session, institutions_index, progress)

    # ── Final save ─────────────────────────────────────────────
    _save_json(all_records)
    _save_csv(all_records)

    errors = [r for r in all_records if r.get("error")]
    log.info("=" * 60)
    log.info(f"✅  Done!  {len(all_records)} records saved.")
    log.info(f"   JSON → {JSON_OUT}")
    log.info(f"   CSV  → {CSV_OUT}")
    log.info(f"   Log  → {LOG_FILE}")
    if errors:
        log.warning(f"   ⚠️  {len(errors)} records had errors (check log for details)")
    log.info("=" * 60)


if __name__ == "__main__":
    main()
