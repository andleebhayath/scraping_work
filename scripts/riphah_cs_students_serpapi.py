"""
Find Riphah International University CS students via SerpAPI (all campuses).

Uses Google Search through SerpAPI to discover public LinkedIn profiles and
extracts name, LinkedIn URL, headline, location, study year, connections, subject, about,
and any contact info visible in search snippets.

After discovery, each profile is looked up directly via Google/SerpAPI for location,
connections, and Riphah education years.

For fields Google does not index (common for study year), the script can estimate from
"Nth semester at Riphah" text, or you can use the optional Playwright helper with your
own LinkedIn login for exact profile data:

  python scripts/riphah_cs_students_linkedin_playwright.py --input output/riphah_cs_students_serpapi.xlsx

Setup:
  1. Sign up at https://serpapi.com/users/sign_up (free trial includes searches)
  2. Copy your API key from https://serpapi.com/manage-api-key
  3. Put your key in project-root .env as SERPAPI_KEY=...  (gitignored)
     or set environment variable / pass --api-key on the command line

Run (10 test profiles, default):
  python scripts/riphah_cs_students_serpapi.py

Run with custom limit:
  python scripts/riphah_cs_students_serpapi.py --max-profiles 10 --enrich-contacts

Incremental Excel/CSV (default):
  Re-runs append new profiles to the existing output file (skips duplicate LinkedIn URLs).
  The file is saved after each new profile is enriched.

  python scripts/riphah_cs_students_serpapi.py --max-profiles 10   # adds 10 new rows
  python scripts/riphah_cs_students_serpapi.py --refresh-existing   # fix rows already in Excel

Enrich your existing Playwright list (190 URLs) with study year, experience, jobs:
  python scripts/riphah_cs_students_serpapi.py --input output/riphah_cs_students_100.csv --refresh-existing --enrich-contacts

Output:
  output/riphah_cs_students_serpapi.csv
  output/riphah_cs_students_serpapi.xlsx
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
import requests

_SCRIPTS = Path(__file__).resolve().parent
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

import uni_data_free_google as udg  # noqa: E402

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT_CSV = PROJECT_ROOT / "output" / "riphah_cs_students_serpapi.csv"
DEFAULT_OUTPUT_XLSX = PROJECT_ROOT / "output" / "riphah_cs_students_serpapi.xlsx"

SERPAPI_URL = "https://serpapi.com/search.json"

DEFAULT_UNIVERSITY = "Riphah International University"
DEFAULT_CAMPUS = "All campuses"
DEFAULT_DEPARTMENT = "Computer Science"

OUTPUT_COLUMNS = [
    "university",
    "campus",
    "department",
    "student_name",
    "linkedin_url",
    "headline",
    "location",
    "study_year",
    "connections",
    "degree_program",
    "subjects",
    "about",
    "experience_summary",
    "experience_roles",
    "current_job",
    "education_detail",
    "email",
    "phone",
    "result_title",
    "snippet",
    "search_query",
    "profile_rank",
    "data_source",
    "scraped_at",
]

YEAR_RE = re.compile(
    r"\b(?:class of|batch|graduat(?:e|ing)|"
    r"(?:1st|2nd|3rd|4th|5th|6th|7th|8th|first|second|third|fourth|final)\s+year|"
    r"semester\s+\d+)\b",
    re.IGNORECASE,
)
YEAR_RANGE_FULL_RE = re.compile(
    r"((?:19|20)\d{2}\s*[-–—]\s*(?:(?:19|20)\d{2}|Present|present|Current|current))",
    re.IGNORECASE,
)
EDUCATION_YEAR_RANGE_RE = YEAR_RANGE_FULL_RE
# Other schools — do not treat their date ranges as Riphah study years
OTHER_UNIVERSITY_RE = re.compile(
    r"\b(comsats|nust|fast|uet|gcu|pu\b|uol\b|iqra university|ksbl|beaconhouse|"
    r"punjab university|virtual university|air university|bahria|nutech|"
    r"piaic|gift university|city university|comsats university|pucit|ncba|"
    r"kips education|national college)\b",
    re.IGNORECASE,
)
EDUCATION_ENTRY_RE = re.compile(
    r"([A-Za-z0-9][^.;·|]{4,120}?"
    r"(?:University|College|FLC|Institute|School|Academy))"
    r"[^.;·|]*?\.\s*"
    r"([^.;·|]{2,140}?)\s*\.\s*"
    r"((?:19|20)\d{2}\s*[-–—]\s*(?:(?:19|20)\d{2}|Present|present|Current|current))",
    re.IGNORECASE,
)
LOCATION_LABEL_RE = re.compile(
    r"Location:\s*([^·\n|]+?)(?:\s*·|\s*\||$)",
    re.IGNORECASE,
)
CONNECTIONS_RE = re.compile(
    r"(\d[\d,]*\+?)\s*connections?\b",
    re.IGNORECASE,
)
LOCATION_CONNECTIONS_RE = re.compile(
    r"Location:\s*[^·]+·\s*(\d[\d,]*\+?)",
    re.IGNORECASE,
)
CONNECTIONS_ON_LINKEDIN_RE = re.compile(
    r"(\d[\d,]*\+?)\s*connections\s+on\s+LinkedIn",
    re.IGNORECASE,
)
FOLLOWERS_CONNECTIONS_RE = re.compile(
    r"(\d[\d,]*)\s*followers\s+(\d[\d,]*)\s*connections\b",
    re.IGNORECASE,
)
CITY_PAKISTAN_RE = re.compile(
    r"([\w][\w\s\-]*,\s*Pakistan)\.?",
    re.IGNORECASE,
)
CITY_PROVINCE_PAKISTAN_RE = re.compile(
    r"\b("
    r"Islamabad|Rawalpindi|Peshawar|Lahore|Karachi|Multan|Faisalabad|Sialkot|"
    r"Quetta|Hyderabad|Abbottabad|Mardan|Gujranwala|Sargodha|Bahawalpur|"
    r"Jaranwala|Jarānwāla|Gujrat|Sheikhupura|Okara|Sahiwal|Wah|Taxila|"
    r"Mirpur|Murree|Swat|Charsadda|Kohat|Dera Ismail Khan|Sukkur|Larkana"
    r")\s*,\s*(?:Punjab|Sindh|Khyber Pakhtunkhwa|KPK|Balochistan|"
    r"Gilgit[- ]Baltistan|Azad Kashmir)?\s*,?\s*Pakistan\b",
    re.IGNORECASE,
)
CITY_STANDALONE_RE = re.compile(
    r"\b("
    r"Islamabad|Rawalpindi|Peshawar|Lahore|Karachi|Multan|Faisalabad|Sialkot|"
    r"Quetta|Hyderabad|Abbottabad|Mardan|Gujranwala|Sargodha|Bahawalpur|"
    r"Jaranwala|Gujrat|Sheikhupura|Okara|Sahiwal|Wah|Taxila|Mirpur|Murree|"
    r"Swat|Charsadda|Kohat|Dera Ismail Khan|Sukkur|Larkana|Rawalpindi"
    r")\b",
    re.IGNORECASE,
)
PAKISTAN_PROVINCES = frozenset(
    {
        "punjab",
        "sindh",
        "khyber pakhtunkhwa",
        "kpk",
        "balochistan",
        "gilgit-baltistan",
        "gilgit baltistan",
        "azad kashmir",
        "pakistan",
    }
)
LOCATION_NOISE_RE = re.compile(
    r"\b(university|linkedin|notification|follow|message|premium|"
    r"skip to main|my network|for business)\b",
    re.IGNORECASE,
)
DEGREE_RE = re.compile(
    r"\b(B\.?S\.?\s*(?:CS|Computer Science|SE|Software Engineering)|"
    r"BSCS|BSSE|Bachelor(?:'s)?(?:\s+of|\s+in)?\s*(?:Computer Science|Software Engineering|IT)|"
    r"M\.?S\.?\s*(?:CS|Computer Science)|"
    r"Associate(?:'s)?\s+degree|"
    r"Master(?:'s)?(?:\s+of|\s+in)?\s*(?:Computer Science|Software Engineering))\b",
    re.IGNORECASE,
)
LOCATION_RE = re.compile(
    r"\b(?:Islamabad|Rawalpindi|Peshawar|Lahore|Karachi|Multan|Faisalabad|"
    r"Sialkot|Quetta|Hyderabad|Abbottabad|Mardan|Gujranwala|"
    r"Pakistan(?:,\s*[A-Za-z .]+)?|[A-Za-z .]+,\s*Pakistan)\b",
    re.IGNORECASE,
)
FACULTY_RE = re.compile(
    r"\b("
    r"professor|assistant professor|associate professor|lecturer|instructor|"
    r"dean|faculty|teacher|head of department|\bhod\b|"
    r"research scholar|mphil|ph\.?d|doctorate"
    r")\b",
    re.IGNORECASE,
)
STUDENT_RE = re.compile(
    r"\b("
    r"student|undergraduate|studying|enrolled|"
    r"bscs|bs cs|bs computer science|bsse|"
    r"bachelor(?:'s)?(?:\s+student|\s+of|\s+in)?"
    r")\b",
    re.IGNORECASE,
)
CS_RE = re.compile(
    r"\b("
    r"computer science|software engineering|bscs|bs cs|bsse|"
    r"artificial intelligence|machine learning|data science|"
    r"information technology|\bit\b|programming|developer"
    r")\b",
    re.IGNORECASE,
)
ALUMNI_RANGE_RE = re.compile(
    r"(?:19|20)\d{2}\s*[-–]\s*((?:19|20)\d{2})",
    re.IGNORECASE,
)
STAFF_SIGNAL_RE = re.compile(
    r"\b(leadership role|dean|faculty member|teaching|professor|lecturer|"
    r"head of|program coordinator|advise startups)\b",
    re.IGNORECASE,
)
SEMESTER_RE = re.compile(
    r"\b(\d{1,2})(?:st|nd|rd|th)?\s+semester\b",
    re.IGNORECASE,
)
BSCS_DURATION_YEARS = 4
SUBJECT_HINTS = (
    "computer science",
    "software engineering",
    "artificial intelligence",
    "machine learning",
    "data science",
    "cyber security",
    "cybersecurity",
    "information technology",
    "web development",
    "programming",
    "networking",
    "database",
    "cloud computing",
    "mobile development",
)


def load_env_file(env_path: Path) -> None:
    """Load KEY=VALUE lines from .env into os.environ (does not override existing vars)."""
    if not env_path.is_file():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def normalize_linkedin_profile_url(url: str) -> str:
    url = (url or "").strip()
    if not url:
        return ""
    parsed = urlparse(url.split("?", 1)[0])
    path = parsed.path.rstrip("/")
    if "/in/" not in path:
        return ""
    slug = path.split("/in/", 1)[1].split("/")[0]
    if not slug:
        return ""
    return f"https://www.linkedin.com/in/{slug}"


def is_linkedin_profile(url: str) -> bool:
    return udg.is_profile_like_linkedin(url)


def name_from_title(title: str) -> str:
    cleaned = re.sub(r"\s*\|\s*LinkedIn\s*$", "", title, flags=re.IGNORECASE).strip()
    for sep in (" - ", " – ", " | ", " — "):
        if sep in cleaned:
            return cleaned.split(sep, 1)[0].strip()
    return cleaned


def name_from_linkedin_url(url: str) -> str:
    try:
        path = urlparse(url).path.strip("/")
        if not path.startswith("in/"):
            return ""
        slug = path.split("/", 1)[1].split("/")[0]
        slug = re.sub(r"-?\d+$", "", slug)
        words = [w for w in slug.replace("-", " ").split() if w and not w.isdigit()]
        if len(words) < 2:
            return ""
        return " ".join(w.capitalize() for w in words[:5])
    except Exception:
        return ""


def contacts_from_text(text: str) -> tuple[str, str]:
    emails = udg.extract_emails(text, limit=3)
    phones = udg.extract_phones(text, limit=2)
    return ("; ".join(emails), phones[0] if phones else "")


def parse_headline(title: str, snippet: str) -> str:
    cleaned = re.sub(r"\s*\|\s*LinkedIn\s*$", "", title, flags=re.IGNORECASE).strip()
    for sep in (" - ", " – ", " — "):
        if sep in cleaned:
            part = cleaned.split(sep, 1)[1].strip()
            if part and "linkedin" not in part.lower():
                return part
    if " | " in cleaned:
        part = cleaned.split(" | ", 1)[1].strip()
        if part and "linkedin" not in part.lower():
            return part
    first_line = snippet.split(".")[0].strip() if snippet else ""
    if first_line and len(first_line) <= 120:
        return first_line
    return ""


def _line_looks_like_location(line: str) -> bool:
    low = line.lower()
    if re.search(r"\bconnections?\b", low) and re.search(r"\d|\+", line):
        return True
    if re.search(r",\s*(punjab|sindh|balochistan|kpk|pakistan)\b", low):
        return True
    if "district" in low and "pakistan" in low:
        return True
    return low in {"contact info", "·"} or low.startswith("contact info")


def infer_headline_from_topcard(topcard: str, *, student_name: str = "") -> str:
    """LinkedIn profile header: name on line 1, headline on line 2."""
    lines = [ln.strip() for ln in topcard.splitlines() if ln.strip() and ln.strip() != "·"]
    name_tokens = [t.lower() for t in student_name.split() if len(t) > 2][:2]
    for idx, line in enumerate(lines):
        if _line_looks_like_location(line):
            continue
        low = line.lower()
        if name_tokens and all(tok in low for tok in name_tokens) and "||" not in line:
            if idx == 0 or len(line.split()) <= 5:
                continue
        if idx == 0 and len(lines) > 1 and "||" not in line and "@" not in line:
            if len(line.split()) <= 5:
                continue
        if 10 <= len(line) <= 240:
            return line
    return ""


def infer_headline_from_about(about: str) -> str:
    for para in about.split("\n\n"):
        for line in para.split("\n"):
            line = line.strip()
            if "||" in line and 15 <= len(line) <= 240:
                return line
    return ""


def linkedin_username_from_url(url: str) -> str:
    path = urlparse(url).path.strip("/")
    if path.startswith("in/"):
        slug = path.split("/", 1)[1].split("/")[0]
        return slug.strip()
    return ""


def profile_slug_matches(link: str, expected_username: str) -> bool:
    """Exact LinkedIn slug match — avoids nayyab-fatima- matching nayyab-fatima-8584651b8."""
    slug = linkedin_username_from_url(link)
    if not slug or not expected_username:
        return False
    a = slug.lower().rstrip("-")
    b = expected_username.lower().rstrip("-")
    return slug.lower() == expected_username.lower() or a == b


def combined_result_text(item: dict) -> str:
    parts = [item.get("title") or "", item.get("snippet") or ""]
    about = item.get("about_this_result") or {}
    source = about.get("source") or {}
    if source.get("description"):
        parts.append(source["description"])
    return " ".join(p for p in parts if p).strip()


def normalize_search_text(text: str) -> str:
    """Strip accents so Islāmābād matches Islamabad in regex."""
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", text)
    return "".join(c for c in text if not unicodedata.combining(c))


def normalize_city_name(location: str) -> str:
    loc = normalize_search_text(location).strip().strip(".")
    if not loc:
        return ""
    loc = re.sub(r"\s+", " ", loc)
    if "," in loc:
        city = city_from_comma_location(loc)
        if city:
            return city
        parts = [p.strip().lower() for p in loc.split(",") if p.strip()]
        if parts and all(p in PAKISTAN_PROVINCES for p in parts):
            return ""
        return ""
    low = loc.lower()
    if "," not in loc and low in {
        "islamabad", "lahore", "karachi", "rawalpindi", "peshawar",
        "multan", "faisalabad", "sialkot", "quetta", "hyderabad",
        "abbottabad", "mardan", "gujranwala", "sargodha", "bahawalpur",
        "jaranwala", "gujrat", "sheikhupura", "okara", "sahiwal",
        "taxila", "mirpur", "murree", "swat", "charsadda", "kohat",
        "sukkur", "larkana", "wah",
    }:
        return loc.title()
    return loc


def city_from_comma_location(location: str) -> str:
    """Lahore, Punjab, Pakistan -> Lahore; Punjab, Pakistan -> empty."""
    parts = [p.strip() for p in location.split(",") if p.strip()]
    if not parts:
        return ""
    for part in parts:
        if part.lower() in PAKISTAN_PROVINCES:
            continue
        if LOCATION_NOISE_RE.search(part):
            continue
        if len(part) > 40 or part.lower().endswith(" university"):
            continue
        city = normalize_city_name(part)
        if city and city.lower() not in PAKISTAN_PROVINCES:
            return city
    return ""


def _location_candidate_score(source: str, value: str) -> int:
    score = 0
    if source == "location_label":
        score += 100
    elif source == "city_province_pakistan":
        score += 80
    elif source == "city_pakistan":
        score += 70
    elif source == "city_standalone":
        score += 40
    val = value.lower()
    if val in PAKISTAN_PROVINCES:
        score -= 50
    if LOCATION_NOISE_RE.search(value):
        score -= 40
    if "university" in val:
        score -= 60
    if "," not in value:
        score += 10
    return score


def infer_location(text: str) -> str:
    text_n = normalize_search_text(text)
    candidates: list[tuple[int, str]] = []

    def add(source: str, raw: str) -> None:
        raw = raw.strip().strip(".")
        if not raw:
            return
        city = city_from_comma_location(raw) if "," in raw else normalize_city_name(raw)
        if not city:
            return
        if LOCATION_NOISE_RE.search(city) and len(city) > 25:
            return
        score = _location_candidate_score(source, city)
        if score > 0:
            candidates.append((score, city))

    m = LOCATION_LABEL_RE.search(text_n)
    if m:
        add("location_label", m.group(1).strip().strip("."))

    for m in CITY_PROVINCE_PAKISTAN_RE.finditer(text_n):
        add("city_province_pakistan", m.group(0))

    for m in re.finditer(
        r"\b([A-Za-z][A-Za-z\s\-]{2,35})\s*,\s*Pakistan\b",
        text_n,
        re.IGNORECASE,
    ):
        add("city_pakistan", m.group(0))

    for m in CITY_STANDALONE_RE.finditer(text_n):
        start = max(0, m.start() - 50)
        prefix = text_n[start : m.start()].lower()
        if "university" in prefix and "location" not in prefix:
            continue
        add("city_standalone", m.group(1))

    if candidates:
        candidates.sort(key=lambda x: (x[0], -len(x[1])), reverse=True)
        return candidates[0][1]

    for match in LOCATION_RE.finditer(text_n):
        loc = match.group(0).strip()
        start = max(0, match.start() - 40)
        prefix = text_n[start : match.start()].lower()
        if "university" in prefix or "education:" in prefix:
            continue
        if "view " in prefix or "profile on linkedin" in text_n[match.start() : match.end() + 30].lower():
            continue
        city = city_from_comma_location(loc) if "," in loc else normalize_city_name(loc)
        if city and city.lower() not in PAKISTAN_PROVINCES:
            return city
    return ""


def university_keywords(university: str) -> tuple[str, ...]:
    u = normalize_search_text(university).lower().strip()
    keywords: list[str] = []
    if u:
        keywords.append(u)
    if "riphah" in u:
        keywords.extend(
            [
                "riphah international university",
                "riphah international",
                "riphah",
            ]
        )
    return tuple(dict.fromkeys(k for k in keywords if k))


def build_profile_enrich_queries(username: str, name: str, university: str) -> list[str]:
    """Queries scoped to the exact profile slug (no broad name-only searches)."""
    uni = university.strip() or DEFAULT_UNIVERSITY
    queries = [
        f"site:linkedin.com/in/{username}",
        f'site:linkedin.com/in/{username} "connections on LinkedIn"',
        f'site:pk.linkedin.com/in/{username} "connections on LinkedIn"',
        f"site:linkedin.com/in/{username} Location",
        f"site:linkedin.com/in/{username} Education",
        f'site:linkedin.com/in/{username} Education Riphah',
        f'site:linkedin.com/in/{username} "Bachelor" Riphah Present',
        f'site:linkedin.com/in/{username} Experience Education Riphah',
        f'site:linkedin.com/in/{username} "Experience"',
        f'site:linkedin.com/in/{username} intern OR developer OR engineer',
        f'site:linkedin.com/in/{username} "Student at" Riphah',
        f'site:linkedin.com/in/{username} "Present" OR "2024" OR "2025"',
        f'"{username}" linkedin Riphah "Bachelor" "Present"',
        f'"{username}" linkedin Riphah "2023" "2027"',
    ]
    profile_url = f"https://www.linkedin.com/in/{username}"
    queries.append(f'"{profile_url}" Riphah education')
    if name.strip():
        n = name.strip()
        queries.append(f'"{n}" site:linkedin.com/in/{username}')
        queries.append(f'"{n}" Riphah site:linkedin.com/in/{username}')
        queries.append(f'"{n}" {uni} BSCS site:linkedin.com/in/{username}')
        queries.append(f'"{n}" linkedin experience Pakistan')
    return list(dict.fromkeys(queries))


EXPERIENCE_BLOCK_RE = re.compile(
    r"Experience[:\s]+(.{15,1200}?)(?=\s*Education[:\s]|\s*Licenses|\s*Certifications|\s*Skills|\s*Activity|$)",
    re.IGNORECASE | re.DOTALL,
)
EDUCATION_BLOCK_RE = re.compile(
    r"Education[:\s]+(.{15,900}?)(?=\s*Experience[:\s]|\s*Licenses|\s*Certifications|\s*Skills|\s*Activity|$)",
    re.IGNORECASE | re.DOTALL,
)
ROLE_AT_RE = re.compile(
    r"([A-Za-z0-9][^·\n|]{4,70}?)\s+at\s+([A-Za-z0-9][^·\n|]{3,70})",
    re.IGNORECASE,
)


def _normalize_blob(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def infer_experience_summary(text: str) -> str:
    text_n = text.replace("\n", " ")
    m = EXPERIENCE_BLOCK_RE.search(text_n)
    if m:
        return _normalize_blob(m.group(1))[:1500]
    parts: list[str] = []
    for m in ROLE_AT_RE.finditer(text_n):
        parts.append(f"{m.group(1).strip()} at {m.group(2).strip()}")
    if parts:
        return "; ".join(dict.fromkeys(parts))[:1500]
    return ""


def infer_experience_roles(text: str) -> str:
    text_n = text.replace("\n", " ")
    roles: list[str] = []
    block = EXPERIENCE_BLOCK_RE.search(text_n)
    chunk = block.group(1) if block else text_n
    for m in ROLE_AT_RE.finditer(chunk):
        roles.append(f"{m.group(1).strip()} @ {m.group(2).strip()}")
    return "; ".join(dict.fromkeys(roles))[:1000]


def infer_education_detail(text: str, university: str) -> str:
    text_n = text.replace("\n", " ")
    m = EDUCATION_BLOCK_RE.search(text_n)
    if m:
        return _normalize_blob(m.group(1))[:1000]
    uni = university or DEFAULT_UNIVERSITY
    key = uni.split()[0] if uni else "Riphah"
    m2 = re.search(
        rf"([^.;|]{{0,80}}{re.escape(key)}[^.;|]{{0,200}}?\.\s*"
        r"[^.;|]{2,120}?\.\s*"
        r"(?:19|20)\d{{2}}\s*[-–—]\s*(?:(?:19|20)\d{{2}}|Present))",
        text_n,
        re.IGNORECASE,
    )
    return _normalize_blob(m2.group(0)) if m2 else ""


def infer_current_job(text: str, headline: str = "", experience_roles: str = "") -> str:
    if headline and "linkedin" not in headline.lower() and len(headline) > 8:
        return headline[:220]
    if experience_roles:
        return experience_roles.split(";")[0].strip()[:220]
    text_n = text.replace("\n", " ")
    m = re.search(
        r"Experience[:\s]+([^.·|]{8,100}?)\s+at\s+([^.·|]{3,80})",
        text_n,
        re.IGNORECASE,
    )
    if m:
        return f"{m.group(1).strip()} at {m.group(2).strip()}"[:220]
    return ""


def apply_extended_fields(
    row: dict[str, str],
    text: str,
    *,
    university: str,
    overwrite: bool = False,
) -> None:
    mapping = {
        "experience_summary": infer_experience_summary(text),
        "experience_roles": infer_experience_roles(text),
        "education_detail": infer_education_detail(text, university),
    }
    mapping["current_job"] = infer_current_job(
        text,
        str(row.get("headline", "")),
        mapping["experience_roles"],
    )
    for key, value in mapping.items():
        if not value:
            continue
        if overwrite or not str(row.get(key, "")).strip():
            row[key] = value


def collect_profile_serpapi_text(
    session: requests.Session,
    api_key: str,
    username: str,
    queries: list[str],
    *,
    gl: str,
    sleep_seconds: float,
) -> str:
    chunks: list[str] = []
    seen_snippets: set[str] = set()
    for query in queries:
        try:
            data = serpapi_google_search(session, api_key, query, num=5, gl=gl)
        except Exception:
            time.sleep(sleep_seconds)
            continue
        for item in data.get("organic_results") or []:
            link = normalize_linkedin_profile_url(item.get("link") or "")
            if not profile_slug_matches(link, username):
                continue
            blob = combined_result_text(item)
            if blob and blob not in seen_snippets:
                seen_snippets.add(blob)
                chunks.append(blob)
        time.sleep(sleep_seconds)
    return " ".join(chunks).strip()


def _connection_count_value(raw: str) -> int:
    """Parse 500+, 46, 1,234 for ranking (higher = more connections)."""
    raw = raw.replace(",", "").strip().rstrip("+")
    if not raw:
        return 0
    try:
        return int(raw)
    except ValueError:
        return 0


def infer_connections(text: str) -> str:
    text_n = normalize_search_text(text)
    candidates: list[tuple[int, str, bool]] = []

    def add(raw: str, has_plus: bool = False) -> None:
        raw = raw.strip()
        if not raw:
            return
        display = f"{raw} connections" if not raw.endswith("+") else f"{raw} connections"
        if has_plus and "+" not in raw:
            display = f"{raw}+ connections"
        candidates.append((_connection_count_value(raw), display, has_plus or "+" in raw))

    for m in FOLLOWERS_CONNECTIONS_RE.finditer(text_n):
        add(m.group(2))
    for m in CONNECTIONS_ON_LINKEDIN_RE.finditer(text_n):
        add(m.group(1), "+" in m.group(1))
    for m in CONNECTIONS_RE.finditer(text_n):
        add(m.group(1), "+" in m.group(1))
    for m in LOCATION_CONNECTIONS_RE.finditer(text_n):
        add(m.group(1), "+" in m.group(1))
    for m in re.finditer(r"(\d[\d,]*\+?)\s*\.\.\.", text_n):
        add(m.group(1), "+" in m.group(1))
    # Arabic LinkedIn: more than 500 connections
    if re.search(r"500|٥٠٠|أكثر من\s*٥٠٠", text_n) and re.search(
        r"زميل|connection|linkedin", text_n, re.I
    ):
        candidates.append((500, "500+ connections", True))

    if not candidates:
        return ""
    # Prefer 500+ tier, then highest numeric count
    candidates.sort(key=lambda x: (x[2], x[0]), reverse=True)
    return candidates[0][1]


def extract_education_entries(text: str) -> list[tuple[str, str, str]]:
    """Parse 'Institution. Degree line. 2023 - 2027' tuples from Education snippets."""
    entries: list[tuple[str, str, str]] = []
    for m in EDUCATION_ENTRY_RE.finditer(text):
        inst = m.group(1).strip()
        degree = m.group(2).strip()
        years = m.group(3).strip()
        if len(inst) < 6 or len(degree) < 2:
            continue
        entries.append((inst, degree, years))
    return entries


def _institution_matches_university(institution: str, keywords: tuple[str, ...]) -> bool:
    inst_l = institution.lower()
    if any(kw in inst_l for kw in keywords):
        return True
    if "riphah" in inst_l and any("riphah" in kw for kw in keywords):
        return True
    return False


def _study_year_score(context: str, value: str, keywords: tuple[str, ...]) -> int:
    """Score a year range by proximity to target university + education vs job experience."""
    ctx = context.lower()
    score = 0
    if any(kw in ctx for kw in keywords):
        score += 60
    if re.search(r"\beducation\b", ctx):
        score += 35
    if re.search(r"\b(bachelor|bscs|bs cs|bsse|degree|diploma|undergraduate)\b", ctx):
        score += 25
    if re.search(r"\bstudent at\b", ctx) and any(kw in ctx for kw in keywords):
        score += 20
    if re.search(r"\b(present|current)\b", value, re.I):
        score += 8
    if re.search(r"\bexperience\b", ctx) and not re.search(r"\beducation\b", ctx):
        score -= 25
    if OTHER_UNIVERSITY_RE.search(ctx) and not any(kw in ctx for kw in keywords):
        score -= 50
    if re.search(r"\b(licenses|certifications|certified)\b", ctx):
        score -= 20
    if re.search(r"\b(internship|intern at|experience:)\b", ctx) and "education" not in ctx:
        score -= 15
    return score


def infer_study_year_at_university(text: str, university: str) -> str:
    """
    Study period at the target university from Education / Experience snippets.
    Prefers year ranges beside the university name (e.g. Riphah ... 2023 - 2027).
    """
    text_n = normalize_search_text(text)
    keywords = university_keywords(university)
    if not keywords:
        return ""

    candidates: list[tuple[int, str]] = []
    institution_bound: list[tuple[int, str]] = []

    # One row per school: only year ranges tied to the target university name
    for inst, degree, years in extract_education_entries(text_n):
        if not _institution_matches_university(inst, keywords):
            continue
        score = 100
        if re.search(r"\b(bscs|bachelor|bs cs|computer science|software engineering)\b", degree, re.I):
            score += 12
        if re.search(r"\b(associate|diploma)\b", degree, re.I):
            score += 4
        if re.search(r"present|current", years, re.I):
            score += 3
        institution_bound.append((score, years))

    if institution_bound:
        institution_bound.sort(key=lambda x: (x[0], len(x[1])), reverse=True)
        return institution_bound[0][1]

    for match in YEAR_RANGE_FULL_RE.finditer(text_n):
        val = match.group(1).strip()
        start = max(0, match.start() - 120)
        end = min(len(text_n), match.end() + 60)
        context = text_n[start:end]
        score = _study_year_score(context, val, keywords)
        # Require university name in the same window (avoids NCBA/other-school dates)
        if score > 0 and any(kw in context.lower() for kw in keywords):
            candidates.append((score, val))

    # "Riphah International University ... 2023 - Present" (range after uni name)
    for kw in keywords:
        pattern = re.compile(
            re.escape(kw) + r".{0,180}?" + YEAR_RANGE_FULL_RE.pattern,
            re.IGNORECASE | re.DOTALL,
        )
        for m in pattern.finditer(text_n):
            val = m.group(1).strip()
            candidates.append((85, val))

    # Education section block before next major section
    for m in re.finditer(
        r"Education[:\s].{0,400}?(?=\s+Experience[:\s]|\s+Licenses|\s+Certifications|$)",
        text_n,
        re.IGNORECASE | re.DOTALL,
    ):
        block = m.group(0)
        if not any(kw in block.lower() for kw in keywords):
            continue
        for yr in YEAR_RANGE_FULL_RE.finditer(block):
            val = yr.group(1).strip()
            candidates.append((_study_year_score(block, val, keywords) + 15, val))

    # "Student at Riphah ..." in experience (dates often precede the role in snippets)
    for kw in keywords:
        for m in re.finditer(
            rf"Student\s+at\s+{re.escape(kw)}",
            text_n,
            re.IGNORECASE,
        ):
            start = max(0, m.start() - 90)
            end = min(len(text_n), m.end() + 60)
            block = text_n[start:end]
            yr = YEAR_RANGE_FULL_RE.search(block)
            if yr:
                val = yr.group(1).strip()
                candidates.append((_study_year_score(block, val, keywords) + 10, val))

    if candidates:
        candidates.sort(key=lambda x: (x[0], len(x[1])), reverse=True)
        return candidates[0][1]

    # Single start year only near university (weak fallback)
    for kw in keywords:
        for m in re.finditer(re.escape(kw), text_n, re.IGNORECASE):
            ctx = text_n[m.end() : m.end() + 80]
            yr = re.search(r"\b((?:19|20)\d{2})\b", ctx)
            if yr and re.search(r"student|bachelor|bscs|degree", ctx, re.I):
                return yr.group(1)

    return ""


def infer_study_year_from_semester(text: str, university: str) -> str:
    """
    Estimate Riphah study period from '5th semester … Riphah' when Google has no dates.
    Marked as approximate in study_year_source.
    """
    text_n = normalize_search_text(text)
    keywords = university_keywords(university)
    if not keywords:
        return ""

    now = datetime.now()
    academic_start_year = now.year if now.month >= 8 else now.year - 1

    for match in SEMESTER_RE.finditer(text_n):
        sem = int(match.group(1))
        if sem < 1 or sem > 10:
            continue
        start = max(0, match.start() - 160)
        end = min(len(text_n), match.end() + 160)
        ctx = text_n[start:end]
        ctx_l = ctx.lower()
        if not any(kw in ctx_l for kw in keywords):
            continue
        if not re.search(r"\b(student|bscs|bachelor|undergraduate|studying)\b", ctx_l):
            continue
        years_in_program = (sem - 1) // 2
        start_year = academic_start_year - years_in_program
        end_year = start_year + BSCS_DURATION_YEARS
        if sem >= 7 or re.search(r"\b(?:final|graduat)\b", ctx_l):
            return f"{start_year} - {end_year}"
        return f"{start_year} - Present"
    return ""


def has_riphah_education_evidence(text: str, university: str) -> bool:
    """True when snippet shows Riphah in Education (not only as employer)."""
    text_n = normalize_search_text(text)
    keywords = university_keywords(university)
    if not keywords:
        return False

    for inst, _, _ in extract_education_entries(text_n):
        if _institution_matches_university(inst, keywords):
            return True

    if re.search(r"education[:;\s].{0,200}riphah", text_n, re.IGNORECASE):
        return True
    m = re.search(
        r"riphah.{0,120}(?:bachelor|bscs|associate|computer science)",
        text_n,
        re.IGNORECASE,
    )
    if m and not OTHER_UNIVERSITY_RE.search(m.group(0)):
        return True

    if re.search(r"student\s+at\s+riphah", text_n, re.IGNORECASE):
        return True

    # Experience-only at Riphah (staff/alumni at other schools) — not education evidence
    if re.search(r"experience:.{0,80}riphah", text_n, re.IGNORECASE):
        if re.search(r"education:.{0,80}(?:pucit|comsats|nust|fast|uet|ksbl)", text_n, re.IGNORECASE):
            return False
    return False


def classify_profile_status(text: str, university: str, *, students_only: bool) -> str:
    """verified_student | needs_review | not_riphah_student"""
    text_n = normalize_search_text(text)
    if students_only and is_likely_faculty(text_n):
        return "not_riphah_student"
    if not has_riphah_education_evidence(text_n, university):
        if re.search(r"\briphah\b", text_n, re.IGNORECASE) and is_likely_student(text_n):
            return "needs_review"
        return "not_riphah_student"
    if not infer_study_year_at_university(text_n, university) and not infer_study_year_from_semester(
        text_n, university
    ):
        return "needs_review"
    return "verified_student"


def compute_data_confidence(row: dict[str, str]) -> str:
    score = 0
    if row.get("location"):
        score += 1
    if row.get("connections"):
        score += 1
    if row.get("study_year"):
        score += 2
    if row.get("study_year_source") == "education":
        score += 2
    elif row.get("study_year_source") == "semester_estimate":
        score += 1
    if row.get("degree_program"):
        score += 1
    if row.get("profile_status") == "verified_student":
        score += 2
    elif row.get("profile_status") == "needs_review":
        score += 1
    if score >= 7:
        return "high"
    if score >= 4:
        return "medium"
    return "low"


def profile_combined_text(row: dict[str, str]) -> str:
    return " ".join(
        str(row.get(k, "") or "")
        for k in ("result_title", "snippet", "headline", "about")
    ).strip()


def finalize_row_quality(
    row: dict[str, str],
    *,
    university: str,
    students_only: bool,
) -> None:
    text = profile_combined_text(row)
    row["profile_status"] = classify_profile_status(text, university, students_only=students_only)
    row["data_confidence"] = compute_data_confidence(row)


def infer_degree_at_university(text: str, university: str) -> str:
    text_n = normalize_search_text(text)
    keywords = university_keywords(university)
    for kw in keywords:
        for m in re.finditer(re.escape(kw), text_n, re.IGNORECASE):
            ctx = text_n[max(0, m.start() - 30) : m.end() + 140]
            d = DEGREE_RE.search(ctx)
            if d:
                return d.group(0).strip()
    return infer_degree_program(text_n)


def apply_parsed_fields(
    row: dict[str, str],
    text: str,
    *,
    university: str,
    overwrite: bool = False,
    topcard: str = "",
) -> None:
    """Fill row fields from combined Google/LinkedIn snippet text."""
    uni = university or row.get("university") or DEFAULT_UNIVERSITY
    study_year = infer_study_year_at_university(text, uni)
    if not study_year:
        study_year = infer_study_year_from_semester(text, uni)
    mapping = {
        "location": infer_location(text),
        "study_year": study_year,
        "connections": infer_connections(text),
        "degree_program": infer_degree_at_university(text, uni),
        "subjects": infer_subjects(text),
    }
    email, phone = contacts_from_text(text)
    if email:
        mapping["email"] = email
    if phone:
        mapping["phone"] = phone

    for key, value in mapping.items():
        if not value:
            continue
        if overwrite or not str(row.get(key, "")).strip():
            row[key] = value
    headline = parse_headline(row.get("result_title", ""), text)
    if not headline and topcard:
        headline = infer_headline_from_topcard(
            topcard, student_name=row.get("student_name", "")
        )
    if headline and (overwrite or not row.get("headline")):
        row["headline"] = headline
    about = infer_about(text, row.get("headline", ""))
    if about and (overwrite or not row.get("about")):
        row["about"] = about
    if not row.get("headline"):
        fallback = infer_headline_from_about(row.get("about", ""))
        if fallback:
            row["headline"] = fallback
    apply_extended_fields(row, text, university=uni, overwrite=overwrite)


def infer_degree_program(text: str) -> str:
    m = DEGREE_RE.search(text)
    return m.group(0).strip() if m else ""


def infer_subjects(text: str) -> str:
    lower = text.lower()
    found = [hint.title() if hint.islower() else hint for hint in SUBJECT_HINTS if hint in lower]
    degree = infer_degree_program(text)
    if degree and degree.lower() not in {s.lower() for s in found}:
        found.insert(0, degree)
    return "; ".join(dict.fromkeys(found))


def infer_about(snippet: str, headline: str) -> str:
    about = snippet.strip()
    if headline and about.lower().startswith(headline.lower()):
        about = about[len(headline) :].lstrip(" .,-")
    return about[:500]


def is_likely_alumni(text: str, current_year: int | None = None) -> bool:
    if re.search(r"\b(?:present|currently studying|current student)\b", text, re.IGNORECASE):
        return False
    year = current_year or datetime.now().year
    education_ctx = re.compile(
        r"\b(education|bachelor|bs\.?|bscs|degree|university|diploma|student at)\b",
        re.IGNORECASE,
    )
    for match in ALUMNI_RANGE_RE.finditer(text):
        start = max(0, match.start() - 50)
        end = min(len(text), match.end() + 50)
        context = text[start:end]
        if not education_ctx.search(context):
            continue
        end_raw = match.group(1)
        end_year = int(end_raw) if len(end_raw) == 4 else 2000 + int(end_raw[-2:])
        if end_year < year - 1:
            return True
    return False


def is_likely_faculty(text: str) -> bool:
    if STAFF_SIGNAL_RE.search(text):
        return True
    return bool(FACULTY_RE.search(text))


def is_likely_student(text: str) -> bool:
    if is_likely_faculty(text):
        return False
    if is_likely_alumni(text):
        return False
    return bool(STUDENT_RE.search(text))


def is_cs_related(text: str) -> bool:
    return bool(CS_RE.search(text))


def build_search_queries(
    university: str,
    department: str,
) -> list[str]:
    """Queries tuned for SerpAPI/Google — all Riphah campuses."""
    dept = department.strip()
    return [
        f"site:linkedin.com/in Riphah {dept} student Pakistan",
        f'site:linkedin.com/in "Riphah International University" student BSCS',
        f"site:linkedin.com/in Riphah BSCS student Pakistan",
        f'site:linkedin.com/in "Riphah International University" student',
        f'site:linkedin.com/in Riphah "BS Computer Science" student',
        f'site:linkedin.com/in "Riphah International University" "Computer Science" student',
        f"site:linkedin.com/in Riphah CS student Pakistan",
        f"site:linkedin.com/in Riphah BSCS Lahore",
        f"site:linkedin.com/in Riphah BSCS Islamabad",
    ]


def serpapi_google_search(
    session: requests.Session,
    api_key: str,
    query: str,
    *,
    num: int = 10,
    start: int = 0,
    gl: str = "pk",
    hl: str = "en",
) -> dict:
    params = {
        "engine": "google",
        "q": query,
        "api_key": api_key,
        "num": min(max(num, 1), 100),
        "start": max(start, 0),
        "gl": gl,
        "hl": hl,
    }
    resp = session.get(SERPAPI_URL, params=params, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    if data.get("error"):
        raise RuntimeError(str(data["error"]))
    return data


def profiles_from_serpapi_results(
    data: dict,
    *,
    query: str,
    university: str,
    campus: str,
    department: str,
    seen_urls: set[str],
    max_new: int,
    students_only: bool = True,
    cs_only: bool = True,
) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    organic = data.get("organic_results") or []
    for item in organic:
        if len(rows) >= max_new:
            break
        raw_url = (item.get("link") or "").strip()
        profile_url = normalize_linkedin_profile_url(raw_url)
        if not profile_url or not is_linkedin_profile(profile_url):
            continue
        if profile_url in seen_urls:
            continue

        title = (item.get("title") or "").strip()
        snippet = (item.get("snippet") or "").strip()
        combined = f"{title} {snippet}"

        if students_only and not is_likely_student(combined):
            continue
        if cs_only and not is_cs_related(combined):
            continue

        seen_urls.add(profile_url)
        headline = parse_headline(title, snippet)
        email, phone = contacts_from_text(combined)
        student_name = name_from_title(title) or name_from_linkedin_url(profile_url)

        row = {
            "university": university,
            "campus": campus,
            "department": department,
            "student_name": student_name,
            "linkedin_url": profile_url,
            "headline": headline,
            "location": "",
            "study_year": "",
            "connections": "",
            "degree_program": "",
            "subjects": "",
            "about": "",
            "email": email,
            "phone": phone,
            "result_title": title,
            "snippet": snippet,
            "search_query": query,
            "profile_rank": "",
            "data_source": "serpapi_google",
            "scraped_at": datetime.now().isoformat(timespec="seconds"),
        }
        apply_parsed_fields(row, combined, university=university)
        rows.append(row)
    return rows


def enrich_profile_via_serpapi(
    session: requests.Session,
    api_key: str,
    row: dict[str, str],
    *,
    university: str,
    gl: str,
    sleep_seconds: float,
) -> None:
    """Run several Google queries per profile to capture location, connections, and years."""
    username = linkedin_username_from_url(row.get("linkedin_url", ""))
    if not username:
        return
    name = row.get("student_name", "").strip()
    queries = build_profile_enrich_queries(username, name, university)
    profile_text = collect_profile_serpapi_text(
        session, api_key, username, queries, gl=gl, sleep_seconds=sleep_seconds
    )
    merged_text = " ".join(
        p
        for p in (
            row.get("snippet", ""),
            row.get("result_title", ""),
            row.get("about", ""),
            profile_text,
        )
        if p
    ).strip()
    if not merged_text:
        return
    if profile_text:
        row["snippet"] = profile_text[:2500]
    apply_parsed_fields(row, merged_text, university=university, overwrite=True)
    row["data_source"] = "serpapi_google+profile_lookup"
    row["scraped_at"] = datetime.now().isoformat(timespec="seconds")


def enrich_contact_via_serpapi(
    session: requests.Session,
    api_key: str,
    row: dict[str, str],
    *,
    sleep_seconds: float,
) -> None:
    name = row.get("student_name", "").strip()
    if not name:
        return
    if row.get("email") and row.get("phone"):
        return

    queries = [
        f'"{name}" "Riphah" email OR phone OR contact',
        f'"{name}" "Computer Science" Islamabad email OR "@gmail.com" OR "@outlook.com"',
    ]
    for query in queries:
        if row.get("email") and row.get("phone"):
            break
        try:
            data = serpapi_google_search(session, api_key, query, num=5)
        except Exception:
            continue
        combined_parts: list[str] = []
        for item in data.get("organic_results") or []:
            combined_parts.append(item.get("title") or "")
            combined_parts.append(item.get("snippet") or "")
        combined = " ".join(combined_parts)
        email, phone = contacts_from_text(combined)
        if email and not row.get("email"):
            row["email"] = email
        if phone and not row.get("phone"):
            row["phone"] = phone
        time.sleep(sleep_seconds)


def profile_row_key(row: dict[str, str]) -> str:
    return normalize_linkedin_profile_url(row.get("linkedin_url", "")).lower()


def load_table_file(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path, dtype=str)
    else:
        df = pd.read_excel(path, dtype=str)
    df = df.fillna("")
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df


def load_existing_outputs(csv_path: Path, xlsx_path: Path) -> list[dict[str, str]]:
    """Load previously saved rows from Excel (preferred) or CSV."""
    if xlsx_path.is_file():
        df = load_table_file(xlsx_path)
    elif csv_path.is_file():
        df = load_table_file(csv_path)
    else:
        return []
    rows = df[OUTPUT_COLUMNS].to_dict("records")
    return [dict(r) for r in rows]


def load_profiles_from_input(
    path: Path,
    *,
    university: str,
    campus: str,
    department: str,
) -> list[dict[str, str]]:
    """Import rows from Playwright or other CSV/XLSX (matched by linkedin_url)."""
    df = load_table_file(path)
    rows: list[dict[str, str]] = []
    for rec in df.to_dict("records"):
        url = normalize_linkedin_profile_url(str(rec.get("linkedin_url", "")))
        if not url:
            continue
        row = {col: "" for col in OUTPUT_COLUMNS}
        row.update(
            {
                "university": university,
                "campus": campus,
                "department": department,
                "linkedin_url": url,
                "student_name": str(rec.get("student_name", "")).strip()
                or name_from_linkedin_url(url),
                "data_source": "imported",
            }
        )
        for col in OUTPUT_COLUMNS:
            if col in rec and str(rec[col]).strip():
                row[col] = str(rec[col]).strip()
        rows.append(row)
    return rows


def assign_profile_ranks(rows: list[dict[str, str]]) -> None:
    for rank, row in enumerate(rows, start=1):
        row["profile_rank"] = str(rank)


def save_outputs(rows: list[dict[str, str]], csv_path: Path, xlsx_path: Path) -> None:
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    df.to_excel(xlsx_path, index=False, sheet_name="students")


def field_richness(value: str) -> int:
    return len(str(value or "").strip())


def merge_profile_row(existing: dict[str, str], new: dict[str, str]) -> dict[str, str]:
    """Keep the richer value for each field when updating an existing row."""
    merged = dict(existing)
    for key in OUTPUT_COLUMNS:
        old_v = str(merged.get(key, "")).strip()
        new_v = str(new.get(key, "")).strip()
        if not new_v:
            continue
        if not old_v or field_richness(new_v) > field_richness(old_v):
            merged[key] = new_v
    merged["scraped_at"] = new.get("scraped_at") or merged.get("scraped_at", "")
    if new.get("data_source"):
        merged["data_source"] = new["data_source"]
    return merged


def upsert_and_save(
    stored_rows: list[dict[str, str]],
    new_row: dict[str, str],
    csv_path: Path,
    xlsx_path: Path,
) -> tuple[list[dict[str, str]], bool]:
    """Insert or update one row (by LinkedIn URL) and write Excel + CSV."""
    key = profile_row_key(new_row)
    if not key:
        return stored_rows, False
    idx = next((i for i, r in enumerate(stored_rows) if profile_row_key(r) == key), None)
    added = idx is None
    if idx is not None:
        stored_rows[idx] = merge_profile_row(stored_rows[idx], new_row)
    else:
        stored_rows.append(new_row)
    assign_profile_ranks(stored_rows)
    save_outputs(stored_rows, csv_path, xlsx_path)
    return stored_rows, added


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Scrape Riphah CS student LinkedIn profiles via SerpAPI (Google Search)."
    )
    parser.add_argument("--university", default=DEFAULT_UNIVERSITY, help="University name")
    parser.add_argument(
        "--campus",
        default=DEFAULT_CAMPUS,
        help="Campus label stored in output (not used for filtering)",
    )
    parser.add_argument("--department", default=DEFAULT_DEPARTMENT, help="Department")
    parser.add_argument(
        "--max-profiles",
        type=int,
        default=10,
        help="How many new profiles to add this run (skips URLs already in the output file)",
    )
    parser.add_argument(
        "--api-key",
        default=os.environ.get("SERPAPI_KEY", "").strip(),
        help="SerpAPI key (default: SERPAPI_KEY env var or .env file)",
    )
    parser.add_argument(
        "--output-csv",
        type=Path,
        default=DEFAULT_OUTPUT_CSV,
        help="CSV output path",
    )
    parser.add_argument(
        "--output-xlsx",
        type=Path,
        default=DEFAULT_OUTPUT_XLSX,
        help="Excel output path",
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=None,
        help="CSV/XLSX with linkedin_url (e.g. riphah_cs_students_100.csv) to enrich via SerpAPI",
    )
    parser.add_argument(
        "--refresh-existing",
        action="store_true",
        help="Re-enrich every profile already in the output file (fixes missing location/connections)",
    )
    parser.add_argument(
        "--fresh",
        action="store_true",
        help="Overwrite output files instead of appending to existing Excel/CSV",
    )
    parser.add_argument(
        "--skip-profile-enrich",
        action="store_true",
        help="Skip per-profile SerpAPI lookup (location/connections/year will be less accurate)",
    )
    parser.add_argument(
        "--enrich-contacts",
        action="store_true",
        help="Run extra SerpAPI searches per student to find public email/phone",
    )
    parser.add_argument(
        "--include-faculty",
        action="store_true",
        help="Do not filter out professor/lecturer profiles",
    )
    parser.add_argument(
        "--sleep",
        type=float,
        default=1.0,
        help="Seconds to wait between SerpAPI requests (default: 1.0)",
    )
    parser.add_argument(
        "--gl",
        default="pk",
        help="Google country code for SerpAPI (default: pk)",
    )
    return parser.parse_args()


def run_search_pass(
    session: requests.Session,
    api_key: str,
    queries: list[str],
    *,
    university: str,
    campus: str,
    department: str,
    seen_urls: set[str],
    all_rows: list[dict[str, str]],
    max_profiles: int,
    students_only: bool,
    cs_only: bool,
    gl: str,
    sleep_seconds: float,
    page_starts: tuple[int, ...] = (0, 10),
    label: str = "",
) -> int:
    added = 0
    if label:
        print(f"\n{label}")
    for qi, query in enumerate(queries, start=1):
        if len(all_rows) >= max_profiles:
            break
        for start in page_starts:
            if len(all_rows) >= max_profiles:
                break
            remaining = max_profiles - len(all_rows)
            try:
                data = serpapi_google_search(
                    session,
                    api_key,
                    query,
                    num=10,
                    start=start,
                    gl=gl,
                )
            except Exception as exc:
                if start == page_starts[0]:
                    safe_print(f"  Query {qi} failed: {exc}")
                break

            new_rows = profiles_from_serpapi_results(
                data,
                query=query,
                university=university,
                campus=campus,
                department=department,
                seen_urls=seen_urls,
                max_new=remaining,
                students_only=students_only,
                cs_only=cs_only,
            )
            if new_rows:
                page_note = f" (page {start // 10 + 1})" if start else ""
                print(f"  Query {qi}{page_note}: +{len(new_rows)} profile(s)")
            added += len(new_rows)
            all_rows.extend(new_rows)
            time.sleep(sleep_seconds)
            if not (data.get("organic_results") or []):
                break
    return added


def safe_print(text: str) -> None:
    try:
        print(text)
    except UnicodeEncodeError:
        print(text.encode("ascii", errors="replace").decode("ascii"))


def main() -> int:
    load_env_file(PROJECT_ROOT / ".env")
    args = parse_args()
    if not args.api_key:
        print(
            "Error: SerpAPI key required.\n"
            "  Put SERPAPI_KEY=... in .env, set env var, or pass --api-key YOUR_KEY\n"
            "  Get a free trial key at https://serpapi.com/users/sign_up",
            file=sys.stderr,
        )
        return 1

    max_profiles = max(1, args.max_profiles)
    queries = build_search_queries(args.university, args.department)
    session = requests.Session()

    if args.fresh:
        stored_rows: list[dict[str, str]] = []
    else:
        stored_rows = load_existing_outputs(args.output_csv, args.output_xlsx)

    seen_urls: set[str] = {profile_row_key(r) for r in stored_rows if profile_row_key(r)}
    new_rows: list[dict[str, str]] = []
    students_only = not args.include_faculty

    print(f"University : {args.university}")
    print(f"Campus     : {args.campus} (label only, no filter)")
    print(f"Department : {args.department}")
    print(f"Target     : {max_profiles} new profile(s) this run")
    print(f"Already in file : {len(stored_rows)} profile(s)")
    print(f"Queries    : {len(queries)}")
    print(f"Output CSV : {args.output_csv}")
    print(f"Output XLSX: {args.output_xlsx}")
    print(f"Mode       : {'overwrite' if args.fresh else 'incremental append'}")
    print()

    if args.refresh_existing and stored_rows:
        print(f"Refreshing {len(stored_rows)} existing profile(s)...")
        for i, row in enumerate(stored_rows, start=1):
            safe_print(f"  [{i}/{len(stored_rows)}] {row.get('student_name') or row.get('linkedin_url')}")
            if not args.skip_profile_enrich:
                enrich_profile_via_serpapi(
                    session,
                    args.api_key,
                    row,
                    university=args.university,
                    gl=args.gl,
                    sleep_seconds=args.sleep,
                )
            if args.enrich_contacts:
                enrich_contact_via_serpapi(
                    session, args.api_key, row, sleep_seconds=args.sleep
                )
            stored_rows, _ = upsert_and_save(stored_rows, row, args.output_csv, args.output_xlsx)
            safe_print(
                f"      -> {row.get('location')} | {row.get('study_year')} | "
                f"{row.get('connections')} | job={str(row.get('current_job', ''))[:40]}"
            )
        print(f"\nRefreshed {len(stored_rows)} profile(s) in file.")
        return 0

    if args.input:
        in_path = args.input
        if not in_path.is_file():
            alt = in_path.with_suffix(".csv" if in_path.suffix.lower() == ".xlsx" else ".xlsx")
            in_path = alt if alt.is_file() else in_path
        if not in_path.is_file():
            print(f"Input not found: {args.input}", file=sys.stderr)
            return 1
        imported = load_profiles_from_input(
            in_path,
            university=args.university,
            campus=args.campus,
            department=args.department,
        )
        if not imported:
            print("No LinkedIn URLs found in input file.", file=sys.stderr)
            return 1
        if args.fresh:
            stored_rows = imported
        else:
            existing = load_existing_outputs(args.output_csv, args.output_xlsx)
            by_url = {profile_row_key(r): r for r in existing if profile_row_key(r)}
            for row in imported:
                key = profile_row_key(row)
                if key in by_url:
                    by_url[key] = merge_profile_row(by_url[key], row)
                else:
                    by_url[key] = row
            stored_rows = list(by_url.values())
        limit = len(stored_rows)
        if args.max_profiles > 0:
            limit = min(args.max_profiles, len(stored_rows))
        to_process = stored_rows[:limit]
        print(
            f"SerpAPI enrich: {len(to_process)} profile(s) from {in_path.name} "
            f"(~{len(build_profile_enrich_queries('user', 'Name', args.university))} searches each)"
        )
        for i, row in enumerate(to_process, start=1):
            safe_print(f"  [{i}/{len(to_process)}] {row.get('student_name') or row.get('linkedin_url')}")
            if not args.skip_profile_enrich:
                enrich_profile_via_serpapi(
                    session,
                    args.api_key,
                    row,
                    university=args.university,
                    gl=args.gl,
                    sleep_seconds=args.sleep,
                )
            if args.enrich_contacts:
                enrich_contact_via_serpapi(
                    session, args.api_key, row, sleep_seconds=args.sleep
                )
            stored_rows, _ = upsert_and_save(
                stored_rows, row, args.output_csv, args.output_xlsx
            )
            safe_print(
                f"      -> yr={row.get('study_year')} | {row.get('current_job', '')[:50]} | "
                f"exp={bool(row.get('experience_summary'))}"
            )
        assign_profile_ranks(stored_rows)
        save_outputs(stored_rows, args.output_csv, args.output_xlsx)
        print(f"\nDone: {len(stored_rows)} profile(s) -> {args.output_xlsx}")
        return 0

    run_search_pass(
        session,
        args.api_key,
        queries,
        university=args.university,
        campus=args.campus,
        department=args.department,
        seen_urls=seen_urls,
        all_rows=new_rows,
        max_profiles=max_profiles,
        students_only=students_only,
        cs_only=True,
        gl=args.gl,
        sleep_seconds=args.sleep,
        page_starts=(0, 10),
        label="Pass 1: CS students (all campuses)",
    )

    if len(new_rows) < max_profiles:
        run_search_pass(
            session,
            args.api_key,
            queries,
            university=args.university,
            campus=args.campus,
            department=args.department,
            seen_urls=seen_urls,
            all_rows=new_rows,
            max_profiles=max_profiles,
            students_only=students_only,
            cs_only=False,
            gl=args.gl,
            sleep_seconds=args.sleep,
            page_starts=(0, 10, 20),
            label="Pass 2: Riphah students (broader match + more pages)",
        )

    added_this_run = 0
    if new_rows:
        print(f"\nProcessing {len(new_rows)} new profile(s)...")
        for i, row in enumerate(new_rows, start=1):
            safe_print(f"  [{i}/{len(new_rows)}] {row.get('student_name') or row.get('linkedin_url')}")
            if not args.skip_profile_enrich:
                enrich_profile_via_serpapi(
                    session,
                    args.api_key,
                    row,
                    university=args.university,
                    gl=args.gl,
                    sleep_seconds=args.sleep,
                )
            if args.enrich_contacts:
                enrich_contact_via_serpapi(
                    session,
                    args.api_key,
                    row,
                    sleep_seconds=args.sleep,
                )
            before = len(stored_rows)
            stored_rows, added = upsert_and_save(
                stored_rows, row, args.output_csv, args.output_xlsx
            )
            if added or len(stored_rows) >= before:
                added_this_run += 1 if added else 0
                action = "added" if added else "updated"
                print(
                    f"      -> {action} ({len(stored_rows)} total) | {row.get('study_year')} | "
                    f"{row.get('current_job', '')[:45]} | {row.get('connections')}"
                )

    print()
    print(f"Added {added_this_run} new profile(s) this run ({len(stored_rows)} total in file):")
    print(f"  CSV  : {args.output_csv}")
    print(f"  XLSX : {args.output_xlsx}")
    if stored_rows:
        print("\nLatest entries:")
        for row in stored_rows[-5:]:
            safe_print(
                f"  - {row.get('student_name')} | {row.get('location')} | "
                f"{row.get('study_year')} | {row.get('connections')} | "
                f"{row.get('linkedin_url')}"
            )
    elif not new_rows:
        print("\nNo new profiles found. Check your SerpAPI quota or try increasing --max-profiles.")
    return 0 if stored_rows else 2


if __name__ == "__main__":
    raise SystemExit(main())
