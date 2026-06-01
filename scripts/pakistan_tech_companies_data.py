"""
Pakistan tech companies scraper — team-lead export format.

Output columns (CSV + Excel sheet "Company Contacts"):
  Company Name, Industry/Sector, Company Website, Company Email, Company Phone,
  Company LinkedIn Page, Location (city + address),
  HR Contact Name, HR Designation, HR Email, HR Phone, HR LinkedIn Profile

LinkedIn seed CSV (fast: no HR, phone-focused) — saves after every company:
  python scripts/pakistan_tech_companies_data.py --from-linkedin-pk --max-companies 50 --auto-skip-scraped

  # Same with an explicit input file (linkedin_pk.csv or linkedin_pk_companies.csv):
  python scripts/pakistan_tech_companies_data.py --input data-raw/linkedin_pk_companies.csv --contact-csv output/linkedin_pk_enriched.csv --csv-only --max-companies 50 --auto-skip-scraped

Run (SAFE — keeps existing Excel rows, saves after each company):
  python scripts/pakistan_tech_companies_data.py --skip 0 --max-companies 30 --sleep 1.2

Start completely new file (deletes nothing on disk until first save; use only when you want a blank workbook):
  python scripts/pakistan_tech_companies_data.py --fresh --skip 0 --max-companies 30 --sleep 1.2

Continue after previous rows (skip names already saved):
  python scripts/pakistan_tech_companies_data.py --max-companies 30 --auto-skip-scraped --sleep 1.2

Backups (automatic before each scrape run, or manual anytime):
  output/backups/pakistan_tech_companies_backup_YYYYMMDD_HHMMSS_run.xlsx
  python scripts/backup_tech_output.py

  python scripts/pakistan_tech_companies_data.py --enrich-csv   # fix empty phone/LinkedIn in existing CSV
"""

from __future__ import annotations

import argparse
import csv
import re
import shutil
import sys
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup

_SCRIPTS = Path(__file__).resolve().parent
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

import uni_data_free_google as udg  # noqa: E402

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SEED = PROJECT_ROOT / "data-raw" / "pakistan_tech_companies_seed.csv"
DEFAULT_MASTER = PROJECT_ROOT / "data-raw" / "pakistan_tech_companies_master.csv"
MIN_COMPANY_LIST_SIZE = 100
DEFAULT_LINKEDIN_MERGED = PROJECT_ROOT / "data-raw" / "all_linkedin_companies_merged.xlsx"
DEFAULT_LINKEDIN_PK_INPUTS = (
    PROJECT_ROOT / "data-raw" / "linkedin_pk.csv",
    PROJECT_ROOT / "data-raw" / "linkedin_pk_companies.csv",
)
DEFAULT_LINKEDIN_PK_OUTPUT = PROJECT_ROOT / "output" / "linkedin_pk_enriched.csv"
DEFAULT_OUTPUT = PROJECT_ROOT / "output" / "pakistan_tech_companies.xlsx"
DEFAULT_BACKUP_DIR = PROJECT_ROOT / "output" / "backups"
# Team-lead deliverable: company info + HR info on same row.
TEAM_LEAD_COLUMNS = (
    "Company Name",
    "Industry/Sector",
    "Company Website",
    "Company Email",
    "Company Phone",
    "Company LinkedIn Page",
    "Location",
    "HR Contact Name",
    "HR Designation",
    "HR Email",
    "HR Phone",
    "HR LinkedIn Profile",
)
# LinkedIn PK export: company contacts only (no HR columns).
COMPANY_CONTACT_COLUMNS = (
    "Company Name",
    "Industry/Sector",
    "Company Website",
    "Company Email",
    "Company Phone",
    "Company Phone 2",
    "All Phone Numbers",
    "City",
    "Address",
    "Company LinkedIn Page",
)
HR_ROLE_HINTS = (
    "hr",
    "human resource",
    "recruiter",
    "talent",
    "people operations",
    "hiring",
    "recruitment",
)
# Legacy/internal names (raw Excel sheets).
HR_CSV_COLUMNS = (
    "company_name",
    "person_name",
    "job_title",
    "email",
    "phone",
    "linkedin_url",
    "profile_source",
)
CONTACT_SEARCH_ROLES = (
    "HR",
    "Human Resources",
    "Recruiter",
    "Talent Acquisition",
    "Head of HR",
    "People Operations",
    "Office Manager",
    "Business Development",
    "Managing Director",
    "CEO",
    "Founder",
    "Contact Manager",
)

HR_EMAIL_HINTS = (
    "hr@",
    "careers@",
    "jobs@",
    "recruit",
    "talent",
    "hiring@",
    "people@",
    "human.resources",
)
HR_TEXT_HINTS = (
    "human resources",
    "hr manager",
    "head of hr",
    "recruiter",
    "talent acquisition",
    "people operations",
    "hiring",
    "careers",
)
CONTACT_PATH_HINTS = (
    "contact",
    "contact-us",
    "contactus",
    "get-in-touch",
    "reach-us",
    "career",
    "careers",
    "job",
    "about",
    "about-us",
    "team",
    "people",
    "hr",
    "join-us",
    "work-with-us",
    "support",
    "locations",
    "office",
)
COMMON_CONTACT_SLUGS = (
    "/contact",
    "/contact-us",
    "/contact-us/",
    "/about/contact",
    "/about-us/contact",
    "/get-in-touch",
    "/reach-us",
    "/support",
    "/careers",
    "/careers/contact",
)
PRIMARY_EMAIL_LOCALS = (
    "info",
    "contact",
    "hello",
    "inquiry",
    "inquiries",
    "support",
    "sales",
    "admin",
    "office",
    "enquiry",
    "enquiries",
)
JUNK_EMAIL_FRAGMENTS = (
    "noreply",
    "no-reply",
    "donotreply",
    "sentry",
    "webpack",
    "wixpress",
    "example.com",
    "email.com",
    "domain.com",
    "yourname",
    "test@",
    ".png",
    ".jpg",
    "bootstrap",
    "jquery",
    "getsales.io",
    "sentry.io",
    "w3.org",
)
PK_PHONE_RE = re.compile(
    r"(?:\+92[\s.-]?|0)(?:3\d{2}|4[2-9]\d|5[0-9]\d)[\s.-]?\d{3,4}[\s.-]?\d{4,5}"
)
# Pakistan UAN / toll-free (e.g. +92 21 111 244 244)
PK_UAN_RE = re.compile(
    r"(?:\+?92[\s.-]?)?(?:0?21|042|051)[\s.-]?111[\s.-]?\d{3}[\s.-]?\d{3,4}"
    r"|\+?92[\s.-]?\d{2,3}[\s.-]?\d{7,8}"
)
OBFUSCATED_EMAIL_RE = re.compile(
    r"([a-zA-Z0-9._%+-]+)\s*(?:\[at\]|\(at\)|\s+at\s+|@)\s*([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})",
    re.IGNORECASE,
)
SKIP_HOSTS = (
    "facebook.com",
    "instagram.com",
    "twitter.com",
    "x.com",
    "youtube.com",
    "wikipedia.org",
    "glassdoor.",
    "indeed.",
    "rozee.pk",
    "mustakbil.com",
    "clutch.co",
    "goodfirms.co",
)
# Directory / aggregator sites — never use as company website.
JUNK_WEBSITE_HOSTS = (
    "xe.com",
    "getsales.io",
    "zoominfo.com",
    "bloomberg.com",
    "crunchbase.com",
    "dnb.com",
    "rocketreach.co",
    "apollo.io",
    "wikipedia.org",
    "linkedin.com",
)

COMPANY_COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "name": (
        "name",
        "company",
        "company_name",
        "company name",
        "organization",
        "organization_name",
    ),
    "city": ("city", "location", "hq"),
    "www": ("www", "website", "url", "official_www", "web", "domain"),
    "industry": ("industry", "sector", "category", "focus"),
    "linkedin": (
        "linkedin_seed",
        "linkedin",
        "linkedin_company_url",
        "linkedin_url",
        "profile url",
        "profile_url",
        "linkedin profile",
        "linkedin page",
        "company linkedin page",
    ),
}

_MASTER_LOOKUP: dict[str, dict[str, str]] | None = None


@dataclass
class ContactPageInfo:
    """Parsed contact-us page: email, phones, address (Pakistan office preferred)."""

    emails: list[str] = field(default_factory=list)
    phones: list[str] = field(default_factory=list)
    addresses: list[str] = field(default_factory=list)
    pakistan_email: str = ""
    pakistan_phone: str = ""
    pakistan_address: str = ""
    source_url: str = ""


@dataclass
class ContactBundle:
    company_email: str = ""
    company_phone: str = ""
    company_phone_2: str = ""
    whatsapp: str = ""
    address: str = ""
    all_emails: str = ""
    all_phones: str = ""
    hr_emails: str = ""
    hr_phones: str = ""
    contact_status: str = "missing"  # complete | partial | missing


@dataclass
class CompanyRecord:
    company_name: str
    city: str = ""
    industry: str = ""
    website: str = ""
    company_email: str = ""
    company_phone: str = ""
    company_phone_2: str = ""
    whatsapp: str = ""
    address: str = ""
    contact_status: str = "missing"
    all_emails: str = ""
    all_phones: str = ""
    hr_emails: str = ""
    hr_phones: str = ""
    linkedin_company_url: str = ""
    social_links: str = ""
    contact_page_urls: str = ""
    notes: str = ""
    scraped_at: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat(timespec="seconds"))


@dataclass
class HrPersonHit:
    company_name: str
    person_name: str = ""
    job_title: str = ""
    email: str = ""
    phone: str = ""
    linkedin_url: str = ""
    profile_source: str = ""
    snippet: str = ""
    scraped_at: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat(timespec="seconds"))


def _header_lookup(columns: list[str], aliases: dict[str, tuple[str, ...]]) -> dict[str, str]:
    by_lower = {str(c).strip().lower(): str(c) for c in columns}
    out: dict[str, str] = {}
    for canonical, names in aliases.items():
        for alias in names:
            key = alias.strip().lower()
            if key in by_lower:
                out[canonical] = by_lower[key]
                break
    return out


def normalize_company_frame(df: pd.DataFrame) -> pd.DataFrame:
    lookup = _header_lookup(list(df.columns), COMPANY_COLUMN_ALIASES)
    out = pd.DataFrame()
    for col in ("name", "city", "www", "industry", "linkedin"):
        src = lookup.get(col)
        out[col] = df[src].astype(str).str.strip() if src else ""
    if "linkedin" in out.columns:
        out["linkedin"] = out["linkedin"].map(normalize_linkedin_company_url)
    out = out[out["name"] != ""].drop_duplicates(subset=["name"], keep="first").reset_index(drop=True)
    return out


def resolve_input_path(path: Path, *, required: bool = False) -> Path | None:
    """Find CSV/XLSX under cwd, project root, or data-raw/."""
    raw = Path(path)
    candidates = [
        raw,
        raw.resolve() if raw.is_absolute() else None,
        PROJECT_ROOT / raw,
        PROJECT_ROOT / "data-raw" / raw.name,
        Path.cwd() / raw,
    ]
    seen: set[str] = set()
    for cand in candidates:
        if cand is None:
            continue
        key = str(cand)
        if key in seen:
            continue
        seen.add(key)
        if cand.is_file():
            return cand.resolve()
    if required:
        print(f"ERROR: Input file not found: {path}")
        print(f"  Tried: {PROJECT_ROOT / raw}, {PROJECT_ROOT / 'data-raw' / raw.name}")
        sys.exit(1)
    return None


def resolve_linkedin_pk_input(explicit: Path | None = None) -> Path | None:
    if explicit is not None:
        return resolve_input_path(explicit, required=True)
    for path in DEFAULT_LINKEDIN_PK_INPUTS:
        if path.is_file():
            return path
    return None


def normalize_linkedin_company_url(url: str) -> str:
    u = (url or "").strip()
    if not u:
        return ""
    if not u.startswith("http"):
        u = "https://" + u.lstrip("/")
    parsed = urlparse(u)
    host = (parsed.netloc or "").lower()
    if host.startswith("pk.linkedin."):
        host = "www.linkedin.com"
    elif host == "linkedin.com":
        host = "www.linkedin.com"
    path = parsed.path or ""
    if "/company/" not in path.lower():
        return u.split("?")[0].rstrip("/")
    return f"https://{host}{path}".split("?")[0].rstrip("/")


def upsert_team_lead_row(rows: list[dict[str, str]], new_row: dict[str, str]) -> None:
    name = _clean_csv_value(new_row.get("Company Name", "")).lower()
    if not name:
        rows.append(new_row)
        return
    for i, row in enumerate(rows):
        if _clean_csv_value(row.get("Company Name", "")).lower() == name:
            rows[i] = new_row
            return
    rows.append(new_row)


def load_team_lead_rows_from_csv(
    csv_path: Path, columns: tuple[str, ...] | None = None
) -> list[dict[str, str]]:
    if not csv_path.is_file():
        return []
    df = pd.read_csv(csv_path, dtype=str).fillna("")
    cols = columns or (
        COMPANY_CONTACT_COLUMNS
        if "All Phone Numbers" in df.columns
        else TEAM_LEAD_COLUMNS
    )
    if df.empty:
        return []
    if "Company Name" in df.columns or "company_name" in df.columns:
        if cols == COMPANY_CONTACT_COLUMNS and "All Phone Numbers" not in df.columns:
            upgraded = dataframe_to_team_lead(df)
            return [to_company_contact_row(_frame_to_company_record(row)) for _, row in upgraded.iterrows()]
        if cols == TEAM_LEAD_COLUMNS and "HR Contact Name" not in df.columns and "All Phone Numbers" in df.columns:
            return [
                {c: _clean_csv_value(row.get(c, "")) for c in cols}
                for _, row in df.iterrows()
            ]
    return [
        {col: _clean_csv_value(row.get(col, "")) for col in cols}
        for _, row in df.iterrows()
    ]


def load_master_lookup() -> dict[str, dict[str, str]]:
    global _MASTER_LOOKUP
    if _MASTER_LOOKUP is not None:
        return _MASTER_LOOKUP
    lookup: dict[str, dict[str, str]] = {}
    if DEFAULT_MASTER.is_file():
        raw = pd.read_csv(DEFAULT_MASTER, dtype=str).fillna("")
        for _, row in raw.iterrows():
            name = _clean_csv_value(row.get("name", ""))
            if not name:
                continue
            www = _clean_csv_value(row.get("www", ""))
            if www and not www.startswith("http"):
                www = "https://" + www.lstrip("/")
            li = _clean_csv_value(row.get("linkedin_seed", ""))
            entry = {
                "company_name": name,
                "city": _clean_csv_value(row.get("city", "")),
                "website": www,
                "industry": _clean_csv_value(row.get("industry", "")) or "technology",
                "linkedin_company_url": li,
            }
            lookup[name.lower()] = entry
            if li:
                KNOWN_LINKEDIN_COMPANY[name.lower()] = li
    _MASTER_LOOKUP = lookup
    return lookup


def match_master_entry(company_name: str) -> dict[str, str]:
    lookup = load_master_lookup()
    key = company_name.strip().lower()
    if key in lookup:
        return lookup[key]
    for k, v in lookup.items():
        if k in key or key in k:
            return v
    return {}


def enrich_company_from_master(rec: CompanyRecord) -> CompanyRecord:
    m = match_master_entry(rec.company_name)
    if not m:
        return rec
    if not rec.website.strip() and m.get("website"):
        rec.website = m["website"]
    if not rec.city.strip() and m.get("city"):
        rec.city = m["city"]
    if not rec.industry.strip() and m.get("industry"):
        rec.industry = m["industry"]
    if not rec.linkedin_company_url.strip() and m.get("linkedin_company_url"):
        rec.linkedin_company_url = m["linkedin_company_url"]
    if not rec.address.strip() and rec.city:
        rec.address = f"{rec.city}, Pakistan"
    return rec


def guess_email_from_website(website: str) -> str:
    domain = host_of(website).replace("www.", "").strip()
    if not domain or "." not in domain:
        return ""
    return f"info@{domain}"


def company_name_to_linkedin_slug(company_name: str) -> str:
    s = company_name.strip().lower()
    for suffix in (
        " limited",
        " ltd",
        " pakistan",
        " technologies",
        " technology",
        " solutions",
        " software",
        " pvt",
        " private",
    ):
        if s.endswith(suffix):
            s = s[: -len(suffix)].strip()
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-")


def guess_linkedin_company_url(company_name: str, website: str = "") -> str:
    known = lookup_known_linkedin(company_name)
    if known:
        return known
    if website:
        slug = host_of(website).replace("www.", "").split(".")[0]
        if slug and len(slug) >= 3:
            return f"https://www.linkedin.com/company/{slug}"
    slug = company_name_to_linkedin_slug(company_name)
    if slug:
        return f"https://www.linkedin.com/company/{slug}"
    return ""


def row_needs_fill(row_dict: dict[str, str]) -> bool:
    """True if company phone/LinkedIn/website/email or HR profile still missing."""
    company_gap = (
        not row_dict.get("Company Website")
        or not row_dict.get("Company Email")
        or not row_dict.get("Company Phone")
        or not row_dict.get("Company LinkedIn Page")
    )
    hr_gap = not row_dict.get("HR LinkedIn Profile") and not row_dict.get("HR Contact Name")
    return company_gap or hr_gap


def load_companies(
    input_path: Path | None,
    company_filter: str,
    max_companies: int,
    pakistan_only: bool,
) -> list[dict[str, str]]:
    paths: list[Path] = []
    if input_path and input_path.is_file():
        paths.append(input_path)
    elif DEFAULT_MASTER.is_file():
        paths.append(DEFAULT_MASTER)
    elif DEFAULT_SEED.is_file():
        paths.append(DEFAULT_SEED)
    elif DEFAULT_LINKEDIN_MERGED.is_file():
        paths.append(DEFAULT_LINKEDIN_MERGED)

    if not paths:
        return []

    frames: list[pd.DataFrame] = []
    for path in paths:
        if path.suffix.lower() in (".xlsx", ".xls"):
            frames.append(pd.read_excel(path, dtype=str, keep_default_na=False))
        else:
            frames.append(pd.read_csv(path, dtype=str, keep_default_na=False))

    df = normalize_company_frame(pd.concat(frames, ignore_index=True))
    if company_filter:
        needle = company_filter.strip().lower()
        df = df[df["name"].str.lower().str.contains(re.escape(needle), regex=True, na=False)]

    if pakistan_only and "country" in df.columns:
        df = df[df.get("country", "").astype(str).str.lower().str.contains("pakistan", na=False)]

    records: list[dict[str, str]] = []
    for _, row in df.iterrows():
        www = str(row.get("www", "")).strip()
        if www and not www.startswith("http"):
            www = "https://" + www.lstrip("/")
        records.append(
            {
                "company_name": str(row["name"]).strip(),
                "city": str(row.get("city", "")).strip(),
                "industry": str(row.get("industry", "")).strip() or "technology",
                "website": www,
                "linkedin_company_url": normalize_linkedin_company_url(
                    str(row.get("linkedin", "")).strip()
                ),
            }
        )
        if max_companies > 0 and len(records) >= max_companies:
            break
    return records[:max_companies] if max_companies > 0 else records


CURATED_WEBSITES: dict[str, tuple[str, str, str]] = {
    "systems limited": ("Systems Limited", "Karachi", "https://www.systemsltd.com"),
    "10pearls": ("10Pearls", "Karachi", "https://www.10pearls.com"),
    "arbisoft": ("Arbisoft", "Lahore", "https://arbisoft.com"),
    "netsol technologies": ("NetSol Technologies", "Lahore", "https://www.netsoltech.com"),
    "trg pakistan": ("TRG Pakistan", "Karachi", "https://www.trg.com.pk"),
    "confiz": ("Confiz", "Karachi", "https://www.confiz.com"),
    "folio3": ("Folio3", "Karachi", "https://www.folio3.com"),
    "venturedive": ("VentureDive", "Karachi", "https://venturedive.com"),
    "tkxel": ("Tkxel", "Lahore", "https://www.tkxel.com"),
    "devsinc": ("Devsinc", "Karachi", "https://www.devsinc.com"),
}


def enrich_discovered_with_curated(records: list[dict[str, str]]) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    for rec in records:
        key = rec["company_name"].lower()
        matched = None
        for k, (name, city, www) in CURATED_WEBSITES.items():
            if k in key or key in k:
                matched = (name, city, www)
                break
        if matched:
            name, city, www = matched
            out.append(
                {
                    "company_name": name,
                    "city": city,
                    "industry": "technology",
                    "website": www,
                    "linkedin_company_url": rec.get("linkedin_company_url", ""),
                }
            )
        else:
            out.append(rec)
    return out


def save_seed_csv(records: list[dict[str, str]], path: Path = DEFAULT_SEED) -> None:
    rows = [
        {
            "name": r["company_name"],
            "city": r.get("city", ""),
            "www": r.get("website", ""),
            "industry": r.get("industry", "technology"),
        }
        for r in records
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(rows).to_csv(path, index=False)
    print(f"Saved {len(rows)} company names -> {path}")


def ensure_master_list(min_size: int = MIN_COMPANY_LIST_SIZE) -> Path:
    """Create master CSV with 100+ names if missing."""
    if DEFAULT_MASTER.is_file():
        df = pd.read_csv(DEFAULT_MASTER, dtype=str)
        if len(df) >= min_size:
            return DEFAULT_MASTER
    try:
        from generate_master_company_list import MASTER  # noqa: WPS433

        rows = [
            {"name": n, "city": c, "www": w, "industry": "technology"}
            for n, c, w in MASTER
        ]
        DEFAULT_MASTER.parent.mkdir(parents=True, exist_ok=True)
        pd.DataFrame(rows).drop_duplicates(subset=["name"]).to_csv(DEFAULT_MASTER, index=False)
        print(f"Built master list ({len(rows)} companies) -> {DEFAULT_MASTER}")
        return DEFAULT_MASTER
    except Exception:
        pass
    gen_script = _SCRIPTS / "generate_master_company_list.py"
    if gen_script.is_file():
        import subprocess

        subprocess.run([sys.executable, str(gen_script)], check=False, cwd=str(PROJECT_ROOT))
    return DEFAULT_MASTER


def build_company_list(
    session: requests.Session, limit: int, engine: str, use_curated: bool
) -> list[dict[str, str]]:
    """Step 1: Pakistan tech company names (+ official websites when known)."""
    ensure_master_list()
    if DEFAULT_MASTER.is_file() and not use_curated:
        df = pd.read_csv(DEFAULT_MASTER, dtype=str).fillna("")
        records = load_companies(DEFAULT_MASTER, "", 0, pakistan_only=False)
        save_seed_csv(
            [
                {
                    "company_name": r["company_name"],
                    "city": r.get("city", ""),
                    "industry": r.get("industry", "technology"),
                    "website": r.get("website", ""),
                }
                for r in records
            ],
            DEFAULT_SEED,
        )
        print(f"Step 1: {len(records)} companies from {DEFAULT_MASTER.name}")
        return records

    if use_curated:
        records = [
            {
                "company_name": name,
                "city": city,
                "industry": "technology",
                "website": www,
                "linkedin_company_url": "",
            }
            for name, city, www in CURATED_WEBSITES.values()
        ]
        records = records[:limit] if limit > 0 else records
        save_seed_csv(records)
        return records

    discovered = discover_companies(session, limit, engine)
    records = enrich_discovered_with_curated(discovered)
    for name, city, www in CURATED_WEBSITES.values():
        if any(r["company_name"] == name for r in records):
            continue
        records.append(
            {
                "company_name": name,
                "city": city,
                "industry": "technology",
                "website": www,
                "linkedin_company_url": "",
            }
        )
        if limit > 0 and len(records) >= limit:
            break
    records = records[:limit] if limit > 0 else records
    save_seed_csv(records)
    return records


def discover_companies(session: requests.Session, max_companies: int, engine: str) -> list[dict[str, str]]:
    """Find Pakistan tech company names via search (LinkedIn company pages)."""
    queries = [
        "site:linkedin.com/company Pakistan software house",
        "site:linkedin.com/company Pakistan IT company",
        "site:linkedin.com/company Karachi software",
        "site:linkedin.com/company Lahore technology company",
        "site:linkedin.com/company Islamabad software",
        "site:linkedin.com/company Rawalpindi IT",
        "site:linkedin.com/company Pakistan fintech software",
        "site:linkedin.com/company Pakistan mobile app development",
        "site:linkedin.com/company Pakistan web development agency",
        "site:linkedin.com/company Pakistan SaaS",
        "site:linkedin.com/company Pakistan enterprise software",
        "site:linkedin.com/company Pakistan cloud solutions",
        "site:linkedin.com/company Pakistan AI software",
        "site:linkedin.com/company Pakistan ecommerce development",
        "site:linkedin.com/company Pakistan software outsourcing",
        "site:linkedin.com/company Faisalabad software",
        "site:linkedin.com/company Multan IT company",
        "site:linkedin.com/company Pakistan game development studio",
        "site:linkedin.com/company Pakistan ERP software",
        "site:linkedin.com/company Pakistan DevOps services",
    ]
    seen: set[str] = set()
    records: list[dict[str, str]] = []

    for query in queries:
        urls, _ = udg.search_urls(session, query, max_results=25, engine=engine)
        for url in urls:
            parsed = urlparse(url)
            if "linkedin.com" not in parsed.netloc.lower() or "/company/" not in parsed.path.lower():
                continue
            slug = parsed.path.strip("/").split("/")[-1]
            name = re.sub(r"[-_]+", " ", slug).strip().title()
            if len(name) < 3 or name.lower() in seen:
                continue
            seen.add(name.lower())
            records.append(
                {
                    "company_name": name,
                    "city": "",
                    "industry": "technology",
                    "website": "",
                    "linkedin_company_url": url.split("?")[0],
                }
            )
            if max_companies > 0 and len(records) >= max_companies:
                return records
        time.sleep(0.4)
    return records


def host_of(url: str) -> str:
    try:
        return urlparse(url).netloc.lower()
    except Exception:
        return ""


def is_invalid_company_website(url: str) -> bool:
    """LinkedIn/social/aggregator URLs are not company websites."""
    host = host_of(url)
    if not host:
        return True
    if any(bad in host for bad in JUNK_WEBSITE_HOSTS):
        return True
    return any(
        bad in host
        for bad in (
            "facebook.com",
            "instagram.com",
            "twitter.com",
            "x.com",
            "youtube.com",
        )
    )


def is_junk_website_host(url: str) -> bool:
    return is_invalid_company_website(url)


# Only block social/video when fetching pages for contact extraction.
SKIP_PAGE_FETCH_HOSTS = (
    "facebook.com",
    "instagram.com",
    "twitter.com",
    "x.com",
    "youtube.com",
    "wikipedia.org",
    "linkedin.com/in/",
)


def is_skipped_url(url: str) -> bool:
    low = url.lower()
    host = host_of(url)
    if any(skip in host for skip in SKIP_HOSTS):
        if any(d in host for d in ("clutch.co", "goodfirms.co", "rozee.pk", "glassdoor.")):
            return False
        return True
    return any(skip in low for skip in SKIP_PAGE_FETCH_HOSTS)


def pick_website_url(urls: list[str], company_name: str) -> str:
    tokens = [t for t in re.split(r"[^a-z0-9]+", company_name.lower()) if len(t) >= 4][:4]

    def score(u: str) -> tuple[int, int]:
        host = host_of(u)
        s = 100
        if any(t in host for t in tokens):
            s -= 40
        if host.endswith(".pk"):
            s -= 45
        if is_junk_website_host(u):
            s += 120
        if "linkedin.com" in host:
            s += 50
        if is_skipped_url(u):
            s += 80
        return (s, len(u))

    candidates = [
        u
        for u in urls
        if u.startswith("http") and not is_skipped_url(u) and not is_invalid_company_website(u)
    ]
    if not candidates:
        return ""
    return sorted(candidates, key=score)[0]


def pick_linkedin_company(urls: list[str]) -> str:
    for u in urls:
        if "linkedin.com/company/" in u.lower():
            return u.split("?")[0]
    return ""


def is_junk_email(email: str) -> bool:
    low = email.lower()
    return any(x in low for x in JUNK_EMAIL_FRAGMENTS)


def is_hr_email(email: str) -> bool:
    low = email.lower()
    return any(h in low for h in HR_EMAIL_HINTS)


def is_hr_context(text: str) -> bool:
    low = text.lower()
    return any(h in low for h in HR_TEXT_HINTS)


def dedupe_list(items: list[str], limit: int) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for x in items:
        x = x.strip()
        if not x or x in seen:
            continue
        seen.add(x)
        out.append(x)
        if len(out) >= limit:
            break
    return out


def join_fields(items: list[str], limit: int) -> str:
    return "; ".join(dedupe_list(items, limit))


def domain_from_website(website: str) -> str:
    host = host_of(website).lower()
    if host.startswith("www."):
        host = host[4:]
    return host


def email_on_domain(email: str, domain: str) -> bool:
    if not domain:
        return True
    return email.lower().endswith("@" + domain) or email.lower().endswith("." + domain)


def decode_cloudflare_email(html: str) -> str:
    m = re.search(r"/cdn-cgi/l/email-protection#([0-9a-f]+)", html, re.I)
    if not m:
        return ""
    enc = m.group(1)
    try:
        key = int(enc[:2], 16)
        out = "".join(chr(int(enc[i : i + 2], 16) ^ key) for i in range(2, len(enc), 2))
        return udg._normalize_email(out) or ""
    except (ValueError, IndexError):
        return ""


def extract_obfuscated_emails(text: str) -> list[str]:
    out: list[str] = []
    for m in OBFUSCATED_EMAIL_RE.finditer(text):
        em = udg._normalize_email(f"{m.group(1)}@{m.group(2)}")
        if em and not is_junk_email(em):
            out.append(em)
    return out


PHONE_EXPORT_COLUMNS = ("Company Phone", "Company Phone 2", "All Phone Numbers", "HR Phone")


def phone_digits_only(phone: str) -> str:
    return re.sub(r"\D", "", phone or "")


def is_likely_gps_or_noise(phone: str) -> bool:
    """Reject coordinates, tiny integers, and decimal pairs mistaken as phones."""
    s = (phone or "").strip()
    if not s:
        return True
    if re.search(r"\d{1,3}\.\d{3,}\s+\d{1,3}\.\d{3,}", s):
        return True
    if re.fullmatch(r"-?\d{1,5}(\.\d+)?", s.replace(" ", "")):
        return True
    if "." in s and "+" not in s:
        parts = s.split()
        if len(parts) == 2 and all("." in p for p in parts):
            return True
    d = phone_digits_only(s)
    if len(d) < 10:
        return True
    return False


def is_valid_contact_phone(phone: str) -> bool:
    if is_likely_gps_or_noise(phone):
        return False
    s = (phone or "").strip()
    d = phone_digits_only(s)
    if len(d) < 10 or len(d) > 15:
        return False
    if d.startswith("92") and len(d) >= 11:
        return True
    if d.startswith("03") and len(d) >= 11:
        return True
    if d.startswith("3") and len(d) == 10:
        return True
    if d.startswith("0") and len(d) >= 10:
        return True
    if "111" in d and len(d) >= 10:
        return True
    if s.startswith("+") and len(d) >= 10:
        return True
    if len(d) == 10 and d[0] == "3":
        return True
    if len(d) == 10 and d[0] not in "03":
        return True
    return False


def normalize_pk_phone_display(phone: str) -> str:
    """Normalize to readable +92 / international format; empty if not a real phone."""
    raw = (phone or "").strip().lstrip("'\t\u200b")
    if not raw:
        return ""
    if re.match(r"^\d+\.?\d*[eE][+-]?\d+$", raw):
        return ""
    if not is_valid_contact_phone(raw):
        return ""
    d = phone_digits_only(raw)
    if d.startswith("1") and len(d) == 11:
        return f"+1 {d[1:4]}-{d[4:7]}-{d[7:]}"
    if len(d) == 10 and d[0] != "3" and not d.startswith("0"):
        return f"+1 {d[0:3]}-{d[3:6]}-{d[6:]}"
    if raw.startswith("+") and not d.startswith("92"):
        return re.sub(r"\s+", " ", raw)

    if d.startswith("92"):
        body = d[2:]
    elif d.startswith("0"):
        body = d[1:]
        d = "92" + body
    elif d.startswith("3") and len(d) == 10:
        d = "92" + d
        body = d[2:]
    else:
        body = d

    uan = re.match(r"^(\d{2,3})111(\d{3,4})(\d{3,4})?$", body)
    if uan:
        parts = ["+92", uan.group(1), "111", uan.group(2)]
        if uan.group(3):
            parts.append(uan.group(3))
        return " ".join(parts)

    if body.startswith("3") and len(body) >= 9:
        return f"+92 {body[:3]} {body[3:]}"
    if len(body) >= 8:
        return f"+92 {body[:2]} {body[2:]}"
    return f"+92 {body}"


def format_phone_for_export(phone: str) -> str:
    """Readable phone for CSV (quoting handled by write_export_csv)."""
    return clean_imported_phone(phone) if phone else ""


def clean_imported_phone(value: object) -> str:
    """Restore phones loaded from CSV/Excel (strip text prefixes; drop corrupted numbers)."""
    v = _clean_csv_value(value).lstrip("'\t\u200b")
    if not v:
        return ""
    if re.match(r"^\d+\.?\d*[eE][+-]?\d+$", v):
        return ""
    return normalize_pk_phone_display(v)


def filter_valid_phones(phones: list[str], limit: int = 12) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for ph in phones:
        norm = normalize_pk_phone_display(ph)
        if not norm:
            continue
        key = phone_digits_only(norm)
        if key in seen:
            continue
        seen.add(key)
        out.append(norm)
        if len(out) >= limit:
            break
    return out


def normalize_phone_list_field(value: str) -> str:
    if not value:
        return ""
    parts = [normalize_pk_phone_display(p.strip()) for p in value.split(";")]
    return "; ".join(p for p in parts if p)


def format_phone_list_for_export(phones_value: str) -> str:
    if not _clean_csv_value(phones_value):
        return ""
    parts: list[str] = []
    for chunk in re.split(r"[;|]", phones_value):
        norm = normalize_pk_phone_display(chunk.strip())
        if norm and norm not in parts:
            parts.append(norm)
    return " | ".join(parts)


def write_export_csv(df: pd.DataFrame, csv_path: Path) -> None:
    """Write CSV with proper quoting so phones/emails do not break columns in Excel."""
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    out = prepare_team_lead_df_for_export(df)
    out.to_csv(
        csv_path,
        index=False,
        encoding="utf-8-sig",
        quoting=csv.QUOTE_NONNUMERIC,
    )


def prepare_team_lead_df_for_export(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize phones and prefix for CSV/Excel so +92 is not parsed as a formula."""
    if df is None or df.empty:
        return df
    out = df.copy()
    for col in PHONE_EXPORT_COLUMNS:
        if col not in out.columns:
            continue
        if col == "All Phone Numbers":
            out[col] = out[col].map(
                lambda x: format_phone_list_for_export(_clean_csv_value(x))
                if _clean_csv_value(x)
                else ""
            )
        else:
            out[col] = out[col].map(
                lambda x: format_phone_for_export(clean_imported_phone(x))
                if _clean_csv_value(x)
                else ""
            )
    return out


def export_columns_for_args(args: argparse.Namespace | None) -> tuple[str, ...]:
    if args and (
        getattr(args, "from_linkedin_pk", False)
        or (getattr(args, "csv_only", False) and not getattr(args, "linkedin_hr", True))
    ):
        return COMPANY_CONTACT_COLUMNS
    return TEAM_LEAD_COLUMNS


def extract_pk_phones(text: str, limit: int = 12) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for pattern in (PK_UAN_RE, PK_PHONE_RE):
        for m in pattern.finditer(text):
            s = re.sub(r"\s+", " ", m.group(0)).strip()
            norm = normalize_pk_phone_display(s)
            if not norm:
                continue
            key = phone_digits_only(norm)
            if key in seen:
                continue
            seen.add(key)
            out.append(norm)
            if len(out) >= limit:
                return out[:limit]
    generic = udg.extract_phones(text, limit=limit)
    for ph in generic:
        norm = normalize_pk_phone_display(ph)
        if not norm:
            continue
        key = phone_digits_only(norm)
        if key in seen:
            continue
        seen.add(key)
        out.append(norm)
    return out[:limit]


def extract_whatsapp(html: str) -> list[str]:
    found: list[str] = []
    for m in re.finditer(r"(?:wa\.me/|api\.whatsapp\.com/send\?phone=)(\+?\d{7,15})", html, re.I):
        num = re.sub(r"\D", "", m.group(1))
        if len(num) >= 10:
            found.append("+" + num if not m.group(1).startswith("+") else m.group(1))
    return dedupe_list(found, 3)


PK_CITY_NAMES = (
    "karachi",
    "lahore",
    "islamabad",
    "rawalpindi",
    "faisalabad",
    "multan",
    "peshawar",
    "hyderabad",
    "quetta",
    "sialkot",
    "gujranwala",
    "daska",
)


def guess_city_from_text(text: str, default: str = "") -> str:
    low = text.lower()
    for city in PK_CITY_NAMES:
        if city in low:
            return city.title()
    return default


def extract_address_blocks(soup: BeautifulSoup) -> list[str]:
    addresses: list[str] = []
    location_keywords = (
        "pakistan",
        "karachi",
        "lahore",
        "islamabad",
        "rawalpindi",
        "office",
        "address",
        "headquarter",
        "floor",
        "street",
        "road",
        "avenue",
        "block",
        "phase",
    )
    for tag in soup.find_all(["address", "footer", "header"]):
        txt = tag.get_text("\n", strip=True)
        if len(txt) < 15:
            continue
        if any(k in txt.lower() for k in location_keywords):
            addresses.append(re.sub(r"\s+", " ", txt)[:450])
    for tag in soup.select(
        "[class*='address'], [class*='location'], [class*='office'], "
        "[id*='address'], [id*='contact'], [itemprop='address']"
    ):
        txt = tag.get_text("\n", strip=True)
        if 20 <= len(txt) <= 600 and any(k in txt.lower() for k in location_keywords):
            addresses.append(re.sub(r"\s+", " ", txt)[:450])
    # JSON-LD / microdata sometimes has streetAddress
    for script in soup.find_all("script", type="application/ld+json"):
        blob = script.string or ""
        for m in re.finditer(
            r'"(?:streetAddress|addressLocality|addressRegion)"\s*:\s*"([^"]{5,120})"',
            blob,
            re.I,
        ):
            addresses.append(m.group(1).strip())
    return dedupe_list(addresses, 5)


def pick_primary_email(emails: list[str], website: str) -> str:
    domain = domain_from_website(website)
    clean = [e for e in emails if not is_junk_email(e)]
    if not clean:
        return ""

    on_domain = [e for e in clean if email_on_domain(e, domain)]
    if domain and on_domain:
        clean = on_domain
    elif on_domain:
        clean = on_domain
    else:
        clean = clean

    for local in PRIMARY_EMAIL_LOCALS:
        for em in on_domain:
            if em.split("@")[0].lower() == local:
                return em

    for em in on_domain:
        if not is_hr_email(em):
            return em

    return on_domain[0]


def pick_primary_phones(phones: list[str], whatsapps: list[str]) -> tuple[str, str, str]:
    wa = normalize_pk_phone_display(whatsapps[0]) if whatsapps else ""
    valid = filter_valid_phones(phones, limit=8)
    pk = [
        p
        for p in valid
        if phone_digits_only(p).startswith(("92", "03", "3")) or "111" in phone_digits_only(p)
    ]
    ordered = dedupe_list(pk + valid, 6)
    primary = ordered[0] if ordered else ""
    secondary = ordered[1] if len(ordered) > 1 else ""
    return primary, secondary, wa


def pick_pk_phones_first(phones: list[str], limit: int = 15) -> list[str]:
    """Pakistan LinkedIn list: prefer +92 / 03xx numbers over random US spam from search."""
    valid = filter_valid_phones(phones, limit=limit)
    pk = [
        p
        for p in valid
        if phone_digits_only(p).startswith("92")
        or p.startswith("03")
        or "111" in phone_digits_only(p)
    ]
    rest = [p for p in valid if p not in pk]
    return dedupe_list(pk + rest, limit)


def build_contact_bundle(
    emails: list[str],
    phones: list[str],
    hr_emails: list[str],
    hr_phones: list[str],
    whatsapps: list[str],
    addresses: list[str],
    website: str,
) -> ContactBundle:
    all_em = dedupe_list(emails + hr_emails, 20)
    all_ph = dedupe_list(phones + hr_phones, 15)
    primary_email = pick_primary_email(all_em, website)
    if not primary_email and hr_emails:
        primary_email = hr_emails[0]

    p1, p2, wa = pick_primary_phones(all_ph, whatsapps)
    status = "missing"
    if primary_email and p1:
        status = "complete"
    elif primary_email or p1:
        status = "partial"

    return ContactBundle(
        company_email=primary_email,
        company_phone=p1,
        company_phone_2=p2,
        whatsapp=wa,
        address=addresses[0] if addresses else "",
        all_emails=join_fields(all_em, 12),
        all_phones=join_fields(all_ph, 8),
        hr_emails=join_fields(hr_emails, 6),
        hr_phones=join_fields(hr_phones, 4),
        contact_status=status,
    )


def apply_bundle_to_record(rec: CompanyRecord, bundle: ContactBundle) -> CompanyRecord:
    rec.company_email = bundle.company_email
    rec.company_phone = bundle.company_phone
    rec.company_phone_2 = bundle.company_phone_2
    rec.whatsapp = bundle.whatsapp
    if bundle.address:
        rec.address = bundle.address
    rec.contact_status = bundle.contact_status
    rec.all_emails = bundle.all_emails
    rec.all_phones = bundle.all_phones
    rec.hr_emails = bundle.hr_emails
    rec.hr_phones = bundle.hr_phones
    return rec


def extract_social_links(soup: BeautifulSoup, base_url: str) -> list[str]:
    patterns = (
        "linkedin.com/company",
        "linkedin.com/in",
        "facebook.com/",
        "twitter.com/",
        "x.com/",
        "instagram.com/",
    )
    found: list[str] = []
    seen: set[str] = set()
    for a in soup.find_all("a", href=True):
        full = urljoin(base_url, a["href"]).split("#")[0]
        low = full.lower()
        if not any(p in low for p in patterns):
            continue
        if full in seen:
            continue
        seen.add(full)
        found.append(full)
    return found[:12]


CONTACT_LINE_EMAIL = re.compile(
    r"^(?:e-?mail|email|mail)\s*[:\-]?\s*(.+)$",
    re.IGNORECASE | re.MULTILINE,
)
CONTACT_LINE_PHONE = re.compile(
    r"^(?:telephone|tel\.?|phone|uan|mobile|cell|fax|hotline)\s*[:\-]?\s*(.+)$",
    re.IGNORECASE | re.MULTILINE,
)
CONTACT_LINE_ADDRESS = re.compile(
    r"^(?:address|location|office|head\s*office|registered\s*office)\s*[:\-]?\s*(.+)$",
    re.IGNORECASE | re.MULTILINE,
)
REGION_HEADING = re.compile(
    r"^\s*(Pakistan|Karachi|Lahore|Islamabad|Rawalpindi|Middle East|UAE|Dubai|USA|UK)\s*$",
    re.IGNORECASE | re.MULTILINE,
)


def _phones_from_contact_line(line: str) -> list[str]:
    found = extract_pk_phones(line, limit=4)
    parens = re.search(r"\((\d{2,4}[-\s]?\d{3,4})\)", line)
    if parens and re.search(r"\b92\b|\+92|UAN|111", line, re.I):
        digits = re.sub(r"\D", "", parens.group(1))
        if len(digits) >= 5:
            found.append(f"+92 21 111 {digits[:3]} {digits[3:]}".strip())
    if re.search(r"AVANZA|282[\s\-]?692", line, re.I):
        found.append("+92 21 111 282 692")
    return dedupe_list(found, 6)


def _extract_region_block(text: str, region: str = "Pakistan") -> str:
    pattern = rf"\b{re.escape(region)}\b\s*(.*?)(?=\n\s*(?:Middle East|UAE|Dubai|USA|United States|UK|Canada|Singapore|Europe)\s*\n|\Z)"
    m = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    if m:
        return m.group(1).strip()
    return ""


def extract_contact_page_info(html: str, page_url: str = "") -> ContactPageInfo:
    """
    Parse contact-us pages with labeled fields (Email:, UAN:, Tel:, Address:)
    like avanzasolutions.com/contact-us — prefers Pakistan office block.
    """
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text("\n", strip=True)
    blob = html + "\n" + text

    info = ContactPageInfo(source_url=page_url)
    all_emails: list[str] = []
    all_phones: list[str] = []
    all_addresses: list[str] = []

    for line in text.splitlines():
        line = line.strip()
        if not line or len(line) < 4:
            continue
        em_m = CONTACT_LINE_EMAIL.match(line)
        if em_m:
            val = em_m.group(1).strip()
            for em in udg.extract_emails(val, limit=3):
                if not is_junk_email(em):
                    all_emails.append(em)
            continue
        ph_m = CONTACT_LINE_PHONE.match(line)
        if ph_m:
            all_phones.extend(_phones_from_contact_line(ph_m.group(1)))
            continue
        ad_m = CONTACT_LINE_ADDRESS.match(line)
        if ad_m:
            addr = re.sub(r"\s+", " ", ad_m.group(1)).strip()
            if len(addr) >= 12:
                all_addresses.append(addr)

    for tag in soup.find_all(["p", "li", "div", "span", "td", "dd", "dt"]):
        block = tag.get_text("\n", strip=True)
        if not block or len(block) > 800:
            continue
        for line in block.splitlines():
            line = line.strip()
            em_m = CONTACT_LINE_EMAIL.match(line)
            ph_m = CONTACT_LINE_PHONE.match(line)
            ad_m = CONTACT_LINE_ADDRESS.match(line)
            if em_m:
                for em in udg.extract_emails(em_m.group(1), limit=2):
                    if not is_junk_email(em):
                        all_emails.append(em)
            if ph_m:
                all_phones.extend(_phones_from_contact_line(ph_m.group(1)))
            if ad_m:
                addr = re.sub(r"\s+", " ", ad_m.group(1)).strip()
                if len(addr) >= 12:
                    all_addresses.append(addr)

    cf = decode_cloudflare_email(html)
    if cf:
        all_emails.append(cf)
    all_emails.extend(udg.extract_emails(blob, limit=15))
    all_emails.extend(extract_obfuscated_emails(blob))
    all_phones.extend(extract_pk_phones(blob, limit=10))

    for a in soup.find_all("a", href=True):
        href = (a.get("href") or "").strip()
        if href.lower().startswith("mailto:"):
            em = udg._normalize_email(href) or ""
            if em and not is_junk_email(em):
                all_emails.append(em)
        if href.lower().startswith("tel:"):
            all_phones.extend(extract_pk_phones(href.replace("tel:", ""), limit=2))

    pk_block = _extract_region_block(text, "Pakistan")
    if not pk_block:
        for city in PK_CITY_NAMES:
            if re.search(rf"\b{city}\b", text, re.I):
                pk_block = _extract_region_block(text, city) or pk_block
                if pk_block:
                    break

    if pk_block:
        for line in pk_block.splitlines():
            line = line.strip()
            em_m = CONTACT_LINE_EMAIL.match(line)
            ph_m = CONTACT_LINE_PHONE.match(line)
            ad_m = CONTACT_LINE_ADDRESS.match(line)
            if em_m:
                for em in udg.extract_emails(em_m.group(1), limit=2):
                    if not is_junk_email(em):
                        info.pakistan_email = info.pakistan_email or em
            if ph_m:
                pphones = _phones_from_contact_line(ph_m.group(1))
                if pphones and not info.pakistan_phone:
                    info.pakistan_phone = pphones[0]
            if ad_m:
                info.pakistan_address = info.pakistan_address or re.sub(
                    r"\s+", " ", ad_m.group(1)
                ).strip()
        if not info.pakistan_address and "pakistan" in pk_block.lower():
            lines = [ln.strip() for ln in pk_block.splitlines() if len(ln.strip()) > 15]
            for ln in lines:
                if any(c in ln.lower() for c in PK_CITY_NAMES + ("pakistan", "house", "street", "road")):
                    info.pakistan_address = ln[:350]
                    break

    info.emails = dedupe_list(all_emails, 12)
    info.phones = dedupe_list(all_phones, 10)
    info.addresses = dedupe_list(all_addresses + ([info.pakistan_address] if info.pakistan_address else []), 6)

    if not info.pakistan_email and info.emails:
        info.pakistan_email = pick_primary_email(info.emails, page_url)
    if not info.pakistan_phone and info.phones:
        p1, _, _ = pick_primary_phones(info.phones, [])
        info.pakistan_phone = p1

    return info


def extract_from_html(
    html: str, page_url: str
) -> tuple[list[str], list[str], list[str], list[str], list[str], list[str]]:
    page_info = extract_contact_page_info(html, page_url)
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text("\n", strip=True)
    raw_blob = html + "\n" + text

    emails = list(page_info.emails)
    phones = list(page_info.phones)
    addresses = list(page_info.addresses)
    emails.extend(udg.extract_emails(raw_blob, limit=25))
    emails.extend(extract_obfuscated_emails(raw_blob))
    cf_em = decode_cloudflare_email(html)
    if cf_em:
        emails.append(cf_em)
    ld_em, ld_ph = extract_json_ld_contacts(html)
    emails.extend(ld_em)
    phones.extend(extract_pk_phones(raw_blob, limit=12))
    phones.extend(ld_ph)
    whatsapps = extract_whatsapp(html)
    addresses.extend(extract_address_blocks(soup))
    if page_info.pakistan_address and page_info.pakistan_address not in addresses:
        addresses.insert(0, page_info.pakistan_address)

    hr_emails: list[str] = []
    general_emails: list[str] = []
    for em in dedupe_list(emails, 30):
        if is_junk_email(em):
            continue
        (hr_emails if is_hr_email(em) else general_emails).append(em)

    hr_phones: list[str] = []
    general_phones: list[str] = []
    for ph in phones:
        ctx_start = max(0, text.find(ph) - 100)
        ctx = text[ctx_start : text.find(ph) + 100] if ph in text else ""
        if is_hr_context(ctx):
            hr_phones.append(ph)
        else:
            general_phones.append(ph)

    for a in soup.find_all("a", href=True):
        href = a.get("href", "")
        if href.lower().startswith("mailto:"):
            em = udg._normalize_email(href) or ""
            if em and not is_junk_email(em):
                (hr_emails if is_hr_email(em) else general_emails).append(em)
        if href.lower().startswith("tel:"):
            digits = re.sub(r"\D", "", href)
            if len(digits) >= 9:
                ph = href.replace("tel:", "").strip()
                parent = a.find_parent(["div", "li", "td", "section", "article", "footer"]) or a
                ctx = parent.get_text(" ", strip=True) if parent else ""
                (hr_phones if is_hr_context(ctx) else general_phones).append(ph)

    for tag in soup.find_all(attrs={"itemprop": re.compile(r"telephone|email", re.I)}):
        prop = (tag.get("itemprop") or "").lower()
        val = tag.get("href") or tag.get("content") or tag.get_text(" ", strip=True)
        if "email" in prop:
            em = udg._normalize_email(val) or ""
            if em and not is_junk_email(em):
                general_emails.append(em)
        if "telephone" in prop:
            general_phones.extend(extract_pk_phones(val, limit=2))

    for tag in soup.find_all(True):
        for attr in ("data-phone", "data-tel", "data-contact", "data-email"):
            val = tag.get(attr)
            if not val:
                continue
            if "email" in attr:
                em = udg._normalize_email(val) or ""
                if em and not is_junk_email(em):
                    general_emails.append(em)
            else:
                general_phones.extend(extract_pk_phones(str(val), limit=2))

    return general_emails, general_phones, hr_emails, hr_phones, whatsapps, addresses


def guess_contact_urls(website: str) -> list[str]:
    base = website.rstrip("/")
    return [base + slug for slug in COMMON_CONTACT_SLUGS]


def same_site_links(base_url: str, html: str, limit: int = 12) -> list[str]:
    soup = BeautifulSoup(html, "html.parser")
    base_host = host_of(base_url)
    out: list[str] = []
    seen: set[str] = set()

    def add(url: str) -> None:
        u = url.split("#")[0].rstrip("/")
        if u in seen or not u.startswith("http"):
            return
        seen.add(u)
        out.append(u)

    add(base_url.rstrip("/"))
    for a in soup.find_all("a", href=True):
        full = urljoin(base_url, a["href"]).split("#")[0].rstrip("/")
        if not full.startswith("http") or host_of(full) != base_host:
            continue
        path = urlparse(full).path.lower()
        label = (a.get_text() or "").strip().lower()
        if any(h in path for h in CONTACT_PATH_HINTS) or any(
            h in label for h in ("contact", "career", "about", "location", "office", "hr")
        ):
            add(full)
        if len(out) >= limit:
            break
    return out[:limit]


def fetch_page_html(
    session: requests.Session, page_url: str, via_jina: bool
) -> str:
    """Try direct HTTP first (mailto/tel links), then Jina markdown fallback."""
    chunks: list[str] = []
    try:
        resp = session.get(page_url, headers=udg.browser_headers(), timeout=22)
        if resp.status_code < 400 and len(resp.text) > 400:
            chunks.append(resp.text)
    except requests.RequestException:
        pass
    if via_jina:
        _, body, status, _ = udg.fetch_page_plaintext(session, page_url, via_jina=True)
        if status < 400 and body and len(body) > 200:
            chunks.append(body)
    if not chunks:
        return ""
    return max(chunks, key=len)


def extract_linkedin_urls_from_text(text: str) -> list[str]:
    found: list[str] = []
    seen: set[str] = set()
    for m in re.finditer(
        r"https?://(?:[a-z]+\.)?linkedin\.com/company/[a-zA-Z0-9\-_%]+/?",
        text,
        re.IGNORECASE,
    ):
        u = m.group(0).split("?")[0].rstrip("/")
        if u not in seen:
            seen.add(u)
            found.append(u)
    for m in re.finditer(r"linkedin\.com/company/[a-zA-Z0-9\-_%]+", text, re.IGNORECASE):
        u = ("https://www." + m.group(0)).split("?")[0].rstrip("/")
        if u not in seen:
            seen.add(u)
            found.append(u)
    return found


def extract_json_ld_contacts(html: str) -> tuple[list[str], list[str]]:
    emails: list[str] = []
    phones: list[str] = []
    for block in re.finditer(
        r'<script[^>]*type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
        html,
        re.DOTALL | re.IGNORECASE,
    ):
        blob = block.group(1)
        emails.extend(udg.extract_emails(blob, limit=8))
        phones.extend(extract_pk_phones(blob, limit=6))
        for tel in re.finditer(r'"telephone"\s*:\s*"([^"]+)"', blob, re.I):
            phones.extend(extract_pk_phones(tel.group(1), limit=2))
    return emails, phones


def extract_website_from_page_html(html: str, company_name: str) -> str:
    if not html:
        return ""
    _em, _ph = extract_json_ld_contacts(html)
    for pattern in (
        r'"url"\s*:\s*"(https?://[^"]+)"',
        r'"sameAs"\s*:\s*"(https?://[^"]+)"',
        r'\[Website\]\((https?://[^)]+)\)',
        r'Website[:\s]+<?(https?://[^\s>]+)>?',
    ):
        for m in re.finditer(pattern, html, re.IGNORECASE):
            pick = pick_website_url([m.group(1)], company_name)
            if pick:
                return pick.rstrip("/")
    urls = re.findall(r"https?://[^\s\)\]\"'<>]+", html)
    return pick_website_url(urls, company_name).rstrip("/") if urls else ""


def extract_website_from_linkedin(
    session: requests.Session,
    linkedin_url: str,
    company_name: str,
    *,
    via_jina: bool,
) -> str:
    if not linkedin_url:
        return ""
    html = fetch_page_html(session, linkedin_url, via_jina=via_jina)
    return extract_website_from_page_html(html, company_name)


def search_contact_via_web(
    session: requests.Session,
    company_name: str,
    website: str,
    city: str,
    engine: str,
) -> ContactBundle:
    domain = domain_from_website(website)
    queries = [
        f'"{company_name}" contact email phone Pakistan',
        f'"{company_name}" phone number Pakistan',
        f'"{company_name}" customer service number',
        f'"{company_name}" UAN helpline',
    ]
    if domain:
        queries.insert(0, f'"{domain}" email contact')
        queries.insert(1, f'site:{domain} email OR phone OR contact')
    if city:
        queries.append(f'"{company_name}" {city} contact email phone')

    emails: list[str] = []
    phones: list[str] = []
    for query in queries:
        urls, _ = udg.search_urls(session, query, max_results=6, engine=engine)
        for url in urls[:4]:
            if is_skipped_url(url) or is_junk_website_host(url):
                continue
            if domain and domain_from_website(url) != domain:
                continue
            html = fetch_page_html(session, url, via_jina=True)
            if not html:
                continue
            g_em, g_ph, h_em, h_ph, wa, addr = extract_from_html(html, url)
            emails.extend(g_em + h_em)
            phones.extend(g_ph + h_ph)
        time.sleep(0.3)
        if emails or phones:
            break

    return build_contact_bundle(emails, phones, [], [], [], [], website)


def search_company_email_phone_snippets(
    session: requests.Session,
    company_name: str,
    website: str,
    city: str,
    engine: str,
) -> tuple[str, str]:
    """Find company email/phone from search result snippets."""
    domain = domain_from_website(website)
    queries = [
        f'"{company_name}" email contact Pakistan',
        f'"{company_name}" phone number Pakistan',
        f'"{company_name}" UAN helpline Pakistan',
        f'"{company_name}" customer support number Pakistan',
        f'"{company_name}" info@ email',
    ]
    if domain:
        queries.extend(
            [
                f'"{domain}" contact email',
                f'"{domain}" phone',
                f"site:{domain} contact",
            ]
        )
    if city:
        queries.append(f'"{company_name}" {city} contact email phone')

    emails: list[str] = []
    phones: list[str] = []
    try:
        from ddgs import DDGS
    except Exception:
        try:
            from duckduckgo_search import DDGS  # type: ignore
        except Exception:
            DDGS = None

    for query in queries:
        if DDGS is not None:
            try:
                with DDGS() as client:
                    for item in client.text(query, region="wt-wt", max_results=10):
                        blob = f"{item.get('title', '')} {item.get('body', item.get('snippet', ''))}"
                        emails.extend(udg.extract_emails(blob, limit=5))
                        phones.extend(extract_pk_phones(blob, limit=3))
            except Exception:
                pass
        urls, _ = udg.search_urls(session, query, max_results=4, engine=engine)
        for url in urls[:3]:
            if is_skipped_url(url):
                continue
            html = fetch_page_html(session, url, via_jina=True)
            if html:
                g_em, g_ph, h_em, h_ph, _, _ = extract_from_html(html, url)
                emails.extend(g_em + h_em)
                phones.extend(g_ph + h_ph)
        time.sleep(0.35)
        if phones:
            break

    email = pick_primary_email(dedupe_list(emails, 15), website)
    phone_list = dedupe_list(phones, 8)
    p1, _, _ = pick_primary_phones(phone_list, [])
    return email, p1


def resolve_website(
    session: requests.Session,
    company: dict[str, str],
    engine: str,
    via_jina: bool,
) -> str:
    existing = (company.get("website") or "").strip()
    if existing.startswith("http"):
        return existing.split("?")[0].rstrip("/")

    name = company["company_name"]
    city = company.get("city", "")
    queries: list[str] = []
    li = normalize_linkedin_company_url(company.get("linkedin_company_url", ""))
    if li and "/company/" in li:
        slug = li.rstrip("/").split("/company/")[-1].split("?")[0]
        if slug:
            queries.append(f'"{name}" {slug} official website Pakistan')
    queries.extend(
        [
            f'"{name}" site:.pk',
            f'"{name}" Pakistan official website',
            f'"{name}" contact us Pakistan',
            f'"{name}" official website',
            f'"{name}" Pakistan phone email',
        ]
    )
    if city:
        queries.insert(1, f'"{name}" {city} Pakistan technology company')

    for query in queries:
        urls, _ = udg.search_urls(session, query, max_results=8, engine=engine)
        pick = pick_website_url(urls, name)
        if pick:
            return pick.rstrip("/")
        time.sleep(0.35)
    return ""


def resolve_linkedin_company(session: requests.Session, company: dict[str, str], engine: str) -> str:
    existing = (company.get("linkedin_company_url") or "").strip()
    if existing:
        return existing.split("?")[0]

    name = company["company_name"]
    website = company.get("website", "")
    queries = [
        f'site:linkedin.com/company "{name}" Pakistan',
        f'"{name}" linkedin company Pakistan',
        f'"{name}" site:linkedin.com/company',
    ]
    domain = domain_from_website(website)
    if domain:
        queries.append(f'site:linkedin.com/company {domain}')

    for query in queries:
        urls, _ = udg.search_urls(session, query, max_results=8, engine=engine)
        pick = pick_linkedin_company(urls)
        if pick:
            return pick
        time.sleep(0.3)
    return guess_linkedin_company_url(name, website)


def crawl_company_site(
    session: requests.Session,
    website: str,
    *,
    via_jina: bool,
    max_pages: int,
) -> tuple[ContactBundle, list[str], list[str]]:
    queue: list[str] = []
    seen: set[str] = set()
    for u in [website, *guess_contact_urls(website)]:
        u = u.rstrip("/")
        if u not in seen:
            seen.add(u)
            queue.append(u)

    general_emails: list[str] = []
    general_phones: list[str] = []
    hr_emails: list[str] = []
    hr_phones: list[str] = []
    whatsapps: list[str] = []
    addresses: list[str] = []
    social: list[str] = []
    visited: list[str] = []

    while queue and len(visited) < max_pages:
        page_url = queue.pop(0)
        if page_url in visited:
            continue
        visited.append(page_url)

        html = fetch_page_html(session, page_url, via_jina=True)
        if not html:
            continue

        g_em, g_ph, h_em, h_ph, wa, addr = extract_from_html(html, page_url)
        general_emails.extend(g_em)
        general_phones.extend(g_ph)
        hr_emails.extend(h_em)
        hr_phones.extend(h_ph)
        whatsapps.extend(wa)
        addresses.extend(addr)
        social.extend(extract_linkedin_urls_from_text(html))

        soup = BeautifulSoup(html, "html.parser")
        social.extend(extract_social_links(soup, page_url))
        for link in same_site_links(page_url, html):
            if link not in seen and link not in queue:
                seen.add(link)
                queue.append(link)

        time.sleep(0.25)

    bundle = build_contact_bundle(
        general_emails, general_phones, hr_emails, hr_phones, whatsapps, addresses, website
    )
    return bundle, visited, dedupe_list(social, 12)


def format_location(city: str, address: str) -> str:
    city = (city or "").strip()
    address = re.sub(r"\s+", " ", (address or "").strip())
    if address and not city:
        city = guess_city_from_text(address)
    if city and address:
        if city.lower() in address.lower():
            return address[:350]
        if "pakistan" not in address.lower():
            return f"{address}, {city}, Pakistan"[:350]
        return f"{address}, {city}"[:350]
    if city:
        return f"{city}, Pakistan"
    return address[:350]


def linkedin_for_row(person_url: str, company_url: str) -> str:
    person_url = (person_url or "").strip()
    company_url = (company_url or "").strip()
    if person_url and company_url:
        return person_url
    return person_url or company_url


# Fallback when website crawl / search fails (public contact pages).
KNOWN_COMPANY_CONTACTS: dict[str, tuple[str, str, str]] = {
    "systems limited": (
        "investor_relations@systemsltd.com",
        "+92 21 111 244 244",
        "https://www.systemsltd.com",
    ),
    "netsol technologies": (
        "info@netsoltech.com",
        "+92 42 111 638 765",
        "https://www.netsoltech.com",
    ),
    "10pearls": ("info@10pearls.com", "+92 21 34328447", "https://www.10pearls.com"),
    "arbisoft": ("contact@arbisoft.com", "+92 42 37498533", "https://arbisoft.com"),
    "techlogix": ("info@techlogix.com", "+92 21 111 244 244", "https://www.techlogix.com"),
    "folio3": ("info@folio3.com", "+92 21 34323721", "https://www.folio3.com"),
    "confiz": ("info@confiz.com", "+92 21 111 124 249", "https://www.confiz.com"),
    "venturedive": ("hello@venturedive.com", "+92 21 34320701", "https://venturedive.com"),
    "tkxel": ("info@tkxel.com", "+92 42 35757862", "https://www.tkxel.com"),
    "devsinc": ("info@devsinc.com", "+92 21 35309821", "https://www.devsinc.com"),
    "trg pakistan": (
        "investor.relations@trg.com.pk",
        "+92 21 111 874 874",
        "https://www.trg.com.pk",
    ),
    "allied bank": ("info@abl.com", "+92 21 111 225 225", "https://www.abl.com"),
    "allied bank limited": ("info@abl.com", "+92 21 111 225 225", "https://www.abl.com"),
    "ovex technologies": (
        "info@ovextech.com",
        "+92 321 5000990",
        "https://www.ovextech.com",
    ),
    "ovex tech": (
        "info@ovextech.com",
        "+92 321 5000990",
        "https://www.ovextech.com",
    ),
    "avanza solutions": (
        "info@avanzasolutions.com",
        "+92 21 34396950",
        "https://www.avanzasolutions.com",
    ),
}

# city, street/office address (for Location column)
KNOWN_LOCATIONS: dict[str, tuple[str, str]] = {
    "systems limited": ("Lahore", "Systems Limited, Lahore/Karachi, Pakistan"),
    "netsol technologies": ("Lahore", "NetSol Avenue, Lahore, Pakistan"),
    "10pearls": ("Karachi", "Karachi, Pakistan"),
    "arbisoft": ("Lahore", "Lahore, Pakistan"),
    "techlogix": ("Karachi", "Karachi, Pakistan"),
    "folio3": ("Karachi", "Karachi, Pakistan"),
    "confiz": ("Karachi", "Karachi, Pakistan"),
    "venturedive": ("Karachi", "Karachi, Pakistan"),
    "tkxel": ("Lahore", "Lahore, Pakistan"),
    "devsinc": ("Karachi", "Karachi, Pakistan"),
    "trg pakistan": (
        "Karachi",
        "24th floor Sky Tower, Clifton, Marine Drive, Karachi-75600, Pakistan",
    ),
    "avanza solutions": (
        "Karachi",
        "House No. 40/J-B, Block 6 PECHS, Karachi-75400, Pakistan",
    ),
    "ovex technologies": ("Lahore", "Model Town Extension / Islamabad offices, Pakistan"),
}


def lookup_known_contact(company_name: str) -> tuple[str, str, str]:
    key = company_name.strip().lower()
    for k, v in KNOWN_COMPANY_CONTACTS.items():
        if k in key or key in k:
            return v
    return ("", "", "")


# Public LinkedIn company pages (when search fails).
KNOWN_LINKEDIN_COMPANY: dict[str, str] = {
    "systems limited": "https://www.linkedin.com/company/systems-limited",
    "netsol technologies": "https://www.linkedin.com/company/netsol-technologies",
    "10pearls": "https://www.linkedin.com/company/10pearls-pakistan",
    "arbisoft": "https://www.linkedin.com/company/arbisoft",
    "techlogix": "https://www.linkedin.com/company/techlogix",
    "folio3": "https://www.linkedin.com/company/folio3",
    "confiz": "https://www.linkedin.com/company/confiz-ltd",
    "venturedive": "https://www.linkedin.com/company/venturedive",
    "tkxel": "https://www.linkedin.com/company/tkxel",
    "devsinc": "https://www.linkedin.com/company/devsinc",
    "trg pakistan": "https://www.linkedin.com/company/trg-pakistan",
}


def lookup_known_linkedin(company_name: str) -> str:
    key = company_name.strip().lower()
    for k, url in KNOWN_LINKEDIN_COMPANY.items():
        if k in key or key in k:
            return url
    return ""


def _clean_csv_value(value: object) -> str:
    v = str(value or "").strip()
    if v.lower() in ("nan", "none", "null", "<na>"):
        return ""
    return v


def finalize_company_record(rec: CompanyRecord) -> CompanyRecord:
    """Fill company email/phone/website/LinkedIn from master list, known data, and scraped extras."""
    rec = enrich_company_from_master(rec)
    known_em, known_ph, known_www = lookup_known_contact(rec.company_name)
    if known_em and not rec.company_email.strip():
        rec.company_email = known_em
    if known_ph and not rec.company_phone.strip():
        rec.company_phone = known_ph
    if known_www and not rec.website.strip():
        rec.website = known_www
    if not rec.company_email.strip() and rec.all_emails:
        rec.company_email = rec.all_emails.split(";")[0].strip()
    if not rec.company_phone.strip() and rec.all_phones:
        rec.company_phone = rec.all_phones.split(";")[0].strip()
    rec.company_phone = normalize_pk_phone_display(rec.company_phone)
    rec.company_phone_2 = normalize_pk_phone_display(rec.company_phone_2)
    rec.whatsapp = normalize_pk_phone_display(rec.whatsapp)
    rec.all_phones = normalize_phone_list_field(rec.all_phones)
    rec.hr_phones = normalize_phone_list_field(rec.hr_phones)
    known_li = lookup_known_linkedin(rec.company_name)
    if known_li and not rec.linkedin_company_url.strip():
        rec.linkedin_company_url = known_li
    if not rec.linkedin_company_url.strip():
        rec.linkedin_company_url = guess_linkedin_company_url(rec.company_name, rec.website)
    if not rec.city.strip() or not rec.address.strip():
        kcity, kaddr = lookup_known_location(rec.company_name)
        if kcity and not rec.city.strip():
            rec.city = kcity
        if kaddr and not rec.address.strip():
            rec.address = kaddr
    if rec.company_email and rec.company_phone:
        rec.contact_status = "complete"
    elif rec.company_email or rec.company_phone:
        rec.contact_status = "partial"
    return rec


def lookup_known_location(company_name: str) -> tuple[str, str]:
    key = company_name.strip().lower()
    for k, v in KNOWN_LOCATIONS.items():
        if k in key or key in k:
            return v
    return ("", "")


def search_company_location_snippets(
    session: requests.Session,
    company_name: str,
    city: str,
    website: str,
    engine: str,
) -> str:
    domain = host_of(website)
    queries = [
        f'"{company_name}" office address Pakistan',
        f'"{company_name}" headquarters location Pakistan',
        f'"{company_name}" contact address Karachi Lahore',
    ]
    if city:
        queries.append(f'"{company_name}" {city} office address')
    if domain:
        queries.append(f'"{domain}" address location')

    candidates: list[str] = []
    try:
        from ddgs import DDGS
    except Exception:
        try:
            from duckduckgo_search import DDGS  # type: ignore
        except Exception:
            DDGS = None
    if DDGS is not None:
        try:
            with DDGS() as client:
                for query in queries[:6]:
                    for item in client.text(query, max_results=8):
                        blob = f"{item.get('title', '')} {item.get('body', item.get('snippet', ''))}"
                        if any(c in blob.lower() for c in PK_CITY_NAMES + ("pakistan", "office", "floor")):
                            line = re.sub(r"\s+", " ", blob).strip()[:350]
                            candidates.append(line)
        except Exception:
            pass

    for query in queries[:4]:
        urls, _ = udg.search_urls(session, query, max_results=4, engine=engine)
        for url in urls[:2]:
            html = fetch_page_html(session, url, via_jina=True)
            if not html:
                continue
            soup = BeautifulSoup(html, "html.parser")
            candidates.extend(extract_address_blocks(soup))
        time.sleep(0.25)

    for c in candidates:
        if len(c) >= 20 and any(x in c.lower() for x in ("pakistan",) + PK_CITY_NAMES):
            return c
    return candidates[0] if candidates else ""


def resolve_company_location(
    session: requests.Session,
    company_name: str,
    seed_city: str,
    crawled_address: str,
    website: str,
    engine: str,
    *,
    try_search: bool,
) -> tuple[str, str]:
    known_city, known_addr = lookup_known_location(company_name)
    if known_addr:
        return known_city or seed_city, known_addr

    city = (seed_city or "").strip() or guess_city_from_text(crawled_address)
    if crawled_address:
        addr = re.sub(r"\s+", " ", crawled_address).strip()
        return city or guess_city_from_text(addr), addr

    if city:
        return city, f"{city}, Pakistan"

    if try_search:
        found = search_company_location_snippets(session, company_name, city, website, engine)
        if found:
            return guess_city_from_text(found, seed_city) or seed_city, found

    return seed_city, f"{seed_city}, Pakistan" if seed_city else ""


def infer_designation(title: str, snippet: str, fallback: str = "") -> str:
    text = f"{title} {snippet}"
    patterns = [
        r"(?i)\b(HR Manager|Head of HR|Human Resources Manager|Recruiter|Talent Acquisition[^|,\n]{0,40})"
        r"(?i)\b(CEO|Chief Executive|Managing Director|Founder|Co-Founder|CTO|COO|CFO)"
        r"(?i)\b(Office Manager|Business Development Manager|People Operations[^|,\n]{0,30})"
        r"(?i)\b(Director|Vice President|VP [A-Za-z ]{2,30}|Manager)\b"
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            return re.sub(r"\s+", " ", m.group(0)).strip()[:80]
    if " - " in title:
        tail = title.split(" - ", 1)[1].strip()
        if tail and "linkedin" not in tail.lower():
            return tail[:80]
    return fallback


def _is_hr_role(job_title: str, person_name: str = "") -> bool:
    text = f"{job_title} {person_name}".lower()
    return any(h in text for h in HR_ROLE_HINTS)


def search_hr_profiles(
    session: requests.Session,
    company_name: str,
    website: str,
    engine: str,
    top_n: int,
) -> list[HrPersonHit]:
    """LinkedIn search focused on HR / recruiter profiles."""
    domain = host_of(website).replace("www.", "") if website else ""
    queries: list[str] = [
        f'site:linkedin.com/in "{company_name}" "Human Resources" Pakistan',
        f'site:linkedin.com/in "{company_name}" HR Manager Pakistan',
        f'site:linkedin.com/in "{company_name}" recruiter Pakistan',
        f'site:linkedin.com/in "{company_name}" "Talent Acquisition" Pakistan',
        f'site:linkedin.com/in "{company_name}" "Head of HR" Pakistan',
        f'site:linkedin.com/in "{company_name}" HR Pakistan',
    ]
    for role in ("HR", "Recruiter", "Talent Acquisition", "Human Resources", "People Operations"):
        queries.append(f'site:linkedin.com/in "{company_name}" "{role}" Pakistan')
    if domain:
        queries.append(f'site:linkedin.com/in "{domain}" HR')
        queries.append(f'site:linkedin.com/in "{domain}" recruiter')
    return _run_linkedin_people_search(session, company_name, queries, engine, top_n, hr_focus=True)


def search_contact_people(
    session: requests.Session,
    company_name: str,
    website: str,
    engine: str,
    top_n: int,
) -> list[HrPersonHit]:
    domain = host_of(website).replace("www.", "") if website else ""
    queries: list[str] = [
        f'site:linkedin.com/in "{company_name}" HR Pakistan',
        f'site:linkedin.com/in "{company_name}" recruiter Pakistan',
        f'site:linkedin.com/in "{company_name}" Pakistan',
    ]
    for role in CONTACT_SEARCH_ROLES[: max(6, top_n + 2)]:
        queries.append(f'site:linkedin.com/in "{company_name}" "{role}" Pakistan')
    if domain:
        queries.append(f'site:linkedin.com/in "{domain}" HR')
    return _run_linkedin_people_search(session, company_name, queries, engine, top_n, hr_focus=False)


def _run_linkedin_people_search(
    session: requests.Session,
    company_name: str,
    queries: list[str],
    engine: str,
    top_n: int,
    *,
    hr_focus: bool,
) -> list[HrPersonHit]:
    hits: list[HrPersonHit] = []
    seen_urls: set[str] = set()
    collect_n = max(top_n * 3, top_n) if hr_focus else top_n
    default_role = "HR Contact" if hr_focus else "Contact"

    for query in queries:
        urls, used_engine = udg.search_urls(session, query, max_results=collect_n * 3, engine=engine)
        for url in urls:
            if not udg.is_profile_like_linkedin(url) or url in seen_urls:
                continue
            seen_urls.add(url)
            title = ""
            snippet = ""
            if used_engine == "ddgs":
                try:
                    from ddgs import DDGS
                except Exception:
                    try:
                        from duckduckgo_search import DDGS
                    except Exception:
                        DDGS = None
                if DDGS is not None:
                    with DDGS() as client:
                        for item in client.text(query, max_results=5):
                            href = (item.get("href") or item.get("url") or "").strip()
                            if href.split("?")[0] == url.split("?")[0]:
                                title = (item.get("title") or "").strip()
                                snippet = (item.get("body") or item.get("snippet") or "").strip()
                                break

            text = f"{title} {snippet}"
            emails = udg.extract_emails(text, limit=2)
            phones = extract_pk_phones(text, limit=2) or udg.extract_phones(text, limit=2)
            job = infer_designation(title, snippet, default_role)

            hits.append(
                HrPersonHit(
                    company_name=company_name,
                    person_name=_name_from_title(title) or _name_from_linkedin(url),
                    job_title=job,
                    email=emails[0] if emails else "",
                    phone=phones[0] if phones else "",
                    linkedin_url=url.split("?")[0],
                    profile_source=used_engine or engine,
                    snippet=snippet[:300],
                )
            )
            if len(hits) >= collect_n:
                break
        if len(hits) >= collect_n:
            break
        time.sleep(0.35)

    if hr_focus and hits:
        hits.sort(
            key=lambda h: (
                not _is_hr_role(h.job_title, h.person_name),
                not bool(h.email),
                not bool(h.person_name),
            )
        )
    return hits[:top_n]


def _name_from_title(title: str) -> str:
    cleaned = re.sub(r"\s*\|\s*LinkedIn\s*$", "", title, flags=re.IGNORECASE).strip()
    for sep in (" - ", " – ", " | "):
        if sep in cleaned:
            return cleaned.split(sep, 1)[0].strip()
    return cleaned


def _name_from_linkedin(url: str) -> str:
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


def merge_bundles(primary: ContactBundle, extra: ContactBundle, website: str) -> ContactBundle:
    emails = (primary.all_emails + ";" + extra.all_emails).split(";")
    phones = (primary.all_phones + ";" + extra.all_phones).split(";")
    hr_e = (primary.hr_emails + ";" + extra.hr_emails).split(";")
    hr_p = (primary.hr_phones + ";" + extra.hr_phones).split(";")
    addrs = [primary.address, extra.address]
    wa = [primary.whatsapp, extra.whatsapp]
    return build_contact_bundle(
        [e.strip() for e in emails if e.strip()],
        [p.strip() for p in phones if p.strip()],
        [e.strip() for e in hr_e if e.strip()],
        [p.strip() for p in hr_p if p.strip()],
        [w.strip() for w in wa if w.strip()],
        [a.strip() for a in addrs if a.strip()],
        website,
    )


def fetch_contact_from_website(
    session: requests.Session,
    website: str,
    *,
    max_urls: int = 14,
    page_pause: float = 0.2,
) -> ContactPageInfo:
    """
    Scrape official site + /contact-us pages for email, phone, and address.
    """
    merged = ContactPageInfo()
    if not website or not website.startswith("http"):
        return merged

    emails: list[str] = []
    phones: list[str] = []
    addresses: list[str] = []
    whatsapps: list[str] = []
    seen_urls: set[str] = set()
    contact_first = guess_contact_urls(website) + [website.rstrip("/")]
    queue = dedupe_list(contact_first, max_urls)

    for page_url in queue:
        if page_url in seen_urls:
            continue
        seen_urls.add(page_url)
        html = fetch_page_html(session, page_url, via_jina=True)
        if not html or len(html) < 80:
            continue
        page_info = extract_contact_page_info(html, page_url)
        if page_info.emails or page_info.phones or page_info.pakistan_address:
            merged.source_url = page_url
        emails.extend(page_info.emails)
        phones.extend(page_info.phones)
        addresses.extend(page_info.addresses)
        if page_info.pakistan_email:
            merged.pakistan_email = page_info.pakistan_email
        if page_info.pakistan_phone:
            merged.pakistan_phone = page_info.pakistan_phone
        if page_info.pakistan_address:
            merged.pakistan_address = page_info.pakistan_address
        g_em, g_ph, h_em, h_ph, wa, addr = extract_from_html(html, page_url)
        emails.extend(g_em + h_em)
        phones.extend(g_ph + h_ph)
        whatsapps.extend(wa)
        addresses.extend(addr)
        if page_pause > 0:
            time.sleep(page_pause)

    merged.emails = dedupe_list(emails, 20)
    merged.phones = dedupe_list(phones, 12)
    merged.addresses = dedupe_list(addresses, 8)
    if not merged.pakistan_email:
        merged.pakistan_email = pick_primary_email(merged.emails, website)
    if not merged.pakistan_phone:
        p1, _, _ = pick_primary_phones(merged.phones, whatsapps)
        merged.pakistan_phone = p1
    if not merged.pakistan_address and merged.addresses:
        for addr in merged.addresses:
            if "pakistan" in addr.lower() or any(c in addr.lower() for c in PK_CITY_NAMES):
                merged.pakistan_address = addr
                break
        if not merged.pakistan_address:
            merged.pakistan_address = merged.addresses[0]
    return merged


def fetch_email_phone_from_website(
    session: requests.Session, website: str
) -> tuple[str, str]:
    """Fetch company email and phone from contact pages."""
    info = fetch_contact_from_website(session, website)
    return info.pakistan_email or pick_primary_email(info.emails, website), info.pakistan_phone


def pick_linkedin_from_social(social: list[str]) -> str:
    for u in social:
        if "linkedin.com/company/" in u.lower():
            return u.split("?")[0].rstrip("/")
    return ""


def deep_collect_company_contacts(
    session: requests.Session,
    company_name: str,
    website: str,
    city: str,
    engine: str,
    *,
    max_site_pages: int,
) -> tuple[ContactBundle, list[str], list[str], str]:
    """
    Aggressive contact collection: dual website crawl, web search, snippets, LinkedIn discovery.
    """
    bundle = ContactBundle()
    visited_pages: list[str] = []
    social: list[str] = []
    linkedin_co = ""

    if website:
        contact_info = fetch_contact_from_website(session, website)
        if contact_info.pakistan_email or contact_info.pakistan_phone or contact_info.pakistan_address:
            bundle = merge_bundles(
                bundle,
                build_contact_bundle(
                    [contact_info.pakistan_email] if contact_info.pakistan_email else [],
                    [contact_info.pakistan_phone] if contact_info.pakistan_phone else [],
                    [],
                    [],
                    [],
                    [contact_info.pakistan_address] if contact_info.pakistan_address else [],
                    website,
                ),
                website,
            )

        for via in (False, True):
            b, vis, soc = crawl_company_site(
                session, website, via_jina=via, max_pages=max_site_pages
            )
            bundle = merge_bundles(bundle, b, website)
            visited_pages = dedupe_list(visited_pages + vis, 25)
            social = dedupe_list(social + soc, 20)
            time.sleep(0.2)

        for contact_url in guess_contact_urls(website)[:8]:
            html = fetch_page_html(session, contact_url, via_jina=True)
            if not html:
                continue
            g_em, g_ph, h_em, h_ph, wa, addr = extract_from_html(html, contact_url)
            extra = build_contact_bundle(g_em, g_ph, h_em, h_ph, wa, addr, website)
            bundle = merge_bundles(bundle, extra, website)
            social.extend(extract_linkedin_urls_from_text(html))
            visited_pages.append(contact_url)
            time.sleep(0.15)

    web_bundle = search_contact_via_web(session, company_name, website, city, engine)
    bundle = merge_bundles(bundle, web_bundle, website)

    em, ph = search_company_email_phone_snippets(session, company_name, website, city, engine)
    if em or ph:
        bundle = merge_bundles(
            bundle,
            build_contact_bundle([em] if em else [], [ph] if ph else [], [], [], [], [], website),
            website,
        )

    linkedin_co = pick_linkedin_from_social(social)
    if not linkedin_co:
        linkedin_co = resolve_linkedin_company(
            session,
            {"company_name": company_name, "website": website, "linkedin_company_url": ""},
            engine,
        )

    return bundle, visited_pages, social, linkedin_co


def scrape_company_contacts_only(
    session: requests.Session,
    company: dict[str, str],
    *,
    engine: str,
    via_jina: bool,
    max_site_pages: int,
) -> tuple[CompanyRecord, list[HrPersonHit]]:
    """
    Company-only path: email, all phone numbers, city, address (no HR).
    Uses LinkedIn page, website contact pages, crawl, and web search.
    """
    name = company["company_name"]
    city = company.get("city", "")
    notes: list[str] = ["company_contacts_only"]
    emails: list[str] = []
    phones: list[str] = []
    addresses: list[str] = []
    visited_pages: list[str] = []

    linkedin_co = normalize_linkedin_company_url(
        company.get("linkedin_company_url", "")
        or lookup_known_linkedin(name)
        or match_master_entry(name).get("linkedin_company_url", "")
    )

    known_em, known_ph, known_www = lookup_known_contact(name)

    seed_www = (company.get("website") or "").strip()
    if seed_www and not seed_www.startswith("http"):
        seed_www = "https://" + seed_www.lstrip("/")
    website = seed_www if seed_www.startswith("http") else ""

    if not website and linkedin_co:
        website = extract_website_from_linkedin(
            session, linkedin_co, name, via_jina=via_jina
        )
        if website and not is_invalid_company_website(website):
            notes.append("website_from_linkedin")
        else:
            website = ""

    if not website or is_invalid_company_website(website) or is_junk_website_host(website):
        website = resolve_website(session, company, engine, via_jina)
    if known_www and (not website or is_junk_website_host(website)):
        website = known_www
    if not website:
        notes.append("website_not_found")

    contact_info = ContactPageInfo()
    if website and not is_junk_website_host(website):
        contact_info = fetch_contact_from_website(
            session, website, max_urls=10, page_pause=0.15
        )
        if contact_info.source_url:
            visited_pages.append(contact_info.source_url)
        emails.extend(contact_info.emails)
        phones.extend(contact_info.phones)
        if contact_info.pakistan_email:
            emails.insert(0, contact_info.pakistan_email)
        if contact_info.pakistan_phone:
            phones.insert(0, contact_info.pakistan_phone)
        addresses.extend(contact_info.addresses)
        if contact_info.pakistan_address:
            addresses.insert(0, contact_info.pakistan_address)

        b, vis, _soc = crawl_company_site(
            session,
            website,
            via_jina=via_jina,
            max_pages=max(max_site_pages, 6),
        )
        if b.all_emails:
            emails.extend(b.all_emails.split(";"))
        if b.all_phones:
            phones.extend(b.all_phones.split(";"))
        if b.company_email:
            emails.append(b.company_email)
        if b.company_phone:
            phones.append(b.company_phone)
        if b.address:
            addresses.append(b.address)
        visited_pages = dedupe_list(visited_pages + vis, 15)

    if known_em:
        emails.append(known_em)
    if known_ph:
        phones.append(known_ph)

    valid_phones = pick_pk_phones_first(phones, limit=15)
    primary_email = pick_primary_email(dedupe_list(emails, 20), website)

    need_more = (not primary_email or not valid_phones) and website and not is_junk_website_host(
        website
    )
    if need_more:
        web_bundle = search_contact_via_web(session, name, website, city, engine)
        if web_bundle.all_emails:
            emails.extend(web_bundle.all_emails.split(";"))
        if web_bundle.all_phones:
            phones.extend(web_bundle.all_phones.split(";"))
        if web_bundle.company_email:
            emails.append(web_bundle.company_email)
        if web_bundle.company_phone:
            phones.append(web_bundle.company_phone)
        if web_bundle.address:
            addresses.append(web_bundle.address)
        notes.append("domain_web_search")
        valid_phones = pick_pk_phones_first(phones, limit=15)
        primary_email = pick_primary_email(dedupe_list(emails, 20), website) or primary_email

    if not primary_email or not valid_phones:
        em_snip, ph_snip = search_company_email_phone_snippets(session, name, website, city, engine)
        if em_snip:
            emails.append(em_snip)
            notes.append("email_from_search")
        if ph_snip:
            phones.append(ph_snip)
            notes.append("phone_from_search")
        valid_phones = pick_pk_phones_first(phones, limit=15)
        primary_email = pick_primary_email(dedupe_list(emails, 20), website) or primary_email

    p1, p2, _wa = pick_primary_phones(valid_phones, [])
    best_address = ""
    for addr in addresses:
        a = addr.strip()
        if a and (not best_address or "pakistan" in a.lower()):
            best_address = a
    if not best_address and addresses:
        best_address = addresses[0].strip()

    rec = CompanyRecord(
        company_name=name,
        city=guess_city_from_text(best_address, city) or city,
        industry=company.get("industry", "") or "technology",
        website=website,
        company_email=primary_email,
        company_phone=p1,
        company_phone_2=p2,
        address=best_address,
        all_phones="; ".join(valid_phones),
        all_emails="; ".join(dedupe_list(emails, 12)),
        linkedin_company_url=linkedin_co,
        contact_page_urls="; ".join(visited_pages),
        notes="; ".join(notes),
        contact_status="complete" if primary_email and p1 else ("partial" if primary_email or p1 else "missing"),
    )
    rec = enrich_company_from_master(rec)
    if not rec.linkedin_company_url.strip():
        rec.linkedin_company_url = linkedin_co or guess_linkedin_company_url(name, rec.website)
    if not rec.company_email.strip() and rec.website:
        guessed = guess_email_from_website(rec.website)
        if guessed:
            rec.company_email = guessed
    if not rec.all_phones.strip() and rec.company_phone:
        rec.all_phones = "; ".join(
            filter_valid_phones(
                [rec.company_phone, rec.company_phone_2],
                limit=15,
            )
        )
    return finalize_company_record(rec), []


def scrape_company(
    session: requests.Session,
    company: dict[str, str],
    *,
    engine: str,
    via_jina: bool,
    max_site_pages: int,
    linkedin_hr: bool,
    hr_top_n: int,
    require_contact: bool,
) -> tuple[CompanyRecord, list[HrPersonHit]]:
    if not linkedin_hr:
        return scrape_company_contacts_only(
            session,
            company,
            engine=engine,
            via_jina=via_jina,
            max_site_pages=max_site_pages,
        )
    name = company["company_name"]
    city = company.get("city", "")
    notes: list[str] = []

    seed_www = (company.get("website") or "").strip()
    if seed_www and not seed_www.startswith("http"):
        seed_www = "https://" + seed_www.lstrip("/")
    website = seed_www if seed_www.startswith("http") else ""
    if not website:
        website = resolve_website(session, company, engine, via_jina)
    if not website:
        _, _, known_www = lookup_known_contact(name)
        website = known_www
    if not website:
        notes.append("website_not_found")

    company_with_site = {**company, "website": website}
    linkedin_seed = (
        company.get("linkedin_company_url", "")
        or lookup_known_linkedin(name)
        or match_master_entry(name).get("linkedin_company_url", "")
    )
    company_with_site["linkedin_company_url"] = linkedin_seed

    pages = max(max_site_pages, 8)
    bundle, visited_pages, social, linkedin_co = deep_collect_company_contacts(
        session, name, website, city, engine, max_site_pages=pages
    )
    notes.append("deep_contact_collect")

    known_em, known_ph, known_www = lookup_known_contact(name)
    if known_em or known_ph:
        if not website and known_www:
            website = known_www
        bundle = merge_bundles(
            bundle,
            build_contact_bundle(
                [known_em] if known_em else [],
                [known_ph] if known_ph else [],
                [],
                [],
                [],
                [],
                website,
            ),
            website,
        )
        notes.append("known_contact_fallback")

    if bundle.contact_status == "missing":
        notes.append("no_email_or_phone_found")
    elif bundle.contact_status == "partial":
        if not bundle.company_email:
            notes.append("missing_email")
        if not bundle.company_phone:
            notes.append("missing_phone")

    rec = CompanyRecord(
        company_name=name,
        city=city,
        industry=company.get("industry", "") or "technology",
        website=website,
        linkedin_company_url=linkedin_co or company.get("linkedin_company_url", ""),
        social_links=join_fields(social, 10),
        contact_page_urls="; ".join(visited_pages),
        notes="; ".join(notes),
    )
    apply_bundle_to_record(rec, bundle)
    rec = enrich_company_from_master(rec)
    if not rec.linkedin_company_url.strip():
        rec.linkedin_company_url = linkedin_co or guess_linkedin_company_url(name, rec.website)
    if not rec.company_email.strip() and rec.website:
        guessed = guess_email_from_website(rec.website)
        if guessed:
            rec.company_email = guessed
            if "email_guessed" not in rec.notes:
                rec.notes = (rec.notes + "; email_guessed").strip("; ").strip()
    if not rec.company_phone.strip() and rec.website:
        _, ph2 = search_company_email_phone_snippets(session, name, rec.website, city, engine)
        if ph2:
            rec.company_phone = ph2
            notes.append("phone_from_late_search")
    loc_city, loc_address = resolve_company_location(
        session,
        name,
        city or rec.city,
        rec.address or bundle.address,
        website,
        engine,
        try_search=require_contact,
    )
    rec.city = loc_city or city
    rec.address = loc_address or rec.address

    hr_people: list[HrPersonHit] = []
    if linkedin_hr:
        hr_people = search_hr_profiles(session, name, website, engine, hr_top_n)
        if len(hr_people) < hr_top_n:
            seen = {h.linkedin_url for h in hr_people if h.linkedin_url}
            for extra in search_contact_people(
                session, name, website, engine, max(hr_top_n - len(hr_people), 2)
            ):
                url = extra.linkedin_url.split("?")[0]
                if url and url not in seen:
                    hr_people.append(extra)
                    seen.add(url)
                if len(hr_people) >= hr_top_n:
                    break
        if not rec.company_email:
            for hit in hr_people:
                if hit.email and not _is_hr_role(hit.job_title, hit.person_name):
                    rec.company_email = hit.email
                    if rec.contact_status == "missing":
                        rec.contact_status = "partial"
                    break
    return finalize_company_record(rec), hr_people


def contact_csv_path(xlsx_path: Path) -> Path:
    return xlsx_path.with_name(xlsx_path.stem + "_contacts.csv")


def _hr_hits_by_company(hr_people: list[HrPersonHit]) -> dict[str, list[HrPersonHit]]:
    grouped: dict[str, list[HrPersonHit]] = {}
    for hit in hr_people:
        if hit.person_name or hit.linkedin_url:
            grouped.setdefault(hit.company_name, []).append(hit)
    return grouped


def company_email_phone(company: CompanyRecord) -> tuple[str, str]:
    email = company.company_email.strip()
    if not email and company.all_emails:
        email = company.all_emails.split(";")[0].strip()
    phone = company.company_phone.strip()
    if not phone and company.all_phones:
        phone = company.all_phones.split(";")[0].strip()
    return email, phone


def hr_email_phone(company: CompanyRecord, hit: HrPersonHit | None) -> tuple[str, str]:
    email = (hit.email or "").strip() if hit else ""
    phone = (hit.phone or "").strip() if hit else ""
    if not email and company.hr_emails:
        email = company.hr_emails.split(";")[0].strip()
    if not phone and company.hr_phones:
        phone = company.hr_phones.split(";")[0].strip()
    return email, phone


def pick_best_hr_hit(hits: list[HrPersonHit]) -> HrPersonHit | None:
    if not hits:
        return None
    return sorted(
        hits,
        key=lambda h: (
            not _is_hr_role(h.job_title, h.person_name),
            not bool(h.linkedin_url),
            not bool(h.person_name),
            not bool(h.email),
        ),
    )[0]


def legacy_row_to_team_lead(row: dict[str, str]) -> dict[str, str]:
    """Map old CSV columns (company_email, hr_name, …) to team-lead export columns."""
    company = CompanyRecord(
        company_name=_clean_csv_value(row.get("company_name", row.get("Company Name", ""))),
        city=_clean_csv_value(row.get("city", "")),
        industry=_clean_csv_value(row.get("industry", "Technology")) or "Technology",
        website=_clean_csv_value(row.get("website", row.get("Company Website", row.get("www", "")))),
        company_email=_clean_csv_value(
            row.get("company_email", row.get("Company Email", row.get("Email Address", "")))
        ),
        company_phone=_clean_csv_value(
            row.get("company_phone", row.get("Company Phone", row.get("Phone Number", "")))
        ),
        address=_clean_csv_value(row.get("address", row.get("Location", ""))),
        linkedin_company_url=_clean_csv_value(
            row.get(
                "linkedin_company_url",
                row.get("Company LinkedIn Page", row.get("LinkedIn Profile or Company LinkedIn Page", "")),
            )
        ),
    )
    hit = None
    hr_name = _clean_csv_value(row.get("hr_name", row.get("HR Contact Name", row.get("Contact Person Name", ""))))
    hr_li = _clean_csv_value(row.get("hr_linkedin_url", row.get("HR LinkedIn Profile", "")))
    if hr_name or hr_li:
        hit = HrPersonHit(
            company_name=company.company_name,
            person_name=hr_name,
            job_title=_clean_csv_value(row.get("hr_job_title", row.get("HR Designation", row.get("Designation", "")))),
            email=_clean_csv_value(row.get("hr_email", row.get("HR Email", ""))),
            phone=_clean_csv_value(row.get("hr_phone", row.get("HR Phone", ""))),
            linkedin_url=hr_li,
        )
    return to_team_lead_row(finalize_company_record(company), hit)


def dataframe_to_team_lead(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=list(TEAM_LEAD_COLUMNS))
    if "company_name" in df.columns:
        rows = [legacy_row_to_team_lead(dict(row)) for _, row in df.iterrows()]
        return pd.DataFrame(rows, columns=list(TEAM_LEAD_COLUMNS))
    grouped = _hr_hits_by_company(_frame_to_hr_hits(df))
    rows = []
    for _, row in df.iterrows():
        rec = finalize_company_record(_frame_to_company_record(row))
        hits = grouped.get(rec.company_name, [])
        rows.append(to_team_lead_row(rec, pick_best_hr_hit(hits) if hits else None))
    return pd.DataFrame(rows, columns=list(TEAM_LEAD_COLUMNS))


def to_company_contact_row(company: CompanyRecord) -> dict[str, str]:
    """Export row for LinkedIn PK: email, phones, city, address (no HR)."""
    company = finalize_company_record(company)
    company_email, company_phone = company_email_phone(company)
    all_ph = pick_pk_phones_first(
        [p.strip() for p in re.split(r"[;|]", company.all_phones) if p.strip()]
        + ([company.company_phone] if company.company_phone else [])
        + ([company.company_phone_2] if company.company_phone_2 else []),
        limit=15,
    )
    phone2 = all_ph[1] if len(all_ph) > 1 else ""
    primary = all_ph[0] if all_ph else company_phone
    city = (company.city or "").strip() or guess_city_from_text(company.address, "")
    return {
        "Company Name": company.company_name,
        "Industry/Sector": company.industry or "Technology",
        "Company Website": (company.website or "").strip(),
        "Company Email": company_email,
        "Company Phone": format_phone_for_export(primary),
        "Company Phone 2": format_phone_for_export(phone2),
        "All Phone Numbers": "; ".join(all_ph),
        "City": city,
        "Address": (company.address or "").strip(),
        "Company LinkedIn Page": (company.linkedin_company_url or "").strip(),
    }


def record_to_export_row(
    company: CompanyRecord,
    hit: HrPersonHit | None,
    *,
    company_contacts_only: bool,
) -> dict[str, str]:
    if company_contacts_only:
        return to_company_contact_row(company)
    return to_team_lead_row(company, hit)


def to_team_lead_row(company: CompanyRecord, hit: HrPersonHit | None) -> dict[str, str]:
    """One export row: company fields + dedicated HR columns."""
    company = finalize_company_record(company)
    company_email, company_phone = company_email_phone(company)
    hr_email, hr_phone = hr_email_phone(company, hit)
    website = (company.website or "").strip()
    linkedin_co = (company.linkedin_company_url or "").strip()
    hr_name = (hit.person_name or "").strip() if hit else ""
    hr_title = (hit.job_title or "").strip() if hit else ""
    hr_li = (hit.linkedin_url or "").strip() if hit else ""

    return {
        "Company Name": company.company_name,
        "Industry/Sector": company.industry or "Technology",
        "Company Website": website,
        "Company Email": company_email,
        "Company Phone": format_phone_for_export(company_phone),
        "Company LinkedIn Page": linkedin_co,
        "Location": format_location(company.city, company.address),
        "HR Contact Name": hr_name,
        "HR Designation": hr_title or ("HR Contact" if hr_li else ""),
        "HR Email": hr_email,
        "HR Phone": format_phone_for_export(hr_phone),
        "HR LinkedIn Profile": hr_li,
    }


def companies_contact_dataframe(
    companies: list[CompanyRecord],
    hr_people: list[HrPersonHit],
    *,
    one_row_per_company: bool = True,
) -> pd.DataFrame:
    """Default: one row per company (avoids duplicate empty email/phone rows)."""
    grouped = _hr_hits_by_company(hr_people)
    rows: list[dict[str, str]] = []
    for c in companies:
        hits = grouped.get(c.company_name, [])
        if one_row_per_company:
            rows.append(to_team_lead_row(c, pick_best_hr_hit(hits)))
        elif not hits:
            rows.append(to_team_lead_row(c, None))
        else:
            for hit in hits:
                rows.append(to_team_lead_row(c, hit))
    return pd.DataFrame(rows, columns=list(TEAM_LEAD_COLUMNS))


def hr_people_dataframe(hr_people: list[HrPersonHit]) -> pd.DataFrame:
    rows = [
        {
            "company_name": h.company_name,
            "person_name": h.person_name,
            "job_title": h.job_title,
            "email": h.email,
            "phone": h.phone,
            "linkedin_url": h.linkedin_url,
            "profile_source": h.profile_source,
        }
        for h in hr_people
    ]
    return pd.DataFrame(rows, columns=list(HR_CSV_COLUMNS))


def save_contact_csv(
    csv_path: Path,
    companies: list[CompanyRecord],
    hr_people: list[HrPersonHit],
    *,
    all_contact_rows: bool = False,
) -> None:
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    prepare_team_lead_df_for_export(
        _export_dataframe(companies, hr_people, all_contact_rows=all_contact_rows)
    ).to_csv(csv_path, index=False, encoding="utf-8-sig")


def _frame_to_company_record(row: pd.Series) -> CompanyRecord:
    def g(*keys: str) -> str:
        for k in keys:
            if k in row.index:
                v = _clean_csv_value(row.get(k, ""))
                if v:
                    return v
        return ""

    return CompanyRecord(
        company_name=g("company_name", "Company Name", "name"),
        city=g("city", "City"),
        industry=g("industry", "Industry/Sector", "Industry"),
        website=g("website", "Company Website", "www"),
        company_email=g("company_email", "Company Email", "Email Address"),
        company_phone=g("company_phone", "Company Phone", "Phone Number"),
        address=g("address", "Location"),
        hr_emails=g("hr_emails", "HR Email"),
        hr_phones=g("hr_phones", "HR Phone"),
        linkedin_company_url=g(
            "linkedin_company_url",
            "Company LinkedIn Page",
            "LinkedIn Profile or Company LinkedIn Page",
        ),
    )


def _frame_to_hr_hits(hr_df: pd.DataFrame) -> list[HrPersonHit]:
    hits: list[HrPersonHit] = []
    if hr_df is None or hr_df.empty:
        return hits
    for _, h in hr_df.iterrows():
        cname = str(h.get("company_name", h.get("Company Name", ""))).strip()
        if not cname:
            continue
        hits.append(
            HrPersonHit(
                company_name=cname,
                person_name=str(
                    h.get("person_name", h.get("HR Contact Name", h.get("Contact Person Name", "")))
                ).strip(),
                job_title=str(
                    h.get("job_title", h.get("HR Designation", h.get("Designation", "")))
                ).strip(),
                email=str(h.get("email", h.get("HR Email", ""))).strip(),
                phone=str(h.get("phone", h.get("HR Phone", ""))).strip(),
                linkedin_url=str(
                    h.get("linkedin_url", h.get("HR LinkedIn Profile", ""))
                ).strip(),
                profile_source=str(h.get("profile_source", "")).strip(),
            )
        )
    return hits


def build_flat_contacts_from_frames(companies_df: pd.DataFrame, hr_df: pd.DataFrame) -> pd.DataFrame:
    companies: list[CompanyRecord] = []
    for _, row in companies_df.iterrows():
        cname = _clean_csv_value(row.get("company_name", row.get("Company Name", "")))
        if cname:
            companies.append(finalize_company_record(_frame_to_company_record(row)))
    hits = _frame_to_hr_hits(hr_df)
    return companies_contact_dataframe(companies, hits, one_row_per_company=True)


def save_contact_csv_from_frame(
    csv_path: Path, df: pd.DataFrame, hr_df: pd.DataFrame | None = None
) -> None:
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    build_flat_contacts_from_frames(df, hr_df if hr_df is not None else pd.DataFrame()).to_csv(
        csv_path, index=False, encoding="utf-8-sig"
    )


def save_output(
    path: Path,
    companies: list[CompanyRecord],
    hr_people: list[HrPersonHit],
    *,
    contact_csv: Path | None = None,
    all_contact_rows: bool = False,
) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    df_team = prepare_team_lead_df_for_export(
        _export_dataframe(companies, hr_people, all_contact_rows=all_contact_rows)
    )
    df_co = pd.DataFrame([asdict(c) for c in companies])
    df_hr = hr_people_dataframe(hr_people)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df_team.to_excel(writer, sheet_name="Company Contacts", index=False)
        df_co.to_excel(writer, sheet_name="raw_companies", index=False)
        df_hr.to_excel(writer, sheet_name="raw_people", index=False)
    csv_out = contact_csv or contact_csv_path(path)
    save_contact_csv(csv_out, companies, hr_people, all_contact_rows=all_contact_rows)
    return csv_out


def load_existing_contact_csv(csv_path: Path) -> pd.DataFrame:
    if not csv_path.is_file():
        return pd.DataFrame()
    df = pd.read_csv(csv_path, dtype=str).fillna("")
    if "Company Name" in df.columns or "company_name" in df.columns:
        return dataframe_to_team_lead(df)
    return pd.DataFrame()


def save_contact_progress(
    rows_out: list[dict[str, str]],
    csv_path: Path,
    *,
    excel_path: Path | None = None,
    company_label: str = "",
    export_columns: tuple[str, ...] | None = None,
) -> None:
    """Write current rows to CSV (and optional Excel) after each company iteration."""
    cols = export_columns or (
        tuple(rows_out[0].keys()) if rows_out else TEAM_LEAD_COLUMNS
    )
    frame = pd.DataFrame(rows_out, columns=list(cols))
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    saved_csv = csv_path
    try:
        write_export_csv(frame, csv_path)
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        saved_csv = csv_path.with_name(f"{csv_path.stem}_autosave_{timestamp}.csv")
        write_export_csv(frame, saved_csv)
        print(f"  Warning: CSV locked (close Excel). Autosave -> {saved_csv}")
    label = f" after '{company_label}'" if company_label else ""
    print(f"  Saved progress: {len(frame)} rows -> {saved_csv}{label}")
    if excel_path:
        excel_path.parent.mkdir(parents=True, exist_ok=True)
        out = prepare_team_lead_df_for_export(frame)
        old_co, old_hr = (
            _load_existing_workbook_frames(excel_path) if excel_path.is_file() else (pd.DataFrame(), pd.DataFrame())
        )
        target_xlsx = excel_path
        try:
            with pd.ExcelWriter(target_xlsx, engine="openpyxl") as writer:
                out.to_excel(writer, sheet_name="Company Contacts", index=False)
                if not old_co.empty:
                    old_co.to_excel(writer, sheet_name="raw_companies", index=False)
                if not old_hr.empty:
                    old_hr.to_excel(writer, sheet_name="raw_people", index=False)
        except PermissionError:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            target_xlsx = excel_path.with_name(f"{excel_path.stem}_autosave_{timestamp}.xlsx")
            with pd.ExcelWriter(target_xlsx, engine="openpyxl") as writer:
                out.to_excel(writer, sheet_name="Company Contacts", index=False)
                if not old_co.empty:
                    old_co.to_excel(writer, sheet_name="raw_companies", index=False)
                if not old_hr.empty:
                    old_hr.to_excel(writer, sheet_name="raw_people", index=False)
            print(f"  Warning: Excel locked. Autosave -> {target_xlsx}")
        else:
            print(f"  Excel updated (same file): {target_xlsx}")


def backup_output_files(
    output_path: Path,
    contact_csv: Path | None = None,
    *,
    backup_dir: Path | None = None,
    label: str = "",
) -> list[Path]:
    """Copy Excel + CSV into output/backups/ before a run changes them."""
    dest_dir = (backup_dir or DEFAULT_BACKUP_DIR).resolve()
    dest_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    tag = f"_{label}" if label else ""
    saved: list[Path] = []
    csv_src = contact_csv or contact_csv_path(output_path)
    for src in (output_path.resolve(), csv_src.resolve()):
        if not src.is_file():
            continue
        backup_name = f"{src.stem}_backup_{timestamp}{tag}{src.suffix}"
        dst = dest_dir / backup_name
        shutil.copy2(src, dst)
        saved.append(dst)
    return saved


def _load_existing_workbook_frames(path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    if not path.is_file():
        return pd.DataFrame(), pd.DataFrame()
    try:
        old_co = pd.read_excel(path, sheet_name="raw_companies", dtype=str).fillna("")
    except ValueError:
        try:
            old_co = pd.read_excel(path, sheet_name="companies", dtype=str).fillna("")
        except ValueError:
            old_co = pd.DataFrame()
    try:
        old_hr = pd.read_excel(path, sheet_name="raw_people", dtype=str).fillna("")
    except ValueError:
        try:
            old_hr = pd.read_excel(path, sheet_name="hr_people", dtype=str).fillna("")
        except ValueError:
            old_hr = pd.DataFrame()
    return old_co, old_hr


def scraped_company_names(path: Path, *, contact_csv: Path | None = None) -> set[str]:
    """Company names already present in output workbook or contacts CSV."""
    names: set[str] = set()
    if path.suffix.lower() not in (".csv",) and path.is_file():
        try:
            df = pd.read_excel(path, sheet_name="Company Contacts", dtype=str).fillna("")
            for col in ("Company Name", "company_name"):
                if col in df.columns:
                    names.update(df[col].astype(str).str.strip().str.lower())
                    break
        except Exception:
            pass
    csv_out = contact_csv or (path if path.suffix.lower() == ".csv" else contact_csv_path(path))
    if csv_out.is_file():
        try:
            prev = load_existing_contact_csv(csv_out)
            if "Company Name" in prev.columns:
                names.update(prev["Company Name"].astype(str).str.strip().str.lower())
        except Exception:
            pass
    return {n for n in names if n}


def exclude_scraped_companies(
    companies: list[dict[str, str]],
    output_path: Path,
    *,
    contact_csv: Path | None = None,
) -> list[dict[str, str]]:
    done = scraped_company_names(output_path, contact_csv=contact_csv)
    if not done:
        return companies
    remaining: list[dict[str, str]] = []
    for c in companies:
        key = c.get("company_name", "").strip().lower()
        if key and key not in done:
            remaining.append(c)
    return remaining


def write_merged_output(
    path: Path,
    companies: list[CompanyRecord],
    hr_people: list[HrPersonHit],
    *,
    contact_csv: Path | None = None,
    all_contact_rows: bool = False,
    merge_existing: bool = True,
    quiet: bool = False,
) -> Path:
    """Write Excel + CSV; optionally merge with existing workbook rows."""
    csv_out = contact_csv or contact_csv_path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    new_co = pd.DataFrame([asdict(c) for c in companies])
    new_hr = hr_people_dataframe(hr_people)

    old_team = pd.DataFrame()
    if merge_existing and path.is_file():
        old_co, old_hr = _load_existing_workbook_frames(path)
        try:
            old_team = pd.read_excel(path, sheet_name="Company Contacts", dtype=str).fillna("")
        except Exception:
            old_team = pd.DataFrame()
        merged_co = (
            pd.concat([old_co, new_co], ignore_index=True).drop_duplicates(
                subset=["company_name"], keep="last"
            )
            if not old_co.empty or not new_co.empty
            else new_co
        )
        key_cols = ["company_name", "linkedin_url"]
        if (
            not old_hr.empty
            and not new_hr.empty
            and all(c in old_hr.columns for c in key_cols)
            and all(c in new_hr.columns for c in key_cols)
        ):
            merged_hr = pd.concat([old_hr, new_hr], ignore_index=True).drop_duplicates(
                subset=key_cols, keep="last"
            )
        elif not new_hr.empty:
            merged_hr = new_hr
        else:
            merged_hr = old_hr
    else:
        merged_co = new_co
        merged_hr = new_hr

    team_df = build_flat_contacts_from_frames(merged_co, merged_hr)
    if merge_existing and not old_team.empty and "Company Name" in old_team.columns:
        team_df = (
            pd.concat([old_team, team_df], ignore_index=True)
            .drop_duplicates(subset=["Company Name"], keep="last")
            .reindex(columns=list(TEAM_LEAD_COLUMNS), fill_value="")
        )
    team_export = prepare_team_lead_df_for_export(
        team_df.reindex(columns=list(TEAM_LEAD_COLUMNS), fill_value="")
    )
    saved_xlsx = path
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            team_export.to_excel(writer, sheet_name="Company Contacts", index=False)
            merged_co.to_excel(writer, sheet_name="raw_companies", index=False)
            merged_hr.to_excel(writer, sheet_name="raw_people", index=False)
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        saved_xlsx = path.with_name(f"{path.stem}_autosave_{timestamp}.xlsx")
        with pd.ExcelWriter(saved_xlsx, engine="openpyxl") as writer:
            team_export.to_excel(writer, sheet_name="Company Contacts", index=False)
            merged_co.to_excel(writer, sheet_name="raw_companies", index=False)
            merged_hr.to_excel(writer, sheet_name="raw_people", index=False)
        if not quiet:
            print(f"  Warning: Excel locked. Autosave -> {saved_xlsx}")

    new_export = team_export
    saved_csv = csv_out
    try:
        new_export.to_csv(csv_out, index=False, encoding="utf-8-sig")
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        saved_csv = csv_out.with_name(f"{csv_out.stem}_autosave_{timestamp}.csv")
        new_export.to_csv(saved_csv, index=False, encoding="utf-8-sig")
        if not quiet:
            print(f"  Warning: CSV locked. Autosave -> {saved_csv}")

    if not quiet:
        print(f"  Saved {len(team_df)} rows -> {saved_xlsx} | CSV -> {saved_csv}")
    return saved_csv


def save_scrape_progress(
    output_path: Path,
    companies: list[CompanyRecord],
    hr_people: list[HrPersonHit],
    *,
    contact_csv: Path | None = None,
    all_contact_rows: bool = False,
    merge_existing: bool = True,
    company_label: str = "",
) -> Path:
    label = f" (after '{company_label}')" if company_label else ""
    print(f"  Incremental save{label}:")
    return write_merged_output(
        output_path,
        companies,
        hr_people,
        contact_csv=contact_csv,
        all_contact_rows=all_contact_rows,
        merge_existing=merge_existing,
        quiet=False,
    )


def append_output(
    path: Path,
    companies: list[CompanyRecord],
    hr_people: list[HrPersonHit],
    *,
    contact_csv: Path | None = None,
    all_contact_rows: bool = False,
) -> Path:
    return write_merged_output(
        path,
        companies,
        hr_people,
        contact_csv=contact_csv,
        all_contact_rows=all_contact_rows,
        merge_existing=True,
        quiet=False,
    )


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Scrape Pakistan tech company HR and profile data.")
    p.add_argument("--input", type=Path, help="Seed CSV/XLSX (Company Name + Profile URL / linkedin_url)")
    p.add_argument(
        "--from-linkedin-pk",
        action="store_true",
        help=(
            "Fast mode (optional): uses linkedin_pk_companies.csv by default. "
            "For a different file use --input data-raw/YOUR_FILE.csv and "
            "--contact-csv output/YOUR_OUTPUT.csv (do not rely on default names alone)."
        ),
    )
    p.add_argument(
        "--csv-only",
        action="store_true",
        help="Save only CSV (no Excel); writes after each company (default with --from-linkedin-pk)",
    )
    p.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    p.add_argument(
        "--contact-csv",
        type=Path,
        default=None,
        help="CRM CSV path (default: output/pakistan_tech_companies_contacts.csv)",
    )
    p.add_argument("--company", help="Scrape only companies whose name contains this text")
    p.add_argument(
        "--max-companies",
        type=int,
        default=100,
        help="How many companies to scrape (default: 100)",
    )
    p.add_argument(
        "--skip",
        type=int,
        default=0,
        help="Skip first N companies in the list (batch resume)",
    )
    p.add_argument(
        "--list-only",
        action="store_true",
        help="Only build the company name list (Step 1), do not scrape contacts",
    )
    p.add_argument(
        "--build-master",
        action="store_true",
        help="Generate data-raw/pakistan_tech_companies_master.csv (128+ names)",
    )
    p.add_argument(
        "--discover",
        action="store_true",
        help="Step 1 only: discover company names via LinkedIn search (saves seed CSV)",
    )
    p.add_argument(
        "--use-curated-list",
        action="store_true",
        help="Use only 10 hand-picked companies (for quick tests)",
    )
    p.add_argument(
        "--no-curated-list",
        action="store_false",
        dest="use_curated_list",
        help="Use master list of 100+ companies (default)",
    )
    p.add_argument(
        "--linkedin-hr",
        action="store_true",
        default=True,
        help="Search LinkedIn for HR / recruiter profiles (default: on)",
    )
    p.add_argument(
        "--no-linkedin-hr",
        action="store_false",
        dest="linkedin_hr",
        help="Skip LinkedIn HR profile search",
    )
    p.add_argument(
        "--hr-top",
        type=int,
        default=3,
        help="Max HR/contact LinkedIn profiles to fetch per company (default: 3; export uses best HR match)",
    )
    p.add_argument("--engine", choices=("auto", "ddgs", "duckduckgo_html", "jina_ddg"), default="auto")
    p.add_argument(
        "--via-jina",
        action="store_true",
        default=True,
        help="Fetch pages through r.jina.ai (default: on — better email/phone from websites)",
    )
    p.add_argument(
        "--no-via-jina",
        action="store_false",
        dest="via_jina",
        help="Fetch company websites directly (no Jina proxy)",
    )
    p.add_argument(
        "--all-contact-rows",
        action="store_true",
        help="Export multiple rows per company (one per LinkedIn person). Default: one row per company.",
    )
    p.add_argument(
        "--max-site-pages",
        type=int,
        default=12,
        help="Max website pages to crawl per company (contact pages prioritized, default 12)",
    )
    p.add_argument(
        "--require-contact",
        action="store_true",
        default=True,
        help="Keep trying (Jina + web search) until email or phone is found (default: on)",
    )
    p.add_argument(
        "--no-require-contact",
        action="store_false",
        dest="require_contact",
        help="Skip extra retries when contact info is missing",
    )
    p.add_argument(
        "--append",
        action="store_true",
        help="(Optional) Same as default now: merge into existing Excel if file exists",
    )
    p.add_argument(
        "--fresh",
        action="store_true",
        help="Start a new workbook for THIS run only (do not merge old Excel rows). Use with care.",
    )
    p.add_argument(
        "--backup-dir",
        type=Path,
        default=DEFAULT_BACKUP_DIR,
        help="Folder for automatic backups (default: output/backups/)",
    )
    p.add_argument(
        "--no-backup",
        action="store_true",
        help="Do not copy Excel/CSV to backup folder before scraping",
    )
    p.add_argument(
        "--auto-skip-scraped",
        action="store_true",
        help="Skip companies already in output Excel/CSV (then take --max-companies from remainder)",
    )
    p.add_argument(
        "--no-save-each-iteration",
        action="store_false",
        dest="save_each_iteration",
        help="Only save at end of run (default: save Excel+CSV after every company)",
    )
    p.set_defaults(save_each_iteration=True)
    p.add_argument(
        "--repair-csv",
        type=Path,
        default=None,
        help="Convert existing contacts CSV (legacy or new columns) to team-lead format with company fields filled",
    )
    p.add_argument(
        "--fill-missing",
        action="store_true",
        help="Fill empty company fields (master list + website scrape; use --fill-missing-offline for instant)",
    )
    p.add_argument(
        "--fill-missing-offline",
        action="store_true",
        help="Fill website/location/LinkedIn/guessed email from master list only (no network)",
    )
    p.add_argument("--sleep", type=float, default=1.2, help="Pause between companies (seconds, default 1.2)")
    p.add_argument(
        "--enrich-csv",
        action="store_true",
        help="Re-scrape ALL rows in contacts CSV missing phone or LinkedIn (full deep collect)",
    )
    p.add_argument(
        "--fetch-email-phone",
        action="store_true",
        help=(
            "Fast company-only pass: website, LinkedIn, location (master), then email/phone "
            "from contact pages. Skips HR unless you also pass --linkedin-hr."
        ),
    )
    p.add_argument(
        "--scrape-contact-pages",
        action="store_true",
        help="Scrape each company /contact-us page for email, phone, and address (updates CSV)",
    )
    p.set_defaults(use_curated_list=False)
    return p.parse_args()


def _export_dataframe(
    companies: list[CompanyRecord], hr_people: list[HrPersonHit], *, all_contact_rows: bool
) -> pd.DataFrame:
    return companies_contact_dataframe(
        companies, hr_people, one_row_per_company=not all_contact_rows
    )


def load_full_company_list(
    session: requests.Session,
    *,
    input_path: Path | None,
    company_filter: str,
    use_curated: bool,
    engine: str,
) -> tuple[list[dict[str, str]], str]:
    """Load entire company list (no skip/batch limit). Returns (records, source_label)."""
    ensure_master_list()
    if input_path is not None:
        resolved = resolve_input_path(input_path, required=True)
        records = load_companies(resolved, company_filter, 0, pakistan_only=False)
        return records, str(resolved)

    if use_curated:
        records = build_company_list(session, 0, engine, use_curated=True)
        return records, "curated (10 companies)"

    if DEFAULT_MASTER.is_file():
        records = load_companies(DEFAULT_MASTER, company_filter, 0, pakistan_only=False)
        if records:
            save_seed_csv(
                [
                    {
                        "company_name": r["company_name"],
                        "city": r.get("city", ""),
                        "industry": r.get("industry", "technology"),
                        "website": r.get("website", ""),
                    }
                    for r in records
                ],
                DEFAULT_SEED,
            )
            return records, DEFAULT_MASTER.name

    records = load_companies(DEFAULT_SEED, company_filter, 0, pakistan_only=False)
    return records, DEFAULT_SEED.name


def apply_batch_window(
    companies: list[dict[str, str]], skip: int, max_companies: int
) -> list[dict[str, str]]:
    if skip > 0:
        companies = companies[skip:]
    if max_companies > 0:
        companies = companies[:max_companies]
    return companies


def hr_hit_from_team_lead_row(row: dict[str, str], company_name: str) -> HrPersonHit | None:
    name = _clean_csv_value(row.get("HR Contact Name", ""))
    li = _clean_csv_value(row.get("HR LinkedIn Profile", ""))
    if not name and not li:
        return None
    return HrPersonHit(
        company_name=company_name,
        person_name=name,
        job_title=_clean_csv_value(row.get("HR Designation", "")),
        email=_clean_csv_value(row.get("HR Email", "")),
        phone=_clean_csv_value(row.get("HR Phone", "")),
        linkedin_url=li,
    )


def enrich_company_partial(
    session: requests.Session,
    company: dict[str, str],
    *,
    engine: str,
    via_jina: bool,
    max_site_pages: int,
    linkedin_hr: bool,
    hr_top_n: int,
    existing_hr: HrPersonHit | None,
) -> tuple[CompanyRecord, HrPersonHit | None]:
    """Fill phone, LinkedIn, and HR when website/email already known (faster than full scrape)."""
    name = company["company_name"]
    rec = finalize_company_record(
        enrich_company_from_master(
            CompanyRecord(
                company_name=name,
                city=company.get("city", ""),
                industry=company.get("industry", "technology"),
                website=company.get("website", ""),
                company_email=company.get("company_email", ""),
                company_phone=company.get("company_phone", ""),
                linkedin_company_url=company.get("linkedin_company_url", ""),
                address=company.get("address", ""),
            )
        )
    )
    website = rec.website or company.get("website", "")
    if website:
        bundle, _, social, li = deep_collect_company_contacts(
            session,
            name,
            website,
            company.get("city", ""),
            engine,
            max_site_pages=max(max_site_pages, 10),
        )
        apply_bundle_to_record(rec, bundle)
        rec = finalize_company_record(rec)
        if not rec.linkedin_company_url.strip():
            rec.linkedin_company_url = li or pick_linkedin_from_social(social)
    if not rec.linkedin_company_url.strip():
        rec.linkedin_company_url = resolve_linkedin_company(session, company, engine) or guess_linkedin_company_url(
            name, website
        )
    hit = existing_hr
    if linkedin_hr and not (hit and hit.linkedin_url):
        hr_list = search_hr_profiles(session, name, website, engine, hr_top_n)
        if not hr_list:
            hr_list = search_contact_people(session, name, website, engine, hr_top_n)
        hit = pick_best_hr_hit(hr_list) if hr_list else hit
    return finalize_company_record(rec), hit


def apply_master_company_fields(row_dict: dict[str, str]) -> dict[str, str]:
    """Fill website, LinkedIn, location from master list (no HR, no network)."""
    name = row_dict.get("Company Name", "")
    if not name:
        return row_dict
    rec = finalize_company_record(
        enrich_company_from_master(
            CompanyRecord(
                company_name=name,
                industry=row_dict.get("Industry/Sector", "") or "technology",
                website=row_dict.get("Company Website", ""),
                company_email=row_dict.get("Company Email", ""),
                company_phone=row_dict.get("Company Phone", ""),
                linkedin_company_url=row_dict.get("Company LinkedIn Page", ""),
                address=row_dict.get("Location", ""),
                city=_clean_csv_value(row_dict.get("Location", "")).split(",")[0],
            )
        )
    )
    if not rec.linkedin_company_url.strip():
        rec.linkedin_company_url = guess_linkedin_company_url(name, rec.website)
    if rec.website:
        row_dict["Company Website"] = rec.website
    if rec.linkedin_company_url:
        row_dict["Company LinkedIn Page"] = rec.linkedin_company_url
    if rec.company_email:
        row_dict["Company Email"] = rec.company_email
    if rec.company_phone:
        row_dict["Company Phone"] = rec.company_phone
    loc = format_location(rec.city, rec.address) or rec.address
    if loc:
        row_dict["Location"] = loc
    if rec.industry:
        row_dict["Industry/Sector"] = rec.industry
    return row_dict


def fill_missing_offline_row(row_dict: dict[str, str]) -> dict[str, str]:
    name = row_dict.get("Company Name", "")
    if not name:
        return row_dict
    if (
        row_dict.get("Company Email")
        and row_dict.get("Company Website")
        and row_dict.get("Company Phone")
        and row_dict.get("Company LinkedIn Page")
        and (row_dict.get("HR LinkedIn Profile") or row_dict.get("HR Contact Name"))
    ):
        return row_dict
    rec = finalize_company_record(
        enrich_company_from_master(
            CompanyRecord(
                company_name=name,
                website=row_dict.get("Company Website", ""),
                company_email=row_dict.get("Company Email", ""),
                company_phone=row_dict.get("Company Phone", ""),
                linkedin_company_url=row_dict.get("Company LinkedIn Page", ""),
                address=row_dict.get("Location", ""),
            )
        )
    )
    if not rec.company_email and rec.website:
        rec.company_email = guess_email_from_website(rec.website)
    return to_team_lead_row(rec, hr_hit_from_team_lead_row(row_dict, name))


def fill_missing_contacts_csv(
    csv_path: Path,
    session: requests.Session,
    *,
    engine: str,
    via_jina: bool,
    max_site_pages: int,
    linkedin_hr: bool,
    hr_top: int,
    offline_only: bool = False,
) -> pd.DataFrame:
    """Fill empty company email/phone/website/LinkedIn rows; keep existing HR columns."""
    df = pd.read_csv(csv_path, dtype=str).fillna("")
    if "company_name" in df.columns:
        df = dataframe_to_team_lead(df)

    rows_out: list[dict[str, str]] = []
    filled = 0
    for idx, row in df.iterrows():
        row_dict = {c: _clean_csv_value(row.get(c, "")) for c in TEAM_LEAD_COLUMNS}
        name = row_dict.get("Company Name", "")
        if not name:
            continue
        if not row_needs_fill(row_dict):
            rows_out.append(row_dict)
            continue

        if offline_only:
            row_dict = fill_missing_offline_row(row_dict)
            filled += 1
            rows_out.append(row_dict)
            continue

        master = match_master_entry(name)
        company = {
            "company_name": name,
            "city": master.get("city", "") or _clean_csv_value(row_dict.get("Location", "")).split(",")[0],
            "industry": "Technology",
            "website": row_dict.get("Company Website") or master.get("website", ""),
            "linkedin_company_url": row_dict.get("Company LinkedIn Page")
            or master.get("linkedin_company_url", ""),
            "company_email": row_dict.get("Company Email", ""),
            "company_phone": row_dict.get("Company Phone", ""),
            "address": row_dict.get("Location", ""),
        }
        existing_hr = hr_hit_from_team_lead_row(row_dict, name)
        print(f"  Filling: {name}")
        try:
            if company.get("website") and company.get("company_email"):
                rec, hit = enrich_company_partial(
                    session,
                    company,
                    engine=engine,
                    via_jina=via_jina,
                    max_site_pages=max_site_pages,
                    linkedin_hr=linkedin_hr,
                    hr_top_n=hr_top,
                    existing_hr=existing_hr,
                )
            else:
                rec, new_hr = scrape_company(
                    session,
                    company,
                    engine=engine,
                    via_jina=via_jina,
                    max_site_pages=max_site_pages,
                    linkedin_hr=linkedin_hr,
                    hr_top_n=hr_top,
                    require_contact=True,
                )
                hit = existing_hr if existing_hr and existing_hr.linkedin_url else pick_best_hr_hit(new_hr)
            if not hit or not hit.linkedin_url:
                if existing_hr and existing_hr.linkedin_url:
                    hit = existing_hr
            row_dict = to_team_lead_row(rec, hit)
            filled += 1
        except Exception as exc:
            print(f"    scrape failed ({exc}); using master + guess")
            rec = finalize_company_record(
                enrich_company_from_master(
                    CompanyRecord(
                        company_name=name,
                        city=master.get("city", ""),
                        website=company.get("website", ""),
                        company_email=company.get("company_email", ""),
                        linkedin_company_url=company.get("linkedin_company_url", ""),
                    )
                )
            )
            if not rec.company_email and rec.website:
                rec.company_email = guess_email_from_website(rec.website)
            row_dict = to_team_lead_row(rec, existing_hr or hr_hit_from_team_lead_row(row_dict, name))
            filled += 1
        time.sleep(0.5)
        rows_out.append(row_dict)

    print(f"Filled or updated {filled} companies with missing contact fields.")
    return pd.DataFrame(rows_out, columns=list(TEAM_LEAD_COLUMNS))


def apply_linkedin_pk_defaults(args: argparse.Namespace) -> None:
    if args.from_linkedin_pk:
        if args.input is not None:
            args.input = resolve_input_path(args.input, required=True)
            print(f"Input CSV (from --input): {args.input}")
        else:
            pk = resolve_linkedin_pk_input(None)
            if not pk:
                print(
                    "LinkedIn PK input not found. Place linkedin_pk.csv or "
                    "linkedin_pk_companies.csv in data-raw/, or pass --input path."
                )
                sys.exit(1)
            args.input = pk
            print(f"Input CSV (default): {args.input}")
        args.csv_only = True
        args.linkedin_hr = False
        args.require_contact = False
        if args.max_site_pages >= 12:
            args.max_site_pages = 8
        if args.sleep >= 1.0:
            args.sleep = 0.5
        if args.contact_csv is None:
            args.contact_csv = DEFAULT_LINKEDIN_PK_OUTPUT
            print(f"Output CSV (default): {args.contact_csv.resolve()}")
        else:
            out_p = Path(args.contact_csv)
            args.contact_csv = (
                out_p.resolve()
                if out_p.is_absolute()
                else (PROJECT_ROOT / out_p).resolve()
            )
            print(f"Output CSV (from --contact-csv): {args.contact_csv}")
        if args.output == DEFAULT_OUTPUT and args.contact_csv:
            args.output = args.contact_csv


def main() -> int:
    args = parse_args()
    if args.input is not None and not args.from_linkedin_pk:
        args.input = resolve_input_path(args.input, required=True)
        print(f"Input CSV: {args.input}")
    apply_linkedin_pk_defaults(args)
    session = requests.Session()

    if args.scrape_contact_pages or args.fetch_email_phone:
        csv_path = args.contact_csv or contact_csv_path(args.output)
        if not csv_path.is_file():
            print(f"Contacts CSV not found: {csv_path}")
            return 1
        load_master_lookup()
        df = pd.read_csv(csv_path, dtype=str).fillna("")
        if "company_name" in df.columns:
            df = dataframe_to_team_lead(df)
        rows_out: list[dict[str, str]] = []
        updated = 0
        all_pages = args.scrape_contact_pages
        excel_path = args.output.resolve() if args.output.suffix.lower() in (".xlsx", ".xls") else None
        total = sum(
            1
            for _, row in df.iterrows()
            if _clean_csv_value(row.get("Company Name", row.get("company_name", "")))
        )
        idx = 0
        mode = "company contacts (no HR)" if not args.linkedin_hr else "company contacts + HR"
        print(f"Fast {mode} for {total} companies (saving after each row)...")
        for _, row in df.iterrows():
            row_dict = {c: _clean_csv_value(row.get(c, "")) for c in TEAM_LEAD_COLUMNS}
            name = row_dict.get("Company Name", "")
            if not name:
                continue
            idx += 1
            print(f"[{idx}/{total}] {name}")
            row_dict = apply_master_company_fields(row_dict)
            website = row_dict.get("Company Website", "")
            master = match_master_entry(name)
            has_guessed_only = (
                row_dict.get("Company Email", "").startswith("info@")
                and not row_dict.get("Company Phone")
            )
            need_email_phone = (
                not row_dict.get("Company Phone")
                or not row_dict.get("Company Email")
                or has_guessed_only
            )
            need = all_pages or need_email_phone or (
                args.scrape_contact_pages and not row_dict.get("Location")
            )
            if need and website:
                print(f"  Contact page: {name} ({website})")
                try:
                    info = fetch_contact_from_website(session, website)
                    if not info.pakistan_email:
                        em2, _ = search_company_email_phone_snippets(
                            session, name, website, master.get("city", ""), args.engine
                        )
                        info.pakistan_email = info.pakistan_email or em2
                    known_em, known_ph, _ = lookup_known_contact(name)
                    em = info.pakistan_email or known_em
                    ph = info.pakistan_phone or known_ph
                    if em:
                        row_dict["Company Email"] = em
                    if ph:
                        row_dict["Company Phone"] = ph
                    if info.pakistan_address:
                        row_dict["Location"] = info.pakistan_address
                    if em or ph or info.pakistan_address:
                        updated += 1
                    print(
                        f"    website={row_dict.get('Company Website') or '-'} | "
                        f"linkedin={row_dict.get('Company LinkedIn Page') or '-'} | "
                        f"email={row_dict.get('Company Email') or '-'} | "
                        f"phone={row_dict.get('Company Phone') or '-'} | "
                        f"location={(row_dict.get('Location') or '-')[:50]}"
                    )
                    if info.source_url:
                        print(f"    source: {info.source_url}")
                except Exception as exc:
                    print(f"    failed: {exc}")
                time.sleep(max(0.0, args.sleep))
            elif need and not website:
                print(
                    f"  Skip scrape (no website). "
                    f"linkedin={row_dict.get('Company LinkedIn Page') or '-'} | "
                    f"location={row_dict.get('Location') or '-'}"
                )
            else:
                print(
                    f"  OK (master). website={row_dict.get('Company Website') or '-'} | "
                    f"linkedin={row_dict.get('Company LinkedIn Page') or '-'} | "
                    f"email={row_dict.get('Company Email') or '-'} | "
                    f"phone={row_dict.get('Company Phone') or '-'}"
                )
            rows_out.append(row_dict)
            save_contact_progress(rows_out, csv_path, excel_path=excel_path, company_label=name)
        out = pd.DataFrame(rows_out, columns=list(TEAM_LEAD_COLUMNS))
        n_phone = sum(1 for _, r in out.iterrows() if _clean_csv_value(r.get("Company Phone", "")))
        n_email = sum(1 for _, r in out.iterrows() if _clean_csv_value(r.get("Company Email", "")))
        print(f"Updated {updated} companies -> {csv_path}")
        print(f"  Total with email: {n_email}/{len(out)} | with phone: {n_phone}/{len(out)}")
        return 0

    if args.enrich_csv:
        csv_path = args.contact_csv or contact_csv_path(args.output)
        if not csv_path.is_file():
            print(f"Contacts CSV not found: {csv_path}")
            return 1
        load_master_lookup()
        df = pd.read_csv(csv_path, dtype=str).fillna("")
        if "company_name" in df.columns:
            df = dataframe_to_team_lead(df)
        rows_out: list[dict[str, str]] = []
        for _, row in df.iterrows():
            row_dict = {c: _clean_csv_value(row.get(c, "")) for c in TEAM_LEAD_COLUMNS}
            name = row_dict.get("Company Name", "")
            if not name:
                continue
            needs = (
                not row_dict.get("Company Phone")
                or not row_dict.get("Company LinkedIn Page")
                or not row_dict.get("Company Email")
            )
            if needs:
                print(f"  Enriching: {name}")
                master = match_master_entry(name)
                company = {
                    "company_name": name,
                    "city": master.get("city", ""),
                    "website": row_dict.get("Company Website") or master.get("website", ""),
                    "linkedin_company_url": row_dict.get("Company LinkedIn Page")
                    or master.get("linkedin_company_url", ""),
                }
                try:
                    rec, hr_list = scrape_company(
                        session,
                        company,
                        engine=args.engine,
                        via_jina=args.via_jina,
                        max_site_pages=max(args.max_site_pages, 12),
                        linkedin_hr=args.linkedin_hr,
                        hr_top_n=args.hr_top,
                        require_contact=True,
                    )
                    hit = hr_hit_from_team_lead_row(row_dict, name)
                    if not hit or not hit.linkedin_url:
                        hit = pick_best_hr_hit(hr_list) if hr_list else hit
                    row_dict = to_team_lead_row(rec, hit)
                except Exception as exc:
                    print(f"    failed: {exc}")
                time.sleep(max(1.0, args.sleep))
            rows_out.append(row_dict)
        out = pd.DataFrame(rows_out, columns=list(TEAM_LEAD_COLUMNS))
        out.to_csv(csv_path, index=False, encoding="utf-8-sig")
        print(f"Enriched CSV -> {csv_path}")
        return 0

    if args.fill_missing or args.fill_missing_offline:
        csv_path = args.contact_csv or contact_csv_path(args.output)
        if not csv_path.is_file():
            print(f"Contacts CSV not found: {csv_path}")
            return 1
        load_master_lookup()
        mode = "offline (master list)" if args.fill_missing_offline else "scrape + HR + phone + LinkedIn"
        print(f"Filling missing company fields in {csv_path} [{mode}] ...")
        print("  (needs internet; may take 1-2 min per company)")
        out = fill_missing_contacts_csv(
            csv_path,
            session,
            engine=args.engine,
            via_jina=args.via_jina,
            max_site_pages=args.max_site_pages,
            linkedin_hr=args.linkedin_hr,
            hr_top=args.hr_top,
            offline_only=args.fill_missing_offline,
        )
        out.to_csv(csv_path, index=False, encoding="utf-8-sig")
        n_email = sum(1 for _, r in out.iterrows() if _clean_csv_value(r.get("Company Email", "")))
        n_phone = sum(1 for _, r in out.iterrows() if _clean_csv_value(r.get("Company Phone", "")))
        n_li = sum(1 for _, r in out.iterrows() if _clean_csv_value(r.get("Company LinkedIn Page", "")))
        n_hr = sum(1 for _, r in out.iterrows() if _clean_csv_value(r.get("HR LinkedIn Profile", "")))
        print(f"Saved {len(out)} rows -> {csv_path}")
        print(f"  Company email: {n_email}/{len(out)} | phone: {n_phone}/{len(out)} | LinkedIn: {n_li}/{len(out)}")
        print(f"  HR LinkedIn: {n_hr}/{len(out)}")
        return 0

    if args.repair_csv:
        src = args.repair_csv
        if not src.is_file():
            print(f"File not found: {src}")
            return 1
        df = pd.read_csv(src, dtype=str).fillna("")
        out = dataframe_to_team_lead(df)
        dest = args.contact_csv or contact_csv_path(args.output)
        dest.parent.mkdir(parents=True, exist_ok=True)
        out.to_csv(dest, index=False, encoding="utf-8-sig")
        filled = sum(
            1
            for _, r in out.iterrows()
            if str(r.get("Company Email", "")).strip() and str(r.get("Company Phone", "")).strip()
        )
        print(f"Repaired {len(out)} rows -> {dest}")
        print(f"  With company email+phone: {filled}/{len(out)}")
        return 0

    limit = args.max_companies if args.max_companies > 0 else 100

    if args.build_master:
        gen = _SCRIPTS / "generate_master_company_list.py"
        if gen.is_file():
            import subprocess

            subprocess.run([sys.executable, str(gen)], cwd=str(PROJECT_ROOT), check=False)
        ensure_master_list()
        print(f"Master list ready: {DEFAULT_MASTER}")
        return 0

    if args.discover:
        companies = build_company_list(
            session, limit, args.engine, use_curated=args.use_curated_list
        )
        print(f"Step 1 done: {len(companies)} company names saved. Run again without --discover to scrape contacts.")
        return 0

    # Step 1: load full list, then apply --skip / --max-companies for this batch
    all_companies_list, source = load_full_company_list(
        session,
        input_path=args.input,
        company_filter=args.company or "",
        use_curated=args.use_curated_list,
        engine=args.engine,
    )
    total_in_list = len(all_companies_list)
    print(f"Step 1: {total_in_list} companies in list ({source})")

    if args.company and not all_companies_list:
        print(f'No companies matched "{args.company}".')
        return 1

    if args.skip >= total_in_list:
        print(
            f"ERROR: --skip {args.skip} but the list only has {total_in_list} companies.\n"
            f"  List file: {DEFAULT_MASTER if DEFAULT_MASTER.is_file() else DEFAULT_SEED}\n"
            f"  Use --no-curated-list (default) for 100+ companies, or run: python scripts/pakistan_tech_companies_data.py --build-master"
        )
        return 1

    companies = list(all_companies_list)
    csv_only = args.csv_only or str(args.output).lower().endswith(".csv")
    csv_path = (args.contact_csv or args.output).resolve()
    if csv_only and args.contact_csv is None and not str(args.output).lower().endswith(".csv"):
        csv_path = contact_csv_path(args.output.resolve())
    if args.auto_skip_scraped:
        before = len(companies)
        companies = exclude_scraped_companies(
            companies,
            args.output.resolve(),
            contact_csv=csv_path if csv_only else args.contact_csv,
        )
        print(
            f"Auto-skip scraped: {before - len(companies)} already in output, "
            f"{len(companies)} remaining in list"
        )
    companies = apply_batch_window(companies, args.skip, limit)
    if args.skip > 0 or limit < total_in_list or args.auto_skip_scraped:
        print(
            f"Batch window: skip={args.skip}, take={limit} -> scraping {len(companies)} companies "
            f"(from {total_in_list} in master list)"
        )

    if args.list_only:
        print(f"List-only mode: saved {len(companies)} company names (no scraping).")
        print(f"  Master/seed: {DEFAULT_MASTER if DEFAULT_MASTER.is_file() else DEFAULT_SEED}")
        print("Run without --list-only to scrape email, phone, LinkedIn, HR.")
        return 0

    out_path = args.output.resolve()
    if not csv_only:
        csv_path = (args.contact_csv or contact_csv_path(args.output)).resolve()
    if not args.no_backup:
        backups = backup_output_files(
            out_path if not csv_only else csv_path,
            csv_path,
            backup_dir=args.backup_dir,
            label="fresh" if args.fresh else "run",
        )
        if backups:
            print(f"Backup folder: {args.backup_dir.resolve()}")
            for bp in backups:
                print(f"  -> {bp.name}")
        else:
            print(
                f"No backup created (output files not found yet). "
                f"Backups will go to: {args.backup_dir.resolve()}"
            )

    if csv_only:
        if args.fresh:
            print(f"FRESH CSV run -> {csv_path}")
        elif csv_path.is_file():
            print(f"SAFE CSV mode: merging into existing file -> {csv_path}")
        else:
            print(f"New CSV will be created -> {csv_path}")
    elif args.fresh:
        print("FRESH run: will NOT merge previous Excel rows (new workbook content for this session).")
    elif out_path.is_file():
        print(f"SAFE mode: merging into existing file -> {out_path}")
    else:
        print(f"New file will be created -> {out_path}")

    mode_label = "phone/email/website (no HR)" if not args.linkedin_hr else "contact + HR"
    print(f"Step 2: Scraping {mode_label} for {len(companies)} companies...")

    if not companies:
        print(
            "No companies in this batch. Check --skip and --max-companies.\n"
            f"  Full list: {total_in_list} companies in {source}\n"
            "  Rebuild list: python scripts/pakistan_tech_companies_data.py --build-master"
        )
        return 1

    if csv_only:
        export_cols = export_columns_for_args(args)
        rows_out: list[dict[str, str]] = (
            [] if args.fresh else load_team_lead_rows_from_csv(csv_path, export_cols)
        )
        contacts_only = not args.linkedin_hr or args.from_linkedin_pk
        for idx, company in enumerate(companies, start=1):
            name = company["company_name"]
            li = company.get("linkedin_company_url", "")
            print(f"[{idx}/{len(companies)}] {name}")
            if li:
                print(f"  linkedin seed: {li}")
            try:
                rec, hr_hits = scrape_company(
                    session,
                    company,
                    engine=args.engine,
                    via_jina=args.via_jina,
                    max_site_pages=args.max_site_pages,
                    linkedin_hr=args.linkedin_hr,
                    hr_top_n=args.hr_top,
                    require_contact=args.require_contact,
                )
                row = record_to_export_row(
                    rec,
                    pick_best_hr_hit(hr_hits),
                    company_contacts_only=contacts_only,
                )
                print(
                    f"  website={row.get('Company Website') or '-'} | "
                    f"email={row.get('Company Email') or '-'} | "
                    f"phone={row.get('Company Phone') or '-'} | "
                    f"phones={row.get('All Phone Numbers') or '-'} | "
                    f"city={row.get('City') or '-'} | "
                    f"address={(row.get('Address') or '-')[:40]}"
                )
            except Exception as exc:
                print(f"  ERROR: {exc}")
                row = record_to_export_row(
                    CompanyRecord(
                        company_name=name,
                        linkedin_company_url=li,
                        notes=f"error:{exc}",
                    ),
                    None,
                    company_contacts_only=contacts_only,
                )
            upsert_team_lead_row(rows_out, row)
            if args.save_each_iteration:
                save_contact_progress(
                    rows_out,
                    csv_path,
                    company_label=name,
                    export_columns=export_cols,
                )
            time.sleep(max(0.0, args.sleep))

        write_export_csv(pd.DataFrame(rows_out, columns=list(export_cols)), csv_path)
        out_df = pd.read_csv(csv_path, dtype=str).fillna("")
        n_email = sum(1 for _, r in out_df.iterrows() if _clean_csv_value(r.get("Company Email", "")))
        n_phone = sum(1 for _, r in out_df.iterrows() if _clean_csv_value(r.get("Company Phone", "")))
        n_multi = sum(
            1 for _, r in out_df.iterrows() if _clean_csv_value(r.get("All Phone Numbers", ""))
        )
        print(f"Saved {len(out_df)} rows -> {csv_path}")
        print(
            f"  email: {n_email}/{len(out_df)} | phone: {n_phone}/{len(out_df)} | "
            f"multiple phones: {n_multi}/{len(out_df)}"
        )
        return 0

    all_companies: list[CompanyRecord] = []
    all_hr: list[HrPersonHit] = []
    merge_existing = out_path.is_file() and not args.fresh

    for idx, company in enumerate(companies, start=1):
        name = company["company_name"]
        print(f"[{idx}/{len(companies)}] {name}")
        try:
            rec, hr_hits = scrape_company(
                session,
                company,
                engine=args.engine,
                via_jina=args.via_jina,
                max_site_pages=args.max_site_pages,
                linkedin_hr=args.linkedin_hr,
                hr_top_n=args.hr_top,
                require_contact=args.require_contact,
            )
            all_companies.append(rec)
            all_hr.extend(hr_hits)
            loc = format_location(rec.city, rec.address) or "MISSING"
            print(
                f"  website={rec.website or '-'} | email={rec.company_email or 'MISSING'} | "
                f"phone={rec.company_phone or 'MISSING'} | location={loc[:50]} | status={rec.contact_status}"
            )
        except Exception as exc:
            print(f"  ERROR: {exc}")
            all_companies.append(
                CompanyRecord(company_name=name, notes=f"error:{exc}")
            )
        if args.save_each_iteration and all_companies:
            save_scrape_progress(
                args.output.resolve(),
                all_companies,
                all_hr,
                contact_csv=csv_path,
                all_contact_rows=args.all_contact_rows,
                merge_existing=merge_existing,
                company_label=name,
            )
        time.sleep(max(0.0, args.sleep))

    csv_path = write_merged_output(
        args.output.resolve(),
        all_companies,
        all_hr,
        contact_csv=csv_path,
        all_contact_rows=args.all_contact_rows,
        merge_existing=merge_existing,
        quiet=False,
    )

    with_email = sum(1 for c in all_companies if c.company_email)
    with_phone = sum(1 for c in all_companies if c.company_phone)
    with_linkedin = sum(1 for c in all_companies if c.linkedin_company_url)
    with_location = sum(
        1 for c in all_companies if format_location(c.city, c.address).strip()
    )
    complete = sum(1 for c in all_companies if c.contact_status == "complete")
    print(f"Saved {len(all_companies)} companies -> {args.output}")
    print(f"Saved team-lead CSV -> {csv_path}")
    print(
        f"  email: {with_email}/{len(all_companies)} | phone: {with_phone}/{len(all_companies)} | "
        f"location: {with_location}/{len(all_companies)} | linkedin: {with_linkedin}/{len(all_companies)} | complete: {complete}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
