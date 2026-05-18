"""
Pakistan university faculty scraper — names, titles, email, phone, LinkedIn.

Primary source: official university faculty/staff web pages (mailto, tables, cards).
Optional: LinkedIn search per department.

Run:
  python scripts/pakistan_faculty_data.py --university "Aga Khan University" --max-institutions 1
  python scripts/pakistan_faculty_data.py --university "Riphah" --campus-only Lahore --campus-only Faisalabad
  (campus-wise is ON by default: Islamabad, Lahore, Faisalabad, ... x each department)
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, unquote, urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag

try:
    from ddgs import DDGS
except Exception:
    try:
        from duckduckgo_search import DDGS
    except Exception:
        DDGS = None

_SCRIPTS = Path(__file__).resolve().parent
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

import uni_data_free_google as udg  # noqa: E402

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = PROJECT_ROOT / "data-raw" / "institutions.csv"
DEFAULT_OUTPUT = PROJECT_ROOT / "output" / "pakistan_faculty.xlsx"

FACULTY_DIVISION_TYPES = frozenset(
    {
        "faculty",
        "department/division",
        "department",
        "school",
        "college",
        "institute",
        "centre",
        "center",
    }
)

CAMPUS_DIVISION_TYPES = frozenset({"campus", "campus abroad"})

# Known multi-campus cities (used with --all-campuses + web discovery).
CAMPUS_PRESETS: dict[str, list[str]] = {
    "riphah": ["Islamabad", "Lahore", "Faisalabad", "Rawalpindi", "Multan", "Peshawar", "Gujrat", "Sargodha"],
    "bahria": ["Islamabad", "Karachi", "Lahore"],
    "fast": ["Karachi", "Lahore", "Islamabad", "Peshawar", "Faisalabad", "Multan"],
    "comsats": ["Islamabad", "Lahore", "Abbottabad", "Wah", "Vehari", "Sahiwal", "Attock"],
    "khyber medical": [
        "Peshawar",
        "Hayatabad",
        "Kohat",
        "Mardan",
        "Swabi",
        "Swat",
        "Islamabad",
        "Abbottabad",
        "Dir",
        "Lower Dir",
        "Parachinar",
        "Lakki Marwat",
        "Hazara",
        "Bannu",
        "Dera Ismail Khan",
        "Timergara",
        "Mansehra",
    ],
    "kmu": [
        "Peshawar",
        "Hayatabad",
        "Kohat",
        "Mardan",
        "Swabi",
        "Swat",
        "Islamabad",
        "Abbottabad",
        "Dir",
        "Lower Dir",
        "Parachinar",
        "Lakki Marwat",
        "Hazara",
        "Bannu",
        "Dera Ismail Khan",
        "Timergara",
        "Mansehra",
    ],
    "aga khan": ["Karachi", "Islamabad", "Hyderabad"],
    "nust": ["Islamabad", "Rawalpindi", "Risalpur"],
    "uol": ["Lahore", "Gujrat", "Sargodha", "Pakpattan"],
    "university of lahore": ["Lahore", "Islamabad", "Gujrat", "Sargodha", "Pakpattan"],
}

# Cities/areas to detect in URLs and page text (longest phrases first when matching).
PAKISTAN_CAMPUS_CITIES: tuple[str, ...] = tuple(
    sorted(
        {
            "Islamabad",
            "Rawalpindi",
            "Karachi",
            "Lahore",
            "Faisalabad",
            "Multan",
            "Peshawar",
            "Hayatabad",
            "Quetta",
            "Hyderabad",
            "Abbottabad",
            "Mardan",
            "Swabi",
            "Swat",
            "Kohat",
            "Gujrat",
            "Sargodha",
            "Rawalakot",
            "Mirpur",
            "Muzaffarabad",
            "Gilgit",
            "Skardu",
            "Bannu",
            "Dera Ismail Khan",
            "Lakki Marwat",
            "Parachinar",
            "Lower Dir",
            "Timergara",
            "Mansehra",
            "Hazara",
            "Chiniot",
            "Sahiwal",
            "Wah",
            "Attock",
            "Vehari",
            "Bahawalpur",
            "Sialkot",
            "Gujranwala",
            "Nawabshah",
            "Larkana",
            "Sukkur",
            "Dir",
        },
        key=len,
        reverse=True,
    )
)

FACULTY_ROLE_KEYWORDS = [
    "professor emeritus",
    "associate professor",
    "assistant professor",
    "professor",
    "senior lecturer",
    "lecturer",
    "instructor",
    "faculty member",
    "head of department",
    "dean",
    "hod",
    "chairperson",
    "chairman",
    "research fellow",
    "teaching fellow",
]
FACULTY_ROLE_RE = re.compile(
    r"\b(" + "|".join(re.escape(x) for x in sorted(FACULTY_ROLE_KEYWORDS, key=len, reverse=True)) + r")\b",
    re.IGNORECASE,
)

FACULTY_URL_HINTS = (
    "faculty",
    "faculties",
    "staff",
    "directory",
    "people",
    "academics",
    "our-team",
    "our-faculty",
    "faculty-members",
    "faculty-staff",
    "teaching-staff",
    "lecturers",
    "professors",
)

NAME_PREFIX_RE = re.compile(
    r"^(?:(?:Dr|Prof|Professor|Mr|Ms|Mrs|Miss|Engr)\.?\s+)+",
    re.IGNORECASE,
)

HEADERS = udg.browser_headers()


@dataclass
class FacultyHit:
    faculty_name: str = ""
    job_title: str = ""
    email: str = ""
    phone: str = ""
    linkedin_url: str = ""
    profile_page_url: str = ""
    result_title: str = ""
    snippet: str = ""
    source_url: str = ""
    data_source: str = ""

    def to_dict(self) -> dict[str, str]:
        return {
            "faculty_name": self.faculty_name,
            "job_title": self.job_title,
            "email": self.email,
            "phone": self.phone,
            "linkedin_url": self.linkedin_url,
            "profile_page_url": self.profile_page_url,
            "result_title": self.result_title,
            "snippet": self.snippet,
            "source_url": self.source_url,
            "data_source": self.data_source,
        }


def parse_divisions(raw: object) -> list[dict[str, str]]:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return []
    s = str(raw).strip()
    if not s or s == "[]":
        return []
    try:
        data = json.loads(s)
    except json.JSONDecodeError:
        return []
    if not isinstance(data, list):
        return []
    out: list[dict[str, str]] = []
    for item in data:
        if not isinstance(item, dict):
            continue
        div_type = str(item.get("type") or "").strip()
        name = str(item.get("name") or "").strip()
        if not name:
            continue
        if div_type and div_type.lower() not in FACULTY_DIVISION_TYPES:
            continue
        out.append(
            {
                "division_type": div_type,
                "department": name,
                "fields_of_study": str(item.get("fields_of_study") or "").strip(),
            }
        )
    return out


def parse_campuses(raw: object) -> list[dict[str, str]]:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return []
    s = str(raw).strip()
    if not s or s == "[]":
        return []
    try:
        data = json.loads(s)
    except json.JSONDecodeError:
        return []
    if not isinstance(data, list):
        return []
    out: list[dict[str, str]] = []
    for item in data:
        if not isinstance(item, dict):
            continue
        div_type = str(item.get("type") or "").strip()
        name = str(item.get("name") or "").strip()
        if not name:
            continue
        if div_type.lower() not in CAMPUS_DIVISION_TYPES:
            continue
        out.append(
            {
                "division_type": div_type,
                "campus": name,
                "fields_of_study": str(item.get("fields_of_study") or "").strip(),
            }
        )
    return out


def preset_campus_names(university: str) -> list[str]:
    u = university.lower()
    for key, cities in CAMPUS_PRESETS.items():
        if key in u:
            return list(cities)
    return []


def extract_cities_from_text(text: str) -> list[str]:
    if not text:
        return []
    low = text.lower()
    found: list[str] = []
    for city in PAKISTAN_CAMPUS_CITIES:
        if re.search(r"\b" + re.escape(city.lower()) + r"\b", low):
            found.append(city)
    return found


def discover_campuses_online(
    session: requests.Session,
    university: str,
    official_www: str,
    engine: str,
    *,
    via_jina: bool = False,
) -> list[str]:
    """Find campus cities from search results and the university website."""
    discovered: list[str] = []
    base = normalize_www(official_www)
    host = host_of(base)

    queries: list[str] = []
    if host:
        queries.extend(
            [
                f"site:{host} campus institute location",
                f"site:{host} constituent colleges affiliated",
                f"site:{host} regional campus",
            ]
        )
    queries.append(f'"{university}" campus cities Pakistan')
    queries.append(f'"{university}" institutes locations Pakistan')

    for q in queries:
        try:
            urls, _eng = udg.search_urls(session, q, 12, engine)
            for u in urls:
                discovered.extend(extract_cities_from_text(u))
                discovered.extend(extract_cities_from_text(urlparse(u).path.replace("-", " ")))
        except Exception:
            continue

    if base:
        for path in ("/institutes", "/institutes/constituent", "/campuses", "/contact-us", "/about"):
            page_url = base + path
            try:
                title, html, status = fetch_page_html(session, page_url, via_jina=via_jina)
                if status != 200 or not html:
                    continue
                discovered.extend(extract_cities_from_text(title))
                discovered.extend(extract_cities_from_text(html))
                if not via_jina:
                    soup = BeautifulSoup(html, "html.parser")
                    for a in soup.find_all("a", href=True):
                        blob = f"{a.get_text(' ', strip=True)} {a.get('href', '')}"
                        discovered.extend(extract_cities_from_text(blob))
            except Exception:
                continue

    return discovered


def collect_campus_names(
    university: str,
    main_city: str,
    divisions_json: str,
    extra_campuses: list[str],
    *,
    use_presets: bool,
    discover_all: bool = False,
    session: requests.Session | None = None,
    official_www: str = "",
    engine: str = "auto",
    via_jina: bool = False,
) -> list[str]:
    seen: set[str] = set()
    names: list[str] = []

    def add(name: str) -> None:
        n = udg._normalize_space(name)
        if not n or len(n) < 3:
            return
        key = n.lower()
        if key in seen:
            return
        seen.add(key)
        names.append(n)

    if main_city:
        add(main_city)
    for row in parse_campuses(divisions_json):
        add(row["campus"])
    for c in extra_campuses:
        add(c)
    if use_presets:
        for c in preset_campus_names(university):
            add(c)
    if discover_all and session is not None:
        for c in discover_campuses_online(
            session, university, official_www, engine, via_jina=via_jina
        ):
            add(c)
    return names


def scrape_key(iau_id: str, university: str, campus: str, department: str) -> str:
    base = iau_id.strip() or university.strip()
    camp = campus.strip().lower() or "(no-campus)"
    dept = department.strip().lower() or "(university-wide)"
    return f"{base}|{camp}|{dept}"


def apply_campus_wise_city(row: dict[str, str]) -> dict[str, str]:
    """
    Normalize city / whed_city / campus on one output row.
    - city = scrape location (campus name when campus-wise)
    - whed_city = WHED head-office city from CSV (e.g. Islamabad for Riphah)
    - campus = scrape location when campus-wise
    """
    out = dict(row)
    campus = str(out.get("campus", "")).strip()
    city = str(out.get("city", "")).strip()
    whed = str(out.get("whed_city", "")).strip()

    if campus:
        if not whed and city and city.lower() != campus.lower():
            out["whed_city"] = city
        out["city"] = campus
    elif city and not whed:
        out["whed_city"] = city
    return out


def repair_rows_campus_cities(rows: list[dict[str, str]]) -> tuple[list[dict[str, str]], int]:
    """Fix legacy rows where city was always WHED HQ but campus had the real location."""
    changed = 0
    repaired: list[dict[str, str]] = []
    for row in rows:
        before_city = str(row.get("city", "")).strip()
        before_whed = str(row.get("whed_city", "")).strip()
        campus = str(row.get("campus", "")).strip()
        fixed = apply_campus_wise_city(row)
        after_city = str(fixed.get("city", "")).strip()
        after_whed = str(fixed.get("whed_city", "")).strip()
        if campus and (before_city != after_city or (not before_whed and after_whed)):
            changed += 1
        elif not campus and not before_whed and after_whed:
            changed += 1
        repaired.append(fixed)
    return repaired, changed


def repair_excel_file(output_path: Path) -> int:
    if not output_path.is_file():
        print(f"File not found: {output_path}")
        return 1
    try:
        prior = pd.read_excel(output_path, dtype=str, engine="openpyxl")
    except Exception as exc:
        print(f"Could not read {output_path}: {exc}")
        return 1
    rows = prior.fillna("").astype(str).to_dict(orient="records")
    repaired, n = repair_rows_campus_cities(rows)
    save_progress(repaired, output_path)
    print(f"Repaired {n} of {len(repaired)} rows -> {output_path}")
    with_campus = sum(1 for r in repaired if str(r.get("campus", "")).strip())
    cities = sorted({str(r.get("city", "")).strip() for r in repaired if str(r.get("city", "")).strip()})
    print(f"Rows with campus set: {with_campus} | distinct city values: {', '.join(cities[:12])}{'...' if len(cities) > 12 else ''}")
    return 0


def load_existing_scrape_keys(output_path: Path) -> set[str]:
    if not output_path.is_file():
        return set()
    try:
        existing = pd.read_excel(output_path, dtype=str, engine="openpyxl")
    except Exception:
        return set()
    keys: set[str] = set()
    for _, row in existing.iterrows():
        campus = str(row.get("campus", "")).strip()
        if not campus:
            campus = str(row.get("city", "")).strip()
        keys.add(
            scrape_key(
                str(row.get("iau_id", "")),
                str(row.get("university_name", "")),
                campus,
                str(row.get("department", "")),
            )
        )
    return keys


def build_scrape_targets(
    *,
    university: str,
    main_city: str,
    divisions_json: str,
    extra_campuses: list[str],
    campus_wise: bool,
    use_campus_presets: bool,
    discover_all_campuses: bool,
    session: requests.Session | None,
    official_www: str,
    engine: str,
    via_jina: bool,
    max_departments: int,
) -> list[dict[str, str]]:
    departments = parse_divisions(divisions_json)
    if max_departments and departments:
        departments = departments[:max_departments]

    if not campus_wise:
        if not departments:
            departments = [{"division_type": "", "department": "", "fields_of_study": ""}]
        return [
            {
                "campus": "",
                "search_city": main_city,
                "division_type": d["division_type"],
                "department": d["department"],
                "fields_of_study": d["fields_of_study"],
            }
            for d in departments
        ]

    campuses = collect_campus_names(
        university,
        main_city,
        divisions_json,
        extra_campuses,
        use_presets=use_campus_presets,
        discover_all=discover_all_campuses,
        session=session,
        official_www=official_www,
        engine=engine,
        via_jina=via_jina,
    )
    if not campuses:
        campuses = [main_city] if main_city else ["All campuses"]

    if not departments:
        departments = [{"division_type": "", "department": "", "fields_of_study": ""}]

    targets: list[dict[str, str]] = []
    for campus in campuses:
        for d in departments:
            targets.append(
                {
                    "campus": campus,
                    "search_city": campus,
                    "division_type": d["division_type"],
                    "department": d["department"],
                    "fields_of_study": d["fields_of_study"],
                }
            )
    return targets


def save_progress(rows_so_far: list[dict[str, str]], output_path: Path) -> None:
    out_df = pd.DataFrame(rows_so_far)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        out_df.to_excel(output_path, index=False, engine="openpyxl")
        print(f"  Saved {len(out_df)} rows -> {output_path}")
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback = output_path.with_name(f"{output_path.stem}_autosave_{timestamp}.xlsx")
        out_df.to_excel(fallback, index=False, engine="openpyxl")
        print(f"  Warning: Excel locked. Saved to: {fallback}")


def normalize_www(www: str) -> str:
    s = (www or "").strip()
    if not s:
        return ""
    if not s.startswith(("http://", "https://")):
        s = "https://" + s
    return s.rstrip("/")


def host_of(url: str) -> str:
    try:
        return urlparse(url).netloc.lower().replace("www.", "")
    except Exception:
        return ""


def clean_name(name: str) -> str:
    return NAME_PREFIX_RE.sub("", udg._normalize_space(name)).strip()


def is_plausible_name(name: str, university: str) -> bool:
    n = clean_name(name)
    if len(n) < 4 or len(n.split()) < 2:
        return False
    try:
        return udg._is_name_like(n, university)
    except Exception:
        low = n.lower()
        bad = ("university", "faculty", "department", "click here", "read more", "contact us")
        return not any(b in low for b in bad)


def name_from_email(email: str) -> str:
    local = email.split("@", 1)[0]
    local = local.replace("_", ".").replace("%20", ".")
    parts = [p for p in re.split(r"[._-]+", local) if p and not p.isdigit() and len(p) > 1]
    if len(parts) >= 2:
        return " ".join(p.capitalize() for p in parts[:4])
    return ""


def contacts_from_text(text: str) -> tuple[str, str]:
    emails = udg.extract_emails(text, limit=3)
    phones = udg.extract_phones(text, limit=2)
    return ("; ".join(emails), phones[0] if phones else "")


def build_faculty_search_query(
    university: str,
    city: str,
    department: str,
    campus: str = "",
) -> str:
    parts: list[str] = ["site:linkedin.com/in"]
    if university:
        parts.append(f'"{university}"')
    location = campus.strip() or city.strip()
    if location:
        parts.append(f'"{location}"')
    if department:
        parts.append(f'"{department}"')
    parts.append("(professor OR lecturer OR faculty OR instructor)")
    return " ".join(parts)


def guess_faculty_urls(base: str, department: str, campus: str = "") -> list[str]:
    if not base:
        return []
    paths = [
        "/faculty",
        "/faculty-staff",
        "/our-faculty",
        "/faculty-members",
        "/people/faculty",
        "/about/faculty",
        "/academics/faculty",
        "/directory/faculty",
        "/staff",
        "/academic-staff",
        "/faculties",
    ]
    urls = [base + p for p in paths]
    if department:
        slug = re.sub(r"[^a-z0-9]+", "-", department.lower()).strip("-")
        if slug:
            urls.extend(
                [
                    f"{base}/faculty/{slug}",
                    f"{base}/departments/{slug}/faculty",
                    f"{base}/department/{slug}/faculty",
                    f"{base}/schools/{slug}/faculty",
                    f"{base}/{slug}/faculty",
                ]
            )
    if campus:
        cslug = re.sub(r"[^a-z0-9]+", "-", campus.lower()).strip("-")
        ctitle = campus.replace(" ", "")
        if cslug:
            urls.extend(
                [
                    f"{base}/campus/{cslug}/faculty",
                    f"{base}/campus/{cslug}/faculty-staff",
                    f"{base}/campuses/{cslug}/faculty",
                    f"{base}/Campus/{ctitle}/Faculty",
                    f"{base}/{cslug}/faculty",
                    f"{base}/{cslug}-campus/faculty",
                ]
            )
    return urls


def discover_faculty_urls(
    session: requests.Session,
    university: str,
    department: str,
    official_www: str,
    city: str,
    engine: str,
    max_search: int,
    campus: str = "",
) -> list[str]:
    base = normalize_www(official_www)
    official_host = host_of(base)
    preset = udg.get_university_preset(university, city)
    preset_domain = str(preset.get("domain") or "").strip()
    if not official_host and preset_domain:
        official_host = preset_domain
        if not base:
            base = f"https://www.{preset_domain}"

    seen: set[str] = set()
    candidates: list[str] = []

    def add(u: str) -> None:
        u = u.split("#")[0].rstrip("/")
        if not u.startswith("http") or u in seen:
            return
        seen.add(u)
        candidates.append(u)

    for u in guess_faculty_urls(base, department, campus):
        add(u)
    for u in preset.get("seed_urls", []):
        add(str(u))
        if "faculty" not in str(u).lower():
            add(str(u).rstrip("/") + "/Faculty")

    loc = campus.strip() or city.strip()
    queries: list[str] = []
    if official_host:
        if loc:
            queries.append(f'site:{official_host} "{loc}" faculty staff directory')
            queries.append(f'site:{official_host} "{loc}" campus lecturers professors')
        queries.append(f"site:{official_host} faculty staff directory")
        if department and loc:
            queries.append(f'site:{official_host} "{loc}" "{department}" faculty')
        if department:
            queries.append(f'site:{official_host} "{department}" faculty professors')
        queries.append(f"site:{official_host} faculty email contact")
    else:
        if loc:
            queries.append(f'"{university}" "{loc}" faculty staff Pakistan site:edu.pk')
        queries.append(f'"{university}" Pakistan faculty staff directory site:edu.pk')
        if department:
            queries.append(f'"{university}" "{department}" faculty members Pakistan')

    engine_used = ""
    for q in queries:
        try:
            urls, engine_used = udg.search_urls(session, q, max_search, engine)
            for u in urls:
                add(u)
        except Exception:
            continue

    def score_url(u: str) -> int:
        path = urlparse(u).path.lower()
        s = 0
        h = host_of(u)
        if official_host and (h == official_host or h.endswith("." + official_host)):
            s += 10
        for hint in FACULTY_URL_HINTS:
            if hint in path:
                s += 5
        if department:
            slug = re.sub(r"[^a-z0-9]+", "-", department.lower())
            if slug and slug in path:
                s += 4
        if campus:
            cslug = re.sub(r"[^a-z0-9]+", "-", campus.lower())
            if cslug and (cslug in path or campus.lower().replace(" ", "-") in path):
                s += 6
        if any(x in path for x in ("/news", "/event", "/blog", "/admission", "/jobs")):
            s -= 10
        if "linkedin.com" in host_of(u):
            s -= 30
        return s

    candidates.sort(key=score_url, reverse=True)
    return candidates


def fetch_page_html(
    session: requests.Session,
    url: str,
    *,
    via_jina: bool,
    timeout: int = 30,
) -> tuple[str, str, int]:
    try:
        if via_jina:
            title, body, status, _note = udg.fetch_page_plaintext(session, url, via_jina=True)
            return title, body, status
        resp = session.get(url, headers=HEADERS, timeout=timeout)
        resp.raise_for_status()
        html = resp.text
        soup = BeautifulSoup(html, "html.parser")
        title_tag = soup.find("title")
        title = title_tag.get_text(strip=True) if title_tag else ""
        return title, html, resp.status_code
    except Exception:
        return "", "", 0


def extract_from_mailto(soup: BeautifulSoup, page_url: str, university: str) -> list[FacultyHit]:
    out: list[FacultyHit] = []
    seen: set[str] = set()
    for a in soup.find_all("a", href=True):
        href = a.get("href", "")
        if not href.lower().startswith("mailto:"):
            continue
        email = udg._normalize_email(href) or ""
        if not email or email in seen:
            continue
        seen.add(email)
        label = udg._normalize_space(a.get_text(" ", strip=True))
        name = ""
        if label and "@" not in label and is_plausible_name(label, university):
            name = clean_name(label)
        if not name:
            name = name_from_email(email)
        parent_text = ""
        parent = a.parent
        for _ in range(4):
            if parent is None or not isinstance(parent, Tag):
                break
            parent_text = parent.get_text(" ", strip=True)
            if len(parent_text) > 20:
                break
            parent = parent.parent
        _em, phone = contacts_from_text(parent_text)
        title_m = FACULTY_ROLE_RE.search(parent_text)
        out.append(
            FacultyHit(
                faculty_name=name,
                job_title=title_m.group(1) if title_m else "",
                email=email,
                phone=phone,
                profile_page_url=page_url,
                source_url=page_url,
                snippet=parent_text[:280],
                data_source="mailto",
            )
        )
    return out


def extract_from_tables(soup: BeautifulSoup, page_url: str, university: str) -> list[FacultyHit]:
    out: list[FacultyHit] = []
    seen: set[str] = set()
    for table in soup.find_all("table"):
        headers = [th.get_text(" ", strip=True).lower() for th in table.find_all("th")]
        if not headers:
            first_row = table.find("tr")
            if first_row:
                headers = [td.get_text(" ", strip=True).lower() for td in first_row.find_all(["td", "th"])]
        name_idx = next((i for i, h in enumerate(headers) if any(x in h for x in ("name", "faculty", "member", "staff"))), -1)
        email_idx = next((i for i, h in enumerate(headers) if "email" in h or "e-mail" in h), -1)
        phone_idx = next((i for i, h in enumerate(headers) if any(x in h for x in ("phone", "tel", "mobile", "contact"))), -1)
        title_idx = next((i for i, h in enumerate(headers) if any(x in h for x in ("title", "designation", "position", "rank"))), -1)
        rows = table.find_all("tr")[1:] if headers else table.find_all("tr")
        for tr in rows:
            cells = tr.find_all(["td", "th"])
            if len(cells) < 2:
                continue
            texts = [c.get_text(" ", strip=True) for c in cells]
            row_html = tr.get_text(" ", strip=True)
            name = ""
            if 0 <= name_idx < len(texts):
                name = clean_name(texts[name_idx])
            if not name or not is_plausible_name(name, university):
                for t in texts:
                    if is_plausible_name(t, university):
                        name = clean_name(t)
                        break
            email = ""
            if 0 <= email_idx < len(texts):
                email = udg.extract_emails(texts[email_idx], limit=1)[0] if texts[email_idx] else ""
            if not email:
                emails = udg.extract_emails(row_html, limit=1)
                email = emails[0] if emails else ""
            if not name and email:
                name = name_from_email(email)
            if not name:
                continue
            phone = ""
            if 0 <= phone_idx < len(texts):
                phones = udg.extract_phones(texts[phone_idx], limit=1)
                phone = phones[0] if phones else ""
            if not phone:
                phones = udg.extract_phones(row_html, limit=1)
                phone = phones[0] if phones else ""
            job = texts[title_idx] if 0 <= title_idx < len(texts) else ""
            if not job:
                m = FACULTY_ROLE_RE.search(row_html)
                job = m.group(1) if m else ""
            key = f"{name.lower()}|{email}"
            if key in seen:
                continue
            seen.add(key)
            out.append(
                FacultyHit(
                    faculty_name=name,
                    job_title=job[:200],
                    email=email,
                    phone=phone,
                    profile_page_url=page_url,
                    source_url=page_url,
                    snippet=row_html[:280],
                    data_source="table",
                )
            )
    return out


def extract_from_cards(soup: BeautifulSoup, page_url: str, university: str) -> list[FacultyHit]:
    out: list[FacultyHit] = []
    seen: set[str] = set()
    selectors = (
        ".faculty-member",
        ".faculty-card",
        ".staff-member",
        ".team-member",
        ".profile-card",
        ".member",
        "article.profile",
        "[class*='faculty']",
        "[class*='staff-card']",
    )
    blocks: list[Tag] = []
    for sel in selectors:
        blocks.extend(soup.select(sel))
    if not blocks:
        for tag in soup.find_all(["div", "li", "article", "section"]):
            cls = " ".join(tag.get("class", [])).lower()
            if any(x in cls for x in ("faculty", "staff", "profile", "member", "people")):
                blocks.append(tag)

    for block in blocks:
        text = block.get_text(" ", strip=True)
        if len(text) < 8 or len(text) > 1200:
            continue
        emails = udg.extract_emails(str(block), limit=2)
        phones = udg.extract_phones(text, limit=1)
        name = ""
        for h in block.find_all(["h2", "h3", "h4", "strong", "b"]):
            cand = clean_name(h.get_text(" ", strip=True))
            if is_plausible_name(cand, university):
                name = cand
                break
        if not name:
            for n in udg.NAME_RE.findall(text):
                cand = udg.clean_name_candidate(n)
                if is_plausible_name(cand, university):
                    name = cand
                    break
        if not name and emails:
            name = name_from_email(emails[0])
        if not name:
            continue
        m = FACULTY_ROLE_RE.search(text)
        key = f"{name.lower()}|{emails[0] if emails else ''}"
        if key in seen:
            continue
        seen.add(key)
        profile = ""
        for a in block.find_all("a", href=True):
            href = a["href"]
            if href.startswith("mailto:"):
                continue
            full = urljoin(page_url, href)
            if is_plausible_name(a.get_text(strip=True), university) or "/faculty" in full.lower():
                profile = full
                break
        out.append(
            FacultyHit(
                faculty_name=name,
                job_title=m.group(1) if m else "",
                email=emails[0] if emails else "",
                phone=phones[0] if phones else "",
                profile_page_url=profile or page_url,
                source_url=page_url,
                snippet=text[:280],
                data_source="card",
            )
        )
    return out


def extract_from_plaintext(body: str, page_url: str, university: str, max_names: int) -> list[FacultyHit]:
    out: list[FacultyHit] = []
    seen: set[str] = set()
    lines = body.splitlines()
    for i, raw in enumerate(lines):
        line = udg._normalize_space(raw)
        if not line:
            continue
        emails = udg.extract_emails(line, limit=2)
        phones = udg.extract_phones(line, limit=1)
        role_m = FACULTY_ROLE_RE.search(line)
        start = max(0, i - 2)
        end = min(len(lines), i + 3)
        context = " | ".join(udg._normalize_space(x) for x in lines[start:end] if udg._normalize_space(x))
        if not emails and not role_m:
            continue
        name = ""
        for n in udg.NAME_RE.findall(context):
            cand = udg.clean_name_candidate(n)
            if is_plausible_name(cand, university):
                name = cand
                break
        if not name and emails:
            name = name_from_email(emails[0])
        if not name:
            continue
        if not emails:
            emails = udg.extract_emails(context, limit=1)
        if not phones:
            phones = udg.extract_phones(context, limit=1)
        key = f"{name.lower()}|{emails[0] if emails else ''}"
        if key in seen:
            continue
        seen.add(key)
        out.append(
            FacultyHit(
                faculty_name=name,
                job_title=role_m.group(1) if role_m else "",
                email=emails[0] if emails else "",
                phone=phones[0] if phones else "",
                profile_page_url=page_url,
                source_url=page_url,
                snippet=context[:280],
                data_source="plaintext",
            )
        )
        if len(out) >= max_names:
            break
    return out


def collect_same_site_links(soup: BeautifulSoup, page_url: str, official_host: str, limit: int) -> list[str]:
    if not official_host:
        return []
    found: list[str] = []
    seen: set[str] = {page_url.rstrip("/")}
    for a in soup.find_all("a", href=True):
        full = urljoin(page_url, a["href"]).split("#")[0].rstrip("/")
        if full in seen or not full.startswith("http"):
            continue
        h = host_of(full)
        if h != official_host and not h.endswith("." + official_host):
            continue
        path = urlparse(full).path.lower()
        if any(hint in path for hint in FACULTY_URL_HINTS):
            seen.add(full)
            found.append(full)
        if len(found) >= limit:
            break
    return found


def scrape_faculty_pages(
    session: requests.Session,
    university: str,
    department: str,
    official_www: str,
    city: str,
    engine: str,
    *,
    campus: str = "",
    max_pages: int,
    max_per_page: int,
    via_jina: bool,
    fetch_delay: float,
    crawl_links: bool,
) -> tuple[list[dict[str, str]], str]:
    candidates = discover_faculty_urls(
        session,
        university,
        department,
        official_www,
        city,
        engine,
        max_search=max(12, max_pages),
        campus=campus,
    )
    if not candidates:
        return [], ""

    official_host = host_of(normalize_www(official_www))
    to_fetch: list[str] = []
    seen_pages: set[str] = set()

    def queue(url: str) -> None:
        u = url.split("#")[0].rstrip("/")
        if u not in seen_pages and u.startswith("http"):
            seen_pages.add(u)
            to_fetch.append(u)

    for u in candidates[:max_pages]:
        queue(u)

    all_hits: list[FacultyHit] = []
    fetched = 0
    idx = 0
    while idx < len(to_fetch) and fetched < max_pages:
        page_url = to_fetch[idx]
        idx += 1
        fetched += 1
        time.sleep(max(fetch_delay, 0))
        title, html, status = fetch_page_html(session, page_url, via_jina=via_jina)
        if status != 200 or not html:
            continue
        if via_jina:
            page_hits = extract_from_plaintext(html, page_url, university, max_per_page)
        else:
            soup = BeautifulSoup(html, "html.parser")
            page_hits: list[FacultyHit] = []
            page_hits.extend(extract_from_mailto(soup, page_url, university))
            page_hits.extend(extract_from_tables(soup, page_url, university))
            page_hits.extend(extract_from_cards(soup, page_url, university))
            page_hits.extend(extract_from_plaintext(soup.get_text("\n", strip=True), page_url, university, max_per_page))
            if crawl_links and len(to_fetch) < max_pages + 10:
                for link in collect_same_site_links(soup, page_url, official_host, limit=8):
                    queue(link)
        all_hits.extend(page_hits)

    rows = [h.to_dict() for h in all_hits]
    return rows, "website_scrape"


# --- LinkedIn search (optional supplement) ---


def clean_search_url(href: str) -> str:
    if not href:
        return ""
    parsed = urlparse(href)
    if "duckduckgo.com" in parsed.netloc.lower() and parsed.path.startswith("/l/"):
        qs = parse_qs(parsed.query)
        target = qs.get("uddg", [None])[0]
        if target:
            return unquote(target)
    if href.startswith("//"):
        return "https:" + href
    return href


def is_linkedin_profile(url: str) -> bool:
    return udg.is_profile_like_linkedin(url)


def name_from_title(title: str) -> str:
    cleaned = re.sub(r"\s*\|\s*LinkedIn\s*$", "", title, flags=re.IGNORECASE).strip()
    for sep in (" - ", " – ", " | "):
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


def infer_job_title(title: str, snippet: str) -> str:
    text = f"{title} {snippet}"
    m = FACULTY_ROLE_RE.search(text)
    if m:
        return m.group(1).strip()
    return ""


def search_profiles_ddgs(query: str, top_n: int) -> list[dict[str, str]]:
    if DDGS is None:
        return []
    rows: list[dict[str, str]] = []
    seen: set[str] = set()
    with DDGS() as client:
        for item in client.text(query, region="wt-wt", safesearch="off", max_results=top_n * 5):
            profile_url = clean_search_url((item.get("href") or item.get("url") or "").strip())
            if not profile_url or not is_linkedin_profile(profile_url) or profile_url in seen:
                continue
            seen.add(profile_url)
            title = (item.get("title") or "").strip()
            snippet = (item.get("body") or item.get("snippet") or "").strip()
            em, ph = contacts_from_text(f"{title} {snippet}")
            rows.append(
                {
                    "faculty_name": name_from_title(title),
                    "job_title": infer_job_title(title, snippet),
                    "email": em,
                    "phone": ph,
                    "linkedin_url": profile_url,
                    "result_title": title,
                    "snippet": snippet,
                    "data_source": "linkedin_ddgs",
                }
            )
            if len(rows) >= top_n:
                break
    return rows


def search_profiles_ddg_html(session: requests.Session, query: str, top_n: int) -> list[dict[str, str]]:
    resp = session.post(
        udg.DDG_HTML_URL,
        data={"q": query, "b": ""},
        headers={**HEADERS, "Referer": "https://html.duckduckgo.com/"},
        timeout=25,
    )
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    rows: list[dict[str, str]] = []
    seen: set[str] = set()
    for result in soup.select("div.result"):
        link_tag = result.select_one("a.result__a, a.result-link, a[data-testid='result-title-a']")
        if link_tag is None:
            continue
        profile_url = clean_search_url((link_tag.get("href") or "").strip())
        if not profile_url or not is_linkedin_profile(profile_url) or profile_url in seen:
            continue
        seen.add(profile_url)
        title = link_tag.get_text(" ", strip=True)
        snippet_tag = result.select_one(".result__snippet")
        snippet = snippet_tag.get_text(" ", strip=True) if snippet_tag else ""
        em, ph = contacts_from_text(f"{title} {snippet}")
        rows.append(
            {
                "faculty_name": name_from_title(title),
                "job_title": infer_job_title(title, snippet),
                "email": em,
                "phone": ph,
                "linkedin_url": profile_url,
                "result_title": title,
                "snippet": snippet,
                "data_source": "linkedin_ddg_html",
            }
        )
        if len(rows) >= top_n:
            break
    return rows


def search_faculty_profiles(
    session: requests.Session,
    query: str,
    engine: str,
    top_n: int,
) -> tuple[list[dict[str, str]], str]:
    top_n = max(1, top_n)
    for eng in (["ddgs", "duckduckgo_html"] if engine == "auto" else [engine]):
        try:
            if eng == "ddgs":
                rows = search_profiles_ddgs(query, top_n)
                if rows:
                    return rows, "ddgs"
            elif eng == "duckduckgo_html":
                rows = search_profiles_ddg_html(session, query, top_n)
                if rows:
                    return rows, "duckduckgo_html"
        except Exception:
            continue
    return [], ""


def profile_key(row: dict[str, str]) -> str:
    email = row.get("email", "").strip().lower()
    url = row.get("linkedin_url", "").strip().lower()
    name = row.get("faculty_name", "").strip().lower()
    if email:
        return f"e:{email}"
    if url:
        return f"u:{url}"
    if name:
        return f"n:{name}"
    return ""


def merge_row(existing: dict[str, str], new: dict[str, str]) -> None:
    for field in (
        "faculty_name",
        "job_title",
        "email",
        "phone",
        "linkedin_url",
        "profile_page_url",
        "result_title",
        "snippet",
        "source_url",
        "data_source",
    ):
        if not str(existing.get(field, "")).strip() and str(new.get(field, "")).strip():
            existing[field] = new[field]


def dedupe_profiles(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    by_key: dict[str, dict[str, str]] = {}
    for row in rows:
        k = profile_key(row)
        if not k:
            continue
        if k in by_key:
            merge_row(by_key[k], row)
        else:
            by_key[k] = dict(row)
    return list(by_key.values())


def base_record(
    *,
    iau_id: str,
    university: str,
    city: str,
    whed_city: str,
    campus: str,
    division_type: str,
    department: str,
    fields_of_study: str,
    official_www: str,
    search_query: str,
    search_engine: str,
) -> dict[str, str]:
    return {
        "iau_id": iau_id,
        "university_name": university,
        "city": city,
        "whed_city": whed_city,
        "campus": campus,
        "division_type": division_type,
        "department": department,
        "fields_of_study": fields_of_study,
        "official_www": official_www,
        "search_query": search_query,
        "search_engine": search_engine,
        "faculty_name": "",
        "job_title": "",
        "email": "",
        "phone": "",
        "linkedin_url": "",
        "profile_page_url": "",
        "result_title": "",
        "snippet": "",
        "source_url": "",
        "data_source": "",
        "profile_rank": "",
        "error": "",
        "scraped_at": datetime.now().isoformat(timespec="seconds"),
    }


def profile_from_hit(hit: dict[str, str], meta: dict[str, str], rank: int) -> dict[str, str]:
    title = hit.get("result_title", "")
    snippet = hit.get("snippet", "")
    faculty_name = (
        hit.get("faculty_name", "")
        or name_from_title(title)
        or name_from_linkedin_url(hit.get("linkedin_url", ""))
    )
    em = hit.get("email", "")
    ph = hit.get("phone", "")
    if not em or not ph:
        extra_em, extra_ph = contacts_from_text(f"{title} {snippet}")
        em = em or extra_em
        ph = ph or extra_ph
    row = base_record(
        iau_id=meta["iau_id"],
        university=meta["university"],
        city=meta["city"],
        whed_city=meta.get("whed_city", ""),
        campus=meta.get("campus", ""),
        division_type=meta["division_type"],
        department=meta["department"],
        fields_of_study=meta["fields_of_study"],
        official_www=meta["official_www"],
        search_query=meta["search_query"],
        search_engine=meta["search_engine"],
    )
    source = hit.get("data_source") or (
        "university_page" if hit.get("source_url") and not hit.get("linkedin_url") else "linkedin_search"
    )
    row.update(
        {
            "faculty_name": faculty_name,
            "job_title": hit.get("job_title", "") or infer_job_title(title, snippet),
            "email": em,
            "phone": ph,
            "linkedin_url": hit.get("linkedin_url", ""),
            "profile_page_url": hit.get("profile_page_url", hit.get("source_url", "")),
            "result_title": title,
            "snippet": snippet[:500],
            "source_url": hit.get("source_url", hit.get("linkedin_url", "")),
            "data_source": source,
            "profile_rank": str(rank),
            "error": hit.get("error", ""),
        }
    )
    return row


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Scrape Pakistan university faculty (names, email, phone, LinkedIn)."
    )
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--engine", choices=("auto", "ddgs", "duckduckgo_html", "jina_ddg"), default="auto")
    parser.add_argument("--top-profiles", type=int, default=25, help="Max LinkedIn profiles per department")
    parser.add_argument("--delay-seconds", type=float, default=1.0)
    parser.add_argument("--max-institutions", type=int, default=0)
    parser.add_argument("--max-departments", type=int, default=0)
    parser.add_argument("--university", action="append", default=[], metavar="NAME")
    parser.add_argument("--iau-id", action="append", default=[], metavar="ID")
    parser.add_argument(
        "--scrape-pages",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Scrape official faculty web pages (default: on; best for email/phone)",
    )
    parser.add_argument("--page-max", type=int, default=15, help="Faculty listing pages to fetch per department")
    parser.add_argument("--max-per-page", type=int, default=80, help="Max people extracted per page")
    parser.add_argument("--via-jina", action="store_true", help="Fetch pages via r.jina.ai")
    parser.add_argument("--no-crawl", action="store_true", help="Do not follow faculty links on same site")
    parser.add_argument(
        "--linkedin-search",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Also search LinkedIn (default: on)",
    )
    parser.add_argument(
        "--resume",
        action="store_true",
        help="Same as default append: skip departments already saved in the output file",
    )
    parser.add_argument(
        "--fresh",
        action="store_true",
        help="Do not load existing Excel; overwrite output from scratch",
    )
    parser.add_argument(
        "--campus-wise",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Scrape each city/campus x department (default: on; needed for Lahore/Faisalabad etc.)",
    )
    parser.add_argument(
        "--no-campus-presets",
        action="store_true",
        help="Do not add known city list for Riphah/Bahria/FAST/KMU etc.",
    )
    parser.add_argument(
        "--all-campuses",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Use presets + search/website discovery for ALL campus cities (default: on)",
    )
    parser.add_argument(
        "--extra-campuses",
        action="append",
        default=[],
        metavar="CITY",
        help='Extra campus cities, e.g. --extra-campuses Faisalabad --extra-campuses Multan',
    )
    parser.add_argument(
        "--campus-only",
        action="append",
        default=[],
        metavar="CITY",
        help="Only these campuses (case-insensitive), e.g. --campus-only Lahore --campus-only Faisalabad",
    )
    parser.add_argument(
        "--repair-cities",
        action="store_true",
        help="Fix existing Excel: set city=campus and whed_city=old HQ city; then exit (no scraping)",
    )
    args = parser.parse_args()

    inp = args.input.resolve()
    out_path = args.output.resolve()

    if args.repair_cities:
        return repair_excel_file(out_path)
    if not inp.is_file():
        print(f"Input not found: {inp}")
        return 1

    df = pd.read_csv(inp, dtype=str, keep_default_na=False)
    if "name" not in df.columns or "country" not in df.columns:
        print("CSV missing required columns: name, country")
        return 1

    pakistan = df[df["country"].str.strip().str.lower() == "pakistan"].copy()
    if args.university:
        needles = [u.strip().lower() for u in args.university if u.strip()]
        pakistan = pakistan[
            pakistan["name"].astype(str).str.lower().apply(lambda n: any(x in n for x in needles))
        ]
    if args.iau_id:
        allowed = {i.strip().upper() for i in args.iau_id if i.strip()}
        if allowed and "iau_id" in pakistan.columns:
            pakistan = pakistan[pakistan["iau_id"].astype(str).str.strip().str.upper().isin(allowed)]
    if pakistan.empty:
        print("No universities matched filters.")
        return 1

    if args.university or args.iau_id:
        print(f"Matched: {', '.join(str(r['name']) for _, r in pakistan.iterrows())}")

    # Append to existing workbook by default (keeps e.g. Aga Khan when scraping Riphah next).
    done_depts: set[str] = set()
    rows_out: list[dict[str, str]] = []
    if not args.fresh and out_path.is_file():
        try:
            prior = pd.read_excel(out_path, dtype=str, engine="openpyxl")
            rows_out = prior.fillna("").astype(str).to_dict(orient="records")
            rows_out, n_repaired = repair_rows_campus_cities(rows_out)
            if n_repaired:
                save_progress(rows_out, out_path)
                print(f"Repaired city/whed_city on {n_repaired} existing rows (campus-wise)")
            done_depts = load_existing_scrape_keys(out_path)
            print(f"Appending to existing file: {len(rows_out)} rows, {len(done_depts)} scrape keys on record")
        except Exception as exc:
            print(f"Could not read existing output ({exc}); starting with empty sheet")
    elif args.fresh:
        print("Fresh run: existing output will be replaced when the first batch saves")

    session = requests.Session()
    session.headers.update(HEADERS)

    n_inst = 0
    for _, row in pakistan.iterrows():
        if args.max_institutions and n_inst >= args.max_institutions:
            break
        uni_name = str(row.get("name", "")).strip()
        if not uni_name:
            continue
        if not str(row.get("www", "")).strip():
            print(f"\nSkipping {uni_name}: no www in CSV")
            continue

        iau = str(row.get("iau_id", "")).strip()
        city = str(row.get("city", "")).strip()
        official_www = str(row.get("www", "")).strip()

        targets = build_scrape_targets(
            university=uni_name,
            main_city=city,
            divisions_json=str(row.get("divisions_json", "")),
            extra_campuses=args.extra_campuses,
            campus_wise=args.campus_wise,
            use_campus_presets=not args.no_campus_presets,
            discover_all_campuses=args.all_campuses,
            session=session,
            official_www=official_www,
            engine=args.engine,
            via_jina=args.via_jina,
            max_departments=args.max_departments,
        )
        if args.campus_only:
            allowed = {c.strip().lower() for c in args.campus_only if c.strip()}
            targets = [t for t in targets if t["campus"].lower() in allowed]

        campuses_in_run = sorted({t["campus"] for t in targets if t["campus"]})
        n_inst += 1
        mode = "campus-wise" if args.campus_wise else "department-only"
        print(f"\n[{n_inst}] {uni_name} — {len(targets)} scrape units ({mode})")
        if campuses_in_run:
            print(f"  Campuses: {', '.join(campuses_in_run)}")

        for target in targets:
            department = target["department"]
            campus = target["campus"]
            search_city = target["search_city"]
            dept_key = scrape_key(iau, uni_name, campus, department or "(university-wide)")
            if dept_key in done_depts:
                label = f"{campus} / {department}" if campus else (department or "all")
                print(f"  Skip: {label}")
                continue

            scrape_city = search_city.strip() or campus.strip() or city.strip()
            meta = {
                "iau_id": iau,
                "university": uni_name,
                "city": scrape_city,
                "whed_city": city,
                "campus": campus,
                "division_type": target["division_type"],
                "department": department,
                "fields_of_study": target["fields_of_study"],
                "official_www": official_www,
                "search_query": "",
                "search_engine": "",
            }
            hits: list[dict[str, str]] = []
            err_parts: list[str] = []

            if args.scrape_pages:
                loc_label = f"{campus} — {department}" if campus else (department or "all")
                print(f"  Website: {loc_label} ...")
                time.sleep(max(args.delay_seconds, 0))
                try:
                    page_hits, eng = scrape_faculty_pages(
                        session,
                        uni_name,
                        department,
                        official_www,
                        city,
                        args.engine,
                        campus=campus,
                        max_pages=args.page_max,
                        max_per_page=args.max_per_page,
                        via_jina=args.via_jina,
                        fetch_delay=args.delay_seconds,
                        crawl_links=not args.no_crawl,
                    )
                    hits.extend(page_hits)
                    meta["search_engine"] = eng
                    print(f"    website -> {len(page_hits)} people")
                except Exception as exc:
                    err_parts.append(f"website:{exc}")

            if args.linkedin_search:
                search_query = build_faculty_search_query(
                    uni_name, city, department, campus=search_city
                )
                meta["search_query"] = search_query
                time.sleep(max(args.delay_seconds, 0))
                try:
                    li_hits, li_eng = search_faculty_profiles(
                        session, search_query, args.engine, args.top_profiles
                    )
                    hits.extend(li_hits)
                    if li_eng:
                        meta["search_engine"] = (meta["search_engine"] + "+" + li_eng).strip("+")
                    print(f"    LinkedIn -> {len(li_hits)} profiles")
                except Exception as exc:
                    err_parts.append(f"linkedin:{exc}")

            hits = dedupe_profiles(hits)
            err = "; ".join(err_parts)

            new_rows: list[dict[str, str]] = []
            if hits:
                for rank, hit in enumerate(hits, start=1):
                    new_rows.append(profile_from_hit(hit, meta, rank))
                with_email = sum(1 for r in new_rows if r.get("email"))
                preview = ", ".join(r["faculty_name"] for r in new_rows[:4] if r["faculty_name"])
                print(f"  Total {len(new_rows)} faculty ({with_email} with email) — {preview}...")
            else:
                placeholder = base_record(
                    iau_id=iau,
                    university=uni_name,
                    city=scrape_city,
                    whed_city=city,
                    campus=campus,
                    division_type=target["division_type"],
                    department=department,
                    fields_of_study=target["fields_of_study"],
                    official_www=official_www,
                    search_query=meta["search_query"],
                    search_engine=meta["search_engine"],
                )
                placeholder["error"] = err or "No faculty found"
                new_rows.append(placeholder)
                print(f"  No faculty ({placeholder['error']})")

            rows_out.extend(new_rows)
            done_depts.add(dept_key)
            save_progress(rows_out, out_path)

    rows_out, n_final = repair_rows_campus_cities(rows_out)
    if n_final:
        save_progress(rows_out, out_path)
        print(f"Final pass: normalized city/campus on {n_final} rows")

    print(f"\nDone. {len(rows_out)} rows -> {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
