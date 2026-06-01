"""
Pakistan university faculty scraper — names, titles, email, phone, LinkedIn.

Primary source: official university faculty/staff web pages (mailto, tables, cards).
Optional: LinkedIn search per department.

Run:
  python scripts/pakistan_faculty_data.py --university "Aga Khan University" --max-institutions 1
  python scripts/pakistan_faculty_data.py --university "Riphah" --max-institutions 1
  python scripts/pakistan_faculty_data.py --university NUTECH --max-institutions 1
    (if not in data-raw/institutions.csv, loads data-raw/NUTECH.xlsx)
  (second run appends to output/pakistan_faculty.xlsx automatically)

  # All Pakistan universities from institutions.csv — one Excel file per university:
  python scripts/pakistan_faculty_data.py --per-university
  python scripts/pakistan_faculty_data.py --per-university --max-institutions 5
  python scripts/pakistan_faculty_data.py --per-university --skip-institutions 10 --max-institutions 20
    (skip first 10 in CSV, then scrape the next 20 universities)
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
DEFAULT_FALLBACK_INPUT = PROJECT_ROOT / "data-raw" / "NUTECH.xlsx"
DEFAULT_OUTPUT = PROJECT_ROOT / "output" / "pakistan_faculty.xlsx"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "output" / "faculty_by_university"

# institutions.csv column mapping (main input — unchanged).
INSTITUTION_CSV_COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "name": ("name", "university", "university_name", "institution", "institution_name"),
    "country": ("country",),
    "city": ("city",),
    "www": ("www", "website", "url", "official_www", "web"),
    "iau_id": ("iau_id", "id"),
    "divisions_json": ("divisions_json", "divisions"),
    "department": ("department", "dept", "division", "school", "faculty"),
    "division_type": ("division_type", "type", "division type"),
    "fields_of_study": ("fields_of_study", "fields", "field_of_study", "programs"),
}

# uniRank-style fallback sheets (e.g. NUTECH.xlsx) — only when CSV has no match.
FALLBACK_EXCEL_EXTRA_ALIASES: dict[str, tuple[str, ...]] = {
    "name": ("university name",),
    "www": ("official website",),
    "city": ("city",),
    "acronym": ("identity: acronym", "acronym"),
}


def _merged_column_aliases(
    base: dict[str, tuple[str, ...]], extra: dict[str, tuple[str, ...]]
) -> dict[str, tuple[str, ...]]:
    merged: dict[str, tuple[str, ...]] = {}
    for key, base_aliases in base.items():
        merged[key] = base_aliases + extra.get(key, ())
    for key, extra_aliases in extra.items():
        if key not in merged:
            merged[key] = extra_aliases
    return merged


FALLBACK_EXCEL_COLUMN_ALIASES = _merged_column_aliases(
    INSTITUTION_CSV_COLUMN_ALIASES, FALLBACK_EXCEL_EXTRA_ALIASES
)

UNIVERSITY_MATCH_ALIASES: dict[str, tuple[str, ...]] = {
    "nutech": ("nutech", "national university of technology"),
    "national university of technology": ("nutech", "national university of technology"),
}

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


def _header_lookup(
    columns: list[str],
    column_aliases: dict[str, tuple[str, ...]],
) -> dict[str, str]:
    by_lower = {str(c).strip().lower(): str(c) for c in columns}
    lookup: dict[str, str] = {}
    for canonical, aliases in column_aliases.items():
        for alias in aliases:
            key = alias.strip().lower()
            if key in by_lower:
                lookup[canonical] = by_lower[key]
                break
        if canonical in lookup:
            continue
        for col in columns:
            col_lower = str(col).strip().lower()
            for alias in aliases:
                alias_lower = alias.strip().lower()
                # Avoid short tokens (e.g. "name") matching "Identity: Name (English)".
                if len(alias_lower) < 5:
                    continue
                if alias_lower in col_lower or col_lower in alias_lower:
                    lookup[canonical] = str(col)
                    break
            if canonical in lookup:
                break
    return lookup


def normalize_institutions_frame(
    df: pd.DataFrame,
    *,
    column_aliases: dict[str, tuple[str, ...]] | None = None,
) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(
            columns=[
                "name",
                "country",
                "city",
                "www",
                "iau_id",
                "divisions_json",
                "department",
                "division_type",
                "fields_of_study",
            ]
        )
    aliases = column_aliases or INSTITUTION_CSV_COLUMN_ALIASES
    lookup = _header_lookup(list(df.columns), aliases)
    out = pd.DataFrame(index=df.index)
    for col in (
        "name",
        "country",
        "city",
        "www",
        "iau_id",
        "divisions_json",
        "department",
        "division_type",
        "fields_of_study",
        "acronym",
    ):
        if col == "acronym" and col not in aliases:
            continue
        src = lookup.get(col)
        out[col] = df[src].astype(str) if src else ""
    out = out.fillna("").astype(str)
    junk_names = frozenset({"", "n.a.", "na", "n/a", "nan", "none", "not reported"})
    out["name"] = out["name"].apply(
        lambda v: "" if str(v).strip().lower() in junk_names else str(v).strip()
    )
    if out["country"].str.strip().eq("").all():
        out["country"] = "Pakistan"
    empty_www = out["www"].str.strip().eq("")
    nutech_mask = out["name"].str.lower().str.contains("nutech", na=False)
    out.loc[empty_www & nutech_mask, "www"] = "https://nutech.edu.pk"
    empty_city = out["city"].str.strip().eq("")
    out.loc[empty_city & nutech_mask, "city"] = "Islamabad"
    empty_name = out["name"].str.strip().eq("")
    out.loc[empty_name & nutech_mask, "name"] = "National University of Technology (NUTECH)"
    return out


def collapse_institution_rows(
    df: pd.DataFrame,
    *,
    column_aliases: dict[str, tuple[str, ...]] | None = None,
) -> pd.DataFrame:
    df = normalize_institutions_frame(df, column_aliases=column_aliases)
    has_department = df["department"].str.strip().ne("").any()
    if not has_department:
        return df.drop(
            columns=["department", "division_type", "fields_of_study", "acronym"],
            errors="ignore",
        )

    group_cols = ["name"]
    if df["iau_id"].str.strip().ne("").any():
        group_cols.append("iau_id")

    merged_rows: list[dict[str, str]] = []
    for _, grp in df.groupby(group_cols, sort=False):
        base = grp.iloc[0].to_dict()
        divisions_json = str(base.get("divisions_json", "")).strip()
        if divisions_json and divisions_json != "[]":
            merged_rows.append({k: str(v) for k, v in base.items()})
            continue
        divisions: list[dict[str, str]] = []
        for _, row in grp.iterrows():
            dept = str(row.get("department", "")).strip()
            if not dept:
                continue
            divisions.append(
                {
                    "type": str(row.get("division_type", "")).strip() or "Department/Division",
                    "name": dept,
                    "fields_of_study": str(row.get("fields_of_study", "")).strip(),
                }
            )
        base["divisions_json"] = json.dumps(divisions) if divisions else ""
        merged_rows.append({k: str(v) for k, v in base.items()})

    collapsed = pd.DataFrame(merged_rows)
    return collapsed.drop(
        columns=["department", "division_type", "fields_of_study", "acronym"],
        errors="ignore",
    )


def resolve_data_path(path: Path) -> Path:
    path = Path(path).expanduser()
    if path.is_file():
        return path.resolve()
    for base in (Path.cwd(), PROJECT_ROOT):
        candidate = (base / path).resolve()
        if candidate.is_file():
            return candidate
    return path.resolve()


def expand_university_needles(needles: list[str]) -> list[str]:
    expanded: list[str] = []
    for raw in needles:
        key = raw.strip().lower()
        if not key:
            continue
        expanded.append(key)
        for anchor, aliases in UNIVERSITY_MATCH_ALIASES.items():
            if anchor in key or key in anchor or any(a in key for a in aliases):
                expanded.extend(aliases)
    seen: set[str] = set()
    unique: list[str] = []
    for item in expanded:
        if item not in seen:
            seen.add(item)
            unique.append(item)
    return unique


def university_name_matches(institution_name: str, needles: list[str]) -> bool:
    name = str(institution_name).strip().lower()
    if not name:
        return False
    for needle in expand_university_needles(needles):
        if needle in name or name in needle:
            return True
    return False


def load_institutions_csv(path: Path) -> pd.DataFrame:
    """Load institutions.csv (original path — CSV aliases only)."""
    path = resolve_data_path(path)
    raw = pd.read_csv(path, dtype=str, keep_default_na=False)
    return collapse_institution_rows(raw, column_aliases=INSTITUTION_CSV_COLUMN_ALIASES)


def load_fallback_excel(path: Path) -> pd.DataFrame:
    """Load fallback Excel (e.g. NUTECH.xlsx) with uniRank-style column names."""
    path = resolve_data_path(path)
    xl = pd.ExcelFile(path, engine="openpyxl")
    frames: list[pd.DataFrame] = []
    for sheet in xl.sheet_names:
        raw = pd.read_excel(
            path, sheet_name=sheet, dtype=str, engine="openpyxl", keep_default_na=False
        )
        if raw.empty:
            continue
        collapsed = collapse_institution_rows(raw, column_aliases=FALLBACK_EXCEL_COLUMN_ALIASES)
        if not collapsed.empty:
            frames.append(collapsed)
    if not frames:
        return normalize_institutions_frame(
            pd.DataFrame(), column_aliases=FALLBACK_EXCEL_COLUMN_ALIASES
        )
    return pd.concat(frames, ignore_index=True)


def filter_pakistan(df: pd.DataFrame, *, lenient: bool = False) -> pd.DataFrame:
    if df.empty:
        return df
    country = df["country"].astype(str).str.strip().str.lower()
    country = country.replace({"nan": "", "none": ""})
    if lenient:
        mask = country.eq("pakistan") | country.eq("")
        return df[mask].copy()
    return df[country.eq("pakistan")].copy()


def filter_universities(df: pd.DataFrame, needles: list[str]) -> pd.DataFrame:
    if not needles:
        return df

    def row_matches(row: pd.Series) -> bool:
        if university_name_matches(str(row.get("name", "")), needles):
            return True
        if "acronym" in row.index:
            return university_name_matches(str(row.get("acronym", "")), needles)
        return False

    return df[df.apply(row_matches, axis=1)]


def apply_university_label(df: pd.DataFrame, university_filters: list[str]) -> pd.DataFrame:
    label = next((u.strip() for u in university_filters if u.strip()), "")
    if not label or df.empty:
        return df
    out = df.copy()
    empty = out["name"].astype(str).str.strip().eq("")
    out.loc[empty, "name"] = label
    return out


def apply_fallback_file_defaults(
    df: pd.DataFrame,
    fallback_path: Path,
    university_filters: list[str],
) -> pd.DataFrame:
    out = apply_university_label(df, university_filters)
    if out.empty:
        return out
    stem = fallback_path.stem.lower()
    nutech_context = "nutech" in stem or any(
        "nutech" in u.lower() or "national university of technology" in u.lower()
        for u in university_filters
    )
    if nutech_context:
        empty_www = out["www"].astype(str).str.strip().eq("")
        out.loc[empty_www, "www"] = "https://nutech.edu.pk"
        empty_city = out["city"].astype(str).str.strip().eq("")
        out.loc[empty_city, "city"] = "Islamabad"
        empty_name = out["name"].astype(str).str.strip().eq("")
        out.loc[empty_name, "name"] = "National University of Technology (NUTECH)"
    return out


def filter_iau_ids(df: pd.DataFrame, iau_ids: list[str]) -> pd.DataFrame:
    allowed = {i.strip().upper() for i in iau_ids if i.strip()}
    if not allowed or "iau_id" not in df.columns:
        return df
    return df[df["iau_id"].astype(str).str.strip().str.upper().isin(allowed)]


def resolve_pakistan_institutions(
    csv_path: Path,
    fallback_path: Path | None,
    university_filters: list[str],
    iau_filters: list[str],
) -> pd.DataFrame:
    csv_path = resolve_data_path(csv_path)
    if not csv_path.is_file():
        print(f"Input not found: {csv_path}")
        pakistan = pd.DataFrame()
    else:
        pakistan = filter_pakistan(load_institutions_csv(csv_path))

    selected = pakistan
    if university_filters:
        selected = filter_universities(pakistan, university_filters)
    if iau_filters:
        selected = filter_iau_ids(selected, iau_filters)

    need_fallback = bool(university_filters) and selected.empty
    if need_fallback and fallback_path:
        fallback_path = resolve_data_path(fallback_path)
    if need_fallback and fallback_path and fallback_path.is_file():
        fb = filter_pakistan(load_fallback_excel(fallback_path), lenient=True)
        fb = apply_fallback_file_defaults(fb, fallback_path, university_filters)
        fb_selected = filter_universities(fb, university_filters)
        if fb_selected.empty and not fb.empty:
            # Dedicated fallback workbook (e.g. NUTECH.xlsx): use all rows in the file.
            fb_selected = fb.copy()
            print(
                f"University not in {csv_path.name}; using all rows from fallback: {fallback_path}"
            )
        elif not fb_selected.empty:
            print(f"University not in {csv_path.name}; using fallback: {fallback_path}")
        if iau_filters:
            fb_selected = filter_iau_ids(fb_selected, iau_filters)
        if not fb_selected.empty:
            selected = fb_selected
    elif need_fallback:
        hint = resolve_data_path(fallback_path) if fallback_path else DEFAULT_FALLBACK_INPUT
        print(f"No match in {csv_path.name}. Place institution rows in: {hint}")

    return selected


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


def department_key(iau_id: str, university: str, department: str) -> str:
    base = iau_id.strip() or university.strip()
    dept = department.strip().lower() or "(university-wide)"
    return f"{base}|{dept}"


def load_existing_department_keys(output_path: Path) -> set[str]:
    if not output_path.is_file():
        return set()
    try:
        existing = pd.read_excel(output_path, dtype=str, engine="openpyxl")
    except Exception:
        return set()
    keys: set[str] = set()
    for _, row in existing.iterrows():
        keys.add(
            department_key(
                str(row.get("iau_id", "")),
                str(row.get("university_name", "")),
                str(row.get("department", "")),
            )
        )
    return keys


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


def slugify_university_filename(name: str, max_len: int = 60) -> str:
    slug = re.sub(r"[^a-zA-Z0-9]+", "_", name.strip()).strip("_")
    return slug[:max_len] if slug else "university"


def university_output_path(output_dir: Path, university: str, iau_id: str) -> Path:
    slug = slugify_university_filename(university)
    iau = iau_id.strip()
    fname = f"{iau}_{slug}.xlsx" if iau else f"{slug}.xlsx"
    return output_dir / fname


def load_output_state(
    output_path: Path, *, fresh: bool
) -> tuple[list[dict[str, str]], set[str]]:
    if fresh or not output_path.is_file():
        return [], set()
    try:
        prior = pd.read_excel(output_path, dtype=str, engine="openpyxl")
        rows = prior.fillna("").astype(str).to_dict(orient="records")
        keys = load_existing_department_keys(output_path)
        return rows, keys
    except Exception as exc:
        print(f"  Could not read existing output ({exc}); starting fresh for this file")
        return [], set()


def scrape_university(
    session: requests.Session,
    row: pd.Series,
    *,
    args: argparse.Namespace,
    output_path: Path,
    rows_out: list[dict[str, str]] | None = None,
    done_depts: set[str] | None = None,
) -> tuple[list[dict[str, str]], set[str], int]:
    """Scrape all departments for one university; save Excel after each department."""
    uni_name = str(row.get("name", "")).strip()
    if not uni_name:
        return rows_out or [], done_depts or set(), 0
    if not str(row.get("www", "")).strip():
        print(f"\nSkipping {uni_name}: no www in CSV")
        return rows_out or [], done_depts or set(), 0

    iau = str(row.get("iau_id", "")).strip()
    city = str(row.get("city", "")).strip()
    official_www = str(row.get("www", "")).strip()

    divisions = parse_divisions(row.get("divisions_json", ""))
    if not divisions:
        divisions = [{"division_type": "", "department": "", "fields_of_study": ""}]
    if args.max_departments:
        divisions = divisions[: args.max_departments]

    if rows_out is None or done_depts is None:
        rows_out, done_depts = load_output_state(output_path, fresh=args.fresh)
    else:
        rows_out = list(rows_out)
        done_depts = set(done_depts)

    print(f"\n{uni_name} ({len(divisions)} departments)")
    print(f"  -> {output_path}")

    for div in divisions:
        department = div["department"]
        dept_key = department_key(iau, uni_name, department or "(university-wide)")
        if dept_key in done_depts:
            print(f"  Skip: {department or '(university-wide)'}")
            continue

        meta = {
            "iau_id": iau,
            "university": uni_name,
            "city": city,
            "division_type": div["division_type"],
            "department": department,
            "fields_of_study": div["fields_of_study"],
            "official_www": official_www,
            "search_query": "",
            "search_engine": "",
        }
        hits: list[dict[str, str]] = []
        err_parts: list[str] = []

        if args.scrape_pages:
            print(f"  Website scrape: {department or 'all'} ...")
            time.sleep(max(args.delay_seconds, 0))
            try:
                page_hits, eng = scrape_faculty_pages(
                    session,
                    uni_name,
                    department,
                    official_www,
                    city,
                    args.engine,
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
            search_query = build_faculty_search_query(uni_name, city, department)
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
                city=city,
                division_type=div["division_type"],
                department=department,
                fields_of_study=div["fields_of_study"],
                official_www=official_www,
                search_query=meta["search_query"],
                search_engine=meta["search_engine"],
            )
            placeholder["error"] = err or "No faculty found"
            new_rows.append(placeholder)
            print(f"  No faculty ({placeholder['error']})")

        rows_out.extend(new_rows)
        done_depts.add(dept_key)
        save_progress(rows_out, output_path)

    return rows_out, done_depts, len(rows_out)


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


def build_faculty_search_query(university: str, city: str, department: str) -> str:
    parts: list[str] = ["site:linkedin.com/in"]
    if university:
        parts.append(f'"{university}"')
    if city:
        parts.append(city)
    if department:
        parts.append(f'"{department}"')
    parts.append("(professor OR lecturer OR faculty OR instructor)")
    return " ".join(parts)


def guess_faculty_urls(base: str, department: str) -> list[str]:
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
    return urls


def discover_faculty_urls(
    session: requests.Session,
    university: str,
    department: str,
    official_www: str,
    city: str,
    engine: str,
    max_search: int,
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

    for u in guess_faculty_urls(base, department):
        add(u)
    for u in preset.get("seed_urls", []):
        add(str(u))
        if "faculty" not in str(u).lower():
            add(str(u).rstrip("/") + "/Faculty")

    queries: list[str] = []
    if official_host:
        queries.append(f"site:{official_host} faculty staff directory")
        if department:
            queries.append(f'site:{official_host} "{department}" faculty professors')
            queries.append(f'site:{official_host} "{department}" lecturers staff')
        queries.append(f"site:{official_host} faculty email contact")
    else:
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
    max_pages: int,
    max_per_page: int,
    via_jina: bool,
    fetch_delay: float,
    crawl_links: bool,
) -> tuple[list[dict[str, str]], str]:
    candidates = discover_faculty_urls(
        session, university, department, official_www, city, engine, max_search=max(12, max_pages)
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
    parser.add_argument(
        "--fallback-input",
        type=Path,
        default=DEFAULT_FALLBACK_INPUT,
        help="Excel used when --university is not found in institutions.csv (default: data-raw/NUTECH.xlsx)",
    )
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument(
        "--per-university",
        action="store_true",
        help="Save each university to its own Excel file (under --output-dir); saves after each department",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help="Directory for per-university Excel files when using --per-university",
    )
    parser.add_argument("--engine", choices=("auto", "ddgs", "duckduckgo_html", "jina_ddg"), default="auto")
    parser.add_argument("--top-profiles", type=int, default=25, help="Max LinkedIn profiles per department")
    parser.add_argument("--delay-seconds", type=float, default=1.0)
    parser.add_argument("--max-institutions", type=int, default=0)
    parser.add_argument(
        "--skip-institutions",
        type=int,
        default=0,
        help="Skip the first N universities in institutions.csv (use with --max-institutions for the next batch)",
    )
    parser.add_argument(
        "--skip-existing",
        action="store_true",
        help="Per-university mode: skip universities that already have an output .xlsx file",
    )
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
    args = parser.parse_args()

    inp = resolve_data_path(args.input)
    fallback_inp = resolve_data_path(args.fallback_input) if args.fallback_input else None
    out_path = args.output.resolve()
    if not inp.is_file() and not (args.university and fallback_inp and fallback_inp.is_file()):
        print(f"Input not found: {inp}")
        return 1

    pakistan = resolve_pakistan_institutions(
        inp,
        fallback_inp,
        args.university,
        args.iau_id,
    )
    if pakistan.empty:
        print("No universities matched filters.")
        return 1

    if args.university or args.iau_id:
        print(f"Matched: {', '.join(str(r['name']) for _, r in pakistan.iterrows())}")

    session = requests.Session()
    session.headers.update(HEADERS)

    n_inst = 0
    total_rows = 0

    if args.per_university:
        output_dir = args.output_dir.resolve()
        output_dir.mkdir(parents=True, exist_ok=True)
        print(f"Per-university mode: one Excel file per institution in {output_dir}")
        if args.skip_institutions:
            print(f"Skipping first {args.skip_institutions} universities in CSV order")

        skipped = 0
        for _, row in pakistan.iterrows():
            uni_name = str(row.get("name", "")).strip()
            if not uni_name:
                continue
            if args.skip_institutions and skipped < args.skip_institutions:
                skipped += 1
                continue
            if args.max_institutions and n_inst >= args.max_institutions:
                break
            iau = str(row.get("iau_id", "")).strip()
            uni_path = university_output_path(output_dir, uni_name, iau)
            if args.skip_existing and uni_path.is_file() and not args.fresh:
                print(f"\nSkip existing: {uni_name} -> {uni_path.name}")
                continue
            n_inst += 1
            print(f"\n[{n_inst}] (CSV row after {skipped} skipped)")
            _, _, n_rows = scrape_university(session, row, args=args, output_path=uni_path)
            total_rows += n_rows

        print(f"\nDone. {n_inst} universities, {total_rows} total rows in {output_dir}")
        return 0

    # Single combined workbook (append across universities by default).
    done_depts: set[str] = set()
    rows_out: list[dict[str, str]] = []
    if not args.fresh and out_path.is_file():
        rows_out, done_depts = load_output_state(out_path, fresh=False)
        print(f"Appending to existing file: {len(rows_out)} rows, {len(done_depts)} dept keys on record")
    elif args.fresh:
        print("Fresh run: existing output will be replaced when the first batch saves")

    skipped = 0
    for _, row in pakistan.iterrows():
        uni_name = str(row.get("name", "")).strip()
        if not uni_name:
            continue
        if args.skip_institutions and skipped < args.skip_institutions:
            skipped += 1
            continue
        if args.max_institutions and n_inst >= args.max_institutions:
            break

        n_inst += 1
        print(f"\n[{n_inst}]")
        rows_out, done_depts, _ = scrape_university(
            session,
            row,
            args=args,
            output_path=out_path,
            rows_out=rows_out,
            done_depts=done_depts,
        )

    print(f"\nDone. {len(rows_out)} rows -> {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
