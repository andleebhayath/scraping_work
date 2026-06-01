"""
Riphah CS students via Google Custom Search API (no LinkedIn login).

Uses your own Google API key + Programmable Search Engine (CSE) to find LinkedIn
profiles and pull public text Google has indexed: headline, location, connections,
study years at Riphah, education, experience, about, email/phone when visible.

Limits (without logging into LinkedIn):
  - You only get what Google shows in search snippets (not the full profile page).
  - Study year and experience are often incomplete or missing.
  - Free Google CSE quota is ~100 searches/day; each profile uses ~12 searches.

Setup:
  1. Google Cloud: enable "Custom Search API", create an API key.
  2. https://programmablesearchengine.google.com/ — create a search engine that
     searches the entire web. Copy the "Search engine ID" (cx).
  3. In project-root .env (gitignored):
       GOOGLE_API_KEY=your_key_here
       GOOGLE_CSE_ID=your_cx_here

Run (discover 10 students + enrich each via Google):
  python scripts/riphah_cs_students_google_api.py --max-profiles 10

Enrich URLs already in a CSV/Excel (no new discovery):
  python scripts/riphah_cs_students_google_api.py --input output/riphah_cs_students_100.csv --refresh-existing

Output:
  output/riphah_cs_students_google.csv
  output/riphah_cs_students_google.xlsx
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests

_SCRIPTS = Path(__file__).resolve().parent
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

from riphah_cs_students_serpapi import (  # noqa: E402
    DEFAULT_CAMPUS,
    DEFAULT_DEPARTMENT,
    DEFAULT_UNIVERSITY,
    apply_extended_fields,
    apply_parsed_fields,
    build_profile_enrich_queries,
    build_search_queries,
    contacts_from_text,
    infer_headline_from_about,
    is_cs_related,
    is_likely_student,
    is_linkedin_profile,
    linkedin_username_from_url,
    load_env_file,
    name_from_linkedin_url,
    name_from_title,
    normalize_linkedin_profile_url,
    parse_headline,
    profile_slug_matches,
)

PROJECT_ROOT = Path(__file__).resolve().parents[1]
GOOGLE_CSE_URL = "https://www.googleapis.com/customsearch/v1"
DEFAULT_OUTPUT_CSV = PROJECT_ROOT / "output" / "riphah_cs_students_google.csv"
DEFAULT_OUTPUT_XLSX = PROJECT_ROOT / "output" / "riphah_cs_students_google.xlsx"

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


class GoogleCseError(RuntimeError):
    pass


def google_custom_search(
    session: requests.Session,
    api_key: str,
    cse_id: str,
    query: str,
    *,
    num: int = 10,
    start: int = 1,
    gl: str = "pk",
    hl: str = "en",
) -> dict:
    """Google Programmable Search Engine JSON API (Custom Search v1)."""
    params = {
        "key": api_key,
        "cx": cse_id,
        "q": query,
        "num": min(max(num, 1), 10),
        "start": max(start, 1),
        "gl": gl,
        "hl": hl,
    }
    resp = session.get(GOOGLE_CSE_URL, params=params, timeout=60)
    if resp.status_code == 429:
        raise GoogleCseError("Google API quota exceeded (429). Try tomorrow or raise quota.")
    resp.raise_for_status()
    data = resp.json()
    err = data.get("error")
    if err:
        msg = err.get("message") or str(err)
        raise GoogleCseError(msg)
    return data


def organic_results(data: dict) -> list[dict]:
    return list(data.get("items") or [])


def build_google_enrich_queries(username: str, name: str, university: str) -> list[str]:
    base = build_profile_enrich_queries(username, name, university)
    extra = [
        f'site:linkedin.com/in/{username} "Experience"',
        f'site:linkedin.com/in/{username} "Present" OR "2024" OR "2025"',
        f'site:linkedin.com/in/{username} intern OR developer OR engineer Pakistan',
        f'site:linkedin.com/in/{username} "Student at" OR "studying at"',
    ]
    return list(dict.fromkeys(base + extra))


def apply_google_fields(
    row: dict[str, str],
    text: str,
    *,
    university: str,
    overwrite: bool = False,
) -> None:
    apply_parsed_fields(row, text, university=university, overwrite=overwrite)
    if not row.get("headline"):
        row["headline"] = infer_headline_from_about(row.get("about", ""))
    apply_extended_fields(row, text, university=university, overwrite=overwrite)


def collect_profile_google_text(
    session: requests.Session,
    api_key: str,
    cse_id: str,
    username: str,
    queries: list[str],
    *,
    sleep_seconds: float,
) -> str:
    chunks: list[str] = []
    seen: set[str] = set()
    for query in queries:
        try:
            data = google_custom_search(session, api_key, cse_id, query, num=8)
        except GoogleCseError:
            raise
        except Exception:
            time.sleep(sleep_seconds)
            continue
        for item in organic_results(data):
            link = normalize_linkedin_profile_url(item.get("link") or "")
            if not profile_slug_matches(link, username):
                continue
            blob = " ".join(
                p
                for p in (
                    item.get("title") or "",
                    item.get("snippet") or "",
                    item.get("htmlSnippet") or "",
                )
                if p
            ).strip()
            if blob and blob not in seen:
                seen.add(blob)
                chunks.append(blob)
        time.sleep(sleep_seconds)
    return " ".join(chunks).strip()


def profiles_from_google_results(
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
    for item in organic_results(data):
        if len(rows) >= max_new:
            break
        raw_url = (item.get("link") or "").strip()
        profile_url = normalize_linkedin_profile_url(raw_url)
        if not profile_url or not is_linkedin_profile(profile_url):
            continue
        if profile_url.lower() in seen_urls:
            continue

        title = (item.get("title") or "").strip()
        snippet = (item.get("snippet") or "").strip()
        combined = f"{title} {snippet}"

        if students_only and not is_likely_student(combined):
            continue
        if cs_only and not is_cs_related(combined):
            continue

        seen_urls.add(profile_url.lower())
        row = {col: "" for col in OUTPUT_COLUMNS}
        row.update(
            {
                "university": university,
                "campus": campus,
                "department": department,
                "student_name": name_from_title(title) or name_from_linkedin_url(profile_url),
                "linkedin_url": profile_url,
                "result_title": title,
                "snippet": snippet,
                "search_query": query,
                "data_source": "google_cse",
                "scraped_at": datetime.now().isoformat(timespec="seconds"),
            }
        )
        headline = parse_headline(title, snippet)
        if headline:
            row["headline"] = headline
        email, phone = contacts_from_text(combined)
        if email:
            row["email"] = email
        if phone:
            row["phone"] = phone
        apply_google_fields(row, combined, university=university)
        rows.append(row)
    return rows


def enrich_profile_via_google(
    session: requests.Session,
    api_key: str,
    cse_id: str,
    row: dict[str, str],
    *,
    university: str,
    sleep_seconds: float,
) -> None:
    username = linkedin_username_from_url(row.get("linkedin_url", ""))
    if not username:
        return
    name = row.get("student_name", "").strip()
    queries = build_google_enrich_queries(username, name, university)
    profile_text = collect_profile_google_text(
        session,
        api_key,
        cse_id,
        username,
        queries,
        sleep_seconds=sleep_seconds,
    )
    merged = " ".join(
        p
        for p in (
            row.get("snippet", ""),
            row.get("result_title", ""),
            row.get("about", ""),
            profile_text,
        )
        if p
    ).strip()
    if not merged:
        return
    if profile_text:
        row["snippet"] = profile_text[:2500]
    apply_google_fields(row, merged, university=university, overwrite=True)
    row["data_source"] = "google_cse+profile_lookup"
    row["scraped_at"] = datetime.now().isoformat(timespec="seconds")


def profile_row_key(row: dict[str, str]) -> str:
    return normalize_linkedin_profile_url(row.get("linkedin_url", "")).lower()


def load_table(path: Path) -> list[dict[str, str]]:
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path, dtype=str)
    else:
        df = pd.read_excel(path, dtype=str)
    df = df.fillna("")
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return [dict(r) for r in df[OUTPUT_COLUMNS].to_dict("records")]


def load_existing_outputs(csv_path: Path, xlsx_path: Path) -> list[dict[str, str]]:
    if xlsx_path.is_file():
        return load_table(xlsx_path)
    if csv_path.is_file():
        return load_table(csv_path)
    return []


def save_outputs(rows: list[dict[str, str]], csv_path: Path, xlsx_path: Path) -> None:
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    for i, row in enumerate(rows, start=1):
        row["profile_rank"] = str(i)
    df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    df.to_excel(xlsx_path, index=False, sheet_name="students")


def merge_row(existing: dict[str, str], new: dict[str, str]) -> dict[str, str]:
    merged = dict(existing)
    for key in OUTPUT_COLUMNS:
        old_v = str(merged.get(key, "")).strip()
        new_v = str(new.get(key, "")).strip()
        if new_v and (not old_v or len(new_v) > len(old_v)):
            merged[key] = new_v
    merged["scraped_at"] = new.get("scraped_at") or merged.get("scraped_at", "")
    if new.get("data_source"):
        merged["data_source"] = new["data_source"]
    return merged


def upsert_row(stored: list[dict[str, str]], row: dict[str, str]) -> bool:
    key = profile_row_key(row)
    if not key:
        return False
    for i, r in enumerate(stored):
        if profile_row_key(r) == key:
            stored[i] = merge_row(r, row)
            return False
    stored.append(row)
    return True


def run_discovery(
    session: requests.Session,
    api_key: str,
    cse_id: str,
    *,
    university: str,
    campus: str,
    department: str,
    max_profiles: int,
    seen_urls: set[str],
    stored: list[dict[str, str]],
    students_only: bool,
    cs_only: bool,
    sleep_seconds: float,
    skip_enrich: bool,
    csv_path: Path,
    xlsx_path: Path,
) -> int:
    queries = build_search_queries(university, department)
    added = 0
    for qi, query in enumerate(queries, start=1):
        if len(stored) >= max_profiles:
            break
        print(f"  Query {qi}/{len(queries)}: {query[:72]}")
        try:
            data = google_custom_search(session, api_key, cse_id, query, num=10)
        except GoogleCseError as exc:
            print(f"    stopped: {exc}")
            break
        need = max_profiles - len(stored)
        batch = profiles_from_google_results(
            data,
            query=query,
            university=university,
            campus=campus,
            department=department,
            seen_urls=seen_urls,
            max_new=need,
            students_only=students_only,
            cs_only=cs_only,
        )
        for row in batch:
            if not skip_enrich:
                try:
                    enrich_profile_via_google(
                        session,
                        api_key,
                        cse_id,
                        row,
                        university=university,
                        sleep_seconds=sleep_seconds,
                    )
                except GoogleCseError as exc:
                    print(f"    enrich stopped: {exc}")
                    save_outputs(stored, csv_path, xlsx_path)
                    return added
            if upsert_row(stored, row):
                added += 1
            save_outputs(stored, csv_path, xlsx_path)
            print(
                f"    + {row.get('student_name', '')[:40]} | "
                f"{row.get('location', '')} | {row.get('study_year', '')} | "
                f"exp={bool(row.get('experience_summary'))}",
                flush=True,
            )
            if len(stored) >= max_profiles:
                break
        time.sleep(sleep_seconds)
    return added


def enrich_stored_rows(
    session: requests.Session,
    api_key: str,
    cse_id: str,
    stored: list[dict[str, str]],
    *,
    university: str,
    sleep_seconds: float,
    refresh: bool,
    limit: int,
    csv_path: Path,
    xlsx_path: Path,
) -> int:
    updated = 0
    for i, row in enumerate(stored):
        if limit > 0 and updated >= limit:
            break
        if not refresh and row.get("data_source", "").endswith("profile_lookup"):
            if row.get("location") and row.get("experience_summary"):
                continue
        url = normalize_linkedin_profile_url(row.get("linkedin_url", ""))
        if not url:
            continue
        print(f"[{i + 1}/{len(stored)}] {row.get('student_name', '')} -> {url}")
        try:
            enrich_profile_via_google(
                session,
                api_key,
                cse_id,
                row,
                university=university,
                sleep_seconds=sleep_seconds,
            )
        except GoogleCseError as exc:
            print(f"  stopped: {exc}")
            save_outputs(stored, csv_path, xlsx_path)
            break
        updated += 1
        save_outputs(stored, csv_path, xlsx_path)
        print(
            f"  -> {row.get('headline', '')[:50]} | {row.get('study_year', '')} | "
            f"{row.get('experience_roles', '')[:60]}",
            flush=True,
        )
    return updated


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Riphah CS students via Google Custom Search API (no LinkedIn login)."
    )
    p.add_argument("--university", default=DEFAULT_UNIVERSITY)
    p.add_argument("--campus", default=DEFAULT_CAMPUS)
    p.add_argument("--department", default=DEFAULT_DEPARTMENT)
    p.add_argument(
        "--max-profiles",
        type=int,
        default=10,
        help="New profiles to add (discovery mode) or max rows to enrich",
    )
    p.add_argument(
        "--api-key",
        default="",
        help="Google API key (default: GOOGLE_API_KEY in .env)",
    )
    p.add_argument(
        "--cse-id",
        default="",
        help="Programmable Search Engine ID / cx (default: GOOGLE_CSE_ID in .env)",
    )
    p.add_argument("--output-csv", type=Path, default=DEFAULT_OUTPUT_CSV)
    p.add_argument("--output-xlsx", type=Path, default=DEFAULT_OUTPUT_XLSX)
    p.add_argument(
        "--input",
        type=Path,
        default=None,
        help="CSV/XLSX with linkedin_url column — enrich only, no discovery",
    )
    p.add_argument(
        "--refresh-existing",
        action="store_true",
        help="Re-run Google lookups for rows in --input or output file",
    )
    p.add_argument(
        "--skip-enrich",
        action="store_true",
        help="Discovery only (faster; less data per row)",
    )
    p.add_argument("--fresh", action="store_true", help="Ignore existing output files")
    p.add_argument("--include-faculty", action="store_true")
    p.add_argument("--sleep", type=float, default=1.2)
    return p.parse_args()


def import_urls_from_input(path: Path, university: str, campus: str, department: str) -> list[dict[str, str]]:
    df = pd.read_csv(path, dtype=str) if path.suffix.lower() == ".csv" else pd.read_excel(path, dtype=str)
    df = df.fillna("")
    rows: list[dict[str, str]] = []
    for _, rec in df.iterrows():
        url = normalize_linkedin_profile_url(str(rec.get("linkedin_url", "")))
        if not url:
            continue
        row = {col: "" for col in OUTPUT_COLUMNS}
        row.update(
            {
                "university": university,
                "campus": campus,
                "department": department,
                "student_name": str(rec.get("student_name", "")).strip()
                or name_from_linkedin_url(url),
                "linkedin_url": url,
                "data_source": "imported",
                "scraped_at": datetime.now().isoformat(timespec="seconds"),
            }
        )
        for col in OUTPUT_COLUMNS:
            if col in rec and str(rec[col]).strip():
                row[col] = str(rec[col]).strip()
        rows.append(row)
    return rows


def main() -> int:
    args = parse_args()
    load_env_file(PROJECT_ROOT / ".env")

    api_key = (args.api_key or os.environ.get("GOOGLE_API_KEY", "")).strip()
    cse_id = (args.cse_id or os.environ.get("GOOGLE_CSE_ID", "")).strip()
    if not api_key or not cse_id:
        print(
            "Missing Google credentials.\n"
            "  Set GOOGLE_API_KEY and GOOGLE_CSE_ID in .env, or pass --api-key and --cse-id.\n"
            "  Create CSE: https://programmablesearchengine.google.com/\n"
            "  Enable API: Google Cloud Console → Custom Search API",
            file=sys.stderr,
        )
        return 1

    csv_path = args.output_csv
    xlsx_path = args.output_xlsx
    if args.fresh:
        stored: list[dict[str, str]] = []
    else:
        stored = load_existing_outputs(csv_path, xlsx_path)

    seen_urls = {profile_row_key(r) for r in stored if profile_row_key(r)}
    session = requests.Session()
    students_only = not args.include_faculty
    cs_only = True

    if args.input:
        path = args.input
        if not path.is_file():
            alt = path.with_suffix(".csv" if path.suffix.lower() == ".xlsx" else ".xlsx")
            path = alt if alt.is_file() else path
        if not path.is_file():
            print(f"Input not found: {args.input}", file=sys.stderr)
            return 1
        imported = import_urls_from_input(
            path, args.university, args.campus, args.department
        )
        for row in imported:
            upsert_row(stored, row)
        print(f"Loaded {len(imported)} URL(s) from {path}")
        n = enrich_stored_rows(
            session,
            api_key,
            cse_id,
            stored,
            university=args.university,
            sleep_seconds=args.sleep,
            refresh=args.refresh_existing,
            limit=args.max_profiles if args.max_profiles > 0 else 0,
            csv_path=csv_path,
            xlsx_path=xlsx_path,
        )
        save_outputs(stored, csv_path, xlsx_path)
        print(f"\nDone: enriched {n} row(s) -> {xlsx_path}")
        return 0

    target = max(args.max_profiles, 0)
    if target <= len(stored) and not args.refresh_existing:
        print(f"Already have {len(stored)} profile(s). Use --refresh-existing or --fresh.")
        return 0

    print(
        "Note: Without LinkedIn login, Google snippets are incomplete. "
        "Use riphah_cs_students_linkedin_playwright.py for full profiles.\n"
    )

    if len(stored) < target:
        added = run_discovery(
            session,
            api_key,
            cse_id,
            university=args.university,
            campus=args.campus,
            department=args.department,
            max_profiles=target,
            seen_urls=seen_urls,
            stored=stored,
            students_only=students_only,
            cs_only=cs_only,
            sleep_seconds=args.sleep,
            skip_enrich=args.skip_enrich,
            csv_path=csv_path,
            xlsx_path=xlsx_path,
        )
        print(f"Discovery added {added} new profile(s).")

    if args.refresh_existing:
        n = enrich_stored_rows(
            session,
            api_key,
            cse_id,
            stored,
            university=args.university,
            sleep_seconds=args.sleep,
            refresh=True,
            limit=0,
            csv_path=csv_path,
            xlsx_path=xlsx_path,
        )
        print(f"Refresh enriched {n} row(s).")

    save_outputs(stored, csv_path, xlsx_path)
    print(f"\nSaved {len(stored)} profile(s) -> {xlsx_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
