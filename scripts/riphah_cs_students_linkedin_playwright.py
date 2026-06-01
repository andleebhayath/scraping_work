"""
Fill accurate LinkedIn fields (study year at Riphah, location, connections) using a
logged-in browser session — the only reliable way to match what you see on LinkedIn.

Discovers all Riphah International University students (any program), not only BSCS/CS.

Google/SerpAPI snippets are often incomplete or outdated.

Setup (one time, PowerShell — run each line separately):
  pip install playwright
  python -m playwright install chromium
  (Package name is playwright, not playwrite. Do not use && in PowerShell.)

  python scripts/riphah_cs_students_linkedin_playwright.py --login
  # Log in to LinkedIn in the opened browser, then press Enter in the terminal.

Run 100 students (discover + scrape to Excel):
  python scripts/riphah_cs_students_linkedin_playwright.py --discover 100 \\
    --output output/riphah_cs_students_100.xlsx

Append 100 more (skips URLs already in the output file):
  python scripts/riphah_cs_students_linkedin_playwright.py \\
    --input output/riphah_cs_students_100.xlsx --discover 100 --max-profiles 100 \\
    --output output/riphah_cs_students_100.xlsx

Re-enrich first N rows only (does not remove other rows):
  python scripts/riphah_cs_students_linkedin_playwright.py \\
    --input output/riphah_cs_students_100.xlsx --max-profiles 10 \\
    --output output/riphah_cs_students_100.xlsx

Re-fill headline, study year, email, phone for ALL rows already in the file:
  python scripts/riphah_cs_students_linkedin_playwright.py \\
    --input output/riphah_cs_students_100.csv --refresh-existing \\
    --output output/riphah_cs_students_100.xlsx

Fill rows that only have a URL (data_source = ddg_discover) — run after append/discover:
  python scripts/riphah_cs_students_linkedin_playwright.py \\
    --input output/riphah_cs_students_100.csv \\
    --output output/riphah_cs_students_100.xlsx

"""

from __future__ import annotations

import argparse
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import quote_plus

import pandas as pd

_SCRIPTS = Path(__file__).resolve().parent
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

import uni_data_free_google as udg  # noqa: E402

from riphah_cs_students_serpapi import (  # noqa: E402
    DEFAULT_CAMPUS,
    DEFAULT_UNIVERSITY,
    OUTPUT_COLUMNS,
    apply_parsed_fields,
    infer_headline_from_about,
    infer_location,
    name_from_linkedin_url,
    normalize_linkedin_profile_url,
)

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_AUTH = PROJECT_ROOT / "data-raw" / "linkedin_auth.json"
DEFAULT_INPUT = PROJECT_ROOT / "output" / "riphah_cs_students_serpapi.xlsx"
DEFAULT_OUTPUT_100 = PROJECT_ROOT / "output" / "riphah_cs_students_100.xlsx"
DEFAULT_DEPARTMENT_ALL = "All departments"


def format_duration(seconds: float) -> str:
    """Human-readable duration for terminal logs."""
    if seconds < 1:
        return f"{seconds * 1000:.0f}ms"
    if seconds < 60:
        return f"{seconds:.1f}s"
    minutes, secs = divmod(int(seconds), 60)
    if minutes < 60:
        return f"{minutes}m {secs}s"
    hours, minutes = divmod(minutes, 60)
    return f"{hours}h {minutes}m {secs}s"


LINKEDIN_DISCOVERY_QUERIES = (
    "Riphah International University student",
    "Riphah International University undergraduate student",
    "student at Riphah International University",
    "Riphah student Pakistan",
    "Riphah student Lahore",
    "Riphah student Islamabad",
    "Riphah student Rawalpindi",
    "Riphah International University alumni student",
    "studying at Riphah International University",
    "Riphah campus student Pakistan",
)

DDG_DISCOVERY_QUERIES = (
    'site:linkedin.com/in "Riphah International University" student',
    'site:linkedin.com/in "student at Riphah International University"',
    'site:linkedin.com/in Riphah student Pakistan',
    'site:linkedin.com/in Riphah undergraduate student',
    'site:linkedin.com/in Riphah student Lahore',
    'site:linkedin.com/in Riphah student Islamabad',
    'site:linkedin.com/in Riphah student Rawalpindi',
)

# Extra queries when appending (broader campuses / programs)
DDG_APPEND_QUERIES = (
    'site:linkedin.com/in Riphah student Faisalabad',
    'site:linkedin.com/in Riphah student Karachi',
    'site:linkedin.com/in Riphah student Multan',
    'site:linkedin.com/in "studying at Riphah"',
    'site:linkedin.com/in Riphah Gulberg student',
    'site:linkedin.com/in Riphah Al Mizan student',
    'site:linkedin.com/in Riphah campus student',
    'site:linkedin.com/in Riphah International University intern',
    'site:linkedin.com/in Riphah "final year" student',
    'site:linkedin.com/in Riphah Raiwind student',
)

DDG_APPEND_QUERIES_2 = (
    'site:linkedin.com/in Riphah Nishtar student',
    'site:linkedin.com/in Riphah Wah Cantt student',
    'site:linkedin.com/in Riphah student Peshawar',
    'site:linkedin.com/in Riphah student Abbottabad',
    'site:linkedin.com/in "Riphah International University" undergraduate',
    'site:linkedin.com/in Riphah medical student',
    'site:linkedin.com/in Riphah business student',
    'site:linkedin.com/in Riphah engineering student',
    'site:linkedin.com/in Riphah law student',
    'site:linkedin.com/in Riphah pharmacy student',
    'site:linkedin.com/in Riphah MBA student',
    'site:linkedin.com/in Riphah student 2024',
    'site:linkedin.com/in Riphah student 2025',
)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Enrich student rows from logged-in LinkedIn.")
    p.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    p.add_argument("--auth", type=Path, default=DEFAULT_AUTH)
    p.add_argument("--login", action="store_true", help="Open browser to save LinkedIn session")
    p.add_argument(
        "--install-browser",
        action="store_true",
        help="Download Chromium for Playwright (~180 MB), then exit",
    )
    p.add_argument("--max-profiles", type=int, default=0, help="0 = all rows")
    p.add_argument(
        "--start-index",
        type=int,
        default=0,
        help="Resume enrichment at this row index (0-based)",
    )
    p.add_argument(
        "--discover",
        type=int,
        default=0,
        help="Find N new LinkedIn profiles (LinkedIn search + DuckDuckGo), then scrape",
    )
    p.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output Excel/CSV path (default: riphah_cs_students_100.xlsx when --discover)",
    )
    p.add_argument("--sleep", type=float, default=2.0, help="Delay between profiles / scrolls")
    p.add_argument("--university", default=DEFAULT_UNIVERSITY)
    p.add_argument("--campus", default=DEFAULT_CAMPUS)
    p.add_argument(
        "--department",
        default=DEFAULT_DEPARTMENT_ALL,
        help="Label stored in output (default: All departments — no CS-only filter)",
    )
    p.add_argument(
        "--skip-linkedin-search",
        action="store_true",
        help="Only use DuckDuckGo for discovery (no LinkedIn people search)",
    )
    p.add_argument(
        "--refresh-existing",
        action="store_true",
        help="Re-visit each profile and overwrite empty or stale fields",
    )
    p.add_argument(
        "--backfill-headlines",
        action="store_true",
        help="Fill empty headline from about text (no browser); then exit",
    )
    return p.parse_args()


def install_chromium() -> None:
    print("Downloading Chromium for Playwright (~180 MB). This may take a few minutes...")
    subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"])
    print("Chromium installed.")


def ensure_chromium_installed() -> None:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise SystemExit(
            "Playwright is not installed.\n"
            "  pip install playwright\n"
            "  python -m playwright install chromium\n"
            "(PowerShell: run those two lines separately, not with &&)"
        ) from exc

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            browser.close()
        return
    except Exception as exc:
        if "Executable doesn't exist" not in str(exc):
            raise
    install_chromium()
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        browser.close()


def save_session(auth_path: Path) -> None:
    ensure_chromium_installed()

    from playwright.sync_api import sync_playwright

    auth_path.parent.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        page.goto("https://www.linkedin.com/login", wait_until="domcontentloaded")
        print("Log in to LinkedIn in the browser window, then press Enter here...")
        input()
        context.storage_state(path=str(auth_path))
        browser.close()
    print(f"Saved session: {auth_path}")


def page_visible_text(page) -> str:
    try:
        return page.inner_text("body")
    except Exception:
        return ""


def resolve_input_file(path: Path) -> Path | None:
    """Accept .xlsx or .csv; try sibling file if one is missing."""
    if path.is_file():
        return path
    alt = path.with_suffix(".csv" if path.suffix.lower() == ".xlsx" else ".xlsx")
    if alt.is_file():
        return alt
    return None


def load_profiles_table(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path, dtype=str)
    else:
        df = pd.read_excel(path, dtype=str)
    return df.fillna("")


def extract_linkedin_topcard_text(page) -> str:
    """Profile header area — best source for city (not full page sidebar noise)."""
    selectors = (
        "section.pv-top-card",
        "div.ph5.pb5",
        "main.scaffold-layout__main",
        "main",
    )
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            if loc.count():
                text = loc.inner_text(timeout=5000)
                if text and len(text.strip()) > 20:
                    return text.strip()
        except Exception:
            continue
    try:
        return page.inner_text("body")[:2500]
    except Exception:
        return ""


def scrape_profile(page, url: str, *, university: str) -> dict[str, str]:
    page.goto(url, wait_until="domcontentloaded", timeout=60_000)
    time.sleep(2.0)
    for _ in range(4):
        try:
            page.evaluate("window.scrollBy(0, 1400)")
        except Exception:
            pass
        time.sleep(0.7)
    topcard = extract_linkedin_topcard_text(page)
    full_text = page_visible_text(page)
    text = f"{topcard}\n{full_text[:16000]}"
    try:
        page_title = page.title()
    except Exception:
        page_title = ""
    row: dict[str, str] = {
        "linkedin_url": normalize_linkedin_profile_url(url),
        "snippet": text[:8000],
        "result_title": page_title,
        "headline": "",
        "about": "",
        "university": university,
    }
    apply_parsed_fields(row, text, university=university, overwrite=True, topcard=topcard)
    if not row.get("headline") and row.get("about"):
        row["headline"] = infer_headline_from_about(row["about"])
    header_loc = infer_location(topcard)
    if header_loc:
        row["location"] = header_loc
    row["scraped_at"] = datetime.now().isoformat(timespec="seconds")
    row["data_source"] = "linkedin_playwright"
    return row


def new_student_row(
    *,
    linkedin_url: str,
    student_name: str,
    university: str,
    campus: str,
    department: str,
    data_source: str,
) -> dict[str, str]:
    row = {col: "" for col in OUTPUT_COLUMNS}
    row.update(
        {
            "university": university,
            "campus": campus,
            "department": department,
            "student_name": student_name,
            "linkedin_url": linkedin_url,
            "data_source": data_source,
            "scraped_at": datetime.now().isoformat(timespec="seconds"),
        }
    )
    return row


def profile_urls_from_page(page) -> list[tuple[str, str]]:
    """Return (url, display_name) from visible /in/ links."""
    found: list[tuple[str, str]] = []
    seen: set[str] = set()
    try:
        hrefs = page.eval_on_selector_all(
            'a[href*="/in/"]',
            """els => els.map(e => ({
                href: e.href || '',
                text: (e.innerText || '').trim().split('\\n')[0]
            }))""",
        )
    except Exception:
        return found
    for item in hrefs or []:
        profile_url = normalize_linkedin_profile_url(str(item.get("href", "")).split("?")[0])
        if not profile_url or not udg.is_profile_like_linkedin(profile_url):
            continue
        key = profile_url.lower()
        if key in seen:
            continue
        seen.add(key)
        name = str(item.get("text", "")).strip()
        if not name or len(name) > 80 or "linkedin" in name.lower():
            name = name_from_linkedin_url(profile_url)
        found.append((profile_url, name))
    return found


def collect_existing_urls(df: pd.DataFrame) -> set[str]:
    urls: set[str] = set()
    if "linkedin_url" not in df.columns:
        return urls
    for raw in df["linkedin_url"].astype(str):
        norm = normalize_linkedin_profile_url(raw)
        if norm:
            urls.add(norm.lower())
    return urls


def discover_via_linkedin_search(
    page,
    target: int,
    *,
    university: str,
    campus: str,
    department: str,
    sleep_seconds: float,
    exclude_urls: set[str] | None = None,
) -> list[dict[str, str]]:
    seen: set[str] = set(exclude_urls or ())
    rows: list[dict[str, str]] = []
    print(f"\nDiscovering up to {target} profiles via LinkedIn people search...", flush=True)
    for qi, query in enumerate(LINKEDIN_DISCOVERY_QUERIES, start=1):
        if len(rows) >= target:
            break
        search_url = (
            "https://www.linkedin.com/search/results/people/?"
            f"keywords={quote_plus(query)}&origin=GLOBAL_SEARCH_HEADER"
        )
        print(f"  Search {qi}: {query}", flush=True)
        q_started = time.perf_counter()
        try:
            page.goto(search_url, wait_until="domcontentloaded", timeout=60_000)
        except Exception as exc:
            print(f"    skip: {exc}")
            continue
        time.sleep(sleep_seconds + 1)
        for scroll in range(12):
            if len(rows) >= target:
                break
            for profile_url, name in profile_urls_from_page(page):
                key = profile_url.lower()
                if key in seen:
                    continue
                seen.add(key)
                rows.append(
                    new_student_row(
                        linkedin_url=profile_url,
                        student_name=name,
                        university=university,
                        campus=campus,
                        department=department,
                        data_source="linkedin_search_discover",
                    )
                )
                if len(rows) >= target:
                    break
            try:
                page.mouse.wheel(0, 2200)
            except Exception:
                page.evaluate("window.scrollBy(0, 2200)")
            time.sleep(sleep_seconds)
        print(f"    total so far: {len(rows)} ({format_duration(time.perf_counter() - q_started)})")
    return rows[:target]


def discover_via_ddg(
    target: int,
    *,
    university: str,
    campus: str,
    department: str,
    sleep_seconds: float,
    exclude_urls: set[str] | None = None,
    extra_queries: tuple[str, ...] = (),
) -> list[dict[str, str]]:
    import requests

    seen: set[str] = set(exclude_urls or ())
    rows: list[dict[str, str]] = []
    session = requests.Session()
    queries = DDG_DISCOVERY_QUERIES + extra_queries
    print(f"\nDiscovering up to {target} profiles via DuckDuckGo...")
    for qi, query in enumerate(queries, start=1):
        if len(rows) >= target:
            break
        print(f"  DDG {qi}: {query[:70]}")
        q_started = time.perf_counter()
        try:
            urls, _engine = udg.search_urls(session, query, max_results=100, engine="auto")
        except Exception as exc:
            print(f"    skip: {exc}")
            time.sleep(sleep_seconds)
            continue
        for raw_url in urls:
            profile_url = normalize_linkedin_profile_url(raw_url)
            if not profile_url or not udg.is_profile_like_linkedin(profile_url):
                continue
            key = profile_url.lower()
            if key in seen:
                continue
            seen.add(key)
            if "riphah" not in f"{query} {profile_url}".lower():
                continue
            rows.append(
                new_student_row(
                    linkedin_url=profile_url,
                    student_name=name_from_linkedin_url(profile_url),
                    university=university,
                    campus=campus,
                    department=department,
                    data_source="ddg_discover",
                )
            )
            if len(rows) >= target:
                break
        print(f"    total so far: {len(rows)} ({format_duration(time.perf_counter() - q_started)})")
        time.sleep(sleep_seconds)
    return rows[:target]


def discover_profiles(
    page,
    target: int,
    *,
    university: str,
    campus: str,
    department: str,
    sleep_seconds: float,
    skip_linkedin_search: bool,
    exclude_urls: set[str] | None = None,
    ddg_extra_queries: tuple[str, ...] = (),
) -> list[dict[str, str]]:
    seen: set[str] = set(exclude_urls or ())
    combined: list[dict[str, str]] = []

    def merge(batch: list[dict[str, str]]) -> None:
        for row in batch:
            key = row["linkedin_url"].lower()
            if key in seen:
                continue
            seen.add(key)
            combined.append(row)
            if len(combined) >= target:
                break

    if not skip_linkedin_search:
        merge(
            discover_via_linkedin_search(
                page,
                target,
                university=university,
                campus=campus,
                department=department,
                sleep_seconds=sleep_seconds,
                exclude_urls=seen.copy(),
            )
        )
    if len(combined) < target:
        need = target - len(combined)
        merge(
            discover_via_ddg(
                need,
                university=university,
                campus=campus,
                department=department,
                sleep_seconds=sleep_seconds,
                exclude_urls=seen.copy(),
                extra_queries=ddg_extra_queries,
            )
        )
    return combined[:target]


def dataframe_from_rows(rows: list[dict[str, str]]) -> pd.DataFrame:
    df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    return df.fillna("")


def row_needs_enrich(df: pd.DataFrame, i: int) -> bool:
    """True when the row only has a URL (DDG) or is missing LinkedIn scrape fields."""
    source = str(df.at[i, "data_source"]).strip().lower()
    if source in ("ddg_discover", "ddg"):
        return True
    if source != "linkedin_playwright":
        return not str(df.at[i, "location"]).strip() or not str(df.at[i, "about"]).strip()
    return not str(df.at[i, "location"]).strip() or not str(df.at[i, "about"]).strip()


def enrich_indices(
    df: pd.DataFrame, *, start_index: int = 0, refresh: bool = False
) -> list[int]:
    indices: list[int] = []
    for i in range(start_index, len(df)):
        if refresh or row_needs_enrich(df, i):
            indices.append(i)
    return indices


def backfill_headlines(df: pd.DataFrame) -> int:
    filled = 0
    for i in range(len(df)):
        if str(df.at[i, "headline"]).strip():
            continue
        about = str(df.at[i, "about"]).strip()
        if not about:
            continue
        headline = infer_headline_from_about(about)
        if headline:
            df.at[i, "headline"] = headline
            filled += 1
    return filled


def save_table(df: pd.DataFrame, out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    csv_path = out_path if out_path.suffix.lower() == ".csv" else out_path.with_suffix(".csv")
    xlsx_path = out_path if out_path.suffix.lower() == ".xlsx" else out_path.with_suffix(".xlsx")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    try:
        df.to_excel(xlsx_path, index=False, sheet_name="students")
    except PermissionError:
        print(
            f"  Warning: could not write {xlsx_path} (close it in Excel). CSV saved.",
            flush=True,
        )


def enrich_dataframe(
    df: pd.DataFrame,
    *,
    page,
    university: str,
    out_path: Path,
    limit: int,
    sleep_seconds: float,
    start_index: int = 0,
    refresh: bool = False,
) -> int:
    updated = 0
    indices = enrich_indices(df, start_index=start_index, refresh=refresh)
    if limit > 0:
        indices = indices[:limit]
    total = len(indices)
    if total == 0:
        print("No rows need LinkedIn enrichment.", flush=True)
        return 0
    print(f"Enriching {total} profile(s) on LinkedIn...", flush=True)
    batch_started = time.perf_counter()
    for n, i in enumerate(indices):
        iter_started = time.perf_counter()
        raw_url = str(df.at[i, "linkedin_url"])
        url = normalize_linkedin_profile_url(raw_url)
        if not url:
            print(f"  skip: invalid LinkedIn URL: {raw_url!r}", flush=True)
            continue
        df.at[i, "linkedin_url"] = url
        name = df.at[i, "student_name"]
        print(f"[{n + 1}/{total}] (row {i + 1}) {name} -> {url}", flush=True)
        try:
            scrape_started = time.perf_counter()
            scraped = scrape_profile(page, url, university=university)
            scrape_elapsed = time.perf_counter() - scrape_started
        except KeyboardInterrupt:
            print(
                f"\nStopped by user after {format_duration(time.perf_counter() - batch_started)} — saving...",
                flush=True,
            )
            save_table(df, out_path)
            raise
        except Exception as exc:
            print(
                f"  failed ({format_duration(time.perf_counter() - iter_started)}): {exc}",
                flush=True,
            )
            time.sleep(sleep_seconds)
            continue
        for key in (
            "location",
            "study_year",
            "connections",
            "degree_program",
            "subjects",
            "about",
            "headline",
            "email",
            "phone",
            "data_source",
            "student_name",
            "scraped_at",
        ):
            val = scraped.get(key, "")
            if not val:
                continue
            if refresh or not str(df.at[i, key]).strip():
                df.at[i, key] = val
        if not str(df.at[i, "student_name"]).strip() and scraped.get("student_name"):
            df.at[i, "student_name"] = scraped["student_name"]
        updated += 1
        save_started = time.perf_counter()
        save_table(df, out_path)
        save_elapsed = time.perf_counter() - save_started
        iter_elapsed = time.perf_counter() - iter_started
        elapsed_so_far = time.perf_counter() - batch_started
        avg = elapsed_so_far / (n + 1)
        eta = avg * (total - n - 1)
        print(
            f"  -> {scraped.get('location')} | {scraped.get('study_year')} | "
            f"{scraped.get('connections')} | "
            f"scrape {format_duration(scrape_elapsed)} | save {format_duration(save_elapsed)} | "
            f"iter {format_duration(iter_elapsed)} | ETA ~{format_duration(eta)}",
            flush=True,
        )
        time.sleep(sleep_seconds)
    batch_elapsed = time.perf_counter() - batch_started
    if updated:
        print(
            f"Enrichment pass: {updated}/{total} in {format_duration(batch_elapsed)} "
            f"(avg {format_duration(batch_elapsed / updated)} per profile)",
            flush=True,
        )
    return updated


def main() -> int:
    args = parse_args()
    if args.install_browser:
        install_chromium()
        return 0
    if args.login:
        save_session(args.auth)
        return 0

    if args.backfill_headlines:
        input_path = resolve_input_file(args.input) or resolve_input_file(DEFAULT_OUTPUT_100)
        if not input_path:
            print("No input file found for --backfill-headlines", file=sys.stderr)
            return 1
        df = load_profiles_table(input_path)
        for col in OUTPUT_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        out_path = args.output or input_path
        n = backfill_headlines(df)
        save_table(df, out_path)
        print(f"Backfilled {n} headline(s) -> {out_path}")
        return 0

    if not args.auth.is_file():
        print(
            f"Missing auth file: {args.auth}\n"
            "Run login first:\n"
            "  python scripts/riphah_cs_students_linkedin_playwright.py --login",
            file=sys.stderr,
        )
        return 1

    discover_n = max(0, args.discover)
    out_path = args.output or (DEFAULT_OUTPUT_100 if discover_n else DEFAULT_INPUT)
    limit = args.max_profiles if args.max_profiles > 0 else (discover_n or 0)

    ensure_chromium_installed()
    from playwright.sync_api import sync_playwright

    updated = 0
    job_started = time.perf_counter()
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            context = browser.new_context(storage_state=str(args.auth))
            page = context.new_page()

            if discover_n > 0:
                existing_df: pd.DataFrame | None = None
                exclude_urls: set[str] = set()
                input_path = resolve_input_file(args.input)
                if not input_path and out_path.is_file():
                    input_path = out_path
                if input_path:
                    existing_df = load_profiles_table(input_path)
                    for col in OUTPUT_COLUMNS:
                        if col not in existing_df.columns:
                            existing_df[col] = ""
                    exclude_urls = collect_existing_urls(existing_df)
                    print(
                        f"Loaded {len(existing_df)} existing row(s), "
                        f"{len(exclude_urls)} URL(s) to skip",
                        flush=True,
                    )

                rows = discover_profiles(
                    page,
                    discover_n,
                    university=args.university,
                    campus=args.campus,
                    department=args.department,
                    sleep_seconds=args.sleep,
                    skip_linkedin_search=args.skip_linkedin_search,
                    exclude_urls=exclude_urls,
                    ddg_extra_queries=(
                        (DDG_APPEND_QUERIES + DDG_APPEND_QUERIES_2) if exclude_urls else ()
                    ),
                )
                if not rows:
                    print("No profiles discovered.", file=sys.stderr)
                    browser.close()
                    return 2
                new_df = dataframe_from_rows(rows)
                base_rank = len(existing_df) if existing_df is not None else 0
                for i, row in enumerate(new_df.to_dict("records"), start=1):
                    row["profile_rank"] = str(base_rank + i)
                new_df = pd.DataFrame(new_df.to_dict("records"), columns=OUTPUT_COLUMNS).fillna("")

                if existing_df is not None:
                    existing_trim = existing_df.reindex(columns=OUTPUT_COLUMNS, fill_value="")
                    df = pd.concat([existing_trim, new_df], ignore_index=True)
                else:
                    df = new_df
                save_table(df, out_path)
                print(
                    f"\nDiscovered {len(new_df)} new profile(s), "
                    f"total {len(df)} -> {out_path}",
                    flush=True,
                )
                scrape_limit = limit if limit > 0 else 0
                enrich_start = args.start_index if args.start_index > 0 else len(df) - len(new_df)
                pending_tail = enrich_indices(
                    df, start_index=enrich_start, refresh=args.refresh_existing
                )
                pending_all = enrich_indices(df, start_index=0, refresh=False)
                print(
                    f"{len(pending_tail)} row(s) to scrape on LinkedIn "
                    f"(from spreadsheet row {enrich_start + 1})",
                    flush=True,
                )
                if len(pending_all) > len(pending_tail):
                    print(
                        f"Note: {len(pending_all) - len(pending_tail)} older row(s) still "
                        "only have URLs. Re-run without --discover to fill them.",
                        flush=True,
                    )
            else:
                input_path = resolve_input_file(args.input)
                if not input_path:
                    print(
                        f"Input not found: {args.input}\n"
                        "Use --discover 100 to find profiles, or create a CSV/XLSX first.",
                        file=sys.stderr,
                    )
                    browser.close()
                    return 1
                df = load_profiles_table(input_path)
                for col in OUTPUT_COLUMNS:
                    if col not in df.columns:
                        df[col] = ""
                out_path = args.output or input_path
                scrape_limit = limit if limit > 0 else 0
                enrich_start = args.start_index
                pending = enrich_indices(
                    df, start_index=enrich_start, refresh=args.refresh_existing
                )
                print(
                    f"{len(pending)} row(s) need LinkedIn scrape "
                    f"(ddg_discover or empty location/about)",
                    flush=True,
                )

            updated = enrich_dataframe(
                df,
                page=page,
                university=args.university,
                out_path=out_path,
                limit=scrape_limit,
                sleep_seconds=args.sleep,
                start_index=enrich_start,
                refresh=args.refresh_existing,
            )
            browser.close()
    except KeyboardInterrupt:
        print(f"\nSaved partial results ({updated} profile(s)) -> {out_path}")
        return 130

    print(
        f"\nDone: {updated} profile(s) enriched -> {out_path} "
        f"(total run {format_duration(time.perf_counter() - job_started)})",
        flush=True,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
