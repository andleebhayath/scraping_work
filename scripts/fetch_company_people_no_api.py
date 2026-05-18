"""
Fetch LinkedIn people search results for company names without paid APIs.

This script reads company names (default: first_50_companies.xlsx), runs web search
queries, and stores top profile matches with basic metadata in Excel.

Run:
    python scripts/fetch_company_people_no_api.py
"""

from __future__ import annotations

import random
import re
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, quote_plus, unquote, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup

try:
    from googlesearch import search as google_search
except Exception:
    google_search = None
try:
    from ddgs import DDGS  # preferred package
except Exception:
    try:
        from duckduckgo_search import DDGS  # backward compatibility
    except Exception:
        DDGS = None

WORKSPACE_ROOT = Path(__file__).resolve().parents[1]
INPUT_50_PATH = WORKSPACE_ROOT / "data-processed" / "first_50_companies.xlsx"
RAW_INPUT_PATH = WORKSPACE_ROOT / "data-raw" / "all_linkedin_companies_merged.xlsx"
OUTPUT_PATH = WORKSPACE_ROOT / "output" / "company_people_linkedin_top5.xlsx"

MAX_COMPANIES = 50
TOP_PROFILES_PER_COMPANY = 5
SEARCH_ENGINE_GOOGLE = "google_scrape"
SEARCH_ENGINE_DDG = "duckduckgo_html"
SEARCH_ENGINE_JINA_DDG = "jina_proxy_ddg"
SEARCH_ENGINE_DDGS = "ddgs"
JINA_PROXY_PREFIX = "https://r.jina.ai/http://"

# Set True for a quick dry run (faster testing).
QUICK_TEST_MODE = False
QUICK_TEST_MAX_COMPANIES = 20

# Modes:
# - fast: much quicker, fewer queries/retries (recommended while iterating)
# - deep: slower, wider query coverage
MODE = "fast"

if MODE == "deep":
    REQUEST_TIMEOUT = 25
    SLEEP_MIN_SECONDS = 1.2
    SLEEP_MAX_SECONDS = 2.4
    MAX_RETRIES_PER_ENGINE = 2
    BACKOFF_BASE_SECONDS = 1.5
    MAX_ROLE_QUERIES = 15
    ENABLE_DDGS_FALLBACK = True
    ENABLE_GOOGLE_FALLBACK = True
    ENABLE_DDG_FALLBACK = True
else:
    REQUEST_TIMEOUT = 8
    SLEEP_MIN_SECONDS = 0.1
    SLEEP_MAX_SECONDS = 0.25
    MAX_RETRIES_PER_ENGINE = 1
    BACKOFF_BASE_SECONDS = 0.4
    MAX_ROLE_QUERIES = 2
    ENABLE_DDGS_FALLBACK = True
    ENABLE_GOOGLE_FALLBACK = False
    ENABLE_DDG_FALLBACK = False

ROLE_KEYWORDS = [
    "CEO",
    "Founder",
    "Co-Founder",
    "CTO",
    "CIO",
    "CISO",
    "VP Engineering",
    "Engineering Manager",
    "Cyber Security",
    "HR",
    "Recruiter",
    "Talent Acquisition",
    "Sales Manager",
    "Marketing Manager",
    "Product Manager",
]

DEFAULT_INDUSTRY_HINT = "quantum jobs"
DEFAULT_DOMAIN_HINT = "quantum"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}


def _check_jina_ddg_health() -> tuple[bool, str]:
    try:
        probe_query = "site:linkedin.com/in quantum"
        probe_url = JINA_PROXY_PREFIX + f"https://html.duckduckgo.com/html/?q={quote_plus(probe_query)}"
        response = requests.get(probe_url, headers=HEADERS, timeout=min(REQUEST_TIMEOUT, 10))
        text = response.text.lower()
        if response.status_code >= 400:
            return False, f"HTTP {response.status_code}"
        if "forbidden" in text or "error-lite" in text or "returned error 403" in text:
            return False, "blocked (DDG 403 via Jina)"
        if "read timed out" in text or "timed out" in text:
            return False, "timeout"
        return True, "ok"
    except Exception as exc:
        return False, str(exc)


def _check_google_health() -> tuple[bool, str]:
    if google_search is None:
        return False, "googlesearch package unavailable"
    try:
        response = requests.get(
            "https://www.google.com/search",
            params={"q": "site:linkedin.com/in test"},
            headers=HEADERS,
            timeout=min(REQUEST_TIMEOUT, 10),
        )
        text = response.text.lower()
        if response.status_code == 429:
            return False, "blocked (HTTP 429)"
        if "unusual traffic" in text or "captcha" in text:
            return False, "blocked (captcha/rate limit)"
        return True, "ok"
    except Exception as exc:
        return False, str(exc)


def _check_ddg_health() -> tuple[bool, str]:
    try:
        response = requests.get(
            "https://html.duckduckgo.com/html/",
            params={"q": "site:linkedin.com/in test"},
            headers=HEADERS,
            timeout=min(REQUEST_TIMEOUT, 10),
        )
        text = response.text.lower()
        if response.status_code in (202, 403):
            return False, f"blocked (HTTP {response.status_code})"
        if "forbidden" in text or "error-lite" in text:
            return False, "blocked (forbidden)"
        return True, "ok"
    except Exception as exc:
        return False, str(exc)


def _check_ddgs_health() -> tuple[bool, str]:
    if DDGS is None:
        return False, "ddgs package unavailable"
    try:
        with DDGS() as client:
            rows = list(client.text("openai", max_results=3))
        if rows:
            return True, "ok"
        return False, "no results from ddgs"
    except Exception as exc:
        return False, str(exc)


def print_engine_health(active_engines: list[str]) -> None:
    checks = {
        SEARCH_ENGINE_DDGS: _check_ddgs_health,
        SEARCH_ENGINE_JINA_DDG: _check_jina_ddg_health,
        SEARCH_ENGINE_GOOGLE: _check_google_health,
        SEARCH_ENGINE_DDG: _check_ddg_health,
    }
    print("Engine health check:")
    for engine in active_engines:
        check_fn = checks.get(engine)
        if check_fn is None:
            print(f"  - {engine}: unknown")
            continue
        ok, detail = check_fn()
        if ok:
            print(f"  - {engine}: OK")
        else:
            print(f"  - {engine}: BLOCKED/UNHEALTHY ({detail}) -> retry later")


def detect_company_column(columns: list[str]) -> str:
    normalized = {c.strip().lower(): c for c in columns}
    for candidate in ("company_name", "name", "company", "organization", "organization_name"):
        if candidate in normalized:
            return normalized[candidate]
    raise ValueError(f"Could not detect company name column. Available columns: {columns}")


def find_optional_column(columns: list[str], candidates: tuple[str, ...]) -> str:
    normalized = {c.strip().lower(): c for c in columns}
    for candidate in candidates:
        if candidate in normalized:
            return normalized[candidate]
    return ""


def infer_domain_hint(company_name: str) -> str:
    tokens = [t for t in re.split(r"[^a-zA-Z0-9]+", company_name.lower()) if t]
    if not tokens:
        return DEFAULT_DOMAIN_HINT
    # Keep meaningful words only.
    stop_words = {"the", "and", "of", "for", "at", "in", "inc", "ltd", "llc", "group", "company"}
    meaningful = [t for t in tokens if t not in stop_words]
    if not meaningful:
        meaningful = tokens[:1]
    return " ".join(meaningful[:2])


def load_company_records() -> list[dict]:
    source_path = INPUT_50_PATH if INPUT_50_PATH.is_file() else RAW_INPUT_PATH
    if not source_path.is_file():
        raise FileNotFoundError(
            "No company input file found. Expected one of:\n"
            f"- {INPUT_50_PATH}\n"
            f"- {RAW_INPUT_PATH}"
        )

    df = pd.read_excel(source_path, dtype=str).fillna("")
    company_col = detect_company_column(list(df.columns))
    domain_col = find_optional_column(list(df.columns), ("domain", "website_domain", "official_www", "website"))
    industry_col = find_optional_column(list(df.columns), ("industry", "sector", "category"))

    company_limit = QUICK_TEST_MAX_COMPANIES if QUICK_TEST_MODE else MAX_COMPANIES
    records: list[dict] = []
    seen_companies: set[str] = set()
    for _, row in df.iterrows():
        company_name = str(row.get(company_col, "")).strip()
        if not company_name:
            continue
        company_key = company_name.lower()
        if company_key in seen_companies:
            continue
        seen_companies.add(company_key)

        raw_domain = str(row.get(domain_col, "")).strip() if domain_col else ""
        raw_industry = str(row.get(industry_col, "")).strip() if industry_col else ""
        domain_hint = raw_domain if raw_domain else infer_domain_hint(company_name)
        industry_hint = raw_industry if raw_industry else DEFAULT_INDUSTRY_HINT

        records.append(
            {
                "company_name": company_name,
                "domain_hint": domain_hint,
                "industry_hint": industry_hint,
            }
        )
        if len(records) >= company_limit:
            break
    return records


def build_query(company_name: str, domain_hint: str, industry_hint: str) -> str:
    return f'site:linkedin.com/in "{company_name}" ("{domain_hint}" OR "{industry_hint}")'


def build_query_variations(company_name: str, domain_hint: str, industry_hint: str) -> list[tuple[str, str]]:
    """
    Return list of (query_text, role_hint) in priority order.
    role_hint is empty for generic company-only query.
    """
    base_strict = build_query(company_name, domain_hint, industry_hint)
    base_broad_quoted = f'site:linkedin.com/in "{company_name}" LinkedIn'
    company_plain = re.sub(r"\s+", " ", re.sub(r"[^a-zA-Z0-9]+", " ", company_name)).strip()
    base_broad_plain = f"site:linkedin.com/in {company_plain} LinkedIn" if company_plain else base_broad_quoted

    # Ordered from strict to broad; broad fallbacks improve hit rate for noisy company names.
    raw_queries: list[tuple[str, str]] = [
        (base_strict, ""),
        (base_broad_quoted, ""),
        (base_broad_plain, ""),
        (f'site:linkedin.com/in "{domain_hint}"', ""),
        (f'site:linkedin.com/in "{industry_hint}"', ""),
    ]
    for role in ROLE_KEYWORDS[:MAX_ROLE_QUERIES]:
        raw_queries.append((f'{base_broad_quoted} "{role}"', role))

    deduped: list[tuple[str, str]] = []
    seen: set[str] = set()
    for query, role_hint in raw_queries:
        key = query.strip().lower()
        if key and key not in seen:
            seen.add(key)
            deduped.append((query, role_hint))
    return deduped


def clean_search_url(href: str) -> str:
    if not href:
        return ""
    if "uddg=" in href:
        query = parse_qs(urlparse(href).query)
        values = query.get("uddg", [])
        if values:
            return unquote(values[0])
    if href.startswith("//"):
        return "https:" + href
    return href


def is_linkedin_profile(url: str) -> bool:
    try:
        parsed = urlparse(url)
        return "linkedin.com" in parsed.netloc.lower() and parsed.path.lower().startswith("/in/")
    except Exception:
        return False


def name_from_title(title: str) -> str:
    cleaned = re.sub(r"\s*\|\s*LinkedIn\s*$", "", title, flags=re.IGNORECASE).strip()
    if " - " in cleaned:
        return cleaned.split(" - ")[0].strip()
    return cleaned


def fetch_from_google_package(company_name: str, query: str, role_hint: str, top_n: int) -> list[dict]:
    if google_search is None:
        return []

    rows: list[dict] = []
    seen: set[str] = set()

    # advanced=True returns result objects with title/description/url in this package.
    for item in google_search(
        query,
        num_results=top_n * 3,
        lang="en",
        advanced=True,
        sleep_interval=0.2 if MODE == "fast" else 1,
        timeout=REQUEST_TIMEOUT,
    ):
        profile_url = (getattr(item, "url", "") or "").strip()
        if not profile_url or not is_linkedin_profile(profile_url) or profile_url in seen:
            continue
        seen.add(profile_url)

        title = (getattr(item, "title", "") or "").strip()
        snippet = (getattr(item, "description", "") or "").strip()

        rows.append(
            {
                "company_name": company_name,
                "person_name": name_from_title(title) if title else "",
                "linkedin_url": profile_url,
                "result_title": title,
                "snippet": snippet,
                "search_query": query,
                "search_engine": SEARCH_ENGINE_GOOGLE,
                "role_hint": role_hint,
                "error_text": "",
            }
        )
        if len(rows) >= top_n:
            break

    return rows


def fetch_from_ddgs(company_name: str, query: str, role_hint: str, top_n: int) -> list[dict]:
    if DDGS is None:
        return []

    rows: list[dict] = []
    seen: set[str] = set()

    with DDGS() as client:
        for item in client.text(query, region="wt-wt", safesearch="off", max_results=top_n * 4):
            profile_url = clean_search_url((item.get("href") or item.get("url") or "").strip())
            if not profile_url or not is_linkedin_profile(profile_url) or profile_url in seen:
                continue
            seen.add(profile_url)

            title = (item.get("title") or "").strip()
            snippet = (item.get("body") or item.get("snippet") or "").strip()
            rows.append(
                {
                    "company_name": company_name,
                    "person_name": name_from_title(title) if title else "",
                    "linkedin_url": profile_url,
                    "result_title": title,
                    "snippet": snippet,
                    "search_query": query,
                    "search_engine": SEARCH_ENGINE_DDGS,
                    "role_hint": role_hint,
                    "error_text": "",
                }
            )
            if len(rows) >= top_n:
                break

    return rows


def fetch_from_ddg(company_name: str, query: str, role_hint: str, top_n: int) -> list[dict]:
    response = requests.get(
        "https://html.duckduckgo.com/html/",
        params={"q": query},
        headers=HEADERS,
        timeout=REQUEST_TIMEOUT,
    )
    response.raise_for_status()
    if response.status_code == 202:
        raise RuntimeError("DuckDuckGo returned 202 (likely bot protection)")

    soup = BeautifulSoup(response.text, "html.parser")
    rows: list[dict] = []
    seen: set[str] = set()

    for result in soup.select("div.result"):
        link_tag = result.select_one("a.result__a")
        if link_tag is None:
            continue

        raw_url = (link_tag.get("href") or "").strip()
        profile_url = clean_search_url(raw_url)
        if not profile_url or not is_linkedin_profile(profile_url) or profile_url in seen:
            continue
        seen.add(profile_url)

        title = link_tag.get_text(" ", strip=True)
        snippet_tag = result.select_one(".result__snippet")
        snippet = snippet_tag.get_text(" ", strip=True) if snippet_tag else ""

        rows.append(
            {
                "company_name": company_name,
                "person_name": name_from_title(title),
                "linkedin_url": profile_url,
                "result_title": title,
                "snippet": snippet,
                "search_query": query,
                "search_engine": SEARCH_ENGINE_DDG,
                "role_hint": role_hint,
                "error_text": "",
            }
        )
        if len(rows) >= top_n:
            break

    return rows


def fetch_from_jina_proxy_ddg(company_name: str, query: str, role_hint: str, top_n: int) -> list[dict]:
    """
    Fetch DuckDuckGo HTML through Jina proxy and parse profile URLs from text.
    This path is often more resilient when direct search pages throttle requests.
    """
    ddg_url = f"https://html.duckduckgo.com/html/?q={quote_plus(query)}"
    proxy_url = JINA_PROXY_PREFIX + ddg_url
    response = requests.get(proxy_url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    text = response.text

    rows: list[dict] = []
    seen: set[str] = set()
    for match in re.finditer(r"https?://(?:[a-z]{2,3}\.)?linkedin\.com/in/[^\s)\]\"'>]+", text, flags=re.IGNORECASE):
        profile_url = match.group(0).strip().rstrip(".,;)]")
        if not is_linkedin_profile(profile_url) or profile_url in seen:
            continue
        seen.add(profile_url)

        rows.append(
            {
                "company_name": company_name,
                "person_name": "",
                "linkedin_url": profile_url,
                "result_title": "",
                "snippet": "",
                "search_query": query,
                "search_engine": SEARCH_ENGINE_JINA_DDG,
                "role_hint": role_hint,
                "error_text": "",
            }
        )
        if len(rows) >= top_n:
            break

    return rows


def run_with_retries(engine_name: str, fn, *args) -> tuple[list[dict], str]:
    errors: list[str] = []
    for attempt in range(1, MAX_RETRIES_PER_ENGINE + 1):
        try:
            return fn(*args), ""
        except Exception as exc:
            errors.append(f"{engine_name} attempt {attempt}: {exc}")
            if attempt < MAX_RETRIES_PER_ENGINE:
                time.sleep(BACKOFF_BASE_SECONDS * attempt)
    return [], " ; ".join(errors)


def fetch_people_for_company(company_record: dict, top_n: int) -> tuple[list[dict], str]:
    """
    Try multiple role-based query variations and multiple engines with retries.
    Returns (deduplicated_rows, combined_error_text).
    """
    errors: list[str] = []
    combined_rows: list[dict] = []
    seen_urls: set[str] = set()
    company_name = company_record["company_name"]
    domain_hint = company_record["domain_hint"]
    industry_hint = company_record["industry_hint"]

    queries = build_query_variations(company_name, domain_hint, industry_hint)
    for query, role_hint in queries:
        if ENABLE_DDGS_FALLBACK:
            rows, err = run_with_retries(
                SEARCH_ENGINE_DDGS,
                fetch_from_ddgs,
                company_name,
                query,
                role_hint,
                top_n,
            )
            if err:
                errors.append(f'{SEARCH_ENGINE_DDGS} ({query}): {err}')
            for row in rows:
                url = row.get("linkedin_url", "").strip()
                if url and url not in seen_urls:
                    seen_urls.add(url)
                    row["domain_hint"] = domain_hint
                    row["industry_hint"] = industry_hint
                    combined_rows.append(row)
                    if len(combined_rows) >= top_n:
                        return combined_rows, ""

        # Jina proxy + DuckDuckGo path (preferred)
        rows, err = run_with_retries(
            SEARCH_ENGINE_JINA_DDG,
            fetch_from_jina_proxy_ddg,
            company_name,
            query,
            role_hint,
            top_n,
        )
        if err:
            errors.append(f'{SEARCH_ENGINE_JINA_DDG} ({query}): {err}')
        for row in rows:
            url = row.get("linkedin_url", "").strip()
            if url and url not in seen_urls:
                seen_urls.add(url)
                row["domain_hint"] = domain_hint
                row["industry_hint"] = industry_hint
                combined_rows.append(row)
                if len(combined_rows) >= top_n:
                    return combined_rows, ""

        if ENABLE_GOOGLE_FALLBACK:
            rows, err = run_with_retries(
                SEARCH_ENGINE_GOOGLE,
                fetch_from_google_package,
                company_name,
                query,
                role_hint,
                top_n,
            )
            if err:
                errors.append(f'{SEARCH_ENGINE_GOOGLE} ({query}): {err}')
            for row in rows:
                url = row.get("linkedin_url", "").strip()
                if url and url not in seen_urls:
                    seen_urls.add(url)
                    row["domain_hint"] = domain_hint
                    row["industry_hint"] = industry_hint
                    combined_rows.append(row)
                    if len(combined_rows) >= top_n:
                        return combined_rows, ""

        if ENABLE_DDG_FALLBACK:
            rows, err = run_with_retries(
                SEARCH_ENGINE_DDG,
                fetch_from_ddg,
                company_name,
                query,
                role_hint,
                top_n,
            )
            if err:
                errors.append(f'{SEARCH_ENGINE_DDG} ({query}): {err}')
            for row in rows:
                url = row.get("linkedin_url", "").strip()
                if url and url not in seen_urls:
                    seen_urls.add(url)
                    row["domain_hint"] = domain_hint
                    row["industry_hint"] = industry_hint
                    combined_rows.append(row)
                    if len(combined_rows) >= top_n:
                        return combined_rows, ""

    return combined_rows, " | ".join(errors)


def print_preview_and_save(rows_so_far: list[dict], current_rows: list[dict], company: str) -> None:
    if current_rows:
        first = current_rows[0]
        print(
            f"  Preview => name='{first.get('person_name', '')}' | "
            f"url='{first.get('linkedin_url', '')}' | "
            f"engine='{first.get('search_engine', '')}'"
        )
    else:
        print(f"  Preview => no profile found for '{company}'")

    out_df = pd.DataFrame(rows_so_far)
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    try:
        out_df.to_excel(OUTPUT_PATH, index=False)
        print(f"  Saved progress: {len(out_df)} rows -> {OUTPUT_PATH}")
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback_path = OUTPUT_PATH.with_name(f"{OUTPUT_PATH.stem}_autosave_{timestamp}.xlsx")
        out_df.to_excel(fallback_path, index=False)
        print(
            "  Warning: output file is locked (likely open in Excel). "
            f"Saved progress to: {fallback_path}"
        )


def main() -> None:
    company_records = load_company_records()
    all_rows: list[dict] = []
    active_engines: list[str] = []
    if ENABLE_DDGS_FALLBACK:
        active_engines.append(SEARCH_ENGINE_DDGS)
    active_engines.append(SEARCH_ENGINE_JINA_DDG)
    if ENABLE_GOOGLE_FALLBACK:
        active_engines.append(SEARCH_ENGINE_GOOGLE)
    if ENABLE_DDG_FALLBACK:
        active_engines.append(SEARCH_ENGINE_DDG)
    print(
        f"Mode={MODE} | quick_test={QUICK_TEST_MODE} | companies={len(company_records)} "
        f"| role_queries/company={MAX_ROLE_QUERIES + 1} "
        f"| retries/engine={MAX_RETRIES_PER_ENGINE} | engines={','.join(active_engines)}"
    )
    print_engine_health(active_engines)

    for idx, company_record in enumerate(company_records, start=1):
        company = company_record["company_name"]
        print(f"[{idx}/{len(company_records)}] Searching: {company}")
        try:
            rows, err = fetch_people_for_company(company_record, TOP_PROFILES_PER_COMPANY)
            if rows:
                for rank, row in enumerate(rows, start=1):
                    row["result_rank_for_company"] = rank
                all_rows.extend(rows)
                print_preview_and_save(all_rows, rows, company)
            else:
                no_result_row = {
                    "company_name": company,
                    "person_name": "",
                    "linkedin_url": "",
                    "result_title": "",
                    "snippet": "",
                    "search_query": build_query(
                        company,
                        company_record["domain_hint"],
                        company_record["industry_hint"],
                    ),
                    "search_engine": "",
                    "domain_hint": company_record["domain_hint"],
                    "industry_hint": company_record["industry_hint"],
                    "error_text": err or "No LinkedIn profile results found",
                    "result_rank_for_company": "",
                }
                all_rows.append(no_result_row)
                print_preview_and_save(all_rows, [], company)
        except Exception as exc:
            error_row = {
                "company_name": company,
                "person_name": "",
                "linkedin_url": "",
                "result_title": "",
                "snippet": "",
                "search_query": build_query(
                    company,
                    company_record["domain_hint"],
                    company_record["industry_hint"],
                ),
                "search_engine": "",
                "domain_hint": company_record["domain_hint"],
                "industry_hint": company_record["industry_hint"],
                "error_text": str(exc),
                "result_rank_for_company": "",
            }
            all_rows.append(error_row)
            print(f"  Error: {exc}")
            print_preview_and_save(all_rows, [], company)

        time.sleep(random.uniform(SLEEP_MIN_SECONDS, SLEEP_MAX_SECONDS))

    print(f"\nDone. Final saved rows: {len(all_rows)} -> {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
