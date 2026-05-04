"""
Read data-raw/institutions.csv (name + officers_json), search per officer like
uni_data_free_google.py, classify LinkedIn / other social / university-ish URLs,
and write an Excel workbook.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
import requests

# Reuse search stack from sibling script
_SCRIPTS = Path(__file__).resolve().parent
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

import uni_data_free_google as udg  # noqa: E402

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = PROJECT_ROOT / "data-raw" / "institutions.csv"
DEFAULT_OUTPUT = PROJECT_ROOT / "output" / "institutions_officers_social.xlsx"

SOCIAL_HOST_FRAGMENTS = (
    "twitter.com",
    "x.com",
    "facebook.com",
    "instagram.com",
    "researchgate.net",
    "orcid.org",
    "scholar.google.com",
    "github.com",
    "youtube.com",
)


def _host_key(url: str) -> str:
    try:
        return urlparse(url).netloc.lower().replace("www.", "")
    except Exception:
        return ""


def _official_host(www: str) -> str:
    s = (www or "").strip()
    if not s:
        return ""
    if not s.startswith(("http://", "https://")):
        s = "https://" + s
    return _host_key(s)


def classify_urls(urls: list[str], official_www: str) -> tuple[str, str, str]:
    """
    From ranked search URLs, pick:
    - linkedin profile URL (first /in/ profile)
    - other social URLs (semicolon-separated)
    - university-related page URL (prefer official www host, else academic-looking)
    """
    official = _official_host(official_www)
    linkedin = ""
    socials: list[str] = []
    uni_candidates: list[str] = []
    seen: set[str] = set()

    for u in urls:
        if not u or not u.startswith("http"):
            continue
        if u in seen:
            continue
        seen.add(u)
        host = _host_key(u)
        low = u.lower()

        if udg.is_profile_like_linkedin(u):
            if not linkedin:
                linkedin = u
            continue

        if "linkedin.com" in host and "/in/" in low and "/company/" not in low:
            if not linkedin:
                linkedin = u
            continue

        if any(frag in host for frag in SOCIAL_HOST_FRAGMENTS):
            socials.append(u)
            continue

        if official and (host == official or host.endswith("." + official)):
            uni_candidates.append(u)
            continue

        if any(
            x in host
            for x in (
                ".edu",
                ".ac.",
                "edu.pk",
                "univ",
                "university",
            )
        ) and "linkedin.com" not in host:
            uni_candidates.append(u)

    uni_page = uni_candidates[0] if uni_candidates else ""
    other_social = "; ".join(socials[:12])
    return linkedin, other_social, uni_page


def parse_officers(raw: str) -> list[dict[str, str]]:
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
        name = str(item.get("name") or "").strip()
        if not name:
            continue
        out.append(
            {
                "role": str(item.get("role") or "").strip(),
                "name": name,
                "job_title": str(item.get("job_title") or "").strip(),
            }
        )
    return out


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Search social/uni URLs for each officer in institutions.csv and save Excel."
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=DEFAULT_INPUT,
        help=f"Path to institutions CSV (default: {DEFAULT_INPUT})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=f"Excel output path (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--engine",
        choices=("auto", "ddgs", "duckduckgo_html", "jina_ddg"),
        default="auto",
        help="Search backend (default: auto)",
    )
    parser.add_argument("--max-results", type=int, default=12, help="Max URLs per search (default: 12)")
    parser.add_argument("--delay-seconds", type=float, default=1.0, help="Pause before each search (default: 1.0)")
    parser.add_argument("--max-institutions", type=int, default=0, help="Process only first N institutions (0=all)")
    parser.add_argument("--max-officers", type=int, default=0, help="Cap officers per institution (0=no cap)")
    args = parser.parse_args()

    inp = args.input.resolve()
    if not inp.is_file():
        print(f"Input not found: {inp}")
        return 1

    df = pd.read_csv(inp, dtype=str, keep_default_na=False)
    for col in ("name", "officers_json"):
        if col not in df.columns:
            print(f"CSV missing required column: {col}")
            return 1

    has_www = "www" in df.columns
    has_id = "iau_id" in df.columns

    session = requests.Session()
    session.headers.update(udg.browser_headers())

    rows_out: list[dict[str, str]] = []
    n_inst = 0
    for idx, row in df.iterrows():
        if args.max_institutions and n_inst >= args.max_institutions:
            break
        uni_name = str(row.get("name", "")).strip()
        officers_raw = row.get("officers_json", "")
        officers = parse_officers(officers_raw)
        if not officers:
            continue
        n_inst += 1
        official_www = str(row.get("www", "")).strip() if has_www else ""
        iau = str(row.get("iau_id", "")).strip() if has_id else ""

        if args.max_officers:
            officers = officers[: args.max_officers]

        for off in officers:
            person = off["name"]
            designation = off["job_title"] or off["role"]
            keyword = "linkedin"
            search_query = udg.build_query(uni_name, person, designation, keyword)

            time.sleep(max(args.delay_seconds, 0))
            try:
                urls, engine_used = udg.search_urls(session, search_query, args.max_results, args.engine)
            except (requests.RequestException, RuntimeError) as exc:
                rows_out.append(
                    {
                        "iau_id": iau,
                        "university_name": uni_name,
                        "officer_role": off["role"],
                        "officer_name": person,
                        "job_title": off["job_title"],
                        "search_query": search_query,
                        "search_engine": "",
                        "error": str(exc),
                        "linkedin_url": "",
                        "other_social_urls": "",
                        "university_page_url": "",
                        "official_www": official_www,
                        "top_urls": "",
                    }
                )
                continue

            linkedin_u, social_u, uni_u = classify_urls(urls, official_www)
            top_joined = " | ".join(urls[: args.max_results])

            rows_out.append(
                {
                    "iau_id": iau,
                    "university_name": uni_name,
                    "officer_role": off["role"],
                    "officer_name": person,
                    "job_title": off["job_title"],
                    "search_query": search_query,
                    "search_engine": engine_used,
                    "error": "",
                    "linkedin_url": linkedin_u,
                    "other_social_urls": social_u,
                    "university_page_url": uni_u,
                    "official_www": official_www,
                    "top_urls": top_joined,
                }
            )

    out_path = args.output.resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    result = pd.DataFrame(rows_out)
    result.to_excel(out_path, index=False, engine="openpyxl")
    print(f"Wrote {len(result)} rows to {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
