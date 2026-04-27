"""
Read output/institutions_officers_social.xlsx (from institutions_officers_social_excel.py)
and print a summary and optional row listings for quick QA.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_EXCEL = PROJECT_ROOT / "output" / "institutions_officers_social.xlsx"

EXPECTED_COLS = (
    "iau_id",
    "university_name",
    "officer_role",
    "officer_name",
    "job_title",
    "search_query",
    "search_engine",
    "error",
    "linkedin_url",
    "other_social_urls",
    "university_page_url",
    "official_www",
    "top_urls",
)


def _nonempty(s) -> bool:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return False
    t = str(s).strip()
    return bool(t)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Read institutions officers social Excel and print review summary."
    )
    parser.add_argument(
        "excel",
        type=Path,
        nargs="?",
        default=DEFAULT_EXCEL,
        help=f"Path to .xlsx (default: {DEFAULT_EXCEL})",
    )
    parser.add_argument(
        "--show-all",
        action="store_true",
        help="Print every row (wide table); default is summary only",
    )
    parser.add_argument(
        "--head",
        type=int,
        default=15,
        help="When using --show-all, limit to first N rows (0 = no limit; default: 15)",
    )
    parser.add_argument(
        "--missing-linkedin",
        action="store_true",
        help="List rows with no linkedin_url",
    )
    parser.add_argument(
        "--errors",
        action="store_true",
        help="List rows with non-empty error",
    )
    parser.add_argument(
        "--export-csv",
        type=Path,
        help="Optional path to re-export the sheet as UTF-8 CSV (for diffs, etc.)",
    )
    args = parser.parse_args()

    path: Path = args.excel
    if not path.is_file():
        print(f"File not found: {path}", file=sys.stderr)
        return 1

    df = pd.read_excel(path, engine="openpyxl", dtype=str)
    df = df.fillna("")

    missing = [c for c in EXPECTED_COLS if c not in df.columns]
    if missing:
        print(f"Note: missing columns (file may be older or different): {', '.join(missing)}")
    print(f"File: {path.resolve()}")
    print(f"Rows: {len(df)}")

    has_in = df["linkedin_url"].map(_nonempty) if "linkedin_url" in df.columns else pd.Series([False] * len(df))
    has_social = (
        df["other_social_urls"].map(_nonempty) if "other_social_urls" in df.columns else pd.Series([False] * len(df))
    )
    has_uni = (
        df["university_page_url"].map(_nonempty)
        if "university_page_url" in df.columns
        else pd.Series([False] * len(df))
    )
    has_err = df["error"].map(_nonempty) if "error" in df.columns else pd.Series([False] * len(df))

    print()
    print("--- Summary ---")
    print(f"  LinkedIn URL found:     {int(has_in.sum())}")
    print(f"  Other social URLs:      {int(has_social.sum())}")
    print(f"  University page guess:  {int(has_uni.sum())}")
    print(f"  Search / network errors: {int(has_err.sum())}")
    no_hit = int(((~has_in) & (~has_social) & (~has_uni) & (~has_err)).sum())
    print(f"  No URL extracted (and no error): {no_hit}")

    if len(df) and "search_engine" in df.columns:
        print()
        print("--- search_engine (row counts) ---")
        print(df["search_engine"].replace("", "(empty)").value_counts().to_string())

    if args.errors and "error" in df.columns:
        bad = df[df["error"].map(_nonempty)]
        print()
        print(f"--- Rows with error ({len(bad)}) ---")
        cols = [c for c in ("university_name", "officer_name", "search_query", "error") if c in bad.columns]
        with pd.option_context("display.max_colwidth", 80, "display.width", 200, "display.max_rows", 200):
            print(bad[cols].to_string(index=True))

    if args.missing_linkedin and "linkedin_url" in df.columns:
        m = df[~df["linkedin_url"].map(_nonempty)]
        print()
        print(f"--- No LinkedIn URL ({len(m)}) ---")
        cols = [c for c in ("university_name", "officer_name", "job_title", "search_query") if c in m.columns]
        with pd.option_context("display.max_colwidth", 100, "display.width", 220, "display.max_rows", 80):
            print(m[cols].to_string(index=True))
        if len(m) > 80:
            print(f"... and {len(m) - 80} more (narrow with grep or export CSV)")

    if args.show_all:
        n = args.head
        to_show = df if n == 0 else df.head(n)
        print()
        print(f"--- All columns (head {n if n else 'all'}) ---")
        with pd.option_context("display.max_colwidth", 60, "display.width", 240, "display.max_rows", 500):
            print(to_show.to_string(index=True))

    if args.export_csv:
        out = args.export_csv.resolve()
        out.parent.mkdir(parents=True, exist_ok=True)
        df.to_csv(out, index=False, encoding="utf-8-sig")
        print()
        print(f"Exported CSV: {out}")

    if not (args.show_all or args.missing_linkedin or args.errors or args.export_csv):
        print()
        print("Tip: run with --show-all, --head 20, --missing-linkedin, --errors, and/or --export-csv path.csv")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
