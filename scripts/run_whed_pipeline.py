"""
Run WHED scraping and officers-social processing in one command.

Default flow:
1) Run whed_scraper.py to generate institutions CSV.
2) Run fetch_officer_socials.py using that CSV as input.

Examples:
  python scripts/run_whed_pipeline.py
  python scripts/run_whed_pipeline.py --institutions-csv data-raw/institutions.csv
  python scripts/run_whed_pipeline.py --whed-args "--max-pages 5"
  python scripts/run_whed_pipeline.py --officers-args "--output output/officers.xlsx --engine auto"
"""

from __future__ import annotations

import argparse
import shlex
import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_WHED = PROJECT_ROOT / "scripts" / "whed_scraper.py"
DEFAULT_OFFICERS = PROJECT_ROOT / "scripts" / "fetch_officer_socials.py"
DEFAULT_INSTITUTIONS_CSV_CANDIDATES = (
    PROJECT_ROOT / "whed_output" / "institutions.csv",
    PROJECT_ROOT / "data-raw" / "institutions.csv",
)


def resolve_path(path_str: str | Path) -> Path:
    p = Path(path_str)
    return p if p.is_absolute() else (PROJECT_ROOT / p)


def run_step(title: str, cmd: list[str]) -> int:
    print(f"\n=== {title} ===")
    print(" ".join(shlex.quote(c) for c in cmd))
    completed = subprocess.run(cmd, cwd=PROJECT_ROOT)
    return completed.returncode


def parse_extra_args(text: str) -> list[str]:
    if not text.strip():
        return []
    return shlex.split(text.strip())


def pick_institutions_csv(path_arg: str) -> Path:
    """
    Resolve institutions CSV from explicit path or known defaults.
    Priority:
      1) user-provided --institutions-csv
      2) whed_output/institutions.csv
      3) data-raw/institutions.csv
    """
    if path_arg.strip():
        return resolve_path(path_arg)
    for candidate in DEFAULT_INSTITUTIONS_CSV_CANDIDATES:
        if candidate.is_file():
            return candidate
    # Keep deterministic error messaging when none exist yet.
    return DEFAULT_INSTITUTIONS_CSV_CANDIDATES[0]


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Automate full pipeline: run whed_scraper.py first, then "
            "fetch_officer_socials.py."
        )
    )
    parser.add_argument(
        "--whed-script",
        default=str(DEFAULT_WHED),
        help=f"Path to whed scraper script (default: {DEFAULT_WHED})",
    )
    parser.add_argument(
        "--officers-script",
        default="",
        help=(
            "Path to officers/social script. "
            "If omitted, auto-detects scripts/fetch_officer_socials.py, "
            "then scripts/fetch_officers_social.py, "
            "then scripts/institution_officers_social.py."
        ),
    )
    parser.add_argument(
        "--institutions-csv",
        default="",
        help=(
            "CSV expected from WHED step and passed to officers step as --input. "
            "If omitted, auto-detects whed_output/institutions.csv, then data-raw/institutions.csv."
        ),
    )
    parser.add_argument(
        "--whed-args",
        default="",
        help='Extra args for whed script, e.g. --whed-args "--max-pages 5"',
    )
    parser.add_argument(
        "--officers-args",
        default="",
        help='Extra args for officers script, e.g. --officers-args "--output output/file.xlsx"',
    )
    parser.add_argument(
        "--skip-whed",
        action="store_true",
        help="Skip WHED step and only run officers step",
    )
    args = parser.parse_args()

    whed_script = resolve_path(args.whed_script)
    if args.officers_script.strip():
        officers_script = resolve_path(args.officers_script)
    else:
        candidates = [
            DEFAULT_OFFICERS,
            PROJECT_ROOT / "scripts" / "fetch_officers_social.py",
            PROJECT_ROOT / "scripts" / "institution_officers_social.py",
        ]
        officers_script = next((p for p in candidates if p.is_file()), candidates[0])
    institutions_csv = pick_institutions_csv(args.institutions_csv)
    whed_extra = parse_extra_args(args.whed_args)
    officers_extra = parse_extra_args(args.officers_args)

    if not args.skip_whed and not whed_script.is_file():
        print(f"WHED script not found: {whed_script}")
        return 1

    if not officers_script.is_file():
        print(f"Officers script not found: {officers_script}")
        return 1

    if not args.skip_whed:
        whed_cmd = [sys.executable, str(whed_script), *whed_extra]
        rc = run_step("Step 1/2: WHED scraping", whed_cmd)
        if rc != 0:
            print(f"WHED step failed with exit code: {rc}")
            return rc

    if not institutions_csv.is_file():
        print(f"Institutions CSV not found after WHED step: {institutions_csv}")
        return 1

    officers_cmd = [
        sys.executable,
        str(officers_script),
        "--input",
        str(institutions_csv),
        *officers_extra,
    ]
    rc = run_step("Step 2/2: Officers social extraction", officers_cmd)
    if rc != 0:
        print(f"Officers step failed with exit code: {rc}")
        return rc

    print("\nPipeline completed successfully.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
