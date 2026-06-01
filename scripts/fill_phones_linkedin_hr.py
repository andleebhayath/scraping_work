"""
Fill missing Company Phone, Company LinkedIn Page, and HR columns in contacts CSV.

Run from project root (needs internet):
  python scripts/fill_phones_linkedin_hr.py
  python scripts/fill_phones_linkedin_hr.py --max-rows 10
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--csv", type=Path, default=ROOT / "output" / "pakistan_tech_companies_contacts.csv")
    p.add_argument("--max-rows", type=int, default=0, help="Only process first N rows that need fill (0=all)")
    args = p.parse_args()
    cmd = [
        sys.executable,
        str(ROOT / "scripts" / "pakistan_tech_companies_data.py"),
        "--fill-missing",
        "--contact-csv",
        str(args.csv),
        "--hr-top",
        "2",
    ]
    if args.max_rows:
        print("Note: --max-rows not yet supported; processing all incomplete rows.")
    return subprocess.call(cmd, cwd=str(ROOT))


if __name__ == "__main__":
    raise SystemExit(main())
