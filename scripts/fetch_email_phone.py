"""
Fetch Company Email + Company Phone into pakistan_tech_companies_contacts.csv.

Run (close Excel first):
  cd scraping_work
  python scripts/fetch_email_phone.py
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def main() -> int:
    cmd = [
        sys.executable,
        str(ROOT / "scripts" / "pakistan_tech_companies_data.py"),
        "--fetch-email-phone",
        "--sleep",
        "1",
        "--engine",
        "auto",
    ]
    print("Fetching company email and phone from websites + search...")
    return subprocess.call(cmd, cwd=str(ROOT))


if __name__ == "__main__":
    raise SystemExit(main())
