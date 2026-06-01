"""
Scrape each company's Contact Us page for:
  - Email address
  - Phone / UAN / Tel
  - Office address (Pakistan block preferred, like avanzasolutions.com/contact-us)

Close Excel, then run:
  cd scraping_work
  python scripts/scrape_company_contact_pages.py

Updates after every company (crash-safe):
  output/pakistan_tech_companies_contacts.csv
  output/pakistan_tech_companies.xlsx (sheet "Company Contacts", if that file exists as --output)
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
        "--scrape-contact-pages",
        "--sleep",
        "1",
    ]
    print("Scraping Contact Us pages for email, phone, and address...")
    return subprocess.call(cmd, cwd=str(ROOT))


if __name__ == "__main__":
    raise SystemExit(main())
