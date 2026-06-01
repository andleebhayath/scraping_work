"""
Add companies from master list to contacts CSV and fill website/location/LinkedIn.

Run:
  python scripts/expand_contacts_from_master.py
  python scripts/pakistan_tech_companies_data.py --fill-missing-offline
  python scripts/pakistan_tech_companies_data.py --fill-missing
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

_SCRIPTS = Path(__file__).resolve().parent
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

import pakistan_tech_companies_data as ptc  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
CSV = ROOT / "output" / "pakistan_tech_companies_contacts.csv"
MAX = 50


def main() -> None:
    ptc.load_master_lookup()
    companies = ptc.load_companies(ptc.DEFAULT_MASTER, "", MAX, pakistan_only=False)

    if CSV.is_file():
        df = pd.read_csv(CSV, dtype=str).fillna("")
        if "company_name" in df.columns:
            df = ptc.dataframe_to_team_lead(df)
        existing = {
            str(r.get("Company Name", "")).strip().lower()
            for _, r in df.iterrows()
            if str(r.get("Company Name", "")).strip()
        }
        rows = [dict(r) for _, r in df.iterrows()]
    else:
        existing = set()
        rows = []

    added = 0
    for co in companies:
        name = co["company_name"]
        if name.lower() in existing:
            continue
        rec = ptc.finalize_company_record(
            ptc.enrich_company_from_master(
                ptc.CompanyRecord(
                    company_name=name,
                    city=co.get("city", ""),
                    website=co.get("website", ""),
                    linkedin_company_url=co.get("linkedin_company_url", ""),
                    industry=co.get("industry", "technology"),
                )
            )
        )
        if not rec.company_email and rec.website:
            rec.company_email = ptc.guess_email_from_website(rec.website)
        rows.append(ptc.to_team_lead_row(rec, None))
        existing.add(name.lower())
        added += 1

    for i, row in enumerate(rows):
        rows[i] = ptc.fill_missing_offline_row(
            {c: str(row.get(c, "")).strip() for c in ptc.TEAM_LEAD_COLUMNS}
        )

    out = pd.DataFrame(rows, columns=list(ptc.TEAM_LEAD_COLUMNS))
    CSV.parent.mkdir(parents=True, exist_ok=True)
    out.to_csv(CSV, index=False, encoding="utf-8-sig")
    with_site = sum(1 for _, r in out.iterrows() if str(r.get("Company Website", "")).strip())
    with_email = sum(1 for _, r in out.iterrows() if str(r.get("Company Email", "")).strip())
    print(f"Wrote {len(out)} rows (+{added} new) -> {CSV}")
    print(f"  With website: {with_site}/{len(out)} | With email: {with_email}/{len(out)}")
    print("Next: python scripts/pakistan_tech_companies_data.py --fill-missing")


if __name__ == "__main__":
    main()
