"""
Build output/pakistan_tech_companies.xlsx from master list + contacts CSV.

Run after scraping (or standalone to merge seed + existing contacts CSV):
  python scripts/build_10_companies_excel.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SEED = PROJECT_ROOT / "data-raw" / "pakistan_tech_companies_master.csv"
FALLBACK_SEED = PROJECT_ROOT / "data-raw" / "pakistan_tech_companies_seed.csv"
CONTACTS_CSV = PROJECT_ROOT / "output" / "pakistan_tech_companies_contacts.csv"
OUTPUT_XLSX = PROJECT_ROOT / "output" / "pakistan_tech_companies.xlsx"
OUTPUT_CSV = PROJECT_ROOT / "output" / "pakistan_tech_companies_contacts.csv"

TEAM_LEAD_COLS = [
    "Company Name",
    "Industry/Sector",
    "Company Website",
    "Company Email",
    "Company Phone",
    "Company LinkedIn Page",
    "Location",
    "HR Contact Name",
    "HR Designation",
    "HR Email",
    "HR Phone",
    "HR LinkedIn Profile",
]


def legacy_row_to_team_lead(row: dict[str, str]) -> dict[str, str]:
    """Convert old company_name / hr_name CSV rows to team-lead columns."""
    return {
        "Company Name": str(row.get("company_name", row.get("Company Name", ""))).strip(),
        "Industry/Sector": str(row.get("industry", "Technology")).strip() or "Technology",
        "Company Website": str(row.get("website", row.get("Company Website", ""))).strip(),
        "Company Email": str(row.get("company_email", row.get("Company Email", row.get("Email Address", "")))).strip(),
        "Company Phone": str(row.get("company_phone", row.get("Company Phone", row.get("Phone Number", "")))).strip(),
        "Company LinkedIn Page": str(
            row.get(
                "linkedin_company_url",
                row.get("Company LinkedIn Page", row.get("LinkedIn Profile or Company LinkedIn Page", "")),
            )
        ).strip(),
        "Location": str(row.get("address", row.get("Location", ""))).strip(),
        "HR Contact Name": str(row.get("hr_name", row.get("HR Contact Name", row.get("Contact Person Name", "")))).strip(),
        "HR Designation": str(row.get("hr_job_title", row.get("HR Designation", row.get("Designation", "")))).strip(),
        "HR Email": str(row.get("hr_email", row.get("HR Email", ""))).strip(),
        "HR Phone": str(row.get("hr_phone", row.get("HR Phone", ""))).strip(),
        "HR LinkedIn Profile": str(
            row.get("hr_linkedin_url", row.get("HR LinkedIn Profile", ""))
        ).strip(),
    }

# Known contacts (website / public pages) — used when scraper has not run yet.
KNOWN: list[dict[str, str]] = [
    {
        "company_name": "TRG Pakistan",
        "company_email": "investor.relations@trg.com.pk",
        "company_phone": "+92 21 111 874 874",
        "address": "24th floor, Sky Tower, Clifton, Karachi-75600, Pakistan",
        "linkedin_company_url": "https://www.linkedin.com/company/trg-pakistan",
        "hr_name": "Rahat Latif",
        "hr_email": "investor.relations@trg.com.pk",
        "hr_phone": "+92 21 111 874 874",
        "hr_job_title": "Investor Relations / Complaints",
        "hr_linkedin_url": "",
    },
    {
        "company_name": "Folio3",
        "company_email": "info@folio3.com",
        "company_phone": "+92 21 34380081",
        "address": "Karachi, Pakistan",
        "linkedin_company_url": "https://www.linkedin.com/company/folio3",
        "hr_name": "",
        "hr_email": "",
        "hr_phone": "",
        "hr_job_title": "",
        "hr_linkedin_url": "",
    },
    {
        "company_name": "NetSol Technologies",
        "company_email": "info@netsoltech.com",
        "company_phone": "+92 42 111 638 765",
        "address": "NetSol Avenue, Lahore, Pakistan",
        "linkedin_company_url": "https://www.linkedin.com/company/netsol-technologies",
        "hr_name": "",
        "hr_email": "",
        "hr_phone": "",
        "hr_job_title": "",
        "hr_linkedin_url": "",
    },
    {
        "company_name": "Confiz",
        "company_email": "info@confiz.com",
        "company_phone": "+92 21 111 124 249",
        "address": "Karachi, Pakistan",
        "linkedin_company_url": "https://www.linkedin.com/company/confiz-ltd",
        "hr_name": "",
        "hr_email": "",
        "hr_phone": "",
        "hr_job_title": "",
        "hr_linkedin_url": "",
    },
    {
        "company_name": "VentureDive",
        "company_email": "hello@venturedive.com",
        "company_phone": "+92 21 34320701",
        "address": "Karachi, Pakistan",
        "linkedin_company_url": "https://www.linkedin.com/company/venturedive",
        "hr_name": "",
        "hr_email": "",
        "hr_phone": "",
        "hr_job_title": "",
        "hr_linkedin_url": "",
    },
    {
        "company_name": "Tkxel",
        "company_email": "info@tkxel.com",
        "company_phone": "+92 42 35757862",
        "address": "Lahore, Pakistan",
        "linkedin_company_url": "https://www.linkedin.com/company/tkxel",
        "hr_name": "",
        "hr_email": "",
        "hr_phone": "",
        "hr_job_title": "",
        "hr_linkedin_url": "",
    },
    {
        "company_name": "Devsinc",
        "company_email": "info@devsinc.com",
        "company_phone": "+92 21 35309821",
        "address": "Karachi, Pakistan",
        "linkedin_company_url": "https://www.linkedin.com/company/devsinc",
        "hr_name": "",
        "hr_email": "",
        "hr_phone": "",
        "hr_job_title": "",
        "hr_linkedin_url": "",
    },
]


def load_seed_names() -> pd.DataFrame:
    path = SEED if SEED.is_file() else FALLBACK_SEED
    df = pd.read_csv(path, dtype=str).fillna("")
    return df.rename(columns={"name": "company_name"})


def load_scraped_contacts() -> pd.DataFrame:
    if not CONTACTS_CSV.is_file():
        return pd.DataFrame(columns=TEAM_LEAD_COLS)
    df = pd.read_csv(CONTACTS_CSV, dtype=str).fillna("")
    if "Company Name" in df.columns:
        for col in TEAM_LEAD_COLS:
            if col not in df.columns:
                df[col] = ""
        return df[TEAM_LEAD_COLS]
    # Legacy columns: company_email, linkedin_company_url, hr_name, …
    if "company_name" in df.columns:
        rows = [legacy_row_to_team_lead(dict(r)) for _, r in df.iterrows()]
        return pd.DataFrame(rows, columns=TEAM_LEAD_COLS)
    return df


def merge_contacts(scraped: pd.DataFrame, known: list[dict[str, str]]) -> pd.DataFrame:
    legacy_cols = list(known[0].keys()) if known else []
    known_df = pd.DataFrame(known, columns=legacy_cols)
    combined = pd.concat([scraped, known_df], ignore_index=True)
    combined = combined[combined["company_name"].astype(str).str.strip() != ""]

    def prefer_row(group: pd.DataFrame) -> pd.Series:
        g = group.copy()
        for col in ("company_email", "company_phone", "linkedin_company_url"):
            filled = g[g[col].astype(str).str.strip() != ""]
            if not filled.empty:
                g = filled
        return g.iloc[0]

    base = (
        combined.groupby("company_name", as_index=False)
        .apply(prefer_row, include_groups=False)
        .reset_index(drop=True)
    )

    hr_rows = combined[
        combined["hr_name"].astype(str).str.strip() != ""
        | combined["hr_linkedin_url"].astype(str).str.strip() != ""
    ]
    if hr_rows.empty:
        return base

    out_rows: list[dict[str, str]] = []
    for _, b in base.iterrows():
        name = str(b["company_name"]).strip()
        hrs = hr_rows[hr_rows["company_name"] == name]
        if hrs.empty:
            out_rows.append(legacy_row_to_team_lead(dict(b)))
            continue
        for _, h in hrs.drop_duplicates(subset=["hr_linkedin_url", "hr_name"]).iterrows():
            merged = dict(b)
            merged["hr_name"] = str(h.get("hr_name", ""))
            merged["hr_email"] = str(h.get("hr_email", "")) or str(merged.get("hr_email", ""))
            merged["hr_phone"] = str(h.get("hr_phone", ""))
            merged["hr_job_title"] = str(h.get("hr_job_title", ""))
            merged["hr_linkedin_url"] = str(h.get("hr_linkedin_url", ""))
            out_rows.append(legacy_row_to_team_lead(merged))
    return pd.DataFrame(out_rows, columns=TEAM_LEAD_COLS)


def main() -> None:
    seed = load_seed_names()
    scraped = load_scraped_contacts()
    if scraped is not None and not scraped.empty and "Company Name" in scraped.columns:
        contacts = scraped
    else:
        merged = merge_contacts(scraped, KNOWN)
        if "Company Name" in merged.columns:
            contacts = merged
        else:
            contacts = pd.DataFrame(
                [legacy_row_to_team_lead(dict(r)) for _, r in merged.iterrows()],
                columns=TEAM_LEAD_COLS,
            )

    if "Company Name" in contacts.columns:
        out = contacts.copy()
        for col in TEAM_LEAD_COLS:
            if col not in out.columns:
                out[col] = ""
        out = out[TEAM_LEAD_COLS]
    else:
        out = contacts

    OUTPUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    out.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")

    companies_sheet = seed.copy()

    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        out.to_excel(writer, sheet_name="Company Contacts", index=False)
        companies_sheet.to_excel(writer, sheet_name="company_list", index=False)

    n_co = out["Company Name"].nunique() if "Company Name" in out.columns else len(out)
    print(f"Wrote {n_co} companies ({len(out)} rows)")
    print(f"  Excel: {OUTPUT_XLSX}")
    print(f"  CSV:   {OUTPUT_CSV}")
    print(f"  Rows in contacts sheet: {len(out)}")


if __name__ == "__main__":
    main()
