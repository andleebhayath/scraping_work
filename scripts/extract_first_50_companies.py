"""
Extract the first N company names from data-raw/all_linkedin_companies_merged.xlsx.

Run:
    python scripts/extract_first_50_companies.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

INPUT_PATH = Path(__file__).resolve().parents[1] / "data-raw" / "all_linkedin_companies_merged.xlsx"
OUTPUT_PATH = Path(__file__).resolve().parents[1] / "data-processed" / "first_50_companies.xlsx"
LIMIT = 50


def detect_company_column(columns: list[str]) -> str:
    """Find the best company-name column using common header variants."""
    normalized = {c.strip().lower(): c for c in columns}
    candidates = ["name", "company", "company_name", "organization", "organization_name"]
    for candidate in candidates:
        if candidate in normalized:
            return normalized[candidate]
    raise ValueError(
        f"Could not detect a company-name column. Available columns: {columns}"
    )


def main() -> None:
    if not INPUT_PATH.is_file():
        raise FileNotFoundError(f"Input file not found: {INPUT_PATH}")

    df = pd.read_excel(INPUT_PATH, dtype=str).fillna("")
    company_col = detect_company_column(list(df.columns))

    out = (
        df[[company_col]]
        .rename(columns={company_col: "company_name"})
        .assign(company_name=lambda d: d["company_name"].astype(str).str.strip())
    )
    out = out[out["company_name"] != ""].head(LIMIT).reset_index(drop=True)
    out.insert(0, "sr_no", range(1, len(out) + 1))

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    out.to_excel(OUTPUT_PATH, index=False)

    print(f"Input rows: {len(df)}")
    print(f"Company column: {company_col}")
    print(f"Saved {len(out)} companies to: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
