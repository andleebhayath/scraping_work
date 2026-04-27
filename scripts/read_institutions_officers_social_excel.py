"""
Load output/institutions_officers_social.xlsx into a pandas DataFrame.

  import sys
  sys.path.insert(0, "path/to/scripts")
  from read_institutions_officers_social_excel import load_institutions_officers_social
  df = load_institutions_officers_social()
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_XLSX = PROJECT_ROOT / "output" / "institutions_officers_social.xlsx"


def load_institutions_officers_social(path: Path | str | None = None) -> pd.DataFrame:
    """
    Read the officers social Excel. All cells as strings; NaN -> empty string.
    """
    p = Path(path) if path is not None else DEFAULT_XLSX
    df = pd.read_excel(p, engine="openpyxl", dtype=str)
    return df.fillna("")


def main() -> int:
    parser = argparse.ArgumentParser(description="Read institutions_officers_social.xlsx into pandas.")
    parser.add_argument(
        "excel",
        type=Path,
        nargs="?",
        default=DEFAULT_XLSX,
        help=f"Path to .xlsx (default: {DEFAULT_XLSX})",
    )
    parser.add_argument("--head", type=int, default=10, help="Rows to print (default: 10, 0 = none)")
    args = parser.parse_args()

    if not args.excel.is_file():
        print(f"File not found: {args.excel.resolve()}")
        return 1

    df = load_institutions_officers_social(args.excel)
    print(f"Path:  {args.excel.resolve()}")
    print(f"Shape: {df.shape[0]} rows x {df.shape[1]} columns")
    print(f"Columns: {list(df.columns)}")
    if args.head and len(df):
        print()
        print(df.head(args.head).to_string())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
