"""Copy company Excel + CSV to output/backups/ (manual backup anytime)."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

_SCRIPTS = Path(__file__).resolve().parent
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

from pakistan_tech_companies_data import (  # noqa: E402
    DEFAULT_BACKUP_DIR,
    DEFAULT_OUTPUT,
    backup_output_files,
    contact_csv_path,
)


def main() -> int:
    p = argparse.ArgumentParser(description="Backup pakistan tech companies output files")
    p.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    p.add_argument("--backup-dir", type=Path, default=DEFAULT_BACKUP_DIR)
    p.add_argument("--label", default="manual", help="Tag in backup filename")
    args = p.parse_args()
    csv_path = contact_csv_path(args.output)
    saved = backup_output_files(
        args.output.resolve(),
        csv_path,
        backup_dir=args.backup_dir,
        label=args.label,
    )
    if not saved:
        print(f"Nothing to backup. Missing: {args.output} and/or {csv_path}")
        return 1
    print(f"Backup folder: {args.backup_dir.resolve()}")
    for path in saved:
        print(f"  {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
