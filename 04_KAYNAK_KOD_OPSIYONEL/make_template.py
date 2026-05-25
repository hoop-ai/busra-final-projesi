"""Generate an empty CSV template that Büşra can fill with her real survey
responses. Same column order as the live Google Form; same Saaty values
expected in the AHP cells and 0-100 (or `Kullanmıyorum`) in the TOPSIS cells.
"""

from __future__ import annotations

import argparse
import csv
from pathlib import Path

import schema


def make(out_path: Path, rows: int = 35) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    cols = schema.csv_columns()
    with out_path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(cols)
        for i in range(1, rows + 1):
            row = [""] * len(cols)
            row[0] = i
            writer.writerow(row)
    print(f"Template saved: {out_path}  ({rows} empty rows)")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=int, default=35)
    parser.add_argument("--out", type=Path, default=Path("../data/responses_template.csv"))
    args = parser.parse_args()
    make(args.out, args.rows)
