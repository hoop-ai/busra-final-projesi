"""Verify the workbook by evaluating its formulas with pycel and comparing the
results against the Python pipeline. Run with Python 3.12 (pycel is not yet
3.14-compatible):  py -3.12 verify_workbook.py

What it catches: wrong cell references, off-by-one row layouts, missing
absolute anchors, typos in formula strings, mismatched aggregation logic.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import numpy as np

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

from pycel import ExcelCompiler  # noqa: E402

TOL = 1e-4


def load_pipeline_results(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def get_value(ec: ExcelCompiler, sheet: str, cell: str):
    return ec.evaluate(f"{sheet}!{cell}")


def almost_eq(a, b, tol=TOL):
    try:
        return abs(float(a) - float(b)) < tol
    except (TypeError, ValueError):
        return False


def check_per_participant(ec, expected):
    failures = 0
    for p in expected["per_participant"][:5]:
        pid = p["participant_id"]
        sheet = f"Katilimci {pid:02d}"
        for i in range(5):
            got = get_value(ec, sheet, f"B{29 + i}")
            want = p["weights"][i]
            if not almost_eq(got, want):
                print(f"FAIL {sheet} weight {i}: got {got} want {want}")
                failures += 1
        # CR is at B40 on each participant sheet
        got_cr = get_value(ec, sheet, "B40")
        if not almost_eq(got_cr, p["cr"], tol=1e-3):
            print(f"FAIL {sheet} CR: got {got_cr} want {p['cr']}")
            failures += 1
    return failures


def check_aggregation(ec, expected):
    failures = 0
    # final normalized W: row = 3 + n_participants + 2 + 2
    final_row = 3 + expected["n_participants"] + 2 + 2
    for i in range(5):
        cell = f"{chr(ord('B') + i)}{final_row}"
        got = get_value(ec, "AHP_Birlestirme", cell)
        want = expected["final_weights"][i]
        if not almost_eq(got, want):
            print(f"FAIL final W[{i}]: got {got} want {want}")
            failures += 1
    return failures


def check_decision_matrix(ec, expected):
    failures = 0
    for a in range(6):
        for k in range(5):
            cell = f"{chr(ord('B') + k)}{4 + a}"
            got = get_value(ec, "Karar_Matrisi", cell)
            want = expected["decision_matrix"][a][k]
            if not almost_eq(got, want, tol=1e-2):
                print(f"FAIL D[{a},{k}]: got {got} want {want}")
                failures += 1
    return failures


def check_topsis(ec, expected):
    failures = 0
    # TOPSIS layout (alt_count = 6):
    #   dm_top=7
    #   cn_row = 7 + 2 + 6 + 1 = 16        (column norms)
    #   norm_top = 16 + 2 = 18             (normalized matrix header)
    #   w_top = 18 + 2 + 6 + 1 = 27        (weighted matrix header)
    #   ideal_top = 27 + 2 + 6 + 1 = 36    (A+ row)
    #   closeness_top = 36 + 3 = 39        (closeness table header)
    closeness_top = 39
    for a in range(6):
        cell = f"D{closeness_top + 1 + a}"
        got = get_value(ec, "TOPSIS", cell)
        want = expected["topsis"]["closeness"][a]
        if not almost_eq(got, want, tol=1e-3):
            print(f"FAIL TOPSIS C[{a}]: got {got} want {want}")
            failures += 1
    return failures


def check_electre(ec, expected):
    failures = 0
    # electre_rank_top: conc_top=16, disc_top=16+2+6+1=25, thresh_top=25+2+6+1=34,
    # out_top=34+3=37, electre_rank_top=37+2+6+1=46
    electre_rank_top = 46
    for a in range(6):
        cell = f"D{electre_rank_top + 1 + a}"
        got = get_value(ec, "ELECTRE", cell)
        want = expected["electre"]["net_score"][a]
        if not almost_eq(got, want, tol=0.5):
            print(f"FAIL ELECTRE net[{a}]: got {got} want {want}")
            failures += 1
    return failures


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlsx", type=Path, default=Path("../workbook/Busra_Final_Project.xlsx"))
    parser.add_argument("--results", type=Path, default=Path("../output/results.json"))
    args = parser.parse_args()

    expected = load_pipeline_results(args.results)
    print(f"Compiling workbook {args.xlsx} ...")
    ec = ExcelCompiler(filename=str(args.xlsx))

    total = 0
    total += check_per_participant(ec, expected)
    total += check_aggregation(ec, expected)
    total += check_decision_matrix(ec, expected)
    total += check_topsis(ec, expected)
    total += check_electre(ec, expected)
    if total == 0:
        print("OK: workbook formulas match the Python pipeline numbers.")
    else:
        print(f"FAILURES: {total}")
        sys.exit(1)


if __name__ == "__main__":
    main()
