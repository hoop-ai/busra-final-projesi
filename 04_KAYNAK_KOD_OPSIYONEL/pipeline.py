"""End-to-end pipeline: CSV -> per-participant AHP -> filter -> aggregate ->
TOPSIS -> ELECTRE -> JSON + console summary.

Run from this directory:
    python pipeline.py --csv ../data/sample_responses.csv --out ../output

The pipeline is deterministic given the same CSV: re-running produces identical
JSON, which is what the assignment expects (every cell traceable).
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

import schema
from ahp import ahp_participant, aggregate_weights
from electre import electre
from topsis import topsis


def load_responses(csv_path: Path) -> pd.DataFrame:
    df = pd.read_csv(csv_path, encoding="utf-8-sig")
    missing = [c for c in schema.csv_columns() if c not in df.columns]
    if missing:
        raise ValueError(f"CSV is missing columns: {missing}")
    return df


def per_participant_ahp(df: pd.DataFrame) -> list[dict]:
    results = []
    for _, row in df.iterrows():
        answers = [float(row[c]) for c in schema.AHP_COLS]
        r = ahp_participant(answers)
        results.append({
            "participant_id": int(row["participant_id"]),
            "weights": r.weights.tolist(),
            "lambda_max": r.lambda_max,
            "ci": r.ci,
            "cr": r.cr,
            "consistent": r.consistent,
        })
    return results


def build_decision_matrix(df: pd.DataFrame) -> np.ndarray:
    """Arithmetic mean per (alternative, criterion), ignoring 'Kullanmıyorum'."""
    m, n = len(schema.ALTERNATIVES_TR), len(schema.CRITERIA_TR)
    decision = np.zeros((m, n), dtype=float)
    for k in range(n):
        for a in range(m):
            col = schema.topsis_col(k, a)
            values = pd.to_numeric(df[col], errors="coerce")
            decision[a, k] = values.mean(skipna=True)
    return decision


def run(csv_path: Path, out_dir: Path) -> dict:
    df = load_responses(csv_path)
    out_dir.mkdir(parents=True, exist_ok=True)

    ahp_rows = per_participant_ahp(df)
    consistent = [r for r in ahp_rows if r["consistent"]]
    consistent_weights = [np.array(r["weights"]) for r in consistent]
    if len(consistent_weights) < 20:
        print(f"WARNING: only {len(consistent_weights)} consistent participants (need >= 20).")

    final_weights = aggregate_weights(consistent_weights)
    decision = build_decision_matrix(df)

    t = topsis(decision, final_weights)
    e = electre(t.weighted, final_weights)

    summary = {
        "n_participants": len(ahp_rows),
        "n_consistent": len(consistent),
        "consistent_threshold": 0.10,
        "criteria_tr": schema.CRITERIA_TR,
        "alternatives_tr": schema.ALTERNATIVES_TR,
        "final_weights": final_weights.tolist(),
        "decision_matrix": decision.tolist(),
        "topsis": {
            "closeness": t.closeness.tolist(),
            "s_pos": t.s_pos.tolist(),
            "s_neg": t.s_neg.tolist(),
            "ideal_pos": t.ideal_pos.tolist(),
            "ideal_neg": t.ideal_neg.tolist(),
            "ranking_idx": t.ranking.tolist(),
            "ranking_named": [schema.ALTERNATIVES_TR[i] for i in t.ranking],
        },
        "electre": {
            "c_bar": e.c_bar,
            "d_bar": e.d_bar,
            "net_score": e.net_score.tolist(),
            "outranking": e.outranking.tolist(),
            "ranking_idx": e.ranking.tolist(),
            "ranking_named": [schema.ALTERNATIVES_TR[i] for i in e.ranking],
        },
        "per_participant": ahp_rows,
    }
    (out_dir / "results.json").write_text(
        json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    print_summary(summary)
    return summary


def print_summary(s: dict) -> None:
    print("=" * 72)
    print(f"Participants total : {s['n_participants']}")
    print(f"Consistent (CR<=0.10): {s['n_consistent']}")
    print()
    print("Final AHP weights (geometric mean, normalized):")
    for c, w in zip(s["criteria_tr"], s["final_weights"]):
        print(f"  {c:<55} {w:.4f}")
    print()
    print("TOPSIS ranking:")
    for rank, (name, idx) in enumerate(zip(s["topsis"]["ranking_named"], s["topsis"]["ranking_idx"]), 1):
        print(f"  {rank}. {name:<20} C={s['topsis']['closeness'][idx]:.4f}")
    print()
    print("ELECTRE ranking:")
    for rank, (name, idx) in enumerate(zip(s["electre"]["ranking_named"], s["electre"]["ranking_idx"]), 1):
        print(f"  {rank}. {name:<20} net={s['electre']['net_score'][idx]:+d}")
    print("=" * 72)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--csv", type=Path, default=Path("../data/sample_responses.csv"))
    parser.add_argument("--out", type=Path, default=Path("../output"))
    args = parser.parse_args()
    run(args.csv, args.out)
