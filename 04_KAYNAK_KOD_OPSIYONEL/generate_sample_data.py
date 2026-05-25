"""Generate 35 synthetic survey responses for demonstration.

WARNING: this is DEMO DATA, not real responses. It exists so Büşra can see the
whole pipeline working end-to-end before plugging in her real survey answers.

Strategy:
    For each participant, draw a "true" weight vector around a population mean
    using a Dirichlet so weights are valid. Map the ideal pairwise ratios
    w_i / w_j onto the discrete Saaty scale {1/9..1..9} with a small chance of
    jitter so a handful of participants end up inconsistent (CR > 0.10) — that
    is realistic and exercises the filtering code path.

    For TOPSIS scores, draw each alternative's per-criterion latent quality
    around brand-level means, jitter per participant, snap to {10,20,..,100}.
"""

from __future__ import annotations

import argparse
import csv
import random
from pathlib import Path

import numpy as np

import schema

SAATY_SCALE = np.array([1/9, 1/7, 1/5, 1/3, 1, 3, 5, 7, 9])

POP_CRITERIA_MEAN = np.array([0.20, 0.16, 0.24, 0.12, 0.28])

ALT_BRAND_MEANS = np.array([
    [80, 80, 75, 80, 80],
    [60, 60, 60, 60, 65],
    [80, 80, 80, 75, 75],
    [55, 50, 55, 50, 55],
    [70, 75, 70, 75, 60],
    [85, 85, 80, 80, 75],
])

GENDER = ["Kadın", "Erkek", "Belirtmek istemiyorum"]
AGE = ["18-24", "25-34", "35-44", "45-54", "55+"]
EDUCATION = ["Lise", "Ön Lisans", "Lisans", "Yüksek Lisans", "Doktora"]
PROFESSION = ["Öğrenci", "Özel Sektör", "Kamu", "Serbest Meslek", "Emekli", "Diğer"]
FREQUENCY = ["Haftada birkaç kez", "Haftada bir", "Ayda birkaç kez", "Ayda bir", "Daha az"]
DEVICE = ["Mobil uygulama", "Mobil tarayıcı", "Masaüstü tarayıcı", "Tablet"]


def snap_to_saaty(ratio: float, jitter_p: float, rng: random.Random) -> float:
    if rng.random() < jitter_p:
        return float(rng.choice([3, 5, 1/3, 1/5]))
    idx = int(np.argmin(np.abs(SAATY_SCALE - ratio)))
    return float(SAATY_SCALE[idx])


def snap_to_topsis(value: float, rng: random.Random) -> int | str:
    if rng.random() < 0.05:
        return "Kullanmıyorum"
    value = max(10, min(100, value))
    return int(round(value / 10.0) * 10)


def generate_participant(pid: int, rng: random.Random, np_rng: np.random.Generator) -> dict:
    weights = np_rng.dirichlet(POP_CRITERIA_MEAN * 60)
    jitter_p = rng.choices([0.05, 0.20, 0.55], weights=[0.7, 0.2, 0.1])[0]
    row = {
        "participant_id": pid,
        "cinsiyet": rng.choices(GENDER, weights=[0.55, 0.42, 0.03])[0],
        "yas": rng.choices(AGE, weights=[0.30, 0.40, 0.18, 0.08, 0.04])[0],
        "egitim": rng.choices(EDUCATION, weights=[0.15, 0.20, 0.45, 0.17, 0.03])[0],
        "meslek": rng.choices(PROFESSION, weights=[0.30, 0.35, 0.10, 0.15, 0.05, 0.05])[0],
        "alisveris_sikligi": rng.choices(FREQUENCY, weights=[0.20, 0.25, 0.30, 0.15, 0.10])[0],
        "cihaz": rng.choices(DEVICE, weights=[0.55, 0.25, 0.15, 0.05])[0],
    }
    for k, (i, j) in enumerate([
        (0, 1), (0, 2), (0, 3), (0, 4),
        (1, 2), (1, 3), (1, 4),
        (2, 3), (2, 4),
        (3, 4),
    ]):
        ratio = weights[i] / weights[j]
        row[f"ahp_{k+1:02d}"] = snap_to_saaty(ratio, jitter_p, rng)

    for k in range(5):
        for a in range(6):
            base = ALT_BRAND_MEANS[a, k]
            noisy = base + np_rng.normal(0, 8)
            row[schema.topsis_col(k, a)] = snap_to_topsis(noisy, rng)
    return row


def generate(n: int, seed: int, out_path: Path) -> None:
    rng = random.Random(seed)
    np_rng = np.random.default_rng(seed)
    rows = [generate_participant(pid, rng, np_rng) for pid in range(1, n + 1)]
    cols = schema.csv_columns()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=cols)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)
    print(f"Wrote {n} rows to {out_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--n", type=int, default=35)
    parser.add_argument("--seed", type=int, default=20260525)
    parser.add_argument("--out", type=Path, default=Path("../data/sample_responses.csv"))
    args = parser.parse_args()
    generate(args.n, args.seed, args.out)
