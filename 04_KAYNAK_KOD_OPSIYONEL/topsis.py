"""TOPSIS (Technique for Order Preference by Similarity to Ideal Solution).

Inputs: decision matrix D (m alternatives x n criteria) of arithmetic-mean scores,
        weight vector W (n,) from the aggregated AHP run.

All criteria are benefit criteria per Final.docx section 7.

Steps:
    1. Column-norm n_j = sqrt(sum_i d_ij^2).
    2. Normalized r_ij = d_ij / n_j.
    3. Weighted v_ij = W_j * r_ij.
    4. Ideal A+ = column max, anti-ideal A- = column min (all benefit).
    5. Distances S+_i, S-_i.
    6. Closeness C_i = S-_i / (S+_i + S-_i).
    7. Rank by C_i descending.
"""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np


@dataclass
class TopsisResult:
    decision: np.ndarray
    col_norms: np.ndarray
    normalized: np.ndarray
    weighted: np.ndarray
    ideal_pos: np.ndarray
    ideal_neg: np.ndarray
    s_pos: np.ndarray
    s_neg: np.ndarray
    closeness: np.ndarray
    ranking: np.ndarray  # 0-based indices, best first


def topsis(decision: np.ndarray, weights: np.ndarray) -> TopsisResult:
    d = np.asarray(decision, dtype=float)
    w = np.asarray(weights, dtype=float)
    col_norms = np.sqrt((d ** 2).sum(axis=0))
    normalized = d / col_norms
    weighted = normalized * w
    ideal_pos = weighted.max(axis=0)
    ideal_neg = weighted.min(axis=0)
    s_pos = np.sqrt(((weighted - ideal_pos) ** 2).sum(axis=1))
    s_neg = np.sqrt(((weighted - ideal_neg) ** 2).sum(axis=1))
    closeness = s_neg / (s_pos + s_neg)
    ranking = np.argsort(-closeness)
    return TopsisResult(
        decision=d,
        col_norms=col_norms,
        normalized=normalized,
        weighted=weighted,
        ideal_pos=ideal_pos,
        ideal_neg=ideal_neg,
        s_pos=s_pos,
        s_neg=s_neg,
        closeness=closeness,
        ranking=ranking,
    )
