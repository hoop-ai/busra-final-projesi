"""ELECTRE I on the same weighted-normalized matrix used by TOPSIS.

Procedure (Final.docx section 9 + standard ELECTRE I):
    1. Build weighted-normalized matrix V (already produced by TOPSIS).
    2. For each ordered pair (p, q) with p != q:
        concordance set C(p,q) = { j : V_pj >= V_qj }   (all criteria are benefit)
        discordance set D(p,q) = { j : V_pj <  V_qj }
        c(p,q) = sum_{j in C} W_j
        d(p,q) = max_{j in D} |V_pj - V_qj|  /  max_j |V_pj - V_qj|
        d(p,q) = 0 if D is empty.
    3. Thresholds: c_bar = mean of off-diagonal c, d_bar = mean of off-diagonal d.
    4. Outranking f(p,q) = 1 if c(p,q) >= c_bar AND d(p,q) <= d_bar.
    5. Net score: row sum of f minus column sum of f. Rank descending.
"""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np


@dataclass
class ElectreResult:
    weighted: np.ndarray
    concordance: np.ndarray
    discordance: np.ndarray
    c_bar: float
    d_bar: float
    outranking: np.ndarray
    net_score: np.ndarray
    ranking: np.ndarray


def electre(weighted: np.ndarray, weights: np.ndarray) -> ElectreResult:
    v = np.asarray(weighted, dtype=float)
    w = np.asarray(weights, dtype=float)
    m = v.shape[0]
    concordance = np.zeros((m, m), dtype=float)
    discordance = np.zeros((m, m), dtype=float)
    for p in range(m):
        for q in range(m):
            if p == q:
                continue
            ge_mask = v[p] >= v[q]
            lt_mask = ~ge_mask
            concordance[p, q] = w[ge_mask].sum()
            if lt_mask.any():
                pair_diffs = np.abs(v[p] - v[q])
                pair_max = pair_diffs.max()
                if pair_max == 0:
                    discordance[p, q] = 0.0
                else:
                    discordance[p, q] = pair_diffs[lt_mask].max() / pair_max
            else:
                discordance[p, q] = 0.0

    off_diag = ~np.eye(m, dtype=bool)
    c_bar = float(concordance[off_diag].mean())
    d_bar = float(discordance[off_diag].mean())

    outranking = ((concordance >= c_bar) & (discordance <= d_bar) & off_diag).astype(int)
    net_score = outranking.sum(axis=1) - outranking.sum(axis=0)
    ranking = np.argsort(-net_score)
    return ElectreResult(
        weighted=v,
        concordance=concordance,
        discordance=discordance,
        c_bar=c_bar,
        d_bar=d_bar,
        outranking=outranking,
        net_score=net_score,
        ranking=ranking,
    )
