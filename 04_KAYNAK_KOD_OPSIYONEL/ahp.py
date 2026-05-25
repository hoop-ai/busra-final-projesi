"""AHP (Analytic Hierarchy Process) for the Decision Support Systems final project.

Per participant:
    1. Build 5x5 reciprocal pairwise matrix A from the 10 survey comparisons.
    2. Column-normalize A -> N.
    3. Eigenvector w = row means of N.
    4. lambda_max = mean of (A @ w) / w.
    5. CI = (lambda_max - n) / (n - 1).
    6. CR = CI / RI, with RI = 1.11 for n = 5.
    7. Keep participant if CR <= 0.10.

Aggregation across consistent participants:
    - Per criterion: geometric mean of weights, then normalize to sum to 1.

Public surface: build_pairwise_matrix, ahp_participant, aggregate_weights.
"""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np

N_CRITERIA = 5
RI_TABLE = {2: 0.0, 3: 0.52, 4: 0.89, 5: 1.11, 6: 1.25, 7: 1.35, 8: 1.40, 9: 1.45, 10: 1.49}
CR_THRESHOLD = 0.10

PAIRWISE_INDEX = [
    (0, 1), (0, 2), (0, 3), (0, 4),
    (1, 2), (1, 3), (1, 4),
    (2, 3), (2, 4),
    (3, 4),
]


@dataclass
class AhpResult:
    pairwise: np.ndarray
    col_sums: np.ndarray
    normalized: np.ndarray
    weights: np.ndarray
    aw: np.ndarray
    aw_over_w: np.ndarray
    lambda_max: float
    ci: float
    cr: float
    consistent: bool


def build_pairwise_matrix(answers: list[float]) -> np.ndarray:
    """Build the 5x5 reciprocal pairwise matrix from the 10 ordered answers.

    answers[k] uses the Saaty convention:
        positive value v >= 1 means the LEFT criterion is preferred by v.
        value 1/v means the RIGHT criterion is preferred by v.
    """
    if len(answers) != len(PAIRWISE_INDEX):
        raise ValueError(f"Expected {len(PAIRWISE_INDEX)} answers, got {len(answers)}")
    matrix = np.ones((N_CRITERIA, N_CRITERIA), dtype=float)
    for value, (i, j) in zip(answers, PAIRWISE_INDEX):
        matrix[i, j] = value
        matrix[j, i] = 1.0 / value
    return matrix


def ahp_participant(answers: list[float]) -> AhpResult:
    a = build_pairwise_matrix(answers)
    col_sums = a.sum(axis=0)
    normalized = a / col_sums
    weights = normalized.mean(axis=1)
    aw = a @ weights
    aw_over_w = aw / weights
    lambda_max = aw_over_w.mean()
    n = N_CRITERIA
    ci = (lambda_max - n) / (n - 1)
    ri = RI_TABLE[n]
    cr = ci / ri
    return AhpResult(
        pairwise=a,
        col_sums=col_sums,
        normalized=normalized,
        weights=weights,
        aw=aw,
        aw_over_w=aw_over_w,
        lambda_max=float(lambda_max),
        ci=float(ci),
        cr=float(cr),
        consistent=bool(cr <= CR_THRESHOLD),
    )


def aggregate_weights(weight_vectors: list[np.ndarray]) -> np.ndarray:
    """Geometric mean per criterion across consistent participants, normalized to sum 1."""
    if not weight_vectors:
        raise ValueError("No consistent participants to aggregate")
    stacked = np.vstack(weight_vectors)
    log_means = np.log(stacked).mean(axis=0)
    gmeans = np.exp(log_means)
    return gmeans / gmeans.sum()
