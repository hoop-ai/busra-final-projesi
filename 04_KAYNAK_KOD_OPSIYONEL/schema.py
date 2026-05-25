"""Canonical column order for the survey CSV.

Mirrors the live Google Form item order documented in SURVEY.md so the Excel
workbook can replay each participant's answers in the same order the survey
posed them. Changing this order invalidates the project per the assignment.
"""

from __future__ import annotations

CRITERIA_TR = [
    "K1 Ödeme Seçenekleri ve Güvenliği",
    "K2 Kullanıcı Deneyimi",
    "K3 Kargo Süresi ve Teslimat Güvenilirliği",
    "K4 Müşteri Yorumları ve Puanlamalar",
    "K5 Fiyatlandırma ve Kampanyalar",
]

CRITERIA_SHORT_TR = ["Ödeme", "Kullanıcı Deneyimi", "Kargo", "Müşteri Yorumları", "Fiyatlandırma"]
CRITERIA_EN = [
    "Payment Options & Security",
    "User Experience",
    "Shipping & Delivery",
    "Customer Reviews & Ratings",
    "Pricing & Campaigns",
]

ALTERNATIVES_TR = [
    "Trendyol",
    "N11",
    "Hepsiburada",
    "PttAVM",
    "Çiçek Sepeti",
    "Amazon Türkiye",
]

PAIRWISE_LABELS_TR = [
    "Ödeme vs Kullanıcı Deneyimi",
    "Ödeme vs Kargo",
    "Ödeme vs Müşteri Yorumları",
    "Ödeme vs Fiyatlandırma",
    "Kullanıcı Deneyimi vs Kargo",
    "Kullanıcı Deneyimi vs Müşteri Yorumları",
    "Kullanıcı Deneyimi vs Fiyatlandırma",
    "Kargo vs Müşteri Yorumları",
    "Kargo vs Fiyatlandırma",
    "Müşteri Yorumları vs Fiyatlandırma",
]

DEMOGRAPHIC_COLS = [
    "cinsiyet",          # gender
    "yas",               # age band
    "egitim",            # education
    "meslek",            # profession
    "alisveris_sikligi", # online-shopping frequency
    "cihaz",             # preferred device
]

AHP_COLS = [f"ahp_{i+1:02d}" for i in range(10)]


def topsis_col(criterion_idx: int, alt_idx: int) -> str:
    """Column name for TOPSIS performance score on (criterion, alternative)."""
    return f"topsis_K{criterion_idx+1}_A{alt_idx+1}"


def topsis_cols() -> list[str]:
    cols = []
    for k in range(len(CRITERIA_TR)):
        for a in range(len(ALTERNATIVES_TR)):
            cols.append(topsis_col(k, a))
    return cols


def csv_columns() -> list[str]:
    return ["participant_id", *DEMOGRAPHIC_COLS, *AHP_COLS, *topsis_cols()]
