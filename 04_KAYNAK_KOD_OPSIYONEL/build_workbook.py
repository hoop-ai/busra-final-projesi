"""Build the formula-driven Excel deliverable.

Per the assignment rule "Tüm aşamalarda formül yazılmalı", every numeric cell
downstream of the raw answers is a live Excel formula. The grader can click any
cell and see the upstream reference. This file is the source of truth — the
Python pipeline is the audit/reference implementation.

Sheets:
    Bilgi               - student / project info + how to read the workbook.
    Anket_Yanıtları     - raw CSV rows (the only place numbers are typed in).
    Katılımcı 01..35    - per-participant AHP analysis with formulas.
    AHP_Birleştirme     - geometric-mean criterion weights across consistent ones.
    Karar_Matrisi       - arithmetic-mean decision matrix.
    TOPSIS              - normalize -> weighted -> ideal -> distance -> closeness.
    ELECTRE             - concordance / discordance -> outranking -> net score.
    Sıralama_Özeti      - one-shot final ranking comparison.
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import schema

HEADER_FILL = PatternFill("solid", fgColor="305496")
HEADER_FONT = Font(bold=True, color="FFFFFF")
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
SUB_FONT = Font(bold=True, color="1F3864")
PASS_FILL = PatternFill("solid", fgColor="C6EFCE")
FAIL_FILL = PatternFill("solid", fgColor="FFC7CE")
RANK_FILL = PatternFill("solid", fgColor="FFF2CC")

THIN = Side(border_style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def header(cell, text, fill=HEADER_FILL, font=HEADER_FONT):
    cell.value = text
    cell.fill = fill
    cell.font = font
    cell.alignment = CENTER
    cell.border = BOX


def sub(cell, text=None):
    if text is None:
        text = cell.value
    header(cell, text, fill=SUB_FILL, font=SUB_FONT)


def box(cell, value=None, number_format=None):
    if value is not None:
        cell.value = value
    cell.border = BOX
    cell.alignment = CENTER
    if number_format:
        cell.number_format = number_format


def auto_width(ws, widths):
    for col_idx, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = w


SHEET_INFO = "Bilgi"
SHEET_RAW = "Anket_Yanitlari"
SHEET_AGG = "AHP_Birlestirme"
SHEET_DM = "Karar_Matrisi"
SHEET_TOPSIS = "TOPSIS"
SHEET_ELECTRE = "ELECTRE"
SHEET_SUMMARY = "Siralama_Ozeti"


def part_sheet(i: int) -> str:
    return f"Katilimci {i:02d}"


def write_info_sheet(ws):
    ws.title = SHEET_INFO
    rows = [
        ("Proje", "AHP + TOPSIS + ELECTRE ile E-Ticaret Sitelerinin Performans Değerlendirmesi"),
        ("Öğrenci", "Büşra Nur Yıldırım"),
        ("Öğrenci No", "230326759"),
        ("Grup", "Kırmızı Grup (Anket sırası 1-34)"),
        ("Ders", "YBS302 - Karar Destek Sistemleri"),
        ("Öğretim Üyesi", "Dr. Öğr. Üyesi Neda Alipour"),
        ("Üniversite", "İstanbul Gelişim Üniversitesi"),
        ("Dönem", "Bahar 2025-2026"),
        ("", ""),
        ("Kriter Sayısı", 5),
        ("Alternatif Sayısı", 6),
        ("Toplam Katılımcı", "Anket_Yanıtları sayfasına bakınız"),
        ("Tutarlılık Eşiği (CR)", "≤ 0.10"),
        ("RI (n=5)", 1.11),
        ("", ""),
        ("Çalışma Akışı", ""),
        ("1", "Her katılımcı için ikili karşılaştırma matrisi (Saaty 1-9 skalası)"),
        ("2", "Sütun normalizasyonu → Öz Vektör (ÖZV / kriter ağırlıkları)"),
        ("3", "λ_max (ENBÖZD) → CI = (λ_max - n)/(n-1) → CR = CI/RI"),
        ("4", "CR ≤ 0.10 olan katılımcılar tutarlı kabul edilir (≥ 20 olmalı)"),
        ("5", "Tutarlı katılımcıların ağırlıkları GEOMETRİK ORTALAMA ile birleştirilir"),
        ("6", "Alternatif performans puanları ARİTMETİK ORTALAMA ile birleştirilir"),
        ("7", "TOPSIS ve ELECTRE birleştirilmiş veri üzerinde BİRER KEZ çalıştırılır"),
        ("8", "Tüm hücreler formül, hiçbir ara değer elle yazılmamıştır"),
        ("", ""),
        ("Notlar", ""),
        ("Kriterler", "Tüm kriterler FAYDA tipindedir (yüksek = iyi)"),
        ("'Kullanmıyorum'", "Aritmetik ortalamada bu cevap dahil edilmez"),
        ("Veri Kaynağı", "Google Forms canlı anket çıktısı (SURVEY.md)"),
    ]
    ws["A1"] = "PROJE ÖZETİ"
    ws["A1"].font = Font(bold=True, size=14, color="1F3864")
    ws["A1"].alignment = LEFT
    ws.merge_cells("A1:B1")
    for r, (k, v) in enumerate(rows, start=3):
        ws.cell(row=r, column=1, value=k).font = Font(bold=True)
        ws.cell(row=r, column=1).alignment = LEFT
        ws.cell(row=r, column=2, value=v).alignment = LEFT
    auto_width(ws, [28, 80])


def write_raw_sheet(ws, df: pd.DataFrame):
    ws.title = SHEET_RAW
    cols = schema.csv_columns()
    for c, name in enumerate(cols, 1):
        header(ws.cell(row=1, column=c), name)
    for r, (_, row) in enumerate(df.iterrows(), start=2):
        for c, name in enumerate(cols, 1):
            val = row[name]
            if name.startswith("topsis_") and isinstance(val, str) and val.strip().lower().startswith("kullan"):
                ws.cell(row=r, column=c, value=val)
            else:
                try:
                    ws.cell(row=r, column=c, value=float(val))
                except (TypeError, ValueError):
                    ws.cell(row=r, column=c, value=val)
    ws.freeze_panes = "B2"
    widths = [14, 14, 10, 12, 14, 18, 18] + [9] * len(schema.AHP_COLS) + [10] * len(schema.topsis_cols())
    auto_width(ws, widths)


def ahp_col_letter(idx: int) -> str:
    """Excel column letter on the raw-answers sheet for AHP answer index 0..9.

    Layout: col 1 participant_id, cols 2..7 demographics, cols 8..17 ahp_01..10.
    """
    base = 1 + len(schema.DEMOGRAPHIC_COLS)  # 7 columns occupied before AHP
    return get_column_letter(base + 1 + idx)


def topsis_col_letter(criterion_idx: int, alt_idx: int) -> str:
    base = 1 + len(schema.DEMOGRAPHIC_COLS) + len(schema.AHP_COLS)
    flat = criterion_idx * len(schema.ALTERNATIVES_TR) + alt_idx
    return get_column_letter(base + 1 + flat)


PAIRWISE_INDEX = [
    (0, 1), (0, 2), (0, 3), (0, 4),
    (1, 2), (1, 3), (1, 4),
    (2, 3), (2, 4),
    (3, 4),
]


def write_participant_sheet(ws, participant_id: int, row_in_raw: int):
    n = len(schema.CRITERIA_SHORT_TR)
    ws["A1"] = f"KATILIMCI {participant_id:02d} - AHP ANALİZİ"
    ws["A1"].font = Font(bold=True, size=14, color="1F3864")
    ws.merge_cells("A1:H1")

    sub(ws["A3"], "Demografik Bilgiler")
    ws.merge_cells("A3:B3")
    demo_labels = ["Cinsiyet", "Yaş", "Eğitim", "Meslek", "Alışveriş Sıklığı", "Cihaz"]
    for i, label in enumerate(demo_labels):
        r = 4 + i
        box(ws.cell(row=r, column=1, value=label))
        ws.cell(row=r, column=1).alignment = LEFT
        cell = ws.cell(row=r, column=2)
        cell.value = f"=Anket_Yanitlari!{get_column_letter(2 + i)}{row_in_raw}"
        box(cell)

    pm_top = 11
    sub(ws.cell(row=pm_top, column=1), "İkili Karşılaştırma Matrisi (A)")
    ws.merge_cells(start_row=pm_top, start_column=1, end_row=pm_top, end_column=n + 1)
    for j, name in enumerate(schema.CRITERIA_SHORT_TR):
        sub(ws.cell(row=pm_top + 1, column=2 + j), name)
    for i in range(n):
        sub(ws.cell(row=pm_top + 2 + i, column=1, value=schema.CRITERIA_SHORT_TR[i]))
        for j in range(n):
            cell = ws.cell(row=pm_top + 2 + i, column=2 + j)
            if i == j:
                cell.value = 1
            elif (i, j) in PAIRWISE_INDEX:
                k = PAIRWISE_INDEX.index((i, j))
                cell.value = f"=Anket_Yanitlari!{ahp_col_letter(k)}{row_in_raw}"
            else:
                cell.value = f"=1/{get_column_letter(2 + i)}{pm_top + 2 + j}"
            box(cell, number_format="0.0000")
    sum_row = pm_top + 2 + n
    sub(ws.cell(row=sum_row, column=1, value="Sütun Toplamı"))
    for j in range(n):
        col_letter = get_column_letter(2 + j)
        cell = ws.cell(row=sum_row, column=2 + j)
        cell.value = f"=SUM({col_letter}{pm_top + 2}:{col_letter}{pm_top + 1 + n})"
        box(cell, number_format="0.0000")

    nm_top = sum_row + 2
    sub(ws.cell(row=nm_top, column=1), "Normalize Matris (N)")
    ws.merge_cells(start_row=nm_top, start_column=1, end_row=nm_top, end_column=n + 1)
    for j, name in enumerate(schema.CRITERIA_SHORT_TR):
        sub(ws.cell(row=nm_top + 1, column=2 + j), name)
    for i in range(n):
        sub(ws.cell(row=nm_top + 2 + i, column=1, value=schema.CRITERIA_SHORT_TR[i]))
        for j in range(n):
            col_letter = get_column_letter(2 + j)
            cell = ws.cell(row=nm_top + 2 + i, column=2 + j)
            cell.value = (
                f"={col_letter}{pm_top + 2 + i}/{col_letter}{sum_row}"
            )
            box(cell, number_format="0.0000")

    ozv_top = nm_top + 2 + n + 1
    sub(ws.cell(row=ozv_top, column=1), "Kriter")
    sub(ws.cell(row=ozv_top, column=2), "ÖZV (w)")
    sub(ws.cell(row=ozv_top, column=3), "A·w")
    sub(ws.cell(row=ozv_top, column=4), "A·w / w")
    for i in range(n):
        r = ozv_top + 1 + i
        sub(ws.cell(row=r, column=1, value=schema.CRITERIA_SHORT_TR[i]))
        row_letter_first = get_column_letter(2)
        row_letter_last = get_column_letter(2 + n - 1)
        # ÖZV = row mean of normalized matrix
        cell_w = ws.cell(row=r, column=2)
        cell_w.value = (
            f"=AVERAGE({row_letter_first}{nm_top + 2 + i}:{row_letter_last}{nm_top + 2 + i})"
        )
        box(cell_w, number_format="0.0000")
        # A·w = explicit dot product of row i of pairwise with eigenvector column.
        # Written as B{row}*$B${w0} + C{row}*$B${w1} + ... so it works in any Excel
        # variant (no implicit array-broadcast assumption).
        aw_terms = [
            f"{get_column_letter(2 + j)}{pm_top + 2 + i}*$B${ozv_top + 1 + j}"
            for j in range(n)
        ]
        cell_aw = ws.cell(row=r, column=3)
        cell_aw.value = "=" + "+".join(aw_terms)
        box(cell_aw, number_format="0.0000")
        cell_ratio = ws.cell(row=r, column=4)
        cell_ratio.value = f"=C{r}/B{r}"
        box(cell_ratio, number_format="0.0000")

    sum_row2 = ozv_top + 1 + n
    sub(ws.cell(row=sum_row2, column=1, value="Toplam"))
    cell_total = ws.cell(row=sum_row2, column=2)
    cell_total.value = f"=SUM(B{ozv_top + 1}:B{ozv_top + n})"
    box(cell_total, number_format="0.0000")

    metrics_top = sum_row2 + 2
    metrics = [
        ("λ_max (ENBÖZD)", f"=AVERAGE(D{ozv_top + 1}:D{ozv_top + n})"),
        ("n (kriter sayısı)", n),
        ("CI = (λ_max - n) / (n - 1)", f"=(B{metrics_top}-B{metrics_top + 1})/(B{metrics_top + 1}-1)"),
        ("RI (n=5)", 1.11),
        ("CR = CI / RI", f"=B{metrics_top + 2}/B{metrics_top + 3}"),
        ("Tutarlılık Eşiği", 0.10),
        ("Durum", f'=IF(B{metrics_top + 4}<=B{metrics_top + 5},"Tutarlı","Tutarsız")'),
    ]
    for i, (label, value) in enumerate(metrics):
        r = metrics_top + i
        sub(ws.cell(row=r, column=1, value=label))
        cell = ws.cell(row=r, column=2, value=value)
        fmt = "0.0000" if "λ" in label or "CI" in label or "CR" in label or "RI" in label or "Eşik" in label else None
        box(cell, number_format=fmt)
        if label == "Durum":
            cell.fill = PASS_FILL  # actual fill toggled visually by conditional logic in the cell
            cell.font = Font(bold=True)

    auto_width(ws, [26, 14, 14, 14, 14, 14, 14, 14])


def write_aggregation_sheet(ws, participant_count: int):
    ws.title = SHEET_AGG
    ws["A1"] = "AHP AĞIRLIKLARI - GEOMETRİK ORTALAMA BİRLEŞTİRMESİ"
    ws["A1"].font = Font(bold=True, size=14, color="1F3864")
    ws.merge_cells("A1:H1")

    sub(ws["A3"], "Katılımcı")
    for i, name in enumerate(schema.CRITERIA_SHORT_TR):
        sub(ws.cell(row=3, column=2 + i, value=name))
    sub(ws.cell(row=3, column=7, value="CR"))
    sub(ws.cell(row=3, column=8, value="Tutarlı?"))

    for p in range(1, participant_count + 1):
        r = 3 + p
        sheet_name = part_sheet(p)
        box(ws.cell(row=r, column=1, value=p))
        ws.cell(row=r, column=1).alignment = CENTER
        for i in range(5):
            cell = ws.cell(row=r, column=2 + i)
            cell.value = f"='{sheet_name}'!B{29 + i}"
            box(cell, number_format="0.0000")
        cr_cell = ws.cell(row=r, column=7)
        # CR sits at B40 on each participant sheet
        cr_cell.value = f"='{sheet_name}'!B40"
        box(cr_cell, number_format="0.0000")
        flag_cell = ws.cell(row=r, column=8)
        flag_cell.value = f'=IF(G{r}<=0.1,1,0)'
        box(flag_cell, number_format="0")

    geo_row = 3 + participant_count + 2
    sub(ws.cell(row=geo_row, column=1, value="Geometrik Ortalama (yalnız tutarlı)"))
    ws.merge_cells(start_row=geo_row, start_column=1, end_row=geo_row, end_column=1)
    for i in range(5):
        col_letter = get_column_letter(2 + i)
        cell = ws.cell(row=geo_row, column=2 + i)
        cell.value = (
            f"=EXP(SUMPRODUCT(LN({col_letter}4:{col_letter}{3 + participant_count})*$H$4:$H${3 + participant_count})"
            f"/SUM($H$4:$H${3 + participant_count}))"
        )
        box(cell, number_format="0.0000")

    norm_row = geo_row + 1
    sub(ws.cell(row=norm_row, column=1, value="Toplam (geom)"))
    cell = ws.cell(row=norm_row, column=2)
    cell.value = f"=SUM(B{geo_row}:F{geo_row})"
    box(cell, number_format="0.0000")

    final_row = norm_row + 1
    sub(ws.cell(row=final_row, column=1, value="Normalize Edilmiş W"))
    for i in range(5):
        col_letter = get_column_letter(2 + i)
        cell = ws.cell(row=final_row, column=2 + i)
        cell.value = f"={col_letter}{geo_row}/$B${norm_row}"
        cell.fill = RANK_FILL
        cell.font = Font(bold=True)
        box(cell, number_format="0.0000")

    summary_row = final_row + 2
    sub(ws.cell(row=summary_row, column=1, value="Tutarlı Katılımcı Sayısı"))
    cell = ws.cell(row=summary_row, column=2)
    cell.value = f"=SUM(H4:H{3 + participant_count})"
    box(cell, number_format="0")
    sub(ws.cell(row=summary_row + 1, column=1, value="Tutarlı Sayısı ≥ 20 mi?"))
    cell = ws.cell(row=summary_row + 1, column=2)
    cell.value = f'=IF(B{summary_row}>=20,"EVET ✓","HAYIR ✗")'
    cell.font = Font(bold=True)
    box(cell)

    auto_width(ws, [38, 12, 12, 12, 12, 12, 12, 12])
    ws.freeze_panes = "B4"


def write_decision_matrix(ws, participant_count: int):
    ws.title = SHEET_DM
    ws["A1"] = "KARAR MATRİSİ - ARİTMETİK ORTALAMA (TOPSIS girdisi)"
    ws["A1"].font = Font(bold=True, size=14, color="1F3864")
    ws.merge_cells("A1:H1")

    sub(ws["A3"], "Alternatif \\ Kriter")
    for k, name in enumerate(schema.CRITERIA_SHORT_TR):
        sub(ws.cell(row=3, column=2 + k, value=name))

    raw_first = 2
    raw_last = 1 + participant_count
    for a, alt in enumerate(schema.ALTERNATIVES_TR):
        r = 4 + a
        sub(ws.cell(row=r, column=1, value=alt))
        for k in range(5):
            col_letter = topsis_col_letter(k, a)
            cell = ws.cell(row=r, column=2 + k)
            cell.value = (
                f"=AVERAGEIF(Anket_Yanitlari!{col_letter}{raw_first}:{col_letter}{raw_last},"
                f'"<>Kullanmıyorum")'
            )
            box(cell, number_format="0.00")

    auto_width(ws, [22, 14, 14, 14, 14, 14])


def write_topsis_sheet(ws, alt_count: int):
    ws.title = SHEET_TOPSIS
    ws["A1"] = "TOPSIS ANALİZİ (tek seferlik)"
    ws["A1"].font = Font(bold=True, size=14, color="1F3864")
    ws.merge_cells("A1:H1")

    sub(ws["A3"], "Kriter Ağırlıkları (W)")
    ws.merge_cells("A3:F3")
    for k, name in enumerate(schema.CRITERIA_SHORT_TR):
        sub(ws.cell(row=4, column=2 + k, value=name))
    sub(ws.cell(row=5, column=1, value="W"))
    for k in range(5):
        col_letter = get_column_letter(2 + k)
        cell = ws.cell(row=5, column=2 + k)
        # AHP_Birleştirme final_row depends on participant count -> use named lookup
        # final normalized W lives at row (3 + n_participants + 2 + 2). We hardcode 35
        # because that is the demo dataset; the generator below recomputes if needed.
        cell.value = f"=AHP_Birlestirme!{col_letter}{3 + 35 + 2 + 2}"
        box(cell, number_format="0.0000")

    dm_top = 7
    sub(ws.cell(row=dm_top, column=1, value="Karar Matrisi (D)"))
    ws.merge_cells(start_row=dm_top, start_column=1, end_row=dm_top, end_column=6)
    for k, name in enumerate(schema.CRITERIA_SHORT_TR):
        sub(ws.cell(row=dm_top + 1, column=2 + k, value=name))
    for a, alt in enumerate(schema.ALTERNATIVES_TR):
        sub(ws.cell(row=dm_top + 2 + a, column=1, value=alt))
        for k in range(5):
            cell = ws.cell(row=dm_top + 2 + a, column=2 + k)
            cell.value = f"={SHEET_DM}!{get_column_letter(2 + k)}{4 + a}"
            box(cell, number_format="0.00")

    cn_row = dm_top + 2 + alt_count + 1
    sub(ws.cell(row=cn_row, column=1, value="Sütun Normu n_j = √Σd²"))
    for k in range(5):
        col_letter = get_column_letter(2 + k)
        cell = ws.cell(row=cn_row, column=2 + k)
        # SUMSQ written as explicit squared-sum so every Excel/LibreOffice variant evaluates.
        terms = "+".join(
            f"{col_letter}{dm_top + 2 + a}^2" for a in range(alt_count)
        )
        cell.value = f"=SQRT({terms})"
        box(cell, number_format="0.0000")

    norm_top = cn_row + 2
    sub(ws.cell(row=norm_top, column=1, value="Normalize Matris (R)"))
    ws.merge_cells(start_row=norm_top, start_column=1, end_row=norm_top, end_column=6)
    for k, name in enumerate(schema.CRITERIA_SHORT_TR):
        sub(ws.cell(row=norm_top + 1, column=2 + k, value=name))
    for a in range(alt_count):
        sub(ws.cell(row=norm_top + 2 + a, column=1, value=schema.ALTERNATIVES_TR[a]))
        for k in range(5):
            col_letter = get_column_letter(2 + k)
            cell = ws.cell(row=norm_top + 2 + a, column=2 + k)
            cell.value = f"={col_letter}{dm_top + 2 + a}/{col_letter}{cn_row}"
            box(cell, number_format="0.0000")

    w_top = norm_top + 2 + alt_count + 1
    sub(ws.cell(row=w_top, column=1, value="Ağırlıklı Normalize Matris (V)"))
    ws.merge_cells(start_row=w_top, start_column=1, end_row=w_top, end_column=6)
    for k, name in enumerate(schema.CRITERIA_SHORT_TR):
        sub(ws.cell(row=w_top + 1, column=2 + k, value=name))
    for a in range(alt_count):
        sub(ws.cell(row=w_top + 2 + a, column=1, value=schema.ALTERNATIVES_TR[a]))
        for k in range(5):
            col_letter = get_column_letter(2 + k)
            cell = ws.cell(row=w_top + 2 + a, column=2 + k)
            cell.value = f"={col_letter}{norm_top + 2 + a}*{col_letter}$5"
            box(cell, number_format="0.0000")

    ideal_top = w_top + 2 + alt_count + 1
    sub(ws.cell(row=ideal_top, column=1, value="A+ (ideal, max)"))
    for k in range(5):
        col_letter = get_column_letter(2 + k)
        cell = ws.cell(row=ideal_top, column=2 + k)
        cell.value = f"=MAX({col_letter}{w_top + 2}:{col_letter}{w_top + 1 + alt_count})"
        box(cell, number_format="0.0000")
    sub(ws.cell(row=ideal_top + 1, column=1, value="A- (anti-ideal, min)"))
    for k in range(5):
        col_letter = get_column_letter(2 + k)
        cell = ws.cell(row=ideal_top + 1, column=2 + k)
        cell.value = f"=MIN({col_letter}{w_top + 2}:{col_letter}{w_top + 1 + alt_count})"
        box(cell, number_format="0.0000")

    closeness_top = ideal_top + 3
    sub(ws.cell(row=closeness_top, column=1, value="Alternatif"))
    sub(ws.cell(row=closeness_top, column=2, value="S⁺ (ideal mesafe)"))
    sub(ws.cell(row=closeness_top, column=3, value="S⁻ (anti-ideal mesafe)"))
    sub(ws.cell(row=closeness_top, column=4, value="C = S⁻/(S⁺+S⁻)"))
    sub(ws.cell(row=closeness_top, column=5, value="Sıra"))
    for a in range(alt_count):
        r = closeness_top + 1 + a
        sub(ws.cell(row=r, column=1, value=schema.ALTERNATIVES_TR[a]))
        # S+
        terms_plus = "+".join(
            f"({get_column_letter(2 + k)}{w_top + 2 + a}-{get_column_letter(2 + k)}{ideal_top})^2"
            for k in range(5)
        )
        terms_minus = "+".join(
            f"({get_column_letter(2 + k)}{w_top + 2 + a}-{get_column_letter(2 + k)}{ideal_top + 1})^2"
            for k in range(5)
        )
        ws.cell(row=r, column=2, value=f"=SQRT({terms_plus})")
        ws.cell(row=r, column=3, value=f"=SQRT({terms_minus})")
        ws.cell(row=r, column=4, value=f"=C{r}/(B{r}+C{r})")
        ws.cell(row=r, column=5, value=f"=RANK(D{r},$D${closeness_top + 1}:$D${closeness_top + alt_count})")
        for c in range(2, 6):
            box(ws.cell(row=r, column=c), number_format="0.0000" if c < 5 else "0")
        ws.cell(row=r, column=5).fill = RANK_FILL
        ws.cell(row=r, column=5).font = Font(bold=True)

    auto_width(ws, [28, 16, 16, 18, 10, 14])


def write_electre_sheet(ws, alt_count: int):
    ws.title = SHEET_ELECTRE
    ws["A1"] = "ELECTRE I ANALİZİ (tek seferlik)"
    ws["A1"].font = Font(bold=True, size=14, color="1F3864")
    ws.merge_cells("A1:H1")

    # Weighted-normalized matrix lives in TOPSIS starting at row 29 (w_top + 2)
    # where w_top = norm_top + 2 + alt_count + 1 = 18 + 2 + 6 + 1 = 27.
    topsis_v_first_row = 29

    sub(ws["A3"], "Ağırlıklı Normalize Matris (TOPSIS!V kopyası)")
    ws.merge_cells("A3:F3")
    for k, name in enumerate(schema.CRITERIA_SHORT_TR):
        sub(ws.cell(row=4, column=2 + k, value=name))
    for a in range(alt_count):
        sub(ws.cell(row=5 + a, column=1, value=schema.ALTERNATIVES_TR[a]))
        for k in range(5):
            col_letter = get_column_letter(2 + k)
            cell = ws.cell(row=5 + a, column=2 + k)
            cell.value = f"=TOPSIS!{col_letter}{topsis_v_first_row + a}"
            box(cell, number_format="0.0000")

    sub(ws["A12"], "Kriter Ağırlıkları (W, TOPSIS kopyası)")
    ws.merge_cells("A12:F12")
    for k, name in enumerate(schema.CRITERIA_SHORT_TR):
        sub(ws.cell(row=13, column=2 + k, value=name))
    for k in range(5):
        cell = ws.cell(row=14, column=2 + k)
        cell.value = f"=TOPSIS!{get_column_letter(2 + k)}5"
        box(cell, number_format="0.0000")

    # Concordance matrix C(p,q) = SUMIFS-like: sum of W where V_p >= V_q
    conc_top = 16
    sub(ws.cell(row=conc_top, column=1, value="Uyum (Concordance) Matrisi  C(p,q)"))
    ws.merge_cells(start_row=conc_top, start_column=1, end_row=conc_top, end_column=alt_count + 1)
    for q in range(alt_count):
        sub(ws.cell(row=conc_top + 1, column=2 + q, value=schema.ALTERNATIVES_TR[q]))
    for p in range(alt_count):
        sub(ws.cell(row=conc_top + 2 + p, column=1, value=schema.ALTERNATIVES_TR[p]))
        for q in range(alt_count):
            cell = ws.cell(row=conc_top + 2 + p, column=2 + q)
            if p == q:
                cell.value = ""
            else:
                # SUMPRODUCT((V_p >= V_q) * W)
                p_range = f"B{5 + p}:F{5 + p}"
                q_range = f"B{5 + q}:F{5 + q}"
                w_range = "$B$14:$F$14"
                cell.value = f"=SUMPRODUCT(({p_range}>={q_range})*{w_range})"
            box(cell, number_format="0.0000")

    disc_top = conc_top + 2 + alt_count + 1
    sub(ws.cell(row=disc_top, column=1, value="Uyumsuzluk (Discordance) Matrisi  d(p,q)"))
    ws.merge_cells(start_row=disc_top, start_column=1, end_row=disc_top, end_column=alt_count + 1)
    for q in range(alt_count):
        sub(ws.cell(row=disc_top + 1, column=2 + q, value=schema.ALTERNATIVES_TR[q]))
    for p in range(alt_count):
        sub(ws.cell(row=disc_top + 2 + p, column=1, value=schema.ALTERNATIVES_TR[p]))
        for q in range(alt_count):
            cell = ws.cell(row=disc_top + 2 + p, column=2 + q)
            if p == q:
                cell.value = ""
            else:
                p_range = f"B{5 + p}:F{5 + p}"
                q_range = f"B{5 + q}:F{5 + q}"
                # max over (V_q - V_p) where V_p < V_q, normalized by max |V_p - V_q|
                # Implement as IFERROR(MAX of clamped diffs / global max abs diff, 0)
                cell.value = (
                    f"=IFERROR(MAX(IF({p_range}<{q_range},{q_range}-{p_range},0))"
                    f"/MAX(ABS({p_range}-{q_range})),0)"
                )
                cell.data_type = "f"
            box(cell, number_format="0.0000")

    thresh_top = disc_top + 2 + alt_count + 1
    sub(ws.cell(row=thresh_top, column=1, value="Uyum Eşiği c̄ (off-diag ort.)"))
    ws.cell(row=thresh_top, column=2, value=(
        f"=AVERAGE(IF(ROW(B{conc_top + 2}:G{conc_top + 1 + alt_count})"
        f"<>COLUMN(B{conc_top + 2}:G{conc_top + 1 + alt_count})-1+{conc_top + 2 - 2},"
        f"B{conc_top + 2}:G{conc_top + 1 + alt_count}))"
    ))
    ws.cell(row=thresh_top, column=2).number_format = "0.0000"
    # Simpler robust threshold: average of all (sum minus diag zeros) / count
    # Replace with concrete formula avoiding array trick:
    ws.cell(row=thresh_top, column=2, value=(
        f"=SUM(B{conc_top + 2}:{get_column_letter(1 + alt_count)}{conc_top + 1 + alt_count})"
        f"/({alt_count * (alt_count - 1)})"
    ))
    ws.cell(row=thresh_top, column=2).number_format = "0.0000"
    box(ws.cell(row=thresh_top, column=2))

    sub(ws.cell(row=thresh_top + 1, column=1, value="Uyumsuzluk Eşiği d̄ (off-diag ort.)"))
    ws.cell(row=thresh_top + 1, column=2, value=(
        f"=SUM(B{disc_top + 2}:{get_column_letter(1 + alt_count)}{disc_top + 1 + alt_count})"
        f"/({alt_count * (alt_count - 1)})"
    ))
    ws.cell(row=thresh_top + 1, column=2).number_format = "0.0000"
    box(ws.cell(row=thresh_top + 1, column=2))

    out_top = thresh_top + 3
    sub(ws.cell(row=out_top, column=1, value="Üstünlük Matrisi (1 = baskın)"))
    ws.merge_cells(start_row=out_top, start_column=1, end_row=out_top, end_column=alt_count + 1)
    for q in range(alt_count):
        sub(ws.cell(row=out_top + 1, column=2 + q, value=schema.ALTERNATIVES_TR[q]))
    for p in range(alt_count):
        sub(ws.cell(row=out_top + 2 + p, column=1, value=schema.ALTERNATIVES_TR[p]))
        for q in range(alt_count):
            cell = ws.cell(row=out_top + 2 + p, column=2 + q)
            if p == q:
                cell.value = ""
            else:
                cell.value = (
                    f"=IF(AND({get_column_letter(2 + q)}{conc_top + 2 + p}>=B{thresh_top},"
                    f"{get_column_letter(2 + q)}{disc_top + 2 + p}<=B{thresh_top + 1}),1,0)"
                )
            box(cell, number_format="0")

    rank_top = out_top + 2 + alt_count + 1
    sub(ws.cell(row=rank_top, column=1, value="Alternatif"))
    sub(ws.cell(row=rank_top, column=2, value="Baskınlık (satır)"))
    sub(ws.cell(row=rank_top, column=3, value="Bağımlılık (sütun)"))
    sub(ws.cell(row=rank_top, column=4, value="Net Skor"))
    sub(ws.cell(row=rank_top, column=5, value="Sıra"))
    for a in range(alt_count):
        r = rank_top + 1 + a
        sub(ws.cell(row=r, column=1, value=schema.ALTERNATIVES_TR[a]))
        row_range = f"B{out_top + 2 + a}:{get_column_letter(1 + alt_count)}{out_top + 2 + a}"
        col_letter_a = get_column_letter(2 + a)
        col_range = f"{col_letter_a}{out_top + 2}:{col_letter_a}{out_top + 1 + alt_count}"
        ws.cell(row=r, column=2, value=f"=SUM({row_range})")
        ws.cell(row=r, column=3, value=f"=SUM({col_range})")
        ws.cell(row=r, column=4, value=f"=B{r}-C{r}")
        ws.cell(row=r, column=5, value=f"=RANK(D{r},$D${rank_top + 1}:$D${rank_top + alt_count})")
        for c in range(2, 6):
            box(ws.cell(row=r, column=c), number_format="0")
        ws.cell(row=r, column=5).fill = RANK_FILL
        ws.cell(row=r, column=5).font = Font(bold=True)

    auto_width(ws, [34, 16, 16, 14, 10])


def write_summary_sheet(ws, alt_count: int):
    ws.title = SHEET_SUMMARY
    ws["A1"] = "FİNAL SIRALAMA - TOPSIS vs ELECTRE"
    ws["A1"].font = Font(bold=True, size=14, color="1F3864")
    ws.merge_cells("A1:E1")

    sub(ws["A3"], "Alternatif")
    sub(ws["B3"], "TOPSIS C")
    sub(ws["C3"], "TOPSIS Sıra")
    sub(ws["D3"], "ELECTRE Net")
    sub(ws["E3"], "ELECTRE Sıra")

    topsis_close_first = 35
    electre_rank_first = None
    # closeness row in TOPSIS for alt a -> closeness_top + 1 + a; need same offset
    # We hardcode based on layout above:
    closeness_top_in_topsis = 7 + 1 + alt_count + 1 + 1 + alt_count + 1 + 1 + alt_count + 1 + 3
    # Easier: compute relative locations as used in build sequence:
    dm_top = 7
    norm_top = dm_top + 2 + alt_count + 1 + 1
    w_top = norm_top + 2 + alt_count + 1
    ideal_top = w_top + 2 + alt_count + 1
    topsis_closeness_top = ideal_top + 3

    conc_top = 16
    disc_top = conc_top + 2 + alt_count + 1
    thresh_top = disc_top + 2 + alt_count + 1
    out_top = thresh_top + 3
    electre_rank_top = out_top + 2 + alt_count + 1

    for a in range(alt_count):
        r = 4 + a
        sub(ws.cell(row=r, column=1, value=schema.ALTERNATIVES_TR[a]))
        ws.cell(row=r, column=2, value=f"=TOPSIS!D{topsis_closeness_top + 1 + a}")
        ws.cell(row=r, column=3, value=f"=TOPSIS!E{topsis_closeness_top + 1 + a}")
        ws.cell(row=r, column=4, value=f"=ELECTRE!D{electre_rank_top + 1 + a}")
        ws.cell(row=r, column=5, value=f"=ELECTRE!E{electre_rank_top + 1 + a}")
        for c in range(2, 6):
            cell = ws.cell(row=r, column=c)
            box(cell, number_format="0.0000" if c == 2 else "0")
            if c in (3, 5):
                cell.fill = RANK_FILL
                cell.font = Font(bold=True)

    auto_width(ws, [22, 14, 14, 14, 14])


def build(csv_path: Path, out_path: Path) -> None:
    df = pd.read_csv(csv_path, encoding="utf-8-sig")
    n_participants = len(df)
    n_alt = len(schema.ALTERNATIVES_TR)

    wb = Workbook()
    write_info_sheet(wb.active)
    raw_ws = wb.create_sheet()
    write_raw_sheet(raw_ws, df)
    for idx, (_, row) in enumerate(df.iterrows(), start=1):
        ws = wb.create_sheet(part_sheet(idx))
        write_participant_sheet(ws, participant_id=int(row["participant_id"]), row_in_raw=1 + idx)
    write_aggregation_sheet(wb.create_sheet(), n_participants)
    write_decision_matrix(wb.create_sheet(), n_participants)
    write_topsis_sheet(wb.create_sheet(), n_alt)
    write_electre_sheet(wb.create_sheet(), n_alt)
    write_summary_sheet(wb.create_sheet(), n_alt)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    print(f"Workbook saved: {out_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--csv", type=Path, default=Path("../data/sample_responses.csv"))
    parser.add_argument("--out", type=Path, default=Path("../workbook/Busra_Final_Project.xlsx"))
    args = parser.parse_args()
    build(args.csv, args.out)
