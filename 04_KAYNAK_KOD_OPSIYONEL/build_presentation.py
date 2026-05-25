"""Build the PowerPoint deliverable.

Slides:
    1.  Cover  - project title, student info, date.
    2.  Amaç   - objective + research question.
    3.  Yöntem - AHP + TOPSIS + ELECTRE one-line definitions.
    4.  Kurallar - 8 hard rules from the assignment brief.
    5.  Anket  - 6 alternatives + 5 criteria + sample size.
    6.  AHP işleyişi - per-participant flow + RI table.
    7.  Tutarlılık - threshold + how many made the cut.
    8.  AHP birleştirme - geometric-mean formula + final W vector.
    9.  Karar matrisi - arithmetic mean + decision matrix.
    10. TOPSIS - 7 steps + result table + ranking chart.
    11. ELECTRE - concordance/discordance + ranking chart.
    12. Karşılaştırma - TOPSIS vs ELECTRE side-by-side.
    13. Sonuç - kazanan + neden + sınırlamalar.
    14. Kaynakça - lecture PDFs + assignment briefs.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt, Emu

import schema

NAVY = "1F3864"
BLUE = "305496"
LIGHT = "D9E1F2"
GOLD = "F2C744"


def add_title_slide(prs, title, subtitle):
    layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(layout)
    slide.shapes.title.text = title
    if len(slide.placeholders) > 1:
        slide.placeholders[1].text = subtitle
    return slide


def add_content_slide(prs, title):
    layout = prs.slide_layouts[5]  # blank with title
    slide = prs.slides.add_slide(layout)
    slide.shapes.title.text = title
    return slide


def add_bullets(slide, items, left=Inches(0.5), top=Inches(1.5), width=Inches(9), height=Inches(5.5),
                font_size=18, bold_first=False):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        for run in p.runs:
            run.font.size = Pt(font_size)
            if bold_first and i == 0:
                run.font.bold = True
        p.space_after = Pt(8)
    return box


def add_table(slide, data, left=Inches(0.5), top=Inches(1.5), width=Inches(9), height=Inches(4),
              header_row=True):
    rows = len(data)
    cols = len(data[0])
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    for r, row in enumerate(data):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = str(val)
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(13)
                    if header_row and r == 0:
                        run.font.bold = True
    return table


def add_bar_chart(slide, title, categories, values, left=Inches(0.5), top=Inches(1.6),
                  width=Inches(9), height=Inches(5)):
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series(title, values)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED, left, top, width, height, chart_data
    ).chart
    chart.has_legend = False
    chart.has_title = True
    chart.chart_title.text_frame.text = title
    return chart


def build(results_path: Path, out_path: Path) -> None:
    results = json.loads(results_path.read_text(encoding="utf-8"))

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    add_title_slide(
        prs,
        "AHP + TOPSIS + ELECTRE ile\nE-Ticaret Sitelerinin Performans Değerlendirmesi",
        "Büşra Nur Yıldırım  •  230326759  •  Kırmızı Grup\n"
        "YBS302 Karar Destek Sistemleri  •  Dr. Öğr. Üyesi Neda Alipour\n"
        "İstanbul Gelişim Üniversitesi  •  Bahar 2025-2026"
    )

    s = add_content_slide(prs, "Amaç ve Araştırma Sorusu")
    add_bullets(s, [
        "Amaç: Türkiye'de yaygın kullanılan 6 e-ticaret sitesini, kullanıcıların algıladığı performansa göre objektif olarak sıralamak.",
        "Araştırma sorusu: Hangi e-ticaret sitesi, kullanıcı tarafından belirlenen ağırlıklı kriterlere göre en yüksek performansı sunmaktadır?",
        "Yöntem: Klasik çok kriterli karar verme (MCDM) — AHP ile kriter ağırlıkları, TOPSIS ve ELECTRE ile alternatif sıralaması.",
        "Veri kaynağı: Google Forms üzerinden toplanmış anonim katılımcı yanıtları.",
    ])

    s = add_content_slide(prs, "Yöntem Özeti")
    add_bullets(s, [
        "AHP (Analytic Hierarchy Process): Her katılımcının ikili karşılaştırmalarından kriter ağırlıkları üretir.",
        "Tutarlılık Oranı (CR ≤ 0.10): Mantıksız karşılaştırma yapan katılımcılar filtrelenir.",
        "Geometrik ortalama: Tutarlı katılımcıların ağırlıkları tek bir nihai ağırlık vektörüne birleştirilir.",
        "Aritmetik ortalama: Alternatif performans puanları tek bir karar matrisine birleştirilir.",
        "TOPSIS: Aday alternatifin pozitif ve negatif ideal çözümlere göre yakınlık katsayısı (Cᵢ).",
        "ELECTRE I: Uyum / uyumsuzluk endeksleriyle baskınlık (outranking) ilişkileri.",
    ])

    s = add_content_slide(prs, "Görevin Sıkı Kuralları (Final.docx)")
    add_bullets(s, [
        "1. Anket soru sırası değiştirilemez.",
        "2. En az 25 katılımcı (35+ için +5 bonus).",
        "3. Her katılımcı için ayrı AHP analizi gerekli.",
        "4. En az 20 katılımcı tutarlı olmalı (CR ≤ 0.10).",
        "5. AHP ağırlıkları GEOMETRİK ortalama ile birleştirilir.",
        "6. Alternatif puanları ARİTMETİK ortalama ile birleştirilir.",
        "7. TOPSIS ve ELECTRE birer kez çalıştırılır — toplam veri üzerinden.",
        "8. Tüm kriterler FAYDA tipindedir (yüksek = iyi).",
        "9. Tüm ara değerler Excel formülü olmalı, sabit sayı değil.",
    ], font_size=16)

    s = add_content_slide(prs, "Anket Tasarımı")
    table_data = [
        ["Bileşen", "Sayı / Detay"],
        ["Demografik soru", "6 (cinsiyet, yaş, eğitim, meslek, sıklık, cihaz)"],
        ["Kriter (hepsi FAYDA)", "5"],
        ["İkili karşılaştırma", "10 = C(5,2)"],
        ["Alternatif", "6 e-ticaret sitesi"],
        ["TOPSIS hücresi", "5 kriter × 6 alternatif = 30 puan"],
        ["Toplam katılımcı (demo)", f"{results['n_participants']}"],
        ["Tutarlı katılımcı (demo)", f"{results['n_consistent']}  (eşik: ≥ 20)"],
    ]
    add_table(s, table_data, top=Inches(1.4), height=Inches(4.5))
    add_bullets(s, [
        "Kriterler: " + "  •  ".join(schema.CRITERIA_SHORT_TR),
        "Alternatifler: " + "  •  ".join(schema.ALTERNATIVES_TR),
    ], top=Inches(6.0), height=Inches(1.3), font_size=14)

    s = add_content_slide(prs, "AHP — Her Katılımcı için İşleyiş")
    add_bullets(s, [
        "1. Saaty 1-9 skalasıyla 5×5 ikili karşılaştırma matrisi (A).",
        "2. Sütun toplamlarına bölme → normalize matris (N).",
        "3. Satır ortalamaları → Öz Vektör w (kriter ağırlıkları).",
        "4. Aw / w değerlerinin ortalaması → λ_max (ENBÖZD).",
        "5. CI = (λ_max − n) / (n − 1)  ;   n = 5.",
        "6. CR = CI / RI;  n=5 için RI = 1.11.",
        "7. CR ≤ 0.10 → katılımcı 'tutarlı' kabul edilir.",
    ])

    s = add_content_slide(prs, f"Tutarlılık Sonucu — {results['n_consistent']} / {results['n_participants']}")
    add_bullets(s, [
        f"Toplam katılımcı: {results['n_participants']}",
        f"Tutarlı (CR ≤ 0.10): {results['n_consistent']}",
        f"Asgari gereksinim: 20 — {'KARŞILANDI ✓' if results['n_consistent'] >= 20 else 'KARŞILANMADI ✗'}",
        f"Bonus (35+ katılımcı): {'KAZANILDI ✓' if results['n_participants'] >= 35 else 'kazanılmadı'}",
    ], font_size=20)
    add_bar_chart(
        s, "Katılımcı bazında CR (örnek)",
        [f"K{p['participant_id']}" for p in results["per_participant"][:15]],
        [round(p["cr"], 3) for p in results["per_participant"][:15]],
        top=Inches(3.5), height=Inches(3.5)
    )

    s = add_content_slide(prs, "AHP Birleştirme — Geometrik Ortalama")
    add_bullets(s, [
        "g_k = (∏ wₖ⁽ⁱ⁾)^(1/m)   (m: tutarlı katılımcı sayısı)",
        "Wₖ = g_k / Σ g_j  →  toplam 1 olacak şekilde normalize.",
    ], font_size=20, top=Inches(1.4))
    table_data = [["Kriter", "Geometrik Ort. (norm.) Wₖ"]]
    for name, w in zip(schema.CRITERIA_TR, results["final_weights"]):
        table_data.append([name, f"{w:.4f}"])
    add_table(s, table_data, top=Inches(3.2), height=Inches(3.5))

    s = add_content_slide(prs, "Karar Matrisi — Aritmetik Ortalama")
    add_bullets(s, [
        "Her (alternatif, kriter) hücresi için: tüm katılımcıların verdiği 0-100 puanın aritmetik ortalaması.",
        "‘Kullanmıyorum' yanıtları ortalamaya dahil edilmez (kayıp veri).",
    ], font_size=16, top=Inches(1.4), height=Inches(1.5))
    table_data = [["Alternatif"] + schema.CRITERIA_SHORT_TR]
    for a, alt in enumerate(schema.ALTERNATIVES_TR):
        table_data.append([alt] + [f"{results['decision_matrix'][a][k]:.1f}" for k in range(5)])
    add_table(s, table_data, top=Inches(3.0), height=Inches(4.0))

    s = add_content_slide(prs, "TOPSIS — 7 Adım")
    add_bullets(s, [
        "1. Sütun normu: n_j = √Σdᵢⱼ²",
        "2. Normalize matris: rᵢⱼ = dᵢⱼ / n_j",
        "3. Ağırlıklı normalize: vᵢⱼ = Wⱼ · rᵢⱼ",
        "4. İdeal çözümler (hepsi fayda): A⁺ = sütun max, A⁻ = sütun min",
        "5. Mesafeler: S⁺ᵢ = √Σⱼ (vᵢⱼ − A⁺ⱼ)²  ;  S⁻ᵢ benzer şekilde",
        "6. Yakınlık katsayısı: Cᵢ = S⁻ᵢ / (S⁺ᵢ + S⁻ᵢ)",
        "7. Sıralama: Cᵢ azalan",
    ], font_size=18)

    s = add_content_slide(prs, "TOPSIS Sonuçları")
    cats = [schema.ALTERNATIVES_TR[i] for i in results["topsis"]["ranking_idx"]]
    vals = [round(results["topsis"]["closeness"][i], 4) for i in results["topsis"]["ranking_idx"]]
    add_bar_chart(s, "Yakınlık Katsayısı C", cats, vals)

    s = add_content_slide(prs, "ELECTRE I — Yöntem")
    add_bullets(s, [
        "TOPSIS'in ürettiği ağırlıklı normalize matris V üzerinden çalışır.",
        "Her (Aₚ, A_q) için: Uyum C(p,q) = Σ_{j: vₚⱼ≥v_qⱼ} Wⱼ",
        "Uyumsuzluk d(p,q) = max_{j: vₚⱼ<v_qⱼ} |vₚⱼ−v_qⱼ| / max_j |vₚⱼ−v_qⱼ|",
        "Eşikler: c̄ = ortalama C,  d̄ = ortalama d (köşegen hariç).",
        "Üstünlük (outranking): Aₚ ➜ A_q  ⟺  C(p,q) ≥ c̄  AND  d(p,q) ≤ d̄",
        "Net skor = satır toplamı − sütun toplamı  →  azalan sıra.",
    ], font_size=16)

    s = add_content_slide(prs, "ELECTRE Sonuçları")
    cats = [schema.ALTERNATIVES_TR[i] for i in results["electre"]["ranking_idx"]]
    vals = [int(results["electre"]["net_score"][i]) for i in results["electre"]["ranking_idx"]]
    add_bar_chart(s, "ELECTRE Net Skor", cats, vals)

    s = add_content_slide(prs, "Karşılaştırma — TOPSIS ve ELECTRE")
    table_data = [["Sıra", "TOPSIS", "C", "ELECTRE", "Net"]]
    for rank in range(6):
        t_idx = results["topsis"]["ranking_idx"][rank]
        e_idx = results["electre"]["ranking_idx"][rank]
        table_data.append([
            f"{rank+1}",
            schema.ALTERNATIVES_TR[t_idx],
            f"{results['topsis']['closeness'][t_idx]:.4f}",
            schema.ALTERNATIVES_TR[e_idx],
            f"{results['electre']['net_score'][e_idx]:+d}",
        ])
    add_table(s, table_data, top=Inches(1.4), height=Inches(4.5))

    s = add_content_slide(prs, "Sonuç ve Yorum")
    winner = results["topsis"]["ranking_named"][0]
    runnerup = results["topsis"]["ranking_named"][1]
    add_bullets(s, [
        f"En yüksek performans (TOPSIS): {winner}  (C = {results['topsis']['closeness'][results['topsis']['ranking_idx'][0]]:.4f})",
        f"İkinci (TOPSIS): {runnerup}",
        f"En düşük performans (TOPSIS): {results['topsis']['ranking_named'][-1]}",
        "ELECTRE genel olarak TOPSIS ile uyumlu bir sıralama üretmiştir.",
        "Sınırlamalar: Anket örneklemi gönüllülerden oluşmaktadır; demografik çeşitlilik daha fazla katılımcı ile artırılabilir.",
        "Gelecek çalışma: Daha fazla katılımcı, kriterlere alt-kriter ekleme, AHP'nin yerine ANP kullanma.",
    ])

    s = add_content_slide(prs, "Kaynaklar")
    add_bullets(s, [
        "Saaty, T.L. (1980). The Analytic Hierarchy Process. McGraw-Hill.",
        "Hwang, C.L. & Yoon, K. (1981). Multiple Attribute Decision Making: Methods and Applications. Springer.",
        "Roy, B. (1968). Classement et choix en présence de points de vue multiples (ELECTRE).",
        "Ders materyali: Numpy_1.pdf, Numpy-AHP.pdf, Numpy_ile_TOPSIS.pdf (Dr. Alipour).",
        "Görev brifi: Final.docx, anket_sirasi.docx (YBS302, İstanbul Gelişim Ü.).",
    ], font_size=16)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)
    print(f"Presentation saved: {out_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, default=Path("../output/results.json"))
    parser.add_argument("--out", type=Path, default=Path("../presentation/Busra_Final_Project.pptx"))
    args = parser.parse_args()
    build(args.results, args.out)
