# Python Kaynak Kodu — Opsiyonel

Bu klasör **opsiyoneldir**. Excel ve PowerPoint dosyalarını üretmek için kullanılan Python kodunu içeriyor. Hocanın istediği sadece Excel + PowerPoint. Yine de "Python ile NumPy kullanarak da hesapladım" diye sunumda söylemek istersen, kod buradadır.

## Dosyalar

| Dosya | Ne yapıyor |
| --- | --- |
| `schema.py` | Anket sütun düzeni, kriter ve alternatif listeleri (Google Form ile birebir). |
| `ahp.py` | AHP hesabı: ikili karşılaştırma matrisi → ÖZV → λ_max → CI → CR. |
| `topsis.py` | TOPSIS algoritması (normalize → ağırlıklı → ideal mesafe → yakınlık katsayısı). |
| `electre.py` | ELECTRE I (uyum/uyumsuzluk → üstünlük matrisi → net skor). |
| `generate_sample_data.py` | 35 sentetik katılımcı veri üretir (demo). |
| `make_template.py` | Gerçek anket için boş CSV şablonu üretir. |
| `pipeline.py` | **Ana orkestratör.** CSV oku → tüm hesapları yap → JSON çıktısı + konsol özeti. |
| `build_workbook.py` | Formül dolu Excel teslim dosyasını üretir. |
| `build_presentation.py` | PowerPoint sunumunu üretir. |
| `verify_workbook.py` | Excel formüllerinin Python pipeline sonuçlarıyla eşleştiğini doğrular. |

## Kendi başına çalıştırmak istersen

**1. Gerekli paketler:**
```
pip install numpy pandas openpyxl python-pptx
```

**2. Pipeline'ı çalıştır:**
```
python pipeline.py --csv ../03_ANKET_VERISI/01_demo_yanitlar_35_kisi.csv --out ../output
```

**3. Excel ve sunumu yeniden üret:**
```
python build_workbook.py --csv ../03_ANKET_VERISI/01_demo_yanitlar_35_kisi.csv --out ../01_TESLIM/Busra_Final_Project.xlsx
python build_presentation.py --results ../output/results.json --out ../01_TESLIM/Busra_Final_Project.pptx
```

**4. (Opsiyonel) Doğrulama:**
```
py -3.12 verify_workbook.py
```

> ⚠️ Python 3.12 gerekli (verify için). Diğer betikler Python 3.10+ ile çalışır.

## Gerçek anket sonuçların geldiğinde

1. CSV'ni `../03_ANKET_VERISI/gercek_yanitlar.csv` adıyla kaydet (sütun düzeni `02_BOS_SABLON...csv` ile aynı olmalı).
2. Yukarıdaki 3 komutu çalıştır (sadece `--csv` parametresini kendi dosyana çevir).
3. `01_TESLIM/` klasöründeki Excel ve PowerPoint güncellenmiş halde hazır olur.

## Kodun mantığı

Akış:

```
CSV (anket cevapları)
   ↓
pipeline.py
   ├─ Her katılımcı için AHP (ahp.py)
   ├─ CR ≤ 0.10 olanları filtrele
   ├─ Geometrik ortalama → kriter ağırlıkları
   ├─ Aritmetik ortalama → karar matrisi
   ├─ TOPSIS bir kez (topsis.py)
   └─ ELECTRE bir kez (electre.py)
        ↓
   results.json (tüm sayılar)
        ↓
   ┌────┴────┐
   ↓         ↓
xlsx       pptx
```

Tüm hesaplar **NumPy** ile yapılıyor (hocanın ders notlarıyla aynı şekilde — `05_Ders_Notu_AHP.pdf` ve `06_Ders_Notu_TOPSIS.pdf`).
