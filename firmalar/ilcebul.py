import json
import pandas as pd
import re


INPUT_FILE = "firmalar.xlsx"
OUTPUT_FILE = "firmalar.xlsx"
ILCELER_FILE = "ilceler.json"


# ============================================================
# YARDIMCI FONKSİYONLAR
# ============================================================

def normalize(text):
    text = str(text).upper().strip()

    replacements = {
        "Ç": "C",
        "Ğ": "G",
        "İ": "I",
        "Ö": "O",
        "Ş": "S",
        "Ü": "U",
    }

    for eski, yeni in replacements.items():
        text = text.replace(eski, yeni)

    return text


def tr_upper(text):
    return text.replace("i", "İ").replace("ı", "I").upper()


# ============================================================
# ŞEHİR LİSTESİ
# ============================================================

with open(ILCELER_FILE, "r", encoding="utf-8") as f:
    SEHIRLER = json.load(f)


SEHIRLER_NORMALIZED = {
    normalize(sehir): sehir
    for sehir in SEHIRLER
}


# ============================================================
# ŞEHİR SEÇİMİ
# ============================================================

print("=" * 60)
print("İLÇE BULMA")
print("=" * 60)
print("Şehir adını girin.")
print("Örnek: istanbul")
print("=" * 60)


while True:
    sehir_girisi = input("Şehir adı: ").strip()
    sehir_normalized = normalize(sehir_girisi)

    if sehir_normalized in SEHIRLER_NORMALIZED:
        sehir = SEHIRLER_NORMALIZED[sehir_normalized]
        break

    print()
    print(f"'{sehir_girisi}' bulunamadı.")
    print("Lütfen geçerli bir şehir adı girin.")
    print()


ILCELER = SEHIRLER[sehir]


print()
print(f"Seçilen şehir: {sehir}")
print(f"İlçe sayısı  : {len(ILCELER)}")
print("=" * 60)


# ============================================================
# İLÇE LİSTESİ
# ============================================================

# Uzun ilçe isimleri önce kontrol edilir.
# Böylece daha kısa isimlerin yanlış eşleşme ihtimali azalır.
ILCELER = sorted(
    set(ILCELER),
    key=lambda x: len(normalize(x)),
    reverse=True
)


ILCELER_NORMALIZED = {
    normalize(ilce): tr_upper(ilce)
    for ilce in ILCELER
}


# ============================================================
# ADRESTEN İLÇE BULMA
# ============================================================

def ilce_bul(adres):
    if pd.isna(adres):
        return ""

    adres = str(adres).strip()

    if not adres:
        return ""

    adres_normalized = normalize(adres)

    for ilce_normalized, ilce_original in ILCELER_NORMALIZED.items():

        pattern = rf"(?<![A-Z]){re.escape(ilce_normalized)}(?![A-Z])"

        if re.search(pattern, adres_normalized):
            return ilce_original

    return ""


# ============================================================
# EXCEL OKUMA
# ============================================================

df = pd.read_excel(INPUT_FILE)

# Sütun isimlerindeki gereksiz boşlukları temizle
df.columns = df.columns.str.strip()


# ILCE sütununu bul
ilce_sutunu = None

for sutun in df.columns:
    sutun_normalized = normalize(sutun)

    if sutun_normalized == "ILCE":
        ilce_sutunu = sutun
        break


# ILCE sütunu yoksa oluştur
if ilce_sutunu is None:
    df["ILCE"] = ""
    ilce_sutunu = "ILCE"


# ============================================================
# İLÇE BİLGİSİNİ BUL
# ============================================================

df[ilce_sutunu] = df["ADRES"].apply(ilce_bul)


# ============================================================
# EXCEL KAYDET
# ============================================================

df.to_excel(OUTPUT_FILE, index=False)


# ============================================================
# SONUÇ
# ============================================================

toplam = len(df)
bulunan = (df[ilce_sutunu] != "").sum()
bulunamayan = toplam - bulunan

print()
print("=" * 60)
print("TAMAMLANDI")
print("=" * 60)
print(f"Şehir            : {sehir}")
print(f"Toplam kayıt     : {toplam}")
print(f"İlçe bulunan     : {bulunan}")
print(f"İlçe bulunamayan : {bulunamayan}")

if toplam > 0:
    print(f"Başarı oranı     : %{bulunan / toplam * 100:.1f}")

print(f"Dosya            : {OUTPUT_FILE}")
print("=" * 60)
