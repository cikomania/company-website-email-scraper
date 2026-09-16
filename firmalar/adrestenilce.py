import json
import re
import sys
import unicodedata

import pandas as pd


# ============================================================
# AYARLAR
# ============================================================

GIRDI_DOSYASI = "firmalar.xlsx"
JSON_DOSYASI = "istanbul_ilce_verileri.json"
CIKTI_DOSYASI = "firmalar_ilce.xlsx"


# ============================================================
# NORMALİZASYON
# ============================================================

def normalize(metin):

    if pd.isna(metin):
        return ""

    metin = str(metin).upper().strip()

    ceviri = str.maketrans({
        "Ç": "C",
        "Ğ": "G",
        "İ": "I",
        "Ö": "O",
        "Ş": "S",
        "Ü": "U",
    })

    metin = metin.translate(ceviri)

    metin = unicodedata.normalize(
        "NFKD",
        metin
    )

    metin = "".join(
        c for c in metin
        if not unicodedata.combining(c)
    )

    # Noktalama işaretlerini boşluğa çevir
    metin = re.sub(
        r"[^A-Z0-9]+",
        " ",
        metin
    )

    # Birden fazla boşluğu teke indir
    metin = re.sub(
        r"\s+",
        " ",
        metin
    ).strip()

    return metin


# ============================================================
# JSON YÜKLE
# ============================================================

print()
print("=" * 65)
print("ADRES → İLÇE EŞLEŞTİRME")
print("=" * 65)
print()

try:

    with open(
        JSON_DOSYASI,
        "r",
        encoding="utf-8"
    ) as dosya:

        referans = json.load(dosya)

except FileNotFoundError:

    print(
        f"HATA: {JSON_DOSYASI} bulunamadı."
    )

    sys.exit(1)

except Exception as e:

    print(
        f"HATA: JSON okunamadı: {e}"
    )

    sys.exit(1)


# ============================================================
# JSON BÖLÜMLERİ
# ============================================================

ILCELER = referans.get(
    "ilceler",
    {}
)

MAHALLE_ILCE_RAW = referans.get(
    "mahalle_ilce",
    {}
)

SEMTLER_RAW = referans.get(
    "semtler",
    {}
)

ADRES_KALIPLARI = referans.get(
    "adres_kaliplari",
    []
)


# ============================================================
# MAHALLE İNDEKSİNİ NORMALİZE ET
# ============================================================

MAHALLE_ILCE = {}

for mahalle, ilceler in MAHALLE_ILCE_RAW.items():

    mahalle_key = normalize(
        mahalle
    )

    if not mahalle_key:
        continue

    if isinstance(
        ilceler,
        list
    ):

        MAHALLE_ILCE[
            mahalle_key
        ] = ilceler

    elif isinstance(
        ilceler,
        str
    ):

        MAHALLE_ILCE[
            mahalle_key
        ] = [ilceler]


# ============================================================
# SEMT İNDEKSİNİ NORMALİZE ET
# ============================================================

SEMTLER = {}

for semt, ilce in SEMTLER_RAW.items():

    semt_key = normalize(
        semt
    )

    if not semt_key:
        continue

    if isinstance(
        ilce,
        list
    ):

        SEMTLER[
            semt_key
        ] = ilce

    else:

        SEMTLER[
            semt_key
        ] = [str(ilce)]


# ============================================================
# İLÇELERİ NORMALİZE ET
# ============================================================

ILCE_NORMALIZE = {}

for ilce_key, bilgi in ILCELER.items():

    if not isinstance(
        bilgi,
        dict
    ):
        continue

    ilce_adi = bilgi.get(
        "ilce_adi"
    )

    if ilce_adi:

        ILCE_NORMALIZE[
            normalize(ilce_adi)
        ] = ilce_adi


# ============================================================
# ADRES KALIPLARINI HAZIRLA
# ============================================================

KALIPLAR = []

for kayit in ADRES_KALIPLARI:

    if not isinstance(
        kayit,
        dict
    ):
        continue

    ifadeler = kayit.get(
        "ifadeler",
        []
    )

    ilce = kayit.get(
        "ilce"
    )

    if not ifadeler or not ilce:
        continue

    ifadeler_normalize = []

    for ifade in ifadeler:

        ifade_norm = normalize(
            ifade
        )

        if ifade_norm:

            ifadeler_normalize.append(
                ifade_norm
            )

    if ifadeler_normalize:

        KALIPLAR.append({
            "ifadeler":
                ifadeler_normalize,

            "ilce":
                ilce
        })


# ============================================================
# ADRESTEN İLÇE BUL
# ============================================================

def adres_ilce_bul(adres):

    adres_norm = normalize(
        adres
    )

    if not adres_norm:
        return ""


    # ========================================================
    # 1. ADRES KALIPLARI
    # ========================================================
    #
    # Daha özel ve güçlü eşleştirme olduğu için
    # önce bunlara bakıyoruz.
    # ========================================================

    for kayit in KALIPLAR:

        ifadeler = kayit[
            "ifadeler"
        ]

        hepsi_var = True

        for ifade in ifadeler:

            if ifade not in adres_norm:

                hepsi_var = False
                break

        if hepsi_var:

            return kayit[
                "ilce"
            ]


    # ========================================================
    # 2. DOĞRUDAN İLÇE ADI
    # ========================================================

    # Uzun ilçe isimlerini önce dene.
    # Örneğin "KÜÇÜKÇEKMECE" gibi isimlerin
    # daha kısa ifadelerle karışmasını önler.

    ilceler_sirali = sorted(
        ILCE_NORMALIZE.items(),
        key=lambda x: len(x[0]),
        reverse=True
    )

    for ilce_norm, ilce_adi in ilceler_sirali:

        pattern = (
            r"(?<![A-Z0-9])"
            + re.escape(ilce_norm)
            + r"(?![A-Z0-9])"
        )

        if re.search(
            pattern,
            adres_norm
        ):

            return ilce_adi


    # ========================================================
    # 3. MAHALLE
    # ========================================================

    # Adres içinde geçen mahalle isimlerini bul.
    #
    # Bir mahalle birden fazla ilçede varsa,
    # doğrudan karar vermiyoruz.
    # Önce adres içinde ilçe adı da var mı diye kontrol
    # ediyoruz; yoksa tek ilçeli mahalleyi kullanıyoruz.

    mahalle_adaylari = []

    for mahalle_norm, ilceler in MAHALLE_ILCE.items():

        if len(mahalle_norm) < 3:
            continue

        pattern = (
            r"(?<![A-Z0-9])"
            + re.escape(mahalle_norm)
            + r"(?![A-Z0-9])"
        )

        if re.search(
            pattern,
            adres_norm
        ):

            mahalle_adaylari.append(
                (
                    mahalle_norm,
                    ilceler
                )
            )


    # En uzun mahalle adı önce
    mahalle_adaylari.sort(
        key=lambda x: len(x[0]),
        reverse=True
    )


    # Tek ilçeye ait mahalle
    for _, ilceler in mahalle_adaylari:

        if len(ilceler) == 1:

            return ilceler[0]


    # Birden fazla ilçeye ait mahalle varsa,
    # adres içinde o ilçelerden biri geçiyor mu?
    for _, ilceler in mahalle_adaylari:

        for ilce in ilceler:

            ilce_norm = normalize(
                ilce
            )

            pattern = (
                r"(?<![A-Z0-9])"
                + re.escape(ilce_norm)
                + r"(?![A-Z0-9])"
            )

            if re.search(
                pattern,
                adres_norm
            ):

                return ilce


    # ========================================================
    # 4. SEMT / BÖLGE
    # ========================================================

    semt_adaylari = []

    for semt_norm, ilceler in SEMTLER.items():

        if len(semt_norm) < 3:
            continue

        pattern = (
            r"(?<![A-Z0-9])"
            + re.escape(semt_norm)
            + r"(?![A-Z0-9])"
        )

        if re.search(
            pattern,
            adres_norm
        ):

            semt_adaylari.append(
                (
                    semt_norm,
                    ilceler
                )
            )


    # Uzun ifadeler önce
    semt_adaylari.sort(
        key=lambda x: len(x[0]),
        reverse=True
    )


    # Tek ilçeye ait semt
    for _, ilceler in semt_adaylari:

        if len(ilceler) == 1:

            return ilceler[0]


    # Birden fazla olasılık varsa,
    # adres içindeki ilçe adıyla doğrulamaya çalış.
    for _, ilceler in semt_adaylari:

        for ilce in ilceler:

            ilce_norm = normalize(
                ilce
            )

            pattern = (
                r"(?<![A-Z0-9])"
                + re.escape(ilce_norm)
                + r"(?![A-Z0-9])"
            )

            if re.search(
                pattern,
                adres_norm
            ):

                return ilce


    # ========================================================
    # BULUNAMADI
    # ========================================================

    return ""


# ============================================================
# EXCEL OKU
# ============================================================

print(
    "Excel okunuyor..."
)

try:

    df = pd.read_excel(
        GIRDI_DOSYASI
    )

except FileNotFoundError:

    print()
    print(
        f"HATA: {GIRDI_DOSYASI} bulunamadı."
    )

    sys.exit(1)

except Exception as e:

    print()
    print(
        f"HATA: Excel okunamadı: {e}"
    )

    sys.exit(1)


# ============================================================
# ADRES KOLONU KONTROL
# ============================================================

if "ADRES" not in df.columns:

    print()
    print(
        "HATA: Excel içinde ADRES kolonu bulunamadı."
    )

    print()
    print(
        "Mevcut kolonlar:"
    )

    for kolon in df.columns:

        print(
            f"- {kolon}"
        )

    sys.exit(1)


# ============================================================
# EŞLEŞTİRME
# ============================================================

print()
print("=" * 65)
print("ADRESLER EŞLEŞTİRİLİYOR")
print("=" * 65)
print()

ilceler_sonuc = []

bulundu = 0
bulunamadi = 0

bulunamayan_adresler = []


toplam = len(df)


for sira, adres in enumerate(
    df["ADRES"],
    start=1
):

    ilce = adres_ilce_bul(
        adres
    )

    ilceler_sonuc.append(
        ilce
    )


    if ilce:

        bulundu += 1

    else:

        bulunamadi += 1

        if not pd.isna(adres):

            bulunamayan_adresler.append(
                str(adres)
            )


    # Her 100 kayıtta ilerleme göster
    if (
        sira % 100 == 0
        or sira == toplam
    ):

        print(
            f"{sira:,} / {toplam:,}"
        )


# ============================================================
# ILCE KOLONU
# ============================================================

df["ILCE"] = ilceler_sonuc


# ============================================================
# EXCEL KAYDET
# ============================================================

print()
print(
    "Excel kaydediliyor..."
)

try:

    df.to_excel(
        CIKTI_DOSYASI,
        index=False
    )

except Exception as e:

    print()
    print(
        f"HATA: Excel kaydedilemedi: {e}"
    )

    sys.exit(1)


# ============================================================
# SONUÇ
# ============================================================

print()
print("=" * 65)
print("TAMAMLANDI")
print("=" * 65)
print()

print(
    f"Toplam adres     : {toplam:,}"
)

print(
    f"Eşleşen          : {bulundu:,}"
)

print(
    f"Eşleşmeyen       : {bulunamadi:,}"
)

if toplam > 0:

    oran = (
        bulundu / toplam
    ) * 100

    print(
        f"Eşleşme oranı    : %{oran:.2f}"
    )

print()
print(
    f"Çıktı dosyası    : {CIKTI_DOSYASI}"
)


# ============================================================
# BULUNAMAYANLAR
# ============================================================

if bulunamayan_adresler:

    print()
    print("=" * 65)
    print(
        "BULUNAMAYAN ADRESLER"
    )
    print("=" * 65)
    print()

    for adres in bulunamayan_adresler:

        print(
            f"- {adres}"
        )

else:

    print()
    print(
        "Tüm adresler eşleştirildi."
    )

print()