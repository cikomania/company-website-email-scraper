import time
import re
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager


# ============================================================
# AYARLAR
# ============================================================

# Buradan aradığın meslek grubu için linki 
# ve oluşturulacak excel dosya adını değiştir cano

SEHIR = "Ankara"

URL = "https://www.atonet.org.tr/KomiteListe?komiteKodu=25"

OUTPUT = "ato_25.xlsx"


# ============================================================
# ANKARA İLÇELERİ
# ============================================================

SEHIR_ILCELERI = {

    "AKYURT",
    "ALTINDAĞ",
    "AYAŞ",
    "BALA",
    "BEYPAZARI",
    "ÇAMLIDERE",
    "ÇANKAYA",
    "ÇUBUK",
    "ELMADAĞ",
    "ETİMESGUT",
    "EVREN",
    "GÖLBAŞI",
    "GÜDÜL",
    "HAYMANA",
    "KAHRAMANKAZAN",
    "KALECİK",
    "KEÇİÖREN",
    "KIZILCAHAMAM",
    "MAMAK",
    "NALLIHAN",
    "POLATLI",
    "PURSAKLAR",
    "SİNCAN",
    "ŞEREFLİKOÇHİSAR",
    "YENİMAHALLE"

}


# ============================================================
# TÜRKÇE KARAKTER NORMALİZASYONU
# ============================================================

def normalize(text):

    text = str(text).strip().upper()

    ceviri = str.maketrans(
        "ÇĞİÖŞÜ",
        "CGIOSU"
    )

    return text.translate(ceviri)


# ============================================================
# ADRESTEN İLÇE BUL
# ============================================================

def ilce_bul(adres):

    if pd.isna(adres):
        return ""

    adres = str(adres).strip()

    if not adres:
        return ""

    adres_normal = normalize(adres)

    sehir_normal = normalize(SEHIR)


    # --------------------------------------------------------
    # ADRESİN SONUNDA /ANKARA VAR MI?
    # --------------------------------------------------------

    if not re.search(
        rf"/\s*{re.escape(sehir_normal)}\s*$",
        adres_normal
    ):
        return ""


    # --------------------------------------------------------
    # SADECE GERÇEK İLÇELERDEN BİRİNİ KABUL ET
    # --------------------------------------------------------

    for ilce in SEHIR_ILCELERI:

        ilce_normal = normalize(ilce)

        pattern = (
            rf"(?:^|[\s,])"
            rf"{re.escape(ilce_normal)}"
            rf"\s*/\s*"
            rf"{re.escape(sehir_normal)}"
            rf"\s*$"
        )

        if re.search(
            pattern,
            adres_normal
        ):

            return ilce


    return ""


# ============================================================
# CHROME
# ============================================================

print()
print("=" * 70)
print(f"ATO {SEHIR.upper()} KOMİTE LİSTESİ")
print("=" * 70)

print()
print("Chrome başlatılıyor...")


options = Options()

options.add_argument(
    "--start-maximized"
)


driver = webdriver.Chrome(
    service=Service(
        ChromeDriverManager().install()
    ),
    options=options
)


# ============================================================
# SAYFAYI AÇ
# ============================================================

print("ATO sayfası açılıyor...")

driver.get(URL)

time.sleep(3)


# ============================================================
# TÜM SAYFALARI TARA
# ============================================================

firmalar = []

sayfa = 1


while True:

    print()
    print(
        f"Sayfa {sayfa} okunuyor..."
    )

    time.sleep(1)


    # --------------------------------------------------------
    # TABLO SATIRLARI
    # --------------------------------------------------------

    rows = driver.find_elements(
        By.CSS_SELECTOR,
        "#TBL_KOMITE tbody tr"
    )

    print(
        f"  Satır sayısı: {len(rows)}"
    )


    if not rows:

        print(
            "  TABLO BULUNAMADI!"
        )

        break


    # --------------------------------------------------------
    # FİRMALARI OKU
    # --------------------------------------------------------

    for row in rows:

        cells = row.find_elements(
            By.TAG_NAME,
            "td"
        )


        if len(cells) < 5:

            continue


        sicil = cells[1].text.strip()

        unvan = cells[2].text.strip()

        adres = cells[3].text.strip()

        telefon = cells[4].text.strip()


        # ----------------------------------------------------
        # BOŞ SATIRLARI ATLA
        # ----------------------------------------------------

        if not sicil or not unvan:

            continue


        # ----------------------------------------------------
        # TELEFON
        # ----------------------------------------------------

        if not telefon:

            telefon = "yok"


        # ----------------------------------------------------
        # İLÇE
        # ----------------------------------------------------

        ilce = ilce_bul(
            adres
        )


        # ----------------------------------------------------
        # FİRMAYI EKLE
        # ----------------------------------------------------

        firmalar.append({

            "Ticaret Sicil No":
                sicil,

            "Unvan":
                unvan,

            "Adres":
                adres,

            "İlçe":
                ilce,

            "Firma Tel":
                telefon

        })


    # ========================================================
    # SONRAKİ BUTONU BUL
    # ========================================================

    next_button = driver.find_element(
        By.CSS_SELECTOR,
        "#TBL_KOMITE_next"
    )


    # --------------------------------------------------------
    # SON SAYFA MI?
    # --------------------------------------------------------

    class_name = (
        next_button
        .get_attribute("class")
        or ""
    )


    if "disabled" in class_name:

        print()
        print(
            "  Son sayfaya ulaşıldı."
        )

        break


    # --------------------------------------------------------
    # ŞU ANKİ İLK FİRMAYI HATIRLA
    # --------------------------------------------------------

    eski_ilk_sicil = ""

    try:

        eski_ilk_sicil = rows[0].find_elements(
            By.TAG_NAME,
            "td"
        )[1].text.strip()

    except:

        pass


    # ========================================================
    # SONRAKİ SAYFAYA TIKLA
    # ========================================================

    driver.execute_script(
        "arguments[0].click();",
        next_button
    )


    # ========================================================
    # TABLONUN DEĞİŞMESİNİ BEKLE
    # ========================================================

    baslangic = time.time()


    while time.time() - baslangic < 10:

        time.sleep(0.2)


        yeni_rows = driver.find_elements(
            By.CSS_SELECTOR,
            "#TBL_KOMITE tbody tr"
        )


        if not yeni_rows:

            continue


        try:

            yeni_ilk_sicil = (
                yeni_rows[0]
                .find_elements(
                    By.TAG_NAME,
                    "td"
                )[1]
                .text
                .strip()
            )

        except:

            continue


        if (
            yeni_ilk_sicil
            and yeni_ilk_sicil
            != eski_ilk_sicil
        ):

            break


    sayfa += 1


# ============================================================
# BROWSER KAPAT
# ============================================================

driver.quit()


# ============================================================
# DATAFRAME
# ============================================================

df = pd.DataFrame(
    firmalar
)


# ============================================================
# AYNI FİRMALARI TEMİZLE
# ============================================================

df = df.drop_duplicates(
    subset=[
        "Ticaret Sicil No"
    ],
    keep="first"
)


# ============================================================
# EXCEL
# ============================================================

df.to_excel(
    OUTPUT,
    index=False
)


# ============================================================
# İSTATİSTİK
# ============================================================

toplam = len(df)


ilce_bulunan = (
    df["İlçe"]
    .fillna("")
    .astype(str)
    .str.strip()
    .ne("")
    .sum()
)


ilce_bulunamayan = (
    toplam
    - ilce_bulunan
)


telefon_var = (
    df["Firma Tel"]
    .fillna("")
    .astype(str)
    .str.strip()
    .ne("yok")
    .sum()
)


telefon_yok = (
    df["Firma Tel"]
    .fillna("")
    .astype(str)
    .str.strip()
    .eq("yok")
    .sum()
)


# ============================================================
# SONUÇ
# ============================================================

print()
print("=" * 70)
print("TAMAMLANDI")
print("=" * 70)

print(
    f"Toplam firma      : {toplam}"
)

print(
    f"İlçe bulunan      : {ilce_bulunan}"
)

print(
    f"İlçe bulunamayan  : {ilce_bulunamayan}"
)

print(
    f"Telefon bulunan   : {telefon_var}"
)

print(
    f"Telefon bulunamayan: {telefon_yok}"
)

print(
    f"Excel             : {OUTPUT}"
)

print("=" * 70)