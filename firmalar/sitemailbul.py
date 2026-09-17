import pandas as pd
import time
import random
import re
import json
import hashlib
import os

from datetime import datetime
from difflib import SequenceMatcher
from urllib.parse import quote, urlparse

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from webdriver_manager.chrome import ChromeDriverManager

from openpyxl import load_workbook
from openpyxl.styles import PatternFill


# =====================================================
# AYARLAR
# =====================================================

MAX_GOOGLE_SONUC = 15
MIN_GOOGLE_PUAN = 70
MIN_SITE_PUAN = 70
MIN_SITE_PUAN_ADRES_YOK = 100

BEKLEME = (1.2, 2.0)

MAX_ILETISIM_SAYFASI = 5
MAX_BODY = 15000

SITE_TIMEOUT = 10
GOOGLE_TIMEOUT = 15

MAX_ADAY_SITE_KONTROL = 5

EXCEL_KAYIT_ARALIGI = 5

ILERLEME_DOSYASI = "sitemailbul_ilerleme.json"
DOSYA = "firmalar_web_mail.xlsx"


# =====================================================
# CHROME
# =====================================================

def chrome_baslat():

    print("\n  → Chrome başlatılıyor...")

    options = Options()
    options.debugger_address = "127.0.0.1:9222"
    options.page_load_strategy = "eager"

    driver = webdriver.Chrome(
        service=Service(
            ChromeDriverManager().install()
        ),
        options=options
    )

    driver.set_page_load_timeout(GOOGLE_TIMEOUT)

    print("  ✓ Chrome hazır")

    return driver


driver = chrome_baslat()


# =====================================================
# EXCEL
# =====================================================

df = pd.read_excel("firmalar.xlsx")

print("\nExcel sütunları:")
print(df.columns.tolist())

print("\nToplam firma:", len(df))


# =====================================================
# ŞEHİR / İLÇE JSON
# =====================================================

with open(
    "ilceler.json",
    "r",
    encoding="utf-8"
) as f:

    SEHIRLER = json.load(f)


# =====================================================
# ŞEHİR SEÇ
# =====================================================

def sehir_sec():

    print("\n" + "=" * 50)
    print("ŞEHİR SEÇİMİ")
    print("=" * 50)

    print("Mevcut şehirler:")
    print(", ".join(sorted(SEHIRLER.keys())))

    print("=" * 50)

    while True:

        secim = input(
            "Şehir adını girin: "
        ).strip()

        bulunan_sehir = None

        for sehir in SEHIRLER:

            if sehir.casefold() == secim.casefold():

                bulunan_sehir = sehir
                break

        if bulunan_sehir:
            return bulunan_sehir

        print("\n⚠ Şehir bulunamadı.")
        print(
            "Lütfen JSON dosyasındaki şehirlerden "
            "birini girin.\n"
        )


# =====================================================
# AKTİF ŞEHİR
# =====================================================

SEHIR = sehir_sec()

SEHIR_ILCELERI = set(
    SEHIRLER[SEHIR]
)

print("\n" + "=" * 50)
print("SEÇİM TAMAMLANDI")
print("=" * 50)

print(f"Şehir : {SEHIR}")
print(f"İlçe  : {len(SEHIR_ILCELERI)} adet")

print("=" * 50)


# =====================================================
# İLERLEME - KAYNAK DOSYA HASH
# =====================================================

def kaynak_dosya_hash():

    sha256 = hashlib.sha256()

    with open("firmalar.xlsx", "rb") as f:

        while True:

            parca = f.read(1024 * 1024)

            if not parca:
                break

            sha256.update(parca)

    return sha256.hexdigest()


KAYNAK_HASH = kaynak_dosya_hash()


# =====================================================
# İLERLEME KAYDET
# =====================================================

def ilerleme_kaydet(sonraki_index, sonuclar):

    veri = {
        "kaynak_dosya": "firmalar.xlsx",
        "kaynak_hash": KAYNAK_HASH,
        "toplam_firma": len(df),
        "sehir": SEHIR,
        "sonraki_index": sonraki_index,
        "tamamlanan_firma": sonraki_index,
        "sonuclar": sonuclar,
        "zaman": datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    }

    try:

        gecici_dosya = ILERLEME_DOSYASI + ".tmp"

        with open(
            gecici_dosya,
            "w",
            encoding="utf-8"
        ) as f:

            json.dump(
                veri,
                f,
                ensure_ascii=False,
                indent=2
            )

        os.replace(
            gecici_dosya,
            ILERLEME_DOSYASI
        )

        print(
            f"  ✓ İlerleme kaydedildi "
            f"({sonraki_index} firma tamamlandı)"
        )

    except Exception as e:

        print(
            "  ⚠ İlerleme kaydedilemedi:",
            e
        )


# =====================================================
# İLERLEME YÜKLE
# =====================================================

def ilerleme_yukle():

    if not os.path.exists(
        ILERLEME_DOSYASI
    ):
        return None

    try:

        with open(
            ILERLEME_DOSYASI,
            "r",
            encoding="utf-8"
        ) as f:

            veri = json.load(f)

        if veri.get(
            "kaynak_dosya"
        ) != "firmalar.xlsx":

            print(
                "\n⚠ İlerleme dosyasının "
                "kaynak dosyası farklı."
            )

            return None

        if veri.get(
            "kaynak_hash"
        ) != KAYNAK_HASH:

            print(
                "\n⚠ firmalar.xlsx değişmiş."
            )

            print(
                "Eski ilerleme dosyası "
                "kullanılmayacak."
            )

            return None

        if veri.get(
            "toplam_firma"
        ) != len(df):

            print(
                "\n⚠ Firma sayısı değişmiş."
            )

            return None

        if veri.get(
            "sehir"
        ) != SEHIR:

            print(
                "\n⚠ Önceki çalışma farklı "
                "bir şehir için yapılmış."
            )

            return None

        if not isinstance(
            veri.get("sonuclar"),
            list
        ):

            print(
                "\n⚠ İlerleme dosyası bozuk."
            )

            return None

        return veri

    except Exception as e:

        print(
            "\n⚠ İlerleme dosyası okunamadı:",
            e
        )

        return None


# =====================================================
# EXCEL KAYDETME
# =====================================================

def excel_kaydet(sonuclar):

    try:

        if sonuclar:

            sonuc_df = pd.DataFrame(
                sonuclar
            )

        else:

            sonuc_df = pd.DataFrame(
                columns=[
                    "UNVAN",
                    "KAYNAK_ADRES",
                    "KAYNAK_ILCE",
                    "WEB",
                    "MAIL",
                    "WEB_ILCE",
                    "ADRES_DURUMU",
                    "DURUM",
                    "SITE_PUANI"
                ]
            )

        sonuc_df.to_excel(
            DOSYA,
            index=False
        )

        wb = load_workbook(DOSYA)
        ws = wb.active

        acik_kirmizi = PatternFill(
            fill_type="solid",
            fgColor="FCE4D6"
        )

        basliklar = {}

        for col in range(
            1,
            ws.max_column + 1
        ):

            baslik = ws.cell(
                row=1,
                column=col
            ).value

            basliklar[baslik] = col

        if "DURUM" in basliklar:

            durum_sutunu = basliklar["DURUM"]

            for row in range(
                2,
                ws.max_row + 1
            ):

                durum = ws.cell(
                    row=row,
                    column=durum_sutunu
                ).value

                if durum in [
                    "SITE BULUNDU - İLÇE FARKLI",
                    "SITE BULUNDU - ADRES YOK"
                ]:

                    for col in range(
                        1,
                        ws.max_column + 1
                    ):

                        ws.cell(
                            row=row,
                            column=col
                        ).fill = acik_kirmizi

        ws.freeze_panes = "A2"

        if ws.max_row >= 1:
            ws.auto_filter.ref = ws.dimensions

        for column in ws.columns:

            max_length = 0

            column_letter = (
                column[0].column_letter
            )

            for cell in column:

                try:

                    length = len(
                        str(cell.value)
                    )

                    if length > max_length:
                        max_length = length

                except:
                    pass

            ws.column_dimensions[
                column_letter
            ].width = min(
                max_length + 2,
                50
            )

        wb.save(DOSYA)

        print(
            f"  ✓ Excel kaydedildi: {DOSYA}"
        )

    except Exception as e:

        print(
            "  ⚠ Excel kaydedilemedi:",
            e
        )


# =====================================================
# ESKİ İLERLEME VAR MI?
# =====================================================

sonuclar = []
baslangic_index = 0

eski_ilerleme = ilerleme_yukle()

if eski_ilerleme:

    eski_index = int(
        eski_ilerleme.get(
            "sonraki_index",
            0
        )
    )

    eski_index = max(
        0,
        min(
            eski_index,
            len(df)
        )
    )

    eski_sonuclar = (
        eski_ilerleme.get(
            "sonuclar",
            []
        )
    )

    print("\n")
    print("=" * 70)
    print("KAYDEDİLMİŞ İLERLEME BULUNDU")
    print("=" * 70)

    print(
        f"Tamamlanan firma : "
        f"{eski_index}/{len(df)}"
    )

    print(
        f"Kalan firma      : "
        f"{len(df) - eski_index}"
    )

    if eski_index < len(df):

        print(
            f"Sıradaki firma   : "
            f"{eski_index + 1}/{len(df)}"
        )

    print(
        f"Şehir            : "
        f"{SEHIR}"
    )

    print(
        f"Son kayıt zamanı  : "
        f"{eski_ilerleme.get('zaman', '-')}"
    )

    print("=" * 70)

    devam = input(
        "\nDaha önce kaydedilmiş ilerleme bulundu. "
        "Devam etmek istiyor musunuz? (E/h): "
    ).strip().lower()

    if devam in ["", "e", "evet"]:

        sonuclar = eski_sonuclar
        baslangic_index = eski_index

        if baslangic_index < len(df):

            print(
                f"\n✓ {baslangic_index + 1}. firmadan "
                "devam edilecek."
            )

        else:

            print(
                "\n✓ Tüm firmalar zaten tamamlanmış."
            )

    else:

        print(
            "\n→ Yeni çalışma başlatılıyor."
        )

        sonuclar = []
        baslangic_index = 0

        try:
            os.remove(
                ILERLEME_DOSYASI
            )
        except:
            pass

else:

    if os.path.exists(
        ILERLEME_DOSYASI
    ):

        print(
            "\n⚠ Eski ilerleme dosyası "
            "bu Excel/şehir ile uyumlu değil."
        )

        print(
            "→ Yeni çalışma başlatılıyor."
        )

    sonuclar = []
    baslangic_index = 0


# =====================================================
# IGNORE DOMAINLER
# =====================================================

IGNORE_DOMAINLER = {

    "google.com",
    "google.com.tr",
    "gstatic.com",

    "linkedin.com",
    "facebook.com",
    "instagram.com",
    "youtube.com",
    "twitter.com",
    "x.com",
    "tiktok.com",

    "bulurum.com",
    "find.com.tr",
    "118.com.tr",
    "firmasec.com",
    "firmabulucu.com",
    "firmaatlas.com",
    "mukellef.info",
    "infobel.com",
    "kompass.com",
    "europages.com",
    "verif.com",
    "listofcompany.com",
    "manuzone.com",
    "b2bhint.com",
    "yelp.com",
    "fobshanghai.com",

    "ticaretsicil.gov.tr",
    "ito.org.tr",
    "atonet.org.tr",

    "sahibinden.com",
    "alibaba.com",
    "amazon.com",
    "trendyol.com",

    "haberler.com",
    "medium.com",
    "eksisozluk.com",
    "emlakkulisi.com",

    "ihalepro.com",
    "ihalekik.com",
    "ihaleciler.com",
    "ihale.com",
    "kamubilgisistemi.com",

    "kariyer.net",
    "happycenter.com",
    "zavis.ai",
    "synevo.com.tr",
    "tesisat.com.tr",
}


# =====================================================
# DOMAIN İÇERİSİNDE RED EDİLECEK KELİMELER
# =====================================================

IGNORE_DOMAIN_KELIMELERI = {

    "rehber",
    "rehberi",
    "firmarehberi",
    "firmabul",
    "firmalar",
    "firmaara",
    "firmasec",
    "firmaatlas",
    "businessdirectory",
    "directory",
    "directorysite",
    "companies",
    "companylist",
    "listofcompany",
    "b2bhint",
    "fobshanghai",
    "yelp",
    "yellowpages",
    "yellowpage",
    "classifieds",
    "ilan",
    "ilanlar",
    "ihale",
    "ihaleler",
    "auction",
    "marketplace",
    "pazar",
    "pazaryeri",
    "haber",
    "haberler",
    "blog",
    "forum",
}


# =====================================================
# DOMAIN
# =====================================================

def domain_adi(url):

    try:

        domain = urlparse(
            url
        ).netloc.lower()

        domain = domain.split(":")[0]

    except:

        domain = str(
            url
        ).lower()

    return domain.replace(
        "www.",
        ""
    )


def domain_koku(url):

    domain = domain_adi(url)

    uzantilar = [
        ".com.tr",
        ".net.tr",
        ".org.tr",
        ".gen.tr",
        ".web.tr",
        ".biz.tr",
        ".info.tr",
        ".tv.tr",
        ".name.tr",
        ".com",
        ".net",
        ".org",
        ".biz",
        ".info",
        ".co"
    ]

    for uzanti in uzantilar:

        if domain.endswith(uzanti):

            return domain[
                :-len(uzanti)
            ]

    parcalar = domain.split(".")

    if len(parcalar) >= 2:
        return parcalar[-2]

    return domain


# =====================================================
# TÜRKÇE TEMİZLE
# =====================================================

def temizle(text):

    text = str(text).lower()

    text = text.replace(
        "i̇",
        "i"
    )

    ceviri = str.maketrans(
        "çğıöşü",
        "cgiosu"
    )

    return text.translate(ceviri)


# =====================================================
# DOMAIN IGNORE KONTROL
# =====================================================

def domain_ignored_mi(url):

    domain = domain_adi(url)

    if not domain:
        return True

    for ignore in IGNORE_DOMAINLER:

        if domain == ignore:
            return True

        if domain.endswith(
            "." + ignore
        ):
            return True

    domain_temiz = temizle(domain)

    for kelime in IGNORE_DOMAIN_KELIMELERI:

        if kelime in domain_temiz:
            return True

    if domain.endswith(
        (
            ".gov.tr",
            ".edu.tr",
            ".bel.tr",
            ".k12.tr"
        )
    ):
        return True

    return False


# =====================================================
# URL KONTROL
# =====================================================

def url_gecerli_mi(url):

    if not url:
        return False

    url_lower = url.lower().strip()

    if url_lower.startswith(
        ("mailto:", "tel:")
    ):
        return False

    if url_lower.endswith(
        (".pdf", ".doc", ".docx")
    ):
        return False

    if domain_ignored_mi(url):
        return False

    return True


# =====================================================
# STOPWORDS
# =====================================================

STOPWORDS = {

    "sanayi",
    "ticaret",
    "limited",
    "ltd",
    "ltdsti",
    "sirketi",
    "anonim",
    "anonimsirketi",
    "as",
    "ve",
    "ithalat",
    "ihracat",
    "hizmet",
    "hizmetleri",
    "teknoloji",
    "teknolojileri",
    "sistem",
    "sistemleri",
    "cozum",
    "cozumleri",
    "makine",
    "makina",
    "muhendislik",
    "mimarlik",
    "insaat",
    "enerji",
    "lojistik",
    "metal",
    "plastik",
    "tekstil",
    "gida",
    "otomotiv",
    "kimya",
    "tasarim",
    "tasarimi",
    "elektrik",
    "elektronik",
    "tic",
    "san",
    "denetim",
    "kalite",
    "guvenlik",
    "cevre",
    "laboratuvar",
    "hastane",
    "geri",
    "donusum",
    "yapi",
    "danismanlik",
    "mobilya",
    "taahhut",
    "malzeme",
    "urun",
    "urunleri",
    "merkez",
    "grup",
    "group",
    "holding",
    "dis",
    "disi",
    "saglik",
    "medikal",
    "firma",
    "sirket",
}


# =====================================================
# FİRMA KELİMELERİ
# =====================================================

def firma_kelimeleri(unvan):

    unvan = temizle(unvan)

    kelimeler = re.findall(
        r"[a-z0-9]+",
        unvan
    )

    sonuc = []

    for kelime in kelimeler:

        if len(kelime) <= 2:
            continue

        if kelime.isdigit():
            continue

        if kelime in STOPWORDS:
            continue

        if kelime not in sonuc:
            sonuc.append(kelime)

    return sonuc


# =====================================================
# BENZERLİK
# =====================================================

def benzerlik(a, b):

    return SequenceMatcher(
        None,
        temizle(a),
        temizle(b)
    ).ratio()


# =====================================================
# DOMAIN - FİRMA EŞLEŞMESİ
# =====================================================

def domain_firma_eslesmesi(unvan, url):

    domain = temizle(
        domain_koku(url)
    )

    kelimeler = firma_kelimeleri(unvan)

    if not kelimeler:
        return 0

    puan = 0

    for kelime in kelimeler:

        if kelime in domain:

            if len(kelime) >= 8:
                puan += 100

            elif len(kelime) >= 5:
                puan += 70

            else:
                puan += 40

    en_iyi = 0

    for kelime in kelimeler:

        oran = benzerlik(
            kelime,
            domain
        )

        if oran > en_iyi:
            en_iyi = oran

    if en_iyi >= 0.90:
        puan += 50

    elif en_iyi >= 0.80:
        puan += 30

    elif en_iyi >= 0.70:
        puan += 15

    return puan


# =====================================================
# GOOGLE PUANI
# =====================================================

def google_puani(unvan, url, baslik):

    domain = temizle(
        domain_koku(url)
    )

    baslik = temizle(baslik)

    kelimeler = firma_kelimeleri(unvan)

    if not kelimeler:
        return 0

    puan = 0

    domain_eslesen = 0
    title_eslesen = 0

    for kelime in kelimeler:

        if kelime in domain:

            domain_eslesen += 1

            if len(kelime) >= 8:
                puan += 100

            elif len(kelime) >= 5:
                puan += 70

            else:
                puan += 35

    for kelime in kelimeler:

        if kelime in baslik:
            title_eslesen += 1

    if title_eslesen >= 3:
        puan += 70

    elif title_eslesen == 2:
        puan += 45

    elif title_eslesen == 1:
        puan += 20

    en_iyi = 0

    for kelime in kelimeler:

        oran = benzerlik(
            kelime,
            domain
        )

        if oran > en_iyi:
            en_iyi = oran

    puan += int(
        en_iyi * 30
    )

    if (
        domain_eslesen == 0
        and title_eslesen == 0
    ):
        return 0

    return puan


# =====================================================
# GOOGLE DOMAIN
# =====================================================

def google_domain_al(sonuc):

    try:

        cite = sonuc.find_element(
            By.CSS_SELECTOR,
            "cite"
        ).text.strip()

        if cite:

            cite = cite.replace(
                "›",
                "/"
            )

            match = re.search(
                r"([A-Za-z0-9-]+\.[A-Za-z0-9.-]+)",
                cite
            )

            if match:

                domain = match.group(1)

                domain = domain.lower()
                domain = domain.split("/")[0]

                return domain.replace(
                    "www.",
                    ""
                )

    except:
        pass

    try:

        vuu = sonuc.find_element(
            By.CSS_SELECTOR,
            "span.VuuXrf"
        ).text.strip()

        if vuu:

            vuu = re.sub(
                r"^https?://",
                "",
                vuu,
                flags=re.IGNORECASE
            )

            vuu = vuu.split("/")[0]
            vuu = vuu.split("›")[0].strip()

            if re.fullmatch(
                r"(?:www\.)?"
                r"[A-Za-z0-9-]+"
                r"(?:\.[A-Za-z0-9-]+)+",
                vuu
            ):

                return vuu.replace(
                    "www.",
                    ""
                ).lower()

    except:
        pass

    return ""


def google_gercek_url(sonuc):

    domain = google_domain_al(sonuc)

    if not domain:
        return ""

    if domain_ignored_mi(
        "https://" + domain
    ):
        return ""

    return "https://" + domain + "/"


# =====================================================
# MAIL
# =====================================================

def mailleri_bul(metin):

    pattern = r"""
        [a-zA-Z0-9._%+\-]+
        @
        [a-zA-Z0-9.\-]+
        \.
        [a-zA-Z]{2,}
    """

    mailler = re.findall(
        pattern,
        metin,
        re.VERBOSE
    )

    sonuc = []

    for mail in mailler:

        mail = mail.lower().strip()

        if len(mail) > 100:
            continue

        if mail in sonuc:
            continue

        if mail.endswith(
            (
                ".png",
                ".jpg",
                ".jpeg",
                ".gif",
                ".webp",
                ".svg"
            )
        ):
            continue

        if mail_adresi_gecersiz_mi(mail):
            continue

        sonuc.append(mail)

    return sonuc


def mail_adresi_gecersiz_mi(mail):

    mail = mail.lower().strip()

    if "@" not in mail:
        return True

    kullanici, domain = mail.split(
        "@",
        1
    )

    if len(kullanici) < 2:
        return True

    if len(domain) < 4:
        return True

    yasak_kullanici = {

        "example",
        "ornek",
        "örnek",
        "test",
        "testing",
        "demo",
        "deneme",
        "sample",
        "user",
        "username",
        "email",
        "mail",
        "yourmail",
        "yourname",
        "name",
        "abc",
        "abcd",
        "xxx",
        "xxxxx",
        "noreply",
        "no-reply",
        "donotreply",
        "do-not-reply",
    }

    if kullanici in yasak_kullanici:
        return True

    yasak_domain = {

        "example.com",
        "example.org",
        "example.net",
        "test.com",
        "test.com.tr",
        "ornek.com",
        "ornek.com.tr",
        "domain.com",
        "domain.com.tr",
        "nginx.com",
    }

    if domain in yasak_domain:
        return True

    if mail in {
        "available@nginx.com",
        "root@localhost",
        "admin@localhost",
    }:
        return True

    return False


# =====================================================
# MAIL PUANI
# =====================================================

def mail_puani(mail, site_domain):

    mail = mail.lower().strip()

    try:

        kullanici, mail_domain = mail.split(
            "@",
            1
        )

    except:

        return -1

    puan = 0

    site_domain = temizle(site_domain)
    mail_domain = temizle(mail_domain)

    if (
        site_domain
        and (
            mail_domain == site_domain
            or mail_domain.endswith(
                "." + site_domain
            )
        )
    ):
        puan += 100

    if kullanici in {
        "info",
        "iletisim",
        "contact",
        "bilgi",
    }:

        puan += 80

    elif kullanici in {
        "satis",
        "sales",
        "pazarlama",
        "muhasebe",
        "accounting",
    }:

        puan += 70

    elif kullanici in {
        "ofis",
        "office",
        "destek",
        "support",
    }:

        puan += 50

    domain_kok = re.sub(
        r"[^a-z0-9]",
        "",
        site_domain
    )

    kullanici_temiz = re.sub(
        r"[^a-z0-9]",
        "",
        kullanici
    )

    if (
        domain_kok
        and len(domain_kok) >= 5
        and domain_kok in kullanici_temiz
    ):
        puan += 120

    ucretsiz = {
        "gmail.com",
        "hotmail.com",
        "hotmail.com.tr",
        "outlook.com",
        "outlook.com.tr",
        "yahoo.com",
        "yahoo.com.tr",
    }

    if mail_domain in ucretsiz:
        puan += 10

    return puan


def en_iyi_mail(mailler, site_domain):

    if not mailler:
        return ""

    puanli = []

    for mail in mailler:

        if mail_adresi_gecersiz_mi(mail):
            continue

        puan = mail_puani(
            mail,
            site_domain
        )

        puanli.append(
            (puan, mail)
        )

    if not puanli:
        return ""

    puanli.sort(
        key=lambda x: x[0],
        reverse=True
    )

    return puanli[0][1]


# =====================================================
# ADRES BÖLÜMÜ
# =====================================================

def adres_bolumu_bul(body):

    text = str(body)

    satirlar = [
        x.strip()
        for x in text.splitlines()
        if x.strip()
    ]

    anahtarlar = [
        "adres",
        "address",
        "iletişim",
        "iletisim",
        "contact",
        "contact us",
        "bize ulaşın",
        "bize ulasin",
        "office",
        "merkez",
        "head office",
        "merkez ofis",
    ]

    aday = []

    for i, satir in enumerate(satirlar):

        satir_temiz = temizle(satir)

        uygun = False

        for anahtar in anahtarlar:

            if temizle(anahtar) in satir_temiz:

                uygun = True
                break

        if not uygun:
            continue

        baslangic = max(0, i)
        bitis = min(
            len(satirlar),
            i + 8
        )

        parca = " ".join(
            satirlar[
                baslangic:bitis
            ]
        )

        aday.append(parca)

    if not aday:

        sehir = temizle(SEHIR)

        adres_patternleri = [

            rf".{{0,150}}mah\.?.{{0,150}}{re.escape(sehir)}",

            rf".{{0,150}}mahallesi.{{0,150}}{re.escape(sehir)}",

            rf".{{0,150}}cad\.?.{{0,150}}{re.escape(sehir)}",

            rf".{{0,150}}sok\.?.{{0,150}}{re.escape(sehir)}",

            rf".{{0,150}}\b\d{{5}}\b.{{0,150}}{re.escape(sehir)}",
        ]

        temiz_body = " ".join(satirlar)

        for pattern in adres_patternleri:

            eslesmeler = re.findall(
                pattern,
                temiz_body,
                re.IGNORECASE
            )

            aday.extend(eslesmeler)

    return "\n".join(aday)


# =====================================================
# ADRESTEN İLÇE BUL
# =====================================================

def adresten_ilce_bul(adres):

    adres = temizle(adres)

    bulunan = []

    for ilce in SEHIR_ILCELERI:

        ilce_temiz = temizle(ilce)

        pattern = (
            rf"\b{re.escape(ilce_temiz)}\b"
        )

        if re.search(
            pattern,
            adres
        ):

            bulunan.append(ilce)

    if bulunan:
        return bulunan[0]

    sehir_temiz = temizle(SEHIR)

    if re.search(
        rf"\b{re.escape(sehir_temiz)}\b",
        adres
    ):
        return SEHIR

    return ""


# =====================================================
# SAYFA VERİSİ
# =====================================================

def sayfa_verisi_al():

    try:

        body = driver.find_element(
            By.TAG_NAME,
            "body"
        ).text[:MAX_BODY]

    except:

        body = ""

    try:
        title = driver.title

    except:
        title = ""

    meta_description = ""

    try:

        meta = driver.find_element(
            By.XPATH,
            "//meta[@name='description']"
        )

        meta_description = (
            meta.get_attribute("content")
            or ""
        )

    except:
        pass

    metin = (
        title
        + "\n"
        + body
        + "\n"
        + meta_description
    )

    mailler = mailleri_bul(metin)

    adres = adres_bolumu_bul(body)

    ilce = adresten_ilce_bul(adres)

    return {
        "metin": metin,
        "body": body,
        "mail": mailler,
        "adres": adres,
        "ilce": ilce,
        "title": title
    }


# =====================================================
# İLETİŞİM LİNKLERİ
# =====================================================

def iletisim_linklerini_bul():

    linkler_sonuc = []

    anahtarlar = [
        "iletisim",
        "contact",
        "adres",
        "address",
        "bize ulasin",
        "bize ulaşın",
        "location",
        "office",
        "contact us",
    ]

    try:

        linkler = driver.find_elements(
            By.TAG_NAME,
            "a"
        )

        for link in linkler:

            try:

                href = (
                    link.get_attribute("href")
                    or ""
                )

                link_text = temizle(
                    link.text
                )

                if not href:
                    continue

                if href.lower().startswith(
                    (
                        "mailto:",
                        "tel:",
                        "javascript:"
                    )
                ):
                    continue

                href_temiz = temizle(href)

                uygun = False

                for kelime in anahtarlar:

                    kelime_temiz = temizle(kelime)

                    if kelime_temiz in link_text:
                        uygun = True
                        break

                    if kelime_temiz in href_temiz:
                        uygun = True
                        break

                if not uygun:
                    continue

                try:

                    link_domain = domain_adi(href)
                    mevcut_domain = domain_adi(
                        driver.current_url
                    )

                    if (
                        link_domain
                        and mevcut_domain
                        and link_domain != mevcut_domain
                    ):
                        continue

                except:
                    pass

                if href not in linkler_sonuc:
                    linkler_sonuc.append(href)

            except:
                continue

    except:
        pass

    return linkler_sonuc


# =====================================================
# SITE - FİRMA İLİŞKİ PUANI
# =====================================================

def site_firma_iliski_puani(
    unvan,
    sayfa_metni,
    url
):

    metin = temizle(sayfa_metni)

    kelimeler = firma_kelimeleri(unvan)

    if not kelimeler:
        return 0

    puan = 0

    domain = temizle(
        domain_koku(url)
    )

    domain_eslesme = 0
    metin_eslesme = 0

    for kelime in kelimeler:

        if kelime in domain:
            domain_eslesme += 1

        if kelime in metin:
            metin_eslesme += 1

    if domain_eslesme >= 2:
        puan += 50

    elif domain_eslesme == 1:
        puan += 30

    if metin_eslesme >= 3:
        puan += 50

    elif metin_eslesme == 2:
        puan += 35

    elif metin_eslesme == 1:
        puan += 15

    return puan


# =====================================================
# SITE DOĞRULA
# =====================================================

def site_dogrula(
    unvan,
    kaynak_ilce,
    url
):

    try:

        if domain_ignored_mi(url):

            return {
                "dogru": False,
                "site_puani": 0,
                "mail": "",
                "web_ilce": "",
                "web_adres": "",
                "adres_durumu": "",
                "durum": "IGNORE DOMAIN"
            }

        domain = domain_adi(url)

        try:

            driver.set_page_load_timeout(
                SITE_TIMEOUT
            )

            driver.get(url)

        except:

            try:
                driver.execute_script(
                    "window.stop();"
                )
            except:
                pass

            time.sleep(0.7)

        try:

            WebDriverWait(
                driver,
                6
            ).until(
                EC.presence_of_element_located(
                    (
                        By.TAG_NAME,
                        "body"
                    )
                )
            )

        except:

            return {
                "dogru": False,
                "site_puani": 0,
                "mail": "",
                "web_ilce": "",
                "web_adres": "",
                "adres_durumu": "",
                "durum": "SITE ACILMADI"
            }

        time.sleep(
            random.uniform(0.4, 0.8)
        )

        ana_veri = sayfa_verisi_al()

        bulunan_mailler = list(
            ana_veri["mail"]
        )

        web_ilce = ana_veri["ilce"]
        web_adres = ana_veri["adres"]

        ilk_iliski = site_firma_iliski_puani(
            unvan,
            ana_veri["metin"],
            url
        )

        if ilk_iliski == 0:

            return {
                "dogru": False,
                "site_puani": 0,
                "mail": "",
                "web_ilce": web_ilce,
                "web_adres": web_adres,
                "adres_durumu": (
                    "ADRES BULUNAMADI"
                    if not web_ilce
                    else "ILCE FARKLI"
                ),
                "durum": "SITE ILISKISI ZAYIF"
            }

        iletisim_linkleri = (
            iletisim_linklerini_bul()
        )

        iletisim_linkleri = (
            iletisim_linkleri[
                :MAX_ILETISIM_SAYFASI
            ]
        )

        ziyaret_edilen = set()

        for link in iletisim_linkleri:

            try:

                if link in ziyaret_edilen:
                    continue

                ziyaret_edilen.add(link)

                driver.set_page_load_timeout(
                    SITE_TIMEOUT
                )

                driver.get(link)

                WebDriverWait(
                    driver,
                    5
                ).until(
                    EC.presence_of_element_located(
                        (
                            By.TAG_NAME,
                            "body"
                        )
                    )
                )

                time.sleep(0.4)

                veri = sayfa_verisi_al()

                for mail in veri["mail"]:

                    if mail not in bulunan_mailler:
                        bulunan_mailler.append(mail)

                if veri["adres"]:
                    web_adres = veri["adres"]

                if veri["ilce"]:
                    web_ilce = veri["ilce"]

                if (
                    bulunan_mailler
                    and web_ilce
                ):
                    break

            except:

                continue

        title = ana_veri["title"]

        google_site_puani = google_puani(
            unvan,
            url,
            title
        )

        domain_eslesme = (
            domain_firma_eslesmesi(
                unvan,
                url
            )
        )

        site_puani = (
            domain_eslesme
            + int(
                google_site_puani * 0.35
            )
        )

        firma_kelime_listesi = (
            firma_kelimeleri(unvan)
        )

        site_metni = temizle(
            ana_veri["metin"]
        )

        site_kelime_eslesmesi = 0

        for kelime in firma_kelime_listesi:

            if kelime in site_metni:
                site_kelime_eslesmesi += 1

        if site_kelime_eslesmesi >= 3:
            site_puani += 45

        elif site_kelime_eslesmesi == 2:
            site_puani += 30

        elif site_kelime_eslesmesi == 1:
            site_puani += 10

        kaynak_ilce_temiz = temizle(
            kaynak_ilce
        ).strip()

        web_ilce_temiz = temizle(
            web_ilce
        ).strip()

        if not web_ilce:

            adres_durumu = "ADRES BULUNAMADI"

        elif web_ilce_temiz == kaynak_ilce_temiz:

            adres_durumu = "ADRES AYNI"

        else:

            adres_durumu = "ILCE FARKLI"

        secilen_mail = en_iyi_mail(
            bulunan_mailler,
            domain
        )

        if site_puani < MIN_SITE_PUAN:

            return {
                "dogru": False,
                "site_puani": site_puani,
                "mail": "",
                "web_ilce": web_ilce,
                "web_adres": web_adres,
                "adres_durumu": adres_durumu,
                "durum": "SITE ISMI ZAYIF"
            }

        if web_ilce:

            if web_ilce_temiz == kaynak_ilce_temiz:

                durum = "DOĞRULANDI"
                dogru = True

            else:

                durum = (
                    "SITE BULUNDU - İLÇE FARKLI"
                )

                dogru = True

        else:

            durum = (
                "SITE BULUNDU - ADRES YOK"
            )

            dogru = (
                site_puani
                >= MIN_SITE_PUAN_ADRES_YOK
            )

        return {
            "dogru": dogru,
            "site_puani": site_puani,
            "mail": secilen_mail,
            "web_ilce": web_ilce,
            "web_adres": web_adres,
            "adres_durumu": adres_durumu,
            "durum": durum
        }

    except Exception:

        return {
            "dogru": False,
            "site_puani": 0,
            "mail": "",
            "web_ilce": "",
            "web_adres": "",
            "adres_durumu": "SITE HATASI",
            "durum": "SITE HATASI"
        }


# =====================================================
# GOOGLE ARAMA
# =====================================================

def google_arama(
    firma,
    max_deneme=3
):

    arama_url = (
        "https://www.google.com/search?q="
        + quote(firma)
    )

    for deneme in range(
        1,
        max_deneme + 1
    ):

        print(
            "  → Google aranıyor..."
        )

        try:

            driver.set_page_load_timeout(
                GOOGLE_TIMEOUT
            )

            # Aynı Chrome penceresinde,
            # aynı sekmede yeni arama yapılır.
            driver.get(arama_url)

            WebDriverWait(
                driver,
                GOOGLE_TIMEOUT
            ).until(
                EC.presence_of_element_located(
                    (
                        By.ID,
                        "search"
                    )
                )
            )

            time.sleep(
                random.uniform(*BEKLEME)
            )

            return True

        except Exception:

            if deneme < max_deneme:

                print(
                    f"  ⚠ Google denemesi "
                    f"{deneme} başarısız, "
                    "aynı Chrome sekmesinde tekrar deneniyor..."
                )

                time.sleep(2)

    print(
        "  ✗ Google araması başarısız."
    )

    return False


# =====================================================
# FİRMALARI TARA
# =====================================================

kaydedilecek_index = baslangic_index


try:

    for i in range(
        baslangic_index,
        len(df)
    ):

        row = df.iloc[i]

        firma = str(
            row["UNVAN"]
        ).strip()

        kaynak_adres = str(
            row["ADRES"]
        ).strip()

        kaynak_ilce = str(
            row["ILCE"]
        ).strip()

        print(
            f"\n[{i + 1}/{len(df)}] {firma}"
        )

        bulunan_site = ""
        bulunan_mail = ""
        bulunan_web_ilce = ""
        bulunan_web_adres = ""

        durum = "SITE BULUNAMADI"
        site_puani = 0
        adres_durumu = ""

        # =================================================
        # GOOGLE
        # =================================================

        google_basarili = google_arama(firma)

        if not google_basarili:

            print()
            print(
                "  ⚠ Google araması başarısız."
            )

            print(
                "  ⚠ Bu firma tamamlanmış sayılmayacak."
            )

            print(
                f"  ⚠ {i + 1}. firma tekrar denenecek."
            )

            print(
                "  ⚠ İlerleme kaydediliyor."
            )

            # i = sıradaki firmanın 0-based indexi
            # Bu nedenle aynı firma tekrar denenir.
            ilerleme_kaydet(
                i,
                sonuclar
            )

            excel_kaydet(
                sonuclar
            )

            kaydedilecek_index = i

            break

        # =================================================
        # GOOGLE SONUÇLARI
        # =================================================

        try:

            google_sonuclari = driver.find_elements(
                By.CSS_SELECTOR,
                "div.yuRUbf"
            )

            adaylar = []

            gorulen_domainler = set()

            for sira, sonuc in enumerate(
                google_sonuclari[
                    :MAX_GOOGLE_SONUC
                ],
                start=1
            ):

                try:

                    baslik = sonuc.find_element(
                        By.TAG_NAME,
                        "h3"
                    ).text.strip()

                except:

                    continue

                if not baslik:
                    continue

                url = google_gercek_url(
                    sonuc
                )

                if not url:
                    continue

                if not url_gecerli_mi(url):
                    continue

                domain = domain_adi(url)

                if not domain:
                    continue

                if domain in gorulen_domainler:
                    continue

                gorulen_domainler.add(domain)

                puan = google_puani(
                    firma,
                    url,
                    baslik
                )

                if sira == 1:
                    puan += 30

                elif sira == 2:
                    puan += 20

                elif sira == 3:
                    puan += 10

                domain_eslesme = (
                    domain_firma_eslesmesi(
                        firma,
                        url
                    )
                )

                title_temiz = temizle(
                    baslik
                )

                firma_kelime_listesi = (
                    firma_kelimeleri(firma)
                )

                title_eslesme = sum(
                    1
                    for kelime
                    in firma_kelime_listesi
                    if kelime in title_temiz
                )

                if (
                    domain_eslesme == 0
                    and title_eslesme == 0
                ):
                    continue

                if domain_ignored_mi(url):
                    continue

                if puan < MIN_GOOGLE_PUAN:
                    continue

                adaylar.append({

                    "url": url,
                    "domain": domain,
                    "puan": puan,
                    "domain_eslesme": domain_eslesme,
                    "title_eslesme": title_eslesme,
                    "sira": sira

                })

            adaylar.sort(
                key=lambda x: (
                    x["puan"],
                    x["domain_eslesme"],
                    x["title_eslesme"],
                    -x["sira"]
                ),
                reverse=True
            )

            adaylar = adaylar[
                :MAX_ADAY_SITE_KONTROL
            ]

            if adaylar:

                print(
                    f"  ✓ {len(adaylar)} uygun "
                    "site adayı bulundu."
                )

            else:

                print(
                    "  ✗ Uygun site bulunamadı."
                )

            # =================================================
            # SITELERI KONTROL ET
            # =================================================

            for aday in adaylar:

                print(
                    f"  → {aday['domain']} "
                    "kontrol ediliyor..."
                )

                sonuc = site_dogrula(
                    firma,
                    kaynak_ilce,
                    aday["url"]
                )

                if sonuc["dogru"]:

                    bulunan_site = aday["url"]

                    bulunan_mail = sonuc["mail"]

                    bulunan_web_ilce = (
                        sonuc["web_ilce"]
                    )

                    bulunan_web_adres = (
                        sonuc["web_adres"]
                    )

                    site_puani = (
                        sonuc["site_puani"]
                    )

                    durum = sonuc["durum"]

                    adres_durumu = (
                        sonuc["adres_durumu"]
                    )

                    print(
                        f"  ✓ Site bulundu | "
                        f"Mail: {sonuc['mail'] or '-'} | "
                        f"Kaynak İlçe: "
                        f"{kaynak_ilce or '-'} | "
                        f"Web İlçe: "
                        f"{sonuc['web_ilce'] or '-'} | "
                        f"Adres: "
                        f"{sonuc['adres_durumu']}"
                    )

                    break

        except Exception as e:

            print(
                "  ✗ Hata:",
                e
            )

            durum = "GENEL HATA"

        # =================================================
        # SONUCU EKLE
        # =================================================

        sonuclar.append({

            "UNVAN": firma,
            "KAYNAK_ADRES": kaynak_adres,
            "KAYNAK_ILCE": kaynak_ilce,
            "WEB": bulunan_site,
            "MAIL": bulunan_mail,
            "WEB_ILCE": bulunan_web_ilce,
            "ADRES_DURUMU": adres_durumu,
            "DURUM": durum,
            "SITE_PUANI": site_puani

        })

        # =================================================
        # BİR SONRAKİ FİRMA INDEXİ
        # =================================================

        kaydedilecek_index = i + 1

        # =================================================
        # İLERLEME KAYDET
        # =================================================

        ilerleme_kaydet(
            kaydedilecek_index,
            sonuclar
        )

        # =================================================
        # HER 5 FİRMADA EXCEL
        # =================================================

        if (
            kaydedilecek_index
            % EXCEL_KAYIT_ARALIGI
            == 0
        ):

            excel_kaydet(
                sonuclar
            )


except KeyboardInterrupt:

    print("\n")
    print("=" * 70)
    print("İŞLEM KULLANICI TARAFINDAN DURDURULDU.")
    print("=" * 70)

    print(
        f"Tamamlanan firma : "
        f"{kaydedilecek_index}/{len(df)}"
    )

    print(
        f"Kalan firma      : "
        f"{len(df) - kaydedilecek_index}"
    )

    if kaydedilecek_index < len(df):

        print(
            f"Sıradaki firma   : "
            f"{kaydedilecek_index + 1}/{len(df)}"
        )

    print()

    ilerleme_kaydet(
        kaydedilecek_index,
        sonuclar
    )

    excel_kaydet(
        sonuclar
    )

    print()
    print(
        "✓ İlerleme kaydedildi."
    )

    if kaydedilecek_index < len(df):

        print(
            "✓ Sonraki çalıştırmada "
            f"{kaydedilecek_index + 1}. firmadan "
            "devam edilecek."
        )

    print(
        f"\nİlerleme dosyası: "
        f"{ILERLEME_DOSYASI}"
    )

    print(
        f"Sonuç dosyası: "
        f"{DOSYA}"
    )

    print("=" * 70)

    try:
        driver.quit()
    except:
        pass

    raise SystemExit


# =====================================================
# GOOGLE HATASI / YARIM KALMA
# =====================================================

if kaydedilecek_index < len(df):

    print("\n")
    print("=" * 70)
    print("İŞLEM TAMAMLANMADI.")
    print("=" * 70)

    print(
        f"Tamamlanan firma : "
        f"{kaydedilecek_index}/{len(df)}"
    )

    print(
        f"Kalan firma      : "
        f"{len(df) - kaydedilecek_index}"
    )

    print(
        f"Sıradaki firma   : "
        f"{kaydedilecek_index + 1}/{len(df)}"
    )

    print(
        "\nİlerleme dosyası korunuyor."
    )

    print(
        "Tekrar çalıştırdığınızda "
        f"{kaydedilecek_index + 1}. firmadan "
        "devam edecektir."
    )

    print("=" * 70)

    try:
        driver.quit()
    except:
        pass

    raise SystemExit


# =====================================================
# CHROME KAPAT
# =====================================================

try:
    driver.quit()
except:
    pass


# =====================================================
# SON EXCEL KAYDI
# =====================================================

excel_kaydet(
    sonuclar
)


# =====================================================
# İLERLEME DOSYASINI SİL
# =====================================================

try:

    if os.path.exists(
        ILERLEME_DOSYASI
    ):

        os.remove(
            ILERLEME_DOSYASI
        )

        print(
            "\n✓ İlerleme dosyası silindi."
        )

except Exception as e:

    print(
        "\n⚠ İlerleme dosyası silinemedi:",
        e
    )


# =====================================================
# İSTATİSTİK
# =====================================================

sonuc_df = pd.DataFrame(
    sonuclar
)

bulunan = (
    sonuc_df["WEB"]
    .fillna("")
    .astype(str)
    .str.strip()
    .ne("")
    .sum()
)

mail_bulunan = (
    sonuc_df["MAIL"]
    .fillna("")
    .astype(str)
    .str.strip()
    .ne("")
    .sum()
)

adres_ayni = (
    sonuc_df["ADRES_DURUMU"]
    .fillna("")
    .astype(str)
    .eq("ADRES AYNI")
    .sum()
)

ilce_farkli = (
    sonuc_df["ADRES_DURUMU"]
    .fillna("")
    .astype(str)
    .eq("ILCE FARKLI")
    .sum()
)

adres_yok = (
    sonuc_df["ADRES_DURUMU"]
    .fillna("")
    .astype(str)
    .eq("ADRES BULUNAMADI")
    .sum()
)


# =====================================================
# TAMAMLANDI
# =====================================================

print(
    "\n"
    + "=" * 70
)

print(
    "İŞLEM TAMAMLANDI."
)

print(
    "=" * 70
)

print(
    f"Toplam firma       : "
    f"{len(sonuc_df)}"
)

print(
    f"Bulunan site       : "
    f"{bulunan}"
)

print(
    f"Mail bulunan       : "
    f"{mail_bulunan}"
)

print(
    f"Adres aynı         : "
    f"{adres_ayni}"
)

print(
    f"İlçe farklı        : "
    f"{ilce_farkli}"
)

print(
    f"Adres bulunamadı   : "
    f"{adres_yok}"
)

print(
    f"Site bulunamayan   : "
    f"{len(sonuc_df) - bulunan}"
)

print(
    "\nDosya: firmalar_web_mail.xlsx"
)

print(
    "=" * 70
)
