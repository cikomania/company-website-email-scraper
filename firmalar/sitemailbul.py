import pandas as pd
import time
import random
import re
import json

from difflib import SequenceMatcher
from urllib.parse import quote, urlparse, urljoin

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

# Google sonucunun minimum puanı
MIN_GOOGLE_PUAN = 70

# Site kimliği minimum puanı
MIN_SITE_PUAN = 70

# Site adres yoksa kabul için daha yüksek puan gerekir
MIN_SITE_PUAN_ADRES_YOK = 100

# Google / site beklemeleri
BEKLEME = (1.2, 2.0)

# İletişim sayfası maksimum
MAX_ILETISIM_SAYFASI = 5

# Sayfadan alınacak maksimum body
MAX_BODY = 15000

# Site açma timeout
SITE_TIMEOUT = 10

# Google timeout
GOOGLE_TIMEOUT = 15

# Google'dan çok zayıf adayları siteye sokmamak için
MAX_ADAY_SITE_KONTROL = 5


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
# CHROME YENİDEN BAŞLAT
# =====================================================

def chrome_yeniden_baslat():

    global driver

    print("  ⚠ Chrome yeniden başlatılıyor...")

    try:
        driver.quit()
    except:
        pass

    time.sleep(2)

    driver = chrome_baslat()

    time.sleep(2)

    try:
        driver.get("https://www.google.com/")
        time.sleep(2)
    except:
        pass

    return driver


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
# IGNORE DOMAINLER
# =====================================================

# Buradaki amaç:
# Firma sitesine benzese bile aslında firma sitesi olmayan
# rehber / portal / sosyal medya / ihale / pazar yeri vb.
# sitelere hiç girmemek.

IGNORE_DOMAINLER = {

    # Arama / Google
    "google.com",
    "google.com.tr",
    "gstatic.com",

    # Sosyal medya
    "linkedin.com",
    "facebook.com",
    "instagram.com",
    "youtube.com",
    "twitter.com",
    "x.com",
    "tiktok.com",

    # Firma rehberleri
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

    # Ticaret / resmi kayıt
    "ticaretsicil.gov.tr",
    "ito.org.tr",
    "atonet.org.tr",

    # İlan / pazar yeri
    "sahibinden.com",
    "alibaba.com",
    "amazon.com",
    "trendyol.com",

    # Haber / içerik
    "haberler.com",
    "medium.com",
    "eksisozluk.com",
    "emlakkulisi.com",

    # İhale
    "ihalepro.com",
    "ihalekik.com",
    "ihaleciler.com",
    "ihale.com",
    "kamubilgisistemi.com",

    # Diğer genel platformlar
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
# TLD / DOMAIN TEMİZLE
# =====================================================

def domain_adi(url):

    try:

        domain = urlparse(url).netloc.lower()

        domain = domain.split(":")[0]

    except:

        domain = str(url).lower()

    domain = domain.replace("www.", "")

    return domain


# =====================================================
# DOMAIN KÖKÜ
# =====================================================

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

            return domain[:-len(uzanti)]

    parcalar = domain.split(".")

    if len(parcalar) >= 2:

        return parcalar[-2]

    return domain


# =====================================================
# TÜRKÇE TEMİZLE
# =====================================================

def temizle(text):

    text = str(text).lower()

    text = text.replace("i̇", "i")

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

    # Tam domain kontrolü
    for ignore in IGNORE_DOMAINLER:

        if domain == ignore:
            return True

        if domain.endswith("." + ignore):
            return True

    # Alt alan adı da kontrol edilir
    domain_temiz = temizle(domain)

    for kelime in IGNORE_DOMAIN_KELIMELERI:

        if kelime in domain_temiz:

            return True

    # Resmi / kamu uzantıları
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

    if url_lower.startswith("mailto:"):
        return False

    if url_lower.startswith("tel:"):
        return False

    if url_lower.endswith(".pdf"):
        return False

    if url_lower.endswith(".doc"):
        return False

    if url_lower.endswith(".docx"):
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

    # Ek genel kelimeler
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

    # Domain ile firma ana kelimesi benzerliği
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

def google_puani(
    unvan,
    url,
    baslik
):

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

    # -------------------------------------------------
    # DOMAIN
    # -------------------------------------------------

    for kelime in kelimeler:

        if kelime in domain:

            domain_eslesen += 1

            if len(kelime) >= 8:
                puan += 100

            elif len(kelime) >= 5:
                puan += 70

            else:
                puan += 35

    # -------------------------------------------------
    # TITLE
    # -------------------------------------------------

    for kelime in kelimeler:

        if kelime in baslik:

            title_eslesen += 1

    if title_eslesen >= 3:
        puan += 70

    elif title_eslesen == 2:
        puan += 45

    elif title_eslesen == 1:
        puan += 20

    # -------------------------------------------------
    # DOMAIN BENZERLİĞİ
    # -------------------------------------------------

    en_iyi = 0

    for kelime in kelimeler:

        oran = benzerlik(
            kelime,
            domain
        )

        if oran > en_iyi:
            en_iyi = oran

    puan += int(en_iyi * 30)

    # -------------------------------------------------
    # HİÇ İLİŞKİ YOKSA
    # -------------------------------------------------

    if (
        domain_eslesen == 0
        and title_eslesen == 0
    ):
        return 0

    return puan


# =====================================================
# GOOGLE GERÇEK DOMAIN
# =====================================================

def google_domain_al(sonuc):

    # -------------------------------------------------
    # 1. CITE
    # -------------------------------------------------

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

                domain = domain.replace(
                    "www.",
                    ""
                )

                return domain

    except:
        pass

    # -------------------------------------------------
    # 2. VuuXrf
    # -------------------------------------------------

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


# =====================================================
# GOOGLE GERÇEK URL
# =====================================================

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
# MAIL BUL
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

        # Dosya uzantısı gibi görünenleri ele
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

        # Teknik / sahte mail
        if mail_adresi_gecersiz_mi(mail):
            continue

        sonuc.append(mail)

    return sonuc


# =====================================================
# GEÇERSİZ MAIL KONTROL
# =====================================================

def mail_adresi_gecersiz_mi(mail):

    mail = mail.lower().strip()

    if "@" not in mail:
        return True

    kullanici, domain = mail.split(
        "@",
        1
    )

    # Çok kısa / anlamsız
    if len(kullanici) < 2:
        return True

    if len(domain) < 4:
        return True

    # Açıkça örnek/test adresleri
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

    # Açıkça sahte domain
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

    # Teknik nginx adresleri
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

def mail_puani(
    mail,
    site_domain
):

    mail = mail.lower().strip()

    try:

        kullanici, mail_domain = mail.split(
            "@",
            1
        )

    except:

        return -1

    puan = 0

    site_domain = temizle(
        site_domain
    )

    mail_domain = temizle(
        mail_domain
    )

    # -------------------------------------------------
    # KURUMSAL DOMAIN
    # -------------------------------------------------

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

    # -------------------------------------------------
    # LOCAL PART
    # -------------------------------------------------

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

    # -------------------------------------------------
    # DOMAIN LOCAL PART'TA VAR MI
    # -------------------------------------------------

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

    # -------------------------------------------------
    # ÜCRETSİZ MAIL
    # -------------------------------------------------

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

        # Yasak değil.
        # Sadece kurumsal maile göre düşük öncelik.
        puan += 10

    return puan


# =====================================================
# EN İYİ MAIL
# =====================================================

def en_iyi_mail(
    mailler,
    site_domain
):

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
            (
                puan,
                mail
            )
        )

    if not puanli:
        return ""

    puanli.sort(
        key=lambda x: x[0],
        reverse=True
    )

    return puanli[0][1]


# =====================================================
# İLÇE BUL
# =====================================================

def ilce_bul(metin):

    metin = temizle(metin)

    bulunan = []

    for ilce in SEHIR_ILCELERI:

        ilce_temiz = temizle(ilce)

        pattern = rf"\b{re.escape(ilce_temiz)}\b"

        if re.search(
            pattern,
            metin
        ):

            bulunan.append(ilce)

    return bulunan


# =====================================================
# ADRES BÖLÜMÜ BUL
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

        baslangic = max(
            0,
            i
        )

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

        pattern = rf"\b{re.escape(ilce_temiz)}\b"

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
# İLETİŞİM LİNKLERİNİ BUL
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
        "iletisim",
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

                # Aynı domain dışına çıkma
                try:

                    link_domain = domain_adi(href)

                    mevcut_url = driver.current_url

                    mevcut_domain = domain_adi(
                        mevcut_url
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
# SITEDE FİRMA İLİŞKİSİ
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

    # Domain
    if domain_eslesme >= 2:
        puan += 50

    elif domain_eslesme == 1:
        puan += 30

    # Site metni
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

        # -------------------------------------------------
        # DOMAIN IGNORE
        # -------------------------------------------------

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

        # -------------------------------------------------
        # ANA SAYFA
        # -------------------------------------------------

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
            random.uniform(
                0.4,
                0.8
            )
        )

        ana_veri = sayfa_verisi_al()

        bulunan_mailler = list(
            ana_veri["mail"]
        )

        web_ilce = ana_veri["ilce"]

        web_adres = ana_veri["adres"]

        # -------------------------------------------------
        # İLK KİMLİK KONTROLÜ
        # -------------------------------------------------

        ilk_iliski = site_firma_iliski_puani(
            unvan,
            ana_veri["metin"],
            url
        )

        # Çok zayıf siteyse iletişim sayfalarına hiç girme
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

        # -------------------------------------------------
        # İLETİŞİM LİNKLERİ
        # -------------------------------------------------

        iletisim_linkleri = (
            iletisim_linklerini_bul()
        )

        iletisim_linkleri = (
            iletisim_linkleri[
                :MAX_ILETISIM_SAYFASI
            ]
        )

        # -------------------------------------------------
        # İLETİŞİM SAYFALARINI TARA
        # -------------------------------------------------

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

                        bulunan_mailler.append(
                            mail
                        )

                if veri["adres"]:

                    web_adres = veri["adres"]

                if veri["ilce"]:

                    web_ilce = veri["ilce"]

                # Mail bulunduysa ve adres de bulunduysa
                # gereksiz diğer iletişim sayfalarına girme
                if (
                    bulunan_mailler
                    and web_ilce
                ):
                    break

            except:

                continue

        # -------------------------------------------------
        # GOOGLE / SITE PUANI
        # -------------------------------------------------

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

        # -------------------------------------------------
        # FİRMA KELİMELERİ SİTEDE
        # -------------------------------------------------

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

        # -------------------------------------------------
        # ADRES DURUMU
        # -------------------------------------------------

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

        # -------------------------------------------------
        # MAIL SEÇ
        # -------------------------------------------------

        secilen_mail = en_iyi_mail(
            bulunan_mailler,
            domain
        )

        # -------------------------------------------------
        # SITE PUANI YETERSİZ
        # -------------------------------------------------

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

        # -------------------------------------------------
        # ADRES VARSA
        # -------------------------------------------------

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

    global driver

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

            driver.get(
                arama_url
            )

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
                random.uniform(
                    *BEKLEME
                )
            )

            return True

        except Exception:

            if deneme < max_deneme:

                try:

                    driver.set_page_load_timeout(
                        10
                    )

                    driver.get(
                        "https://www.google.com/"
                    )

                    time.sleep(2)

                except:
                    pass

            if deneme == 2:

                try:
                    chrome_yeniden_baslat()
                except:
                    pass

    print(
        "  ✗ Google araması başarısız."
    )

    return False


# =====================================================
# FİRMALARI TARA
# =====================================================

sonuclar = []


for i, (_, row) in enumerate(
    df.iterrows(),
    start=1
):

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
        f"\n[{i}/{len(df)}] {firma}"
    )

    bulunan_site = ""

    bulunan_mail = ""

    bulunan_web_ilce = ""

    bulunan_web_adres = ""

    durum = "SITE BULUNAMADI"

    site_puani = 0

    adres_durumu = ""

    try:

        # =================================================
        # GOOGLE
        # =================================================

        google_basarili = google_arama(
            firma
        )

        if not google_basarili:

            durum = "GOOGLE HATASI"

            sonuclar.append({

                "UNVAN": firma,

                "KAYNAK_ADRES": kaynak_adres,

                "KAYNAK_ILCE": kaynak_ilce,

                "WEB": "",

                "MAIL": "",

                "WEB_ILCE": "",

                "ADRES_DURUMU": "",

                "DURUM": durum,

                "SITE_PUANI": 0
            })

            continue

        # =================================================
        # GOOGLE SONUÇLARI
        # =================================================

        google_sonuclari = driver.find_elements(
            By.CSS_SELECTOR,
            "div.yuRUbf"
        )

        adaylar = []

        gorulen_domainler = set()

        # =================================================
        # GOOGLE SONUÇLARINI DEĞERLENDİR
        # =================================================

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

            # -------------------------------------------------
            # GERÇEK URL
            # -------------------------------------------------

            url = google_gercek_url(
                sonuc
            )

            if not url:
                continue

            # -------------------------------------------------
            # URL GEÇERLİ Mİ
            # -------------------------------------------------

            if not url_gecerli_mi(url):
                continue

            domain = domain_adi(url)

            if not domain:
                continue

            # -------------------------------------------------
            # AYNI DOMAIN
            # -------------------------------------------------

            if domain in gorulen_domainler:
                continue

            gorulen_domainler.add(domain)

            # -------------------------------------------------
            # GOOGLE PUANI
            # -------------------------------------------------

            puan = google_puani(
                firma,
                url,
                baslik
            )

            # -------------------------------------------------
            # SIRA BONUSU
            # -------------------------------------------------

            if sira == 1:

                puan += 30

            elif sira == 2:

                puan += 20

            elif sira == 3:

                puan += 10

            # -------------------------------------------------
            # DOMAIN EŞLEŞMESİ
            # -------------------------------------------------

            domain_eslesme = (
                domain_firma_eslesmesi(
                    firma,
                    url
                )
            )

            # -------------------------------------------------
            # TITLE EŞLEŞMESİ
            # -------------------------------------------------

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

            # -------------------------------------------------
            # HİÇ İLİŞKİ YOK
            # -------------------------------------------------

            if (
                domain_eslesme == 0
                and title_eslesme == 0
            ):

                continue

            # -------------------------------------------------
            # GENEL REHBER GİBİ GÖRÜNEN DOMAIN
            # -------------------------------------------------

            if domain_ignored_mi(url):

                continue

            # -------------------------------------------------
            # MIN GOOGLE PUANI
            # -------------------------------------------------

            if puan < MIN_GOOGLE_PUAN:

                continue

            # -------------------------------------------------
            # ADAY
            # -------------------------------------------------

            adaylar.append({

                "url": url,

                "domain": domain,

                "puan": puan,

                "domain_eslesme":
                    domain_eslesme,

                "title_eslesme":
                    title_eslesme,

                "sira": sira
            })

        # =================================================
        # ADAYLARI SIRALA
        # =================================================

        adaylar.sort(

            key=lambda x: (
                x["puan"],
                x["domain_eslesme"],
                x["title_eslesme"],
                -x["sira"]
            ),

            reverse=True
        )

        # =================================================
        # SADECE EN İYİ ADAYLAR
        # =================================================

        adaylar = adaylar[
            :MAX_ADAY_SITE_KONTROL
        ]

        if adaylar:

            print(
                f"  ✓ {len(adaylar)} uygun site adayı bulundu."
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
                f"  → {aday['domain']} kontrol ediliyor..."
            )

            sonuc = site_dogrula(

                firma,

                kaynak_ilce,

                aday["url"]
            )

            # =================================================
            # SADECE KABUL EDİLEN SITE
            # =================================================

            if sonuc["dogru"]:

                bulunan_site = (
                    aday["url"]
                )

                bulunan_mail = (
                    sonuc["mail"]
                )

                bulunan_web_ilce = (
                    sonuc["web_ilce"]
                )

                bulunan_web_adres = (
                    sonuc["web_adres"]
                )

                site_puani = (
                    sonuc["site_puani"]
                )

                durum = (
                    sonuc["durum"]
                )

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

        "UNVAN":
            firma,

        "KAYNAK_ADRES":
            kaynak_adres,

        "KAYNAK_ILCE":
            kaynak_ilce,

        "WEB":
            bulunan_site,

        "MAIL":
            bulunan_mail,

        "WEB_ILCE":
            bulunan_web_ilce,

        "ADRES_DURUMU":
            adres_durumu,

        "DURUM":
            durum,

        "SITE_PUANI":
            site_puani
    })


# =====================================================
# CHROME KAPAT
# =====================================================

try:
    driver.quit()
except:
    pass


# =====================================================
# EXCEL
# =====================================================

DOSYA = "firmalar_web_mail.xlsx"

sonuc_df = pd.DataFrame(
    sonuclar
)

sonuc_df.to_excel(
    DOSYA,
    index=False
)


# =====================================================
# EXCEL RENKLENDİRME
# =====================================================

wb = load_workbook(
    DOSYA
)

ws = wb.active


# =====================================================
# RENKLER
# =====================================================

acik_kirmizi = PatternFill(
    fill_type="solid",
    fgColor="FCE4D6"
)


# =====================================================
# BAŞLIKLAR
# =====================================================

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


durum_sutunu = basliklar[
    "DURUM"
]


# =====================================================
# SATIRLARI RENKLENDİR
# =====================================================

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


# =====================================================
# EXCEL KULLANIM KOLAYLIKLARI
# =====================================================

ws.freeze_panes = "A2"

ws.auto_filter.ref = ws.dimensions


# =====================================================
# SÜTUN GENİŞLİKLERİ
# =====================================================

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


# =====================================================
# KAYDET
# =====================================================

wb.save(
    DOSYA
)


# =====================================================
# İSTATİSTİK
# =====================================================

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
    f"Toplam firma       : {len(sonuc_df)}"
)

print(
    f"Bulunan site       : {bulunan}"
)

print(
    f"Mail bulunan       : {mail_bulunan}"
)

print(
    f"Adres aynı         : {adres_ayni}"
)

print(
    f"İlçe farklı        : {ilce_farkli}"
)

print(
    f"Adres bulunamadı   : {adres_yok}"
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
