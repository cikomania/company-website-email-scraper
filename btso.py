# python3 btso.py çalıştırdıktan sonra meslek grubu seçimini ve doğrulama kodunu elle gir
# Altta liste açıldıktan sonra terminalde enter bas
# Unvan, Adres, İlçe, Web, Meslek Grubu No, Meslek Grubu Adı sütunlarıyla excel oluşturacak
# Eğer hata verip durursa terminale python3 btso.py ve aynı aşamaları yap, kaldığı sayfadan devam eder

# python -m pip install selenium webdriver-manager pandas openpyxl beautifulsoup4 requests

import os
import re
import json
import time
import random
import requests
import pandas as pd

from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    WebDriverException,
)

from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


# =============================================================================
# AYARLAR
# =============================================================================

BASE_URL = (
    "https://www.btso.org.tr/"
    "hizmetler/online-islemler/kayitli-uyeler/"
)

DETAIL_URL = (
    "https://www.btso.org.tr/"
    "hizmetler/online-islemler/kayitli-uyeler/"
    "kayitli-uyeler-detay"
)

# Normal durumda firma POST'ları arasındaki kısa bekleme
DETAIL_WAIT_MIN = 0.5
DETAIL_WAIT_MAX = 0.9

# Sayfa değişiminden sonra
PAGE_WAIT_MIN = 1.0
PAGE_WAIT_MAX = 1.8

# Rate-limit olursa ilk bekleme
RATE_LIMIT_INITIAL_WAIT = 8

# Rate-limit maksimum bekleme
RATE_LIMIT_MAX_WAIT = 60

# Aynı firma isteğini kaç kez deneyeceğiz?
MAX_DETAIL_RETRY = 5

# Aynı sayfaya kaç kez tekrar girmeyi deneyelim?
MAX_PAGE_RETRY = 5

# Detay POST timeout
DETAIL_TIMEOUT = 30

# Sayfa yükleme timeout
PAGE_TIMEOUT = 25


# =============================================================================
# BURSA İLÇELERİ
# =============================================================================

BURSA_ILCELERI = [
    "BÜYÜKORHAN",
    "GEMLİK",
    "GÜRSU",
    "HARMANCIK",
    "İNEGÖL",
    "İZNİK",
    "KARACABEY",
    "KELES",
    "KESTEL",
    "MUDANYA",
    "MUSTAFAKEMALPAŞA",
    "NİLÜFER",
    "ORHANELİ",
    "ORHANGAZİ",
    "OSMANGAZİ",
    "YENİŞEHİR",
    "YILDIRIM",
]


# =============================================================================
# GENEL YARDIMCI FONKSİYONLAR
# =============================================================================

def temiz_metin(text):
    if text is None:
        return ""

    text = str(text)

    text = text.replace("\xa0", " ")
    text = text.replace("\r", " ")
    text = text.replace("\n", " ")
    text = text.replace("\t", " ")

    text = re.sub(r"\s+", " ", text)

    return text.strip()


def turkce_upper(text):
    """
    Türkçe karakterleri bozmadan büyük harfe çevirir.

    Nilüfer  -> NİLÜFER
    nilüfer  -> NİLÜFER
    NİLÜFER  -> NİLÜFER
    """

    if not text:
        return ""

    text = temiz_metin(text)

    ceviri = str.maketrans({
        "i": "İ",
        "ı": "I",
        "ş": "Ş",
        "ğ": "Ğ",
        "ü": "Ü",
        "ö": "Ö",
        "ç": "Ç",
    })

    return text.translate(ceviri).upper()


# =============================================================================
# İLÇE BUL
# =============================================================================

def ilce_bul(adres):
    """
    Sadece "... İLÇE/BURSA" yapısını dikkate alır.

    Örnek:
    NİLÜFER/BURSA
    Nilüfer/Bursa
    nilüfer / BURSA

    hepsi -> NİLÜFER
    """

    if not adres:
        return ""

    adres_norm = turkce_upper(adres)

    for ilce in BURSA_ILCELERI:

        pattern = (
            rf"(?:^|[\s,;/.-])"
            rf"{re.escape(ilce)}"
            rf"\s*/\s*BURSA"
            rf"(?:\b|$)"
        )

        if re.search(
            pattern,
            adres_norm,
            flags=re.IGNORECASE
        ):
            return ilce

    return ""


# =============================================================================
# MESLEK GRUBU
# =============================================================================

def meslek_grubu_ayristir(meslek):
    """
    04. GRUP : MADENLER VE...

    ->
    04. GRUP
    MADENLER VE...
    """

    if not meslek:
        return "", ""

    meslek = temiz_metin(meslek)

    meslek = re.sub(
        r"^\s*:\s*",
        "",
        meslek
    )

    match = re.match(
        r"^(\d{2}\.\s*GRUP)\s*:\s*(.+)$",
        meslek,
        flags=re.IGNORECASE
    )

    if match:

        return (
            temiz_metin(match.group(1)),
            temiz_metin(match.group(2)),
        )

    return "", meslek


# =============================================================================
# DETAIL HTML PARSE
# =============================================================================

def parse_detail_html(html, liste_unvani=""):
    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    unvan = ""
    adres = ""
    web = ""
    meslek = ""

    # -------------------------------------------------------------------------
    # Unvan
    # -------------------------------------------------------------------------

    h3 = soup.find("h3")

    if h3:

        unvan = temiz_metin(
            h3.get_text(
                " ",
                strip=True
            )
        )

    if not unvan:
        unvan = temiz_metin(
            liste_unvani
        )

    # -------------------------------------------------------------------------
    # Satırlar
    # -------------------------------------------------------------------------

    for row in soup.find_all("tr"):

        cells = row.find_all("td")

        if len(cells) < 2:
            continue

        alan = temiz_metin(
            cells[0].get_text(
                " ",
                strip=True
            )
        )

        deger = temiz_metin(
            cells[1].get_text(
                " ",
                strip=True
            )
        )

        deger = re.sub(
            r"^\s*:\s*",
            "",
            deger
        )

        alan_norm = turkce_upper(
            alan
        )

        if alan_norm == "ADRES":

            adres = deger

        elif alan_norm == "WEB":

            a = cells[1].find("a")

            if a:

                href = (
                    a.get("href") or ""
                ).strip()

                if href not in (
                    "",
                    "http://",
                    "https://"
                ):
                    web = href

        elif alan_norm == "MESLEK GRUBU":

            meslek = deger

    grup_no, grup_adi = (
        meslek_grubu_ayristir(
            meslek
        )
    )

    return {
        "Unvan": unvan,
        "Adres": adres,
        "İlçe": ilce_bul(adres),
        "Web": web,
        "Meslek Grubu No": grup_no,
        "Meslek Grubu": grup_adi,
    }


# =============================================================================
# CHROME
# =============================================================================

def chrome_baslat():

    options = webdriver.ChromeOptions()

    options.add_argument(
        "--start-maximized"
    )

    options.add_argument(
        "--disable-notifications"
    )

    options.add_argument(
        "--disable-popup-blocking"
    )

    options.add_experimental_option(
        "detach",
        True
    )

    service = Service(
        ChromeDriverManager().install()
    )

    return webdriver.Chrome(
        service=service,
        options=options
    )


# =============================================================================
# REQUESTS SESSION
# =============================================================================

def session_olustur(driver):

    session = requests.Session()

    session.headers.update({
        "User-Agent": driver.execute_script(
            "return navigator.userAgent;"
        ),

        "Accept": (
            "text/html,application/xhtml+xml,"
            "application/xml;q=0.9,image/avif,"
            "image/webp,*/*;q=0.8"
        ),

        "Accept-Language": (
            "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7"
        ),

        "X-Requested-With": "XMLHttpRequest",
    })

    # Selenium'daki cookie'leri requests.Session'a aktar
    for cookie in driver.get_cookies():

        try:
            session.cookies.set(
                cookie["name"],
                cookie["value"],
                domain=cookie.get("domain"),
                path=cookie.get("path", "/"),
            )

        except Exception:
            pass

    return session


def session_cookie_guncelle(session, driver):

    try:

        for cookie in driver.get_cookies():

            try:
                session.cookies.set(
                    cookie["name"],
                    cookie["value"],
                    domain=cookie.get("domain"),
                    path=cookie.get("path", "/"),
                )

            except Exception:
                pass

    except Exception:
        pass


# =============================================================================
# CSRF TOKEN
# =============================================================================

def csrf_token_al(driver):

    elements = driver.find_elements(
        By.CSS_SELECTOR,
        "input[name='__RequestVerificationToken']"
    )

    for element in elements:

        try:

            value = element.get_attribute(
                "value"
            )

            if value:
                return value

        except StaleElementReferenceException:
            continue

    return ""


# =============================================================================
# MESLEK GRUBUNU BUL
# =============================================================================

def grup_bilgisi_al(driver):

    grup_no = ""
    grup_adi = ""

    # Önce select
    try:

        selected = driver.find_element(
            By.CSS_SELECTOR,
            "#SelectedSector option:checked"
        )

        secili = temiz_metin(
            selected.text
        )

        grup_no, grup_adi = (
            meslek_grubu_ayristir(
                secili
            )
        )

    except Exception:
        pass

    # URL fallback
    if not grup_no:

        try:

            url = driver.current_url

            match = re.search(
                r"/unvan/(.+)$",
                url
            )

            if match:

                grup_no, grup_adi = (
                    meslek_grubu_ayristir(
                        match.group(1)
                    )
                )

        except Exception:
            pass

    return grup_no, grup_adi


# =============================================================================
# FİRMA LİSTESİ
# =============================================================================

def firma_listesini_oku(driver):

    elements = driver.find_elements(
        By.CSS_SELECTOR,
        "a.show-details[data-id][data-token]"
    )

    firmalar = []

    for element in elements:

        try:

            firma_id = (
                element.get_attribute(
                    "data-id"
                )
                or ""
            ).strip()

            token = (
                element.get_attribute(
                    "data-token"
                )
                or ""
            ).strip()

            unvan = temiz_metin(
                element.text
            )

            if (
                firma_id
                and token
                and unvan
            ):

                firmalar.append({
                    "id": firma_id,
                    "token": token,
                    "unvan": unvan,
                })

        except StaleElementReferenceException:
            continue

    return firmalar


# =============================================================================
# RATE LIMIT KONTROLÜ
# =============================================================================

def rate_limit_var_mi(response):

    if response.status_code in (
        429,
        403,
    ):
        return True

    text = turkce_upper(
        response.text
    )

    ifadeler = [
        "ÇOK HIZLI İŞLEM YAPTINIZ",
        "LÜTFEN BİRKAÇ SANİYE SONRA TEKRAR DENEYİN",
        "ÇOK HIZLI",
        "TOO MANY REQUESTS",
        "TOO MANY",
    ]

    return any(
        ifade in text
        for ifade in ifadeler
    )


def hata_sayfasi_mi(html):

    text = turkce_upper(
        html
    )

    return (
        "İŞLEM BAŞARISIZ OLDU"
        in text
        or
        "BEKLENMEYEN BİR HATA"
        in text
        or
        "TOO MANY REQUESTS"
        in text
    )


# =============================================================================
# FİRMA DETAIL POST
# =============================================================================

def firma_detayi_getir(
    session,
    firma,
    csrf_token,
    referer
):

    payload = {
        "id": firma["id"],
        "token": firma["token"],
        "__RequestVerificationToken": csrf_token,
    }

    headers = {
        "Referer": referer,
        "Origin": "https://www.btso.org.tr",
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "*/*",
    }

    bekleme = RATE_LIMIT_INITIAL_WAIT

    for deneme in range(
        1,
        MAX_DETAIL_RETRY + 1
    ):

        try:

            response = session.post(
                DETAIL_URL,
                data=payload,
                headers=headers,
                timeout=DETAIL_TIMEOUT,
            )

            # ---------------------------------------------------------------
            # Rate-limit
            # ---------------------------------------------------------------

            if rate_limit_var_mi(
                response
            ):

                print(
                    f"        ! Rate-limit "
                    f"(deneme {deneme}/{MAX_DETAIL_RETRY})"
                )

                if deneme >= MAX_DETAIL_RETRY:
                    break

                print(
                    f"          {bekleme} saniye bekleniyor..."
                )

                time.sleep(
                    bekleme
                )

                bekleme = min(
                    bekleme * 2,
                    RATE_LIMIT_MAX_WAIT
                )

                continue

            # ---------------------------------------------------------------
            # HTTP hata
            # ---------------------------------------------------------------

            if response.status_code != 200:

                print(
                    f"        ! HTTP "
                    f"{response.status_code}"
                )

                time.sleep(
                    min(
                        3 * deneme,
                        15
                    )
                )

                continue

            # ---------------------------------------------------------------
            # BTSO hata HTML'i
            # ---------------------------------------------------------------

            if hata_sayfasi_mi(
                response.text
            ):

                print(
                    f"        ! BTSO hata döndürdü "
                    f"(deneme {deneme}/{MAX_DETAIL_RETRY})"
                )

                if deneme >= MAX_DETAIL_RETRY:
                    break

                print(
                    f"          {bekleme} saniye bekleniyor..."
                )

                time.sleep(
                    bekleme
                )

                bekleme = min(
                    bekleme * 2,
                    RATE_LIMIT_MAX_WAIT
                )

                continue

            # ---------------------------------------------------------------
            # Parse
            # ---------------------------------------------------------------

            detay = parse_detail_html(
                response.text,
                firma["unvan"]
            )

            return detay

        except requests.RequestException as e:

            print(
                f"        ! Request hatası: {e}"
            )

            if deneme < MAX_DETAIL_RETRY:

                time.sleep(
                    min(
                        3 * deneme,
                        15
                    )
                )

    # Başarısız olsa bile boş kayıt oluştur
    return {
        "Unvan": firma["unvan"],
        "Adres": "",
        "İlçe": "",
        "Web": "",
        "Meslek Grubu No": "",
        "Meslek Grubu": "",
    }


# =============================================================================
# EXCEL
# =============================================================================

COLUMNS = [
    "Unvan",
    "Adres",
    "İlçe",
    "Web",
    "Meslek Grubu No",
    "Meslek Grubu",
]


def excel_kaydet(
    results,
    filename
):

    if not results:
        return

    df = pd.DataFrame(
        results
    )

    for column in COLUMNS:

        if column not in df.columns:
            df[column] = ""

    df = df[
        COLUMNS
    ]

    df = df.drop_duplicates(
        subset=[
            "Unvan",
            "Adres",
        ],
        keep="first"
    )

    df.to_excel(
        filename,
        index=False,
        engine="openpyxl"
    )


def excel_oku(
    filename
):

    if not os.path.exists(
        filename
    ):
        return []

    try:

        df = pd.read_excel(
            filename
        )

        for column in COLUMNS:

            if column not in df.columns:
                df[column] = ""

        df = df[
            COLUMNS
        ]

        return df.to_dict(
            orient="records"
        )

    except Exception as e:

        print(
            f"Excel okunamadı: {e}"
        )

        return []


# =============================================================================
# PROGRESS
# =============================================================================

def progress_yolu(grup_no):

    temiz = re.sub(
        r"[^0-9A-Za-zÇĞİÖŞÜçğıöşü]+",
        "_",
        grup_no
    )

    return (
        f"BTSO_{temiz}_progress.json"
    )


def output_yolu(grup_no):

    temiz = re.sub(
        r"[^0-9A-Za-zÇĞİÖŞÜçğıöşü]+",
        "_",
        grup_no
    )

    return (
        f"BTSO_{temiz}.xlsx"
    )


def progress_oku(filename):

    if not os.path.exists(
        filename
    ):
        return {
            "last_completed_page": 0
        }

    try:

        with open(
            filename,
            "r",
            encoding="utf-8"
        ) as f:

            return json.load(f)

    except Exception:

        return {
            "last_completed_page": 0
        }


def progress_kaydet(
    filename,
    last_completed_page,
    total
):

    data = {
        "last_completed_page":
            last_completed_page,

        "total_firma":
            total,

        "updated_at":
            time.strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
    }

    with open(
        filename,
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            data,
            f,
            ensure_ascii=False,
            indent=4
        )


# =============================================================================
# MEVCUT SAYFA
# =============================================================================

def mevcut_sayfa(driver):

    url = driver.current_url

    match = re.search(
        r"/kayitli-uyeler/(\d+)/unvan",
        url
    )

    if match:

        return int(
            match.group(1)
        )

    return 1


def hata_sayfasinda_mi(driver):

    return (
        "/Home/Error"
        in driver.current_url
    )


# =============================================================================
# SAYFANIN YÜKLENDİĞİNİ KONTROL
# =============================================================================

def firma_listesi_bekle(driver):

    WebDriverWait(
        driver,
        PAGE_TIMEOUT
    ).until(
        lambda d:
        len(
            d.find_elements(
                By.CSS_SELECTOR,
                "a.show-details[data-id][data-token]"
            )
        ) > 0
    )


# =============================================================================
# SAYFA NUMARASI LINKİ
# =============================================================================

def sayfa_linki_bul(
    driver,
    hedef_sayfa
):
    """
    Sayfadaki gerçek pagination linkini bulur.

    Doğrudan URL oluşturmuyoruz.
    """

    hedef = str(
        hedef_sayfa
    )

    links = driver.find_elements(
        By.TAG_NAME,
        "a"
    )

    for link in links:

        try:

            if not link.is_displayed():
                continue

            text = temiz_metin(
                link.text
            )

            if text != hedef:
                continue

            href = (
                link.get_attribute(
                    "href"
                )
                or ""
            )

            if (
                "/kayitli-uyeler/"
                in href
                and
                "/unvan/"
                in href
            ):
                return link

        except StaleElementReferenceException:
            continue

    return None


# =============================================================================
# SAYFA DEĞİŞTİR
# =============================================================================

def sonraki_sayfaya_gec(
    driver,
    current_page
):

    hedef = current_page + 1

    link = sayfa_linki_bul(
        driver,
        hedef
    )

    if link is None:

        print(
            f"\n{hedef}. sayfa linki bulunamadı."
        )

        return False

    eski_url = driver.current_url

    try:

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            link
        )

        time.sleep(0.4)

        try:

            link.click()

        except Exception:

            driver.execute_script(
                "arguments[0].click();",
                link
            )

        # URL değişmesini bekle
        WebDriverWait(
            driver,
            PAGE_TIMEOUT
        ).until(
            lambda d:
            d.current_url != eski_url
        )

        # Liste gerçekten geldi mi?
        firma_listesi_bekle(
            driver
        )

        time.sleep(
            random.uniform(
                PAGE_WAIT_MIN,
                PAGE_WAIT_MAX
            )
        )

        return True

    except Exception as e:

        print(
            f"\nSayfa {hedef} geçişi başarısız: {e}"
        )

        return False


# =============================================================================
# ERROR SAYFASINDAN GERİ DÖN
# =============================================================================

def hata_sayfasindan_geri_don(
    driver
):

    if not hata_sayfasinda_mi(
        driver
    ):
        return True

    print(
        "\nBTSO hata sayfasına yönlendirdi."
    )

    print(
        "Önceki sayfaya dönülüyor..."
    )

    try:

        driver.back()

        firma_listesi_bekle(
            driver
        )

        time.sleep(
            random.uniform(
                2,
                4
            )
        )

        return True

    except Exception:

        return False


# =============================================================================
# KALINAN SAYFAYA GİT
# =============================================================================

def resume_sayfasina_git(
    driver,
    hedef_sayfa
):

    if hedef_sayfa <= 1:

        return True

    mevcut = mevcut_sayfa(
        driver
    )

    print(
        f"\nKayıt dosyasına göre "
        f"{hedef_sayfa}. sayfaya dönülüyor..."
    )

    while mevcut < hedef_sayfa:

        print(
            f"  {mevcut} → {mevcut + 1}"
        )

        basarili = False

        for deneme in range(
            1,
            MAX_PAGE_RETRY + 1
        ):

            if hata_sayfasinda_mi(
                driver
            ):

                if not hata_sayfasindan_geri_don(
                    driver
                ):
                    time.sleep(
                        10
                    )
                    continue

            if sonraki_sayfaya_gec(
                driver,
                mevcut
            ):

                basarili = True
                break

            print(
                f"    Deneme {deneme}/"
                f"{MAX_PAGE_RETRY} başarısız."
            )

            # BTSO biraz nefes alsın
            time.sleep(
                min(
                    5 * deneme,
                    30
                )
            )

            if hata_sayfasinda_mi(
                driver
            ):

                hata_sayfasindan_geri_don(
                    driver
                )

        if not basarili:

            print(
                f"\n{mevcut + 1}. sayfaya "
                "ulaşılamadı."
            )

            return False

        mevcut = mevcut_sayfa(
            driver
        )

    return True


# =============================================================================
# ANA PROGRAM
# =============================================================================

def main():

    driver = chrome_baslat()

    results = []

    try:

        # =====================================================================
        # BTSO
        # =====================================================================

        print(
            "\nBTSO açılıyor..."
        )

        driver.get(
            BASE_URL
        )

        print(
            "\n" + "=" * 70
        )

        print(
            "MANUEL AŞAMA"
        )

        print(
            "=" * 70
        )

        print(
            "1. Meslek grubunu seç."
        )

        print(
            "2. CAPTCHA'yı gir."
        )

        print(
            "3. Ara butonuna bas."
        )

        print(
            "4. Firma listesi geldiğinde ENTER'a bas."
        )

        print(
            "=" * 70
        )

        input(
            "\nListe geldiyse ENTER'a basın..."
        )

        # =====================================================================
        # Firma listesi gelmesini bekle
        # =====================================================================

        firma_listesi_bekle(
            driver
        )

        # =====================================================================
        # Grup
        # =====================================================================

        grup_no, grup_adi = (
            grup_bilgisi_al(
                driver
            )
        )

        if not grup_no:

            raise RuntimeError(
                "Meslek grubu bulunamadı."
            )

        print(
            "\nMeslek grubu:"
        )

        print(
            f"  No  : {grup_no}"
        )

        print(
            f"  Adı : {grup_adi}"
        )

        # =====================================================================
        # Dosyalar
        # =====================================================================

        output_file = output_yolu(
            grup_no
        )

        progress_file = progress_yolu(
            grup_no
        )

        print(
            f"\nExcel:"
            f" {output_file}"
        )

        print(
            f"Progress:"
            f" {progress_file}"
        )

        # =====================================================================
        # Önceki kayıt var mı?
        # =====================================================================

        progress = progress_oku(
            progress_file
        )

        last_completed_page = int(
            progress.get(
                "last_completed_page",
                0
            )
        )

        if last_completed_page > 0:

            results = excel_oku(
                output_file
            )

            print(
                "\n" + "=" * 70
            )

            print(
                "ÖNCEKİ İLERLEME BULUNDU"
            )

            print(
                "=" * 70
            )

            print(
                f"Tamamlanan son sayfa: "
                f"{last_completed_page}"
            )

            print(
                f"Mevcut firma sayısı: "
                f"{len(results):,}"
            )

            print(
                "Kaldığı yerden devam edilecek."
            )

            print(
                "=" * 70
            )

            # -----------------------------------------------------------------
            # Örneğin 16 tamamlandıysa 17'ye git
            # -----------------------------------------------------------------

            if not resume_sayfasina_git(
                driver,
                last_completed_page + 1
            ):

                print(
                    "\nKaldığı sayfaya ulaşılamadı."
                )

                print(
                    "İşlem sonlandırıldı."
                )

                return

        # =====================================================================
        # Session
        # =====================================================================

        session = session_olustur(
            driver
        )

        # =====================================================================
        # SAYFA DÖNGÜSÜ
        # =====================================================================

        visited_pages = set()

        while True:

            # Error'a düşmüşse geri dön
            if hata_sayfasinda_mi(
                driver
            ):

                if not hata_sayfasindan_geri_don(
                    driver
                ):

                    print(
                        "\nHata sayfasından dönülemedi."
                    )

                    break

            current_page = mevcut_sayfa(
                driver
            )

            # Aynı sayfa tekrar işlenmesin
            if current_page in visited_pages:

                print(
                    f"\nSayfa {current_page} "
                    "zaten işlendi."
                )

                break

            visited_pages.add(
                current_page
            )

            print(
                "\n" + "=" * 70
            )

            print(
                f"SAYFA {current_page}"
            )

            print(
                "=" * 70
            )

            # =================================================================
            # Session / CSRF güncelle
            # =================================================================

            session_cookie_guncelle(
                session,
                driver
            )

            csrf_token = csrf_token_al(
                driver
            )

            if not csrf_token:

                print(
                    "CSRF token bulunamadı."
                )

                print(
                    "Sayfa yeniden okunmaya çalışılıyor..."
                )

                driver.refresh()

                time.sleep(
                    3
                )

                csrf_token = csrf_token_al(
                    driver
                )

            if not csrf_token:

                print(
                    "CSRF token alınamadı."
                )

                print(
                    "İşlem sonlandırılıyor."
                )

                break

            # =================================================================
            # Firmalar
            # =================================================================

            firmalar = firma_listesini_oku(
                driver
            )

            print(
                f"Bu sayfada "
                f"{len(firmalar)} firma bulundu."
            )

            if not firmalar:

                print(
                    "Bu sayfada firma bulunamadı."
                )

                break

            # =================================================================
            # Detail POST
            # =================================================================

            referer = driver.current_url

            # Önceden alınmış unvanları set yap
            mevcut_unvanlar = {
                temiz_metin(
                    row.get("Unvan", "")
                )
                for row in results
            }

            for index, firma in enumerate(
                firmalar,
                start=1
            ):

                unvan = temiz_metin(
                    firma["unvan"]
                )

                # -----------------------------------------------------------------
                # Daha önce alınmışsa atla
                # -----------------------------------------------------------------

                if unvan in mevcut_unvanlar:

                    print(
                        f"    [{index}/{len(firmalar)}] "
                        f"Zaten mevcut → {unvan}"
                    )

                    continue

                print(
                    f"    [{index}/{len(firmalar)}] "
                    f"{unvan}"
                )

                detay = firma_detayi_getir(
                    session=session,
                    firma=firma,
                    csrf_token=csrf_token,
                    referer=referer,
                )

                # Grup detail'da bulunamadıysa
                if not detay["Meslek Grubu No"]:

                    detay[
                        "Meslek Grubu No"
                    ] = grup_no

                if not detay["Meslek Grubu"]:

                    detay[
                        "Meslek Grubu"
                    ] = grup_adi

                results.append(
                    detay
                )

                mevcut_unvanlar.add(
                    unvan
                )

                print(
                    f"        "
                    f"İlçe: "
                    f"{detay['İlçe'] or '-'}"
                    f" | Web: "
                    f"{detay['Web'] or '-'}"
                )

                # Normal çok kısa bekleme
                # Rate-limit olursa fonksiyon zaten daha uzun bekler.
                if index < len(firmalar):

                    time.sleep(
                        random.uniform(
                            DETAIL_WAIT_MIN,
                            DETAIL_WAIT_MAX
                        )
                    )

            # =================================================================
            # SAYFA TAMAMLANDI
            # =================================================================

            excel_kaydet(
                results,
                output_file
            )

            progress_kaydet(
                progress_file,
                current_page,
                len(results)
            )

            print(
                "\n  ✓ Sayfa tamamlandı."
            )

            print(
                f"  ✓ Toplam firma: "
                f"{len(results):,}"
            )

            print(
                f"  ✓ Progress: "
                f"sayfa {current_page}"
            )

            # =================================================================
            # SONRAKİ SAYFA
            # =================================================================

            next_page = (
                current_page + 1
            )

            next_link = (
                sayfa_linki_bul(
                    driver,
                    next_page
                )
            )

            # Sonraki sayfa yoksa tamamlandı
            if next_link is None:

                print(
                    "\nSonraki sayfa bulunamadı."
                )

                print(
                    "Son sayfaya ulaşılmış olabilir."
                )

                break

            print(
                f"\n{next_page}. sayfaya geçiliyor..."
            )

            basarili = False

            for deneme in range(
                1,
                MAX_PAGE_RETRY + 1
            ):

                # Eğer hata sayfasına düştüysek geri dön
                if hata_sayfasinda_mi(
                    driver
                ):

                    if not hata_sayfasindan_geri_don(
                        driver
                    ):

                        time.sleep(
                            10
                        )

                        continue

                # Gerçek pagination linkini tekrar bul
                next_link = (
                    sayfa_linki_bul(
                        driver,
                        next_page
                    )
                )

                if next_link is None:

                    print(
                        f"    {next_page}. sayfa linki "
                        "bulunamadı."
                    )

                    time.sleep(
                        5 * deneme
                    )

                    continue

                if sonraki_sayfaya_gec(
                    driver,
                    current_page
                ):

                    basarili = True
                    break

                print(
                    f"    Sayfa geçişi başarısız "
                    f"(deneme {deneme}/{MAX_PAGE_RETRY})"
                )

                wait_time = min(
                    5 * deneme,
                    30
                )

                print(
                    f"    {wait_time} saniye bekleniyor..."
                )

                time.sleep(
                    wait_time
                )

                if hata_sayfasinda_mi(
                    driver
                ):

                    hata_sayfasindan_geri_don(
                        driver
                    )

            if not basarili:

                print(
                    "\n" + "=" * 70
                )

                print(
                    "İŞLEM GEÇİCİ OLARAK DURDURULDU"
                )

                print(
                    "=" * 70
                )

                print(
                    f"Son tamamlanan sayfa: "
                    f"{current_page}"
                )

                print(
                    f"Sonraki hedef: "
                    f"{next_page}"
                )

                print(
                    "Tekrar çalıştırırsanız "
                    f"{next_page}. sayfadan devam edecek."
                )

                print(
                    "=" * 70
                )

                break

        # =====================================================================
        # FINAL
        # =====================================================================

        excel_kaydet(
            results,
            output_file
        )

        print(
            "\n" + "=" * 70
        )

        print(
            "İŞLEM TAMAMLANDI / DURDURULDU"
        )

        print(
            "=" * 70
        )

        print(
            f"Toplam firma: "
            f"{len(results):,}"
        )

        print(
            f"Excel: "
            f"{output_file}"
        )

        print(
            "=" * 70
        )

    except KeyboardInterrupt:

        print(
            "\n\nKullanıcı tarafından durduruldu."
        )

        if results:

            excel_kaydet(
                results,
                output_file
            )

        print(
            "Mevcut veriler Excel'e kaydedildi."
        )

    except Exception as e:

        print(
            "\nBeklenmeyen hata:"
        )

        print(
            repr(e)
        )

        try:

            if results:

                excel_kaydet(
                    results,
                    output_file
                )

                print(
                    f"Toplanan {len(results):,} "
                    "firma kaydedildi."
                )

        except Exception as save_error:

            print(
                "Excel kaydedilirken de hata oluştu:"
            )

            print(
                repr(save_error)
            )

    finally:

        print(
            "\nChrome açık bırakıldı."
        )


# =============================================================================
# ÇALIŞTIR
# =============================================================================

if __name__ == "__main__":
    main()