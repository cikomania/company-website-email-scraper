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
from selenium.common.exceptions import StaleElementReferenceException
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

# Firma detay istekleri arasındaki bekleme
DETAIL_WAIT_MIN = 0.5
DETAIL_WAIT_MAX = 0.9

# Sayfa değişiminden sonra bekleme
PAGE_WAIT_MIN = 1.0
PAGE_WAIT_MAX = 1.8

# Rate-limit ilk bekleme
RATE_LIMIT_INITIAL_WAIT = 8

# Rate-limit maksimum bekleme
RATE_LIMIT_MAX_WAIT = 60

# Aynı firma için maksimum deneme
MAX_DETAIL_RETRY = 5

# Aynı sayfaya geçiş için maksimum deneme
MAX_PAGE_RETRY = 5

# Detail POST timeout
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
# EXCEL SÜTUNLARI
# =============================================================================

COLUMNS = [
    "Unvan",
    "Adres",
    "İlçe",
    "Web",
    "Meslek Grubu No",
    "Meslek Grubu",
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

    Nilüfer -> NİLÜFER
    nilüfer -> NİLÜFER
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
    Adres içerisindeki:

        İLÇE/BURSA
        İlçe / Bursa
        ilçe/BURSA

    yapısını bulur.
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
# MESLEK GRUBU AYRIŞTIR
# =============================================================================

def meslek_grubu_ayristir(meslek):
    """
    Örnek:

        04. GRUP : MADENLER VE...

    sonuç:

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
    # UNVAN
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
    # TABLO SATIRLARI
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

        alan_norm = turkce_upper(alan)

        # ADRES
        if alan_norm == "ADRES":

            adres = deger

        # WEB
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

        # MESLEK GRUBU
        elif alan_norm == "MESLEK GRUBU":

            meslek = deger

    grup_no, grup_adi = meslek_grubu_ayristir(
        meslek
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

    # Chrome açık kalsın
    options.add_experimental_option(
        "detach",
        True
    )

    service = Service(
        ChromeDriverManager().install()
    )

    driver = webdriver.Chrome(
        service=service,
        options=options
    )

    return driver


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

    # Selenium cookie'lerini requests'e aktar
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
# MESLEK GRUBU BİLGİSİ
# =============================================================================

def grup_bilgisi_al(driver):

    grup_no = ""
    grup_adi = ""

    # -------------------------------------------------------------------------
    # Önce SELECT
    # -------------------------------------------------------------------------

    try:

        selected = driver.find_element(
            By.CSS_SELECTOR,
            "#SelectedSector option:checked"
        )

        secili = temiz_metin(
            selected.text
        )

        grup_no, grup_adi = meslek_grubu_ayristir(
            secili
        )

    except Exception:
        pass

    # -------------------------------------------------------------------------
    # URL FALLBACK
    # -------------------------------------------------------------------------

    if not grup_no:

        try:

            url = driver.current_url

            match = re.search(
                r"/unvan/(.+)$",
                url
            )

            if match:

                grup_no, grup_adi = meslek_grubu_ayristir(
                    match.group(1)
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
                element.get_attribute("data-id")
                or ""
            ).strip()

            token = (
                element.get_attribute("data-token")
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
# RATE LIMIT
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


def hata_html_mi(html):

    text = turkce_upper(
        html
    )

    ifadeler = [
        "İŞLEM BAŞARISIZ OLDU",
        "BEKLENMEYEN BİR HATA",
        "TOO MANY REQUESTS",
    ]

    return any(
        ifade in text
        for ifade in ifadeler
    )


# =============================================================================
# FİRMA DETAY POST
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

            # -----------------------------------------------------------------
            # RATE LIMIT
            # -----------------------------------------------------------------

            if rate_limit_var_mi(response):

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

            # -----------------------------------------------------------------
            # HTTP HATA
            # -----------------------------------------------------------------

            if response.status_code != 200:

                print(
                    f"        ! HTTP "
                    f"{response.status_code}"
                )

                if deneme < MAX_DETAIL_RETRY:

                    time.sleep(
                        min(
                            3 * deneme,
                            15
                        )
                    )

                continue

            # -----------------------------------------------------------------
            # BTSO HATA HTML
            # -----------------------------------------------------------------

            if hata_html_mi(
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

            # -----------------------------------------------------------------
            # PARSE
            # -----------------------------------------------------------------

            return parse_detail_html(
                response.text,
                firma["unvan"]
            )

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

    # Başarısız olsa bile unvan kaybolmasın
    return {
        "Unvan": firma["unvan"],
        "Adres": "",
        "İlçe": "",
        "Web": "",
        "Meslek Grubu No": "",
        "Meslek Grubu": "",
    }


# =============================================================================
# EXCEL KAYDET
# =============================================================================

def excel_kaydet(
    results,
    filename
):

    if not results:
        return

    df = pd.DataFrame(
        results
    )

    # Eksik sütunları oluştur
    for column in COLUMNS:

        if column not in df.columns:
            df[column] = ""

    # Sıralama
    df = df[
        COLUMNS
    ]

    # Aynı firma tekrar etmesin
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


# =============================================================================
# EXCEL OKU
# =============================================================================

def excel_oku(filename):

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
# DOSYA YOLLARI
# =============================================================================

def dosya_temizle(grup_no):

    return re.sub(
        r"[^0-9A-Za-zÇĞİÖŞÜçğıöşü]+",
        "_",
        grup_no
    )


def progress_yolu(grup_no):

    temiz = dosya_temizle(
        grup_no
    )

    return (
        f"BTSO_{temiz}_progress.json"
    )


def output_yolu(grup_no):

    temiz = dosya_temizle(
        grup_no
    )

    return (
        f"BTSO_{temiz}.xlsx"
    )


# =============================================================================
# PROGRESS OKU
# =============================================================================

def progress_oku(filename):

    if not os.path.exists(
        filename
    ):

        return {
            "last_completed_page": 0,
            "total_firma": 0,
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
            "last_completed_page": 0,
            "total_firma": 0,
        }


# =============================================================================
# PROGRESS KAYDET
# =============================================================================

def progress_kaydet(
    filename,
    last_completed_page,
    total
):

    data = {
        "last_completed_page": last_completed_page,
        "total_firma": total,
        "updated_at": time.strftime(
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


# =============================================================================
# ERROR SAYFASI
# =============================================================================

def hata_sayfasinda_mi(driver):

    return (
        "/Home/Error"
        in driver.current_url
    )


# =============================================================================
# FİRMA LİSTESİ BEKLE
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
# SAYFA LINKİ BUL
# =============================================================================

def sayfa_linki_bul(
    driver,
    hedef_sayfa
):
    """
    Önce doğrudan hedef sayfa numarasını arar.

    Örneğin:
        47

    linki varsa onu döndürür.
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
                link.get_attribute("href")
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
# SONRAKİ SAYFA LINKİNİ BUL
# =============================================================================

def sonraki_linki_bul(
    driver,
    current_page
):
    """
    BTSO pagination yapısında:

        Sonraki
        Next
        >
        »
        ›
        veya doğrudan current+1

    gibi seçenekleri arar.
    """

    hedef_sayfa = current_page + 1

    # -------------------------------------------------------------------------
    # 1. Önce doğrudan hedef numara
    # -------------------------------------------------------------------------

    link = sayfa_linki_bul(
        driver,
        hedef_sayfa
    )

    if link is not None:
        return link

    # -------------------------------------------------------------------------
    # 2. Sonraki / Next / > / » / ›
    # -------------------------------------------------------------------------

    links = driver.find_elements(
        By.TAG_NAME,
        "a"
    )

    sonraki_textler = {
        "SONRAKİ",
        "SONRAKI",
        "NEXT",
        ">",
        ">>",
        "›",
        "»",
        "→",
    }

    for link in links:

        try:

            if not link.is_displayed():
                continue

            text = temiz_metin(
                link.text
            ).upper()

            aria = temiz_metin(
                link.get_attribute("aria-label")
                or ""
            ).upper()

            title = temiz_metin(
                link.get_attribute("title")
                or ""
            ).upper()

            if (
                text in sonraki_textler
                or aria in sonraki_textler
                or title in sonraki_textler
                or "SONRAKİ" in aria
                or "SONRAKI" in aria
                or "NEXT" in aria
            ):

                href = (
                    link.get_attribute("href")
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

    link = sonraki_linki_bul(
        driver,
        current_page
    )

    if link is None:

        print(
            f"\n{hedef}. sayfaya ait "
            "pagination linki bulunamadı."
        )

        return False

    eski_url = driver.current_url

    try:

        driver.execute_script(
            """
            arguments[0].scrollIntoView({
                block: 'center'
            });
            """,
            link
        )

        time.sleep(
            0.4
        )

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

        # Yeni firma listesi gelsin
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
            f"\nSayfa {hedef} geçişi başarısız:"
        )

        print(
            f"    {e}"
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
# ANA PROGRAM
# =============================================================================

def main():

    driver = chrome_baslat()

    results = []

    output_file = ""
    progress_file = ""

    try:

        # =====================================================================
        # BTSO'YU AÇ
        # =====================================================================

        print(
            "\nBTSO açılıyor..."
        )

        driver.get(
            BASE_URL
        )

        # =====================================================================
        # MANUEL AŞAMA
        # =====================================================================

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
            "2. CAPTCHA / doğrulama kodunu gir."
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
        # FİRMA LİSTESİ
        # =====================================================================

        firma_listesi_bekle(
            driver
        )

        # =====================================================================
        # MESLEK GRUBU
        # =====================================================================

        grup_no, grup_adi = grup_bilgisi_al(
            driver
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
        # DOSYALAR
        # =====================================================================

        output_file = output_yolu(
            grup_no
        )

        progress_file = progress_yolu(
            grup_no
        )

        print(
            "\nExcel:"
        )

        print(
            f"  {output_file}"
        )

        print(
            "\nProgress:"
        )

        print(
            f"  {progress_file}"
        )

        # =====================================================================
        # ÖNCEKİ İLERLEME
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

        # ---------------------------------------------------------------------
        # DAHA ÖNCE KAYIT VARSA
        # ---------------------------------------------------------------------

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
                f"Tamamlanan son sayfa : "
                f"{last_completed_page}"
            )

            print(
                f"Mevcut firma sayısı  : "
                f"{len(results):,}"
            )

            print(
                "=" * 70
            )

            print(
                "\nŞimdi Chrome'da devam edeceğiniz "
                "sayfayı MANUEL olarak açın."
            )

            print(
                f"Önerilen sayfa: "
                f"{last_completed_page + 1}"
            )

            print(
                "\nÖrneğin son kayıt 100 ise:"
            )

            print(
                "Chrome'da 101. sayfayı açın."
            )

            print(
                "Firma listesi göründüğünde ENTER'a basın."
            )

            input(
                "\nHazır olduğunuzda ENTER..."
            )

        else:

            print(
                "\nİlk çalışma."
            )

            print(
                "Mevcut sayfadan başlanacak."
            )

        # =====================================================================
        # MEVCUT SAYFAYI BELİRLE
        # =====================================================================

        current_page = mevcut_sayfa(
            driver
        )

        print(
            "\n" + "=" * 70
        )

        print(
            f"BAŞLANGIÇ SAYFASI: {current_page}"
        )

        print(
            "=" * 70
        )

        # ---------------------------------------------------------------------
        # UYARI
        # ---------------------------------------------------------------------

        if last_completed_page > 0:

            beklenen = last_completed_page + 1

            if current_page != beklenen:

                print(
                    "\nUYARI!"
                )

                print(
                    f"Excel'e göre beklenen sayfa: "
                    f"{beklenen}"
                )

                print(
                    f"Chrome'daki mevcut sayfa: "
                    f"{current_page}"
                )

                print(
                    "Program Chrome'daki mevcut "
                    "sayfadan devam edecek."
                )

        # =====================================================================
        # SESSION
        # =====================================================================

        session = session_olustur(
            driver
        )

        # =====================================================================
        # SAYFA DÖNGÜSÜ
        # =====================================================================

        visited_pages = set()

        while True:

            # -----------------------------------------------------------------
            # ERROR SAYFASI
            # -----------------------------------------------------------------

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

            # -----------------------------------------------------------------
            # MEVCUT SAYFA
            # -----------------------------------------------------------------

            current_page = mevcut_sayfa(
                driver
            )

            # -----------------------------------------------------------------
            # AYNI SAYFA KONTROLÜ
            # -----------------------------------------------------------------

            if current_page in visited_pages:

                print(
                    f"\nSayfa {current_page} "
                    "zaten işlendi."
                )

                break

            visited_pages.add(
                current_page
            )

            # -----------------------------------------------------------------
            # BAŞLIK
            # -----------------------------------------------------------------

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
            # SESSION / COOKIE
            # =================================================================

            session_cookie_guncelle(
                session,
                driver
            )

            # =================================================================
            # CSRF
            # =================================================================

            csrf_token = csrf_token_al(
                driver
            )

            if not csrf_token:

                print(
                    "\nCSRF token bulunamadı."
                )

                print(
                    "Sayfa yenileniyor..."
                )

                driver.refresh()

                time.sleep(
                    3
                )

                firma_listesi_bekle(
                    driver
                )

                csrf_token = csrf_token_al(
                    driver
                )

            if not csrf_token:

                print(
                    "\nCSRF token alınamadı."
                )

                print(
                    "İşlem durduruldu."
                )

                break

            # =================================================================
            # FİRMA LİSTESİ
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
                    "\nBu sayfada firma bulunamadı."
                )

                print(
                    "İşlem durduruldu."
                )

                break

            # =================================================================
            # DETAIL POST
            # =================================================================

            referer = driver.current_url

            # Daha önce alınmış unvanlar
            mevcut_unvanlar = {
                temiz_metin(
                    row.get(
                        "Unvan",
                        ""
                    )
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

                # -------------------------------------------------------------
                # ZATEN VARSA ATLA
                # -------------------------------------------------------------

                if unvan in mevcut_unvanlar:

                    print(
                        f"    [{index}/{len(firmalar)}] "
                        f"Zaten mevcut → {unvan}"
                    )

                    continue

                # -------------------------------------------------------------
                # FİRMA
                # -------------------------------------------------------------

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

                # -------------------------------------------------------------
                # GRUP BİLGİSİ BOŞSA ANA GRUPTAN DOLDUR
                # -------------------------------------------------------------

                if not detay["Meslek Grubu No"]:

                    detay[
                        "Meslek Grubu No"
                    ] = grup_no

                if not detay["Meslek Grubu"]:

                    detay[
                        "Meslek Grubu"
                    ] = grup_adi

                # -------------------------------------------------------------
                # SONUCA EKLE
                # -------------------------------------------------------------

                results.append(
                    detay
                )

                mevcut_unvanlar.add(
                    unvan
                )

                print(
                    f"        İlçe: "
                    f"{detay['İlçe'] or '-'}"
                    f" | Web: "
                    f"{detay['Web'] or '-'}"
                )

                # -------------------------------------------------------------
                # NORMAL BEKLEME
                # -------------------------------------------------------------

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
                "\n" + "-" * 70
            )

            print(
                "✓ SAYFA TAMAMLANDI"
            )

            print(
                f"✓ Sayfa       : {current_page}"
            )

            print(
                f"✓ Toplam firma: {len(results):,}"
            )

            print(
                f"✓ Excel       : {output_file}"
            )

            print(
                "-" * 70
            )

            # =================================================================
            # SONRAKİ SAYFA
            # =================================================================

            next_page = current_page + 1

            print(
                f"\n{next_page}. sayfaya geçiliyor..."
            )

            basarili = False

            for deneme in range(
                1,
                MAX_PAGE_RETRY + 1
            ):

                # -------------------------------------------------------------
                # ERROR SAYFASI
                # -------------------------------------------------------------

                if hata_sayfasinda_mi(
                    driver
                ):

                    if not hata_sayfasindan_geri_don(
                        driver
                    ):

                        print(
                            "Hata sayfasından "
                            "dönülemedi."
                        )

                        time.sleep(
                            10
                        )

                        continue

                # -------------------------------------------------------------
                # SAYFAYA GEÇ
                # -------------------------------------------------------------

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

            # =================================================================
            # SAYFA GEÇİŞİ BAŞARISIZ
            # =================================================================

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
                    "\nExcel ve progress kaydedildi."
                )

                print(
                    "Tekrar çalıştırdığınızda "
                    "Chrome'da sonraki sayfayı "
                    "manuel açabilirsiniz."
                )

                print(
                    "=" * 70
                )

                break

        # =====================================================================
        # FINAL EXCEL
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

    # =========================================================================
    # CTRL + C
    # =========================================================================

    except KeyboardInterrupt:

        print(
            "\n\nKullanıcı tarafından durduruldu."
        )

        if results and output_file:

            excel_kaydet(
                results,
                output_file
            )

            print(
                "Mevcut veriler Excel'e kaydedildi."
            )

    # =========================================================================
    # BEKLENMEYEN HATA
    # =========================================================================

    except Exception as e:

        print(
            "\nBeklenmeyen hata:"
        )

        print(
            repr(e)
        )

        try:

            if results and output_file:

                excel_kaydet(
                    results,
                    output_file
                )

                print(
                    f"Toplanan {len(results):,} "
                    "firma Excel'e kaydedildi."
                )

        except Exception as save_error:

            print(
                "\nExcel kaydedilirken de hata oluştu:"
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
