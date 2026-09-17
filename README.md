# Şirket Web Sitesi ve E-posta Bulucu

Bu program Excel dosyasındaki firma bilgilerini kullanarak şirketlerin web sitelerini ve e-posta adreslerini bulup doğrulayan Python tabanlı bir araçtır.

Google Search üzerinden web sitesi adaylarını bulur, firma adı ve adres bilgilerine göre değerlendirir, Selenium ile siteleri kontrol eder ve bulduğu web sitesi ile e-posta bilgilerini Excel dosyasına kaydeder.  

Farklı şehirlerde kullanılabilir. Şehir ve ilçe listeleri `ilceler.json` dosyasında tutulur ve program başlatıldığında yalnızca işlem yapılacak şehrin seçilmesi yeterlidir.

## Kullanılan Teknolojiler

- Python 3
- Selenium
- Pandas
- OpenPyXL
- WebDriver Manager
- Google Search
- Web Scraping
- Veri Çıkarma
- Otomasyon

## Girdi

Programın çalışması için proje klasöründe `firmalar.xlsx` adlı Excel dosyasının bulunması gerekir.

Dosyada şu üç sütun yer almalıdır:

`UNVAN`, `ADRES`, `ILCE`

Sütun isimlerinde yazım hatası, eksik karakter veya farklı adlandırma bulunmamalıdır.  

Örnek:

| UNVAN | ADRES | ILCE |
|---|---|---|
| Örnek Ltd. Şti. | Örnek Mah. Örnek Sok. | Şişli |

### İstanbul İlçe Bilgisi

İstanbul için hazırlanan ve `ILCE` bilgisi bulunmayan firma listelerinde, adreslerden ilçe bilgisini oluşturmak için `adrestenilce.py` programı kullanılabilir. Programın çalışabilmesi için `istanbul_ilce_verileri.json` dosyasının proje klasöründe bulunması gerekir.

## Şehir ve İlçe Ayarları

Şehir ve ilçe bilgileri `ilceler.json` dosyasından okunur.  

Program çalıştırıldığında mevcut şehirler arasından işlem yapılacak şehir seçilir. Seçilen şehre ait ilçeler otomatik olarak yüklenir.

Başka şehirlerde çalışabilmek için yalnızca `ilceler.json` dosyasını güncellemeniz yeterlidir. Python dosyasında değişiklik yapmanız gerekmez.

## Çıktı

Program, sonuçları `firmalar_web_mail.xlsx` dosyasına kaydeder.  

Çıktı dosyasında aşağıdaki sütunlar bulunur:

`UNVAN`, `KAYNAK_ADRES`, `KAYNAK_ILCE`, `WEB`, `MAIL`, `WEB_ILCE`, `ADRES_DURUMU`, `DURUM`, `SITE_PUANI`. 

Excel dosyasında bazı satırlar **turuncu renkle** işaretlenir. Bu satırlar, web sitesi bulunmuş olsa da ilçe veya adres bilgilerinde farklılık bulunduğu ya da adres bilgilerinin eksik olduğu durumları gösterir ve **manuel olarak tekrar kontrol edilmelidir**.

**Beyaz renkteki satırlar**, otomatik doğrulama kriterlerini karşılayan ve firma ile web sitesi arasında güçlü eşleşme bulunan sonuçlardır.

## İlerleme Kaydı

Program, işlem sırasında ilerlemeyi `sitemailbul_ilerleme.json` dosyasına otomatik olarak kaydeder. Program yarıda kesilirse, sonraki çalıştırmada tamamlanan firmalar atlanarak kaldığı yerden devam edilebilir.

Kaynak `firmalar.xlsx` dosyası değiştirildiğinde mevcut ilerleme kaydı geçersiz kabul edilir ve işlem yeniden başlatılır.

## Proje Yapısı

```text
firmalar/
├── venv/
├── firmalar.xlsx
├── firmalar_web_mail.xlsx
├── ilceler.json
└── sitemailbul.py

Çalışma sırasında oluşabilir:
└── sitemailbul_ilerleme.json
```

## Kurulum

Proje Python 3.14.3 ile geliştirilmiştir.  

Projeyi indirdikten veya klonladıktan sonra proje klasörüne geçin ve sanal ortam oluşturun:

- ### macOS Terminal

     ```bash
     cd ~/Desktop/firmalar
     python3 -m venv venv
     source venv/bin/activate
     ```

- ### Windows PowerShell

     ```bash
     cd C:\firmalar
     python -m venv venv
     venv\Scripts\Activate
     ```

Gerekli paketleri yükleyin:

- ### macOS

     ```bash
     python3 -m pip install pandas openpyxl selenium webdriver-manager
     ```

- ### Windows

     ```bash
     python -m pip install pandas openpyxl selenium webdriver-manager
     ```

## Chrome Uzaktan Hata Ayıklama

Program, Selenium ile açık bir Chrome oturumuna bağlanarak çalışır. Bu nedenle Chrome'u uzaktan hata ayıklama özelliği etkin olacak şekilde başlatmanız gerekir.

- ### macOS

  Terminal'i açın ve aşağıdaki komutu çalıştırın:

     ```bash
     "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --remote-debugging-port=9222 --user-data-dir="$HOME/chrome_selenium"
     ```

  Program çalışırken açılan Chrome penceresini kapatmayın.

- ### Windows

  PowerShell'i açın ve aşağıdaki komutu çalıştırın:

     ```bash
     & "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="$env:USERPROFILE\chrome_selenium"
     ```

Chrome farklı bir klasöre kurulmuşsa komuttaki dosya yolunu kendi sisteminize göre güncelleyin.

### Google CAPTCHA veya Chrome Profil Sorunları

Google aramalarında CAPTCHA ile karşılaşılması veya Selenium ile Chrome bağlantısında beklenmeyen bir sorun oluşması durumunda, programın kullandığı ayrı Chrome profilini sıfırlayabilirsiniz.

- ### macOS

     ```bash
     rm -rf "$HOME/chrome_selenium"
     ```

- ### Windows PowerShell

     ```bash
     Remove-Item -Recurse -Force "$env:USERPROFILE\chrome_selenium"
     ```
Ardından Chrome'u tamamen kapatıp, **Chrome Uzaktan Hata Ayıklama** bölümündeki komutu tekrar çalıştırın ve programı yeniden başlatın.

> `chrome_selenium` yalnızca program için kullanılan ayrı Chrome profilidir. Bu klasörün silinmesi normal Chrome profilinizi etkilemez.  

## Programı Çalıştırma

Öncelikle proje klasörüne geçin ve sanal ortamı etkinleştirin.  

macOS'ta Chrome Uzaktan Hata Ayıklama ve program için ayrı Terminal pencereleri kullanılmalıdır. Windows'ta ise Chrome ve program aynı PowerShell penceresinden çalıştırılabilir.  

- ### macOS

    ```bash
     cd ~/Desktop/firmalar
     source venv/bin/activate
     python3 sitemailbul.py
     ```
  
- ### Windows
  
    ```bash
     cd C:\firmalar
     venv\Scripts\Activate
     python sitemailbul.py
     ```
    
Proje klasörünün yolu bilgisayarınızdaki konuma göre değişebilir. 

> `firmalar.xlsx`, `firmalar_web_mail.xlsx` ve `venv/` yerel veri veya sanal ortam dosyaları içerdiği için public repository'ye eklenmemelidir.
