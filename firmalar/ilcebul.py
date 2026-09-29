import pandas as pd
import re

# =========================================================
# AYARLAR
# =========================================================

INPUT_FILE = "firmalar.xlsx"
OUTPUT_FILE = "firmalar.xlsx"


# =========================================================
# TÜRKİYE İLÇELERİ
# =========================================================

ILCELER = [

    # ADANA
    "Aladağ", "Ceyhan", "Çukurova", "Feke", "İmamoğlu",
    "Karaisalı", "Karataş", "Kozan", "Pozantı", "Saimbeyli",
    "Sarıçam", "Seyhan", "Tufanbeyli", "Yumurtalık", "Yüreğir",

    # ADIYAMAN
    "Besni", "Çelikhan", "Gerger", "Gölbaşı", "Kahta",
    "Merkez", "Samsat", "Sincik", "Tut",

    # AFYONKARAHİSAR
    "Başmakçı", "Bayat", "Bolvadin", "Çay", "Çobanlar",
    "Dazkırı", "Dinar", "Emirdağ", "Evciler", "Hocalar",
    "İhsaniye", "İscehisar", "Kızılören", "Merkez",
    "Sandıklı", "Sinanpaşa", "Sultandağı", "Şuhut",

    # AĞRI
    "Diyadin", "Doğubayazıt", "Eleşkirt", "Hamur",
    "Merkez", "Patnos", "Taşlıçay", "Tutak",

    # AKSARAY
    "Ağaçören", "Eskil", "Gülağaç", "Güzelyurt",
    "Merkez", "Ortaköy", "Sarıyahşi", "Sultanhanı",

    # AMASYA
    "Göynücek", "Gümüşhacıköy", "Hamamözü", "Merkez",
    "Merzifon", "Suluova", "Taşova",

    # ANKARA
    "Akyurt", "Altındağ", "Ayaş", "Bala", "Beypazarı",
    "Çamlıdere", "Çankaya", "Çubuk", "Elmadağ", "Etimesgut",
    "Evren", "Gölbaşı", "Güdül", "Haymana", "Kahramankazan",
    "Kalecik", "Keçiören", "Kızılcahamam", "Mamak", "Nallıhan",
    "Polatlı", "Pursaklar", "Sincan", "Şereflikoçhisar",
    "Yenimahalle",

    # ANTALYA
    "Akseki", "Aksu", "Alanya", "Demre", "Döşemealtı",
    "Elmalı", "Finike", "Gazipaşa", "Gündoğmuş", "İbradı",
    "Kaş", "Kemer", "Kepez", "Konyaaltı", "Korkuteli",
    "Kumluca", "Manavgat", "Muratpaşa", "Serik",

    # ARDAHAN
    "Çıldır", "Damal", "Göle", "Hanak", "Merkez", "Posof",

    # ARTVİN
    "Ardanuç", "Arhavi", "Borçka", "Hopa", "Kemalpaşa",
    "Merkez", "Murgul", "Şavşat", "Yusufeli",

    # AYDIN
    "Bozdoğan", "Buharkent", "Çine", "Didim", "Efeler",
    "Germencik", "İncirliova", "Karacasu", "Karpuzlu",
    "Koçarlı", "Köşk", "Kuşadası", "Kuyucak", "Nazilli",
    "Söke", "Sultanhisar", "Yenipazar",

    # BALIKESİR
    "Altıeylül", "Ayvalık", "Balya", "Bandırma", "Bigadiç",
    "Burhaniye", "Dursunbey", "Edremit", "Erdek", "Gömeç",
    "Gönen", "Havran", "İvrindi", "Karesi", "Kepsut",
    "Manyas", "Marmara", "Savaştepe", "Sındırgı", "Susurluk",

    # BARTIN
    "Amasra", "Kurucaşile", "Merkez", "Ulus",

    # BATMAN
    "Beşiri", "Gercüş", "Hasankeyf", "Kozluk",
    "Merkez", "Sason",

    # BAYBURT
    "Aydıntepe", "Demirözü", "Merkez",

    # BİLECİK
    "Bozüyük", "Gölpazarı", "İnhisar", "Merkez",
    "Osmaneli", "Pazaryeri", "Söğüt", "Yenipazar",

    # BİNGÖL
    "Adaklı", "Genç", "Karlıova", "Kiğı", "Merkez",
    "Solhan", "Yayladere", "Yedisu",

    # BİTLİS
    "Adilcevaz", "Ahlat", "Güroymak", "Hizan",
    "Merkez", "Mutki", "Tatvan",

    # BOLU
    "Dörtdivan", "Gerede", "Göynük", "Kıbrıscık",
    "Mengen", "Merkez", "Mudurnu", "Seben", "Yeniçağa",

    # BURDUR
    "Ağlasun", "Altınyayla", "Bucak", "Çavdır", "Çeltikçi",
    "Gölhisar", "Karamanlı", "Kemer", "Merkez", "Tefenni",
    "Yeşilova",

    # BURSA
    "Büyükorhan", "Gemlik", "Gürsu", "Harmancık", "İnegöl",
    "İznik", "Karacabey", "Keles", "Kestel", "Mudanya",
    "Mustafakemalpaşa", "Nilüfer", "Orhaneli", "Orhangazi",
    "Osmangazi", "Yenişehir", "Yıldırım",

    # ÇANAKKALE
    "Ayvacık", "Bayramiç", "Biga", "Bozcaada", "Çan",
    "Eceabat", "Ezine", "Gelibolu", "Gökçeada", "Lapseki",
    "Merkez", "Yenice",

    # ÇANKIRI
    "Atkaracalar", "Bayramören", "Çerkeş", "Eldivan",
    "Ilgaz", "Kızılırmak", "Korgun", "Kurşunlu", "Merkez",
    "Orta", "Şabanözü", "Yapraklı",

    # ÇORUM
    "Alaca", "Bayat", "Boğazkale", "Dodurga", "İskilip",
    "Kargı", "Laçin", "Mecitözü", "Merkez", "Oğuzlar",
    "Ortaköy", "Osmancık", "Sungurlu", "Uğurludağ",

    # DENİZLİ
    "Acıpayam", "Babadağ", "Baklan", "Bekilli", "Beyağaç",
    "Bozkurt", "Buldan", "Çal", "Çameli", "Çardak",
    "Çivril", "Güney", "Honaz", "Kale", "Merkezefendi",
    "Pamukkale", "Sarayköy", "Serinhisar", "Tavas",

    # DİYARBAKIR
    "Bağlar", "Bismil", "Çermik", "Çınar", "Çüngüş",
    "Dicle", "Eğil", "Ergani", "Hani", "Hazro",
    "Kayapınar", "Kocaköy", "Kulp", "Lice", "Silvan",
    "Sur", "Yenişehir",

    # DÜZCE
    "Akçakoca", "Çilimli", "Cumayeri", "Gölyaka",
    "Gümüşova", "Kaynaşlı", "Merkez", "Yığılca",

    # EDİRNE
    "Enez", "Havsa", "İpsala", "Keşan", "Lalapaşa",
    "Meriç", "Merkez", "Süloğlu", "Uzunköprü",

    # ELAZIĞ
    "Ağın", "Alacakaya", "Arıcak", "Baskil", "Karakoçan",
    "Keban", "Kovancılar", "Maden", "Merkez", "Palu",
    "Sivrice",

    # ERZİNCAN
    "Çayırlı", "İliç", "Kemah", "Kemaliye", "Merkez",
    "Otlukbeli", "Refahiye", "Tercan", "Üzümlü",

    # ERZURUM
    "Aşkale", "Aziziye", "Çat", "Hınıs", "Horasan",
    "İspir", "Karaçoban", "Karayazı", "Köprüköy", "Narman",
    "Oltu", "Olur", "Palandöken", "Pasinler", "Pazaryolu",
    "Şenkaya", "Tekman", "Tortum", "Uzundere", "Yakutiye",

    # ESKİŞEHİR
    "Alpu", "Beylikova", "Çifteler", "Günyüzü", "Han",
    "İnönü", "Mahmudiye", "Mihalgazi", "Mihalıççık",
    "Odunpazarı", "Sarıcakaya", "Seyitgazi", "Sivrihisar",
    "Tepebaşı",

    # GAZİANTEP
    "Araban", "İslahiye", "Karkamış", "Nizip", "Nurdağı",
    "Oğuzeli", "Şahinbey", "Şehitkamil", "Yavuzeli",

    # GİRESUN
    "Alucra", "Bulancak", "Çamoluk", "Çanakçı", "Dereli",
    "Doğankent", "Espiye", "Eynesil", "Görele", "Güce",
    "Keşap", "Merkez", "Piraziz", "Şebinkarahisar",
    "Tirebolu", "Yağlıdere",

    # GÜMÜŞHANE
    "Kelkit", "Köse", "Kürtün", "Merkez", "Şiran", "Torul",

    # HAKKARİ
    "Çukurca", "Derecik", "Merkez", "Şemdinli", "Yüksekova",

    # HATAY
    "Altınözü", "Antakya", "Arsuz", "Belen", "Defne",
    "Dörtyol", "Erzin", "Hassa", "İskenderun", "Kırıkhan",
    "Kumlu", "Payas", "Reyhanlı", "Samandağ", "Yayladağı",

    # IĞDIR
    "Aralık", "Karakoyunlu", "Merkez", "Tuzluca",

    # ISPARTA
    "Aksu", "Atabey", "Eğirdir", "Gelendost", "Gönen",
    "Keçiborlu", "Merkez", "Senirkent", "Sütçüler",
    "Şarkikaraağaç", "Uluborlu", "Yalvaç", "Yenişarbademli",

    # İSTANBUL
    "Adalar", "Arnavutköy", "Ataşehir", "Avcılar", "Bağcılar",
    "Bahçelievler", "Bakırköy", "Başakşehir", "Bayrampaşa",
    "Beşiktaş", "Beykoz", "Beylikdüzü", "Beyoğlu",
    "Büyükçekmece", "Çatalca", "Çekmeköy", "Esenler",
    "Esenyurt", "Eyüpsultan", "Fatih", "Gaziosmanpaşa",
    "Güngören", "Kadıköy", "Kağıthane", "Kartal",
    "Küçükçekmece", "Maltepe", "Pendik", "Sancaktepe",
    "Sarıyer", "Silivri", "Sultanbeyli", "Sultangazi",
    "Şile", "Şişli", "Tuzla", "Ümraniye", "Üsküdar",
    "Zeytinburnu",

    # İZMİR
    "Aliağa", "Balçova", "Bayındır", "Bayraklı", "Bergama",
    "Beydağ", "Bornova", "Buca", "Çeşme", "Çiğli", "Dikili",
    "Foça", "Gaziemir", "Güzelbahçe", "Karabağlar",
    "Karaburun", "Karşıyaka", "Kemalpaşa", "Kınık", "Kiraz",
    "Konak", "Menderes", "Menemen", "Narlıdere", "Ödemiş",
    "Seferihisar", "Selçuk", "Tire", "Torbalı", "Urla",

    # KAHRAMANMARAŞ
    "Afşin", "Andırın", "Çağlayancerit", "Dulkadiroğlu",
    "Ekinözü", "Elbistan", "Göksun", "Nurhak", "Onikişubat",
    "Pazarcık", "Türkoğlu",

    # KARABÜK
    "Eflani", "Eskipazar", "Merkez", "Ovacık",
    "Safranbolu", "Yenice",

    # KARAMAN
    "Ayrancı", "Başyayla", "Ermenek", "Kazımkarabekir",
    "Merkez", "Sarıveliler",

    # KARS
    "Akyaka", "Arpaçay", "Digor", "Kağızman", "Merkez",
    "Sarıkamış", "Selim", "Susuz",

    # KASTAMONU
    "Abana", "Ağlı", "Araç", "Azdavay", "Bozkurt",
    "Cide", "Çatalzeytin", "Daday", "Devrekani", "Doğanyurt",
    "Hanönü", "İhsangazi", "İnebolu", "Küre", "Merkez",
    "Pınarbaşı", "Seydiler", "Şenpazar", "Taşköprü", "Tosya",

    # KAYSERİ
    "Akkışla", "Bünyan", "Develi", "Felahiye", "Hacılar",
    "İncesu", "Kocasinan", "Melikgazi", "Özvatan",
    "Pınarbaşı", "Sarıoğlan", "Sarız", "Talas", "Tomarza",
    "Yahyalı", "Yeşilhisar",

    # KIRKLARELİ
    "Babaeski", "Demirköy", "Kofçaz", "Lüleburgaz",
    "Merkez", "Pehlivanköy", "Pınarhisar", "Vize",

    # KIRŞEHİR
    "Akçakent", "Akpınar", "Boztepe", "Çiçekdağı",
    "Kaman", "Merkez", "Mucur",

    # KIRIKKALE
    "Bahşılı", "Balışeyh", "Çelebi", "Delice", "Karakeçili",
    "Keskin", "Merkez", "Sulakyurt", "Yahşihan",

    # KİLİS
    "Elbeyli", "Merkez", "Musabeyli", "Polateli",

    # KOCAELİ
    "Başiskele", "Çayırova", "Darıca", "Derince", "Dilovası",
    "Gebze", "Gölcük", "İzmit", "Kandıra", "Karamürsel",
    "Kartepe", "Körfez",

    # KONYA
    "Ahırlı", "Akören", "Akşehir", "Altınekin", "Beyşehir",
    "Bozkır", "Cihanbeyli", "Çeltik", "Çumra", "Derbent",
    "Derebucak", "Doğanhisar", "Emirgazi", "Ereğli",
    "Güneysınır", "Hadim", "Halkapınar", "Hüyük", "Ilgın",
    "Kadınhanı", "Karapınar", "Karatay", "Kulu", "Meram",
    "Sarayönü", "Selçuklu", "Seydişehir", "Taşkent",
    "Tuzlukçu", "Yalıhüyük", "Yunak",

    # KÜTAHYA
    "Altıntaş", "Aslanapa", "Çavdarhisar", "Domaniç",
    "Dumlupınar", "Emet", "Gediz", "Hisarcık", "Merkez",
    "Pazarlar", "Şaphane", "Simav", "Tavşanlı",

    # MALATYA
    "Akçadağ", "Arapgir", "Arguvan", "Battalgazi", "Darende",
    "Doğanşehir", "Doğanyol", "Hekimhan", "Kale",
    "Kuluncak", "Pütürge", "Yazıhan", "Yeşilyurt",

    # MANİSA
    "Ahmetli", "Akhisar", "Alaşehir", "Demirci", "Gölmarmara",
    "Gördes", "Kırkağaç", "Köprübaşı", "Kula", "Salihli",
    "Sarıgöl", "Saruhanlı", "Şehzadeler", "Selendi", "Soma",
    "Turgutlu", "Yunusemre",

    # MARDİN
    "Artuklu", "Dargeçit", "Derik", "Kızıltepe", "Mazıdağı",
    "Midyat", "Nusaybin", "Ömerli", "Savur", "Yeşilli",

    # MERSİN
    "Akdeniz", "Anamur", "Aydıncık", "Bozyazı", "Çamlıyayla",
    "Erdemli", "Gülnar", "Mezitli", "Mut", "Silifke",
    "Tarsus", "Toroslar", "Yenişehir",

    # MUĞLA
    "Bodrum", "Dalaman", "Datça", "Fethiye", "Kavaklıdere",
    "Köyceğiz", "Marmaris", "Menteşe", "Milas", "Ortaca",
    "Seydikemer", "Ula", "Yatağan",

    # MUŞ
    "Bulanık", "Hasköy", "Korkut", "Malazgirt", "Merkez", "Varto",

    # NEVŞEHİR
    "Acıgöl", "Avanos", "Derinkuyu", "Gülşehir", "Hacıbektaş",
    "Kozaklı", "Merkez", "Ürgüp",

    # NİĞDE
    "Altunhisar", "Bor", "Çamardı", "Çiftlik", "Merkez",
    "Ulukışla",

    # ORDU
    "Akkuş", "Altınordu", "Aybastı", "Çamaş", "Çatalpınar",
    "Çaybaşı", "Fatsa", "Gölköy", "Gülyalı", "Gürgentepe",
    "İkizce", "Kabadüz", "Kabataş", "Korgan", "Kumru",
    "Mesudiye", "Perşembe", "Ulubey", "Ünye",

    # OSMANİYE
    "Bahçe", "Düziçi", "Hasanbeyli", "Kadirli", "Merkez",
    "Sumbas", "Toprakkale",

    # RİZE
    "Ardeşen", "Çamlıhemşin", "Çayeli", "Derepazarı",
    "Fındıklı", "Güneysu", "Hemşin", "İkizdere", "İyidere",
    "Kalkandere", "Merkez", "Pazar",

    # SAKARYA
    "Adapazarı", "Akyazı", "Arifiye", "Erenler", "Ferizli",
    "Geyve", "Hendek", "Karapürçek", "Karasu", "Kaynarca",
    "Kocaali", "Pamukova", "Sapanca", "Serdivan", "Söğütlü",
    "Taraklı",

    # SAMSUN
    "19 Mayıs", "Alaçam", "Asarcık", "Atakum", "Ayvacık",
    "Bafra", "Canik", "Çarşamba", "Havza", "İlkadım",
    "Kavak", "Ladik", "Salıpazarı", "Tekkeköy", "Terme",
    "Vezirköprü", "Yakakent",

    # SİİRT
    "Aydınlar", "Baykan", "Eruh", "Kurtalan", "Merkez",
    "Pervari", "Şirvan",

    # SİNOP
    "Ayancık", "Boyabat", "Dikmen", "Durağan", "Erfelek",
    "Gerze", "Merkez", "Saraydüzü", "Türkeli",

    # SİVAS
    "Akıncılar", "Altınyayla", "Divriği", "Doğanşar",
    "Gemerek", "Gölova", "Gürün", "Hafik", "İmranlı",
    "Kangal", "Koyulhisar", "Merkez", "Şarkışla", "Suşehri",
    "Ulaş", "Yıldızeli", "Zara",

    # ŞANLIURFA
    "Akçakale", "Birecik", "Bozova", "Ceylanpınar",
    "Eyyübiye", "Halfeti", "Haliliye", "Harran", "Hilvan",
    "Karaköprü", "Siverek", "Suruç", "Viranşehir",

    # ŞIRNAK
    "Beytüşşebap", "Cizre", "Güçlükonak", "İdil", "Merkez",
    "Silopi", "Uludere",

    # TEKİRDAĞ
    "Çerkezköy", "Çorlu", "Ergene", "Hayrabolu", "Kapaklı",
    "Malkara", "Marmaraereğlisi", "Muratlı", "Saray",
    "Şarköy", "Süleymanpaşa",

    # TOKAT
    "Almus", "Artova", "Başçiftlik", "Erbaa", "Merkez",
    "Niksar", "Pazar", "Reşadiye", "Sulusaray", "Turhal",
    "Yeşilyurt", "Zile",

    # TRABZON
    "Akçaabat", "Araklı", "Arsin", "Beşikdüzü", "Çarşıbaşı",
    "Çaykara", "Dernekpazarı", "Düzköy", "Hayrat", "Köprübaşı",
    "Maçka", "Of", "Ortahisar", "Şalpazarı", "Sürmene",
    "Tonya", "Vakfıkebir", "Yomra",

    # TUNCELİ
    "Çemişgezek", "Hozat", "Mazgirt", "Merkez", "Nazımiye",
    "Ovacık", "Pertek", "Pülümür",

    # UŞAK
    "Banaz", "Eşme", "Karahallı", "Merkez", "Sivaslı", "Ulubey",

    # VAN
    "Bahçesaray", "Başkale", "Çaldıran", "Çatak", "Edremit",
    "Erciş", "Gevaş", "Gürpınar", "İpekyolu", "Muradiye",
    "Özalp", "Saray", "Tuşba",

    # YALOVA
    "Altınova", "Armutlu", "Çınarcık", "Çiftlikköy",
    "Merkez", "Termal",

    # YOZGAT
    "Akdağmadeni", "Aydıncık", "Boğazlıyan", "Çandır",
    "Çayıralan", "Çekerek", "Kadışehri", "Merkez",
    "Saraykent", "Sarıkaya", "Sefaatli", "Sorgun",
    "Yenifakılı", "Yerköy",

    # ZONGULDAK
    "Alaplı", "Çaycuma", "Devrek", "Ereğli", "Gökçebey",
    "Kilimli", "Kozlu", "Merkez"
]


# =========================================================
# TÜRKÇE KARAKTER YARDIMCILARI
# =========================================================

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
    # Türkçe büyük harf dönüşümü: i -> İ, ı -> I
    return text.replace("i", "İ").replace("ı", "I").upper()


# =========================================================
# İLÇE BUL
# =========================================================

# Uzun ilçe isimlerini önce kontrol et
# Örn. "MARMARAEREĞLİSİ" önce gelsin.
ILCELER = sorted(
    set(ILCELER),
    key=lambda x: len(normalize(x)),
    reverse=True
)

ILCELER_NORMALIZED = {
    normalize(ilce): tr_upper(ilce)
    for ilce in ILCELER
}


def ilce_bul(adres):

    if pd.isna(adres):
        return ""

    adres = str(adres).strip()

    if not adres:
        return ""

    adres_normalized = normalize(adres)

    for ilce_normalized, ilce_original in ILCELER_NORMALIZED.items():

        # Kelime sınırlarıyla ara
        pattern = rf"(?<![A-Z]){re.escape(ilce_normalized)}(?![A-Z])"

        if re.search(pattern, adres_normalized):
            return ilce_original

    return ""


# =========================================================
# EXCEL OKU
# =========================================================

df = pd.read_excel(INPUT_FILE)

# =========================================================
# İLÇE SÜTUNUNU BUL / OLUŞTUR
# =========================================================

# Sütun isimlerinde boşluk varsa temizle
df.columns = df.columns.str.strip()

# ILCE / İLÇE / ilce gibi farklı yazımları yakala
ilce_sutunu = None

for sutun in df.columns:

    sutun_normalized = normalize(sutun)

    if sutun_normalized == "ILCE":
        ilce_sutunu = sutun
        break


# =========================================================
# İLÇE SÜTUNU YOKSA OLUŞTUR
# =========================================================

if ilce_sutunu is None:

    # ILCE sütunu yoksa oluştur
    df["ILCE"] = ""

    ilce_sutunu = "ILCE"


# =========================================================
# ADRESLERDEN İLÇELERİ BUL
# =========================================================

df[ilce_sutunu] = df["ADRES"].apply(ilce_bul)


# =========================================================
# EXCEL'E AYNI DOSYANIN ÜZERİNE YAZ
# =========================================================

df.to_excel(OUTPUT_FILE, index=False)

# =========================================================
# SONUÇ
# =========================================================

toplam = len(df)
bulunan = (df["İLÇE"] != "").sum()
bulunamayan = toplam - bulunan

print("=" * 60)
print("TAMAMLANDI")
print("=" * 60)
print(f"Toplam kayıt     : {toplam}")
print(f"İlçe bulunan     : {bulunan}")
print(f"İlçe bulunamayan : {bulunamayan}")
print(f"Başarı oranı     : %{bulunan / toplam * 100:.1f}")
print(f"Dosya            : {OUTPUT_FILE}")
print("=" * 60)