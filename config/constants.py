import logging  # Uygulama günlükleme işlemleri için
import sys      # Sistem çıktıları/logları için
import matplotlib.pyplot as plt  # Grafik çizimi için Matplotlib

# Her sayfada kaç grafik gösterileceği (sayfalama için)
GRAPHS_PER_PAGE = 1  # Örn: DailyGraphsPage'de 1 grafik göster

# Excel dosyasında beklenen sayfa isimleri
REQ_SHEETS = {"SMD-OEE", "ROBOT", "DALGA_LEHİM", "KAPLAMA-OEE"}  # Gerekli sheet'ler

# ------------------------------------------
# Loglama ayarları
# ------------------------------------------
logging.basicConfig(
    level=logging.INFO,  # Log seviyesini INFO olarak ayarla
    format="%(asctime)s [%(levelname)s] %(message)s",  # Log formatı: zaman - seviye - mesaj
    handlers=[logging.StreamHandler(sys.stdout)],  # Log'ları terminale yazdır
)

# ------------------------------------------
# Matplotlib yazı tipi ayarları (Türkçe karakter desteği için)
# ------------------------------------------
plt.rcParams['font.family'] = 'DejaVu Sans'  # Varsayılan font ailesi
plt.rcParams['font.sans-serif'] = [          # Türkçe karakterleri destekleyen alternatif sans-serif fontlar
    'SimSun', 'Arial', 'Liberation Sans', 'Bitstream Vera Sans', 'sans-serif'
]
plt.rcParams['axes.unicode_minus'] = False  # Eksi işaretleri düzgün çıksın diye

# ------------------------------------------
# Genel Matplotlib ayarları (grafik görünümü)
# ------------------------------------------
plt.rcParams['axes.grid'] = True           # Grafiklerde grid (kılavuz çizgisi) göster
plt.rcParams['grid.alpha'] = 0.7           # Grid çizgilerinin saydamlığı
plt.rcParams['grid.linestyle'] = '--'      # Grid çizgi stili: kesik çizgi
plt.rcParams['grid.linewidth'] = 0.5       # Grid çizgi kalınlığı
plt.rcParams['figure.dpi'] = 100           # Ekranda gösterim DPI değeri
plt.rcParams['savefig.dpi'] = 300          # Kaydedilen görsellerin DPI değeri

# ------------------------------------------
# X/Y eksen tik işaretleri ayarları
# ------------------------------------------
plt.rcParams['xtick.direction'] = 'out'    # X ekseninde tikler dışa baksın
plt.rcParams['ytick.direction'] = 'out'    # Y ekseninde tikler dışa baksın
plt.rcParams['xtick.major.size'] = 7       # X ekseni büyük tik uzunluğu
plt.rcParams['xtick.minor.size'] = 4       # X ekseni küçük tik uzunluğu
plt.rcParams['ytick.major.size'] = 7       # Y ekseni büyük tik uzunluğu
plt.rcParams['ytick.minor.size'] = 4       # Y ekseni küçük tik uzunluğu
plt.rcParams['xtick.major.width'] = 1.5    # X ekseni büyük tik kalınlığı
plt.rcParams['xtick.minor.width'] = 1      # X ekseni küçük tik kalınlığı
plt.rcParams['xtick.top'] = False          # X ekseninde üst tik çizgisi gösterme
plt.rcParams['ytick.right'] = False        # Y ekseninde sağ tik çizgisi gösterme
plt.rcParams['axes.edgecolor'] = 'black'   # Grafik çerçeve rengi
plt.rcParams['axes.linewidth'] = 1.5       # Grafik çerçeve kalınlığı
