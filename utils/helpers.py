import sys
import logging
import datetime
import pandas as pd
import matplotlib.pyplot as plt

# --- Genel Sabitler ---
GRAPHS_PER_PAGE = 1  # Her sayfada gösterilecek grafik sayısı
REQ_SHEETS = {"SMD-OEE", "ROBOT", "DALGA_LEHİM", "KAPLAMA-OEE"}  # Gerekli Excel sayfaları

# --- Loglama Ayarları ---
logging.basicConfig(
    level=logging.INFO,  # INFO ve üzeri seviyedeki mesajlar loglanacak
    format="%(asctime)s [%(levelname)s] %(message)s",  # Zaman damgası ve seviye dahil
    handlers=[logging.StreamHandler(sys.stdout)],  # Konsola yazdır
)

# --- Matplotlib Genel Ayarları ---

# Türkçe karakter ve font ayarları
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.sans-serif'] = [
    'SimSun', 'Arial', 'Liberation Sans', 'Bitstream Vera Sans', 'sans-serif'
]
plt.rcParams['axes.unicode_minus'] = False  # Negatif işaretler için Unicode kapalı

# Izgara çizgileri ayarları
plt.rcParams['axes.grid'] = True
plt.rcParams['grid.alpha'] = 0.7
plt.rcParams['grid.linestyle'] = '--'
plt.rcParams['grid.linewidth'] = 0.5

# DPI ayarları
plt.rcParams['figure.dpi'] = 100  # Ekranda
plt.rcParams['savefig.dpi'] = 300  # Kaydederken yüksek çözünürlük

# Tick (işaretleyici) ayarları
plt.rcParams['xtick.direction'] = 'out'
plt.rcParams['ytick.direction'] = 'out'
plt.rcParams['xtick.major.size'] = 7
plt.rcParams['xtick.minor.size'] = 4
plt.rcParams['ytick.major.size'] = 7
plt.rcParams['ytick.minor.size'] = 4
plt.rcParams['xtick.major.width'] = 1.5
plt.rcParams['xtick.minor.width'] = 1
plt.rcParams['xtick.top'] = False
plt.rcParams['ytick.right'] = False

plt.rcParams['axes.edgecolor'] = 'black'
plt.rcParams['axes.linewidth'] = 1.5


def excel_col_to_index(col_letter: str) -> int:
    """
    Excel sütun harfini sıfır tabanlı sayısal indekse dönüştürür.
    Örneğin:
        'A'  -> 0
        'Z'  -> 25
        'AA' -> 26
        'BD' -> 55
    Parametre:
        col_letter: Excel sütun harfi (büyük/küçük harf farketmez)
    Dönen:
        Sıfır tabanlı sütun indeksi (int)
    Hata:
        Geçersiz sütun harfi verilirse ValueError fırlatır.
    """
    index = 0
    for char in col_letter.upper():
        if not char.isalpha():
            raise ValueError(f"Geçersiz sütun harfi: {col_letter}. Sadece alfabetik karakterler kabul edilir.")
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1


def seconds_from_timedelta(series: pd.Series) -> pd.Series:
    """
    Pandas Serisindeki farklı zaman formatlarındaki değerleri (datetime.time, timedelta string, sayısal vb.)
    toplam saniye cinsine çevirir. Dönüştürülemeyenler 0.0 olarak işaretlenir.

    İşleyiş:
    - datetime.time objeleri saat, dakika, saniye + mikrosaniyeye göre hesaplanır.
    - string veya timedelta formatları pd.to_timedelta ile dönüştürülür.
    - Sayısal değerler gün olarak kabul edilip saniyeye çevrilir.
    - Sonuç tüm indeksler için float saniye cinsinden döner.

    Parametre:
        series: pd.Series, zaman bilgileri içeren

    Dönen:
        pd.Series, aynı indeksle saniye cinsinden değerler (float)
    """
    seconds_series = pd.Series(0.0, index=series.index, dtype=float)

    # 1) datetime.time objelerini işleme
    is_time_obj = series.apply(lambda x: isinstance(x, datetime.time))
    if is_time_obj.any():
        time_objects = series[is_time_obj]
        seconds_series.loc[is_time_obj] = time_objects.apply(
            lambda t: t.hour * 3600 + t.minute * 60 + t.second + t.microsecond / 1e6
        )

    # 2) timedelta string veya diğer formatları işleme
    str_and_timedelta_mask = ~is_time_obj & series.notna()
    if str_and_timedelta_mask.any():
        converted_td = pd.to_timedelta(series.loc[str_and_timedelta_mask].astype(str).str.strip(), errors='coerce')
        valid_td_mask = pd.notna(converted_td)
        seconds_series.loc[str_and_timedelta_mask & valid_td_mask] = converted_td[valid_td_mask].dt.total_seconds()

    # 3) Kalan NaN veya dönüştürülemeyenleri sayısal gün olarak işleme
    remaining_nan_mask = seconds_series.isna()
    if remaining_nan_mask.any():
        numeric_candidates = series.loc[remaining_nan_mask]
        numeric_values = pd.to_numeric(numeric_candidates, errors='coerce')
        valid_numeric_mask = pd.notna(numeric_values)
        if valid_numeric_mask.any():
            converted_from_numeric = pd.to_timedelta(numeric_values[valid_numeric_mask], unit='D', errors='coerce')
            valid_num_td_mask = pd.notna(converted_from_numeric)
            seconds_series.loc[remaining_nan_mask[valid_numeric_mask] & valid_num_td_mask] = converted_from_numeric[
                valid_num_td_mask].dt.total_seconds()

    # Dönüştürülemeyenler 0.0 ile doldurulur
    return seconds_series.fillna(0.0)
