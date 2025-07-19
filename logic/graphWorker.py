import logging  # Hata ve bilgi loglama
from typing import List, Tuple
import pandas as pd  # Veri işleme
from PyQt5.QtCore import QThread, pyqtSignal  # PyQt5 iş parçacığı ve sinyal sistemi

from utils.helpers import seconds_from_timedelta  # Yardımcı fonksiyon: timedelta -> saniye

class GraphWorker(QThread):
    """Arka planda grafik verisi işleyen iş parçacığı sınıfı."""

    finished = pyqtSignal(list)  # İşlem tamamlandığında sonuç listesi gönderilir
    progress = pyqtSignal(int)   # Yüzdelik ilerleme bilgisi yayınlanır
    error = pyqtSignal(str)      # Hata mesajı yayınlanır

    def __init__(
            self,
            df: pd.DataFrame,  # Ham veri DataFrame
            grouping_col_name: str,  # Ana gruplama sütunu adı
            grouped_col_name: str,  # Alt gruplama sütunu adı
            grouped_values: List[str],  # Gruplanacak alt değerler
            metric_cols: List[str],  # Süre içeren metrik sütunlar
            oee_col_name: str | None,  # OEE sütunu (varsa)
            selected_grouping_val: str  # Seçilen grup (örn. 'Tarih' veya 'Hat')
    ) -> None:
        super().__init__()  # QThread constructor
        # Gerekli sütunları içeren yeni bir DataFrame oluştur (güvenli kopya)
        self.df = df[[grouping_col_name, grouped_col_name, oee_col_name] + metric_cols].copy() if oee_col_name else \
                  df[[grouping_col_name, grouped_col_name] + metric_cols].copy()
        self.grouping_col_name = grouping_col_name
        self.grouped_col_name = grouped_col_name
        self.grouped_values = grouped_values
        self.metric_cols = metric_cols
        self.oee_col_name = oee_col_name
        self.selected_grouping_val = selected_grouping_val

    def run(self) -> None:
        """İş parçacığı çalıştığında veri işleyip grafik sonuçlarını üretir."""
        try:
            results: List[Tuple[str, pd.Series, str]] = []  # Sonuç listesi: (grup değeri, metrik toplamları, OEE)
            total = len(self.grouped_values)  # Toplam alt grup sayısı

            # Metrik sütunlarını saniyeye çevir
            for col in self.metric_cols:
                if col in self.df.columns:
                    self.df[col] = seconds_from_timedelta(self.df[col])

            # Gruplama sütunlarını string'e dönüştür (karşılaştırmalar için güvenli)
            if self.grouping_col_name in self.df.columns:
                self.df[self.grouping_col_name] = self.df[self.grouping_col_name].astype(str)
            if self.grouped_col_name in self.df.columns:
                self.df[self.grouped_col_name] = self.df[self.grouped_col_name].astype(str)

            # Her alt grup için işlem yap
            for i, current_grouped_val in enumerate(self.grouped_values, 1):
                # Belirli grup ve alt grup için alt küme oluştur
                subset_df_for_chart = self.df[
                    (self.df[self.grouping_col_name] == self.selected_grouping_val) &
                    (self.df[self.grouped_col_name] == current_grouped_val)
                ]

                # Metrik sütunların toplamını al (0'dan büyük olanları filtrele)
                sums = subset_df_for_chart[[col for col in self.metric_cols if col in subset_df_for_chart.columns]].sum()
                sums = sums[sums > 0]

                oee_display_value = "0%"  # Varsayılan OEE değeri

                # OEE varsa, formatla
                if self.oee_col_name and self.oee_col_name in subset_df_for_chart.columns and not subset_df_for_chart.empty:
                    oee_value_raw = subset_df_for_chart[self.oee_col_name].values[0]
                    if pd.notna(oee_value_raw):
                        if isinstance(oee_value_raw, str) and oee_value_raw.strip().upper() == "\u00dcRET\u0130M YAPILMADI":
                            oee_display_value = ""  # Özel durum: üretim yapılmadı
                        else:
                            try:
                                oee_value_float: float
                                if isinstance(oee_value_raw, str):
                                    oee_value_str = oee_value_raw.replace('%', '').strip()
                                    oee_value_float = float(oee_value_str)
                                elif isinstance(oee_value_raw, (int, float)):
                                    oee_value_float = float(oee_value_raw)
                                else:
                                    raise ValueError("Desteklenmeyen OEE değeri tipi veya formatı")

                                # 0-1 arası ise % çevir, >1 zaten % olarak kabul edilir
                                if 0.0 <= oee_value_float <= 1.0 and oee_value_float != 0:
                                    oee_display_value = f"{oee_value_float * 100:.0f}%"
                                elif oee_value_float > 1.0:
                                    oee_display_value = f"{oee_value_float:.0f}%"
                                else:
                                    oee_display_value = "0%"
                            except (ValueError, TypeError):
                                logging.warning(
                                    f"OEE de\u011feri d\u00f6n\u00fc\u015ft\u00fcr\u00fclemedi: {oee_value_raw}. Varsay\u0131lan '0%' kullan\u0131lacak.")
                                oee_display_value = "0%"

                if not sums.empty:
                    results.append((current_grouped_val, sums, oee_display_value))  # Sonuçlara ekle

                self.progress.emit(int(i / total * 100))  # İlerleme sinyali gönder

            self.finished.emit(results)  # İşlem tamamlandığında sonuçları gönder

        except Exception as exc:
            logging.exception("GraphWorker hatas\u0131 olu\u015ftu.")  # Log'a yaz
            self.error.emit(f"Grafik olu\u015fturulurken bir hata olu\u015ftu: {str(exc)}")  # Hata sinyali g\u00f6nder
