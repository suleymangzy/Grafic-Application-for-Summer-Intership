import sys
import logging
import datetime
from pathlib import Path
import re  # Düzenli ifadeler için eklendi
from typing import List, Tuple, Any, Union

import pandas as pd
import numpy as np

import matplotlib

# matplotlib.use("Agg") # Masaüstü uygulaması için yoruma alındı
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.ticker import PercentFormatter

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QFileDialog,
    QPushButton,
    QLabel,
    QVBoxLayout,
    QHBoxLayout,
    QListWidget,
    QListWidgetItem,
    QComboBox,
    QMessageBox,
    QProgressBar,
    QStackedWidget,
    QScrollArea,
    QFrame,
    QCheckBox,
    QSpacerItem,  # Import QSpacerItem for flexible spacing
    QSizePolicy,  # Import QSizePolicy for size policies
    QLineEdit,  # For input fields
)

GRAPHS_PER_PAGE = 1  # Her sayfada kaç grafik gösterileceği
REQ_SHEETS = {"SMD-OEE", "ROBOT", "DALGA_LEHİM"}  # Gerekli sayfa isimleri

# Loglama ayarları
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)

# Matplotlib font ayarları (Türkçe karakter desteği)
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.sans-serif'] = ['SimSun', 'Arial', 'Liberation Sans', 'Bitstream Vera Sans', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False  # Negatif işaretler için

# Global matplotlib ayarları
plt.rcParams['axes.grid'] = True  # Bar grafiği için aşağıda özel olarak False yapılacak
plt.rcParams['grid.alpha'] = 0.7
plt.rcParams['grid.linestyle'] = '--'
plt.rcParams['grid.linewidth'] = 0.5
plt.rcParams['figure.dpi'] = 100  # Ekran çözünürlüğü için
plt.rcParams['savefig.dpi'] = 300  # Kaydedilen resim çözünürlüğü için

# Tick ayarları: Düz çizgiler ve bitiş noktalarında uyumlu noktalar
plt.rcParams['xtick.direction'] = 'out'  # tick markların dışa doğru olmasını sağlar
plt.rcParams['ytick.direction'] = 'out'
plt.rcParams['xtick.major.size'] = 7  # Büyük tick uzunluğu
plt.rcParams['xtick.minor.size'] = 4  # Küçük tick uzunluğu
plt.rcParams['ytick.major.size'] = 7
plt.rcParams['ytick.minor.size'] = 4
plt.rcParams['xtick.major.width'] = 1.5  # Büyük tick kalınlığı
plt.rcParams['xtick.minor.width'] = 1  # Küçük tick kalınlığı
plt.rcParams['xtick.top'] = False  # Üst tickleri kapat
plt.rcParams['ytick.right'] = False  # Sağ tickleri kapat
plt.rcParams['axes.edgecolor'] = 'black'  # Eksen çizgisi rengi
plt.rcParams['axes.linewidth'] = 1.5  # Eksen çizgisi kalınlığı


def excel_col_to_index(col_letter: str) -> int:
    """Excel sütun harfini sıfır tabanlı indekse dönüştürür."""
    index = 0
    for char in col_letter.upper():
        if not char.isalpha():
            raise ValueError(f"Geçersiz sütun harfi: {col_letter}")
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1


def seconds_from_timedelta(series: pd.Series) -> pd.Series:
    """Pandas Serisi'ndeki zaman değerlerini saniyeye dönüştürür.
    Çeşitli zaman formatlarını (timedelta, time objesi, HH:MM:SS string, sadece sayı) destekler.
    Geçersiz değerleri 0 olarak işler.
    Optimizasyon: apply kullanmadan datetime.time objelerini ve sayısal değerleri daha verimli işleme.
    """
    seconds_series = pd.Series(0.0, index=series.index, dtype=float)

    # Convert datetime.time objects
    is_time_obj = series.apply(lambda x: isinstance(x, datetime.time))
    if is_time_obj.any():
        time_objects = series[is_time_obj]
        seconds_series.loc[is_time_obj] = time_objects.apply(
            lambda t: t.hour * 3600 + t.minute * 60 + t.second + t.microsecond / 1e6
        )

    # Convert timedelta objects or strings convertible to timedelta
    str_and_timedelta_mask = ~is_time_obj & series.notna()
    if str_and_timedelta_mask.any():
        converted_td = pd.to_timedelta(series.loc[str_and_timedelta_mask].astype(str).strip(), errors='coerce')
        valid_td_mask = pd.notna(converted_td)
        seconds_series.loc[str_and_timedelta_mask & valid_td_mask] = converted_td[valid_td_mask].dt.total_seconds()

    # Convert numeric values (which might represent days in Excel)
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

    return seconds_series.fillna(0.0)


class GraphWorker(QThread):
    """Grafik oluşturma işlemlerini arka planda yürüten işçi sınıfı."""
    finished = pyqtSignal(list)  # İşlem bittiğinde sonuçları gönderir
    progress = pyqtSignal(int)  # İlerleme bilgisini gönderir
    error = pyqtSignal(str)  # Hata mesajını gönderir

    def __init__(
            self,
            df: pd.DataFrame,
            grouping_col_name: str,
            grouped_col_name: str,
            grouped_values: List[str],
            metric_cols: List[str],
            oee_col_name: str | None,
            selected_grouping_val: str
    ) -> None:
        super().__init__()
        # Veri çerçevesinin bir kopyasıyla çalışmak yerine,
        # sadece gerekli sütunları kopyalayarak bellek kullanımını optimize et.
        # Bu iş parçacığı sadece okuma yapacaksa, kopyalama gereksiz olabilir.
        # Ancak, güvenlik için burada kopyalama tercih ediliyor.
        self.df = df[[grouping_col_name, grouped_col_name, oee_col_name] + metric_cols].copy() if oee_col_name else \
            df[[grouping_col_name, grouped_col_name] + metric_cols].copy()
        self.grouping_col_name = grouping_col_name
        self.grouped_col_name = grouped_col_name
        self.grouped_values = grouped_values
        self.metric_cols = metric_cols
        self.oee_col_name = oee_col_name
        self.selected_grouping_val = selected_grouping_val

    def run(self) -> None:
        """İş parçacığı başladığında çalışacak metod."""
        try:
            results: List[Tuple[str, pd.Series, str]] = []  # (Gruplama değeri, Metrik toplamları, OEE değeri)
            total = len(self.grouped_values)  # Toplam işlem sayısı

            # Tüm metrik sütunlarını saniyeye dönüştür (sadece bir kere yap)
            for col in self.metric_cols:
                if col in self.df.columns:
                    self.df[col] = seconds_from_timedelta(self.df[col])

            # Gruplama ve gruplanan sütunları bir kere string'e çevir
            if self.grouping_col_name in self.df.columns:
                self.df[self.grouping_col_name] = self.df[self.grouping_col_name].astype(str)
            if self.grouped_col_name in self.df.columns:
                self.df[self.grouped_col_name] = self.df[self.grouped_col_name].astype(str)

            for i, current_grouped_val in enumerate(self.grouped_values, 1):
                # Mevcut gruplama ve gruplanan değerlere göre alt veri çerçevesini filtrele
                subset_df_for_chart = self.df[
                    (self.df[self.grouping_col_name] == self.selected_grouping_val) &
                    (self.df[self.grouped_col_name] == current_grouped_val)
                    ]

                # Metrik sütunlarının toplamlarını al
                # Sadece mevcut sütunlarda sum al
                sums = subset_df_for_chart[
                    [col for col in self.metric_cols if col in subset_df_for_chart.columns]].sum()
                sums = sums[sums > 0]  # Sadece sıfırdan büyük toplamları dikkate al

                oee_display_value = "0%"  # Varsayılan OEE değeri
                if self.oee_col_name and self.oee_col_name in subset_df_for_chart.columns and not subset_df_for_chart.empty:
                    # OEE değerini almak için .iloc[0] yerine .values[0] kullanmak biraz daha hızlı olabilir.
                    oee_value_raw = subset_df_for_chart[self.oee_col_name].values[0]
                    if pd.notna(oee_value_raw):
                        try:
                            oee_value_float: float
                            if isinstance(oee_value_raw, str):
                                oee_value_str = oee_value_raw.replace('%', '').strip()
                                oee_value_float = float(oee_value_str)
                            elif isinstance(oee_value_raw, (int, float)):
                                oee_value_float = float(oee_value_raw)
                            else:
                                raise ValueError("Unsupported OEE value type or format")

                            if 0.0 <= oee_value_float <= 1.0 and oee_value_float != 0:
                                oee_display_value = f"{oee_value_float * 100:.0f}%"
                            elif oee_value_float > 1.0:  # Yüzde 100'den büyükse olduğu gibi göster
                                oee_display_value = f"{oee_value_float:.0f}%"
                            else:
                                oee_display_value = "0%"
                        except (ValueError, TypeError):
                            logging.warning(
                                f"OEE değeri dönüştürülemedi: {oee_value_raw}. Varsayılan '0%' kullanılacak.")
                            oee_display_value = "0%"

                if not sums.empty:  # Eğer metrik toplamı varsa grafiğe ekle
                    results.append((current_grouped_val, sums, oee_display_value))
                self.progress.emit(int(i / total * 100))  # İlerleme bilgisini gönder

            self.finished.emit(results)  # İşlem tamamlandığında sonuçları gönder
        except Exception as exc:
            logging.exception("GraphWorker hatası oluştu.")
            self.error.emit(f"Grafik oluşturulurken bir hata oluştu: {str(exc)}")


class GraphPlotter:
    """Matplotlib grafikleri oluşturmak için yardımcı sınıf."""

    @staticmethod
    def create_donut_chart(
            ax: plt.Axes,
            sorted_metrics_series: pd.Series,
            oee_display_value: str,
            chart_colors: List[Any],
            fig: plt.Figure
    ) -> None:
        """Donut grafiği oluşturur."""
        # Donut grafiğini tüm metriklerle oluştur
        wedges, texts = ax.pie(
            sorted_metrics_series,  # Tüm metrikler burada kullanılıyor
            autopct=None,
            startangle=90,
            wedgeprops=dict(width=0.4, edgecolor='w'),
            colors=chart_colors[:len(sorted_metrics_series)]  # Tüm metrikler için renkleri kullan
        )

        # OEE değerini grafik merkezine yerleştir
        ax.text(0, 0, f"OEE\n{oee_display_value}",
                horizontalalignment='center', verticalalignment='center',
                fontsize=24, fontweight='bold', color='black')

        # Metrik etiketlerini grafiğin solunda alt alta yerleştirme ve numaralandırma
        label_y_start = 0.25 + (30 / (fig.get_size_inches()[1] * fig.dpi))
        label_line_height = 0.05

        # Yalnızca ilk 3 metriğin etiketlerini oluştur
        top_3_metrics = sorted_metrics_series.head(3)
        top_3_colors = chart_colors[:len(top_3_metrics)]

        for i, (metric_name, metric_value) in enumerate(top_3_metrics.items()):
            label_text = (
                f"{i + 1}. {metric_name}; "  # Numaralandırma eklendi
                f"{int(metric_value // 3600):02d}:"
                f"{int((metric_value % 3600) // 60):02d}; "
                f"{metric_value / sorted_metrics_series.sum() * 100:.0f}%"  # Yüzdeyi genel toplama göre hesapla
            )
            y_pos = label_y_start - (i * label_line_height)
            bbox_props = dict(boxstyle="round,pad=0.3", fc=top_3_colors[i], ec=top_3_colors[i], lw=0.5)
            r, g, b, _ = matplotlib.colors.to_rgba(top_3_colors[i])
            luminance = (0.299 * r + 0.587 * g + 0.114 * b)
            text_color = 'white' if luminance < 0.5 else 'black'

            fig.text(0.02,  # X koordinatını sola kaydır
                     y_pos,
                     label_text,
                     horizontalalignment='left',
                     verticalalignment='top',
                     fontsize=10,
                     bbox=bbox_props,
                     transform=fig.transFigure,
                     color=text_color
                     )

        ax.set_title("")  # Donut grafik için başlık yok
        ax.axis("equal")  # Pie chart'ın daire şeklinde görünmesini sağlar
        fig.tight_layout(rect=[0.25, 0.1, 1, 0.95])

    @staticmethod
    def create_bar_chart(
            ax: plt.Axes,
            sorted_metrics_series: pd.Series,
            oee_display_value: str,
            chart_colors: List[Any]
    ) -> None:
        """Bar grafiği oluşturur."""
        metrics = sorted_metrics_series.index.tolist()
        values = sorted_metrics_series.values.tolist()

        # Saniye değerlerini dakikaya çevir
        values_minutes = [v / 60 for v in values]

        y_pos = np.arange(len(metrics))

        ax.barh(y_pos, values_minutes, color=chart_colors)
        ax.set_yticks(y_pos)
        ax.set_yticklabels(metrics, fontsize=10)  # Y ekseni etiketlerinde numaralandırma yok
        ax.invert_yaxis()  # En büyük değeri en üste getir

        # X ekseni değerlerini ve etiketlerini kaldır
        ax.set_xlabel("")
        ax.set_xticks([])

        ax.set_title(f"OEE: {oee_display_value}", fontsize=16, fontweight='bold')

        # Grafiğin içinde bulunan hatları (grid) kaldır
        ax.grid(False)

        # Bar grafiğinin genel şablonunu iki ışının birleşimi şeklinde yap
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        # Sol ve alt eksenler zaten varsayılan olarak görünür olmalı, ancak emin olmak için:
        ax.spines['left'].set_visible(True)
        ax.spines['bottom'].set_visible(True)

        # Her barın üzerine değeri ve yüzdesini yaz
        total_sum = sorted_metrics_series.sum()
        for i, (value, metric_name) in enumerate(zip(values, metrics)):
            percentage = (value / total_sum) * 100 if total_sum > 0 else 0
            duration_hours = int(value // 3600)
            duration_minutes = int((value % 3600) // 60)
            text_label = f"{duration_hours:02d}:{duration_minutes:02d} ({percentage:.0f}%)"

            # Metin konumunu barın dışına (sağına) al
            # ha='left' yaparak metnin sol kenarını belirteç pozisyonuna hizala
            # x_position'a küçük bir boşluk ekleyerek barın dışına taşı
            text_x_position = (value / 60) + 0.5  # 0.5 dakika = 30 saniye boşluk
            ax.text(text_x_position, i, text_label,
                    va='center', ha='left',
                    fontsize=11, fontweight='bold',  # Fontu kalın ve daha büyük yap
                    color='black')  # Metin barın dışında olduğu için her zaman siyah olsun

        ax.set_xlim(left=0)
        plt.tight_layout(rect=[0.1, 0.1, 0.95, 0.9])


class FileSelectionPage(QWidget):
    """Dosya seçimi sayfasını temsil eder."""

    def __init__(self, main_window: "MainWindow") -> None:
        super().__init__()
        self.main_window = main_window
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)

        title_label = QLabel("<h2>Dosya Seçimi</h2>")
        title_label.setObjectName("title_label")
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        self.lbl_path = QLabel("Henüz dosya seçilmedi")
        self.lbl_path.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.lbl_path)

        self.btn_browse = QPushButton(".xlsx dosyası seç…")
        self.btn_browse.clicked.connect(self.browse)
        layout.addWidget(self.btn_browse)

        layout.addStretch(1)  # Boşluk ekle

        # Yeni butonlar: Günlük Grafikler ve Aylık Grafikler
        h_layout_buttons = QHBoxLayout()
        self.btn_daily_graphs = QPushButton("Günlük Grafikler")
        self.btn_daily_graphs.clicked.connect(self.go_to_daily_graphs)
        self.btn_daily_graphs.setEnabled(False)  # Enable after file selection
        h_layout_buttons.addWidget(self.btn_daily_graphs)

        self.btn_monthly_graphs = QPushButton("Aylık Grafikler")
        self.btn_monthly_graphs.clicked.connect(self.go_to_monthly_graphs)
        self.btn_monthly_graphs.setEnabled(False)  # Enable after file selection
        h_layout_buttons.addWidget(self.btn_monthly_graphs)

        layout.addLayout(h_layout_buttons)
        layout.addStretch(1)

    def browse(self) -> None:
        """Kullanıcının Excel dosyası seçmesini sağlar."""
        path, _ = QFileDialog.getOpenFileName(self, "Excel seç", str(Path.home()), "Excel Files (*.xlsx)")
        if not path:
            return

        try:
            xls = pd.ExcelFile(path)
            # İstenen sayfa isimlerinden hangilerinin dosyada olduğunu bul
            sheets = sorted(list(REQ_SHEETS.intersection(set(xls.sheet_names))))

            if not sheets:
                QMessageBox.warning(self, "Uygun sayfa yok",
                                    "Seçilen dosyada istenen (SMD-OEE, ROBOT, DALGA_LEHİM) sheet bulunamadı.")
                self.reset_page()
                return

            self.main_window.excel_path = Path(path)
            self.lbl_path.setText(f"Seçilen Dosya: <b>{Path(path).name}</b>")

            # Bulunan sayfaları MainWindow'a kaydet
            self.main_window.available_sheets = sheets
            if "SMD-OEE" in sheets:
                self.main_window.selected_sheet = "SMD-OEE"  # SMD-OEE varsa varsayılan olarak seç
            elif sheets:
                self.main_window.selected_sheet = sheets[0]  # Yoksa ilk uygun sayfayı varsayılan olarak seç
            else:
                self.main_window.selected_sheet = None

            self.btn_daily_graphs.setEnabled(True)  # Enable daily graphs button
            self.btn_monthly_graphs.setEnabled(True)  # Enable monthly graphs button

            logging.info("Dosya seçildi: %s", path)

        except Exception as e:
            QMessageBox.critical(self, "Okuma hatası",
                                 f"Dosya okunurken bir hata oluştu: {e}\nLütfen dosyanın bozuk olmadığından ve Excel formatında olduğundan emin olun.")
            self.reset_page()

    def go_to_daily_graphs(self) -> None:
        """Günlük grafikler sayfasına geçer."""
        # Veri seçimi sayfasına gitmeden önce load_excel'i burada çağırmıyoruz,
        # DataSelectionPage'in refresh metodu çağıracak.
        self.main_window.goto_page(1)  # Veri seçimi sayfasına git (mevcut işlev)

    def go_to_monthly_graphs(self) -> None:
        """Aylık grafikler sayfasına geçer."""
        # Aylık grafikler için SMD-OEE sayfasının seçili olduğundan emin ol
        if "SMD-OEE" not in self.main_window.available_sheets:
            QMessageBox.warning(self, "Uyarı", "Aylık grafikler için 'SMD-OEE' sayfası Excel dosyasında bulunmalıdır.")
            return

        # Seçili sayfayı SMD-OEE olarak ayarla ve veriyi yükle
        self.main_window.selected_sheet = "SMD-OEE"
        self.main_window.load_excel()  # SMD-OEE verisini yükle

        self.main_window.goto_page(3)  # Aylık grafikler sayfasına git (yeni sayfa)

    def reset_page(self):
        """Sayfayı başlangıç durumuna döndürür."""
        self.main_window.excel_path = None
        self.main_window.selected_sheet = None
        self.main_window.available_sheets = []  # Reset available sheets
        self.lbl_path.setText("Henüz dosya seçilmedi")
        self.btn_daily_graphs.setEnabled(False)
        self.btn_monthly_graphs.setEnabled(False)


class DataSelectionPage(QWidget):
    """Veri seçimi sayfasını temsil eder (gruplama, metrikler vb. - Günlük Grafikler için)."""

    def __init__(self, main_window: "MainWindow") -> None:
        super().__init__()
        self.main_window = main_window
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setAlignment(Qt.AlignTop)

        title_label = QLabel("<h2>Günlük Grafik Veri Seçimi</h2>")
        title_label.setObjectName("title_label")
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # Yeni eklenen sayfa seçimi alanı
        sheet_selection_group = QHBoxLayout()
        self.sheet_selection_label = QLabel("İşlenecek Sayfa:")
        self.sheet_selection_label.setAlignment(Qt.AlignLeft)
        sheet_selection_group.addWidget(self.sheet_selection_label)

        self.cmb_sheet = QComboBox()
        self.cmb_sheet.setEnabled(False)  # Başlangıçta devre dışı
        self.cmb_sheet.currentIndexChanged.connect(self.on_sheet_selected)
        sheet_selection_group.addWidget(self.cmb_sheet)
        main_layout.addLayout(sheet_selection_group)

        # Gruplama değişkeni seçimi
        grouping_group = QHBoxLayout()
        grouping_group.addWidget(QLabel("<b>Gruplama Değişkeni (Tarihler):</b>"))
        self.cmb_grouping = QComboBox()
        self.cmb_grouping.currentIndexChanged.connect(self.populate_grouped)
        grouping_group.addWidget(self.cmb_grouping)
        main_layout.addLayout(grouping_group)

        # Gruplanan değişkenler (ürünler) seçimi
        grouped_group = QHBoxLayout()
        grouped_group.addWidget(QLabel("<b>Gruplanan Değişkenler (Ürünler):</b>"))
        self.lst_grouped = QListWidget()
        self.lst_grouped.setSelectionMode(QListWidget.MultiSelection)  # Çoklu seçim
        self.lst_grouped.itemSelectionChanged.connect(self.update_next_button_state)
        grouped_group.addWidget(self.lst_grouped)
        main_layout.addLayout(grouped_group)

        # Metrikler checkbox'ları
        metrics_group = QVBoxLayout()
        metrics_group.addWidget(QLabel("<b>Metrikler :</b>"))
        self.metrics_scroll_area = QScrollArea()
        self.metrics_scroll_area.setWidgetResizable(True)
        self.metrics_content_widget = QWidget()
        self.metrics_layout = QVBoxLayout(self.metrics_content_widget)
        self.metrics_scroll_area.setWidget(self.metrics_content_widget)
        metrics_group.addWidget(self.metrics_scroll_area)
        main_layout.addLayout(metrics_group)

        # Navigasyon butonları
        nav_layout = QHBoxLayout()
        self.btn_back = QPushButton("← Geri")
        self.btn_back.clicked.connect(lambda: self.main_window.goto_page(0))
        nav_layout.addWidget(self.btn_back)

        self.btn_next = QPushButton("İleri →")
        self.btn_next.setEnabled(False)  # Başlangıçta devre dışı
        self.btn_next.clicked.connect(self.go_next)
        nav_layout.addStretch(1)
        nav_layout.addWidget(self.btn_next)
        main_layout.addLayout(nav_layout)

    def refresh(self) -> None:
        """Sayfa her gösterildiğinde verileri yeniler."""
        # Sayfa seçimi ComboBox'ını doldur
        self.cmb_sheet.clear()
        if self.main_window.available_sheets:
            self.cmb_sheet.addItems(self.main_window.available_sheets)
            self.cmb_sheet.setEnabled(True)
            # Eğer daha önce bir sayfa seçilmişse onu ayarla
            if self.main_window.selected_sheet in self.main_window.available_sheets:
                self.cmb_sheet.setCurrentText(self.main_window.selected_sheet)
            else:
                # İlk uygun sayfayı varsayılan olarak seç ve sinyali tetikle
                self.cmb_sheet.setCurrentText(self.main_window.available_sheets[0])
            # setCurrentText zaten currentIndexChanged sinyalini tetikleyecek,
            # bu da on_sheet_selected'ı çağırıp load_excel'i çalıştıracak.
        else:
            self.cmb_sheet.setEnabled(False)
            self.main_window.selected_sheet = None
            QMessageBox.warning(self, "Uyarı", "Seçilen Excel dosyasında uygun sayfa bulunamadı.")
            self.main_window.goto_page(0)  # Dosya seçimine geri dön
            return

    def on_sheet_selected(self) -> None:
        """Sayfa seçimi değiştiğinde ana penceredeki seçimi günceller ve veriyi yeniden yükler."""
        self.main_window.selected_sheet = self.cmb_sheet.currentText()
        self.main_window.load_excel()  # Yeni seçilen sayfaya göre veriyi yeniden yükle
        self._populate_data_selection_fields()  # Sayfanın diğer alanlarını yeniden doldur

    def _populate_data_selection_fields(self):
        """Gruplama, gruplanan ve metrik alanlarını doldurur."""
        df = self.main_window.df
        if df.empty:
            QMessageBox.critical(self, "Hata", "Veri yüklenemedi. Lütfen dosyayı kontrol edin.")
            self.main_window.goto_page(0)  # Dosya seçimine geri dön
            return

        # Gruplama sütunu doldur
        self.cmb_grouping.clear()
        grouping_col_name = self.main_window.grouping_col_name
        if grouping_col_name and grouping_col_name in df.columns:
            grouping_vals = sorted(df[grouping_col_name].dropna().astype(str).unique())
            grouping_vals = [s for s in grouping_vals if s.strip()]  # Boş stringleri filtrele
            self.cmb_grouping.addItems(grouping_vals)
            if not grouping_vals:
                QMessageBox.warning(self, "Uyarı", "Gruplama sütunu (A) boş veya geçerli değer içermiyor.")
        else:
            QMessageBox.warning(self, "Uyarı", "Gruplama sütunu (A) bulunamadı veya boş.")
            self.cmb_grouping.clear()
            self.lst_grouped.clear()
            self.clear_metrics_checkboxes()
            self.update_next_button_state()  # Update button state after clearing
            return

        self.populate_metrics_checkboxes()
        self.populate_grouped()

    def populate_grouped(self) -> None:
        """Gruplanan değişkenler listesini (ürünler) doldurur."""
        self.lst_grouped.clear()
        selected_grouping_val = self.cmb_grouping.currentText()
        df = self.main_window.df

        if selected_grouping_val and self.main_window.grouping_col_name and self.main_window.grouped_col_name:
            # Optimize by filtering directly without `astype(str)` if already converted or if not needed
            filtered_df = df[df[self.main_window.grouping_col_name].astype(str) == selected_grouping_val]
            grouped_vals = sorted(filtered_df[self.main_window.grouped_col_name].dropna().astype(str).unique())
            grouped_vals = [s for s in grouped_vals if s.strip()]

            for gv in grouped_vals:
                item = QListWidgetItem(gv)
                item.setSelected(True)  # Varsayılan olarak hepsini seç
                self.lst_grouped.addItem(item)

        self.update_next_button_state()

    def populate_metrics_checkboxes(self):
        """Metrik sütunları için checkbox'ları oluşturur ve doldurur."""
        self.clear_metrics_checkboxes()

        self.main_window.selected_metrics = []

        if not self.main_window.metric_cols:
            empty_label = QLabel("Seçilebilir metrik bulunamadı.", parent=self.metrics_content_widget)
            empty_label.setAlignment(Qt.AlignCenter)
            self.metrics_layout.addWidget(empty_label)
            self.btn_next.setEnabled(False)
            return

        for col_name in self.main_window.metric_cols:
            checkbox = QCheckBox(col_name)
            is_entirely_empty = self.main_window.df[col_name].dropna().empty

            if is_entirely_empty:
                checkbox.setChecked(False)
                checkbox.setEnabled(False)
                checkbox.setText(f"{col_name} (Boş)")
                checkbox.setStyleSheet("color: gray;")
            else:
                checkbox.setChecked(True)
                self.main_window.selected_metrics.append(col_name)

            checkbox.stateChanged.connect(self.on_metric_checkbox_changed)
            self.metrics_layout.addWidget(checkbox)

        # Add a stretch to push checkboxes to the top
        self.metrics_layout.addStretch(1)
        self.update_next_button_state()

    def clear_metrics_checkboxes(self):
        """Metrik checkbox'larını temizler."""
        while self.metrics_layout.count():
            item = self.metrics_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
            elif isinstance(item, QSpacerItem):  # Also remove spacers if any
                self.metrics_layout.removeItem(item)

    def on_metric_checkbox_changed(self, state):
        """Bir metrik checkbox'ının durumu değiştiğinde çağrılır."""
        sender_checkbox = self.sender()
        metric_name = sender_checkbox.text().replace(" (Boş)", "")

        if state == Qt.Checked:
            if metric_name not in self.main_window.selected_metrics:
                self.main_window.selected_metrics.append(metric_name)
        else:
            if metric_name in self.main_window.selected_metrics:
                self.main_window.selected_metrics.remove(metric_name)

        self.update_next_button_state()

    def update_next_button_state(self):
        """İleri butonunun etkinleştirme durumunu günceller."""
        is_grouped_selected = bool(self.lst_grouped.selectedItems())
        is_metric_selected = bool(self.main_window.selected_metrics)
        self.btn_next.setEnabled(is_grouped_selected and is_metric_selected)

    def go_next(self) -> None:
        """Bir sonraki sayfaya geçmek için verileri hazırlar."""
        self.main_window.grouped_values = [i.text() for i in self.lst_grouped.selectedItems()]
        self.main_window.selected_grouping_val = self.cmb_grouping.currentText()
        if not self.main_window.grouped_values or not self.main_window.selected_metrics:
            QMessageBox.warning(self, "Seçim Eksik", "Lütfen en az bir gruplanan değişken ve bir metrik seçin.")
            return
        self.main_window.goto_page(2)  # Grafik sayfasına git (Günlük Grafiklerin gösterildiği sayfa)


class DailyGraphsPage(QWidget):
    """Oluşturulan günlük grafikleri gösteren ve kaydetme seçenekleri sunan sayfa."""

    def __init__(self, main_window: "MainWindow") -> None:
        super().__init__()
        self.main_window = main_window
        self.worker: GraphWorker | None = None
        self.figures_data: List[Tuple[str, Figure, str]] = []  # (Ürün adı, figür objesi, OEE değeri)
        self.current_page = 0  # Mevcut sayfa numarası
        self.current_graph_type = "Donut"  # Varsayılan grafik tipi
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setAlignment(Qt.AlignTop)

        title_label = QLabel("<h2>Günlük Grafikler</h2>")
        title_label.setObjectName("title_label")
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        self.progress = QProgressBar()
        self.progress.setAlignment(Qt.AlignCenter)
        self.progress.setTextVisible(True)
        self.progress.hide()  # Başlangıçta gizle
        main_layout.addWidget(self.progress)

        # Üst navigasyon ve kaydet butonu
        nav_top = QHBoxLayout()
        self.btn_back = QPushButton("← Veri Seçimi")
        self.btn_back.clicked.connect(lambda: self.main_window.goto_page(1))
        nav_top.addWidget(self.btn_back)

        self.lbl_chart_info = QLabel("")
        self.lbl_chart_info.setAlignment(Qt.AlignCenter)
        self.lbl_chart_info.setStyleSheet("font-weight: bold; font-size: 12pt;")
        nav_top.addWidget(self.lbl_chart_info)

        # Grafik tipi seçimi için ComboBox
        self.cmb_graph_type = QComboBox()
        self.cmb_graph_type.addItems(["Donut", "Bar"])
        self.cmb_graph_type.setCurrentText(self.current_graph_type)  # Varsayılanı ayarla
        self.cmb_graph_type.currentIndexChanged.connect(self.on_graph_type_changed)
        nav_top.addWidget(self.cmb_graph_type)

        nav_top.addStretch(1)
        self.lbl_page = QLabel("Sayfa 0 / 0")
        self.lbl_page.setAlignment(Qt.AlignCenter)
        nav_top.addWidget(self.lbl_page)
        nav_top.addStretch(1)

        self.btn_save_image = QPushButton("Grafiği Kaydet (PNG/JPEG)")
        self.btn_save_image.clicked.connect(self.save_single_graph_as_image)
        self.btn_save_image.setEnabled(False)
        nav_top.addWidget(self.btn_save_image)
        main_layout.addLayout(nav_top)

        # Grafiklerin gösterileceği kaydırılabilir alan
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.canvas_holder = QWidget()
        self.canvas_centered_layout = QHBoxLayout(self.canvas_holder)
        self.vbox_canvases = QVBoxLayout()
        self.canvas_centered_layout.addStretch(1)
        self.canvas_centered_layout.addLayout(self.vbox_canvases)
        self.canvas_centered_layout.addStretch(1)

        self.scroll.setWidget(self.canvas_holder)
        main_layout.addWidget(self.scroll)

        # Alt navigasyon butonları (önceki/sonraki sayfa)
        nav_bottom = QHBoxLayout()
        nav_bottom.addStretch(1)
        self.btn_prev = QPushButton("← Önceki Sayfa")
        self.btn_prev.clicked.connect(self.prev_page)
        self.btn_prev.setEnabled(False)
        nav_bottom.addWidget(self.btn_prev)

        self.btn_next = QPushButton("Sonraki Sayfa →")
        self.btn_next.clicked.connect(self.next_page)
        self.btn_next.setEnabled(False)
        nav_bottom.addWidget(self.btn_next)
        nav_bottom.addStretch(1)
        main_layout.addLayout(nav_bottom)

    def on_graph_type_changed(self, index: int) -> None:
        """Grafik tipi değiştiğinde çağrılır ve grafikleri yeniden çizer."""
        self.current_graph_type = self.cmb_graph_type.currentText()
        self.enter_page()  # Re-generate graphs with the new type

    def on_results(self, results: List[Tuple[str, pd.Series, str]]) -> None:
        """GraphWorker'dan gelen sonuçları işler ve grafikleri oluşturur."""
        self.progress.setValue(100)
        self.progress.hide()

        if not results:
            QMessageBox.information(self, "Veri yok", "Grafik oluşturulamadı. Seçilen kriterlere göre veri bulunamadı.")
            self.btn_save_image.setEnabled(False)
            self.lbl_chart_info.setText("Gösterilecek grafik bulunamadı.")
            return

        self.figures_data.clear()

        fig_width_inches = 700 / 100
        fig_height_inches = 460 / 100

        for grouped_val, metric_sums, oee_display_value in results:
            fig, ax = plt.subplots(figsize=(fig_width_inches, fig_height_inches))
            background_color = 'white'
            fig.patch.set_facecolor(background_color)
            ax.set_facecolor(background_color)

            sorted_metrics_series = metric_sums.sort_values(ascending=False) if not metric_sums.empty else pd.Series()

            num_metrics = len(sorted_metrics_series)
            if num_metrics == 1 and sorted_metrics_series.index[0] == 'HAT ÇALIŞMADI':
                chart_colors = ['#FF9841']
            else:
                colors_palette = matplotlib.colormaps.get_cmap('tab20')
                chart_colors = [colors_palette(i % 20) for i in range(num_metrics)] if num_metrics > 0 else []

            if self.current_graph_type == "Donut":
                GraphPlotter.create_donut_chart(ax, sorted_metrics_series, oee_display_value, chart_colors, fig)
            elif self.current_graph_type == "Bar":
                GraphPlotter.create_bar_chart(ax, sorted_metrics_series, oee_display_value, chart_colors)

            total_duration_seconds = sorted_metrics_series.sum()
            total_duration_hours = int(total_duration_seconds // 3600)
            total_duration_minutes = int((total_duration_seconds % 3600) // 60)
            total_duration_text = f"TOPLAM DURUŞ\n{total_duration_hours} SAAT {total_duration_minutes} DAKİKA"

            fig.text(0.01, 0.05, total_duration_text, transform=fig.transFigure,
                     fontsize=14, fontweight='bold', verticalalignment='bottom')

            self.figures_data.append((grouped_val, fig, oee_display_value))
            plt.close(fig)

        self.display_current_page_graphs()
        if self.figures_data:
            self.btn_save_image.setEnabled(True)

    def enter_page(self) -> None:
        """Bu sayfaya girildiğinde grafikleri yeniden oluşturma sürecini başlatır."""
        self.figures_data.clear()
        self.clear_canvases()
        self.progress.setValue(0)
        self.progress.show()  # Show progress bar when starting
        self.btn_save_image.setEnabled(False)
        self.lbl_chart_info.setText("Grafikler oluşturuluyor...")
        self.update_page_label()  # Reset page label
        self.update_navigation_buttons()  # Disable navigation buttons

        if self.worker and self.worker.isRunning():
            self.worker.quit()
            self.worker.wait()

        self.worker = GraphWorker(
            df=self.main_window.df,
            grouping_col_name=self.main_window.grouping_col_name,
            grouped_col_name=self.main_window.grouped_col_name,
            grouped_values=self.main_window.grouped_values,
            metric_cols=self.main_window.selected_metrics,
            oee_col_name=self.main_window.oee_col_name,
            selected_grouping_val=self.main_window.selected_grouping_val
        )
        self.worker.finished.connect(self.on_results)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_error(self, message: str) -> None:
        """GraphWorker'dan gelen hata mesajını gösterir."""
        QMessageBox.critical(self, "Hata", message)
        self.progress.setValue(0)
        self.progress.hide()
        self.lbl_chart_info.setText("Grafik oluşturma hatası.")
        self.btn_save_image.setEnabled(False)

    def clear_canvases(self) -> None:
        """Mevcut grafik tuvallerini temizler."""
        while self.vbox_canvases.count():
            item = self.vbox_canvases.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def display_current_page_graphs(self) -> None:
        """Mevcut sayfadaki grafikleri gösterir."""
        self.clear_canvases()

        total_pages = (len(self.figures_data) + GRAPHS_PER_PAGE - 1) // GRAPHS_PER_PAGE if self.figures_data else 0

        # Ensure current_page is within valid range
        if self.current_page >= total_pages and total_pages > 0:
            self.current_page = total_pages - 1
        elif total_pages == 0:
            self.current_page = 0

        start_index = self.current_page * GRAPHS_PER_PAGE
        end_index = start_index + GRAPHS_PER_PAGE

        graphs_to_display = self.figures_data[start_index:end_index]

        if not graphs_to_display:
            self.lbl_chart_info.setText("Gösterilecek grafik bulunamadı.")
            self.btn_save_image.setEnabled(False)
            self.update_page_label()
            self.update_navigation_buttons()
            return

        for grouped_val, fig, oee_display_value in graphs_to_display:
            canvas = FigureCanvas(fig)
            canvas.setFixedSize(700, 460)
            self.vbox_canvases.addWidget(canvas)
            self.lbl_chart_info.setText(f"{self.main_window.selected_grouping_val} - {grouped_val}")

        self.update_page_label()
        self.update_navigation_buttons()
        self.btn_save_image.setEnabled(True)

    def update_page_label(self) -> None:
        """Sayfa etiketini günceller."""
        total_pages = (len(self.figures_data) + GRAPHS_PER_PAGE - 1) // GRAPHS_PER_PAGE if self.figures_data else 0
        self.lbl_page.setText(f"Sayfa {self.current_page + 1} / {total_pages}")

    def update_navigation_buttons(self) -> None:
        """Gezinme butonlarının etkinleştirme durumunu günceller."""
        total_pages = (len(self.figures_data) + GRAPHS_PER_PAGE - 1) // GRAPHS_PER_PAGE if self.figures_data else 0
        self.btn_prev.setEnabled(self.current_page > 0)
        self.btn_next.setEnabled(self.current_page < total_pages - 1)

    def prev_page(self) -> None:
        """Önceki sayfaya geçer."""
        if self.current_page > 0:
            self.current_page -= 1
            self.display_current_page_graphs()

    def next_page(self) -> None:
        """Sonraki sayfaya geçer."""
        total_pages = (len(self.figures_data) + GRAPHS_PER_PAGE - 1) // GRAPHS_PER_PAGE if self.figures_data else 0
        if self.current_page < total_pages - 1:
            self.current_page += 1
            self.display_current_page_graphs()

    def save_single_graph_as_image(self) -> None:
        """Mevcut sayfadaki grafiği PNG/JPEG olarak kaydeder."""
        if not self.figures_data:
            QMessageBox.warning(self, "Kaydedilecek Grafik Yok", "Görüntülenecek bir grafik bulunmamaktadır.")
            return

        total_graphs = len(self.figures_data)
        fig_index_on_page = self.current_page * GRAPHS_PER_PAGE

        if not (0 <= fig_index_on_page < total_graphs):
            QMessageBox.warning(self, "Geçersiz Sayfa", "Mevcut sayfada kaydedilecek bir grafik yok.")
            return

        grouped_val, fig, _ = self.figures_data[fig_index_on_page]

        default_filename = f"grafik_{grouped_val}_{self.main_window.selected_grouping_val}_{self.current_graph_type}.png".replace(
            " ", "_").replace("/", "-")
        filepath, _ = QFileDialog.getSaveFileName(
            self, "Grafiği Kaydet", default_filename, "PNG (*.png);;JPEG (*.jpeg);;JPG (*.jpg)"
        )

        if filepath:
            try:
                fig.savefig(filepath, dpi=plt.rcParams['savefig.dpi'], bbox_inches='tight',
                            facecolor=fig.get_facecolor())
                QMessageBox.information(self, "Kaydedildi", f"Grafik başarıyla kaydedildi: {Path(filepath).name}")
                logging.info("Grafik kaydedildi: %s", filepath)
            except Exception as e:
                QMessageBox.critical(self, "Kaydetme Hatası", f"Grafik kaydedilirken bir hata oluştu: {e}")
                logging.exception("Grafik kaydetme hatası.")


class MonthlyGraphWorker(QThread):
    """Aylık grafik oluşturma işlemlerini arka planda yürüten işçi sınıfı."""
    finished = pyqtSignal(list)  # List of (hat_name, figure_object)
    progress = pyqtSignal(int)
    error = pyqtSignal(str)

    def __init__(self, df: pd.DataFrame, grouping_col_name: str, grouped_col_name: str, oee_col_name: str,
                 prev_year_oee: float | None, prev_month_oee: float | None):
        super().__init__()
        self.df = df.copy()
        self.grouping_col_name = grouping_col_name
        self.grouped_col_name = grouped_col_name
        self.oee_col_name = oee_col_name
        self.prev_year_oee = prev_year_oee
        self.prev_month_oee = prev_month_oee

    def run(self):
        try:
            figures_data: List[Tuple[str, Figure]] = []

            df_smd_oee = self.df
            logging.info(f"MonthlyGraphWorker: Başlangıç veri çerçevesi boyutu: {df_smd_oee.shape}")

            # Rename columns for internal consistency within this function
            # Check if columns exist before renaming
            col_mapping = {}
            if self.grouped_col_name in df_smd_oee.columns:
                col_mapping[self.grouped_col_name] = 'U_Agaci_Sev'
            if self.grouping_col_name in df_smd_oee.columns:
                col_mapping[self.grouping_col_name] = 'Tarih'
            if self.oee_col_name and self.oee_col_name in df_smd_oee.columns:
                col_mapping[self.oee_col_name] = 'OEE_Degeri'

            if col_mapping:
                df_smd_oee.rename(columns=col_mapping, inplace=True)
                logging.info(f"MonthlyGraphWorker: Sütunlar yeniden adlandırıldı: {col_mapping}")
            else:
                self.error.emit("Gerekli sütunlar (Ürün, Tarih, OEE) Excel dosyasında bulunamadı veya adlandırılamadı.")
                return

            # Convert 'Tarih' to datetime and filter out NaT
            if 'Tarih' in df_smd_oee.columns:
                df_smd_oee['Tarih'] = pd.to_datetime(df_smd_oee['Tarih'], errors='coerce')
                df_smd_oee.dropna(subset=['Tarih'], inplace=True)
                logging.info(
                    f"MonthlyGraphWorker: Tarih sütunu datetime'a dönüştürüldü ve NaT değerleri temizlendi. Yeni boyut: {df_smd_oee.shape}")
            else:
                self.error.emit("'Tarih' sütunu bulunamadı.")
                return

            # Convert 'OEE_Degeri' to float using pd.to_numeric
            if 'OEE_Degeri' in df_smd_oee.columns:
                # Replace comma with dot for decimal conversion and handle non-numeric values
                df_smd_oee['OEE_Degeri'] = pd.to_numeric(
                    df_smd_oee['OEE_Degeri'].astype(str).str.replace('%', '').str.replace(',', '.'),
                    errors='coerce'
                )
                df_smd_oee.dropna(subset=['OEE_Degeri'], inplace=True)
                logging.info(
                    f"MonthlyGraphWorker: OEE_Degeri sütunu float'a dönüştürüldü ve geçersiz değerler temizlendi. Yeni boyut: {df_smd_oee.shape}")
            else:
                self.error.emit("'OEE_Degeri' sütunu bulunamadı.")
                return

            # Extract 'Group_Key' (e.g., "HAT-4")
            def extract_group_key(s):
                s = str(s).upper()
                match = re.search(r'HAT(\d+)', s)
                if match:
                    hat_number = match.group(1)
                    return f"HAT-{hat_number}"
                return None

            if 'U_Agaci_Sev' in df_smd_oee.columns:
                df_smd_oee['Group_Key'] = df_smd_oee['U_Agaci_Sev'].apply(extract_group_key)
                df_smd_oee.dropna(subset=['Group_Key'],
                                  inplace=True)  # Remove rows where Group_Key couldn't be extracted
                logging.info(
                    f"MonthlyGraphWorker: Group_Key sütunu oluşturuldu ve boş değerler temizlendi. Yeni boyut: {df_smd_oee.shape}")
            else:
                self.error.emit("'U_Agaci_Sev' sütunu bulunamadı.")
                return

            unique_hats = sorted(df_smd_oee['Group_Key'].unique())

            # Filter unique_hats to include only HAT-1, HAT-2, HAT-3, HAT-4
            target_hat_patterns = {"HAT-1", "HAT-2", "HAT-3", "HAT-4"}
            filtered_hats = [hat for hat in unique_hats if hat in target_hat_patterns]
            unique_hats = sorted(filtered_hats)
            logging.info(f"MonthlyGraphWorker: Hedef hatlar filtrelendi: {unique_hats}")

            total_hats = len(unique_hats)

            if not unique_hats:
                self.error.emit(
                    "Grafik oluşturulacak hat verisi bulunamadı. Lütfen Excel dosyasında 'HAT-1', 'HAT-2', 'HAT-3' veya 'HAT-4' içeren verilerin olduğundan emin olun.")
                return

            for i, selected_hat in enumerate(unique_hats):
                logging.info(f"MonthlyGraphWorker: '{selected_hat}' için grafik oluşturuluyor...")
                df_smd_oee_filtered_by_hat = df_smd_oee[df_smd_oee['Group_Key'] == selected_hat].copy()

                if df_smd_oee_filtered_by_hat.empty:
                    logging.warning(
                        f"MonthlyGraphWorker: Seçilen '{selected_hat}' hattı için veri bulunamadı, atlanıyor.")
                    self.progress.emit(int((i + 1) / total_hats * 100))
                    continue

                # Group by date and calculate mean OEE for the selected hat
                grouped_oee = df_smd_oee_filtered_by_hat.groupby(pd.Grouper(key='Tarih', freq='D'))[
                    'OEE_Degeri'].mean().reset_index()

                if grouped_oee.empty:
                    logging.warning(
                        f"MonthlyGraphWorker: Seçilen '{selected_hat}' hattı için günlük OEE ortalaması bulunamadı, atlanıyor.")
                    self.progress.emit(int((i + 1) / total_hats * 100))
                    continue

                grouped_oee_sorted = grouped_oee.sort_values(by='Tarih')

                dates = grouped_oee_sorted['Tarih']
                oee_values = grouped_oee_sorted['OEE_Degeri']

                fig, ax = plt.subplots(figsize=(10, 6), dpi=100)
                ax.set_facecolor('#f9f9f9')
                fig.patch.set_facecolor('#f0f2f5')

                ax.spines['top'].set_visible(False)
                ax.spines['right'].set_visible(False)
                ax.spines['left'].set_visible(False)
                ax.spines['bottom'].set_visible(False)
                ax.grid(False)

                line_color = '#1f77b4'

                ax.plot(dates, oee_values, marker='o', markersize=8, color=line_color, linewidth=2, label=selected_hat)
                ax.plot(dates, oee_values, 'o', markersize=6, color='white', markeredgecolor=line_color,
                        markeredgewidth=1.5, zorder=5)

                for x, y in zip(dates, oee_values):
                    ax.annotate(f'{y:.0f}%', (x, y), textcoords="offset points", xytext=(0, 10), ha='center',
                                fontsize=8, fontweight='bold')

                overall_calculated_average = np.mean(oee_values) if not oee_values.empty else 0

                if self.prev_year_oee is not None:
                    ax.axhline(self.prev_year_oee, color='red', linestyle='--', linewidth=1.5,
                               label=f'Önceki Yıl OEE ({self.prev_year_oee:.1f}%)')
                if self.prev_month_oee is not None:
                    ax.axhline(self.prev_month_oee, color='orange', linestyle='--', linewidth=1.5,
                               label=f'Önceki Ay OEE ({self.prev_month_oee:.1f}%)')
                if overall_calculated_average > 0:
                    ax.axhline(overall_calculated_average, color='purple', linestyle='--', linewidth=1.5,
                               label=f'Hesaplanan Ortalama OEE ({overall_calculated_average:.1f}%)')

                ax.xaxis.set_major_formatter(plt.matplotlib.dates.DateFormatter('%d.%m.%Y'))
                fig.autofmt_xdate(rotation=45)

                ax.yaxis.set_major_formatter(PercentFormatter())
                ax.set_ylim(bottom=0)

                ax.set_xlabel("Tarih", fontsize=12, fontweight='bold')
                ax.set_ylabel("OEE (%)", fontsize=12, fontweight='bold')

                first_date_in_data = dates.min()
                month_name = first_date_in_data.strftime('%B').capitalize()
                chart_title = f"{selected_hat} {month_name} OEE"
                ax.set_title(chart_title, fontsize=16, color='#2c3e50', fontweight='bold')

                ax.legend(loc='upper left', bbox_to_anchor=(1, 1), fontsize=10)

                figures_data.append((selected_hat, fig))
                plt.close(fig)  # Close the figure to free up memory

                self.progress.emit(int((i + 1) / total_hats * 100))
                logging.info(
                    f"MonthlyGraphWorker: '{selected_hat}' için grafik oluşturuldu. İlerleme: {int((i + 1) / total_hats * 100)}%")

            self.finished.emit(figures_data)
            logging.info("MonthlyGraphWorker: Tüm aylık grafikler başarıyla oluşturuldu.")
        except Exception as exc:
            logging.exception("MonthlyGraphWorker hatası oluştu.")
            self.error.emit(f"Aylık grafik oluşturulurken bir hata oluştu: {str(exc)}")


class MonthlyGraphsPage(QWidget):
    """Aylık grafikler ve veri seçimi sayfasını temsil eder."""

    def __init__(self, main_window: "MainWindow") -> None:
        super().__init__()
        self.main_window = main_window

        # Initialize chart container and layout first to avoid AttributeError
        self.monthly_chart_container = QFrame(objectName="chartContainer")
        self.monthly_chart_layout = QVBoxLayout(self.monthly_chart_container)
        self.monthly_chart_layout.setAlignment(Qt.AlignCenter)  # Center the content

        self.current_monthly_chart_figure = None  # To store the monthly chart for saving
        self.figures_data_monthly: List[Tuple[str, Figure]] = []  # (Hat adı, figür objesi)
        self.current_page_monthly = 0
        self.monthly_worker: MonthlyGraphWorker | None = None

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setAlignment(Qt.AlignTop)

        title_label = QLabel("<h2>Aylık Grafikler ve Veri Seçimi</h2>")
        title_label.setObjectName("title_label")
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # Grafik tipi seçimi
        graph_type_selection_layout = QHBoxLayout()
        graph_type_selection_layout.addWidget(QLabel("<b>Grafik Tipi:</b>"))
        self.cmb_monthly_graph_type = QComboBox()
        self.cmb_monthly_graph_type.addItems(["OEE Grafikleri", "Dizgi Duruş Grafiği", "Dizgi Onay Dağılım Grafiği"])
        self.cmb_monthly_graph_type.currentIndexChanged.connect(self.on_monthly_graph_type_changed)
        graph_type_selection_layout.addWidget(self.cmb_monthly_graph_type)
        main_layout.addLayout(graph_type_selection_layout)

        # OEE Grafikleri için özel alanlar
        self.oee_options_widget = QWidget()
        oee_options_layout = QVBoxLayout(self.oee_options_widget)

        # Önceki Yılın OEE Değeri
        prev_year_oee_layout = QHBoxLayout()
        prev_year_oee_layout.addWidget(QLabel("Önceki Yılın OEE Değeri (%):"))
        self.txt_prev_year_oee = QLineEdit()
        self.txt_prev_year_oee.setPlaceholderText("Örn: 85.5")
        prev_year_oee_layout.addWidget(self.txt_prev_year_oee)
        oee_options_layout.addLayout(prev_year_oee_layout)

        # Önceki Ayın OEE Değeri
        prev_month_oee_layout = QHBoxLayout()
        prev_month_oee_layout.addWidget(QLabel("Önceki Ayın OEE Değeri (%):"))
        self.txt_prev_month_oee = QLineEdit()
        self.txt_prev_month_oee.setPlaceholderText("Örn: 82.0")
        prev_month_oee_layout.addWidget(self.txt_prev_month_oee)
        oee_options_layout.addLayout(prev_month_oee_layout)

        # Hat Grafikleri ve Sayfa Grafikleri butonları
        oee_buttons_layout = QHBoxLayout()
        self.btn_line_chart = QPushButton("Hat Grafikleri")
        self.btn_line_chart.clicked.connect(self._start_monthly_graph_worker)
        self.btn_line_chart.setEnabled(False)  # Initially disabled until a graph type is selected
        oee_buttons_layout.addWidget(self.btn_line_chart)

        self.btn_page_chart = QPushButton("Sayfa Grafikleri")
        self.btn_page_chart.setEnabled(False)  # Not implemented yet
        oee_buttons_layout.addWidget(self.btn_page_chart)
        oee_options_layout.addLayout(oee_buttons_layout)

        main_layout.addWidget(self.oee_options_widget)

        # Diğer grafik tipleri için placeholder (şimdilik gizli)
        self.other_graphs_widget = QWidget()
        other_graphs_layout = QVBoxLayout(self.other_graphs_widget)
        other_graphs_layout.addWidget(QLabel("Bu grafik tipi için seçenekler burada olacak."))
        main_layout.addWidget(self.other_graphs_widget)
        self.other_graphs_widget.hide()

        # Grafik görüntüleme alanı
        # self.monthly_chart_container ve self.monthly_chart_layout zaten __init__ içinde tanımlandı
        self.monthly_chart_container.setMinimumHeight(460)
        main_layout.addWidget(self.monthly_chart_container)

        self.monthly_progress = QProgressBar()
        self.monthly_progress.setAlignment(Qt.AlignCenter)
        self.monthly_progress.setTextVisible(True)
        self.monthly_progress.hide()
        main_layout.addWidget(self.monthly_progress)

        # Alt navigasyon butonları (önceki/sonraki sayfa)
        nav_bottom = QHBoxLayout()
        self.btn_monthly_back = QPushButton("← Geri")
        self.btn_monthly_back.clicked.connect(lambda: self.main_window.goto_page(0))
        nav_bottom.addWidget(self.btn_monthly_back)

        self.lbl_monthly_page = QLabel("Sayfa 0 / 0")
        self.lbl_monthly_page.setAlignment(Qt.AlignCenter)
        nav_bottom.addWidget(self.lbl_monthly_page)

        self.btn_prev_monthly = QPushButton("← Önceki Hat")
        self.btn_prev_monthly.clicked.connect(self.prev_monthly_page)
        self.btn_prev_monthly.setEnabled(False)
        nav_bottom.addWidget(self.btn_prev_monthly)

        self.btn_next_monthly = QPushButton("Sonraki Hat →")
        self.btn_next_monthly.clicked.connect(self.next_monthly_page)
        self.btn_next_monthly.setEnabled(False)
        nav_bottom.addWidget(self.btn_next_monthly)

        self.btn_save_monthly_chart = QPushButton("Grafiği Kaydet (PNG/JPEG)")
        self.btn_save_monthly_chart.clicked.connect(self._save_monthly_chart_as_image)
        self.btn_save_monthly_chart.setEnabled(False)
        nav_bottom.addStretch(1)
        nav_bottom.addWidget(self.btn_save_monthly_chart)
        main_layout.addLayout(nav_bottom)

        self.on_monthly_graph_type_changed(0)  # Set initial visibility

    def enter_page(self) -> None:
        """Bu sayfaya girildiğinde grafiği temizler ve buton durumlarını günceller."""
        self.clear_monthly_chart_canvas()
        self.btn_save_monthly_chart.setEnabled(False)
        self.update_monthly_page_label()
        self.update_monthly_navigation_buttons()
        # Enable line chart button if OEE Graphs is selected
        if self.cmb_monthly_graph_type.currentText() == "OEE Grafikleri":
            self.btn_line_chart.setEnabled(True)
        else:
            self.btn_line_chart.setEnabled(False)

    def on_monthly_graph_type_changed(self, index: int):
        """Aylık grafik tipi seçimi değiştiğinde ilgili seçenekleri gösterir/gizler."""
        selected_type = self.cmb_monthly_graph_type.currentText()
        if selected_type == "OEE Grafikleri":
            self.oee_options_widget.show()
            self.other_graphs_widget.hide()
            self.btn_line_chart.setEnabled(True)  # Enable line chart button for OEE
        else:
            self.oee_options_widget.hide()
            self.other_graphs_widget.show()
            self.btn_line_chart.setEnabled(False)  # Disable for other types

        # Clear existing chart when graph type changes
        self.clear_monthly_chart_canvas()
        self.btn_save_monthly_chart.setEnabled(False)
        self.figures_data_monthly.clear()  # Clear stored figures
        self.current_page_monthly = 0
        self.update_monthly_page_label()
        self.update_monthly_navigation_buttons()

    def clear_monthly_chart_canvas(self):
        """Aylık grafik tuvallerini temizler."""
        while self.monthly_chart_layout.count():
            item = self.monthly_chart_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def _start_monthly_graph_worker(self):
        """Aylık grafik oluşturma işçisini başlatır."""
        self.clear_monthly_chart_canvas()
        self.btn_save_monthly_chart.setEnabled(False)
        self.figures_data_monthly.clear()
        self.current_page_monthly = 0
        self.monthly_progress.setValue(0)
        self.monthly_progress.show()
        self.update_monthly_page_label()
        self.update_monthly_navigation_buttons()

        if self.monthly_worker and self.monthly_worker.isRunning():
            self.monthly_worker.quit()
            self.monthly_worker.wait()

        try:
            prev_year_oee = float(
                self.txt_prev_year_oee.text().replace(",", ".")) if self.txt_prev_year_oee.text() else None
            prev_month_oee = float(
                self.txt_prev_month_oee.text().replace(",", ".")) if self.txt_prev_month_oee.text() else None
        except ValueError:
            QMessageBox.warning(self, "Geçersiz Giriş",
                                "Lütfen Önceki Yıl/Ay OEE değerlerini geçerli sayı olarak girin.")
            self.monthly_progress.hide()
            return

        self.monthly_worker = MonthlyGraphWorker(
            df=self.main_window.df,
            grouping_col_name=self.main_window.grouping_col_name,
            grouped_col_name=self.main_window.grouped_col_name,
            oee_col_name=self.main_window.oee_col_name,
            prev_year_oee=prev_year_oee,
            prev_month_oee=prev_month_oee
        )
        self.monthly_worker.finished.connect(self._on_monthly_graphs_generated)
        self.monthly_worker.progress.connect(self.monthly_progress.setValue)
        self.monthly_worker.error.connect(self._on_monthly_graph_error)
        self.monthly_worker.start()

    def _on_monthly_graphs_generated(self, figures_data: List[Tuple[str, Figure]]):
        """MonthlyGraphWorker'dan gelen sonuçları işler."""
        self.monthly_progress.setValue(100)
        self.monthly_progress.hide()

        if not figures_data:
            QMessageBox.information(self, "Veri Yok",
                                    "Aylık grafik oluşturulamadı. Seçilen kriterlere göre veri bulunamadı.")
            self.btn_save_monthly_chart.setEnabled(False)
            return

        self.figures_data_monthly = figures_data
        self.display_current_page_graphs_monthly()
        self.btn_save_monthly_chart.setEnabled(True)

    def _on_monthly_graph_error(self, message: str):
        """MonthlyGraphWorker'dan gelen hata mesajını gösterir."""
        QMessageBox.critical(self, "Hata", message)
        self.monthly_progress.setValue(0)
        self.monthly_progress.hide()
        self.btn_save_monthly_chart.setEnabled(False)

    def display_current_page_graphs_monthly(self) -> None:
        """Mevcut sayfadaki aylık grafiği gösterir."""
        self.clear_monthly_chart_canvas()

        total_pages = len(self.figures_data_monthly)

        # Ensure current_page_monthly is within valid range
        if self.current_page_monthly >= total_pages and total_pages > 0:
            self.current_page_monthly = total_pages - 1
        elif total_pages == 0:
            self.current_page_monthly = 0

        if not self.figures_data_monthly:
            no_data_label = QLabel("Gösterilecek aylık grafik bulunamadı.", alignment=Qt.AlignCenter)
            self.monthly_chart_layout.addWidget(no_data_label)
            self.current_monthly_chart_figure = None
            self.btn_save_monthly_chart.setEnabled(False)
            self.update_monthly_page_label()
            self.update_monthly_navigation_buttons()
            return

        hat_name, fig = self.figures_data_monthly[self.current_page_monthly]

        canvas = FigureCanvas(fig)
        canvas.setFixedSize(700, 460)  # Fixed size for consistency
        self.monthly_chart_layout.addWidget(canvas, stretch=1)
        canvas.draw()

        self.current_monthly_chart_figure = fig  # Store for saving
        self.btn_save_monthly_chart.setEnabled(True)
        self.update_monthly_page_label()
        self.update_monthly_navigation_buttons()

    def update_monthly_page_label(self) -> None:
        """Aylık grafik sayfa etiketini günceller."""
        total_pages = len(self.figures_data_monthly)
        self.lbl_monthly_page.setText(f"Sayfa {self.current_page_monthly + 1} / {total_pages}")

    def update_monthly_navigation_buttons(self) -> None:
        """Aylık grafik gezinme butonlarının etkinleştirme durumunu günceller."""
        total_pages = len(self.figures_data_monthly)
        self.btn_prev_monthly.setEnabled(self.current_page_monthly > 0)
        self.btn_next_monthly.setEnabled(self.current_page_monthly < total_pages - 1)

    def prev_monthly_page(self) -> None:
        """Önceki aylık grafik sayfasına geçer."""
        if self.current_page_monthly > 0:
            self.current_page_monthly -= 1
            self.display_current_page_graphs_monthly()

    def next_monthly_page(self) -> None:
        """Sonraki aylık grafik sayfasına geçer."""
        total_pages = len(self.figures_data_monthly)
        if self.current_page_monthly < total_pages - 1:
            self.current_page_monthly += 1
            self.display_current_page_graphs_monthly()

    def _save_monthly_chart_as_image(self):
        """Aylık grafiği PNG/JPEG olarak kaydeder."""
        if self.current_monthly_chart_figure is None:
            QMessageBox.warning(self, "Kaydedilecek Grafik Yok", "Görüntülenecek bir aylık grafik bulunmamaktadır.")
            return

        # Get the hat name for the current chart to use in the filename
        current_hat_name = "grafik"
        if self.figures_data_monthly and 0 <= self.current_page_monthly < len(self.figures_data_monthly):
            current_hat_name = self.figures_data_monthly[self.current_page_monthly][0].replace(" ", "_").replace("/",
                                                                                                                 "-")

        default_filename = f"aylik_oee_{current_hat_name}.png"
        filepath, _ = QFileDialog.getSaveFileName(
            self, "Aylık Grafiği Kaydet", default_filename, "PNG (*.png);;JPEG (*.jpeg);;JPG (*.jpg)"
        )

        if filepath:
            try:
                self.current_monthly_chart_figure.savefig(filepath, dpi=100, bbox_inches='tight',
                                                          facecolor=self.current_monthly_chart_figure.get_facecolor())
                QMessageBox.information(self, "Kaydedildi", f"Aylık grafik başarıyla kaydedildi: {Path(filepath).name}")
                logging.info("Aylık grafik kaydedildi: %s", filepath)
            except Exception as e:
                QMessageBox.critical(self, "Kaydetme Hatası", f"Aylık grafik kaydedilirken bir hata oluştu: {e}")
                logging.exception("Aylık grafik kaydetme hatası.")


class MainWindow(QMainWindow):
    """Ana uygulama penceresini temsil eder."""

    def __init__(self) -> None:
        super().__init__()
        self.excel_path: Path | None = None
        self.selected_sheet: str | None = None
        self.available_sheets: List[str] = []  # Yeni eklendi: uygun sayfaların listesi
        self.df: pd.DataFrame = pd.DataFrame()
        self.grouping_col_name: str | None = None  # For daily graphs and monthly OEE
        self.grouped_col_name: str | None = None  # For daily graphs and monthly OEE (e.g., 'Ürün' column)
        self.metric_cols: List[str] = []  # For daily graphs
        self.oee_col_name: str | None = None  # For daily graphs and monthly OEE
        self.selected_grouping_val: str | None = None  # For daily graphs
        self.grouped_values: List[str] = []  # For daily graphs
        self.selected_metrics: List[str] = []  # For daily graphs

        self.stacked_widget = QStackedWidget()
        self.file_selection_page = FileSelectionPage(self)
        self.data_selection_page = DataSelectionPage(self)  # Daily graphs data selection
        self.daily_graphs_page = DailyGraphsPage(self)  # Daily graphs display
        self.monthly_graphs_page = MonthlyGraphsPage(self)  # New monthly graphs page

        self.stacked_widget.addWidget(self.file_selection_page)
        self.stacked_widget.addWidget(self.data_selection_page)
        self.stacked_widget.addWidget(self.daily_graphs_page)
        self.stacked_widget.addWidget(self.monthly_graphs_page)  # Add new page

        self.setCentralWidget(self.stacked_widget)
        self.setWindowTitle("OEE ve Durum Grafiği Uygulaması")
        self.setGeometry(100, 100, 1200, 800)  # Pencere boyutunu ayarla

        self.apply_stylesheet()
        self.goto_page(0)  # Başlangıçta dosya seçimi sayfasına git

    def apply_stylesheet(self):
        """Uygulamaya modern bir stil uygular."""
        self.setStyleSheet("""
            QWidget {
                background-color: #f0f2f5; /* Açık gri arkaplan */
                font-family: 'Segoe UI', Arial, sans-serif; /* Modern font */
                color: #333333;
            }
            QLabel#title_label {
                font-size: 28pt; /* Slightly larger title */
                font-weight: bold;
                color: #2c3e50; /* Koyu gri/mavi başlık */
                margin-bottom: 25px;
                padding: 10px;
            }
            QLabel {
                font-size: 11pt;
            }
            QPushButton {
                background-color: #3498db; /* Mavi tonu */
                color: white;
                padding: 12px 25px; /* Daha büyük padding */
                border-radius: 8px; /* Daha yumuşak kenarlar */
                border: none;
                font-weight: bold;
                font-size: 11pt;
                margin: 8px; /* Daha fazla dış boşluk */
                box-shadow: 2px 2px 5px rgba(0, 0, 0, 0.2); /* Hafif gölge */
            }
            QPushButton:hover {
                background-color: #2980b9; /* Koyu mavi tonu */
                box-shadow: 3px 3px 8px rgba(0, 0, 0, 0.3); /* Daha belirgin gölge */
            }
            QPushButton:pressed {
                background-color: #21618c; /* Daha da koyu */
            }
            QPushButton:disabled {
                background-color: #cccccc; /* Gri devre dışı buton */
                color: #666666;
                box-shadow: none;
            }
            QComboBox, QListWidget, QScrollArea, QProgressBar, QFrame, QLineEdit {
                border: 1px solid #dcdcdc; /* Daha yumuşak kenarlık */
                border-radius: 6px; /* Daha yumuşak kenarlar */
                padding: 8px; /* Daha fazla iç boşluk */
                background-color: white;
                selection-background-color: #aed6f1; /* Seçim rengi */
                selection-color: black;
            }
            QComboBox::drop-down {
                border: 0px; /* Okun kenarlığını kaldır */
            }
            QComboBox::down-arrow {
                image: url(down_arrow.png); /* Özel ok ikonu kullanılabilir */
                width: 12px;
                height: 12px;
            }
            QListWidget::item {
                padding: 5px; /* Daha fazla item boşluğu */
            }
            QListWidget::item:selected {
                background-color: #3498db;
                color: white;
                border-radius: 3px;
            }
            QCheckBox {
                spacing: 8px; /* Daha fazla boşluk */
                padding: 5px;
                font-size: 10pt;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
                border-radius: 4px;
                border: 1px solid #3498db;
                background-color: white;
            }
            QCheckBox::indicator:checked {
                background-color: #3498db;
                border: 1px solid #2980b9;
                image: url(check_mark.png); /* Özel onay işareti ikonu kullanılabilir */
            }
            QMessageBox {
                background-color: #ffffff;
                color: #333333;
                font-size: 10pt;
            }
            QProgressBar {
                border: 1px solid #b0b0b0;
                border-radius: 7px;
                text-align: center;
                height: 20px;
                margin: 10px 0;
            }
            QProgressBar::chunk {
                background-color: #2ecc71; /* Yeşil ilerleme çubuğu */
                border-radius: 7px;
            }
            /* Scrollbar styling */
            QScrollArea > QWidget > QWidget { /* Targets the content widget within QScrollArea */
                background-color: white;
            }
            QScrollBar:vertical {
                border: 1px solid #999;
                background: #f0f0f0;
                width: 12px;
                margin: 0px 0px 0px 0px;
            }
            QScrollBar::handle:vertical {
                background: #c0c0c0;
                min-height: 20px;
                border-radius: 5px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                border: none;
                background: none;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }
        """)

    def goto_page(self, index: int) -> None:
        """Belirli bir sayfaya geçiş yapar ve sayfayı yeniler."""
        self.stacked_widget.setCurrentIndex(index)
        if index == 1:  # Günlük grafikler için veri seçimi sayfası
            self.data_selection_page.refresh()
        elif index == 2:  # Günlük grafikler görüntüleme sayfası
            self.daily_graphs_page.enter_page()
        elif index == 3:  # Aylık grafikler sayfası
            self.monthly_graphs_page.enter_page()

    def load_excel(self) -> None:
        """Seçilen Excel dosyasını ve sayfasını yükler."""
        if not self.excel_path or not self.selected_sheet:
            logging.warning("load_excel: Excel yolu veya seçili sayfa boş. Veri yüklenemiyor.")
            return

        # Eğer veri zaten yüklenmişse ve aynı dosya/sayfa ise tekrar yükleme
        if not self.df.empty and self.df.attrs.get('excel_path') == self.excel_path and \
                self.df.attrs.get('selected_sheet') == self.selected_sheet:
            logging.info(f"Veri '{self.selected_sheet}' sayfasından zaten yüklü. Tekrar yüklenmiyor.")
            return

        try:
            # Load with header=None and skiprows=[0] to treat first row as data if needed,
            # but for monthly graphs, we need the actual header for column names.
            # Let's assume for SMD-OEE, the first row is the header, so no skiprows for this sheet.
            if self.selected_sheet == "SMD-OEE":
                self.df = pd.read_excel(self.excel_path, sheet_name=self.selected_sheet)
                # Ensure column names are strings to avoid issues with mixed types
                self.df.columns = self.df.columns.astype(str)
            else:
                self.df = pd.read_excel(self.excel_path, sheet_name=self.selected_sheet, header=None, skiprows=[0])
                # Ensure column names are strings for consistency
                self.df.columns = [str(col) for col in self.df.columns]

            # Yüklenen verinin kaynağını sakla
            self.df.attrs['excel_path'] = self.excel_path
            self.df.attrs['selected_sheet'] = self.selected_sheet

            logging.info("'%s' sayfasından veri yüklendi. Satır sayısı: %d", self.selected_sheet, len(self.df))

            # Sütun isimlerini belirle (A, B, BP, H-BD)
            # A sütunu: Gruplama değişkeni (tarih)
            self.grouping_col_name = self.df.columns[excel_col_to_index('A')]
            # B sütunu: Gruplanan değişken (ürün/hat)
            self.grouped_col_name = self.df.columns[excel_col_to_index('B')]
            # BP sütunu: OEE değeri
            self.oee_col_name = self.df.columns[excel_col_to_index('BP')] if excel_col_to_index('BP') < len(
                self.df.columns) else None  # Ensure BP column exists

            # H'den BD'ye kadar olan sütunlar: Metrikler (duruş sebepleri)
            start_col_index = excel_col_to_index('H')
            end_col_index = excel_col_to_index('BD')
            # AP sütununu metriklerden hariç tut
            ap_col_index = excel_col_to_index('AP')

            self.metric_cols = []
            for i in range(start_col_index, end_col_index + 1):
                if i < len(self.df.columns) and i != ap_col_index:
                    self.metric_cols.append(self.df.columns[i])

            logging.info("Tanımlanan gruplama sütunu: %s", self.grouping_col_name)
            logging.info("Tanımlanan gruplanan sütun: %s", self.grouped_col_name)
            logging.info("Tanımlanan OEE sütunu: %s", self.oee_col_name)
            logging.info("Tanımlanan metrik sütunları: %s", self.metric_cols)

        except Exception as e:
            QMessageBox.critical(self, "Veri Yükleme Hatası", f"Veri yüklenirken bir hata oluştu: {e}")
            logging.exception("Excel veri yükleme hatası.")
            self.df = pd.DataFrame()  # Hata durumunda boş DataFrame ayarla


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # Daha modern bir stil kullan

    try:
        win = MainWindow()
        win.show()
        sys.exit(app.exec_())
    except Exception as e:
        logging.exception("Uygulama başlatılırken kritik hata oluştu.")
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText("Uygulama başlatılırken kritik bir hata oluştu.")
        msg.setInformativeText(str(e))
        msg.setWindowTitle("Kritik Hata")
        msg.exec_()
        sys.exit(1)