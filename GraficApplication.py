import sys
import logging
import datetime
from pathlib import Path
from typing import List, Tuple, Any

import pandas as pd
import numpy as np

import matplotlib

# matplotlib.use("Agg") # Masaüstü uygulaması için yoruma alındı
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

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
plt.rcParams['axes.unicode_minus'] = False # Negatif işaretler için

# Global matplotlib ayarları
plt.rcParams['axes.grid'] = True
plt.rcParams['grid.alpha'] = 0.7
plt.rcParams['grid.linestyle'] = '--'
plt.rcParams['grid.linewidth'] = 0.5
plt.rcParams['figure.dpi'] = 100 # Ekran çözünürlüğü için
plt.rcParams['savefig.dpi'] = 300 # Kaydedilen resim çözünürlüğü için

# Tick ayarları: Düz çizgiler ve bitiş noktalarında uyumlu noktalar
plt.rcParams['xtick.direction'] = 'out' # tick markların dışa doğru olmasını sağlar
plt.rcParams['ytick.direction'] = 'out'
plt.rcParams['xtick.major.size'] = 7 # Büyük tick uzunluğu
plt.rcParams['xtick.minor.size'] = 4 # Küçük tick uzunluğu
plt.rcParams['ytick.major.size'] = 7
plt.rcParams['ytick.minor.size'] = 4
plt.rcParams['xtick.major.width'] = 1.5 # Büyük tick kalınlığı
plt.rcParams['xtick.minor.width'] = 1 # Küçük tick kalınlığı
plt.rcParams['ytick.major.width'] = 1.5
plt.rcParams['ytick.minor.width'] = 1
plt.rcParams['xtick.top'] = False # Üst tickleri kapat
plt.rcParams['ytick.right'] = False # Sağ tickleri kapat
plt.rcParams['axes.edgecolor'] = 'black' # Eksen çizgisi rengi
plt.rcParams['axes.linewidth'] = 1.5 # Eksen çizgisi kalınlığı


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
    """
    seconds_series = pd.Series(0.0, index=series.index, dtype=float)

    # datetime.time objelerini işleme
    is_time_obj = series.apply(lambda x: isinstance(x, datetime.time))
    if is_time_obj.any():
        time_objects = series[is_time_obj]
        seconds_series.loc[is_time_obj] = time_objects.apply(
            lambda t: t.hour * 3600 + t.minute * 60 + t.second + t.microsecond / 1e6
        )

    # timedelta objelerini veya timedelta'a dönüştürülebilen stringleri işleme
    remaining_indices = series.index[~is_time_obj & series.notna()]
    if not remaining_indices.empty:
        remaining_series_str = series.loc[remaining_indices].astype(str).str.strip()
        remaining_series_str = remaining_series_str.replace('', np.nan)  # Boş stringleri NaN yap
        converted_td = pd.to_timedelta(remaining_series_str, errors='coerce')
        valid_td_mask = pd.notna(converted_td)
        seconds_series.loc[remaining_indices[valid_td_mask]] = converted_td[valid_td_mask].dt.total_seconds()

    # Sayısal değerleri (gün olarak kabul ederek) işleme
    # Önceki adımlarda işlenmemiş ve hala NaN olan değerleri kontrol et
    remaining_nan_indices = seconds_series.index[seconds_series.isna()]
    if not remaining_nan_indices.empty:
        numeric_values = pd.to_numeric(series.loc[remaining_nan_indices], errors='coerce')
        valid_numeric_mask = pd.notna(numeric_values)
        if valid_numeric_mask.any():
            # Excel'den gelen sayılar bazen gün olarak yorumlanabilir (timedelta gibi)
            converted_from_numeric = pd.to_timedelta(numeric_values[valid_numeric_mask], unit='D', errors='coerce')
            valid_num_td_mask = pd.notna(converted_from_numeric)
            seconds_series.loc[remaining_nan_indices[valid_numeric_mask & valid_num_td_mask]] = converted_from_numeric[
                valid_num_td_mask].dt.total_seconds()

    return seconds_series.fillna(0.0)  # Tüm NaN değerleri 0 ile doldur


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
        self.df = df.copy()  # Veri çerçevesinin bir kopyasıyla çalış
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
            df_processed_times = self.df.copy()
            cols_to_process = list(self.metric_cols)

            for col in cols_to_process:
                if col in df_processed_times.columns:
                    df_processed_times[col] = seconds_from_timedelta(df_processed_times[col])

            for i, current_grouped_val in enumerate(self.grouped_values, 1):
                # Mevcut gruplama ve gruplanan değerlere göre alt veri çerçevesini filtrele
                subset_df_for_chart = df_processed_times[
                    (df_processed_times[self.grouping_col_name].astype(str) == self.selected_grouping_val) &
                    (df_processed_times[self.grouped_col_name].astype(str) == current_grouped_val)
                    ].copy()

                # Metrik sütunlarının toplamlarını al
                sums = subset_df_for_chart[self.metric_cols].sum()
                sums = sums[sums > 0]  # Sadece sıfırdan büyük toplamları dikkate al

                oee_display_value = "0%"  # Varsayılan OEE değeri
                if self.oee_col_name and self.oee_col_name in self.df.columns:
                    matching_rows = self.df[
                        (self.df[self.grouping_col_name].astype(str) == self.selected_grouping_val) &
                        (self.df[self.grouped_col_name].astype(str) == current_grouped_val)
                        ]
                    if not matching_rows.empty:
                        oee_value_raw = matching_rows[self.oee_col_name].iloc[0]
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

        self.sheet_selection_label = QLabel("İşlenecek Sayfa:")
        self.sheet_selection_label.setAlignment(Qt.AlignCenter)
        self.sheet_selection_label.hide()  # Başlangıçta gizli
        layout.addWidget(self.sheet_selection_label)

        self.cmb_sheet = QComboBox()
        self.cmb_sheet.setEnabled(False)  # Başlangıçta devre dışı
        self.cmb_sheet.currentIndexChanged.connect(self.on_sheet_selected)
        self.cmb_sheet.hide()  # Başlangıçta gizli
        layout.addWidget(self.cmb_sheet)

        self.btn_next = QPushButton("İleri →")
        self.btn_next.setEnabled(False)  # Başlangıçta devre dışı
        self.btn_next.clicked.connect(self.go_next)
        layout.addWidget(self.btn_next, alignment=Qt.AlignRight)

        layout.addStretch(1)  # Boşluk ekle

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
            self.cmb_sheet.clear()
            self.cmb_sheet.addItems(sheets)
            self.cmb_sheet.setEnabled(True)
            self.sheet_selection_label.show()
            self.cmb_sheet.show()
            self.btn_next.setEnabled(True)

            if len(sheets) == 1:  # Eğer sadece bir uygun sayfa varsa, otomatik seç
                self.main_window.selected_sheet = sheets[0]
                self.sheet_selection_label.setText(f"İşlenecek Sayfa: <b>{self.main_window.selected_sheet}</b>")
                self.cmb_sheet.hide()  # ComboBox'ı gizle
            else:
                self.main_window.selected_sheet = self.cmb_sheet.currentText()  # Seçili sayfayı al

            logging.info("Dosya seçildi: %s", path)

        except Exception as e:
            QMessageBox.critical(self, "Okuma hatası",
                                 f"Dosya okunurken bir hata oluştu: {e}\nLütfen dosyanın bozuk olmadığından ve Excel formatında olduğundan emin olun.")
            self.reset_page()

    def on_sheet_selected(self) -> None:
        """Sayfa seçimi değiştiğinde ana penceredeki seçimi günceller."""
        self.main_window.selected_sheet = self.cmb_sheet.currentText()
        self.btn_next.setEnabled(bool(self.main_window.selected_sheet))

    def go_next(self) -> None:
        """Bir sonraki sayfaya geçer."""
        self.main_window.load_excel()  # Excel verilerini yükle
        self.main_window.goto_page(1)  # Veri seçimi sayfasına git

    def reset_page(self):
        """Sayfayı başlangıç durumuna döndürür."""
        self.main_window.excel_path = None
        self.main_window.selected_sheet = None
        self.lbl_path.setText("Henüz dosya seçilmedi")
        self.cmb_sheet.clear()
        self.cmb_sheet.setEnabled(False)
        self.cmb_sheet.hide()
        self.sheet_selection_label.hide()
        self.btn_next.setEnabled(False)


class DataSelectionPage(QWidget):
    """Veri seçimi sayfasını temsil eder (gruplama, metrikler vb.)."""

    def __init__(self, main_window: "MainWindow") -> None:
        super().__init__()
        self.main_window = main_window
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setAlignment(Qt.AlignTop)

        title_label = QLabel("<h2>Veri Seçimi</h2>")
        title_label.setObjectName("title_label")
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

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
        df = self.main_window.df
        if df.empty:
            QMessageBox.critical(self, "Hata", "Veri yüklenemedi. Lütfen dosyayı kontrol edin.")
            self.main_window.goto_page(0)  # Dosya seçimine geri dön
            return

        # Gruplama sütunu doldur
        self.cmb_grouping.clear()
        if self.main_window.grouping_col_name and self.main_window.grouping_col_name in df.columns:
            grouping_vals = sorted(df[self.main_window.grouping_col_name].dropna().astype(str).unique())
            grouping_vals = [s for s in grouping_vals if s.strip()]  # Boş stringleri filtrele
            self.cmb_grouping.addItems(grouping_vals)
            if not grouping_vals:
                QMessageBox.warning(self, "Uyarı", "Gruplama sütunu (A) boş veya geçerli değer içermiyor.")
        else:
            QMessageBox.warning(self, "Uyarı", "Gruplama sütunu (A) bulunamadı veya boş.")
            # Eğer gruplama sütunu yoksa, diğer alanları da boşalt
            self.cmb_grouping.clear()
            self.lst_grouped.clear()
            self.clear_metrics_checkboxes()  # Metrik checkbox'larını da temizle
            return  # Fonksiyondan çık

        self.populate_metrics_checkboxes()  # Metrik checkbox'larını doldur
        self.populate_grouped()  # Gruplanan değişkenleri doldur

    def populate_grouped(self) -> None:
        """Gruplanan değişkenler listesini (ürünler) doldurur."""
        self.lst_grouped.clear()
        selected_grouping_val = self.cmb_grouping.currentText()
        df = self.main_window.df

        if selected_grouping_val and self.main_window.grouping_col_name and self.main_window.grouped_col_name:
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
        self.clear_metrics_checkboxes()  # Mevcut checkbox'ları temizle

        self.main_window.selected_metrics = []  # Seçili metrikleri sıfırla

        if not self.main_window.metric_cols:
            empty_label = QLabel("Seçilebilir metrik bulunamadı.", parent=self.metrics_content_widget)
            empty_label.setAlignment(Qt.AlignCenter)
            self.metrics_layout.addWidget(empty_label)
            self.btn_next.setEnabled(False)
            return

        for col_name in self.main_window.metric_cols:
            checkbox = QCheckBox(col_name)
            # Sütunun tamamen boş olup olmadığını kontrol et
            is_entirely_empty = self.main_window.df[col_name].dropna().empty

            if is_entirely_empty:
                checkbox.setChecked(False)  # Boşsa seçili olmasın
                checkbox.setEnabled(False)  # Ve devre dışı olsun
                checkbox.setText(f"{col_name} (Boş)")
                checkbox.setStyleSheet("color: gray;")
            else:
                checkbox.setChecked(True)  # Doluysa seçili olsun
                self.main_window.selected_metrics.append(col_name)  # Seçili metriklere ekle

            checkbox.stateChanged.connect(self.on_metric_checkbox_changed)
            self.metrics_layout.addWidget(checkbox)

        self.update_next_button_state()  # İleri butonunun durumunu güncelle

    def clear_metrics_checkboxes(self):
        """Metrik checkbox'larını temizler."""
        while self.metrics_layout.count():
            item = self.metrics_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def on_metric_checkbox_changed(self, state):
        """Bir metrik checkbox'ının durumu değiştiğinde çağrılır."""
        sender_checkbox = self.sender()
        metric_name = sender_checkbox.text().replace(" (Boş)", "")  # "(Boş)" kısmını temizle

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
        self.main_window.goto_page(2)  # Grafik sayfasına git


class GraphsPage(QWidget):
    """Oluşturulan grafikleri gösteren ve kaydetme seçenekleri sunan sayfa."""

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

        title_label = QLabel("<h2>Grafikler</h2>")
        title_label.setObjectName("title_label")
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        self.progress = QProgressBar()
        self.progress.setAlignment(Qt.AlignCenter)
        self.progress.setTextVisible(True)
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
        self.cmb_graph_type.setCurrentText(self.current_graph_type) # Varsayılanı ayarla
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
        # Sadece mevcut sayfanın grafiğini yeniden çizmek için mevcut veriyi kullanabiliriz.
        # Ancak, mevcut figures_data'daki tüm figürleri yeniden oluşturmak daha sağlam olur.
        # Bu yüzden enter_page'i yeniden çağırarak tüm grafiklerin yeni tipe göre oluşturulmasını sağlarız.
        self.enter_page()


    def enter_page(self) -> None:
        """Bu sayfaya geçildiğinde grafik oluşturma işlemini başlatır."""
        self.clear_canvases()
        self.figures_data.clear()
        self.current_page = 0
        self.update_page_label()
        self.progress.setValue(0)
        self.btn_save_image.setEnabled(False)
        self.lbl_chart_info.setText("")

        df = self.main_window.df

        self.worker = GraphWorker(
            df=df,
            grouping_col_name=self.main_window.grouping_col_name,
            grouped_col_name=self.main_window.grouped_col_name,
            grouped_values=self.main_window.grouped_values,
            metric_cols=self.main_window.selected_metrics,
            oee_col_name=self.main_window.oee_col_name,
            selected_grouping_val=self.main_window.selected_grouping_val
        )
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished.connect(self.on_results)
        self.worker.error.connect(lambda m: QMessageBox.critical(self, "Hata", m))
        self.worker.start()

    def on_results(self, results: List[Tuple[str, pd.Series, str]]) -> None:
        """GraphWorker'dan gelen sonuçları işler ve grafikleri oluşturur."""
        self.progress.setValue(100)

        if not results:
            QMessageBox.information(self, "Veri yok", "Grafik oluşturulamadı. Seçilen kriterlere göre veri bulunamadı.")
            self.btn_save_image.setEnabled(False)
            return

        fig_width_inches = 700 / 100
        fig_height_inches = 460 / 100

        for grouped_val, metric_sums, oee_display_value in results:
            fig, ax = plt.subplots(figsize=(fig_width_inches, fig_height_inches)) # subplot_kw aspect="equal" kaldırıldı
            background_color = 'white'
            fig.patch.set_facecolor(background_color)
            ax.set_facecolor(background_color)

            if not metric_sums.empty:
                sorted_metrics_series = metric_sums.sort_values(ascending=False)
            else:
                sorted_metrics_series = pd.Series()

            num_metrics = len(sorted_metrics_series)
            if num_metrics == 1 and sorted_metrics_series.index[0] == 'HAT ÇALIŞMADI':
                chart_colors = ['#FF9841']
            else:
                colors_palette = plt.cm.get_cmap('tab20', num_metrics)
                chart_colors = [colors_palette(i) for i in range(num_metrics)]

            if self.current_graph_type == "Donut":
                wedges, texts = ax.pie(
                    sorted_metrics_series,
                    autopct=None,
                    startangle=90,
                    wedgeprops=dict(width=0.4, edgecolor='w'),
                    colors=chart_colors
                )

                # OEE değerini grafik merkezine yerleştir
                ax.text(0, 0, f"OEE\n{oee_display_value}",
                        horizontalalignment='center', verticalalignment='center',
                        fontsize=24, fontweight='bold', color='black')

                # Her dilimin üzerine sıra numarasını yaz
                radius_text = 0.7  # Metinlerin dilimlerin ortasına ne kadar yakın olacağı
                for i, wedge in enumerate(wedges):
                    angle = (wedge.theta2 - wedge.theta1) / 2. + wedge.theta1
                    x = radius_text * np.cos(np.deg2rad(angle))
                    y = radius_text * np.sin(np.deg2rad(angle))

                    r, g, b, _ = matplotlib.colors.to_rgba(chart_colors[i])
                    luminance = (0.299 * r + 0.587 * g + 0.114 * b)
                    text_color = 'white' if luminance < 0.5 else 'black'

                    ax.text(x, y, str(i + 1),
                            horizontalalignment='center',
                            verticalalignment='center',
                            fontsize=12,
                            color=text_color,
                            fontweight='bold')

                # Metrik etiketlerini grafiğin solunda alt alta yerleştirme ve numaralandırma
                label_y_start = 0.9  # Adjusted starting position for labels (figure coordinates)
                label_line_height = 0.05  # Approximate line height for each label

                for i, (metric_name, metric_value) in enumerate(sorted_metrics_series.items()):
                    label_text = (
                        f"{i+1}. {metric_name}; "
                        f"{int(metric_value // 3600):02d}:"
                        f"{int((metric_value % 3600) // 60):02d}; "
                        f"{metric_value / sorted_metrics_series.sum() * 100:.0f}%"
                    )
                    y_pos = label_y_start - (i * label_line_height)
                    bbox_props = dict(boxstyle="round,pad=0.3", fc=chart_colors[i], ec=chart_colors[i], lw=0.5)
                    r, g, b, _ = matplotlib.colors.to_rgba(chart_colors[i])
                    luminance = (0.299 * r + 0.587 * g + 0.114 * b)
                    text_color = 'white' if luminance < 0.5 else 'black'

                    fig.text(0.02,
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

            elif self.current_graph_type == "Bar":
                # Bar grafiği çizimi
                metrics = sorted_metrics_series.index.tolist()
                values = sorted_metrics_series.values.tolist()

                y_pos = np.arange(len(metrics))

                ax.barh(y_pos, values, color=chart_colors) # Yatay bar grafiği
                ax.set_yticks(y_pos)
                ax.set_yticklabels([f"{i+1}. {m}" for i, m in enumerate(metrics)], fontsize=10) # Numaralandırma ve etiketler
                ax.invert_yaxis()  # En büyük değeri en üste getir

                ax.set_xlabel("Duruş Süresi (Saniye)")
                ax.set_title(f"Metrik Duruş Süreleri\nOEE: {oee_display_value}", fontsize=16, fontweight='bold') # OEE üstte

                # Her barın üzerine değeri ve yüzdesini yaz
                total_sum = sorted_metrics_series.sum()
                for i, (value, metric_name) in enumerate(zip(values, metrics)):
                    percentage = (value / total_sum) * 100 if total_sum > 0 else 0
                    duration_hours = int(value // 3600)
                    duration_minutes = int((value % 3600) // 60)
                    # Sütunun sonuna hizala, metnin rengini ve boyutunu ayarla
                    text_label = f"{duration_hours:02d}:{duration_minutes:02d} ({percentage:.0f}%)"
                    ax.text(value, i, text_label,
                            va='center', ha='left',  # value'ya göre sağa hizala
                            fontsize=9, color='black',
                            bbox=dict(facecolor='white', alpha=0.7, edgecolor='none', boxstyle="round,pad=0.2")) # Beyaz kutu arkası

                ax.set_xlim(left=0) # X ekseninin 0'dan başlamasını sağla
                fig.tight_layout(rect=[0.1, 0.1, 0.95, 0.9]) # Etiketler için biraz boşluk bırak

            # TOPLAM DURUŞ hesapla ve göster (her iki grafik tipi için)
            total_duration_seconds = sorted_metrics_series.sum()
            total_duration_hours = int(total_duration_seconds // 3600)
            total_duration_minutes = int((total_duration_seconds % 3600) // 60)
            total_duration_text = f"TOPLAM DURUŞ\n{total_duration_hours} SAAT {total_duration_minutes} DAKİKA"

            # TOPLAM DURUŞ metnini sol alt köşeye yerleştir
            fig.text(0.05, 0.05, total_duration_text, transform=fig.transFigure,
                     fontsize=14, fontweight='bold', verticalalignment='bottom')


            self.figures_data.append((grouped_val, fig, oee_display_value)) # OEE değeri de figures_data'da tutulacak
            plt.close(fig)

        self.display_current_page_graphs()
        self.btn_save_image.setEnabled(True)

    def clear_canvases(self) -> None:
        """Mevcut grafik tuvallerini temizler."""
        while self.vbox_canvases.count():
            item = self.vbox_canvases.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
            elif item.layout():
                self._clear_layout(item.layout())

    def _clear_layout(self, layout):
        """Alt layoutları ve widget'ları temizler."""
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
            elif item.layout():
                self._clear_layout(item.layout())

    def display_current_page_graphs(self) -> None:
        """Mevcut sayfadaki grafikleri gösterir."""
        self.clear_canvases()

        if not self.figures_data:
            self.btn_prev.setEnabled(False)
            self.btn_next.setEnabled(False)
            self.update_page_label()
            self.lbl_chart_info.setText("")
            return

        start_index = self.current_page * GRAPHS_PER_PAGE
        end_index = start_index + GRAPHS_PER_PAGE

        graphs_to_display = self.figures_data[start_index:end_index]

        for grouped_val, fig, oee_display_value in graphs_to_display:
            canvas = FigureCanvas(fig)
            canvas.setFixedSize(700, 460)
            self.vbox_canvases.addWidget(canvas)
            self.lbl_chart_info.setText(f"{self.main_window.selected_grouping_val} - {grouped_val}")

        self.update_page_label()
        self.update_navigation_buttons()

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
        """Bir önceki sayfaya geçer."""
        if self.current_page > 0:
            self.current_page -= 1
            self.display_current_page_graphs()

    def next_page(self) -> None:
        """Bir sonraki sayfaya geçer."""
        total_pages = (len(self.figures_data) + GRAPHS_PER_PAGE - 1) // GRAPHS_PER_PAGE if self.figures_data else 0
        if self.current_page < total_pages - 1:
            self.current_page += 1
            self.display_current_page_graphs()

    def save_single_graph_as_image(self) -> None:
        """Mevcut sayfadaki grafiği PNG/JPEG olarak kaydeder."""
        if not self.figures_data:
            QMessageBox.warning(self, "Kaydedilecek Grafik Yok", "Görüntülenecek bir grafik bulunmamaktadır.")
            return

        if self.current_page >= len(self.figures_data) // GRAPHS_PER_PAGE:
            QMessageBox.warning(self, "Geçersiz Sayfa", "Mevcut sayfada kaydedilecek bir grafik yok.")
            return

        fig_index_on_page = self.current_page * GRAPHS_PER_PAGE
        if fig_index_on_page < len(self.figures_data):
            grouped_val, fig, _ = self.figures_data[fig_index_on_page]
        else:
            QMessageBox.warning(self, "Kaydedilecek Grafik Yok", "Mevcut sayfada kaydedilecek bir grafik yok.")
            return

        default_filename = f"grafik_{grouped_val}_{self.main_window.selected_grouping_val}_{self.current_graph_type}.png".replace(" ", "_").replace("/", "-")
        filepath, _ = QFileDialog.getSaveFileName(
            self, "Grafiği Kaydet", default_filename, "PNG (*.png);;JPEG (*.jpeg);;JPG (*.jpg)"
        )

        if filepath:
            try:
                fig.savefig(filepath, dpi=100, bbox_inches='tight', facecolor=fig.get_facecolor())
                QMessageBox.information(self, "Kaydedildi", f"Grafik başarıyla kaydedildi: {Path(filepath).name}")
                logging.info("Grafik kaydedildi: %s", filepath)
            except Exception as e:
                QMessageBox.critical(self, "Kaydetme Hatası", f"Grafik kaydedilirken bir hata oluştu: {e}")
                logging.exception("Grafik kaydetme hatası.")


class MainWindow(QMainWindow):
    """Ana uygulama penceresini temsil eder."""

    def __init__(self) -> None:
        super().__init__()
        self.excel_path: Path | None = None
        self.selected_sheet: str | None = None
        self.df: pd.DataFrame = pd.DataFrame()
        self.grouping_col_name: str | None = None
        self.grouped_col_name: str | None = None
        self.metric_cols: List[str] = []
        self.oee_col_name: str | None = None
        self.selected_grouping_val: str | None = None
        self.grouped_values: List[str] = []
        self.selected_metrics: List[str] = []

        self.stacked_widget = QStackedWidget()
        self.file_selection_page = FileSelectionPage(self)
        self.data_selection_page = DataSelectionPage(self)
        self.graphs_page = GraphsPage(self)

        self.stacked_widget.addWidget(self.file_selection_page)
        self.stacked_widget.addWidget(self.data_selection_page)
        self.stacked_widget.addWidget(self.graphs_page)

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
                font-family: Arial, sans-serif;
                color: #333333;
            }
            QLabel#title_label {
                font-size: 24pt;
                font-weight: bold;
                color: #007bff; /* Mavi başlık */
                margin-bottom: 20px;
            }
            QLabel {
                font-size: 11pt;
            }
            QPushButton {
                background-color: #007bff; /* Mavi buton */
                color: white;
                padding: 10px 20px;
                border-radius: 5px;
                border: none;
                font-weight: bold;
                margin: 5px;
            }
            QPushButton:hover {
                background-color: #0056b3; /* Mavi buton hover */
            }
            QPushButton:disabled {
                background-color: #cccccc; /* Gri devre dışı buton */
                color: #666666;
            }
            QComboBox, QListWidget, QScrollArea, QProgressBar, QFrame {
                border: 1px solid #c0c0c0; /* Açık gri kenarlık */
                border-radius: 4px;
                padding: 5px;
                background-color: white; /* Beyaz arkaplan */
            }
            QListWidget::item {
                padding: 3px;
            }
            QCheckBox {
                spacing: 5px;
                padding: 3px;
            }
            QMessageBox {
                background-color: #ffffff; /* Beyaz mesaj kutusu */
                color: #333333;
            }
            QProgressBar::chunk {
                background-color: #007bff; /* Mavi ilerleme çubuğu dolgusu */
                border-radius: 4px;
            }
        """)

    def goto_page(self, index: int) -> None:
        """Belirli bir sayfaya geçiş yapar ve sayfayı yeniler."""
        self.stacked_widget.setCurrentIndex(index)
        if index == 1:  # Veri seçimi sayfası
            self.data_selection_page.refresh()
        elif index == 2:  # Grafik sayfası
            self.graphs_page.enter_page()

    def load_excel(self) -> None:
        """Seçilen Excel dosyasını ve sayfasını yükler."""
        if not self.excel_path or not self.selected_sheet:
            QMessageBox.warning(self, "Dosya Seçilmedi", "Lütfen önce bir Excel dosyası ve sayfa seçin.")
            return

        try:
            self.df = pd.read_excel(self.excel_path, sheet_name=self.selected_sheet)
            logging.info("'%s' sayfasından veri yüklendi. Satır sayısı: %d", self.selected_sheet, len(self.df))

            # Sütun isimlerini belirle (A, B, BP, H-BD)
            # A sütunu: Gruplama değişkeni (tarih)
            self.grouping_col_name = self.df.columns[excel_col_to_index('A')]
            # B sütunu: Gruplanan değişken (ürün/hat)
            self.grouped_col_name = self.df.columns[excel_col_to_index('B')]
            # BP sütunu: OEE değeri
            self.oee_col_name = self.df.columns[excel_col_to_index('BP')]
            # H'den BD'ye kadar olan sütunlar: Metrikler (duruş sebepleri)
            start_col_index = excel_col_to_index('H')
            end_col_index = excel_col_to_index('BD')
            # AP sütununu metriklerden hariç tut
            ap_col_name = self.df.columns[excel_col_to_index('AP')] if excel_col_to_index('AP') < len(
                self.df.columns) else None

            self.metric_cols = []
            for i in range(start_col_index, end_col_index + 1):
                if i < len(self.df.columns):  # Sütun indeksi dataframe'de mevcutsa
                    col_name = self.df.columns[i]
                    if col_name != ap_col_name:
                        self.metric_cols.append(col_name)

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