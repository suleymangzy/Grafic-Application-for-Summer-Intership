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
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'DejaVu Sans', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False  # Eksi işaretinin doğru gösterimi


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
        self.init_ui()
        self.figures_data: List[Tuple[str, Figure]] = []  # Oluşturulan figürleri tutar (ürün adı, figür objesi)
        self.current_page = 0  # Mevcut sayfa numarası

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

        self.lbl_chart_info = QLabel("") # Added for displaying date and product info
        self.lbl_chart_info.setAlignment(Qt.AlignCenter)
        self.lbl_chart_info.setStyleSheet("font-weight: bold; font-size: 12pt;")
        nav_top.addWidget(self.lbl_chart_info)

        nav_top.addStretch(1)
        self.lbl_page = QLabel("Sayfa 0 / 0")
        self.lbl_page.setAlignment(Qt.AlignCenter)
        nav_top.addWidget(self.lbl_page)
        nav_top.addStretch(1)

        self.btn_save_image = QPushButton("Grafiği Kaydet (PNG/JPEG)")
        self.btn_save_image.clicked.connect(self.save_single_graph_as_image)
        self.btn_save_image.setEnabled(False)  # Başlangıçta devre dışı
        nav_top.addWidget(self.btn_save_image)
        main_layout.addLayout(nav_top)

        # Grafiklerin gösterileceği kaydırılabilir alan
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.canvas_holder = QWidget()
        self.vbox_canvases = QVBoxLayout(self.canvas_holder)
        self.canvas_holder.setLayout(self.vbox_canvases)
        self.scroll.setWidget(self.canvas_holder)
        main_layout.addWidget(self.scroll)

        # Alt navigasyon butonları (önceki/sonraki sayfa)
        nav_bottom = QHBoxLayout()
        nav_bottom.addStretch(1)
        self.btn_prev = QPushButton("← Önceki Sayfa")
        self.btn_prev.clicked.connect(self.prev_page)
        self.btn_prev.setEnabled(False)  # Başlangıçta devre dışı
        nav_bottom.addWidget(self.btn_prev)

        self.btn_next = QPushButton("Sonraki Sayfa →")
        self.btn_next.clicked.connect(self.next_page)
        self.btn_next.setEnabled(False)  # Başlangıçta devre dışı
        nav_bottom.addWidget(self.btn_next)
        nav_bottom.addStretch(1)
        main_layout.addLayout(nav_bottom)

    def enter_page(self) -> None:
        """Bu sayfaya geçildiğinde grafik oluşturma işlemini başlatır."""
        self.clear_canvases()  # Önceki grafikleri temizle
        self.figures_data.clear()  # Figür verilerini temizle
        self.current_page = 0  # Sayfayı sıfırla
        self.update_page_label()  # Sayfa etiketini güncelle
        self.progress.setValue(0)  # İlerleme çubuğunu sıfırla
        self.btn_save_image.setEnabled(False)  # Kaydet butonunu devre dışı bırak
        self.lbl_chart_info.setText("") # Clear chart info label

        df = self.main_window.df

        # GraphWorker'ı başlat
        self.worker = GraphWorker(
            df=df,
            grouping_col_name=self.main_window.grouping_col_name,
            grouped_col_name=self.main_window.grouped_col_name,
            grouped_values=self.main_window.grouped_values,
            metric_cols=self.main_window.selected_metrics,
            oee_col_name=self.main_window.oee_col_name,
            selected_grouping_val=self.main_window.selected_grouping_val
        )
        self.worker.progress.connect(self.progress.setValue)  # İlerleme çubuğunu güncelle
        self.worker.finished.connect(self.on_results)  # İşlem bitince sonuçları al
        self.worker.error.connect(lambda m: QMessageBox.critical(self, "Hata", m))  # Hata durumunda mesaj göster
        self.worker.start()  # İş parçacığını başlat

    def on_results(self, results: List[Tuple[str, pd.Series, str]]) -> None:
        """GraphWorker'dan gelen sonuçları işler ve grafikleri oluşturur."""
        self.progress.setValue(100)  # İlerleme çubuğunu tamamla

        if not results:
            QMessageBox.information(self, "Veri yok", "Grafik oluşturulamadı. Seçilen kriterlere göre veri bulunamadı.")
            self.btn_save_image.setEnabled(False)
            return

        # Convert cm to inches for figure size
        cm_to_inches = 0.393701
        fig_width_cm = 14.93
        fig_height_cm = 9.89
        fig_width_inches = fig_width_cm * cm_to_inches
        fig_height_inches = fig_height_cm * cm_to_inches

        for grouped_val, metric_sums, oee_display_value in results:
            fig, ax = plt.subplots(figsize=(fig_width_inches, fig_height_inches), subplot_kw=dict(aspect="equal"))

            # Grafik arka plan rengi
            background_color = '#d3d3d3'  # Açık gri
            fig.patch.set_facecolor(background_color)
            ax.set_facecolor(background_color)

            # Dinamik renk paleti oluştur
            num_metrics = len(metric_sums)  # Directly use the number of effective metrics

            # Eğer sadece bir metrik varsa ve bu "HAT ÇALIŞMADI" ise özel rengi kullan
            if num_metrics == 1 and metric_sums.index[0] == 'HAT ÇALIŞMADI':
                chart_colors = ['#FF9841']  # Resim1.png'deki turuncu/şeftali rengi
            else:
                colors_palette = plt.cm.get_cmap('tab20', num_metrics)
                chart_colors = [colors_palette(i) for i in range(num_metrics)]

            wedges, texts, autotexts = ax.pie(
                metric_sums,
                autopct='',  # Yüzdeleri direk pasta üzerinde gösterme
                startangle=90,
                wedgeprops=dict(width=0.4, edgecolor='w'),  # Donut grafik için
                colors=chart_colors
            )

            # OEE değerini grafik merkezine yerleştir
            ax.text(0, 0, f"OEE\n{oee_display_value}",
                    horizontalalignment='center', verticalalignment='center',
                    fontsize=24, fontweight='bold', color='black')

            # Metrik etiketlerini çizgilerle bağlayarak yerleştirme
            bbox_props = dict(boxstyle="round,pad=0.3", fc="white", ec="0.5", lw=0.5)
            # Changed arrow color to bright gray
            kw = dict(arrowprops=dict(arrowstyle="-", connectionstyle="arc3,rad=0.3", color='dimgray', lw=1.5),
                      bbox=bbox_props, zorder=0, va="center")

            # "HAT ÇALIŞMADI" için özel etiketleme
            if num_metrics == 1 and metric_sums.index[0] == 'HAT ÇALIŞMADI':
                label_text = (
                    f"{metric_sums.index[0]}; "
                    f"{int(metric_sums.values[0] // 3600):02d}:"
                    f"{int((metric_sums.values[0] % 3600) // 60):02d}:00; "  # Saniye de eklendi
                    f"{metric_sums.values[0] / metric_sums.sum() * 100:.0f}%"
                )
                # Resim1.png'deki konum ve stil için fig.text kullan
                # Adjusted position to not conflict with top menu, further down
                fig.text(0.5, 0.2,  # Figür koordinatları (0-1), merkezi alt kısımda
                         label_text,
                         horizontalalignment='center', verticalalignment='center',
                         fontsize=12,  # Resimdeki font boyutuna yakın
                         bbox=dict(boxstyle="square,pad=0.5", fc="white", ec="black", lw=1),
                         # Kare kutu, siyah kenarlık
                         transform=fig.transFigure  # Figür koordinat sistemini kullan
                         )
            else:
                # Other metric labels
                for i, p in enumerate(wedges):
                    ang = (p.theta2 - p.theta1) / 2. + p.theta1
                    y = np.sin(np.deg2rad(ang))
                    x = np.cos(np.deg2rad(ang))

                    # Adjusted these multipliers to move labels further out to reduce overlap
                    outside_x = 1.6 * x if x > 0 else 1.6 * x
                    outside_y = 1.6 * y if y > 0 else 1.6 * y


                    horizontalalignment = "left" if x > 0 else "right"

                    label_text = (
                        f"{metric_sums.index[i]}; "
                        f"{int(metric_sums.values[i] // 3600):02d}:"
                        f"{int((metric_sums.values[i] % 3600) // 60):02d}; "
                        f"{metric_sums.values[i] / metric_sums.sum() * 100:.0f}%"
                    )

                    ax.annotate(label_text, xy=(x, y), xytext=(outside_x, outside_y),
                                horizontalalignment=horizontalalignment, **kw,
                                color='dimgray') # Set label text color to dimgray

            # TOPLAM DURUŞ hesapla ve göster
            total_duration_seconds = metric_sums.sum()
            total_duration_hours = int(total_duration_seconds // 3600)
            total_duration_minutes = int((total_duration_seconds % 3600) // 60)
            # Modified to only show total duration, date and product moved to top menu
            total_duration_text = f"TOPLAM DURUŞ\n{total_duration_hours} SAAT {total_duration_minutes} DAKİKA"

            # TOPLAM DURUŞ metnini sol alt köşeye yerleştir
            fig.text(0.05, 0.05, total_duration_text, transform=fig.transFigure,
                     fontsize=14, fontweight='bold', verticalalignment='bottom')

            ax.set_title("")  # Üstteki başlığı kaldır
            ax.axis("equal")  # Pie chart'ın daire şeklinde görünmesini sağlar
            # Adjusted rect to make more room for labels and bottom text
            fig.tight_layout(rect=[0, 0.1, 1, 0.95])

            self.figures_data.append((grouped_val, fig))
            plt.close(fig)  # Bellek sızıntısını önlemek için figürü kapat

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
            self.lbl_chart_info.setText("") # Clear info when no graphs
            return

        start_index = self.current_page * GRAPHS_PER_PAGE
        end_index = start_index + GRAPHS_PER_PAGE

        graphs_to_display = self.figures_data[start_index:end_index]

        for grouped_val, fig in graphs_to_display:
            canvas = FigureCanvas(fig)
            self.vbox_canvases.addWidget(canvas)
            # Update the label with current product and date
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

        # Mevcut sayfadaki ilk (veya tek) grafiği al
        fig_index_on_page = self.current_page * GRAPHS_PER_PAGE
        if fig_index_on_page < len(self.figures_data):
            grouped_val, fig = self.figures_data[fig_index_on_page]
        else:
            QMessageBox.warning(self, "Kaydedilecek Grafik Yok", "Mevcut sayfada kaydedilecek bir grafik yok.")
            return

        default_filename = f"grafik_{grouped_val}_{self.main_window.selected_grouping_val}.png".replace(" ",
                                                                                                        "_").replace(
            "/", "-")
        filepath, _ = QFileDialog.getSaveFileName(
            self, "Grafiği Kaydet", default_filename, "PNG (*.png);;JPEG (*.jpeg);;JPG (*.jpg)"
        )

        if filepath:
            try:
                fig.savefig(filepath, dpi=300, bbox_inches='tight', facecolor=fig.get_facecolor())
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