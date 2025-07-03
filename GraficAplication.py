from __future__ import annotations

import sys
import logging
import datetime
from pathlib import Path
from typing import List, Tuple, Any

import pandas as pd
import numpy as np  # NaN ve sayısal işlemler için

import matplotlib

matplotlib.use("Agg")  # GUI bağımlı olmayan backend
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.backends.backend_pdf import PdfPages  # PDF'e kaydetmek için

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
    QCheckBox,  # Metrikler için checkbox
)

# ────────────────────────────────────────────────────────────────────────────────
# Sabitler & Logging
# ────────────────────────────────────────────────────────────────────────────────
GRAPHS_PER_PAGE = 4
REQ_SHEETS = {"SMD-OEE", "ROBOT", "DALGA_LEHİM"}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)

# Matplotlib için Türkçe karakter desteği ve stil
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'DejaVu Sans', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False  # Negatif işaretlerini düzgün göster


# ────────────────────────────────────────────────────────────────────────────────
# Yardımcı Fonksiyonlar
# ────────────────────────────────────────────────────────────────────────────────

def excel_col_to_index(col_letter: str) -> int:
    """'AP' -> 41 gibi 0-tabanlı index döndürür."""
    index = 0
    for char in col_letter.upper():
        if not char.isalpha():
            raise ValueError(f"Geçersiz sütun harfi: {col_letter}")
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1  # 0-tabanlı yapmak için -1


# Süre formatındaki verileri güvenli bir şekilde saniyeye çeviren fonksiyon
def seconds_from_timedelta(series: pd.Series) -> pd.Series:
    """
    Pandas Serisindeki çeşitli süre formatlarını (timedelta string, datetime.time, float)
    güvenli bir şekilde toplam saniyeye çevirir.
    """
    seconds_series = pd.Series(0.0, index=series.index, dtype=float)

    # 1. datetime.time objelerini işle
    is_time_obj = series.apply(lambda x: isinstance(x, datetime.time))
    if is_time_obj.any():
        time_objects = series[is_time_obj]
        seconds_series.loc[is_time_obj] = time_objects.apply(
            lambda t: t.hour * 3600 + t.minute * 60 + t.second + t.microsecond / 1e6
        )

    # 2. Geri kalan (henüz işlenmemiş ve NaN olmayan) değerleri stringe çevir ve timedelta olarak dene
    # NaN'ları ve daha önce işlenmiş time objelerini dışarıda bırak
    remaining_indices = series.index[~is_time_obj & series.notna()]
    if not remaining_indices.empty:
        remaining_series_str = series.loc[remaining_indices].astype(str).str.strip()

        # Boş stringleri NaN yap
        remaining_series_str = remaining_series_str.replace('', np.nan)

        converted_td = pd.to_timedelta(remaining_series_str, errors='coerce')

        # Geçerli timedelta dönüşümlerini ata
        valid_td_mask = pd.notna(converted_td)
        seconds_series.loc[remaining_indices[valid_td_mask]] = converted_td[valid_td_mask].dt.total_seconds()

    # 3. Hala NaN olan (string/timedelta dönüşümü başarısız olan) değerleri sayısal olarak dene (örn: Excel'deki gün tabanlı süreler)
    remaining_nan_indices = seconds_series.index[seconds_series.isna()]
    if not remaining_nan_indices.empty:
        numeric_values = pd.to_numeric(series.loc[remaining_nan_indices], errors='coerce')

        # Sadece geçerli sayısal değerleri işle
        valid_numeric_mask = pd.notna(numeric_values)
        if valid_numeric_mask.any():
            converted_from_numeric = pd.to_timedelta(numeric_values[valid_numeric_mask], unit='D', errors='coerce')

            valid_num_td_mask = pd.notna(converted_from_numeric)
            seconds_series.loc[remaining_nan_indices[valid_numeric_mask & valid_num_td_mask]] = converted_from_numeric[
                valid_num_td_mask].dt.total_seconds()

    return seconds_series.fillna(0.0)  # Tüm NaN'ları 0.0 ile doldur


# ────────────────────────────────────────────────────────────────────────────────
# QThread: GraphWorker
# ────────────────────────────────────────────────────────────────────────────────

class GraphWorker(QThread):
    # finished sinyali: List[Tuple[gruplanmış_değişken_değeri_str, metrik_toplamları_PandasSerisi, bp_toplam_saniye_float]]
    finished = pyqtSignal(list)
    progress = pyqtSignal(int)
    error = pyqtSignal(str)

    def __init__(
            self,
            df: pd.DataFrame,
            grouping_col_name: str,  # A sütununun başlığı (Gruplama Değişkeni)
            grouped_col_name: str,  # B sütununun başlığı (Gruplanan Değişken)
            grouped_values: List[str],  # Seçilen B sütunu değerleri (yani grafik başına bir)
            metric_cols: List[str],  # Seçilen metrik sütun başlıkları
            bp_col_name: str | None  # BP sütununun başlığı (varsa)
    ) -> None:
        super().__init__()
        self.df = df.copy()
        self.grouping_col_name = grouping_col_name
        self.grouped_col_name = grouped_col_name
        self.grouped_values = grouped_values
        self.metric_cols = metric_cols
        self.bp_col_name = bp_col_name

    # ------------------------------------------------------------------
    def run(self) -> None:
        try:
            results: List[Tuple[str, pd.Series, float]] = []
            total = len(self.grouped_values)

            # İlgili tüm sütunları (metrikler ve BP) bir kere saniyeye dönüştür
            # Bu, her döngüde aynı dönüşümü tekrarlamayı önler.
            all_cols_to_process = list(set(self.metric_cols + ([self.bp_col_name] if self.bp_col_name else [])))
            df_processed_times = self.df.copy()  # Orijinal DataFrame'i değiştirmemek için kopyala

            for col in all_cols_to_process:
                if col in df_processed_times.columns:
                    df_processed_times[col] = seconds_from_timedelta(df_processed_times[col])

            # Şimdi her bir 'Gruplanan Değişken' değeri için grafik verisini hazırla
            for i, current_grouped_val in enumerate(self.grouped_values, 1):
                # Sadece mevcut gruplanmış değişkenin verilerini içeren alt küme
                subset_df_for_chart = df_processed_times[
                    df_processed_times[self.grouped_col_name].astype(str) == current_grouped_val
                    ].copy()  # Filter by B column value

                # Metriklerin toplamını al
                sums = subset_df_for_chart[self.metric_cols].sum()
                sums = sums[sums > 0]  # Sadece pozitif toplamı olan metrikleri dahil et

                bp_total_seconds = 0.0
                if self.bp_col_name and self.bp_col_name in subset_df_for_chart.columns:
                    # BP sütununun toplam saniyesini al
                    bp_total_seconds = subset_df_for_chart[self.bp_col_name].sum()

                if not sums.empty:  # Eğer çizilecek geçerli metrik toplamı varsa
                    results.append((current_grouped_val, sums, bp_total_seconds))
                self.progress.emit(int(i / total * 100))

            self.finished.emit(results)
        except Exception as exc:
            logging.exception("GraphWorker hata")
            self.error.emit(str(exc))


# ────────────────────────────────────────────────────────────────────────────────
# Sayfalar
# ────────────────────────────────────────────────────────────────────────────────

class FileSelectionPage(QWidget):
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
        self.sheet_selection_label.hide()
        layout.addWidget(self.sheet_selection_label)

        self.cmb_sheet = QComboBox()
        self.cmb_sheet.setEnabled(False)
        self.cmb_sheet.currentIndexChanged.connect(self.on_sheet_selected)
        self.cmb_sheet.hide()
        layout.addWidget(self.cmb_sheet)

        self.btn_next = QPushButton("İleri →")
        self.btn_next.setEnabled(False)
        self.btn_next.clicked.connect(self.go_next)
        layout.addWidget(self.btn_next, alignment=Qt.AlignRight)

        layout.addStretch(1)

    def browse(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Excel seç", str(Path.home()), "Excel Files (*.xlsx)")
        if not path:
            return

        try:
            xls = pd.ExcelFile(path)
            sheets = sorted(list(REQ_SHEETS.intersection(set(xls.sheet_names))))

            if not sheets:
                QMessageBox.warning(self, "Uygun sayfa yok", "Seçilen dosyada istenen sheet bulunamadı.")
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

            # Eğer tek sayfa varsa otomatik seç
            if len(sheets) == 1:
                self.main_window.selected_sheet = sheets[0]
                self.sheet_selection_label.setText(f"İşlenecek Sayfa: <b>{self.main_window.selected_sheet}</b>")
                self.cmb_sheet.hide()  # ComboBox'ı gizle
            else:
                self.main_window.selected_sheet = self.cmb_sheet.currentText()  # Varsayılanı ata

            logging.info("Dosya seçildi: %s", path)

        except Exception as e:
            QMessageBox.critical(self, "Okuma hatası", f"Dosya okunurken bir hata oluştu: {e}")
            self.reset_page()

    def on_sheet_selected(self) -> None:
        self.main_window.selected_sheet = self.cmb_sheet.currentText()
        self.btn_next.setEnabled(bool(self.main_window.selected_sheet))

    def go_next(self) -> None:
        self.main_window.load_excel()
        self.main_window.goto_page(1)

    def reset_page(self):
        self.main_window.excel_path = None
        self.main_window.selected_sheet = None
        self.lbl_path.setText("Henüz dosya seçilmedi")
        self.cmb_sheet.clear()
        self.cmb_sheet.setEnabled(False)
        self.cmb_sheet.hide()
        self.sheet_selection_label.hide()
        self.btn_next.setEnabled(False)


class DataSelectionPage(QWidget):
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

        # Gruplama Değişkeni (A sütunu)
        grouping_group = QHBoxLayout()
        grouping_group.addWidget(QLabel("<b>Gruplama Değişkeni (A Sütunu):</b>"))
        self.cmb_grouping = QComboBox()
        self.cmb_grouping.currentIndexChanged.connect(self.populate_grouped)
        grouping_group.addWidget(self.cmb_grouping)
        main_layout.addLayout(grouping_group)

        # Gruplanan Değişkenler (B sütunu)
        grouped_group = QHBoxLayout()
        grouped_group.addWidget(QLabel("<b>Gruplanan Değişkenler (B Sütunu):</b>"))
        self.lst_grouped = QListWidget()
        self.lst_grouped.setSelectionMode(QListWidget.MultiSelection)
        grouped_group.addWidget(self.lst_grouped)
        main_layout.addLayout(grouped_group)

        # Metrikler (H-BD, AP hariç)
        metrics_group = QVBoxLayout()
        metrics_group.addWidget(QLabel("<b>Metrikler (H-BD, AP hariç):</b>"))
        self.metrics_scroll_area = QScrollArea()
        self.metrics_scroll_area.setWidgetResizable(True)
        self.metrics_content_widget = QWidget()
        self.metrics_layout = QVBoxLayout(self.metrics_content_widget)
        self.metrics_scroll_area.setWidget(self.metrics_content_widget)
        metrics_group.addWidget(self.metrics_scroll_area)
        main_layout.addLayout(metrics_group)

        # Butonlar
        nav_layout = QHBoxLayout()
        self.btn_back = QPushButton("← Geri")
        self.btn_back.clicked.connect(lambda: self.main_window.goto_page(0))
        nav_layout.addWidget(self.btn_back)

        self.btn_next = QPushButton("İleri →")
        self.btn_next.setEnabled(False)  # Başlangıçta pasif
        self.btn_next.clicked.connect(self.go_next)
        nav_layout.addStretch(1)  # Boşluk ekle
        nav_layout.addWidget(self.btn_next)
        main_layout.addLayout(nav_layout)

    def refresh(self) -> None:
        df = self.main_window.df
        if df.empty:
            QMessageBox.critical(self, "Hata", "Veri yüklenemedi. Lütfen dosyayı kontrol edin.")
            self.main_window.goto_page(0)  # Dosya seçim sayfasına geri dön
            return

        # Gruplama değişkeni (A sütunu başlığı)
        self.cmb_grouping.clear()
        if self.main_window.grouping_col_name:
            grouping_vals = sorted(df[self.main_window.grouping_col_name].dropna().astype(str).unique())
            grouping_vals = [s for s in grouping_vals if s.strip()]  # Boş stringleri de filtrele
            self.cmb_grouping.addItems(grouping_vals)
        else:
            QMessageBox.warning(self, "Uyarı", "Gruplama sütunu (A) bulunamadı veya boş.")

        self.populate_metrics_checkboxes()
        self.populate_grouped()  # Gruplama değişkeni seçildiğinde otomatik doldur

    def populate_grouped(self) -> None:
        self.lst_grouped.clear()
        selected_grouping_val = self.cmb_grouping.currentText()
        df = self.main_window.df

        if selected_grouping_val and self.main_window.grouping_col_name and self.main_window.grouped_col_name:
            # Gruplama değişkeniyle filtrele
            filtered_df = df[df[self.main_window.grouping_col_name].astype(str) == selected_grouping_val]

            # Gruplanan değişkenleri (B sütunu) al
            grouped_vals = sorted(filtered_df[self.main_window.grouped_col_name].dropna().astype(str).unique())
            grouped_vals = [s for s in grouped_vals if s.strip()]  # Boş stringleri filtrele

            for gv in grouped_vals:
                item = QListWidgetItem(gv)
                item.setSelected(True)  # Varsayılan olarak seçili gelmeli
                self.lst_grouped.addItem(item)

        self.update_next_button_state()  # Gruplanan değişkenler listesi güncellendiğinde buton durumunu kontrol et

    def populate_metrics_checkboxes(self):
        # Önceki checkbox'ları temizle
        while self.metrics_layout.count():
            item = self.metrics_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        self.main_window.selected_metrics = []  # Seçili metrik listesini sıfırla

        if not self.main_window.metric_cols:
            empty_label = QLabel("Seçilebilir metrik bulunamadı.", parent=self.metrics_content_widget)
            empty_label.setAlignment(Qt.AlignCenter)
            self.metrics_layout.addWidget(empty_label)
            self.btn_next.setEnabled(False)
            return

        for col_name in self.main_window.metric_cols:
            checkbox = QCheckBox(col_name)

            # Sütunun tamamının boş olup olmadığını kontrol et
            is_entirely_empty = self.main_window.df[col_name].dropna().empty

            checkbox.setChecked(not is_entirely_empty)  # Boş değilse varsayılan olarak seçili
            checkbox.setEnabled(not is_entirely_empty)  # Boş ise devre dışı bırak

            if is_entirely_empty:
                checkbox.setText(f"{col_name} (Boş)")
                checkbox.setStyleSheet("color: gray;")
            else:
                # Başlangıçta seçili olanları main_window'a ekle
                self.main_window.selected_metrics.append(col_name)

            checkbox.stateChanged.connect(self.on_metric_checkbox_changed)
            self.metrics_layout.addWidget(checkbox)

        self.update_next_button_state()  # Metrik checkbox'ları güncellendiğinde buton durumunu kontrol et

    def on_metric_checkbox_changed(self, state):
        sender_checkbox = self.sender()
        metric_name = sender_checkbox.text().replace(" (Boş)", "")  # "(Boş)" etiketini kaldır

        if state == Qt.Checked:
            if metric_name not in self.main_window.selected_metrics:
                self.main_window.selected_metrics.append(metric_name)
        else:
            if metric_name in self.main_window.selected_metrics:
                self.main_window.selected_metrics.remove(metric_name)

        self.update_next_button_state()

    def update_next_button_state(self):
        # Gruplanan değişken seçimi ve en az bir metrik seçili ise ileri butonu aktif
        is_grouped_selected = bool(self.lst_grouped.selectedItems())
        is_metric_selected = bool(self.main_window.selected_metrics)
        self.btn_next.setEnabled(is_grouped_selected and is_metric_selected)

    def go_next(self) -> None:
        self.main_window.grouped_values = [i.text() for i in self.lst_grouped.selectedItems()]

        if not self.main_window.grouped_values or not self.main_window.selected_metrics:
            QMessageBox.warning(self, "Seçim Eksik", "Lütfen en az bir gruplanan değişken ve bir metrik seçin.")
            return

        self.main_window.goto_page(2)


class GraphsPage(QWidget):
    def __init__(self, main_window: "MainWindow") -> None:
        super().__init__()
        self.main_window = main_window
        self.worker: GraphWorker | None = None
        self.init_ui()

        self.figures_data: List[Tuple[str, Figure]] = []  # (grouped_val, Figure) listesi
        self.current_page = 0

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

        # Navigasyon butonları (üst kısım)
        nav_top = QHBoxLayout()
        self.btn_back = QPushButton("← Veri Seçimi")
        self.btn_back.clicked.connect(lambda: self.main_window.goto_page(1))
        nav_top.addWidget(self.btn_back)
        nav_top.addStretch(1)  # Boşluk
        self.lbl_page = QLabel("Sayfa 0 / 0")
        self.lbl_page.setAlignment(Qt.AlignCenter)
        nav_top.addWidget(self.lbl_page)
        nav_top.addStretch(1)  # Boşluk
        self.btn_save = QPushButton("Grafikleri Kaydet (PDF)")
        self.btn_save.clicked.connect(self.save_all_graphs_to_pdf)
        nav_top.addWidget(self.btn_save)
        main_layout.addLayout(nav_top)

        # Kaydırılabilir grafik alanı
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.canvas_holder = QWidget()
        self.vbox_canvases = QVBoxLayout(self.canvas_holder)
        self.canvas_holder.setLayout(self.vbox_canvases)
        self.scroll.setWidget(self.canvas_holder)
        main_layout.addWidget(self.scroll)

        # Sayfa navigasyon butonları (alt kısım)
        nav_bottom = QHBoxLayout()
        nav_bottom.addStretch(1)
        self.btn_prev = QPushButton("← Önceki Sayfa")
        self.btn_prev.clicked.connect(self.prev_page)
        nav_bottom.addWidget(self.btn_prev)
        self.btn_next = QPushButton("Sonraki Sayfa →")
        self.btn_next.clicked.connect(self.next_page)
        nav_bottom.addWidget(self.btn_next)
        nav_bottom.addStretch(1)
        main_layout.addLayout(nav_bottom)

    def enter_page(self) -> None:
        self.clear_canvases()
        self.figures_data.clear()
        self.current_page = 0
        self.update_page_label()
        self.progress.setValue(0)

        df = self.main_window.df

        # GraphWorker'a doğru sütun adlarını ve BP sütun adını ilet
        self.worker = GraphWorker(
            df=df,
            grouping_col_name=self.main_window.grouping_col_name,
            grouped_col_name=self.main_window.grouped_col_name,
            grouped_values=self.main_window.grouped_values,
            metric_cols=self.main_window.selected_metrics,
            bp_col_name=self.main_window.bp_col_name  # BP sütun adını ilet
        )
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished.connect(self.on_results)
        self.worker.error.connect(lambda m: QMessageBox.critical(self, "Hata", m))
        self.worker.start()

    def on_results(self, results: List[Tuple[str, pd.Series, float]]) -> None:
        self.progress.setValue(100)
        if not results:
            QMessageBox.information(self, "Veri yok", "Grafik oluşturulamadı. Seçilen kriterlere göre veri bulunamadı.")
            return

        # Tek tip bir renk paleti kullanarak tutarlılık sağlayın
        # Metriklerin sayısı kadar renk alalım.
        colors_palette = plt.cm.get_cmap('tab20', len(self.main_window.selected_metrics))
        metric_colors = {metric: colors_palette(i) for i, metric in enumerate(self.main_window.selected_metrics)}

        for grouped_val, metric_sums, bp_total_seconds in results:
            fig = Figure(figsize=(8, 8), dpi=100)  # A4 sayfa üzerinde yer alacak uygun boyut
            ax = fig.add_subplot(111)

            # Pasta grafiği oluştur
            # labels=metric_sums.index (metrik isimleri), values=metric_sums.values (toplam saniyeler)
            wedges, texts, autotexts = ax.pie(
                metric_sums.values,
                labels=metric_sums.index,
                autopct="%1.1f%%",  # Yüzdeyi göster
                startangle=90,
                counterclock=False,
                colors=[metric_colors[m] for m in metric_sums.index]  # Metrik renklerini ata
            )

            # Başlık oluşturma
            title = f"{self.main_window.grouped_col_name}: {grouped_val}"
            if self.main_window.bp_col_name and bp_total_seconds > 0:
                bp_hours = int(bp_total_seconds // 3600)
                bp_minutes = int((bp_total_seconds % 3600) // 60)
                bp_seconds = int(bp_total_seconds % 60)
                bp_formatted_time = f"{bp_hours:02d}:{bp_minutes:02d}:{bp_seconds:02d}"
                title += f"\n{self.main_window.bp_col_name}: {bp_formatted_time}"
            ax.set_title(title, fontweight="bold", fontsize=12)

            # Etiketlerin stilini ayarla
            for autotext in autotexts:
                autotext.set_color('black')
                autotext.set_fontsize(9)
            for text in texts:
                text.set_fontsize(9)

            ax.axis("equal")  # Oranların eşit olmasını sağlar (daire şeklinde görünüm)
            fig.tight_layout()  # Grafiğin sıkıca yerleşmesini sağlar

            self.figures_data.append((grouped_val, fig))  # Figürü listeye ekle
        self.display_page()

    # ------------------------------------------------------------------
    def display_page(self) -> None:
        self.clear_canvases()
        start = self.current_page * GRAPHS_PER_PAGE
        end = start + GRAPHS_PER_PAGE

        for _, fig in self.figures_data[start:end]:
            canvas = FigureCanvas(fig)
            frame = QFrame()
            frame.setFrameShape(QFrame.StyledPanel)  # Çerçeve ekle
            frame.setLineWidth(1)
            vb = QVBoxLayout(frame)
            vb.addWidget(canvas)
            self.vbox_canvases.addWidget(frame)
        self.vbox_canvases.addStretch(1)  # Sayfadaki grafikleri üste hizala
        self.update_page_label()

    def clear_canvases(self) -> None:
        while self.vbox_canvases.count():
            item = self.vbox_canvases.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
            elif item.layout():  # İç içe layoutları da temizle
                self.clear_layout(item.layout())

    def clear_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
            elif item.layout():
                self.clear_layout(item.layout())

    def update_page_label(self) -> None:
        total_pages = max(1, (len(self.figures_data) + GRAPHS_PER_PAGE - 1) // GRAPHS_PER_PAGE)
        self.lbl_page.setText(f"Sayfa {self.current_page + 1} / {total_pages}")
        self.btn_prev.setEnabled(self.current_page > 0)
        self.btn_next.setEnabled((self.current_page + 1) * GRAPHS_PER_PAGE < len(self.figures_data))

    def next_page(self) -> None:
        if (self.current_page + 1) * GRAPHS_PER_PAGE < len(self.figures_data):
            self.current_page += 1
            self.display_page()

    def prev_page(self) -> None:
        if self.current_page > 0:
            self.current_page -= 1
            self.display_page()

    def save_all_graphs_to_pdf(self) -> None:
        if not self.figures_data:
            QMessageBox.warning(self, "Grafik yok", "Kaydedilecek grafik bulunamadı.")
            return

        file_name, _ = QFileDialog.getSaveFileName(self, "Grafikleri PDF Olarak Kaydet", "", "PDF Dosyaları (*.pdf)")

        if not file_name:
            return

        try:
            with PdfPages(file_name) as pdf:
                for _, fig in self.figures_data:
                    # PDF'e kaydederken figürü sıkıca sığdır
                    pdf.savefig(fig, bbox_inches='tight', pad_inches=0.5)
                    # plt.close(fig) # Figürü kaydettikten sonra kapat, bellek yönetimi için
                    # Ancak display_page() sırasında çizilen figürleri tutmak gerekiyor,
                    # bu yüzden burada kapatmamak, onun yerine clear_canvases'te kapatmak daha iyi.

                # Tüm grafiklerin açıklamasını kapsayacak şekilde genel renk legend'ı
                if self.main_window.selected_metrics:
                    legend_fig = Figure(figsize=(8.27, 11.69))  # A4 boyutunda boş bir sayfa (dikey)
                    legend_ax = legend_fig.add_subplot(111)
                    legend_ax.set_axis_off()  # Eksenleri gizle

                    handles = []
                    labels = []
                    colors_palette = plt.cm.get_cmap('tab20', len(self.main_window.selected_metrics))

                    # Seçili metriklerin her biri için legend öğesi oluştur
                    for i, metric in enumerate(self.main_window.selected_metrics):
                        color = colors_palette(i)
                        handle = plt.Rectangle((0, 0), 1, 1, fc=color, edgecolor='black')
                        handles.append(handle)
                        labels.append(metric)

                    # Legend'ı A4 sayfasının sağ alt köşesine yerleştir
                    # bbox_to_anchor=(0.95, 0.05) ile konumu ayarla (sağ alt köşe)
                    # ncol=2 ile iki sütunlu düzen
                    legend_ax.legend(handles, labels, title="Metrik Legendı",
                                     loc='lower right', bbox_to_anchor=(0.95, 0.05),
                                     fontsize=9, title_fontsize=10, frameon=True, fancybox=True, shadow=True, ncol=2)

                    pdf.savefig(legend_fig, bbox_inches='tight', pad_inches=0.5)
                    plt.close(legend_fig)  # Oluşturulan legend figürünü kapat

            QMessageBox.information(self, "Başarılı", f"Grafikler '{file_name}' konumuna başarıyla kaydedildi.")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Grafikler kaydedilirken bir hata oluştu: {e}")
        finally:
            # Kaydetme işlemi bittikten sonra figürleri ve canvasları temizle
            # Bu, bellek yönetimini optimize eder.
            self.clear_canvases()
            for _, fig in self.figures_data:
                plt.close(fig)  # Kaydedilmeyen figürleri de kapat
            self.figures_data.clear()
            self.update_page_label()


# ────────────────────────────────────────────────────────────────────────────────
# Main Window
# ────────────────────────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Pasta Grafik Rapor Uygulaması")
        self.resize(1200, 900)

        # Uygulama genelinde kullanılacak paylaşılan veriler
        self.excel_path: Path | None = None
        self.selected_sheet: str = ""
        self.df: pd.DataFrame = pd.DataFrame()
        self.grouping_col_name: str | None = None  # A sütununun başlığı
        self.grouped_col_name: str | None = None  # B sütununun başlığı
        self.bp_col_name: str | None = None  # BP sütununun başlığı
        self.metric_cols: List[str] = []  # Tüm geçerli metrik sütun başlıkları
        self.grouped_values: List[str] = []  # Seçilen gruplanan değişken değerleri (B sütunu)
        self.selected_metrics: List[str] = []  # Kullanıcının seçtiği metrik sütun başlıkları

        self.init_ui()

    def init_ui(self):
        self.stacked_widget = QStackedWidget(self)
        self.setCentralWidget(self.stacked_widget)  # QMainWindow'ın merkezi widget'ı yap

        self.file_selection_page = FileSelectionPage(self)
        self.data_selection_page = DataSelectionPage(self)
        self.graphs_page = GraphsPage(self)

        self.stacked_widget.addWidget(self.file_selection_page)
        self.stacked_widget.addWidget(self.data_selection_page)
        self.stacked_widget.addWidget(self.graphs_page)

        self.stacked_widget.setCurrentWidget(self.file_selection_page)

    def goto_page(self, index: int) -> None:
        """Belirtilen indexteki sayfaya geçiş yapar ve sayfayı yeniler."""
        self.stacked_widget.setCurrentIndex(index)
        if index == 1:  # Data Selection Page
            self.data_selection_page.refresh()
        elif index == 2:  # Graphs Page
            self.graphs_page.enter_page()

    def load_excel(self) -> None:
        """Seçilen Excel dosyasını okur ve sütunları işler."""
        if not self.excel_path or not self.selected_sheet:
            QMessageBox.critical(self, "Hata", "Dosya yolu veya sayfa adı belirtilmedi.")
            return

        logging.info("Excel okunuyor: %s | Sheet: %s", self.excel_path, self.selected_sheet)
        try:
            # İlk satırı başlık olarak oku
            df_raw = pd.read_excel(self.excel_path, sheet_name=self.selected_sheet, header=0)

            # Sütun isimlerini temizle ve büyük harfe çevir
            df_raw.columns = df_raw.columns.astype(str).str.strip().str.upper()
            self.df = df_raw  # DataFrame'i ata

            # A ve B sütun adlarını dinamik olarak belirle
            a_idx = excel_col_to_index('A')
            b_idx = excel_col_to_index('B')
            bp_idx = excel_col_to_index('BP')

            if a_idx < len(self.df.columns):
                self.grouping_col_name = self.df.columns[a_idx]
            else:
                QMessageBox.warning(self, "Uyarı", f"Excel'de 'A' ({a_idx + 1}. sütun) sütunu bulunamadı.")
                self.grouping_col_name = None

            if b_idx < len(self.df.columns):
                self.grouped_col_name = self.df.columns[b_idx]
            else:
                QMessageBox.warning(self, "Uyarı", f"Excel'de 'B' ({b_idx + 1}. sütun) sütunu bulunamadı.")
                self.grouped_col_name = None

            # BP sütun adını belirle
            self.bp_col_name = None
            if bp_idx < len(self.df.columns):
                self.bp_col_name = self.df.columns[bp_idx]
                logging.info("BP sütunu: %s", self.bp_col_name)
            else:
                logging.warning(
                    "BP sütunu ('BP' indeksi) Excel dosyasında bulunamadı. Grafik başlıklarında BP değeri gösterilmeyecek.")

            # Metrik sütunlarını belirle (H'den BD'ye kadar, AP hariç)
            h_idx = excel_col_to_index("H")
            bd_idx = excel_col_to_index("BD")
            ap_idx = excel_col_to_index("AP")

            potential_metrics_from_range = []
            if h_idx < len(self.df.columns) and bd_idx < len(self.df.columns) and h_idx <= bd_idx:
                for i in range(h_idx, bd_idx + 1):
                    col_name = self.df.columns[i]
                    if i != ap_idx:  # AP sütununu hariç tut
                        potential_metrics_from_range.append(col_name)
            else:
                QMessageBox.warning(self, "Uyarı",
                                    f"Metrik aralığı (H-BD) geçersiz veya sütunlar bulunamadı. (H:{h_idx + 1}, BD:{bd_idx + 1}, Toplam Sütun:{len(self.df.columns)})")

            # Sadece değeri olan metrik sütunlarını dahil et
            # Sütunun tamamı boş veya sadece boşluk içeren stringlerse dışla
            self.metric_cols = [
                c for c in potential_metrics_from_range
                if c in self.df.columns and not self.df[c].dropna().empty and not self.df[c].astype(str).str.strip().eq(
                    '').all()
            ]

            logging.info("%d geçerli metrik bulundu", len(self.metric_cols))

        except Exception as e:
            QMessageBox.critical(self, "Okuma hatası",
                                 f"Excel dosyası yüklenirken veya işlenirken bir hata oluştu: {e}\nLütfen dosyanın bozuk olmadığından ve formatının doğru olduğundan emin olun.")
            self.df = pd.DataFrame()  # Hata durumunda DataFrame'i sıfırla
            self.excel_path = None
            self.selected_sheet = None


# ────────────────────────────────────────────────────────────────────────────────
# main()
# ────────────────────────────────────────────────────────────────────────────────

def main() -> None:
    app = QApplication(sys.argv)

    # Global stil ayarları (kullanıcı dostu arayüz için)
    app.setStyleSheet("""
        QWidget {
            font-family: 'Segoe UI', Arial, sans-serif;
            font-size: 10pt;
            background-color: #f0f2f5; /* Açık gri arka plan */
            color: #333333;
        }
        QLabel {
            margin-bottom: 5px;
            color: #555555;
        }
        QLabel#title_label { /* ID seçici */
            color: #2c3e50; /* Koyu mavi */
            font-size: 18pt;
            font-weight: bold;
            margin-bottom: 20px;
        }
        QPushButton {
            background-color: #007bff; /* Mavi */
            color: white;
            padding: 10px 20px;
            border-radius: 5px;
            border: none;
            font-weight: bold;
            margin: 5px;
        }
        QPushButton:hover {
            background-color: #0056b3; /* Koyu mavi hover */
        }
        QPushButton:disabled {
            background-color: #cccccc;
            color: #666666;
        }
        QComboBox, QListWidget, QScrollArea, QProgressBar, QFrame {
            border: 1px solid #c0c0c0;
            border-radius: 4px;
            padding: 5px;
            background-color: white;
        }
        QListWidget::item {
            padding: 3px;
        }
        QCheckBox {
            spacing: 5px;
            padding: 3px;
        }
        QMessageBox {
            background-color: #ffffff;
            color: #333333;
        }
        QProgressBar::chunk {
            background-color: #007bff;
            border-radius: 4px;
        }
    """)

    try:
        win = MainWindow()
        win.show()
        sys.exit(app.exec_())
    except Exception as e:
        # Genel uygulama seviyesi hata yakalama
        QMessageBox.critical(None, "Uygulama Hatası", f"Beklenmeyen bir hata oluştu: {e}\nUygulama kapatılıyor.")
        sys.exit(1)


if __name__ == "__main__":
    print(">> GraficApplication – Sürüm 3 – 3 Tem 2025 – page 4 grafik")
    main()