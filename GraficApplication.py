import sys
import logging
import datetime
from pathlib import Path
import re
from typing import List, Tuple, Any, Union, Dict

import pandas as pd
import numpy as np

import matplotlib
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
    QSpacerItem,
    QSizePolicy,
    QLineEdit,
)
from PyQt5 import QtGui # Added for QDoubleValidator

# Her sayfada kaç grafik gösterileceği
GRAPHS_PER_PAGE = 1
# Gerekli Excel sayfa isimleri
REQ_SHEETS = {"SMD-OEE", "ROBOT", "DALGA_LEHİM"}

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
plt.rcParams['axes.grid'] = True
plt.rcParams['grid.alpha'] = 0.7
plt.rcParams['grid.linestyle'] = '--'
plt.rcParams['grid.linewidth'] = 0.5
plt.rcParams['figure.dpi'] = 100
plt.rcParams['savefig.dpi'] = 300

# Tick ayarları
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
    """Excel sütun harfini sıfır tabanlı indekse dönüştürür."""
    index = 0
    for char in col_letter.upper():
        if not char.isalpha():
            raise ValueError(f"Geçersiz sütun harfi: {col_letter}")
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1


def seconds_from_timedelta(series: pd.Series) -> pd.Series:
    """Pandas Serisi'ndeki zaman değerlerini saniyeye dönüştürür."""
    seconds_series = pd.Series(0.0, index=series.index, dtype=float)

    is_time_obj = series.apply(lambda x: isinstance(x, datetime.time))
    if is_time_obj.any():
        time_objects = series[is_time_obj]
        seconds_series.loc[is_time_obj] = time_objects.apply(
            lambda t: t.hour * 3600 + t.minute * 60 + t.second + t.microsecond / 1e6
        )

    str_and_timedelta_mask = ~is_time_obj & series.notna()
    if str_and_timedelta_mask.any():
        # Corrected: Use .str.strip() for Series of strings
        converted_td = pd.to_timedelta(series.loc[str_and_timedelta_mask].astype(str).str.strip(), errors='coerce')
        valid_td_mask = pd.notna(converted_td)
        seconds_series.loc[str_and_timedelta_mask & valid_td_mask] = converted_td[valid_td_mask].dt.total_seconds()

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
    """Günlük grafik oluşturma işlemlerini arka planda yürüten işçi sınıfı."""
    finished = pyqtSignal(list)
    progress = pyqtSignal(int)
    error = pyqtSignal(str)

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
            results: List[Tuple[str, pd.Series, str]] = []
            total = len(self.grouped_values)

            for col in self.metric_cols:
                if col in self.df.columns:
                    self.df[col] = seconds_from_timedelta(self.df[col])

            if self.grouping_col_name in self.df.columns:
                self.df[self.grouping_col_name] = self.df[self.grouping_col_name].astype(str)
            if self.grouped_col_name in self.df.columns:
                self.df[self.grouped_col_name] = self.df[self.grouped_col_name].astype(str)

            for i, current_grouped_val in enumerate(self.grouped_values, 1):
                subset_df_for_chart = self.df[
                    (self.df[self.grouping_col_name] == self.selected_grouping_val) &
                    (self.df[self.grouped_col_name] == current_grouped_val)
                    ]

                sums = subset_df_for_chart[
                    [col for col in self.metric_cols if col in subset_df_for_chart.columns]].sum()
                sums = sums[sums > 0]

                oee_display_value = "0%"
                if self.oee_col_name and self.oee_col_name in subset_df_for_chart.columns and not subset_df_for_chart.empty:
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
                            elif oee_value_float > 1.0:
                                oee_display_value = f"{oee_value_float:.0f}%"
                            else:
                                oee_display_value = "0%"
                        except (ValueError, TypeError):
                            logging.warning(
                                f"OEE değeri dönüştürülemedi: {oee_value_raw}. Varsayılan '0%' kullanılacak.")
                            oee_display_value = "0%"

                if not sums.empty:
                    results.append((current_grouped_val, sums, oee_display_value))
                self.progress.emit(int(i / total * 100))

            self.finished.emit(results)
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
        wedges, texts = ax.pie(
            sorted_metrics_series,
            autopct=None,
            startangle=90,
            wedgeprops=dict(width=0.4, edgecolor='w'),
            colors=chart_colors[:len(sorted_metrics_series)]
        )

        ax.text(0, 0, f"OEE\n{oee_display_value}",
                horizontalalignment='center', verticalalignment='center',
                fontsize=24, fontweight='bold', color='black')

        label_y_start = 0.25 + (30 / (fig.get_size_inches()[1] * fig.dpi))
        label_line_height = 0.05

        top_3_metrics = sorted_metrics_series.head(3)
        top_3_colors = chart_colors[:len(top_3_metrics)]

        for i, (metric_name, metric_value) in enumerate(top_3_metrics.items()):
            duration_hours = int(metric_value // 3600)
            duration_minutes = int((metric_value % 3600) // 60)
            duration_seconds = int(metric_value % 60) # Calculate seconds for donut chart
            label_text = (
                f"{i + 1}. {metric_name}; "
                f"{duration_hours:02d}:"
                f"{duration_minutes:02d}:"
                f"{duration_seconds:02d}; " # Updated to HH:MM:SS
                f"{metric_value / sorted_metrics_series.sum() * 100:.0f}%"
            )
            y_pos = label_y_start - (i * label_line_height)
            bbox_props = dict(boxstyle="round,pad=0.3", fc=top_3_colors[i], ec=top_3_colors[i], lw=0.5)
            r, g, b, _ = matplotlib.colors.to_rgba(top_3_colors[i])
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

        ax.set_title("")
        ax.axis("equal")
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

        values_minutes = [v / 60 for v in values]

        y_pos = np.arange(len(metrics))

        ax.barh(y_pos, values_minutes, color=chart_colors)
        ax.set_yticks(y_pos)
        ax.set_yticklabels(metrics, fontsize=10)
        ax.invert_yaxis()

        ax.set_xlabel("")
        ax.set_xticks([])

        ax.set_title(f"OEE: {oee_display_value}", fontsize=16, fontweight='bold')

        ax.grid(False)

        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        ax.spines['left'].set_visible(True)
        ax.spines['bottom'].set_visible(True)

        total_sum = sorted_metrics_series.sum()
        for i, (value, metric_name) in enumerate(zip(values, metrics)):
            percentage = (value / total_sum) * 100 if total_sum > 0 else 0
            duration_hours = int(value // 3600)
            duration_minutes = int((value % 3600) // 60)
            duration_seconds = int(value % 60) # Calculate seconds
            text_label = f"{duration_hours:02d}:{duration_minutes:02d}:{duration_seconds:02d} ({percentage:.0f}%)" # Update format

            text_x_position = (value / 60) + 0.5
            ax.text(text_x_position, i, text_label,
                    va='center', ha='left',
                    fontsize=11, fontweight='bold',
                    color='black')

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

        layout.addStretch(1)

        h_layout_buttons = QHBoxLayout()
        self.btn_daily_graphs = QPushButton("Günlük Grafikler")
        self.btn_daily_graphs.clicked.connect(self.go_to_daily_graphs)
        self.btn_daily_graphs.setEnabled(False)
        h_layout_buttons.addWidget(self.btn_daily_graphs)

        self.btn_monthly_graphs = QPushButton("Aylık Grafikler")
        self.btn_monthly_graphs.clicked.connect(self.go_to_monthly_graphs)
        self.btn_monthly_graphs.setEnabled(False)
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
            sheets = sorted(list(REQ_SHEETS.intersection(set(xls.sheet_names))))

            if not sheets:
                QMessageBox.warning(self, "Uygun sayfa yok",
                                    "Seçilen dosyada istenen (SMD-OEE, ROBOT, DALGA_LEHİM) sheet bulunamadı.")
                self.reset_page()
                return

            self.main_window.excel_path = Path(path)
            self.lbl_path.setText(f"Seçilen Dosya: <b>{Path(path).name}</b>")

            self.main_window.available_sheets = sheets
            if "SMD-OEE" in sheets:
                self.main_window.selected_sheet = "SMD-OEE"
            elif sheets:
                self.main_window.selected_sheet = sheets[0]
            else:
                self.main_window.selected_sheet = None

            self.btn_daily_graphs.setEnabled(True)
            self.btn_monthly_graphs.setEnabled(True)

            logging.info("Dosya seçildi: %s", path)

        except Exception as e:
            QMessageBox.critical(self, "Okuma hatası",
                                 f"Dosya okunurken bir hata oluştu: {e}\nLütfen dosyanın bozuk olmadığından ve Excel formatında olduğundan emin olun.")
            self.reset_page()

    def go_to_daily_graphs(self) -> None:
        """Günlük grafikler sayfasına geçer."""
        self.main_window.goto_page(1)

    def go_to_monthly_graphs(self) -> None:
        """Aylık grafiklar sayfasına geçer."""
        if "SMD-OEE" not in self.main_window.available_sheets:
            QMessageBox.warning(self, "Uyarı", "Aylık grafikler için 'SMD-OEE' sayfası Excel dosyasında bulunmalıdır.")
            return

        self.main_window.selected_sheet = "SMD-OEE"
        self.main_window.load_excel()

        self.main_window.goto_page(3)

    def reset_page(self):
        """Sayfayı başlangıç durumuna döndürür."""
        self.main_window.excel_path = None
        self.main_window.selected_sheet = None
        self.main_window.available_sheets = []
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

        sheet_selection_group = QHBoxLayout()
        self.sheet_selection_label = QLabel("İşlenecek Sayfa:")
        self.sheet_selection_label.setAlignment(Qt.AlignLeft)
        sheet_selection_group.addWidget(self.sheet_selection_label)

        self.cmb_sheet = QComboBox()
        self.cmb_sheet.setEnabled(False)
        self.cmb_sheet.currentIndexChanged.connect(self.on_sheet_selected)
        sheet_selection_group.addWidget(self.cmb_sheet)
        main_layout.addLayout(sheet_selection_group)

        grouping_group = QHBoxLayout()
        grouping_group.addWidget(QLabel("<b>Gruplama Değişkeni (Tarihler):</b>"))
        self.cmb_grouping = QComboBox()
        self.cmb_grouping.currentIndexChanged.connect(self.populate_grouped)
        grouping_group.addWidget(self.cmb_grouping)
        main_layout.addLayout(grouping_group)

        grouped_group = QHBoxLayout()
        grouped_group.addWidget(QLabel("<b>Gruplanan Değişkenler (Ürünler):</b>"))
        self.lst_grouped = QListWidget()
        self.lst_grouped.setSelectionMode(QListWidget.MultiSelection)
        self.lst_grouped.itemSelectionChanged.connect(self.update_next_button_state)
        grouped_group.addWidget(self.lst_grouped)
        main_layout.addLayout(grouped_group)

        metrics_group = QVBoxLayout()
        metrics_group.addWidget(QLabel("<b>Metrikler :</b>"))
        self.metrics_scroll_area = QScrollArea()
        self.metrics_scroll_area.setWidgetResizable(True)
        self.metrics_content_widget = QWidget()
        self.metrics_layout = QVBoxLayout(self.metrics_content_widget)
        self.metrics_scroll_area.setWidget(self.metrics_content_widget)
        metrics_group.addWidget(self.metrics_scroll_area)
        main_layout.addLayout(metrics_group)

        nav_layout = QHBoxLayout()
        self.btn_back = QPushButton("← Geri")
        self.btn_back.clicked.connect(lambda: self.main_window.goto_page(0))
        nav_layout.addWidget(self.btn_back)

        self.btn_next = QPushButton("İleri →")
        self.btn_next.setEnabled(False)
        self.btn_next.clicked.connect(self.go_next)
        nav_layout.addStretch(1)
        nav_layout.addWidget(self.btn_next)
        main_layout.addLayout(nav_layout)

    def refresh(self) -> None:
        """Sayfa her gösterildiğinde verileri yeniler."""
        self.cmb_sheet.clear()
        if self.main_window.available_sheets:
            self.cmb_sheet.addItems(self.main_window.available_sheets)
            self.cmb_sheet.setEnabled(True)
            if self.main_window.selected_sheet in self.main_window.available_sheets:
                self.cmb_sheet.setCurrentText(self.main_window.selected_sheet)
            else:
                self.cmb_sheet.setCurrentText(self.main_window.available_sheets[0])
        else:
            self.cmb_sheet.setEnabled(False)
            self.main_window.selected_sheet = None
            QMessageBox.warning(self, "Uyarı", "Seçilen Excel dosyasında uygun sayfa bulunamadı.")
            self.main_window.goto_page(0)
            return

    def on_sheet_selected(self) -> None:
        """Sayfa seçimi değiştiğinde ana penceredeki seçimi günceller ve veriyi yeniden yükler."""
        self.main_window.selected_sheet = self.cmb_sheet.currentText()
        self.main_window.load_excel()
        self._populate_data_selection_fields()

    def _populate_data_selection_fields(self):
        """Gruplama, gruplanan ve metrik alanlarını doldurur."""
        df = self.main_window.df
        if df.empty:
            QMessageBox.critical(self, "Hata", "Veri yüklenemedi. Lütfen dosyayı kontrol edin.")
            self.main_window.goto_page(0)
            return

        self.cmb_grouping.clear()
        grouping_col_name = self.main_window.grouping_col_name
        if grouping_col_name and grouping_col_name in df.columns:
            grouping_vals = sorted(df[grouping_col_name].dropna().astype(str).unique())
            grouping_vals = [s for s in grouping_vals if s.strip()]
            self.cmb_grouping.addItems(grouping_vals)
            if not grouping_vals:
                QMessageBox.warning(self, "Uyarı", "Gruplama sütunu (A) boş veya geçerli değer içermiyor.")
        else:
            QMessageBox.warning(self, "Uyarı", "Gruplama sütunu (A) bulunamadı veya boş.")
            self.cmb_grouping.clear()
            self.lst_grouped.clear()
            self.clear_metrics_checkboxes()
            self.update_next_button_state()
            return

        self.populate_metrics_checkboxes()
        self.populate_grouped()

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
                item.setSelected(True)
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

        self.metrics_layout.addStretch(1)
        self.update_next_button_state()

    def clear_metrics_checkboxes(self):
        """Metrik checkbox'larını temizler."""
        while self.metrics_layout.count():
            item = self.metrics_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
            elif isinstance(item, QSpacerItem):
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
        self.main_window.goto_page(2)


class DailyGraphsPage(QWidget):
    """Oluşturulan günlük grafikleri gösteren ve kaydettirme seçenekleri sunan sayfa."""

    def __init__(self, main_window: "MainWindow") -> None:
        super().__init__()
        self.main_window = main_window
        self.worker: GraphWorker | None = None
        self.figures_data: List[Tuple[str, Figure, str]] = []
        self.current_page = 0
        self.current_graph_type = "Donut"
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
        self.progress.hide()
        main_layout.addWidget(self.progress)

        nav_top = QHBoxLayout()
        self.btn_back = QPushButton("← Veri Seçimi")
        self.btn_back.clicked.connect(lambda: self.main_window.goto_page(1))
        nav_top.addWidget(self.btn_back)

        self.lbl_chart_info = QLabel("")
        self.lbl_chart_info.setAlignment(Qt.AlignCenter)
        self.lbl_chart_info.setStyleSheet("font-weight: bold; font-size: 12pt;")
        nav_top.addWidget(self.lbl_chart_info)

        self.cmb_graph_type = QComboBox()
        self.cmb_graph_type.addItems(["Donut", "Bar"])
        self.cmb_graph_type.setCurrentText(self.current_graph_type)
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
        self.enter_page()

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
            total_duration_text = f"TOPLAM DURUŞ\n{total_duration_hours} SAAT {total_duration_minutes} DAKİDA"

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
        self.progress.show()
        self.btn_save_image.setEnabled(False)
        self.lbl_chart_info.setText("Grafikler oluşturuluyor...")
        self.update_page_label()
        self.update_navigation_buttons()

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
            # MODIFICATION 1: Remove "HAT-#" from the displayed title
            display_grouped_val = grouped_val.replace("HAT-#", "").strip()
            self.lbl_chart_info.setText(f"{self.main_window.selected_grouping_val} - {display_grouped_val}")

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
    # Sinyal, hat adı, OEE verisi (liste dict olarak), önceki yıl OEE, önceki ay OEE alacak şekilde güncellendi.
    finished = pyqtSignal(list, object, object)
    progress = pyqtSignal(int)
    error = pyqtSignal(str)

    def __init__(self, df: pd.DataFrame, grouping_col_name: str, grouped_col_name: str, oee_col_name: str,
                 prev_year_oee: float | None, prev_month_oee: float | None, graph_type: str):
        super().__init__()
        self.df = df.copy()
        self.grouping_col_name = grouping_col_name
        self.grouped_col_name = grouped_col_name
        self.oee_col_name = oee_col_name
        self.prev_year_oee = prev_year_oee
        self.prev_month_oee = prev_month_oee
        self.graph_type = graph_type  # Yeni eklenen parametre

    def run(self):
        """İş parçacığı başladığında çalışacak metod."""
        try:
            figures_data: List[Tuple[str, List[dict[str, Any]]]] = []  # Figür yerine veri saklanacak

            df_smd_oee = self.df
            logging.info(f"MonthlyGraphWorker: Başlangıç veri çerçevesi boyutu: {df_smd_oee.shape}")

            # Sütunları dahili tutarlılık için yeniden adlandır
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

            # 'Tarih' sütununu datetime'a dönüştür ve geçersiz tarihleri temizle
            if 'Tarih' in df_smd_oee.columns:
                df_smd_oee['Tarih'] = pd.to_datetime(df_smd_oee['Tarih'], errors='coerce')
                # Sadece geçerli tarihleri tut
                df_smd_oee.dropna(subset=['Tarih'], inplace=True)
                logging.info(
                    f"MonthlyGraphWorker: Tarih sütunu datetime'a dönüştürüldü ve NaT değerleri temizlendi. Yeni boyut: {df_smd_oee.shape}")
            else:
                self.error.emit("'Tarih' sütunu bulunamadı.")
                return

            # 'OEE_Degeri' sütununu float'a dönüştür (sadece OEE grafikleri için)
            if self.graph_type == "OEE Grafikleri":
                if 'OEE_Degeri' in df_smd_oee.columns:
                    logging.info(
                        f"MonthlyGraphWorker: 'OEE_Degeri' sütunu dönüştürme öncesi ilk 5 değer ve tipleri:\n{df_smd_oee['OEE_Degeri'].head().apply(lambda x: f'{x} ({type(x).__name__})')}")

                    # FutureWarning'ı önlemek için inplace=True kaldırıldı
                    df_smd_oee['OEE_Degeri'] = pd.to_numeric(
                        df_smd_oee['OEE_Degeri'].astype(str).str.replace('%', '').str.replace(',', '.'),
                        errors='coerce'
                    )
                    # NaN değerleri 0.0 ile doldur
                    df_smd_oee['OEE_Degeri'] = df_smd_oee['OEE_Degeri'].fillna(0.0)

                    # Eğer OEE değerleri 1'den büyükse (yani % cinsinden ise), 100'e bölerek 0-1 aralığına getir
                    # Örneğin, 85.5 değeri 0.855'e dönüşür.
                    if not df_smd_oee['OEE_Degeri'].empty and df_smd_oee['OEE_Degeri'].max() > 1.0:
                        df_smd_oee['OEE_Degeri'] = df_smd_oee['OEE_Degeri'] / 100.0
                        logging.info("MonthlyGraphWorker: OEE_Degeri sütunu 0-1 aralığına ölçeklendi (100'e bölündü).")

                    logging.info(
                        f"MonthlyGraphWorker: 'OEE_Degeri' sütunu dönüştürme sonrası ve NaN doldurma sonrası ilk 5 değer ve tipleri:\n{df_smd_oee['OEE_Degeri'].head().apply(lambda x: f'{x} ({type(x).__name__})')}")
                    logging.info(
                        f"MonthlyGraphWorker: OEE_Degeri sütunu float'a dönüştürüldü ve geçersiz/boş değerler 0.0 ile dolduruldu. Yeni boyut: {df_smd_oee.shape}")
                else:
                    self.error.emit("'OEE_Degeri' sütunu bulunamadı.")
                    return

            # Dizgi Onay Dağılım Grafiği için 19. indeks (T sütunu)
            dizgi_onay_col_index = excel_col_to_index('T')  # T sütunu (19. indeks)
            dizgi_onay_col_name = self.df.columns[dizgi_onay_col_index] if dizgi_onay_col_index < len(
                self.df.columns) else None

            if self.graph_type == "Dizgi Onay Dağılım Grafiği":
                if not dizgi_onay_col_name or dizgi_onay_col_name not in df_smd_oee.columns:
                    self.error.emit(f"'{dizgi_onay_col_name}' (Dizgi Onay) sütunu bulunamadı veya geçersiz.")
                    return
                # Dizgi Onay sütununu saniyeye dönüştür
                df_smd_oee[dizgi_onay_col_name] = seconds_from_timedelta(df_smd_oee[dizgi_onay_col_name])
                logging.info(f"MonthlyGraphWorker: '{dizgi_onay_col_name}' sütunu saniyeye dönüştürüldü.")

            # 'Group_Key' (örn: "HAT-4") çıkar
            def extract_group_key(s):
                s = str(s).upper()
                match = re.search(r'HAT(\d+)', s)
                if match:
                    hat_number = match.group(1)
                    return f"HAT-{hat_number}"
                return None

            if 'U_Agaci_Sev' in df_smd_oee.columns:
                df_smd_oee['Group_Key'] = df_smd_oee['U_Agaci_Sev'].apply(extract_group_key)
                df_smd_oee.dropna(subset=['Group_Key'], inplace=True)
                logging.info(
                    f"MonthlyGraphWorker: Group_Key sütunu oluşturuldu ve boş değerler temizlendi. Yeni boyut: {df_smd_oee.shape}")
            else:
                self.error.emit("'U_Agaci_Sev' sütunu bulunamadı.")
                return

            unique_hats = sorted(df_smd_oee['Group_Key'].unique())

            # Yalnızca HAT-1, HAT-2, HAT-3, HAT-4'ü filtrele
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

                if self.graph_type == "OEE Grafikleri":
                    # Tarihe göre grupla ve OEE ortalamasını hesapla
                    grouped_oee = df_smd_oee_filtered_by_hat.groupby(pd.Grouper(key='Tarih', freq='D'))[
                        'OEE_Degeri'].mean().reset_index()
                    grouped_oee.dropna(subset=['OEE_Degeri'], inplace=True)
                    figures_data.append((selected_hat, grouped_oee.to_dict('records')))
                    logging.info(f"MonthlyGraphWorker: '{selected_hat}' için OEE verisi hazırlandı.")

                elif self.graph_type == "Dizgi Onay Dağılım Grafiği":
                    # Dizgi Onay metriği için toplam değerleri al
                    current_hat_onay_sum = df_smd_oee_filtered_by_hat[dizgi_onay_col_name].sum()

                    # Diğer hatların Dizgi Onay toplamını hesapla
                    other_hats_df = df_smd_oee[df_smd_oee['Group_Key'] != selected_hat].copy()
                    other_hats_onay_sum = other_hats_df[dizgi_onay_col_name].sum()

                    total_onay_sum = current_hat_onay_sum + other_hats_onay_sum

                    logging.info(
                        f"Dizgi Onay Dağılımı - Hat: {selected_hat}, Bu Hat Toplam: {current_hat_onay_sum:.2f} saniye, Diğer Hatlar Toplam: {other_hats_onay_sum:.2f} saniye")

                    if total_onay_sum > 0:
                        figures_data.append((selected_hat, [
                            {"label": selected_hat, "value": current_hat_onay_sum},
                            {"label": "DİĞER HATLAR", "value": other_hats_onay_sum}
                        ]))
                        logging.info(f"MonthlyGraphWorker: '{selected_hat}' için Dizgi Onay verisi hazırlandı.")
                    else:
                        logging.warning(
                            f"MonthlyGraphWorker: '{selected_hat}' için Dizgi Onay verisi bulunamadı veya toplamı sıfır, atlanıyor.")

                self.progress.emit(int((i + 1) / total_hats * 100))
                logging.info(
                    f"MonthlyGraphWorker: '{selected_hat}' için veri hazırlandı. İlerleme: {int((i + 1) / total_hats * 100)}%")

            # Bitince tüm verileri ve OEE değerlerini sinyal ile gönder
            self.finished.emit(figures_data, self.prev_year_oee, self.prev_month_oee)
            logging.info("MonthlyGraphWorker: Tüm aylık grafik verileri başarıyla hazırlandı.")
        except Exception as exc:
            logging.exception("MonthlyGraphWorker hatası oluştu.")
            self.error.emit(f"Aylık grafik oluşturulurken bir hata oluştu: {str(exc)}")


class MonthlyGraphsPage(QWidget):
    """Aylık grafikler ve veri seçimi sayfasını temsil eder."""

    def __init__(self, main_window: "MainWindow") -> None:
        super().__init__()
        self.main_window = main_window

        self.monthly_chart_container = QFrame(objectName="chartContainer")
        self.monthly_chart_layout = QVBoxLayout(self.monthly_chart_container)
        self.monthly_chart_layout.setAlignment(Qt.AlignCenter)

        self.current_monthly_chart_figure = None
        # Figür yerine sadece veri saklanacak
        self.figures_data_monthly: List[Tuple[str, List[dict[str, Any]]]] = []
        self.current_page_monthly = 0
        self.monthly_worker: MonthlyGraphWorker | None = None
        self.prev_year_oee_for_plot: float | None = None
        self.prev_month_oee_for_plot: float | None = None

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setAlignment(Qt.AlignTop)

        title_label = QLabel("<h2>Aylık Grafikler ve Veri Seçimi</h2>")
        title_label.setObjectName("title_label")
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        graph_type_selection_layout = QHBoxLayout()
        graph_type_selection_layout.addWidget(QLabel("<b>Grafik Tipi:</b>"))
        self.cmb_monthly_graph_type = QComboBox()
        self.cmb_monthly_graph_type.addItems(["OEE Grafikleri", "Dizgi Duruş Grafiği", "Dizgi Onay Dağılım Grafiği"])
        self.cmb_monthly_graph_type.currentIndexChanged.connect(self.on_monthly_graph_type_changed)
        graph_type_selection_layout.addWidget(self.cmb_monthly_graph_type)
        main_layout.addLayout(graph_type_selection_layout)

        self.oee_options_widget = QWidget()
        oee_options_layout = QVBoxLayout(self.oee_options_widget)

        prev_year_oee_layout = QHBoxLayout()
        prev_year_oee_layout.addWidget(QLabel("Önceki Yılın OEE Değeri (%):"))
        self.txt_prev_year_oee = QLineEdit()
        self.txt_prev_year_oee.setPlaceholderText("Örn: 85.5")
        self.txt_prev_year_oee.setValidator(QtGui.QDoubleValidator(0.0, 100.0, 2)) # Added validator
        prev_year_oee_layout.addWidget(self.txt_prev_year_oee)
        oee_options_layout.addLayout(prev_year_oee_layout)

        prev_month_oee_layout = QHBoxLayout()
        prev_month_oee_layout.addWidget(QLabel("Önceki Ayın OEE Değeri (%):"))
        self.txt_prev_month_oee = QLineEdit()
        self.txt_prev_month_oee.setPlaceholderText("Örn: 82.0")
        self.txt_prev_month_oee.setValidator(QtGui.QDoubleValidator(0.0, 100.0, 2)) # Added validator
        prev_month_oee_layout.addWidget(self.txt_prev_month_oee)
        oee_options_layout.addLayout(prev_month_oee_layout)

        oee_buttons_layout = QHBoxLayout()
        self.btn_line_chart = QPushButton("Hat Grafikleri")
        self.btn_line_chart.clicked.connect(self._start_monthly_graph_worker)  # İşlevsellik yeniden bağlandı
        self.btn_line_chart.setEnabled(False)  # Başlangıçta devre dışı
        oee_buttons_layout.addWidget(self.btn_line_chart)

        self.btn_page_chart = QPushButton("Sayfa Grafikleri")
        self.btn_page_chart.setEnabled(False)
        oee_buttons_layout.addWidget(self.btn_page_chart)
        oee_options_layout.addLayout(oee_buttons_layout)

        main_layout.addWidget(self.oee_options_widget)

        self.other_graphs_widget = QWidget()
        other_graphs_layout = QVBoxLayout(self.other_graphs_widget)
        # Placeholder metin kaldırıldı
        main_layout.addWidget(self.other_graphs_widget)
        self.other_graphs_widget.hide()

        # Add a scroll area for the chart container
        self.monthly_chart_scroll_area = QScrollArea()
        self.monthly_chart_scroll_area.setWidgetResizable(True)
        # Yatay kaydırmayı etkinleştir
        self.monthly_chart_scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.monthly_chart_scroll_area.setWidget(self.monthly_chart_container)
        main_layout.addWidget(self.monthly_chart_scroll_area)  # Add the scroll area to the main layout

        self.monthly_progress = QProgressBar()
        self.monthly_progress.setAlignment(Qt.AlignCenter)
        self.monthly_progress.setTextVisible(True)
        self.monthly_progress.hide()
        main_layout.addWidget(self.monthly_progress)

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

        self.on_monthly_graph_type_changed(0)

    def enter_page(self) -> None:
        """Bu sayfaya girildiğinde grafiği temizler ve buton durumlarını günceller."""
        self.clear_monthly_chart_canvas()
        self.btn_save_monthly_chart.setEnabled(False)
        self.update_monthly_page_label()
        self.update_monthly_navigation_buttons()
        # Sadece OEE Grafikleri seçeneği aktifken Hat Grafikleri butonu aktif olacak
        if self.cmb_monthly_graph_type.currentText() == "OEE Grafikleri":
            self.btn_line_chart.setEnabled(True)
        else:
            self.btn_line_chart.setEnabled(False)

    def on_monthly_graph_type_changed(self, index: int):
        """Aylık grafik tipi seçimi değiştiğinde ilgili seçenekleri gösterir/gizler."""
        selected_type = self.cmb_monthly_graph_type.currentText()
        self.clear_monthly_chart_canvas()
        self.btn_save_monthly_chart.setEnabled(False)
        self.figures_data_monthly.clear()
        self.current_page_monthly = 0
        self.update_monthly_page_label()
        self.update_monthly_navigation_buttons()

        if selected_type == "OEE Grafikleri":
            self.oee_options_widget.show()
            self.other_graphs_widget.hide()
            self.btn_line_chart.setEnabled(True)
        elif selected_type == "Dizgi Onay Dağılım Grafiği":
            self.oee_options_widget.hide()
            self.other_graphs_widget.show()
            self.btn_line_chart.setEnabled(False)  # Disable the button as it will be triggered automatically
            self._start_monthly_graph_worker()  # Automatically start worker for Dizgi Onay
        else:
            self.oee_options_widget.hide()
            self.other_graphs_widget.show()
            self.btn_line_chart.setEnabled(False)

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

        prev_year_oee = None
        prev_month_oee = None

        try:
            if self.txt_prev_year_oee.text():
                prev_year_oee = float(self.txt_prev_year_oee.text().replace(",", "."))
            if self.txt_prev_month_oee.text():
                prev_month_oee = float(self.txt_prev_month_oee.text().replace(",", "."))
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
            prev_month_oee=prev_month_oee,
            graph_type=self.cmb_monthly_graph_type.currentText()  # Yeni parametre
        )
        self.monthly_worker.finished.connect(self._on_monthly_graphs_generated)
        self.monthly_worker.progress.connect(self.monthly_progress.setValue)
        self.monthly_worker.error.connect(self._on_monthly_graph_error)
        self.monthly_worker.start()

    def _on_monthly_graphs_generated(self, figures_data_raw: List[Tuple[str, List[dict[str, Any]]]],
                                     prev_year_oee: float | None, prev_month_oee: float | None):
        """MonthlyGraphWorker'dan gelen sonuçları işler."""
        self.monthly_progress.setValue(100)
        self.monthly_progress.hide()

        if not figures_data_raw:
            QMessageBox.information(self, "Veri Yok",
                                    "Aylık grafik oluşturulamadı. Seçilen kriterlere göre veri bulunamadı.")
            self.btn_save_monthly_chart.setEnabled(False)
            return

        self.figures_data_monthly = figures_data_raw
        self.prev_year_oee_for_plot = prev_year_oee
        self.prev_month_oee_for_plot = prev_month_oee
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

        hat_name, data_list = self.figures_data_monthly[self.current_page_monthly]

        # Güncellenen boyutlar
        fig_width_inches = 8.0  # 800 piksel
        fig_height_inches = 5.0  # 500 piksel

        fig, ax = plt.subplots(figsize=(fig_width_inches, fig_height_inches), dpi=100)
        background_color = 'white'
        fig.patch.set_facecolor(background_color)
        ax.set_facecolor(background_color)

        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_visible(False)
        ax.spines['bottom'].set_visible(False)
        ax.grid(False)

        if self.cmb_monthly_graph_type.currentText() == "OEE Grafikleri":
            # Saklanan veriden DataFrame'i yeniden oluştur
            grouped_oee = pd.DataFrame(data_list)
            # 'Tarih' sütununun datetime olduğundan emin ol
            grouped_oee['Tarih'] = pd.to_datetime(grouped_oee['Tarih'])

            dates = grouped_oee['Tarih']
            oee_values = grouped_oee['OEE_Degeri']

            line_color = '#1f77b4'

            # X ekseninde eşit aralıklarla konumlandırmak için sayısal bir indeks kullan
            x_indices = np.arange(len(dates))
            ax.plot(x_indices, oee_values, marker='o', markersize=8, color=line_color, linewidth=2, label=hat_name)
            ax.plot(x_indices, oee_values, 'o', markersize=6, color='white', markeredgecolor=line_color,
                    markeredgewidth=1.5, zorder=5)

            for i, (x, y) in enumerate(zip(x_indices, oee_values)):
                # %0 olan değerlerin üzerinde "%0" belirteci olmasın
                if pd.notna(y) and y > 0:
                    ax.annotate(f'{y * 100:.1f}%', (x, y), textcoords="offset points", xytext=(0, 10), ha='center',
                                fontsize=8, fontweight='bold')

            overall_calculated_average = np.mean(oee_values) if not oee_values.empty else 0

            # Horizontal line for previous year OEE
            if self.prev_year_oee_for_plot is not None:
                y_val = self.prev_year_oee_for_plot / 100
                ax.axhline(y_val, color='red', linestyle='--', linewidth=1.5,
                           label=f'Önceki Yıl OEE ({self.prev_year_oee_for_plot:.1f}%)')
                # Yüzde değerini sağa, eksenin dışına yaz
                ax.text(1.01, y_val, f'{self.prev_year_oee_for_plot:.1f}%',
                        transform=ax.transAxes, color='red', va='center', ha='left', fontsize=9, fontweight='bold')

            # Horizontal line for previous month OEE
            if self.prev_month_oee_for_plot is not None:
                y_val = self.prev_month_oee_for_plot / 100
                ax.axhline(y_val, color='orange', linestyle='--', linewidth=1.5,
                           label=f'Önceki Ay OEE ({self.prev_month_oee_for_plot:.1f}%)')
                # Yüzde değerini sağa, eksenin dışına yaz
                ax.text(1.01, y_val, f'{self.prev_month_oee_for_plot:.1f}%',
                        transform=ax.transAxes, color='orange', va='center', ha='left', fontsize=9, fontweight='bold')

            # Horizontal line for calculated average OEE
            if overall_calculated_average > 0:
                y_val = overall_calculated_average
                # Ay adını Türkçe ve büyük harfle al
                month_names_turkish = {
                    1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
                    7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
                }
                month_name = ""
                if not dates.empty:
                    first_date_in_data = dates.min()
                    month_name = month_names_turkish.get(first_date_in_data.month,
                                                         first_date_in_data.strftime('%B')).capitalize()
                else:
                    month_name = datetime.date.today().strftime('%B').capitalize()

                ax.axhline(y_val, color='purple', linestyle='--', linewidth=1.5,
                           label=f'{month_name} OEE ({overall_calculated_average * 100:.1f}%)')
                # Yüzde değerini sağa, eksenin dışına yaz
                ax.text(1.01, y_val, f'{overall_calculated_average * 100:.1f}%',
                        transform=ax.transAxes, color='purple', va='center', ha='left', fontsize=9, fontweight='bold')

            # X ekseni işaretçilerini sayısal indekslere ayarla
            ax.set_xticks(x_indices)
            # X ekseni etiketlerini orijinal tarihlerle ayarla
            ax.set_xticklabels([d.strftime('%d.%m.%Y') for d in dates])

            fig.autofmt_xdate(rotation=45)  # Tarih etiketlerini döndür

            # Y ekseni etiketlerini yüzde olarak formatla
            # Y ekseni etiketlerini tam sayı yüzde olarak formatla
            ax.yaxis.set_major_formatter(PercentFormatter(xmax=1, decimals=0))

            # Y ekseni ana işaretçilerini %25'luk artışlarla ayarla
            ax.set_yticks(np.arange(0.0, 1.001, 0.25))
            # Y ekseni limitlerini 0% ile 100% aralığına sabitle, alt ve üst limitler için küçük bir boşluk bırak
            ax.set_ylim(bottom=-0.05, top=1.05)

            ax.set_xlabel("Tarih", fontsize=12, fontweight='bold')
            ax.set_ylabel("OEE (%)", fontsize=12, fontweight='bold')

            # Ay adını Türkçe ve büyük harfle al
            month_names_turkish = {
                1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
                7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
            }
            month_name = ""
            if not dates.empty:
                first_date_in_data = dates.min()
                month_name = month_names_turkish.get(first_date_in_data.month,
                                                     first_date_in_data.strftime('%B')).capitalize()
            else:
                month_name = datetime.date.today().strftime('%B').capitalize()

            chart_title = f"{hat_name} {month_name} OEE"
            ax.set_title(chart_title, fontsize=16, color='#2c3e50', fontweight='bold')

            # Legend'ı sağa, dikeyde ortaya yerleştir ve grafiğe alan aç
            ax.legend(loc='upper left', bbox_to_anchor=(1.02, 0), fontsize=10)
            # Legend ve yeni eklenen yüzde etiketleri için yeterli boşluk bırak
            fig.subplots_adjust(right=0.60) # Reverted right margin

        elif self.cmb_monthly_graph_type.currentText() == "Dizgi Onay Dağılım Grafiği":
            labels = [d["label"] for d in data_list]
            values = [d["value"] for d in data_list]

            # Güncellenmiş renkler: Koyu Mavi ve Turuncu
            colors = ['#00008B', '#FFA500']

            total_sum = sum(values)

            # Autopct formatını özelleştir
            def func(pct, allvals):
                absolute = int(np.round(pct / 100. * total_sum))
                # Saniyeyi HH:MM:SS formatına dönüştür
                hours = absolute // 3600
                minutes = (absolute % 3600) // 60
                seconds = absolute % 60
                return f"{hours:02d}:{minutes:02d}:{seconds:02d}; {pct:.0f}%"

            wedges, texts, autotexts = ax.pie(
                values,
                autopct=lambda pct: func(pct, values),
                startangle=90,
                colors=colors,
                # width parametresi kaldırıldı, böylece donut değil tam pasta grafiği olur
                wedgeprops=dict(edgecolor='black', linewidth=1.5)
            )

            # Autotexts (yüzde etiketleri) için renk ve stil ayarı
            for autotext in autotexts:
                autotext.set_color('white')  # Metin rengi beyaza çevrildi
                autotext.set_fontsize(14)  # Font boyutu artırıldı
                autotext.set_fontweight('bold')

            ax.axis('equal')  # Dairenin orantılı olmasını sağlar

            # Legend'ı sağ üst köşeye taşı
            ax.legend(wedges, labels,
                      title="Hatlar",
                      loc="upper right",
                      bbox_to_anchor=(1.2, 1),
                      fontsize=10,
                      title_fontsize=12)

            # MODIFICATION: Use the exact chart title as requested
            chart_title = f"Dizgi Onay Dağılımı"
            ax.set_title(chart_title, fontsize=16, color='#2c3e50', fontweight='bold')
            fig.tight_layout()  # Grafiğin sıkışmasını önle

        canvas = FigureCanvas(fig)
        # Canvas boyutunu da grafiğin boyutuna göre ayarla
        canvas.setFixedSize(int(fig_width_inches * fig.dpi), int(fig_height_inches * fig.dpi))
        self.monthly_chart_layout.addWidget(canvas, stretch=1)
        canvas.draw()

        self.current_monthly_chart_figure = fig
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
        self.available_sheets: List[str] = []
        self.df: pd.DataFrame = pd.DataFrame()
        self.grouping_col_name: str | None = None
        self.grouped_col_name: str | None = None
        self.oee_col_name: str | None = None
        self.metric_cols: List[str] = []
        self.grouped_values: List[str] = []
        self.selected_metrics: List[str] = []
        self.selected_grouping_val: str = "" # Initialize selected_grouping_val

        self.stacked_widget = QStackedWidget()
        self.file_selection_page = FileSelectionPage(self)
        self.data_selection_page = DataSelectionPage(self)
        self.daily_graphs_page = DailyGraphsPage(self)
        self.monthly_graphs_page = MonthlyGraphsPage(self)

        self.stacked_widget.addWidget(self.file_selection_page)
        self.stacked_widget.addWidget(self.data_selection_page)
        self.stacked_widget.addWidget(self.daily_graphs_page)
        self.stacked_widget.addWidget(self.monthly_graphs_page)

        self.setCentralWidget(self.stacked_widget)
        self.setWindowTitle("OEE ve Durum Grafiği Uygulaması")
        self.setGeometry(100, 100, 1200, 800)

        self.apply_stylesheet()
        self.goto_page(0)

    def apply_stylesheet(self):
        """Uygulamaya modern bir stil uygular."""
        self.setStyleSheet("""
            QWidget {
                background-color: #f0f2f5;
                font-family: 'Segoe UI', Arial, sans-serif;
                color: #333333;
            }
            QLabel#title_label {
                font-size: 28pt;
                font-weight: bold;
                color: #2c3e50;
                margin-bottom: 25px;
                padding: 10px;
            }
            QLabel {
                font-size: 11pt;
            }
            QPushButton {
                background-color: #3498db;
                color: white;
                padding: 12px 25px;
                border-radius: 8px;
                border: none;
                font-weight: bold;
                font-size: 11pt;
                margin: 8px;
                box-shadow: 2px 2px 5px rgba(0, 0, 0, 0.2);
            }
            QPushButton:hover {
                background-color: #2980b9;
                box-shadow: 3px 3px 8px rgba(0, 0, 0, 0.3);
            }
            QPushButton:pressed {
                background-color: #21618c;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
                box-shadow: none;
            }
            QComboBox, QListWidget, QScrollArea, QProgressBar, QFrame, QLineEdit {
                border: 1px solid #dcdcdc;
                border-radius: 6px;
                padding: 8px;
                background-color: white;
                selection-background-color: #aed6f1;
                selection-color: black;
            }
            QComboBox::drop-down {
                border: 0px;
            }
            QComboBox::down-arrow {
                image: url(down_arrow.png);
                width: 12px;
                height: 12px;
            }
            QListWidget::item {
                padding: 5px;
            }
            QListWidget::item:selected {
                background-color: #3498db;
                color: white;
                border-radius: 3px;
            }
            QCheckBox {
                spacing: 8px;
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
                image: url(check_mark.png);
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
                background-color: #2ecc71;
                border-radius: 7px;
            }
            QScrollArea > QWidget > QWidget {
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
        if index == 1:
            self.data_selection_page.refresh()
        elif index == 2:
            self.daily_graphs_page.enter_page()
        elif index == 3:
            self.monthly_graphs_page.enter_page()

    def load_excel(self) -> None:
        """Seçilen Excel dosyasını ve sayfasını yükler."""
        if not self.excel_path or not self.selected_sheet:
            logging.warning("load_excel: Excel yolu veya seçili sayfa boş. Veri yüklenemiyor.")
            return

        if not self.df.empty and self.df.attrs.get('excel_path') == self.excel_path and \
                self.df.attrs.get('selected_sheet') == self.selected_sheet:
            logging.info(f"Veri '{self.selected_sheet}' sayfasından zaten yüklü. Tekrar yüklenmiyor.")
            return

        try:
            # Always read the header (first row) for all sheets
            self.df = pd.read_excel(self.excel_path, sheet_name=self.selected_sheet, header=0)
            self.df.columns = self.df.columns.astype(str) # Ensure column names are strings

            self.df.attrs['excel_path'] = self.excel_path
            self.df.attrs['selected_sheet'] = self.selected_sheet

            logging.info("'%s' sayfasından veri yüklendi. Satır sayısı: %d", self.selected_sheet, len(self.df))

            # Initialize with default values, then override based on sheet
            self.grouping_col_name = self.df.columns[excel_col_to_index('A')]
            self.grouped_col_name = self.df.columns[excel_col_to_index('B')]
            self.oee_col_name = None # Default to None, set if applicable
            self.metric_cols = []

            if self.selected_sheet == "SMD-OEE":
                self.oee_col_name = self.df.columns[excel_col_to_index('BP')] if excel_col_to_index('BP') < len(self.df.columns) else None
                start_col_index = excel_col_to_index('H')
                end_col_index = excel_col_to_index('BD')
                ap_col_index = excel_col_to_index('AP')
                for i in range(start_col_index, end_col_index + 1):
                    if i < len(self.df.columns) and i != ap_col_index:
                        self.metric_cols.append(self.df.columns[i])
            elif self.selected_sheet == "ROBOT":
                # For ROBOT, metrics are from H to AU
                # OEE column is explicitly None based on previous logs
                start_col_index = excel_col_to_index('H')
                end_col_index = excel_col_to_index('AU') # As per user request
                ao_col_index = excel_col_to_index('AO') # Get index for 'AO' column

                for i in range(start_col_index, end_col_index + 1):
                    if i < len(self.df.columns) and i != ao_col_index: # Skip 'AO' column
                        self.metric_cols.append(self.df.columns[i])
            elif self.selected_sheet == "DALGA_LEHİM":
                # Assuming similar structure to SMD-OEE for now, adjust if needed
                self.oee_col_name = self.df.columns[excel_col_to_index('BP')] if excel_col_to_index('BP') < len(self.df.columns) else None
                start_col_index = excel_col_to_index('H')
                end_col_index = excel_col_to_index('BD')
                ap_col_index = excel_col_to_index('AP')
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
            self.df = pd.DataFrame()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

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
