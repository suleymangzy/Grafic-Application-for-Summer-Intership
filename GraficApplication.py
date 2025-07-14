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
from PyQt5 import QtGui

# Her sayfada kaç grafik gösterileceği
GRAPHS_PER_PAGE = 1
# Gerekli Excel sayfa isimleri
REQ_SHEETS = {"SMD-OEE", "ROBOT", "DALGA_LEHİM", "KAPLAMA-OEE"}  # "KAPLAMA-OEE" eklendi

# Loglama ayarları
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)

# Matplotlib font settings (Turkish character support)
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.sans-serif'] = ['SimSun', 'Arial', 'Liberation Sans', 'Bitstream Vera Sans', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False  # For negative signs

# Global matplotlib settings
plt.rcParams['axes.grid'] = True
plt.rcParams['grid.alpha'] = 0.7
plt.rcParams['grid.linestyle'] = '--'
plt.rcParams['grid.linewidth'] = 0.5
plt.rcParams['figure.dpi'] = 100
plt.rcParams['savefig.dpi'] = 300

# Tick settings
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
    """Pandas Serisindeki zaman değerlerini saniyeye dönüştürür."""
    seconds_series = pd.Series(0.0, index=series.index, dtype=float)

    is_time_obj = series.apply(lambda x: isinstance(x, datetime.time))
    if is_time_obj.any():
        time_objects = series[is_time_obj]
        seconds_series.loc[is_time_obj] = time_objects.apply(
            lambda t: t.hour * 3600 + t.minute * 60 + t.second + t.microsecond / 1e6
        )

    str_and_timedelta_mask = ~is_time_obj & series.notna()
    if str_and_timedelta_mask.any():
        converted_td = pd.to_timedelta(series.loc[str_and_timedelta_mask].astype(str).strip(), errors='coerce')
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
    """Arka planda günlük grafik oluşturma için çalışan sınıf."""
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
                                raise ValueError("Desteklenmeyen OEE değeri tipi veya formatı")

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
        """Donut grafik oluşturur."""
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
            duration_seconds = int(metric_value % 60)
            label_text = (
                f"{i + 1}. {metric_name}; "
                f"{duration_hours:02d}:"
                f"{duration_minutes:02d}:"
                f"{duration_seconds:02d}; "
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
        """Çubuk grafik oluşturur."""
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
            duration_seconds = int(value % 60)
            text_label = f"{duration_hours:02d}:{duration_minutes:02d}:{duration_seconds:02d} ({percentage:.0f}%)"

            text_x_position = (value / 60) + 0.5
            ax.text(text_x_position, i, text_label,
                    va='center', ha='left',
                    fontsize=11, fontweight='bold',
                    color='black')

        ax.set_xlim(left=0)
        plt.tight_layout(rect=[0.1, 0.1, 0.95, 0.9])


class FileSelectionPage(QWidget):
    """Dosya seçim sayfasını temsil eder."""

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
        """Kullanıcının bir Excel dosyası seçmesine izin verir."""
        path, _ = QFileDialog.getOpenFileName(self, "Excel seç", str(Path.home()), "Excel Files (*.xlsx)")
        if not path:
            return

        try:
            xls = pd.ExcelFile(path)
            sheets = sorted(list(REQ_SHEETS.intersection(set(xls.sheet_names))))

            if not sheets:
                QMessageBox.warning(self, "Uygun sayfa yok",
                                    f"Seçilen dosyada istenen ({', '.join(REQ_SHEETS)}) sheet bulunamadı.")
                self.reset_page()
                return

            self.main_window.excel_path = Path(path)
            self.lbl_path.setText(f"Seçilen Dosya: <b>{Path(path).name}</b>")

            self.main_window.available_sheets = sheets
            # SMD-OEE'yi varsayılan olarak seç, yoksa ilk uygun sayfayı seç
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
        """Günlük grafikler sayfasına gider."""
        self.main_window.goto_page(1)

    def go_to_monthly_graphs(self) -> None:
        """Aylık grafikler sayfasına gider."""
        # Aylık grafikler için SMD-OEE sayfasının yüklenmesi önemli,
        # çünkü Hat Grafikleri bu sayfadaki verilere dayanır.
        # Sayfa Grafikleri ise kendi içinde ilgili sayfaları yükleyecektir.
        if "SMD-OEE" in self.main_window.available_sheets:
            self.main_window.selected_sheet = "SMD-OEE"
        elif self.main_window.available_sheets:
            self.main_window.selected_sheet = self.main_window.available_sheets[0]
        else:
            QMessageBox.warning(self, "Uyarı", "Aylık grafikler için uygun sayfa bulunamadı.")
            return

        self.main_window.load_excel()
        self.main_window.goto_page(3)

    def reset_page(self):
        """Sayfayı başlangıç durumuna sıfırlar."""
        self.main_window.excel_path = None
        self.main_window.selected_sheet = None
        self.main_window.available_sheets = []
        self.lbl_path.setText("Henüz dosya seçilmedi")
        self.btn_daily_graphs.setEnabled(False)
        self.btn_monthly_graphs.setEnabled(False)


class DataSelectionPage(QWidget):
    """Veri seçim sayfasını (gruplama, metrikler vb. - Günlük Grafikler için) temsil eder."""

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
        """Sayfa görüntülendiğinde verileri yeniler."""
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
        """Sayfa seçimi değiştiğinde ana penceredeki seçili sayfayı günceller ve verileri yeniden yükler."""
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
        """Gruplanan değişkenler (ürünler) listesini doldurur."""
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
        """Metrik sütunları için onay kutuları oluşturur ve doldurur."""
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
        """Metrik onay kutularını temizler."""
        while self.metrics_layout.count():
            item = self.metrics_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
            elif isinstance(item, QSpacerItem):
                self.metrics_layout.removeItem(item)

    def on_metric_checkbox_changed(self, state):
        """Bir metrik onay kutusunun durumu değiştiğinde çağrılır."""
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
        """İleri düğmesinin etkin durumunu günceller."""
        is_grouped_selected = bool(self.lst_grouped.selectedItems())
        is_metric_selected = bool(self.main_window.selected_metrics)
        self.btn_next.setEnabled(is_grouped_selected and is_metric_selected)

    def go_next(self) -> None:
        """Bir sonraki sayfaya gitmek için verileri hazırlar."""
        self.main_window.grouped_values = [i.text() for i in self.lst_grouped.selectedItems()]
        self.main_window.selected_grouping_val = self.cmb_grouping.currentText()
        if not self.main_window.grouped_values or not self.main_window.selected_metrics:
            QMessageBox.warning(self, "Seçim Eksik", "Lütfen en az bir gruplanan değişken ve bir metrik seçin.")
            return
        self.main_window.goto_page(2)


class DailyGraphsPage(QWidget):
    """Oluşturulan günlük grafikleri görüntüleyen ve kaydetme seçenekleri sunan sayfa."""

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
        """Grafik türü değiştiğinde çağrılır ve grafikleri yeniden çizer."""
        self.current_graph_type = self.cmb_graph_type.currentText()
        self.enter_page()

    def on_results(self, results: List[Tuple[str, pd.Series, str]]) -> None:
        """GraphWorker'dan gelen sonuçları işler ve grafikler oluşturur."""
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
        """Bu sayfaya girildiğinde grafik yeniden oluşturma sürecini başlatır."""
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
        """GraphWorker'dan gelen bir hata mesajını görüntüler."""
        QMessageBox.critical(self, "Hata", message)
        self.progress.setValue(0)
        self.progress.hide()
        self.lbl_chart_info.setText("Grafik oluşturma hatası.")
        self.btn_save_image.setEnabled(False)

    def clear_canvases(self) -> None:
        """Mevcut grafik tuvalini temizler."""
        while self.vbox_canvases.count():
            item = self.vbox_canvases.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def display_current_page_graphs(self) -> None:
        """Mevcut sayfadaki grafikleri görüntüler."""
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
        """Gezinme düğmelerinin etkin durumunu günceller."""
        total_pages = (len(self.figures_data) + GRAPHS_PER_PAGE - 1) // GRAPHS_PER_PAGE if self.figures_data else 0
        self.btn_prev.setEnabled(self.current_page > 0)
        self.btn_next.setEnabled(self.current_page < total_pages - 1)

    def prev_page(self) -> None:
        """Önceki sayfaya gider."""
        if self.current_page > 0:
            self.current_page -= 1
            self.display_current_page_graphs()

    def next_page(self) -> None:
        """Sonraki sayfaya gider."""
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
    """Aylık grafik oluşturma için arka planda çalışan sınıf."""
    finished = pyqtSignal(list, object, object)
    progress = pyqtSignal(int)
    error = pyqtSignal(str)

    def __init__(self, excel_path: Path, current_df: pd.DataFrame, graph_mode: str, graph_type: str,
                 prev_year_oee: float | None, prev_month_oee: float | None, main_window: "MainWindow"):
        super().__init__()
        self.excel_path = excel_path
        self.current_df = current_df  # This is the df from main_window, usually SMD-OEE
        self.graph_mode = graph_mode
        self.graph_type = graph_type
        self.prev_year_oee = prev_year_oee
        self.prev_month_oee = prev_month_oee
        self.main_window = main_window  # Keep reference to main_window to access its properties

    def run(self):
        """İş parçacığı başladığında çalışacak metod."""
        try:
            figures_data: List[Tuple[str, Union[List[dict[str, Any]], Dict[str, Any]]]] = []

            if self.graph_mode == "hat":
                df_to_process = self.current_df.copy()
                grouping_col_name = self.main_window.grouping_col_name
                grouped_col_name = self.main_window.grouped_col_name
                oee_col_name = self.main_window.oee_col_name

                # Sütunları dahili tutarlılık için yeniden adlandır
                col_mapping = {}
                if grouping_col_name in df_to_process.columns:
                    col_mapping[grouping_col_name] = 'Tarih'
                if grouped_col_name in df_to_process.columns:
                    col_mapping[grouped_col_name] = 'U_Agaci_Sev'
                if oee_col_name and oee_col_name in df_to_process.columns:
                    col_mapping[oee_col_name] = 'OEE_Degeri'

                if col_mapping:
                    df_to_process.rename(columns=col_mapping, inplace=True)
                else:
                    self.error.emit("Gerekli sütunlar Excel dosyasında bulunamadı veya adlandırılamadı.")
                    return

                # 'Tarih' sütununu datetime'a dönüştür ve geçersiz tarihleri temizle
                if 'Tarih' in df_to_process.columns:
                    df_to_process['Tarih'] = pd.to_datetime(df_to_process['Tarih'], errors='coerce')
                    df_to_process.dropna(subset=['Tarih'], inplace=True)
                else:
                    self.error.emit("'Tarih' sütunu bulunamadı.")
                    return

                # 'OEE_Degeri' sütununu float'a dönüştür (sadece OEE grafikleri için)
                if self.graph_type == "OEE Grafikleri":
                    if 'OEE_Degeri' in df_to_process.columns:
                        df_to_process['OEE_Degeri'] = pd.to_numeric(
                            df_to_process['OEE_Degeri'].astype(str).replace('%', '').replace(',', '.'),
                            errors='coerce'
                        )
                        df_to_process['OEE_Degeri'] = df_to_process['OEE_Degeri'].fillna(0.0)

                        if not df_to_process['OEE_Degeri'].empty and df_to_process['OEE_Degeri'].max() > 1.0:
                            df_to_process['OEE_Degeri'] = df_to_process['OEE_Degeri'] / 100.0
                    else:
                        self.error.emit("'OEE_Degeri' sütunu bulunamadı.")
                        return

                # Dizgi Onay Dağılım Grafiği için 19. indeksi (T sütunu) al
                dizgi_onay_col_index = excel_col_to_index('T')
                dizgi_onay_col_name = self.current_df.columns[dizgi_onay_col_index] if dizgi_onay_col_index < len(
                    self.current_df.columns) else None

                # Dizgi Duruş Grafiği için metrik sütunlarını al
                dizgi_durusu_metric_cols = []
                start_col_index = excel_col_to_index('H')
                end_col_index = excel_col_to_index('BD')
                for i in range(start_col_index, end_col_index + 1):
                    col_name = self.current_df.columns[i]
                    if i < len(self.current_df.columns):
                        dizgi_durusu_metric_cols.append(col_name)

                if self.graph_type == "Dizgi Onay Dağılım Grafiği":
                    if not dizgi_onay_col_name or dizgi_onay_col_name not in df_to_process.columns:
                        self.error.emit(f"'{dizgi_onay_col_name}' (Dizgi Onay) sütunu bulunamadı veya geçersiz.")
                        return
                    df_to_process[dizgi_onay_col_name] = seconds_from_timedelta(df_to_process[dizgi_onay_col_name])
                elif self.graph_type == "Dizgi Duruş Grafiği":
                    if not dizgi_durusu_metric_cols:
                        self.error.emit("Dizgi Duruş Grafiği için metrik sütunları bulunamadı.")
                        return
                    for col in dizgi_durusu_metric_cols:
                        if col in df_to_process.columns:
                            df_to_process[col] = seconds_from_timedelta(df_to_process[col])

                # 'Group_Key' (örn: "HAT-4") çıkar
                def extract_group_key(s):
                    s = str(s).upper()
                    match = re.search(r'HAT(\d+)', s)
                    if match:
                        hat_number = match.group(1)
                        return f"HAT-{hat_number}"
                    return None

                if 'U_Agaci_Sev' in df_to_process.columns:
                    df_to_process['Group_Key'] = df_to_process['U_Agaci_Sev'].apply(extract_group_key)
                    df_to_process.dropna(subset=['Group_Key'], inplace=True)
                else:
                    self.error.emit("'U_Agaci_Sev' sütunu bulunamadı.")
                    return

                unique_hats = sorted(df_to_process['Group_Key'].unique())
                target_hat_patterns = {"HAT-1", "HAT-2", "HAT-3", "HAT-4"}
                filtered_hats = [hat for hat in unique_hats if hat in target_hat_patterns]
                unique_hats = sorted(filtered_hats)
                total_items = len(unique_hats)

                if not unique_hats and self.graph_type != "Dizgi Duruş Grafiği":
                    self.error.emit(
                        "Grafik oluşturmak için hat verisi bulunamadı. Lütfen Excel dosyasının 'HAT-1', 'HAT-2', 'HAT-3' veya 'HAT-4' için veri içerdiğinden emin olun.")
                    return

                if self.graph_type == "Dizgi Duruş Grafiği":
                    all_metrics_sum = df_to_process[dizgi_durusu_metric_cols].sum()
                    total_sum_of_all_metrics = all_metrics_sum.sum()
                    metric_sums = all_metrics_sum[all_metrics_sum > 0].sort_values(ascending=False)
                    cumulative_sum_for_line = metric_sums.cumsum()
                    cumulative_percentage_for_line = (cumulative_sum_for_line / total_sum_of_all_metrics) * 100

                    pareto_metrics_to_plot = pd.Series(dtype=float)
                    current_cumulative_percent = 0.0
                    for idx, (metric_name, value) in enumerate(metric_sums.items()):
                        percent_of_total = (value / total_sum_of_all_metrics) * 100
                        current_cumulative_percent += percent_of_total
                        pareto_metrics_to_plot[metric_name] = value
                        if current_cumulative_percent >= 80:
                            if current_cumulative_percent - percent_of_total >= 80 and (
                                    current_cumulative_percent - 80) > 10:
                                pareto_metrics_to_plot = pareto_metrics_to_plot.iloc[:-1]
                            break
                    if pareto_metrics_to_plot.empty and not metric_sums.empty:
                        pareto_metrics_to_plot = metric_sums.head(1)

                    figures_data.append(("Genel Dizgi Duruş", {
                        "metrics": pareto_metrics_to_plot.to_dict(),
                        "total_overall_sum": total_sum_of_all_metrics,
                        "cumulative_percentages": cumulative_percentage_for_line[pareto_metrics_to_plot.index].to_dict()
                    }))
                    self.progress.emit(100)

                else:
                    for i, selected_hat in enumerate(unique_hats):
                        df_smd_oee_filtered_by_hat = df_to_process[df_to_process['Group_Key'] == selected_hat].copy()
                        if df_smd_oee_filtered_by_hat.empty:
                            self.progress.emit(int((i + 1) / total_items * 100))
                            continue

                        if self.graph_type == "OEE Grafikleri":
                            grouped_oee = df_smd_oee_filtered_by_hat.groupby(pd.Grouper(key='Tarih', freq='D'))[
                                'OEE_Degeri'].mean().reset_index()
                            grouped_oee.dropna(subset=['OEE_Degeri'], inplace=True)
                            figures_data.append((selected_hat, grouped_oee.to_dict('records')))

                        elif self.graph_type == "Dizgi Onay Dağılım Grafiği":
                            current_hat_onay_sum = df_smd_oee_filtered_by_hat[dizgi_onay_col_name].sum()
                            other_hats_df = df_to_process[df_to_process['Group_Key'] != selected_hat].copy()
                            other_hats_onay_sum = other_hats_df[dizgi_onay_col_name].sum()
                            total_onay_sum = current_hat_onay_sum + other_hats_onay_sum
                            if total_onay_sum > 0:
                                figures_data.append((selected_hat, [
                                    {"label": selected_hat, "value": current_hat_onay_sum},
                                    {"label": "DİĞER HATLAR", "value": other_hats_onay_sum}
                                ]))
                        self.progress.emit(int((i + 1) / total_items * 100))

            elif self.graph_mode == "page" and self.graph_type == "OEE Grafikleri":
                # Define sheets and their OEE column letters
                sheets_to_process_info = [
                    ("DALGA_LEHİM", "BP"),
                    ("ROBOT", "BG"),
                    ("SMD-OEE", "BP"),  # KAPLAMA-OEE yerine SMD-OEE kullanıldı, BP sütunu için
                    ("KAPLAMA-OEE", "BG")  # Eğer ayrı bir KAPLAMA-OEE sayfası varsa ve BG sütunu kullanıyorsa
                ]

                # Filter sheets based on availability in the loaded Excel file
                available_sheets_for_page_mode = [
                    (sheet_name, oee_col) for sheet_name, oee_col in sheets_to_process_info
                    if sheet_name in self.main_window.available_sheets
                ]

                total_items = len(available_sheets_for_page_mode)
                if not available_sheets_for_page_mode:
                    self.error.emit(
                        "Sayfa grafikleri için işlenecek uygun sayfa bulunamadı (DALGA_LEHİM, ROBOT, SMD-OEE/KAPLAMA-OEE).")
                    return

                for i, (sheet_name, oee_col_letter) in enumerate(available_sheets_for_page_mode):
                    logging.info(
                        f"MonthlyGraphWorker (Page Mode): '{sheet_name}' sayfası için OEE grafiği oluşturuluyor...")

                    try:
                        sheet_df = pd.read_excel(self.excel_path, sheet_name=sheet_name, header=0)
                        sheet_df.columns = sheet_df.columns.astype(str)  # Ensure columns are strings
                    except Exception as e:
                        logging.warning(f"'{sheet_name}' sayfası yüklenirken hata oluştu: {e}. Atlanıyor.")
                        self.progress.emit(int((i + 1) / total_items * 100))
                        continue

                    # Tarih sütununu al (A sütunu)
                    tarih_col_name = sheet_df.columns[excel_col_to_index('A')] if excel_col_to_index('A') < len(
                        sheet_df.columns) else None
                    if not tarih_col_name or tarih_col_name not in sheet_df.columns:
                        logging.warning(
                            f"MonthlyGraphWorker (Page Mode): '{sheet_name}' sayfasında 'A' sütunu (Tarih) bulunamadı. Atlanıyor.")
                        self.progress.emit(int((i + 1) / total_items * 100))
                        continue

                    # OEE sütununu al
                    current_oee_col_index = excel_col_to_index(oee_col_letter)
                    current_oee_col_name = sheet_df.columns[current_oee_col_index] if current_oee_col_index < len(
                        sheet_df.columns) else None

                    if not current_oee_col_name or current_oee_col_name not in sheet_df.columns:
                        logging.warning(
                            f"MonthlyGraphWorker (Page Mode): '{sheet_name}' sayfası için '{oee_col_letter}' ({current_oee_col_name}) sütunu bulunamadı veya geçersiz. Atlanıyor.")
                        self.progress.emit(int((i + 1) / total_items * 100))
                        continue

                    sheet_df['Tarih'] = pd.to_datetime(sheet_df[tarih_col_name], errors='coerce')
                    sheet_df.dropna(subset=['Tarih'], inplace=True)

                    if sheet_df.empty:
                        logging.warning(
                            f"MonthlyGraphWorker (Page Mode): '{sheet_name}' sayfası için tarih verisi bulunamadı. Atlanıyor.")
                        self.progress.emit(int((i + 1) / total_items * 100))
                        continue

                    # OEE sütununu sayısal değere dönüştür
                    sheet_df['OEE_Degeri_Processed'] = pd.to_numeric(
                        sheet_df[current_oee_col_name].astype(str).replace('%', '').replace(',', '.'),
                        errors='coerce'
                    )
                    sheet_df['OEE_Degeri_Processed'] = sheet_df['OEE_Degeri_Processed'].fillna(0.0)
                    if not sheet_df['OEE_Degeri_Processed'].empty and sheet_df['OEE_Degeri_Processed'].max() > 1.0:
                        sheet_df['OEE_Degeri_Processed'] = sheet_df['OEE_Degeri_Processed'] / 100.0

                    # Tarihe göre grupla ve OEE ortalamasını hesapla
                    grouped_oee = sheet_df.groupby(pd.Grouper(key='Tarih', freq='D'))[
                        'OEE_Degeri_Processed'].mean().reset_index()
                    grouped_oee.dropna(subset=['OEE_Degeri_Processed'], inplace=True)

                    if grouped_oee.empty:
                        logging.warning(
                            f"MonthlyGraphWorker (Page Mode): '{sheet_name}' sayfası için işlenecek OEE verisi bulunamadı. Atlanıyor.")
                        self.progress.emit(int((i + 1) / total_items * 100))
                        continue

                    figures_data.append((sheet_name, grouped_oee.to_dict('records')))
                    self.progress.emit(int((i + 1) / total_items * 100))

            self.finished.emit(figures_data, self.prev_year_oee, self.prev_month_oee)
        except Exception as exc:
            logging.exception("MonthlyGraphWorker hatası oluştu.")
            self.error.emit(f"Aylık grafik oluşturulurken bir hata oluştu: {str(exc)}")


class MonthlyGraphsPage(QWidget):
    """Aylık grafikler ve veri seçim sayfasını temsil eder."""

    def __init__(self, main_window: "MainWindow") -> None:
        super().__init__()
        self.main_window = main_window

        self.monthly_chart_container = QFrame(objectName="chartContainer")
        self.monthly_chart_layout = QVBoxLayout(self.monthly_chart_container)
        self.monthly_chart_layout.setAlignment(Qt.AlignCenter)

        self.current_monthly_chart_figure = None
        self.figures_data_monthly: List[Tuple[str, Union[List[dict[str, Any]], Dict[str, Any]]]] = []
        self.current_page_monthly = 0
        self.monthly_worker: MonthlyGraphWorker | None = None
        self.prev_year_oee_for_plot: float | None = None
        self.prev_month_oee_for_plot: float | None = None
        self.current_graph_mode: str = "hat"  # Default graph mode

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
        self.txt_prev_year_oee.setValidator(QtGui.QDoubleValidator(0.0, 100.0, 2))
        prev_year_oee_layout.addWidget(self.txt_prev_year_oee)
        oee_options_layout.addLayout(prev_year_oee_layout)

        prev_month_oee_layout = QHBoxLayout()
        prev_month_oee_layout.addWidget(QLabel("Önceki Ayın OEE Değeri (%):"))
        self.txt_prev_month_oee = QLineEdit()
        self.txt_prev_month_oee.setPlaceholderText("Örn: 82.0")
        self.txt_prev_month_oee.setValidator(QtGui.QDoubleValidator(0.0, 100.0, 2))
        prev_month_oee_layout.addWidget(self.txt_prev_month_oee)
        oee_options_layout.addLayout(prev_month_oee_layout)

        oee_buttons_layout = QHBoxLayout()
        self.btn_line_chart = QPushButton("Hat Grafikleri")
        # _start_monthly_graph_worker'a graph_mode parametresi eklendi
        self.btn_line_chart.clicked.connect(lambda: self._start_monthly_graph_worker(graph_mode="hat"))
        self.btn_line_chart.setEnabled(False)
        oee_buttons_layout.addWidget(self.btn_line_chart)

        self.btn_page_chart = QPushButton("Sayfa Grafikleri")
        # _start_monthly_graph_worker'a graph_mode parametresi eklendi
        self.btn_page_chart.clicked.connect(lambda: self._start_monthly_graph_worker(graph_mode="page"))
        self.btn_page_chart.setEnabled(False)
        oee_buttons_layout.addWidget(self.btn_page_chart)
        oee_options_layout.addLayout(oee_buttons_layout)

        main_layout.addWidget(self.oee_options_widget)

        self.other_graphs_widget = QWidget()
        other_graphs_layout = QVBoxLayout(self.other_graphs_widget)
        main_layout.addWidget(self.other_graphs_widget)
        self.other_graphs_widget.hide()

        self.monthly_chart_scroll_area = QScrollArea()
        self.monthly_chart_scroll_area.setWidgetResizable(True)
        self.monthly_chart_scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.monthly_chart_scroll_area.setWidget(self.monthly_chart_container)
        main_layout.addWidget(self.monthly_chart_scroll_area)

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
        """Grafiği temizler ve bu sayfaya girildiğinde düğme durumlarını günceller."""
        self.clear_monthly_chart_canvas()
        self.btn_save_monthly_chart.setEnabled(False)
        self.update_monthly_page_label(graph_mode=self.current_graph_mode)
        self.update_monthly_navigation_buttons(graph_mode=self.current_graph_mode)
        selected_type = self.cmb_monthly_graph_type.currentText()
        if selected_type == "OEE Grafikleri":
            self.btn_line_chart.setEnabled(True)
            self.btn_page_chart.setEnabled(True)
        else:
            self.btn_line_chart.setEnabled(False)
            self.btn_page_chart.setEnabled(False)

    def on_monthly_graph_type_changed(self, index: int):
        """Aylık grafik türü seçimi değiştiğinde ilgili seçenekleri gösterir/gizler."""
        selected_type = self.cmb_monthly_graph_type.currentText()
        self.clear_monthly_chart_canvas()
        self.btn_save_monthly_chart.setEnabled(False)
        self.figures_data_monthly.clear()
        self.current_page_monthly = 0

        # Reset current_graph_mode based on the selected type
        if selected_type == "OEE Grafikleri":
            self.current_graph_mode = "hat"  # Default to hat when OEE is selected
            self.oee_options_widget.show()
            self.other_graphs_widget.hide()
            self.btn_line_chart.setEnabled(True)
            self.btn_page_chart.setEnabled(True)
        elif selected_type in ["Dizgi Onay Dağılım Grafiği", "Dizgi Duruş Grafiği"]:
            self.current_graph_mode = "hat"  # These types are always hat mode
            self.oee_options_widget.hide()
            self.other_graphs_widget.show()
            self.btn_line_chart.setEnabled(False)
            self.btn_page_chart.setEnabled(False)
            self._start_monthly_graph_worker(graph_mode="hat")  # Auto-start for these types
        else:
            self.current_graph_mode = "hat"  # Fallback
            self.oee_options_widget.hide()
            self.other_graphs_widget.show()
            self.btn_line_chart.setEnabled(False)
            self.btn_page_chart.setEnabled(False)

        self.update_monthly_page_label(graph_mode=self.current_graph_mode)
        self.update_monthly_navigation_buttons(graph_mode=self.current_graph_mode)

    def clear_monthly_chart_canvas(self):
        """Aylık grafik tuvallerini temizler."""
        while self.monthly_chart_layout.count():
            item = self.monthly_chart_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def _start_monthly_graph_worker(self, graph_mode: str):
        """Aylık grafik çalışanını başlatır."""
        self.clear_monthly_chart_canvas()
        self.btn_save_monthly_chart.setEnabled(False)
        self.figures_data_monthly.clear()
        self.current_page_monthly = 0
        self.monthly_progress.setValue(0)
        self.monthly_progress.show()
        self.current_graph_mode = graph_mode  # Update current mode
        self.update_monthly_page_label(graph_mode=self.current_graph_mode)
        self.update_monthly_navigation_buttons(graph_mode=self.current_graph_mode)

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
            excel_path=self.main_window.excel_path,
            current_df=self.main_window.df,
            graph_mode=self.current_graph_mode,  # Use the updated mode
            graph_type=self.cmb_monthly_graph_type.currentText(),
            prev_year_oee=prev_year_oee,
            prev_month_oee=prev_month_oee,
            main_window=self.main_window
        )
        self.monthly_worker.finished.connect(self._on_monthly_graphs_generated)
        self.monthly_worker.progress.connect(self.monthly_progress.setValue)
        self.monthly_worker.error.connect(self._on_monthly_graph_error)
        self.monthly_worker.start()

    def _on_monthly_graphs_generated(self,
                                     figures_data_raw: List[Tuple[str, Union[List[dict[str, Any]], Dict[str, Any]]]],
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
        """MonthlyGraphWorker'dan gelen bir hata mesajını görüntüler."""
        QMessageBox.critical(self, "Hata", message)
        self.monthly_progress.setValue(0)
        self.monthly_progress.hide()
        self.btn_save_monthly_chart.setEnabled(False)

    def display_current_page_graphs_monthly(self) -> None:
        """Mevcut sayfadaki aylık grafiği görüntüler."""
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
            self.update_monthly_page_label(graph_mode=self.current_graph_mode)
            self.update_monthly_navigation_buttons(graph_mode=self.current_graph_mode)
            return

        name, data_container = self.figures_data_monthly[self.current_page_monthly]

        fig_width_inches = 17.0
        fig_height_inches = 8.0

        fig, ax = plt.subplots(figsize=(fig_width_inches, fig_height_inches), dpi=100)
        # Grafik arka plan rengini beyaza yap
        background_color = 'white'
        fig.patch.set_facecolor(background_color)
        ax.set_facecolor(background_color)

        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_linewidth(1.5)
        ax.spines['bottom'].set_linewidth(1.5)
        ax.grid(False)

        if self.cmb_monthly_graph_type.currentText() == "OEE Grafikleri":
            grouped_oee = pd.DataFrame(data_container)
            grouped_oee['Tarih'] = pd.to_datetime(grouped_oee['Tarih'])

            dates = grouped_oee['Tarih']
            oee_values = grouped_oee['OEE_Degeri_Processed'] if 'OEE_Degeri_Processed' in grouped_oee.columns else \
            grouped_oee['OEE_Degeri']

            line_color = '#1f77b4'

            x_indices = np.arange(len(dates))
            ax.plot(x_indices, oee_values, marker='o', markersize=8, color=line_color, linewidth=2, label=name)
            ax.plot(x_indices, oee_values, 'o', markersize=6, color='white', markeredgecolor=line_color,
                    markeredgewidth=1.5, zorder=5)

            for i, (x, y) in enumerate(zip(x_indices, oee_values)):
                if pd.notna(y) and y > 0:
                    ax.annotate(f'{y * 100:.1f}%', (x, y), textcoords="offset points", xytext=(0, 10), ha='center',
                                fontsize=8, fontweight='bold')

            overall_calculated_average = np.mean(oee_values) if not oee_values.empty else 0

            if self.prev_year_oee_for_plot is not None:
                y_val = self.prev_year_oee_for_plot / 100
                ax.axhline(y_val, color='red', linestyle='--', linewidth=1.5,
                           label=f'Önceki Yıl OEE ({self.prev_year_oee_for_plot:.1f}%)')
                ax.text(1.01, y_val, f'{self.prev_year_oee_for_plot:.1f}%',
                        transform=ax.transAxes, color='red', va='center', ha='left', fontsize=9, fontweight='bold')

            if self.prev_month_oee_for_plot is not None:
                y_val = self.prev_month_oee_for_plot / 100
                ax.axhline(y_val, color='orange', linestyle='--', linewidth=1.5,
                           label=f'Önceki Ay OEE ({self.prev_month_oee_for_plot:.1f}%)')
                ax.text(1.01, y_val, f'{self.prev_month_oee_for_plot:.1f}%',
                        transform=ax.transAxes, color='orange', va='center', ha='left', fontsize=9, fontweight='bold')

            if overall_calculated_average > 0:
                y_val = overall_calculated_average
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
                ax.text(1.01, y_val, f'{overall_calculated_average * 100:.1f}%',
                        transform=ax.transAxes, color='purple', va='center', ha='left', fontsize=9, fontweight='bold')

            ax.set_xticks(x_indices)
            ax.set_xticklabels([d.strftime('%d.%m.%Y') for d in dates])

            fig.autofmt_xdate(rotation=45)

            ax.yaxis.set_major_formatter(PercentFormatter(xmax=1, decimals=0))

            ax.set_yticks(np.arange(0.0, 1.001, 0.25))
            ax.set_ylim(bottom=-0.05, top=1.05)

            ax.set_xlabel("Tarih", fontsize=12, fontweight='bold')
            ax.set_ylabel("OEE (%)", fontsize=12, fontweight='bold')

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

            # Grafik başlığı dinamik olarak ayarlandı
            if self.current_graph_mode == "page":
                chart_title = f"{month_name} {name} OEE"  # name burada sayfa adıdır
            else:  # hat mode
                chart_title = f"{name} {month_name} OEE"  # name burada hat adıdır

            ax.set_title(chart_title, fontsize=16, color='#2c3e50', fontweight='bold')

            ax.legend(loc='upper left', bbox_to_anchor=(1.02, 0), fontsize=10)
            fig.subplots_adjust(right=0.60)

        elif self.cmb_monthly_graph_type.currentText() == "Dizgi Onay Dağılım Grafiği":
            labels = [d["label"] for d in data_container]
            values = [d["value"] for d in data_container]

            colors = ['#00008B', '#ff7f0e']

            total_sum = sum(values)

            def func(pct, allvals):
                absolute = int(np.round(pct / 100. * total_sum))
                hours = absolute // 3600
                minutes = (absolute % 3600) // 60
                seconds = absolute % 60
                return f"{hours:02d}:{minutes:02d}:{minutes:02d}; {pct:.0f}%"

            wedges, texts, autotexts = ax.pie(
                values,
                autopct=lambda pct: func(pct, values),
                startangle=90,
                colors=colors,
                wedgeprops=dict(edgecolor='black', linewidth=1.5)
            )

            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontsize(14)
                autotext.set_fontweight('bold')

            ax.axis('equal')

            ax.legend(wedges, labels,
                      title="Hatlar",
                      loc="upper right",
                      bbox_to_anchor=(1.2, 1),
                      fontsize=10,
                      title_fontsize=12)

            chart_title = f"Dizgi Onay Dağılımı"
            ax.set_title(chart_title, fontsize=16, color='#2c3e50', fontweight='bold')
            fig.tight_layout()

        elif self.cmb_monthly_graph_type.currentText() == "Dizgi Duruş Grafiği":
            # Extract metrics and total_overall_sum from the data_container
            metric_sums_dict = data_container["metrics"]
            total_overall_sum = data_container["total_overall_sum"]
            cumulative_percentages_dict = data_container["cumulative_percentages"]

            metric_sums = pd.Series(metric_sums_dict)
            cumulative_percentage = pd.Series(cumulative_percentages_dict)

            ax2 = ax.twinx()

            # Grafik renklendirmesi (görseldeki gibi)
            bar_color = '#AECDCB'  # Açık teal/grimsi yeşil
            line_color = '#6B0000'  # Koyu kırmızımsı kahverengi

            # Sütunları çiz
            bars = ax.bar(metric_sums.index, metric_sums.values / 60, color=bar_color, alpha=0.8, edgecolor='black',
                          linewidth=1.5)

            # Kümülatif yüzde çizgisini çiz (marker olmadan, daha düşük zorder)
            ax2.plot(metric_sums.index, cumulative_percentage, color=line_color, linestyle='-', linewidth=2, zorder=1)
            # Kümülatif yüzde noktalarını çiz (içi boş daireler, daha yüksek zorder)
            ax2.plot(metric_sums.index, cumulative_percentage, 'o', markersize=8, markerfacecolor='white',
                     markeredgecolor=line_color, markeredgewidth=2, zorder=2)

            # Get the x-axis limits after plotting to determine the range for axhline
            x_min_data, x_max_data = ax.get_xlim()

            # %80 noktasında yatay kesikli kırmızı çizgi ekle (görseldeki gibi gri ve ince)
            # Calculate normalized xmin and xmax based on the actual data range
            # The cumulative line is plotted over the indices 0 to len(metric_sums.index) - 1
            # Convert these data coordinates to normalized axis coordinates.
            normalized_xmin = (0 - x_min_data) / (x_max_data - x_min_data)
            normalized_xmax = (len(metric_sums.index) - 1 - x_min_data) / (x_max_data - x_min_data)

            ax2.axhline(80, color='#B0B0B0', linestyle='--', linewidth=1.5, xmin=normalized_xmin, xmax=normalized_xmax)

            # Yatay ızgara çizgilerini kaldır
            ax.grid(False)
            ax2.grid(False)

            # "Duruş Nedenleri" başlığını kaldır
            ax.set_xlabel("")

            # Metrik isimlerini kalın fontta yaz
            # Set explicit ticks for x-axis to avoid UserWarning
            ax.set_xticks(np.arange(len(metric_sums.index)))
            ax.set_xticklabels(metric_sums.index, fontsize=10, fontweight='bold', rotation=45, ha='right')

            # Sütunların üzerindeki % ve süre değerleri
            for i, bar in enumerate(bars):
                value_seconds = metric_sums.values[i]
                percentage = (value_seconds / total_overall_sum) * 100 if total_overall_sum > 0 else 0
                duration_hours = int(value_seconds // 3600)
                duration_minutes = int((value_seconds % 3600) // 60)
                duration_seconds = int(value_seconds % 60)

                text_label = f"{duration_hours:02d}:{duration_minutes:02d}:{duration_seconds:02d}\n({percentage:.1f}%)"
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.5, text_label,
                        ha='center', va='bottom', fontsize=10, fontweight='bold', color='black')

            ax.set_ylabel("Süre (Dakika)", fontsize=12, fontweight='bold',
                          color=bar_color)
            ax2.set_ylabel("Kümülatif Yüzde (%)", fontsize=12, fontweight='bold', color=line_color)

            ax.tick_params(axis='x', rotation=45)

            # Pareto grafiği için dinamik başlık
            chart_title = "Genel Dizgi Duruş Pareto Analizi"  # Varsayılan başlık

            if not self.main_window.df.empty and 'Tarih' in self.main_window.df.columns:
                df_dates = pd.to_datetime(self.main_window.df['Tarih'], errors='coerce').dropna()

                if not df_dates.empty:
                    min_date = df_dates.min()
                    max_date = df_dates.max()

                    month_names_turkish = {
                        1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
                        7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
                    }

                    if min_date.month == max_date.month and min_date.year == max_date.year:
                        month_name = month_names_turkish.get(min_date.month, min_date.strftime('%B')).capitalize()
                        chart_title = f"{month_name} Ayı Dizgi Duruşları"
                    elif min_date.year == max_date.year:
                        first_month_name = month_names_turkish.get(min_date.month, min_date.strftime('%B')).capitalize()
                        last_month_name = month_names_turkish.get(max_date.month, max_date.strftime('%B')).capitalize()
                        chart_title = f"{min_date.year} Yılı {first_month_name}-{last_month_name} Ayları Dizgi Duruşları"
                    else:
                        chart_title = f"{min_date.year}-{max_date.year} Yılları Dizgi Duruşları"

            # Başlık rengini görseldeki gibi koyu gri yap
            ax.set_title(chart_title, fontsize=20, color='#363636', fontweight='bold')  # Başlık boyutu arttırıldı

            ax.set_ylim(bottom=0)
            ax2.set_ylim(0, 100)

            ax.spines['top'].set_visible(False)
            ax2.spines['top'].set_visible(False)

            fig.subplots_adjust(right=0.75)

            fig.tight_layout()

        canvas = FigureCanvas(fig)
        canvas.setFixedSize(int(fig_width_inches * fig.dpi), int(fig_height_inches * fig.dpi))
        self.monthly_chart_layout.addWidget(canvas, stretch=1)
        canvas.draw()

        self.current_monthly_chart_figure = fig
        self.btn_save_monthly_chart.setEnabled(True)
        self.update_monthly_page_label(graph_mode=self.current_graph_mode)
        self.update_monthly_navigation_buttons(graph_mode=self.current_graph_mode)

    def update_monthly_page_label(self, graph_mode: str) -> None:
        """Aylık grafik sayfa etiketini günceller."""
        total_pages = len(self.figures_data_monthly)
        self.lbl_monthly_page.setText(f"Sayfa {self.current_page_monthly + 1} / {total_pages}")

    def update_monthly_navigation_buttons(self, graph_mode: str) -> None:
        """Aylık grafik gezinme düğmelerinin etkin durumunu günceller."""
        total_pages = len(self.figures_data_monthly)
        self.btn_prev_monthly.setEnabled(self.current_page_monthly > 0)
        self.btn_next_monthly.setEnabled(self.current_page_monthly < total_pages - 1)
        if graph_mode == "hat":
            self.btn_prev_monthly.setText("← Önceki Hat")
            self.btn_next_monthly.setText("Sonraki Hat →")
        elif graph_mode == "page":
            self.btn_prev_monthly.setText("← Önceki Sayfa")
            self.btn_next_monthly.setText("Sonraki Sayfa →")

    def prev_monthly_page(self) -> None:
        """Önceki aylık grafik sayfasına gider."""
        if self.current_page_monthly > 0:
            self.current_page_monthly -= 1
            self.display_current_page_graphs_monthly()

    def next_monthly_page(self) -> None:
        """Sonraki aylık grafik sayfasına gider."""
        total_pages = len(self.figures_data_monthly)
        if self.current_page_monthly < total_pages - 1:
            self.current_page_monthly += 1
            self.display_current_page_graphs_monthly()

    def _save_monthly_chart_as_image(self):
        """Aylık grafiği PNG/JPEG olarak kaydeder."""
        if self.current_monthly_chart_figure is None:
            QMessageBox.warning(self, "Kaydedilecek Grafik Yok", "Görüntülenecek bir aylık grafik bulunmamaktadır.")
            return

        current_name = "grafik"
        if self.figures_data_monthly and 0 <= self.current_page_monthly < len(self.figures_data_monthly):
            current_name = self.figures_data_monthly[self.current_page_monthly][0].replace(" ", "_").replace("/",
                                                                                                             "-")

        graph_type_name = self.cmb_monthly_graph_type.currentText().replace(" ", "_").replace("/", "-")
        default_filename = f"{graph_type_name}_{current_name}.png"

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
        self.selected_grouping_val: str = ""

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
        """Uygulamaya modern bir stil sayfası uygular."""
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
        """Belirli bir sayfaya gider ve onu yeniler."""
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
            logging.warning("load_excel: Excel yolu veya seçilen sayfa boş. Veri yüklenemiyor.")
            return

        # Eğer aynı dosya ve sayfa zaten yüklüyse tekrar yükleme
        if not self.df.empty and self.df.attrs.get('excel_path') == self.excel_path and \
                self.df.attrs.get('selected_sheet') == self.selected_sheet:
            logging.info(f"'{self.selected_sheet}' sayfasındaki veriler zaten yüklü. Yeniden yüklenmiyor.")
            return

        try:
            self.df = pd.read_excel(self.excel_path, sheet_name=self.selected_sheet, header=0)
            self.df.columns = self.df.columns.astype(str)

            self.df.attrs['excel_path'] = self.excel_path
            self.df.attrs['selected_sheet'] = self.selected_sheet

            logging.info("Veri '%s' sayfasından yüklendi. Satır sayısı: %d", self.selected_sheet, len(self.df))

            # Sütun isimlerini dinamik olarak al
            # A sütunu gruplama (tarih), B sütunu gruplanan (ürün)
            self.grouping_col_name = self.df.columns[excel_col_to_index('A')]
            self.grouped_col_name = self.df.columns[excel_col_to_index('B')]
            self.oee_col_name = None
            self.metric_cols = []

            # Seçilen sayfaya göre OEE ve metrik sütunlarını belirle
            if self.selected_sheet == "SMD-OEE":
                self.oee_col_name = self.df.columns[excel_col_to_index('BP')] if excel_col_to_index('BP') < len(
                    self.df.columns) else None
                start_col_index = excel_col_to_index('H')
                end_col_index = excel_col_to_index('BD')
                # AP sütunu metriklerden hariç tutulacak
                ap_col_index = excel_col_to_index('AP')
                for i in range(start_col_index, end_col_index + 1):
                    if i < len(self.df.columns) and i != ap_col_index:
                        self.metric_cols.append(self.df.columns[i])
            elif self.selected_sheet == "ROBOT":
                # ROBOT sayfası için OEE sütunu BG olarak belirtildi, ancak mevcut kodda kullanılmıyor.
                # Eğer ROBOT sayfası için de OEE grafiği çizilecekse bu kısım güncellenmeli.
                # Günlük grafiklerde OEE sütunu kullanılmadığı için burada sadece metrikler tanımlanır.
                start_col_index = excel_col_to_index('H')
                end_col_index = excel_col_to_index('AU')
                # AO sütunu metriklerden hariç tutulacak
                ao_col_index = excel_col_to_index('AO')

                for i in range(start_col_index, end_col_index + 1):
                    if i < len(self.df.columns) and i != ao_col_index:
                        self.metric_cols.append(self.df.columns[i])
            elif self.selected_sheet == "DALGA_LEHİM":
                self.oee_col_name = self.df.columns[excel_col_to_index('BP')] if excel_col_to_index('BP') < len(
                    self.df.columns) else None
                start_col_index = excel_col_to_index('H')
                end_col_index = excel_col_to_index('BD')
                # AP sütunu metriklerden hariç tutulacak
                ap_col_index = excel_col_to_index('AP')
                for i in range(start_col_index, end_col_index + 1):
                    if i < len(self.df.columns) and i != ap_col_index:
                        self.metric_cols.append(self.df.columns[i])
            elif self.selected_sheet == "KAPLAMA-OEE":
                self.oee_col_name = self.df.columns[excel_col_to_index('BG')] if excel_col_to_index('BG') < len(
                    self.df.columns) else None
                # KAPLAMA-OEE için özel metrik sütunları tanımlanmadıysa, boş bırakılır veya varsayılan atanır.
                # Bu sayfa için sadece OEE grafiği istendiği için metrikler boş kalabilir.
                self.metric_cols = []

            logging.info("Gruplama sütunu tanımlandı: %s", self.grouping_col_name)
            logging.info("Gruplanan sütun tanımlandı: %s", self.grouped_col_name)
            logging.info("OEE sütunu tanımlandı: %s", self.oee_col_name)
            logging.info("Metrik sütunları tanımlandı: %s", self.metric_cols)

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
        logging.exception("Uygulama başlatılırken kritik bir hata oluştu.")
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText("Uygulama başlatılırken kritik bir hata oluştu.")
        msg.setInformativeText(str(e))
        msg.setWindowTitle("Kritik Hata")
        msg.exec_()
        sys.exit(1)
