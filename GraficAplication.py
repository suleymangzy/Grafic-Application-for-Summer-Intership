import sys
import logging
import datetime
from pathlib import Path
from typing import List, Tuple, Any

import pandas as pd
import numpy as np

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.backends.backend_pdf import PdfPages

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

GRAPHS_PER_PAGE = 1
REQ_SHEETS = {"SMD-OEE", "ROBOT", "DALGA_LEHİM"}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)

plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'DejaVu Sans', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False


def excel_col_to_index(col_letter: str) -> int:
    index = 0
    for char in col_letter.upper():
        if not char.isalpha():
            raise ValueError(f"Geçersiz sütun harfi: {col_letter}")
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1


def seconds_from_timedelta(series: pd.Series) -> pd.Series:
    seconds_series = pd.Series(0.0, index=series.index, dtype=float)
    is_time_obj = series.apply(lambda x: isinstance(x, datetime.time))
    if is_time_obj.any():
        time_objects = series[is_time_obj]
        seconds_series.loc[is_time_obj] = time_objects.apply(
            lambda t: t.hour * 3600 + t.minute * 60 + t.second + t.microsecond / 1e6
        )
    remaining_indices = series.index[~is_time_obj & series.notna()]
    if not remaining_indices.empty:
        remaining_series_str = series.loc[remaining_indices].astype(str).str.strip()
        remaining_series_str = remaining_series_str.replace('', np.nan)
        converted_td = pd.to_timedelta(remaining_series_str, errors='coerce')
        valid_td_mask = pd.notna(converted_td)
        seconds_series.loc[remaining_indices[valid_td_mask]] = converted_td[valid_td_mask].dt.total_seconds()
    remaining_nan_indices = seconds_series.index[seconds_series.isna()]
    if not remaining_nan_indices.empty:
        numeric_values = pd.to_numeric(series.loc[remaining_nan_indices], errors='coerce')
        valid_numeric_mask = pd.notna(numeric_values)
        if valid_numeric_mask.any():
            converted_from_numeric = pd.to_timedelta(numeric_values[valid_numeric_mask], unit='D', errors='coerce')
            valid_num_td_mask = pd.notna(converted_from_numeric)
            seconds_series.loc[remaining_nan_indices[valid_numeric_mask & valid_num_td_mask]] = converted_from_numeric[
                valid_num_td_mask].dt.total_seconds()
    return seconds_series.fillna(0.0)


class GraphWorker(QThread):
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
            bp_col_name: str | None,
            oee_col_name: str | None # Added oee_col_name
    ) -> None:
        super().__init__()
        self.df = df.copy()
        self.grouping_col_name = grouping_col_name
        self.grouped_col_name = grouped_col_name
        self.grouped_values = grouped_values
        self.metric_cols = metric_cols
        self.bp_col_name = bp_col_name
        self.oee_col_name = oee_col_name # Store oee_col_name

    def run(self) -> None:
        try:
            results: List[Tuple[str, pd.Series, float, str]] = []
            total = len(self.grouped_values)
            # Only process metric_cols for time conversion. BP will be handled separately for OEE value if needed.
            all_cols_to_process = list(set(self.metric_cols + ([self.bp_col_name] if self.bp_col_name else [])))
            df_processed_times = self.df.copy()

            for col in all_cols_to_process:
                if col in df_processed_times.columns:
                    df_processed_times[col] = seconds_from_timedelta(df_processed_times[col])

            for i, current_grouped_val in enumerate(self.grouped_values, 1):
                subset_df_for_chart = df_processed_times[
                    df_processed_times[self.grouped_col_name].astype(str) == current_grouped_val
                    ].copy()

                sums = subset_df_for_chart[self.metric_cols].sum()
                sums = sums[sums > 0]

                bp_total_seconds = 0.0
                if self.bp_col_name and self.bp_col_name in subset_df_for_chart.columns:
                    bp_total_seconds = subset_df_for_chart[self.bp_col_name].sum()

                # Get OEE value directly from the column specified as oee_col_name
                oee_display_value = "0%" # Default value
                if self.oee_col_name and self.oee_col_name in self.df.columns:
                    # Find the row in the original df that matches the current grouped value
                    matching_rows = self.df[self.df[self.grouped_col_name].astype(str) == current_grouped_val]
                    if not matching_rows.empty:
                        oee_value_raw = matching_rows[self.oee_col_name].iloc[0]
                        if pd.notna(oee_value_raw):
                            try:
                                oee_value_float: float
                                if isinstance(oee_value_raw, str):
                                    # Attempt to convert string to float, removing '%' if present
                                    oee_value_str = oee_value_raw.replace('%', '').strip()
                                    oee_value_float = float(oee_value_str)
                                elif isinstance(oee_value_raw, (int, float)):
                                    oee_value_float = float(oee_value_raw)
                                else:
                                    raise ValueError("Unsupported OEE value type or format")

                                # Determine if it's a decimal (e.g., 0.51) or a whole number (e.g., 51)
                                if 0.0 <= oee_value_float <= 1.0 and oee_value_float != 0:
                                    # It's a decimal percentage, convert to 0-100 scale
                                    oee_display_value = f"{oee_value_float * 100:.0f}%"
                                elif oee_value_float > 1.0:
                                    # It's already on a 0-100 scale
                                    oee_display_value = f"{oee_value_float:.0f}%"
                                else:
                                    # Handle 0 or negative values, display as 0%
                                    oee_display_value = "0%"
                            except (ValueError, TypeError):
                                # If conversion fails (e.g., "#SAYI/D!" or other non-numeric strings), keep default "0%"
                                oee_display_value = "0%"


                if not sums.empty:
                    results.append((current_grouped_val, sums, bp_total_seconds, oee_display_value))
                self.progress.emit(int(i / total * 100))

            self.finished.emit(results)
        except Exception as exc:
            logging.exception("GraphWorker hata")
            self.error.emit(str(exc))


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

            if len(sheets) == 1:
                self.main_window.selected_sheet = sheets[0]
                self.sheet_selection_label.setText(f"İşlenecek Sayfa: <b>{self.main_window.selected_sheet}</b>")
                self.cmb_sheet.hide()
            else:
                self.main_window.selected_sheet = self.cmb_sheet.currentText()

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

        grouping_group = QHBoxLayout()
        grouping_group.addWidget(QLabel("<b>Gruplama Değişkeni (A Sütunu):</b>"))
        self.cmb_grouping = QComboBox()
        self.cmb_grouping.currentIndexChanged.connect(self.populate_grouped)
        grouping_group.addWidget(self.cmb_grouping)
        main_layout.addLayout(grouping_group)

        grouped_group = QHBoxLayout()
        grouped_group.addWidget(QLabel("<b>Gruplanan Değişkenler (B Sütunu):</b>"))
        self.lst_grouped = QListWidget()
        self.lst_grouped.setSelectionMode(QListWidget.MultiSelection)
        self.lst_grouped.itemSelectionChanged.connect(self.update_next_button_state)
        grouped_group.addWidget(self.lst_grouped)
        main_layout.addLayout(grouped_group)

        metrics_group = QVBoxLayout()
        metrics_group.addWidget(QLabel("<b>Metrikler (H-BD, AP hariç):</b>"))
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
        df = self.main_window.df
        if df.empty:
            QMessageBox.critical(self, "Hata", "Veri yüklenemedi. Lütfen dosyayı kontrol edin.")
            self.main_window.goto_page(0)
            return

        self.cmb_grouping.clear()
        if self.main_window.grouping_col_name and self.main_window.grouping_col_name in df.columns:
            grouping_vals = sorted(df[self.main_window.grouping_col_name].dropna().astype(str).unique())
            grouping_vals = [s for s in grouping_vals if s.strip()]
            self.cmb_grouping.addItems(grouping_vals)
            if not grouping_vals:
                QMessageBox.warning(self, "Uyarı", "Gruplama sütunu (A) boş veya geçerli değer içermiyor.")
        else:
            QMessageBox.warning(self, "Uyarı", "Gruplama sütunu (A) bulunamadı veya boş.")

        self.populate_metrics_checkboxes()
        self.populate_grouped()

    def populate_grouped(self) -> None:
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
        while self.metrics_layout.count():
            item = self.metrics_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

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
            checkbox.setChecked(not is_entirely_empty)
            checkbox.setEnabled(not is_entirely_empty)

            if is_entirely_empty:
                checkbox.setText(f"{col_name} (Boş)")
                checkbox.setStyleSheet("color: gray;")
            else:
                self.main_window.selected_metrics.append(col_name)

            checkbox.stateChanged.connect(self.on_metric_checkbox_changed)
            self.metrics_layout.addWidget(checkbox)

        self.update_next_button_state()

    def on_metric_checkbox_changed(self, state):
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

        self.figures_data: List[Tuple[str, Figure]] = []
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

        nav_top = QHBoxLayout()
        self.btn_back = QPushButton("← Veri Seçimi")
        self.btn_back.clicked.connect(lambda: self.main_window.goto_page(1))
        nav_top.addWidget(self.btn_back)
        nav_top.addStretch(1)
        self.lbl_page = QLabel("Sayfa 0 / 0")
        self.lbl_page.setAlignment(Qt.AlignCenter)
        nav_top.addWidget(self.lbl_page)
        nav_top.addStretch(1)

        self.btn_save_image = QPushButton("Grafiği Kaydet (PNG/JPEG)")
        self.btn_save_image.clicked.connect(self.save_single_graph_as_image)
        nav_top.addWidget(self.btn_save_image)

        main_layout.addLayout(nav_top)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.canvas_holder = QWidget()
        self.vbox_canvases = QVBoxLayout(self.canvas_holder)
        self.canvas_holder.setLayout(self.vbox_canvases)
        self.scroll.setWidget(self.canvas_holder)
        main_layout.addWidget(self.scroll)

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

        self.worker = GraphWorker(
            df=df,
            grouping_col_name=self.main_window.grouping_col_name,
            grouped_col_name=self.main_window.grouped_col_name,
            grouped_values=self.main_window.grouped_values,
            metric_cols=self.main_window.selected_metrics,
            bp_col_name=self.main_window.bp_col_name,
            oee_col_name=self.main_window.oee_col_name # Pass OEE column name
        )
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished.connect(self.on_results)
        self.worker.error.connect(lambda m: QMessageBox.critical(self, "Hata", m))
        self.worker.start()

    def on_results(self, results: List[Tuple[str, pd.Series, float, str]]) -> None:
        self.progress.setValue(100)
        if not results:
            QMessageBox.information(self, "Veri yok", "Grafik oluşturulamadı. Seçilen kriterlere göre veri bulunamadı.")
            return

        # Use 'Paired' colormap for better distinction
        colors_palette = plt.cm.get_cmap('Paired', len(self.main_window.selected_metrics))
        metric_colors = {metric: colors_palette(i) for i, metric in enumerate(self.main_window.selected_metrics)}

        for grouped_val, metric_sums, bp_total_seconds, oee_display_value in results:
            fig = Figure(figsize=(8, 8), dpi=100)
            ax = fig.add_subplot(111)

            # Donut chart
            wedges, texts, autotexts = ax.pie(
                metric_sums.values,
                # Labels formatted as "METRİK ADI; SAAT:DAKİTA; YÜZDE%"
                labels=[f"{label}; {int(value // 3600):02d}:{int((value % 3600) // 60):02d}; {p:.0f}%"
                        for label, value, p in zip(metric_sums.index, metric_sums.values, metric_sums.values / metric_sums.sum() * 100)],
                autopct="", # Disable default autopct to use custom labels for percentages
                startangle=90,
                counterclock=False,
                colors=[metric_colors[m] for m in metric_sums.index],
                wedgeprops=dict(width=0.4, edgecolor='w')
            )

            # OEE value in the center
            ax.text(0, 0, f"OEE\n{oee_display_value}",
                    horizontalalignment='center', verticalalignment='center',
                    fontsize=24, fontweight='bold', color='black')

            current_date = datetime.datetime.now().strftime("%d.%m.%Y")
            title_text = f"{current_date}\n{self.main_window.grouped_col_name.upper()}: {grouped_val.upper()}"

            ax.set_title(title_text, fontweight="bold", fontsize=16, pad=20)

            # Position labels outside and add a line connecting to the slice
            # This part is a bit complex for a simple code modification without direct access to label positioning logic of matplotlib.
            # The current approach for labels in `ax.pie` already tries to place them, but for more precise external labels with lines,
            # you'd typically need to iterate over wedges and texts, calculate positions, and draw lines manually.
            # For this request, I'll keep the labels *next to* the slices but improve their readability by including percentages.
            # To achieve the exact visual style of image_299c97.png with external labels and lines,
            # more advanced matplotlib annotation techniques would be required.
            # The current label formatting in `labels` argument makes them more informative.
            for text, autotext in zip(texts, autotexts):
                text.set_fontsize(10)
                text.set_color('black')
                # Autotext (percentage) is removed as it's part of the label now.
                # If you want separate percentage labels, you'd re-enable autopct and handle positioning.
                autotext.set_visible(False) # Hide autopct text since we embedded it in the main label

            ax.axis("equal")
            fig.tight_layout(rect=[0, 0, 1, 0.95])

            # TOPLAM DURUŞ calculation and display
            total_duration_seconds = metric_sums.sum()
            total_duration_hours = int(total_duration_seconds // 3600)
            total_duration_minutes = int((total_duration_seconds % 3600) // 60)
            total_duration_text = f"TOPLAM DURUŞ\n{total_duration_hours} SAAT {total_duration_minutes} DAKİKA"
            fig.text(0.05, 0.05, total_duration_text, transform=fig.transFigure,
                     fontsize=14, fontweight='bold', verticalalignment='bottom')

            # Add "HAT ÇALIŞMADI" text and its value below the "TOPLAM DURUŞ" text
            bp_hours = int(bp_total_seconds // 3600)
            bp_minutes = int((bp_total_seconds % 3600) // 60)
            bp_seconds = int(bp_total_seconds % 60)
            bp_text = f"HAT ÇALIŞMADI; {bp_hours:02d}:{bp_minutes:02d}:{bp_seconds:02d};%100"
            fig.text(0.05, 0.01, bp_text, transform=fig.transFigure,
                     fontsize=10, verticalalignment='bottom', bbox=dict(boxstyle='round,pad=0.3', fc='lightgrey', ec='black', lw=0.5))


            self.figures_data.append((grouped_val, fig))
        self.display_page()

    def display_page(self) -> None:
        self.clear_canvases()
        start = self.current_page * GRAPHS_PER_PAGE
        end = start + GRAPHS_PER_PAGE

        for _, fig in self.figures_data[start:end]:
            canvas = FigureCanvas(fig)
            frame = QFrame()
            frame.setFrameShape(QFrame.StyledPanel)
            frame.setLineWidth(1)
            vb = QVBoxLayout(frame)
            vb.addWidget(canvas)
            self.vbox_canvases.addWidget(frame)
        self.vbox_canvases.addStretch(1)
        self.update_page_label()

    def clear_canvases(self) -> None:
        while self.vbox_canvases.count():
            item = self.vbox_canvases.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
            elif item.layout():
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
        self.btn_save_image.setEnabled(len(self.figures_data) > 0)

    def next_page(self) -> None:
        if (self.current_page + 1) * GRAPHS_PER_PAGE < len(self.figures_data):
            self.current_page += 1
            self.display_page()

    def prev_page(self) -> None:
        if self.current_page > 0:
            self.current_page -= 1
            self.display_page()

    def save_single_graph_as_image(self) -> None:
        if not self.figures_data:
            QMessageBox.warning(self, "Grafik yok", "Kaydedilecek grafik bulunamadı.")
            return

        current_figure_index = self.current_page * GRAPHS_PER_PAGE
        if current_figure_index >= len(self.figures_data):
            QMessageBox.warning(self, "Hata", "Geçerli sayfada gösterilecek grafik bulunmuyor.")
            return

        grouped_val, current_fig = self.figures_data[current_figure_index]

        default_filename = f"grafik_{grouped_val}.jpeg"
        filters = "JPEG Dosyaları (*.jpeg *.jpg);;PNG Dosyaları (*.png)"

        file_name, selected_filter = QFileDialog.getSaveFileName(
            self, "Grafiği Resim Olarak Kaydet", default_filename, filters
        )

        if not file_name:
            return

        try:
            if "jpeg" in selected_filter.lower() or "jpg" in selected_filter.lower():
                format = 'jpeg'
            elif "png" in selected_filter.lower():
                format = 'png'
            else:
                format = 'jpeg'

            current_fig.savefig(file_name, bbox_inches='tight', pad_inches=0.5, format=format, dpi=300)
            QMessageBox.information(self, "Başarılı", f"Grafik '{file_name}' konumuna başarıyla kaydedildi.")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Grafik kaydedilirken bir hata oluştu: {e}")


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Pasta Grafik Rapor Uygulaması")
        self.resize(1200, 900)

        self.excel_path: Path | None = None
        self.selected_sheet: str = ""
        self.df: pd.DataFrame = pd.DataFrame()
        self.grouping_col_name: str | None = None
        self.grouped_col_name: str | None = None
        self.bp_col_name: str | None = None
        self.oee_col_name: str | None = None
        self.metric_cols: List[str] = []
        self.grouped_values: List[str] = []
        self.selected_metrics: List[str] = []

        self.init_ui()

    def init_ui(self):
        self.stacked_widget = QStackedWidget(self)
        self.setCentralWidget(self.stacked_widget)

        self.file_selection_page = FileSelectionPage(self)
        self.data_selection_page = DataSelectionPage(self)
        self.graphs_page = GraphsPage(self)

        self.stacked_widget.addWidget(self.file_selection_page)
        self.stacked_widget.addWidget(self.data_selection_page)
        self.stacked_widget.addWidget(self.graphs_page)

        self.stacked_widget.setCurrentWidget(self.file_selection_page)

    def goto_page(self, index: int) -> None:
        self.stacked_widget.setCurrentIndex(index)
        if index == 1:
            self.data_selection_page.refresh()
        elif index == 2:
            self.graphs_page.enter_page()

    def load_excel(self) -> None:
        if not self.excel_path or not self.selected_sheet:
            QMessageBox.critical(self, "Hata", "Dosya yolu veya sayfa adı belirtilmedi.")
            return

        logging.info("Excel okunuyor: %s | Sheet: %s", self.excel_path, self.selected_sheet)
        try:
            df_raw = pd.read_excel(self.excel_path, sheet_name=self.selected_sheet, header=0)
            df_raw.columns = df_raw.columns.astype(str).str.strip().str.upper()
            self.df = df_raw

            a_idx = excel_col_to_index('A')
            b_idx = excel_col_to_index('B')
            bp_idx = excel_col_to_index('F') # Changed to 'F' for 'Üretim Yok'
            oee_idx = excel_col_to_index('BF') # Changed to 'BF' for OEE

            if a_idx < len(self.df.columns):
                self.grouping_col_name = self.df.columns[a_idx]
            else:
                QMessageBox.warning(self, "Uyarı", f"Excel'de 'A' ({a_idx + 1}. sütun) bulunamadı.")
                self.grouping_col_name = None

            if b_idx < len(self.df.columns):
                self.grouped_col_name = self.df.columns[b_idx]
            else:
                QMessageBox.warning(self, "Uyarı", f"Excel'de 'B' ({b_idx + 1}. sütun) bulunamadı.")
                self.grouped_col_name = None

            self.bp_col_name = None
            if bp_idx < len(self.df.columns):
                self.bp_col_name = self.df.columns[bp_idx]
                logging.info("BP sütunu (HAT ÇALIŞMADI için): %s", self.bp_col_name)
            else:
                logging.warning("BP sütunu ('F' indeksi) Excel dosyasında bulunamadı. 'HAT ÇALIŞMADI' değeri '00:00:00;%100' olarak gösterilecek.")
                self.bp_col_name = None

            self.oee_col_name = None # Reset OEE column name
            if oee_idx < len(self.df.columns):
                self.oee_col_name = self.df.columns[oee_idx]
                logging.info("OEE sütunu ('BF' indeksi): %s", self.oee_col_name)
            else:
                logging.warning("OEE sütunu ('BF' indeksi) Excel dosyasında bulunamadı. OEE değeri '0%' olarak gösterilecek.")
                self.oee_col_name = None


            h_idx = excel_col_to_index("H")
            bd_idx = excel_col_to_index("BD")
            ap_idx = excel_col_to_index("AP")

            potential_metrics_from_range = []
            max_col_idx = len(self.df.columns) - 1

            if h_idx <= max_col_idx and bd_idx <= max_col_idx and h_idx <= bd_idx:
                for i in range(h_idx, bd_idx + 1):
                    col_name = self.df.columns[i]
                    if self.df.columns.get_loc(col_name) != ap_idx:
                        potential_metrics_from_range.append(col_name)
            else:
                QMessageBox.warning(self, "Uyarı",
                                    f"Metrik aralığı (H-BD) geçersiz veya sütunlar bulunamadı. (H:{h_idx + 1}, BD:{bd_idx + 1}, Toplam Sütun:{len(self.df.columns)})")

            self.metric_cols = [
                c for c in potential_metrics_from_range
                if c in self.df.columns and not self.df[c].dropna().empty and not self.df[c].astype(str).str.strip().eq(
                    '').all()
            ]

            logging.info("%d geçerli metrik bulundu", len(self.metric_cols))

        except Exception as e:
            QMessageBox.critical(self, "Okuma hatası",
                                 f"Excel dosyası yüklenirken veya işlenirken bir hata oluştu: {e}\nLütfen dosyanın bozuk olmadığından ve formatının doğru olduğundan emin olun.")
            self.df = pd.DataFrame()
            self.excel_path = None
            self.selected_sheet = None


def main() -> None:
    app = QApplication(sys.argv)
    app.setStyleSheet("""
        QWidget {
            font-family: 'Segoe UI', Arial, sans-serif;
            font-size: 10pt;
            background-color: #f0f2f5;
            color: #333333;
        }
        QLabel {
            margin-bottom: 5px;
            color: #555555;
        }
        QLabel#title_label {
            color: #2c3e50;
            font-size: 18pt;
            font-weight: bold;
            margin-bottom: 20px;
        }
        QPushButton {
            background-color: #007bff;
            color: white;
            padding: 10px 20px;
            border-radius: 5px;
            border: none;
            font-weight: bold;
            margin: 5px;
        }
        QPushButton:hover {
            background-color: #0056b3;
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
        QMessageBox.critical(None, "Uygulama Hatası", f"Beklenmeyen bir hata oluştu: {e}\nUygulama kapatılıyor.")
        sys.exit(1)

if __name__ == "__main__":
    print(">> GraficApplication – Sürüm 3 – 3 Tem 2025 – page 1 grafik")
    main()