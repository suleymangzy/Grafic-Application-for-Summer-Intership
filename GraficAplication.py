import sys
import os
import numpy as np
import pandas as pd
from scipy import stats
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QAction, QMenu, QInputDialog,
    QMessageBox, QFileDialog, QLabel, QVBoxLayout, QWidget,
    QScrollArea, QProgressDialog, QPushButton, QComboBox, QLineEdit,
    QListWidget, QStackedWidget, QHBoxLayout, QListWidgetItem
)
from PyQt5.QtGui import QIcon, QFont, QPixmap
from PyQt5.QtCore import Qt, QDateTime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar

# Matplotlib varsayılan arka planını daha koyu bir temaya uyduralım
plt.style.use('dark_background')
plt.rcParams.update({
    'axes.facecolor': '#282828',
    'axes.edgecolor': '#888888',
    'axes.labelcolor': '#E0E0E0',
    'xtick.color': '#E0E0E0',
    'ytick.color': '#E0E0E0',
    'grid.color': '#444444',
    'text.color': '#E0E0E0',
    'figure.facecolor': '#282828',
    'savefig.facecolor': '#282828'
})


class AppState:
    def __init__(self):
        self.file_path = None
        self.df = None
        self.grouping_variable = None
        self.selected_grouping_value = "Tümünü Seç"  # New: to store selected value within grouping variable
        self.grouped_variable = None
        self.chart_variables = []
        self.chart_type = "Çizgi Grafiği"
        self.chart_count = 1


class StartPage(QWidget):
    def __init__(self, stacked_widget, state):
        super().__init__()
        self.stacked_widget = stacked_widget
        self.state = state

        layout = QVBoxLayout()
        title = QLabel("Grafik ve Rapor Uygulaması")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #E0E0E0;")

        btn = QPushButton("Rapor ve Grafik Oluştur")
        btn.clicked.connect(self.select_file)
        btn.setStyleSheet(
            "padding: 10px 20px; font-size: 16px; background-color: #555; color: #FFF; border-radius: 5px;")

        layout.addStretch()
        layout.addWidget(title)
        layout.addSpacing(50)
        layout.addWidget(btn, alignment=Qt.AlignCenter)
        layout.addStretch()

        self.setLayout(layout)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Dosya Seç",
                                                   filter="Excel Dosyaları (*.xlsx *.xls);;"
                                                          "CSV Dosyaları (*.csv);;"
                                                          "Tüm Desteklenen Dosyalar (*.xlsx *.xls *.csv);;"
                                                          "Tüm Dosyalar (*)")
        if file_path:
            ext = file_path.split('.')[-1].lower()
            try:
                na_values = ['', '#N/A', '#N/A N/A', '#NA', '-1.#IND', '-1.#QNAN', '-NaN', '-nan',
                             '1.#IND', '1.#QNAN', '<NA>', 'N/A', 'NA', 'NULL', 'NaN', 'n/a',
                             'nan', 'null', '?', '*', '-', ' ']

                df = None
                if ext in ['xlsx', 'xls']:
                    sheet_name, ok = QInputDialog.getText(
                        self, "Sayfa Adı", "Lütfen Excel sayfasının adını girin (örneğin: SMD-OEE):"
                    )
                    if not ok or not sheet_name:
                        QMessageBox.warning(self, "Uyarı", "Excel sayfası adı girilmedi. Dosya yükleme iptal edildi.")
                        return
                    df = pd.read_excel(file_path, sheet_name=sheet_name, na_values=na_values)
                elif ext == 'csv':
                    df = pd.read_csv(file_path, na_values=na_values)
                else:
                    raise ValueError("Desteklenmeyen dosya formatı. Lütfen Excel veya CSV dosyası seçin.")

                if df.empty:
                    raise ValueError("Dosya boş veya okunamadı")

                for col in df.columns:
                    # Attempt to convert to datetime. If over 50% success, convert.
                    # Prioritize DD.MM.YYYY format first, then general parse
                    converted_series = pd.to_datetime(df[col], format='%d.%m.%Y', errors='coerce')
                    if (converted_series.notna().sum() > 0) and \
                            (converted_series.notna().sum() / len(df) > 0.5):
                        df[col] = converted_series
                    else:
                        converted_series = pd.to_datetime(df[col], errors='coerce')
                        if (converted_series.notna().sum() > 0) and \
                                (converted_series.notna().sum() / len(df) > 0.5):
                            df[col] = converted_series

                self.state.df = df
                self.state.file_path = file_path
                self.stacked_widget.setCurrentIndex(1)
            except Exception as e:
                QMessageBox.warning(self, "Hata", f"Dosya yüklenirken hata oluştu: {str(e)}")


class ChartSettingsPage(QWidget):
    def __init__(self, stacked_widget, state):
        super().__init__()
        self.stacked_widget = stacked_widget
        self.state = state

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignTop | Qt.AlignHCenter)

        self.file_label = QLabel()
        self.file_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #ADD8E6;")
        layout.addWidget(self.file_label, alignment=Qt.AlignCenter)
        layout.addSpacing(20)

        chart_type_label = QLabel("Grafik Türü Seç:")
        chart_type_label.setStyleSheet("font-size: 14px; color: #E0E0E0;")
        layout.addWidget(chart_type_label)
        self.chart_combo = QComboBox()
        self.chart_combo.addItems(["Çizgi Grafiği", "Bar Grafiği", "Histogram", "Pasta Grafiği", "Dağılım Grafiği"])
        self.chart_combo.setStyleSheet(
            "font-size: 14px; padding: 5px; background-color: #444; color: #FFF; border-radius: 3px;")
        layout.addWidget(self.chart_combo)
        layout.addSpacing(15)

        count_label = QLabel("Kaç adet grafik oluşturulsun (1-10):")
        count_label.setStyleSheet("font-size: 14px; color: #E0E0E0;")
        layout.addWidget(count_label)
        self.count_input = QLineEdit("1")
        self.count_input.setPlaceholderText("1-10 arası adet girin")
        self.count_input.setStyleSheet(
            "font-size: 14px; padding: 5px; background-color: #444; color: #FFF; border-radius: 3px;")
        layout.addWidget(self.count_input)
        layout.addSpacing(30)

        btn_layout = QHBoxLayout()
        back = QPushButton("Geri")
        back.clicked.connect(lambda: self.stacked_widget.setCurrentIndex(0))
        next_btn = QPushButton("İleri")
        next_btn.clicked.connect(self.next_page)

        back.setStyleSheet(
            "padding: 10px 20px; font-size: 14px; background-color: #777; color: #FFF; border-radius: 5px;")
        next_btn.setStyleSheet(
            "padding: 10px 20px; font-size: 14px; background-color: #555; color: #FFF; border-radius: 5px;")

        btn_layout.addStretch()
        btn_layout.addWidget(back)
        btn_layout.addWidget(next_btn)
        btn_layout.addStretch()

        layout.addLayout(btn_layout)
        layout.addStretch()

        self.setLayout(layout)

    def showEvent(self, event):
        if self.state.file_path:
            filename = self.state.file_path.split('/')[-1]
            self.file_label.setText(f"Seçilen dosya: {filename}")
        else:
            self.file_label.setText("Lütfen bir dosya seçin.")

    def next_page(self):
        self.state.chart_type = self.chart_combo.currentText()
        try:
            count = int(self.count_input.text())
            if not (1 <= count <= 10):
                raise ValueError
            self.state.chart_count = count
            self.stacked_widget.setCurrentIndex(2)
        except ValueError:
            QMessageBox.warning(self, "Hata", "Grafik adedi için 1 ile 10 arasında geçerli bir sayı girin.")


class DataSelectionPage(QWidget):
    def __init__(self, stacked_widget, state):
        super().__init__()
        self.stacked_widget = stacked_widget
        self.state = state

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignTop | Qt.AlignHCenter)

        title_label = QLabel("Veri Sütunlarını Seçin")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #E0E0E0;")
        layout.addWidget(title_label, alignment=Qt.AlignCenter)
        layout.addSpacing(20)

        # Gruplama Değişkeni (Zorunlu)
        self.grouping_var_label = QLabel("Gruplama Değişkeni (Zorunlu):")
        self.grouping_var_label.setStyleSheet("font-size: 14px; color: #E0E0E0;")
        layout.addWidget(self.grouping_var_label)
        self.grouping_var_combo = QComboBox()
        self.grouping_var_combo.setStyleSheet(
            "font-size: 14px; padding: 5px; background-color: #444; color: #FFF; border-radius: 3px;")
        # Connect signal to populate grouping values
        self.grouping_var_combo.currentIndexChanged.connect(self.populate_grouping_values)
        layout.addWidget(self.grouping_var_combo)
        layout.addSpacing(15)

        # Gruplama Değişkeni Değeri Seçimi (Yeni)
        self.grouping_value_label = QLabel("Gruplama Değişkeni İçeriği Seç (İsteğe Bağlı):")
        self.grouping_value_label.setStyleSheet("font-size: 14px; color: #E0E0E0;")
        layout.addWidget(self.grouping_value_label)
        self.grouping_value_combo = QComboBox()
        self.grouping_value_combo.setStyleSheet(
            "font-size: 14px; padding: 5px; background-color: #444; color: #FFF; border-radius: 3px;")
        layout.addWidget(self.grouping_value_combo)
        layout.addSpacing(15)

        # Gruplanan Değişken (İsteğe Bağlı)
        self.grouped_var_label = QLabel("Gruplanan Değişken (İsteğe Bağlı):")
        self.grouped_var_label.setStyleSheet("font-size: 14px; color: #E0E0E0;")
        layout.addWidget(self.grouped_var_label)
        self.grouped_var_combo = QComboBox()
        self.grouped_var_combo.setStyleSheet(
            "font-size: 14px; padding: 5px; background-color: #444; color: #FFF; border-radius: 3px;")
        self.grouped_var_combo.addItem("Yok (Gruplama yapma)")
        layout.addWidget(self.grouped_var_combo)
        layout.addSpacing(15)

        # Grafik Değişkenleri (Zorunlu)
        self.chart_vars_label = QLabel("Grafik Değişkenleri (Zorunlu):")
        self.chart_vars_label.setStyleSheet("font-size: 14px; color: #E0E0E0;")
        layout.addWidget(self.chart_vars_label)
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.MultiSelection)
        self.list_widget.setStyleSheet("""
            QListWidget {
                background-color: #3A3A3A;
                color: #E0E0E0;
                border: 1px solid #555;
                border-radius: 5px;
                padding: 5px;
            }
            QListWidget::item {
                padding: 3px;
            }
            QListWidget::item:selected {
                background-color: #555;
                color: #FFF;
            }
        """)
        layout.addWidget(self.list_widget)
        layout.addSpacing(30)

        btn_layout = QHBoxLayout()
        back = QPushButton("Geri")
        back.clicked.connect(lambda: self.stacked_widget.setCurrentIndex(1))
        next_btn = QPushButton("İleri")
        next_btn.clicked.connect(self.next_page)

        back.setStyleSheet(
            "padding: 10px 20px; font-size: 14px; background-color: #777; color: #FFF; border-radius: 5px;")
        next_btn.setStyleSheet(
            "padding: 10px 20px; font-size: 14px; background-color: #555; color: #FFF; border-radius: 5px;")

        btn_layout.addStretch()
        btn_layout.addWidget(back)
        btn_layout.addWidget(next_btn)
        btn_layout.addStretch()

        layout.addLayout(btn_layout)
        layout.addStretch()
        self.setLayout(layout)

    def showEvent(self, event):
        if self.state.df is not None:
            cols = list(self.state.df.columns)
            self.grouping_var_combo.clear()
            self.grouped_var_combo.clear()
            self.list_widget.clear()

            # Populate Grouping Variable (typically date/time columns)
            datetime_cols = [col for col in cols if pd.api.types.is_datetime64_any_dtype(self.state.df[col])]
            if not datetime_cols:
                # If no datetime columns detected, offer all columns and warn user
                QMessageBox.warning(self, "Uyarı", "Veride tarih/zaman formatında bir sütun bulunamadı. "
                                                   "Lütfen 'Gruplama Değişkeni' olarak uygun bir sütun seçtiğinizden emin olun.")
                self.grouping_var_combo.addItems(cols)
            else:
                self.grouping_var_combo.addItems(datetime_cols)

            # Populate Grouped Variable (all columns + "Yok")
            self.grouped_var_combo.addItems(["Yok (Gruplama yapma)"] + cols)

            # Populate Chart Variables (numeric columns)
            numeric_cols = self.state.df.select_dtypes(include=np.number).columns.tolist()
            for col in numeric_cols:
                item = QListWidgetItem(col)
                item.setCheckState(Qt.Unchecked)
                self.list_widget.addItem(item)

            # Initial population of grouping values based on the first item in grouping_var_combo
            self.populate_grouping_values()  # Call this after populating grouping_var_combo

            self.update_labels_based_on_chart_type()

        else:
            QMessageBox.warning(self, "Uyarı", "Veri dosyası yüklenmemiş. Lütfen ilk sayfadan bir dosya seçin.")
            self.stacked_widget.setCurrentIndex(0)

    def populate_grouping_values(self):
        self.grouping_value_combo.clear()
        selected_grouping_var = self.grouping_var_combo.currentText()
        if selected_grouping_var and self.state.df is not None and selected_grouping_var in self.state.df.columns:
            # Get unique values, handle datetime formatting for display
            unique_values = self.state.df[selected_grouping_var].dropna().unique().tolist()
            # If datetime, format them for display in combo box
            if pd.api.types.is_datetime64_any_dtype(self.state.df[selected_grouping_var]):
                # Sort unique datetime values and convert to string for display
                unique_values = sorted([val.strftime('%Y-%m-%d %H:%M:%S') for val in unique_values])
            else:
                # Sort other unique values and convert to string for display
                unique_values = sorted([str(val) for val in unique_values])

            self.grouping_value_combo.addItem("Tümünü Seç")  # Option to select all
            self.grouping_value_combo.addItems(unique_values)

    def update_labels_based_on_chart_type(self):
        chart_type = self.state.chart_type
        if chart_type == "Çizgi Grafiği":
            self.grouping_var_label.setText("Gruplama Değişkeni (X Ekseni - Tarih/Zaman - Zorunlu):")
            self.grouped_var_label.setText("Gruplanan Değişken (Ürün/Kategori - İsteğe Bağlı):")
            self.chart_vars_label.setText("Grafik Değişkenleri (Y Ekseni - Sayısal - Zorunlu, çoklu seçim):")
        elif chart_type == "Bar Grafiği":
            self.grouping_var_label.setText("Gruplama Değişkeni (X Ekseni - Kategori/Tarih - Zorunlu):")
            self.grouped_var_label.setText("Gruplanan Değişken (Seri Gruplama - İsteğe Bağlı):")
            self.chart_vars_label.setText("Grafik Değişkenleri (Y Ekseni - Sayısal - Zorunlu, çoklu seçim):")
        elif chart_type == "Histogram":
            self.grouping_var_label.setText("Gruplama Değişkeni (İsteğe Bağlı - Veriyi bölmek için):")
            self.grouped_var_label.setText("Gruplanan Değişken (Seri Gruplama - İsteğe Bağlı):")
            self.chart_vars_label.setText("Grafik Değişkenleri (Değerler - Sayısal - Zorunlu, çoklu seçim):")
        elif chart_type == "Pasta Grafiği":
            self.grouping_var_label.setText("Gruplama Değişkeni (İsteğe Bağlı - Veriyi bölmek için):")
            self.grouped_var_label.setText("Gruplanan Değişken (Kategori Gruplama - İsteğe Bağlı):")
            self.chart_vars_label.setText("Grafik Değişkenleri (Değerler - Sayısal - Zorunlu, tek seçim önerilir):")
        elif chart_type == "Dağılım Grafiği":
            self.grouping_var_label.setText("Gruplama Değişkeni (İsteğe Bağlı - Veriyi bölmek için):")
            self.grouped_var_label.setText("Gruplanan Değişken (Seri Gruplama - İsteğe Bağlı):")
            self.chart_vars_label.setText("Grafik Değişkenleri (X ve Y Ekseni - Sayısal - Zorunlu, en az iki seçim):")
        else:
            self.grouping_var_label.setText("Gruplama Değişkeni (Zorunlu):")
            self.grouped_var_label.setText("Gruplanan Değişken (İsteğe Bağlı):")
            self.chart_vars_label.setText("Grafik Değişkenleri (Zorunlu):")

    def next_page(self):
        selected_chart_vars = [self.list_widget.item(i).text() for i in range(self.list_widget.count())
                               if self.list_widget.item(i).checkState() == Qt.Checked]

        self.state.grouping_variable = self.grouping_var_combo.currentText()
        self.state.selected_grouping_value = self.grouping_value_combo.currentText()  # Store selected value

        # Handle "Yok" for grouped variable
        current_grouped_text = self.grouped_var_combo.currentText()
        if current_grouped_text == "Yok (Gruplama yapma)":
            self.state.grouped_variable = None
        else:
            self.state.grouped_variable = current_grouped_text

        self.state.chart_variables = selected_chart_vars

        # --- Validation based on Chart Type ---
        if not self.state.grouping_variable:
            QMessageBox.warning(self, "Hata", "Lütfen bir 'Gruplama Değişkeni' seçin.")
            return

        if not self.state.chart_variables:
            QMessageBox.warning(self, "Hata", "Grafik için en az bir 'Grafik Değişkeni' seçmelisiniz.")
            return

        chart_type = self.state.chart_type
        if chart_type == "Dağılım Grafiği":
            if len(self.state.chart_variables) < 2:
                QMessageBox.warning(self, "Hata",
                                    "Dağılım grafiği için en az iki 'Grafik Değişkeni' seçmelisiniz (X ve Y eksenleri için).")
                return
            # For scatter plots, the grouping variable typically isn't used as an axis itself, but for splitting the data.
            # Ensure the grouping variable is compatible (e.g., date, or a categorical for separate plots)

        # Check if grouping variable is actually a datetime type for time-series charts
        # This is a soft check, as the user might select a non-datetime, which will cause issues later.
        # A stricter check could be implemented here if needed.
        if chart_type in ["Çizgi Grafiği", "Bar Grafiği"]:
            if not pd.api.types.is_datetime64_any_dtype(self.state.df[self.state.grouping_variable]):
                reply = QMessageBox.question(self, 'Uyarı',
                                             f"Seçilen '{self.state.grouping_variable}' sütunu tarih/zaman formatında görünmüyor. "
                                             "Çizgi veya Bar grafiği için tarih/zaman sütunu seçmek performansı artırır. "
                                             "Devam etmek istiyor musunuz?", QMessageBox.Yes | QMessageBox.No,
                                             QMessageBox.No)
                if reply == QMessageBox.No:
                    return

        self.stacked_widget.setCurrentIndex(3)


class ResultPage(QWidget):
    def __init__(self, stacked_widget, state):
        super().__init__()
        self.stacked_widget = stacked_widget
        self.state = state

        layout = QVBoxLayout()
        self.canvas = FigureCanvas(plt.Figure(figsize=(18, 12)))
        layout.addWidget(self.canvas)

        btn_layout = QHBoxLayout()
        back = QPushButton("Geri")
        back.clicked.connect(lambda: self.stacked_widget.setCurrentIndex(2))

        back.setStyleSheet(
            "padding: 10px 20px; font-size: 14px; background-color: #777; color: #FFF; border-radius: 5px;")

        btn_layout.addStretch()
        btn_layout.addWidget(back)
        btn_layout.addStretch()

        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def showEvent(self, event):
        self.plot_graphs()

    def plot_graphs(self):
        df = self.state.df.copy()

        # Ensure grouping variable exists and is processed
        grouping_col = self.state.grouping_variable
        if not grouping_col or grouping_col not in df.columns:
            QMessageBox.warning(self, "Hata", "Gruplama değişkeni seçilmedi veya dosyada bulunamadı.")
            self.stacked_widget.setCurrentIndex(2)
            return

        # Attempt to convert grouping column to datetime if it's not already
        if not pd.api.types.is_datetime64_any_dtype(df[grouping_col]):
            df[grouping_col] = pd.to_datetime(df[grouping_col], errors='coerce')
            df = df.dropna(subset=[grouping_col])  # Remove rows where grouping_col became NaT

        # Filter by selected grouping value if not "Tümünü Seç"
        if self.state.selected_grouping_value != "Tümünü Seç":
            # Need to handle datetime comparison carefully
            if pd.api.types.is_datetime64_any_dtype(df[grouping_col]):
                try:
                    # Convert selected_grouping_value back to datetime for comparison
                    filter_value = pd.to_datetime(self.state.selected_grouping_value)
                    # For date comparison, compare just the date part if times might differ
                    df = df[df[grouping_col].dt.date == filter_value.date()].copy()
                except Exception as e:
                    QMessageBox.warning(self, "Hata",
                                        f"Tarih filtresi uygulanırken hata: {str(e)}. Tüm veriler kullanılacak.")
                    # Continue without filter if error
            else:
                df = df[df[grouping_col].astype(str) == self.state.selected_grouping_value].copy()

        if df.empty:
            QMessageBox.information(self, "Bilgi", "Seçilen kriterlere göre işlenecek veri bulunamadı.")
            return

        # Prepare data for plotting
        grouped_df = None
        group_keys = []
        if self.state.grouped_variable and self.state.grouped_variable in df.columns:
            # Drop NaN values in the grouped variable column to ensure proper grouping
            df = df.dropna(subset=[self.state.grouped_variable])
            # First group by the grouping_variable (e.g., Date), then by the grouped_variable (e.g., Product)
            grouped_df = df.groupby([grouping_col, self.state.grouped_variable])
            group_keys = list(grouped_df.groups.keys())
        else:
            # If no grouped variable, just group by the grouping variable
            grouped_df = df.groupby(grouping_col)
            group_keys = list(grouped_df.groups.keys())

        fig = self.canvas.figure
        fig.clear()

        plots_to_draw = []

        # Determine the actual plots to draw based on chart_count and available data
        if self.state.grouped_variable:  # Grouped by both grouping_col and grouped_variable
            for (group_val_main, group_val_sub), group_data in grouped_df:
                for chart_var in self.state.chart_variables:
                    plots_to_draw.append(((group_val_main, group_val_sub), chart_var, group_data))
                    if len(plots_to_draw) >= self.state.chart_count:
                        break
                if len(plots_to_draw) >= self.state.chart_count:
                    break
        else:  # Grouped only by grouping_col or no explicit grouping for all selected chart variables
            if self.state.chart_type in ["Pasta Grafiği", "Histogram"]:
                # For pie/histograms without a specific 'grouped_variable', we plot each chart_variable for each grouping_variable instance
                for group_val_main, group_data in grouped_df:
                    for chart_var in self.state.chart_variables:
                        plots_to_draw.append((group_val_main, chart_var, group_data))
                        if len(plots_to_draw) >= self.state.chart_count:
                            break
                    if len(plots_to_draw) >= self.state.chart_count:
                        break
            elif self.state.chart_type in ["Çizgi Grafiği", "Bar Grafiği", "Dağılım Grafiği"]:
                # If no 'grouped_variable', we usually want to plot for the entire dataset for the selected chart variables,
                # possibly using the grouping_col as X-axis for line/bar charts.
                # 'df' is already filtered by 'selected_grouping_value' if applicable.
                plots_to_draw.append(
                    (None, self.state.chart_variables, df))  # Represents overall plot for selected chart vars

        num_plots_to_draw = len(plots_to_draw)
        if num_plots_to_draw == 0:
            QMessageBox.information(self, "Bilgi", "Seçilen kriterlere göre çizilebilecek grafik bulunamadı.")
            return

        num_cols = int(np.ceil(np.sqrt(num_plots_to_draw)))
        num_rows = int(np.ceil(num_plots_to_draw / num_cols))
        current_plot_index = 1

        for plot_info in plots_to_draw:
            if current_plot_index > self.state.chart_count:
                break

            ax = fig.add_subplot(num_rows, num_cols, current_plot_index)
            ax.tick_params(axis='x', rotation=45, labelsize=6)
            ax.tick_params(axis='y', labelsize=6)
            ax.legend(loc='lower right', fontsize=6)

            base_title = "Tüm Veri"  # Default title part if not specified by grouping
            data_to_plot = None  # Initialize
            chart_var = None  # Initialize for single chart variable plots

            if self.state.grouped_variable:
                (group_val_main, group_val_sub), chart_var, data_to_plot = plot_info
                # If grouping_col is datetime, format for title
                if pd.api.types.is_datetime64_any_dtype(self.state.df[grouping_col]):
                    base_title = f"{pd.to_datetime(group_val_main).date()} - {group_val_sub}"
                else:
                    base_title = f"{group_val_main} - {group_val_sub}"

            else:  # No grouped_variable
                if self.state.chart_type in ["Pasta Grafiği", "Histogram"]:
                    group_val_main, chart_var, data_to_plot = plot_info
                    if pd.api.types.is_datetime64_any_dtype(self.state.df[grouping_col]) and group_val_main:
                        base_title = f"{pd.to_datetime(group_val_main).date()}"
                    elif group_val_main:
                        base_title = f"{group_val_main}"
                    else:
                        base_title = "Tüm Veri"  # Fallback if no specific grouping value used
                else:  # Line, Bar, Scatter without grouped_variable
                    # This branch assumes plot_info is (None, self.state.chart_variables, df_filtered)
                    # where df_filtered is the data after any selected_grouping_value filter.
                    _, chart_vars_for_plot, data_to_plot = plot_info
                    # chart_var remains None here, handled by iterating self.state.chart_variables
                    if self.state.selected_grouping_value != "Tümünü Seç" and self.state.selected_grouping_value:
                        base_title = f"Gruplama Değeri: {self.state.selected_grouping_value}"
                    else:
                        base_title = "Tüm Veri"

            # Plotting logic based on chart type
            if self.state.chart_type == "Çizgi Grafiği":
                if self.state.grouped_variable:  # Grouped by both grouping_col and grouped_variable
                    ax.plot(data_to_plot[grouping_col], data_to_plot[chart_var], label=chart_var)
                    ax.set_title(f"Çizgi Grafiği: {base_title} - {chart_var}", fontsize=8)
                else:  # No grouped_variable, plot all selected chart variables on one plot
                    for var in self.state.chart_variables:
                        ax.plot(data_to_plot[grouping_col], data_to_plot[var], label=var)
                    ax.set_title(f"Çizgi Grafiği: {base_title}", fontsize=8)
                ax.set_xlabel(grouping_col, fontsize=7)
                ax.set_ylabel("Değer", fontsize=7)

            elif self.state.chart_type == "Bar Grafiği":
                if self.state.grouped_variable:  # Grouped by both grouping_col and grouped_variable
                    summary_data = data_to_plot.groupby(grouping_col)[chart_var].sum()  # Example: sum
                    summary_data.plot(kind='bar', ax=ax, label=chart_var)
                    ax.set_title(f"Bar Grafiği: {base_title} - {chart_var}", fontsize=8)
                else:  # No grouped_variable
                    for var in self.state.chart_variables:
                        summary_data = data_to_plot.groupby(grouping_col)[var].sum()
                        summary_data.plot(kind='bar', ax=ax, label=var)
                    ax.set_title(f"Bar Grafiği: {base_title}", fontsize=8)
                ax.set_xlabel(grouping_col, fontsize=7)
                ax.set_ylabel("Toplam Değer", fontsize=7)

            elif self.state.chart_type == "Histogram":
                # For histograms, we generally plot distributions of numerical columns.
                # If grouped, it's the distribution within that group.
                # If not grouped, it's the distribution of the overall column.
                data_for_hist = data_to_plot[chart_var].dropna()
                if not data_for_hist.empty:
                    ax.hist(data_for_hist, bins=10, alpha=0.7, label=chart_var)
                    ax.set_title(f"Histogram: {base_title} - {chart_var}", fontsize=8)
                    ax.set_xlabel("Değer Aralığı", fontsize=7)
                    ax.set_ylabel("Frekans", fontsize=7)
                else:
                    ax.set_title(f"Histogram: {base_title} - {chart_var} (Veri Yok)", fontsize=8)


            elif self.state.chart_type == "Pasta Grafiği":
                # For pie charts, we need counts/proportions of categorical data or sum of numerical data
                # against a categorical grouping.
                if pd.api.types.is_numeric_dtype(data_to_plot[chart_var]):
                    if self.state.grouped_variable:
                        counts = data_to_plot.groupby(self.state.grouped_variable)[chart_var].sum()
                        pie_title_var = self.state.grouped_variable
                    else:
                        counts = data_to_plot.groupby(grouping_col)[chart_var].sum()
                        pie_title_var = grouping_col

                    if not counts.empty:
                        wedges, texts = ax.pie(counts, autopct='%1.1f%%', textprops={'fontsize': 7})
                        ax.set_title(f"Pasta Grafiği: {base_title} - {chart_var} by {pie_title_var}", fontsize=8)
                        ax.axis('equal')
                        ax.legend(wedges, counts.index.astype(str), title=pie_title_var, loc='lower right', fontsize=6,
                                  bbox_to_anchor=(1.05, 0), fancybox=True, shadow=True, borderpad=0.5, labelspacing=0.5)
                    else:
                        ax.set_title(f"Pasta Grafiği: {base_title} - {chart_var} (Veri Yok)", fontsize=8)

                else:  # Treat chart_var as categorical for pie slices (value counts)
                    counts = data_to_plot[chart_var].value_counts()
                    if not counts.empty:
                        wedges, texts = ax.pie(counts, autopct='%1.1f%%', textprops={'fontsize': 7})
                        ax.set_title(f"Pasta Grafiği: {base_title} - {chart_var}", fontsize=8)
                        ax.axis('equal')
                        ax.legend(wedges, counts.index, title="Kategoriler", loc='lower right', fontsize=6,
                                  bbox_to_anchor=(1.05, 0), fancybox=True, shadow=True, borderpad=0.5, labelspacing=0.5)
                    else:
                        ax.set_title(f"Pasta Grafiği: {base_title} - {chart_var} (Veri Yok)", fontsize=8)


            elif self.state.chart_type == "Dağılım Grafiği":
                # Requires at least two chart variables (X and Y).
                x_col = self.state.chart_variables[0]
                y_col = self.state.chart_variables[1] if len(self.state.chart_variables) > 1 else \
                self.state.chart_variables[0]  # Fallback

                ax.scatter(data_to_plot[x_col], data_to_plot[y_col], label=f'{x_col} vs {y_col}')
                ax.set_title(f"Dağılım Grafiği: {base_title}", fontsize=8)
                ax.set_xlabel(x_col, fontsize=7)
                ax.set_ylabel(y_col, fontsize=7)

            current_plot_index += 1

        fig.tight_layout()
        self.canvas.draw()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Çok Sayfalı Grafik ve Rapor Uygulaması")
        self.state = AppState()

        self.stacked_widget = QStackedWidget()
        self.stacked_widget.addWidget(StartPage(self.stacked_widget, self.state))
        self.stacked_widget.addWidget(ChartSettingsPage(self.stacked_widget, self.state))
        self.stacked_widget.addWidget(DataSelectionPage(self.stacked_widget, self.state))
        self.stacked_widget.addWidget(ResultPage(self.stacked_widget, self.state))

        self.setCentralWidget(self.stacked_widget)
        self.setStyleSheet("""
            QMainWindow { background-color: #202020; }
            QLabel { color: #E0E0E0; }
            QPushButton {
                background-color: #555;
                color: #FFF;
                border: 1px solid #777;
                padding: 10px 20px;
                border-radius: 5px;
            }
            QPushButton:hover { background-color: #666; }
            QPushButton:pressed { background-color: #444; }
            QComboBox {
                background-color: #444;
                color: #FFF;
                border: 1px solid #666;
                padding: 5px;
                border-radius: 3px;
            }
            QComboBox::drop-down { border-left: 1px solid #666; }
            QComboBox QAbstractItemView {
                background-color: #444;
                color: #FFF;
                selection-background-color: #666;
            }
            QLineEdit {
                background-color: #444;
                color: #FFF;
                border: 1px solid #666;
                padding: 5px;
                border-radius: 3px;
            }
            QListWidget {
                background-color: #3A3A3A;
                color: #E0E0E0;
                border: 1px solid #555;
                border-radius: 5px;
                padding: 5px;
            }
            QListWidget::item {
                padding: 3px;
            }
            QListWidget::item:selected {
                background-color: #555;
                color: #FFF;
            }
        """)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(1000, 750)
    window.show()
    sys.exit(app.exec_())