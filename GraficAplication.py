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
) # Yeni eklenenler: QPushButton, QComboBox, QLineEdit, QListWidget, QStackedWidget, QHBoxLayout, QListWidgetItem
from PyQt5.QtGui import QIcon, QFont, QPixmap
from PyQt5.QtCore import Qt, QDateTime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
# fpdf importları eksikse ekleyin, önceki kodda vardı
# from fpdf import FPDF
# from fpdf.enums import XPos, YPos

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

# AppState, ChartSettingsPage, DataSelectionPage, ResultPage sınıflarının
# ve MainWindow sınıfının tanımlarının burada olduğunu varsayıyorum.
# Tam bir çalışan kod için bu sınıfların hepsi mevcut olmalı.

# AppState sınıfı (önceki yanıtımdan alınmıştır)
class AppState:
    def __init__(self):
        self.file_path = None
        self.df = None
        self.date_column = None
        self.product_column = None
        self.selected_columns = []
        self.chart_type = "Çizgi Grafiği"
        self.chart_count = 1

# Diğer sınıflar (ChartSettingsPage, DataSelectionPage, ResultPage) da burada olmalı
# Sizin tarafınızdan sağlandığı varsayılıyor.

class StartPage(QWidget):
    def __init__(self, stacked_widget, state):
        super().__init__()
        self.stacked_widget = stacked_widget
        self.state = state

        layout = QVBoxLayout()
        title = QLabel("Grafik ve Rapor Uygulaması")
        title.setAlignment(Qt.AlignCenter)
        # Stil ekle
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #E0E0E0;")

        btn = QPushButton("Rapor ve Grafik Oluştur")
        btn.clicked.connect(self.select_file)
        # Stil ekle
        btn.setStyleSheet(
            "padding: 10px 20px; font-size: 16px; background-color: #555; color: #FFF; border-radius: 5px;")

        layout.addStretch()  # Üste it
        layout.addWidget(title)
        layout.addSpacing(50)
        layout.addWidget(btn, alignment=Qt.AlignCenter)
        layout.addStretch()  # Alta it

        self.setLayout(layout)

    def select_file(self):
        # Dosya filtresi güncellendi
        file_path, _ = QFileDialog.getOpenFileName(self, "Dosya Seç",
                                                   filter="Excel Dosyaları (*.xlsx *.xls);;"
                                                          "CSV Dosyaları (*.csv);;"
                                                          "Word Dosyaları (*.docx);;"
                                                          "PowerPoint Dosyaları (*.pptx);;"
                                                          "Tüm Desteklenen Dosyalar (*.xlsx *.xls *.csv *.docx *.pptx);;"
                                                          "Tüm Dosyalar (*)")
        if file_path:
            ext = file_path.split('.')[-1].lower()  # Uzantıyı küçük harfe çevir
            try:
                na_values = ['', '#N/A', '#N/A N/A', '#NA', '-1.#IND', '-1.#QNAN', '-NaN', '-nan',
                             '1.#IND', '1.#QNAN', '<NA>', 'N/A', 'NA', 'NULL', 'NaN', 'n/a',
                             'nan', 'null', '?', '*', '-', ' ']

                if ext in ['xlsx', 'xls']:
                    sheet_name, ok = QInputDialog.getText(
                        self, "Sayfa Adı", "Lütfen Excel sayfasının adını girin (örneğin: Sayfa1):"
                    )
                    if not ok or not sheet_name:
                        QMessageBox.warning(self, "Uyarı", "Excel sayfası adı girilmedi. Dosya yükleme iptal edildi.")
                        return
                    df = pd.read_excel(file_path, sheet_name=sheet_name, na_values=na_values)
                elif ext == 'csv':
                    df = pd.read_csv(file_path, na_values=na_values)
                elif ext in ['docx', 'pptx']: # Yeni eklenen formatlar için uyarı
                    raise ValueError(f"'{ext}' formatındaki dosyalar şu an için desteklenmemektedir. "
                                     "Lütfen Excel veya CSV dosyası seçin.")
                else:
                    raise ValueError("Desteklenmeyen dosya formatı")

                if df.empty:
                    raise ValueError("Dosya boş veya okunamadı")

                for col in df.columns:
                    converted_series = pd.to_datetime(df[col], format='%d.%m.%Y', errors='coerce')
                    # Yüzde 50'den fazlası tarih olarak başarıyla dönüştürülebiliyorsa dönüştür
                    if (converted_series.notna().sum() > 0) and \
                            (converted_series.notna().sum() / len(df) > 0.5): # Yüzde 50'den fazlası
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
        layout.setAlignment(Qt.AlignTop | Qt.AlignHCenter)  # Üste ve ortaya hizala

        self.file_label = QLabel()
        self.file_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #ADD8E6;")
        layout.addWidget(self.file_label, alignment=Qt.AlignCenter)
        layout.addSpacing(20)

        # Grafik türü seçimi
        chart_type_label = QLabel("Grafik Türü Seç:")
        chart_type_label.setStyleSheet("font-size: 14px; color: #E0E0E0;")
        layout.addWidget(chart_type_label)
        self.chart_combo = QComboBox()
        self.chart_combo.addItems(["Çizgi Grafiği", "Bar Grafiği", "Histogram", "Pasta Grafiği", "Dağılım Grafiği"])
        self.chart_combo.setStyleSheet(
            "font-size: 14px; padding: 5px; background-color: #444; color: #FFF; border-radius: 3px;")
        layout.addWidget(self.chart_combo)
        layout.addSpacing(15)

        # Grafik adedi girişi
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

        # Buton stilleri
        back.setStyleSheet(
            "padding: 10px 20px; font-size: 14px; background-color: #777; color: #FFF; border-radius: 5px;")
        next_btn.setStyleSheet(
            "padding: 10px 20px; font-size: 14px; background-color: #555; color: #FFF; border-radius: 5px;")

        btn_layout.addStretch()  # Sağa it
        btn_layout.addWidget(back)
        btn_layout.addWidget(next_btn)
        btn_layout.addStretch()  # Sağa it

        layout.addLayout(btn_layout)
        layout.addStretch()  # Sayfa içeriğini üste it

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
            if not (1 <= count <= 10):  # 1 ile 10 arasında kontrol
                raise ValueError
            self.state.chart_count = count
            self.stacked_widget.setCurrentIndex(2)  # DataSelectionPage'e geç
        except ValueError:
            QMessageBox.warning(self, "Hata", "Grafik adedi için 1 ile 10 arasında geçerli bir sayı girin.")


class DataSelectionPage(QWidget):
    def __init__(self, stacked_widget, state):
        super().__init__()
        self.stacked_widget = stacked_widget
        self.state = state

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignTop | Qt.AlignHCenter)

        # Başlık
        title_label = QLabel("Veri Sütunlarını Seçin")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #E0E0E0;")
        layout.addWidget(title_label, alignment=Qt.AlignCenter)
        layout.addSpacing(20)

        # Tarih sütunu seçimi
        date_label = QLabel("Tarih/Zaman sütununu seçin:")
        date_label.setStyleSheet("font-size: 14px; color: #E0E0E0;")
        layout.addWidget(date_label)
        self.date_combo = QComboBox()
        self.date_combo.setStyleSheet(
            "font-size: 14px; padding: 5px; background-color: #444; color: #FFF; border-radius: 3px;")
        layout.addWidget(self.date_combo)
        layout.addSpacing(15)

        # Ürün/Kategori sütunu seçimi
        product_label = QLabel("Ürün/Kategori sütununu seçin (İsteğe Bağlı):")
        product_label.setStyleSheet("font-size: 14px; color: #E0E0E0;")
        layout.addWidget(product_label)
        self.product_combo = QComboBox()
        self.product_combo.setStyleSheet(
            "font-size: 14px; padding: 5px; background-color: #444; color: #FFF; border-radius: 3px;")
        self.product_combo.addItem("Yok (Gruplama yapma)")  # Ürün sütunu olmadan da devam edebilme seçeneği
        layout.addWidget(self.product_combo)
        layout.addSpacing(15)

        # Diğer sütunları seçme
        other_cols_label = QLabel("Grafikte kullanılacak diğer sayısal sütunları seçin:")
        other_cols_label.setStyleSheet("font-size: 14px; color: #E0E0E0;")
        layout.addWidget(other_cols_label)
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.MultiSelection)  # Çoklu seçim
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

        # Buton stilleri
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
            self.date_combo.clear()
            self.product_combo.clear()
            self.list_widget.clear()

            self.date_combo.addItems(cols)
            # Ürün/kategori sütunu için "Yok" seçeneğini koru
            self.product_combo.addItems(["Yok (Gruplama yapma)"] + cols)

            # Sadece sayısal sütunları list_widget'a ekle
            numeric_cols = self.state.df.select_dtypes(include=np.number).columns.tolist()
            for col in numeric_cols:
                item = QListWidgetItem(col)
                item.setCheckState(Qt.Unchecked)
                self.list_widget.addItem(item)
        else:
            QMessageBox.warning(self, "Uyarı", "Veri dosyası yüklenmemiş. Lütfen ilk sayfadan bir dosya seçin.")
            self.stacked_widget.setCurrentIndex(0)  # Dosya seçme sayfasına dön

    def next_page(self):
        selected = [self.list_widget.item(i).text() for i in range(self.list_widget.count())
                    if self.list_widget.item(i).checkState() == Qt.Checked]

        if not selected:
            QMessageBox.warning(self, "Hata", "Grafikte kullanılacak en az bir sayısal sütun seçmelisiniz.")
            return

        if not self.date_combo.currentText():
            QMessageBox.warning(self, "Hata", "Lütfen bir tarih/zaman sütunu seçin.")
            return

        self.state.date_column = self.date_combo.currentText()

        # Eğer 'Yok' seçilmişse product_column None olarak ayarla
        self.state.product_column = self.product_combo.currentText()
        if self.state.product_column == "Yok (Gruplama yapma)":
            self.state.product_column = None

        self.state.selected_columns = selected
        self.stacked_widget.setCurrentIndex(3)  # ResultPage'e geç


class ResultPage(QWidget):
    def __init__(self, stacked_widget, state):
        super().__init__()
        self.stacked_widget = stacked_widget
        self.state = state

        layout = QVBoxLayout()
        # Figür boyutunu artırarak daha fazla yer sağlayın
        # Ekran boyutunuza ve çizilecek alt grafik sayısına göre ayarlayın
        self.canvas = FigureCanvas(plt.Figure(figsize=(18, 12)))
        layout.addWidget(self.canvas)

        btn_layout = QHBoxLayout()
        back = QPushButton("Geri")
        back.clicked.connect(lambda: self.stacked_widget.setCurrentIndex(2))

        # Buton stilleri
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

        # Tarih sütununu kontrol et ve dönüştür
        if self.state.date_column and self.state.date_column in df.columns:
            df[self.state.date_column] = pd.to_datetime(df[self.state.date_column], errors='coerce')
            df = df.dropna(subset=[self.state.date_column])  # Tarih sütunundaki NaN'leri sil
        else:
            QMessageBox.warning(self, "Hata", "Geçerli bir tarih sütunu seçilmedi veya dosyada bulunamadı.")
            self.stacked_widget.setCurrentIndex(2)  # Veri seçme sayfasına geri dön
            return

        # Ürün sütununa göre gruplama yapılacaksa
        if self.state.product_column and self.state.product_column in df.columns:
            df = df.dropna(subset=[self.state.product_column])  # Ürün sütunundaki NaN'leri sil
            grouped_data = df.groupby([self.state.date_column, self.state.product_column])
        else:
            # Sadece tarih sütununa göre grupla veya hiç gruplama yapma
            # Örneğin, her bir tarih için ayrı bir plot istiyorsak:
            grouped_data = df.groupby(self.state.date_column)
            # Ya da tüm veri seti için tek bir plot istiyorsak, gruplamayı farklı yapmalıyız.
            # Şimdilik, product_column yoksa sadece date_column'a göre gruplama yapalım.
            # Eğer ChartType Pasta Grafiği ise ve product_column yoksa, tüm selected_columns için çizmeliyiz.

        fig = self.canvas.figure
        fig.clear()

        plots_generated = 0
        total_possible_plots = 0

        # Gruplanmış veri sayısını ve seçilen sütun sayısını kullanarak toplam potansiyel plot sayısını hesaplayalım
        if self.state.product_column:  # Hem tarih hem ürün bazında gruplama varsa
            total_possible_plots = len(grouped_data) * len(self.state.selected_columns)
        else:  # Sadece tarih bazında gruplama varsa veya hiç gruplama yoksa (sadece selected_columns)
            # Bu senaryoda her bir seçili sütun için ayrı bir pasta grafiği çizilebilir
            if self.state.chart_type == "Pasta Grafiği":
                total_possible_plots = len(grouped_data) * len(self.state.selected_columns)
            else:  # Diğer grafik türleri için daha farklı olabilir. Şimdilik her selected_column için bir plot varsayalım
                total_possible_plots = len(self.state.selected_columns)  # Her sütun için ayrı plot

        # Çizilecek alt grafiklerin maksimum sayısı, kullanıcının seçimi veya potansiyel sayının küçüğü
        num_plots_to_draw = min(self.state.chart_count, total_possible_plots)

        if num_plots_to_draw == 0:
            QMessageBox.information(self, "Bilgi", "Seçilen kriterlere göre çizilebilecek grafik bulunamadı.")
            return

        # Alt grafikler için uygun bir ızgara boyutu hesaplayın
        num_cols = int(np.ceil(np.sqrt(num_plots_to_draw)))
        num_rows = int(np.ceil(num_plots_to_draw / num_cols))

        # Eğer tek bir grup ve tek bir sütun seçiliyse, yani tek bir grafik çizilecekse,
        # eksenlerin yerleşimini doğrudan ayarlayabiliriz.
        # Bu sadece pasta grafiği için geçerlidir. Diğer grafik türleri için farklı yaklaşımlar gerekebilir.

        current_plot_index = 1  # Alt grafik indeksini takip etmek için

        if self.state.chart_type == "Pasta Grafiği":
            # Pasta Grafiği için gruplama mantığı
            if self.state.product_column:  # Hem tarih hem ürün bazında gruplama varsa
                for (date, product), group in grouped_data:
                    if plots_generated >= self.state.chart_count:
                        break
                    for col in self.state.selected_columns:
                        if plots_generated >= self.state.chart_count:
                            break

                        data_to_plot = group[col].dropna()
                        if data_to_plot.empty: continue

                        counts = data_to_plot.value_counts()
                        if counts.empty: continue

                        ax = fig.add_subplot(num_rows, num_cols, current_plot_index)
                        # Etiket ve autopct kaldırıldı, sadece renk ve dilimler çiziliyor
                        wedges, texts = ax.pie(counts, textprops={'fontsize': 7})  # wedges ve texts döndürür

                        ax.set_title(f"{date.date()} - {product} - {col}", fontsize=8)
                        ax.axis('equal')  # Pasta grafiklerinin daire şeklinde olmasını sağlar

                        # Legend (açıklama) ekle
                        # loc='lower right' ile alt grafik içinde sağ alta yerleştir
                        ax.legend(wedges, counts.index, title="Kategoriler", loc='lower right', fontsize=6,
                                  bbox_to_anchor=(1.05, 0),
                                  fancybox=True, shadow=True, borderpad=0.5,
                                  labelspacing=0.5)  # bbox_to_anchor ile biraz dışarı it

                        current_plot_index += 1
                        plots_generated += 1
            else:  # Sadece tarih veya hiç gruplama yoksa (her seçili sütun için ayrı pasta grafiği)
                # Tüm veri setini kullan veya tarih bazında toplamları al
                for col in self.state.selected_columns:
                    if plots_generated >= self.state.chart_count:
                        break

                    data_to_plot = df[col].dropna()  # Tüm veri setinden seçili sütunu al
                    if data_to_plot.empty: continue

                    counts = data_to_plot.value_counts()
                    if counts.empty: continue

                    ax = fig.add_subplot(num_rows, num_cols, current_plot_index)
                    wedges, texts = ax.pie(counts, textprops={'fontsize': 7})

                    # Eğer tarih gruplaması varsa, başlığa tarihi ekleyebiliriz
                    title_text = f"{col}"
                    if self.state.date_column:
                        # Tek bir tarih veya tüm verinin ortalama tarihi olabilir.
                        # Burada sadece sütun adını başlık yapalım.
                        pass
                    ax.set_title(title_text, fontsize=8)
                    ax.axis('equal')

                    ax.legend(wedges, counts.index, title="Kategoriler", loc='lower right', fontsize=6,
                              bbox_to_anchor=(1.05, 0),
                              fancybox=True, shadow=True, borderpad=0.5, labelspacing=0.5)

                    current_plot_index += 1
                    plots_generated += 1
        elif self.state.chart_type == "Çizgi Grafiği":
            # Çizgi grafiği örneği (şimdilik basit bir örnek, daha detaylı gruplama ve sütun seçimi gerekebilir)
            # Tarih sütunu X ekseninde, selected_columns Y ekseninde olacak.

            # Gruplama varsa her grup için ayrı çizgi
            if self.state.product_column:
                for (date, product), group in grouped_data:
                    if plots_generated >= self.state.chart_count:
                        break
                    ax = fig.add_subplot(num_rows, num_cols, current_plot_index)
                    for col in self.state.selected_columns:
                        ax.plot(group[self.state.date_column], group[col], label=col)
                    ax.set_title(f"{date.date()} - {product}", fontsize=8)
                    ax.set_xlabel("Tarih", fontsize=7)
                    ax.set_ylabel("Değer", fontsize=7)
                    ax.tick_params(axis='x', rotation=45, labelsize=6)
                    ax.tick_params(axis='y', labelsize=6)
                    ax.legend(loc='lower right', fontsize=6)
                    plots_generated += 1
                    current_plot_index += 1
            else:  # Gruplama yoksa, her selected_column için ayrı bir plot
                ax = fig.add_subplot(num_rows, num_cols, current_plot_index)
                for col in self.state.selected_columns:
                    ax.plot(df[self.state.date_column], df[col], label=col)
                ax.set_title("Çizgi Grafiği", fontsize=8)
                ax.set_xlabel("Tarih", fontsize=7)
                ax.set_ylabel("Değer", fontsize=7)
                ax.tick_params(axis='x', rotation=45, labelsize=6)
                ax.tick_params(axis='y', labelsize=6)
                ax.legend(loc='lower right', fontsize=6)
                plots_generated += 1
                current_plot_index += 1
        # Diğer grafik türleri (Bar, Histogram, Dağılım) için de benzer şekilde plot_graphs metodunu genişletmeniz gerekecektir.
        # Örneğin:
        elif self.state.chart_type == "Bar Grafiği":
            if self.state.product_column:
                for (date, product), group in grouped_data:
                    if plots_generated >= self.state.chart_count:
                        break
                    ax = fig.add_subplot(num_rows, num_cols, current_plot_index)
                    for col in self.state.selected_columns:
                        # Bar grafiği için genellikle kategorik veriye ihtiyaç duyulur.
                        # Burada her seçili sütunun ortalamasını veya toplamını gösterebiliriz.
                        summary_data = group.groupby(self.state.date_column)[col].sum()  # Örnek toplama
                        summary_data.plot(kind='bar', ax=ax, label=col)
                        ax.set_title(f"{date.date()} - {product}", fontsize=8)
                        ax.set_xlabel("Tarih", fontsize=7)
                        ax.set_ylabel("Toplam Değer", fontsize=7)
                        ax.tick_params(axis='x', rotation=45, labelsize=6)
                        ax.tick_params(axis='y', labelsize=6)
                        ax.legend(loc='lower right', fontsize=6)
                        plots_generated += 1
                        current_plot_index += 1
            else:
                ax = fig.add_subplot(num_rows, num_cols, current_plot_index)
                for col in self.state.selected_columns:
                    summary_data = df.groupby(self.state.date_column)[col].sum()
                    summary_data.plot(kind='bar', ax=ax, label=col)
                ax.set_title("Bar Grafiği", fontsize=8)
                ax.set_xlabel("Tarih", fontsize=7)
                ax.set_ylabel("Toplam Değer", fontsize=7)
                ax.tick_params(axis='x', rotation=45, labelsize=6)
                ax.tick_params(axis='y', labelsize=6)
                ax.legend(loc='lower right', fontsize=6)
                plots_generated += 1
                current_plot_index += 1
        elif self.state.chart_type == "Histogram":
            if self.state.product_column:
                for (date, product), group in grouped_data:
                    if plots_generated >= self.state.chart_count:
                        break
                    ax = fig.add_subplot(num_rows, num_cols, current_plot_index)
                    for col in self.state.selected_columns:
                        group[col].hist(ax=ax, alpha=0.7, label=col)
                    ax.set_title(f"Histogram: {date.date()} - {product}", fontsize=8)
                    ax.set_xlabel("Değer Aralığı", fontsize=7)
                    ax.set_ylabel("Frekans", fontsize=7)
                    ax.tick_params(axis='x', labelsize=6)
                    ax.tick_params(axis='y', labelsize=6)
                    ax.legend(loc='lower right', fontsize=6)
                    plots_generated += 1
                    current_plot_index += 1
            else:
                ax = fig.add_subplot(num_rows, num_cols, current_plot_index)
                for col in self.state.selected_columns:
                    df[col].hist(ax=ax, alpha=0.7, label=col)
                ax.set_title("Histogram", fontsize=8)
                ax.set_xlabel("Değer Aralığı", fontsize=7)
                ax.set_ylabel("Frekans", fontsize=7)
                ax.tick_params(axis='x', labelsize=6)
                ax.tick_params(axis='y', labelsize=6)
                ax.legend(loc='lower right', fontsize=6)
                plots_generated += 1
                current_plot_index += 1
        elif self.state.chart_type == "Dağılım Grafiği":
            # Dağılım grafiği için en az iki sayısal sütuna ihtiyaç vardır (x ve y eksenleri)
            # Şimdilik, seçilen ilk iki sütunu X ve Y olarak kullanalım.
            if len(self.state.selected_columns) < 2:
                QMessageBox.warning(self, "Hata", "Dağılım grafiği için en az iki sayısal sütun seçmelisiniz.")
                return

            if self.state.product_column:
                for (date, product), group in grouped_data:
                    if plots_generated >= self.state.chart_count:
                        break
                    ax = fig.add_subplot(num_rows, num_cols, current_plot_index)
                    x_col = self.state.selected_columns[0]
                    y_col = self.state.selected_columns[1]
                    ax.scatter(group[x_col], group[y_col], label=f'{x_col} vs {y_col}')
                    ax.set_title(f"Dağılım: {date.date()} - {product}", fontsize=8)
                    ax.set_xlabel(x_col, fontsize=7)
                    ax.set_ylabel(y_col, fontsize=7)
                    ax.tick_params(axis='x', labelsize=6)
                    ax.tick_params(axis='y', labelsize=6)
                    ax.legend(loc='lower right', fontsize=6)
                    plots_generated += 1
                    current_plot_index += 1
            else:
                ax = fig.add_subplot(num_rows, num_cols, current_plot_index)
                x_col = self.state.selected_columns[0]
                y_col = self.state.selected_columns[1]
                ax.scatter(df[x_col], df[y_col], label=f'{x_col} vs {y_col}')
                ax.set_title("Dağılım Grafiği", fontsize=8)
                ax.set_xlabel(x_col, fontsize=7)
                ax.set_ylabel(y_col, fontsize=7)
                ax.tick_params(axis='x', labelsize=6)
                ax.tick_params(axis='y', labelsize=6)
                ax.legend(loc='lower right', fontsize=6)
                plots_generated += 1
                current_plot_index += 1
        else:
            QMessageBox.information(self, "Bilgi",
                                    f"'{self.state.chart_type}' grafik türü henüz desteklenmiyor veya uygulanmadı.")

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
        """)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(1000, 750)  # Pencere boyutunu biraz daha büyüttük
    window.show()
    sys.exit(app.exec_())