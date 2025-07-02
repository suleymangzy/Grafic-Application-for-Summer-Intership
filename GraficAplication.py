import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QFileDialog, QLabel, QComboBox, QListWidget, QMessageBox,
    QStackedWidget, QCheckBox, QScrollArea, QSizePolicy
)
from PyQt5.QtCore import Qt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages # PDF'e kaydetmek için
import numpy as np # NaN kontrolü için

# Matplotlib için Türkçe karakter desteği
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'DejaVu Sans', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False # Negatif işaretlerini düzgün göster

# --- Yardımcı Fonksiyonlar ---
# Excel sütun harfini sayısal indekse çevirir (0-tabanlı)
def excel_col_to_index(col_letter):
    index = 0
    for i, char in enumerate(reversed(col_letter.upper())):
        index += (ord(char) - ord('A') + 1) * (26 ** i)
    return index - 1 # 0-tabanlı yapmak için -1

# Sayısal indeksi Excel sütun harfine çevirir (debugging veya bilgi amaçlı)
def index_to_excel_col(index):
    result = ""
    if index < 0:
        return ""
    while index >= 0:
        result = chr(index % 26 + ord('A')) + result
        index = index // 26 - 1
    return result

# --- 1. Dosya Seçimi Sayfası ---
class FileSelectionPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignCenter)

        self.title_label = QLabel("<h2>Dosya Seçimi</h2>")
        self.title_label.setObjectName("title_label") # CSS için ID ataması
        self.title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.title_label)

        self.select_file_button = QPushButton("Excel Dosyası Seç")
        self.select_file_button.clicked.connect(self.select_excel_file)
        layout.addWidget(self.select_file_button)

        self.file_path_label = QLabel("Seçilen Dosya: Yok")
        self.file_path_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.file_path_label)

        self.sheet_selection_label = QLabel("İşlem Yapılacak Sayfayı Seçin:")
        self.sheet_selection_label.setAlignment(Qt.AlignCenter)
        self.sheet_selection_label.hide()
        layout.addWidget(self.sheet_selection_label)

        self.sheet_combo_box = QComboBox()
        self.sheet_combo_box.hide()
        self.sheet_combo_box.currentIndexChanged.connect(self.on_sheet_selected)
        layout.addWidget(self.sheet_combo_box)

        self.next_button = QPushButton("İleri")
        self.next_button.clicked.connect(self.main_window.go_to_data_selection)
        self.next_button.setEnabled(False)
        layout.addWidget(self.next_button)

        layout.addStretch() # İçeriği dikeyde ortalamak için
        self.setLayout(layout)

    def select_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Excel Dosyası Seç", "", "Excel Dosyaları (*.xlsx)")

        if file_path:
            self.main_window.file_path = file_path
            self.file_path_label.setText(f"Seçilen Dosya: <b>{file_path.split('/')[-1]}</b>")
            self.check_excel_sheets(file_path)
        else:
            self.file_path_label.setText("Seçilen Dosya: Yok")
            self.sheet_selection_label.hide()
            self.sheet_combo_box.clear()
            self.sheet_combo_box.hide()
            self.next_button.setEnabled(False)
            self.main_window.selected_sheet = None

    def check_excel_sheets(self, file_path):
        try:
            xls = pd.ExcelFile(file_path)
            all_sheets = xls.sheet_names
            valid_sheets = []

            for sheet_name in ["SMD-OEE", "ROBOT", "DALGA_LEHİM"]:
                if sheet_name in all_sheets:
                    valid_sheets.append(sheet_name)

            self.sheet_combo_box.clear()

            if not valid_sheets:
                QMessageBox.warning(self, "Uyarı", "Üzerinde işlem yapılmaya uygun değildir. Dosya 'SMD-OEE', 'ROBOT' veya 'DALGA_LEHİM' sayfalarından hiçbirini içermiyor.")
                self.sheet_selection_label.hide()
                self.sheet_combo_box.hide()
                self.next_button.setEnabled(False)
                self.main_window.selected_sheet = None
            elif len(valid_sheets) == 1:
                self.main_window.selected_sheet = valid_sheets[0]
                self.sheet_selection_label.setText(f"İşlem Yapılacak Sayfa: <b>{self.main_window.selected_sheet}</b>")
                self.sheet_selection_label.show()
                self.sheet_combo_box.hide()
                self.next_button.setEnabled(True)
            else:
                self.sheet_selection_label.setText("İşlem Yapılacak Sayfayı Seçin:")
                self.sheet_selection_label.show()
                self.sheet_combo_box.addItems(valid_sheets)
                self.sheet_combo_box.show()
                # Varsayılan olarak ilkini seçili yap ve main_window'a ata
                self.main_window.selected_sheet = self.sheet_combo_box.currentText()
                self.next_button.setEnabled(True)

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Dosya okunurken bir hata oluştu: {e}")
            self.sheet_selection_label.hide()
            self.sheet_combo_box.clear()
            self.sheet_combo_box.hide()
            self.next_button.setEnabled(False)
            self.main_window.selected_sheet = None

    def on_sheet_selected(self):
        self.main_window.selected_sheet = self.sheet_combo_box.currentText()
        self.next_button.setEnabled(True)

# --- 2. Veri Seçimi Sayfası ---
class DataSelectionPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setAlignment(Qt.AlignTop)

        self.title_label = QLabel("<h2>Veri Seçimi</h2>")
        self.title_label.setObjectName("title_label")
        self.title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.title_label)

        # Gruplama Değişkeni
        grouping_group = QHBoxLayout()
        grouping_group.addWidget(QLabel("<b>Gruplama Değişkeni:</b>"))
        self.grouping_variable_combo = QComboBox()
        self.grouping_variable_combo.currentIndexChanged.connect(self.populate_grouped_variables)
        grouping_group.addWidget(self.grouping_variable_combo)
        main_layout.addLayout(grouping_group)

        # Gruplanan Değişkenler
        grouped_group = QHBoxLayout()
        grouped_group.addWidget(QLabel("<b>Gruplanan Değişkenler:</b>"))
        self.grouped_variable_list = QListWidget()
        self.grouped_variable_list.setSelectionMode(QListWidget.NoSelection) # Seçilemez yap
        grouped_group.addWidget(self.grouped_variable_list)
        main_layout.addLayout(grouped_group)

        # Metrikler
        metrics_group = QVBoxLayout()
        metrics_group.addWidget(QLabel("<b>Metrikler:</b>"))
        self.metrics_scroll_area = QScrollArea()
        self.metrics_scroll_area.setWidgetResizable(True)
        self.metrics_content_widget = QWidget()
        self.metrics_layout = QVBoxLayout(self.metrics_content_widget)
        self.metrics_scroll_area.setWidget(self.metrics_content_widget)
        metrics_group.addWidget(self.metrics_scroll_area)
        main_layout.addLayout(metrics_group)

        # Butonlar
        button_layout = QHBoxLayout()
        self.back_button = QPushButton("Geri")
        self.back_button.clicked.connect(lambda: self.main_window.go_to_previous_page(self))
        button_layout.addWidget(self.back_button)

        self.next_button = QPushButton("İleri")
        self.next_button.clicked.connect(self.main_window.go_to_graphs_page)
        self.next_button.setEnabled(False)
        button_layout.addWidget(self.next_button)
        main_layout.addLayout(button_layout)

        self.setLayout(main_layout)

    def load_data(self):
        if self.main_window.file_path and self.main_window.selected_sheet:
            try:
                # Veriyi okurken ilk satırı sütun başlığı olarak kabul et
                self.main_window.df = pd.read_excel(self.main_window.file_path, sheet_name=self.main_window.selected_sheet, header=0)
                
                # Sütun isimlerini stringe çevirip boşlukları temizleyelim ve büyük harfe dönüştürelim
                # Bu, Excel'deki başlıkların temizlenmesi ve tutarlılık sağlanması için önemlidir.
                self.main_window.df.columns = self.main_window.df.columns.astype(str).str.strip().str.upper()

                # Excel'in A ve B sütunlarına karşılık gelen başlıkları bulalım
                # A sütunu 0. indekse, B sütunu 1. indekse karşılık gelir
                a_idx = excel_col_to_index('A')
                b_idx = excel_col_to_index('B')

                if len(self.main_window.df.columns) > a_idx and len(self.main_window.df.columns) > b_idx:
                    self.main_window.grouping_variable_col = self.main_window.df.columns[a_idx]
                    self.main_window.grouped_variable_col = self.main_window.df.columns[b_idx]
                else:
                    raise ValueError(f"Excel dosyasında 'A' ({a_idx+1}. sütun) veya 'B' ({b_idx+1}. sütun) sütunlarına karşılık gelen yeterli sayıda sütun bulunamadı. Lütfen dosyanızı kontrol edin.")

                self.main_window.metrics_columns = self.get_metrics_columns()
                self.populate_grouping_variable()
                self.populate_metrics()
                self.next_button.setEnabled(True if self.main_window.selected_metrics else False)

            except Exception as e:
                QMessageBox.critical(self, "Hata", f"Veri okunurken veya işlenirken bir hata oluştu: {e}\nLütfen Excel dosyanızın formatını ve sütun başlıklarını kontrol edin.")
                self.main_window.df = None
                self.main_window.grouping_variable_values = []
                self.main_window.grouped_variable_dict = {}
                self.main_window.metrics_columns = []
                self.main_window.selected_metrics = []
                self.grouping_variable_combo.clear()
                self.grouped_variable_list.clear()
                self.clear_metrics_checkboxes()
                self.next_button.setEnabled(False)

    def populate_grouping_variable(self):
        self.grouping_variable_combo.clear()
        if self.main_window.df is not None and self.main_window.grouping_variable_col:
            grouping_data = self.main_window.df[self.main_window.grouping_variable_col].dropna().unique().tolist()
            self.main_window.grouping_variable_values = [str(x) for x in grouping_data if str(x).strip()]
            self.grouping_variable_combo.addItems(self.main_window.grouping_variable_values)
            self.populate_grouped_variables()

    def populate_grouped_variables(self):
        self.grouped_variable_list.clear()
        selected_grouping_value = self.grouping_variable_combo.currentText()
        if self.main_window.df is not None and selected_grouping_value and self.main_window.grouped_variable_col:
            filtered_df = self.main_window.df[self.main_window.df[self.main_window.grouping_variable_col] == selected_grouping_value]
            grouped_data = filtered_df[self.main_window.grouped_variable_col].dropna().unique().tolist()
            self.main_window.grouped_variable_dict = {selected_grouping_value: [str(x) for x in grouped_data if str(x).strip()]}
            self.grouped_variable_list.addItems(self.main_window.grouped_variable_dict[selected_grouping_value])

    def get_metrics_columns(self):
        column_names = self.main_window.df.columns.tolist()
        print(f"Pandas tarafından okunan tüm sütun başlıkları: {column_names}")

        # H indexi (0-tabanlı) = excel_col_to_index('H')
        # BD indexi (0-tabanlı) = excel_col_to_index('BD')
        h_col_numeric_index = excel_col_to_index('H')
        bd_col_numeric_index = excel_col_to_index('BD')

        if h_col_numeric_index >= len(column_names):
            QMessageBox.warning(self, "Uyarı", f"Excel dosyasında 'H' ({h_col_numeric_index+1}. sütun) sütununa karşılık gelen sütun bulunamadı. Lütfen dosyanızın sütun genişliğini kontrol edin.")
            return []
        
        if bd_col_numeric_index >= len(column_names):
             QMessageBox.warning(self, "Uyarı", f"Excel dosyasında 'BD' ({bd_col_numeric_index+1}. sütun) sütununa karşılık gelen sütun bulunamadı. Lütfen dosyanızın sütun genişliğini kontrol edin.")
             return []

        if h_col_numeric_index > bd_col_numeric_index:
             QMessageBox.warning(self, "Uyarı", f"'H' ({index_to_excel_col(h_col_numeric_index)}) sütunu 'BD' ({index_to_excel_col(bd_col_numeric_index)}) sütunundan sonra geliyor. Metrik aralığı geçersiz.")
             return []

        # Excel'deki H sütunundan BD sütununa kadar olan başlıkları al
        metrics_potential_cols = column_names[h_col_numeric_index : bd_col_numeric_index + 1]
        print(f"Excel H'den BD'ye kadar olan potansiyel metrik sütun başlıkları: {metrics_potential_cols}")

        valid_metrics = []
        for col_name in metrics_potential_cols:
            # Sütundaki değerlerin tümüyle boş veya NaN olup olmadığını kontrol et
            # Süre değerleri genelde string veya timedelta olarak okunur, np.nan değil.
            # Dolayısıyla, boş string veya pandas'ın NaT (Not a Time) kontrolü önemlidir.
            
            # str'ye çevirip boşlukları temizleyerek boşluğa eşit mi bak
            is_all_empty_string = self.main_window.df[col_name].astype(str).str.strip().eq('').all()
            # Ya da tüm değerler pandas'ın tanıdığı NaN mı (boş hücreler NaN olarak okunur)
            is_all_nan = self.main_window.df[col_name].isnull().all()

            if not (is_all_empty_string or is_all_nan):
                valid_metrics.append(col_name)
            else:
                print(f"'{col_name}' sütunu tamamen boş veya sadece boşluk içeren değerler içeriyor, metrik olarak dahil edilmiyor.")

        print(f"Geçerli metrikler: {valid_metrics}")
        return valid_metrics

    def populate_metrics(self):
        self.clear_metrics_checkboxes()
        self.main_window.selected_metrics = []

        if not self.main_window.metrics_columns:
            empty_label = QLabel("Seçilebilir metrik bulunamadı.", parent=self.metrics_content_widget)
            empty_label.setAlignment(Qt.AlignCenter)
            self.metrics_layout.addWidget(empty_label)
            self.next_button.setEnabled(False)
            return

        for metric in self.main_window.metrics_columns:
            checkbox = QCheckBox(metric)
            checkbox.setChecked(True)
            checkbox.stateChanged.connect(self.on_metric_checkbox_changed)
            self.metrics_layout.addWidget(checkbox)
            self.main_window.selected_metrics.append(metric)

        self.next_button.setEnabled(True)

    def clear_metrics_checkboxes(self):
        while self.metrics_layout.count():
            item = self.metrics_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def on_metric_checkbox_changed(self, state):
        sender_checkbox = self.sender()
        metric_name = sender_checkbox.text()

        if state == Qt.Checked:
            if metric_name not in self.main_window.selected_metrics:
                self.main_window.selected_metrics.append(metric_name)
        else:
            if metric_name in self.main_window.selected_metrics:
                self.main_window.selected_metrics.remove(metric_name)
        
        self.next_button.setEnabled(bool(self.main_window.selected_metrics))


# --- 3. Grafikler Sayfası ---
class GraphsPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.init_ui()
        self.figures = []

    def init_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setAlignment(Qt.AlignTop)

        self.title_label = QLabel("<h2>Grafikler</h2>")
        self.title_label.setObjectName("title_label")
        self.title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.title_label)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.graphs_container = QWidget()
        self.graphs_layout = QVBoxLayout(self.graphs_container)
        self.graphs_container.setLayout(self.graphs_layout)
        self.scroll_area.setWidget(self.graphs_container)
        main_layout.addWidget(self.scroll_area)

        # Butonlar
        button_layout = QHBoxLayout()
        self.back_button = QPushButton("Geri")
        self.back_button.clicked.connect(lambda: self.main_window.go_to_previous_page(self))
        button_layout.addWidget(self.back_button)

        self.save_graphs_button = QPushButton("Grafikleri Kaydet/İndir (PDF)")
        self.save_graphs_button.clicked.connect(self.save_all_graphs_to_pdf)
        button_layout.addWidget(self.save_graphs_button)
        main_layout.addLayout(button_layout)

        self.setLayout(main_layout)

    def clear_graphs(self):
        for fig in self.figures:
            plt.close(fig)
        self.figures = []
        while self.graphs_layout.count():
            item = self.graphs_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def generate_graphs(self):
        self.clear_graphs()

        df = self.main_window.df
        selected_grouping_value = self.main_window.grouping_variable_combo.currentText()
        selected_metrics = self.main_window.selected_metrics

        if df is None or not selected_grouping_value or not selected_metrics:
            QMessageBox.warning(self, "Uyarı", "Grafik oluşturmak için yeterli veri veya seçim bulunmuyor.")
            return

        filtered_df_by_grouping = df[df[self.main_window.grouping_variable_col] == selected_grouping_value].copy()
        
        grouped_variables = filtered_df_by_grouping[self.main_window.grouped_variable_col].dropna().unique().tolist()
        grouped_variables = [str(x) for x in grouped_variables if str(x).strip()]

        if not grouped_variables:
            QMessageBox.information(self, "Bilgi", f"'{selected_grouping_value}' için gruplanmış değişken bulunamadı.")
            return

        # Tek tip bir renk paleti kullanarak tutarlılık sağlayın
        # Matplotlib'in 20 farklı renk tonu olan 'tab20' paletini kullanıyoruz.
        # Daha fazla metrik için renklerin tekrar etmesi normaldir.
        colors_palette = plt.cm.get_cmap('tab20', len(selected_metrics))
        metric_colors = {metric: colors_palette(i) for i, metric in enumerate(selected_metrics)}

        # BP sütun başlığını tanımlayın (varsayım: 'BP' adında bir sütun var)
        # Eğer BP sütununuzun adı farklıysa, bu satırı güncelleyin.
        bp_column_name = 'BP'

        for grouped_var in grouped_variables:
            subset_df = filtered_df_by_grouping[filtered_df_by_grouping[self.main_window.grouped_variable_col] == grouped_var]

            plot_data = {}
            for metric in selected_metrics:
                # Metrik değerlerini al ve süre formatından saniyeye çevirerek topla
                # NaN olmayan ve boş string olmayan değerleri seç
                time_values = subset_df[metric].dropna()

                # 'total_seconds' fonksiyonunu çağırabilmek için timedelta objesine çevirmeliyiz.
                # Eğer Excel'deki hücreler zaten bir datetime.time objesi olarak okunuyorsa (ki genelde öyle olur),
                # doğrudan timedelta'a çevirebiliriz. string'e çevirmek ekstra bir adım olabilir.
                
                # try-except bloğu, farklı formatlardaki süreleri işlemek için daha sağlamdır.
                converted_to_timedelta = pd.to_timedelta(time_values.astype(str), errors='coerce')
                
                # Sadece geçerli timedelta değerlerini (NaT olmayanları) al
                valid_timedeltas = converted_to_timedelta.dropna()
                
                total_seconds = 0
                if not valid_timedeltas.empty:
                    # Tüm timedelta değerlerini saniyeye çevirip topla
                    total_seconds = valid_timedeltas.apply(lambda x: x.total_seconds()).sum()
                
                if total_seconds > 0:
                    plot_data[metric] = total_seconds
                
            if not plot_data:
                continue

            fig = Figure(figsize=(8.27, 8.27)) # A4 genişliğine yakın bir kare figür
            canvas = FigureCanvas(fig)
            self.graphs_layout.addWidget(canvas)
            self.figures.append(fig)

            ax = fig.add_subplot(111)

            # BP değeri alınması
            bp_value = 0
            if bp_column_name in subset_df.columns:
                bp_col_data = subset_df[bp_column_name].dropna()
                
                # BP değerinin formatına göre işlem yapın (süre mi, sayısal mı?)
                # Varsayalım ki BP de süre formatında ve saniyeye çeviriyoruz:
                converted_bp_times = pd.to_timedelta(bp_col_data.astype(str), errors='coerce').dropna()
                if not converted_bp_times.empty:
                    bp_value = converted_bp_times.apply(lambda x: x.total_seconds()).sum()
                
                # Eğer BP sayısal bir değerse (süre değil, doğrudan sayı):
                # bp_value = pd.to_numeric(bp_col_data, errors='coerce').sum()

            # Pasta grafiği oluştur
            wedges, texts, autotexts = ax.pie(
                plot_data.values(),
                labels=plot_data.keys(),
                autopct=lambda p: f'{p:.1f}%\n({bp_value:.0f})' if p > 0 else '', # BP değeri ve yüzdesi
                startangle=90,
                colors=[metric_colors[m] for m in plot_data.keys()] # Renkleri metrik bazında ata
            )

            ax.set_title(f'Gruplama: {selected_grouping_value}\nGruplanan: {grouped_var}', fontsize=12)
            ax.axis('equal') # Oranların eşit olmasını sağlar (daire şeklinde görünüm)

            # Yüzde ve BP değeri etiketlerinin stilini ayarla
            for autotext in autotexts:
                autotext.set_color('black')
                autotext.set_fontsize(8)

            # Label etiketlerinin stilini ayarla
            for text in texts:
                text.set_fontsize(8)

            fig.tight_layout() # Grafiğin sıkıca yerleşmesini sağlar

    def save_all_graphs_to_pdf(self):
        if not self.figures:
            QMessageBox.warning(self, "Uyarı", "Kaydedilecek grafik bulunmuyor.")
            return

        file_name, _ = QFileDialog.getSaveFileName(self, "Grafikleri PDF Olarak Kaydet", "", "PDF Dosyaları (*.pdf)")

        if file_name:
            try:
                with PdfPages(file_name) as pdf:
                    for fig in self.figures:
                        # bbox_inches='tight' ve pad_inches ile grafik etrafındaki boşlukları ayarla
                        pdf.savefig(fig, bbox_inches='tight', pad_inches=0.5)

                    # Tüm grafiklerin açıklamasını kapsayacak şekilde genel renk legend'ı
                    if self.main_window.selected_metrics:
                        legend_fig = Figure(figsize=(8.27, 11.69)) # A4 boyutunda boş bir sayfa
                        legend_ax = legend_fig.add_subplot(111)
                        legend_ax.set_axis_off() # Eksenleri gizle

                        handles = []
                        labels = []
                        colors_palette = plt.cm.get_cmap('tab20', len(self.main_window.selected_metrics))
                        
                        for i, metric in enumerate(self.main_window.selected_metrics):
                            color = colors_palette(i)
                            handle = plt.Rectangle((0, 0), 1, 1, fc=color, edgecolor='black')
                            handles.append(handle)
                            labels.append(metric)
                        
                        # Legend'ı A4 sayfasının sağ alt köşesine yerleştir
                        legend_ax.legend(handles, labels, title="Metrik Legendı", 
                                         loc='lower right', bbox_to_anchor=(0.95, 0.05),
                                         fontsize=9, title_fontsize=10, frameon=True, fancybox=True, shadow=True, ncol=2)

                        pdf.savefig(legend_fig, bbox_inches='tight', pad_inches=0.5)
                        plt.close(legend_fig) # Oluşturulan legend figürünü kapat

                QMessageBox.information(self, "Başarılı", f"Grafikler '{file_name}' konumuna kaydedildi.")
            except Exception as e:
                QMessageBox.critical(self, "Hata", f"Grafikler kaydedilirken bir hata oluştu: {e}")

# --- Ana Uygulama Penceresi ---
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Veri Analiz Uygulaması")
        self.setGeometry(100, 100, 1200, 900)

        # Uygulama genelinde kullanılacak veriler
        self.file_path = None
        self.selected_sheet = None
        self.df = None
        self.grouping_variable_col = None # Pandas DataFrame'deki sütun adı (örn: 'SET-UP')
        self.grouped_variable_col = None  # Pandas DataFrame'deki sütun adı (örn: 'MALZEME BEKLEME')
        self.grouping_variable_values = []
        self.grouped_variable_dict = {}
        self.metrics_columns = []         # Pandas DataFrame'deki metrik sütun adları
        self.selected_metrics = []        # Kullanıcının seçtiği metrik sütun adları

        self.init_ui()

    def init_ui(self):
        self.stacked_widget = QStackedWidget(self)

        self.file_selection_page = FileSelectionPage(self)
        self.data_selection_page = DataSelectionPage(self)
        self.graphs_page = GraphsPage(self)

        self.stacked_widget.addWidget(self.file_selection_page)
        self.stacked_widget.addWidget(self.data_selection_page)
        self.stacked_widget.addWidget(self.graphs_page)

        main_layout = QVBoxLayout()
        main_layout.addWidget(self.stacked_widget)
        self.setLayout(main_layout)

        self.stacked_widget.setCurrentWidget(self.file_selection_page)

    def go_to_data_selection(self):
        if self.selected_sheet:
            self.stacked_widget.setCurrentWidget(self.data_selection_page)
            self.data_selection_page.load_data()
        else:
            QMessageBox.warning(self, "Uyarı", "Lütfen işlem yapmak için bir Excel sayfası seçin.")

    def go_to_graphs_page(self):
        if self.selected_metrics:
            self.stacked_widget.setCurrentWidget(self.graphs_page)
            self.graphs_page.generate_graphs()
        else:
            QMessageBox.warning(self, "Uyarı", "Lütfen grafik oluşturmak için en az bir metrik seçin.")

    def go_to_previous_page(self, current_page):
        current_index = self.stacked_widget.indexOf(current_page)
        if current_index > 0:
            self.stacked_widget.setCurrentIndex(current_index - 1)
            # Geri düğmesine basıldığında sayfa durumunu kontrol edebiliriz
            if current_page == self.data_selection_page:
                # Dosya seçimi sayfasına geri dönüldüğünde mevcut dosya bilgisini tekrar göster
                if self.file_path:
                    self.file_selection_page.file_path_label.setText(f"Seçilen Dosya: <b>{self.file_path.split('/')[-1]}</b>")
                    # Sayfa seçimini de doğru yansıtmak gerekebilir.
                    # Eğer tek bir geçerli sayfa varsa direkt gösterilir, yoksa combo box tekrar doldurulur.
                    self.file_selection_page.check_excel_sheets(self.file_path)


# --- Uygulama Başlatma ---
if __name__ == "__main__":
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
        QComboBox, QListWidget, QScrollArea {
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
    """)

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())