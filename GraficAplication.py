import sys
import os
import numpy as np
import pandas as pd
from scipy import stats
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QAction, QMenu, QInputDialog,
    QMessageBox, QFileDialog, QLabel, QVBoxLayout, QWidget,
    QScrollArea, QProgressDialog
)
from PyQt5.QtGui import QIcon, QFont, QPixmap
from PyQt5.QtCore import Qt, QDateTime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
from fpdf import FPDF
from fpdf.enums import XPos, YPos

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


class MplWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.figure = plt.figure()
        self.canvas = FigureCanvas(self.figure)
        self.toolbar = NavigationToolbar(self.canvas, self)
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.toolbar)
        self.layout.addWidget(self.canvas)
        self.setLayout(self.layout)
        self.figure.clear()
        self.canvas.draw_idle()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Grafik Uygulaması")
        self.setGeometry(100, 100, 1024, 768)
        self.current_plot_data = []

        # Veri yükleme ve seçimi için değişkenler
        self.loaded_data_df = None  # Pandas DataFrame olarak tutalım
        self.column_names = []
        self.selected_x_column = None
        self.selected_y_column = None
        self.selected_category_column = None
        self.selected_color_column = None
        self.current_chart_type = ""
        self.current_chart_count = 0

        # UI bileşenlerini oluştur
        self.init_ui()
        self.create_menu()
        self.update_plot_info_label()  # Başlangıçta boş

    def init_ui(self):
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        # Dosya bilgi etiketi
        self.file_info_label = QLabel("Lütfen bir veri dosyası seçin...", self)
        self.file_info_label.setAlignment(Qt.AlignCenter)
        self.file_info_label.setFont(QFont("Arial", 12))
        self.file_info_label.setObjectName("fileInfoLabel")
        self.file_info_label.setFixedHeight(30)

        # Grafik bilgi etiketi
        self.plot_info_label = QLabel("Henüz bir grafik oluşturulmadı.", self)
        self.plot_info_label.setAlignment(Qt.AlignCenter)
        self.plot_info_label.setFont(QFont("Arial", 12))
        self.plot_info_label.setObjectName("plotInfoLabel")
        self.plot_info_label.setFixedHeight(30)

        self.main_layout.addWidget(self.file_info_label)
        self.main_layout.addWidget(self.plot_info_label)

        # Grafik alanı
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_content_widget = QWidget()
        self.scroll_content_layout = QVBoxLayout(self.scroll_content_widget)
        self.matplotlib_widget = MplWidget(self)
        self.scroll_content_layout.addWidget(self.matplotlib_widget)
        self.scroll_area.setWidget(self.scroll_content_widget)
        self.main_layout.addWidget(self.scroll_area)

        # Stil ayarları (mevcut haliyle bırakıldı, sorunsuz görünüyor)
        self.setStyleSheet("""
            QMainWindow { background-color: #202020; }
            QMenuBar {
                background-color: #333; color: #FFF;
                font-size: 14px; padding: 5px 0px;
            }
            QMenuBar::item { padding: 8px 15px; background-color: transparent; }
            QMenuBar::item:selected { background-color: #555; }
            QMenu {
                background-color: #444; color: #FFF;
                border: 1px solid #666;
            }
            QMenu::item { padding: 8px 25px; background-color: transparent; }
            QMenu::item:selected { background-color: #666; }
            QMenu::separator {
                height: 1px; background-color: #555;
                margin-left: 10px; margin-right: 10px;
            }
            QLabel#fileInfoLabel {
                color: #ADD8E6; font-weight: bold;
                padding: 5px; background-color: #282828;
                border-bottom: 1px solid #444;
            }
            QLabel#plotInfoLabel {
                color: #A0FFA0; font-weight: bold;
                padding: 5px; background-color: #3A3A3A;
                border-bottom: 1px solid #555;
            }
            QScrollArea { border: none; background-color: #282828; }
            QScrollBar:vertical {
                border: 1px solid #555; background: #333;
                width: 15px; margin: 15px 0 15px 0;
            }
            QScrollBar::handle:vertical {
                background: #666; min-height: 20px; border-radius: 5px;
            }
            QScrollBar:horizontal {
                border: 1px solid #555; background: #333;
                height: 15px; margin: 0 15px 0 15px;
            }
            QScrollBar::handle:horizontal {
                background: #666; min-width: 20px; border-radius: 5px;
            }
        """)

    def create_menu(self):
        menubar = self.menuBar()

        # Yardımcı fonksiyon: Action oluşturma
        def create_action(text, icon_path, shortcut, tip, func):
            # İkon yoksa veya yüklenemezse metin tabanlı action oluştur
            icon = QIcon(icon_path) if os.path.exists(icon_path) else QIcon()
            action = QAction(icon, text, self)
            action.setShortcut(shortcut)
            action.setStatusTip(tip)
            action.triggered.connect(func)
            return action

        # Dosya Menüsü
        file_menu = menubar.addMenu("&Dosya")
        file_menu.addAction(create_action(
            "Veri Dosyası &Aç...", 'icons/data_file.png', "Ctrl+D",
            "Bir veri dosyasını (Excel, CSV) açar ve verilerini yükler",
            lambda: self.open_file("Excel Dosyaları (*.xlsx *.xls);;CSV Dosyaları (*.csv);;Tüm Dosyalar (*)", "data")
        ))
        file_menu.addSeparator()
        file_menu.addAction(create_action(
            "Çı&kış", 'icons/exit.png', "Ctrl+Q",
            "Uygulamadan çıkar", self.close
        ))

        # Grafik Oluştur Menüsü
        self.plot_create_menu = menubar.addMenu("&Grafik Oluştur")
        self.chart_types = [
            "Çizgi Grafiği (plot)", "Bar Grafiği (bar)", "Histogram (hist)",
            "Pasta Grafiği (pie)", "Dağılım Grafiği (scatter)", "Alan Grafiği (fill_between)",
            "Kutu Grafiği (boxplot)", "Violin Grafiği (violinplot)", "Stem Grafiği (stem)",
            "Hata Çubuklu Grafik (errorbar)"
        ]
        for chart_type in self.chart_types:
            action = QAction(chart_type + "...", self)
            action.setStatusTip(f"'{chart_type}' grafiği oluştur")
            action.triggered.connect(lambda checked, name=chart_type: self.get_plot_count(name))
            self.plot_create_menu.addAction(action)
        self.plot_create_menu.setEnabled(False) # Başlangıçta pasif

        # İndir/Yazdır Menüsü
        download_print_menu = menubar.addMenu("&İndir / Yazdır")
        save_as_menu = QMenu("Farklı Kaydet", self)
        formats = {"PNG G&örseli": "png", "JPEG G&örseli": "jpeg",
                   "PDF &Belgesi": "pdf", "SVG &Vektörü": "svg"}
        for name, ext in formats.items():
            save_action = create_action(
                name, f'icons/save_{ext}.png', "",
                f"Grafiği .{ext} formatında kaydet",
                lambda checked, fmt=ext: self.save_graph(fmt)
            )
            save_as_menu.addAction(save_action)
        download_print_menu.addMenu(save_as_menu)
        download_print_menu.addSeparator()
        download_print_menu.addAction(create_action(
            "&Rapor Oluştur (PDF)", 'icons/report_pdf.png', "Ctrl+R",
            "Mevcut grafiğin raporunu PDF olarak oluştur", self.generate_pdf_report
        ))
        download_print_menu.addAction(create_action(
            "&Yazdır...", 'icons/print.png', "Ctrl+P",
            "Mevcut grafiği yazdır", self.print_graph
        ))
        download_print_menu.setEnabled(False) # Başlangıçta pasif
        self.download_print_menu = download_print_menu # Referansını sakla

        # Veri Seç Menüsü
        self.data_menu = menubar.addMenu("&Veri Seç")
        self.x_axis_menu = QMenu("X Ekseni &Seç", self)
        self.y_axis_menu = QMenu("Y Ekseni &Seç", self)
        self.category_selection_menu = QMenu("Kategori Verisi &Seç", self)
        self.color_by_data_menu = QMenu("Veriye Göre &Renklendir", self)

        self.data_menu.addMenu(self.x_axis_menu)
        self.data_menu.addMenu(self.y_axis_menu)
        self.data_menu.addSeparator()
        self.data_menu.addMenu(self.category_selection_menu)
        self.data_menu.addMenu(self.color_by_data_menu)
        self.data_menu.setEnabled(False)  # Başlangıçta pasif

    def populate_data_selection_menus(self):
        """Yüklenen verilere göre veri seçim menülerini doldurur."""
        self.x_axis_menu.clear()
        self.y_axis_menu.clear()
        self.category_selection_menu.clear()
        self.color_by_data_menu.clear()

        if self.loaded_data_df is None or self.loaded_data_df.empty:
            self.data_menu.setEnabled(False)
            return

        self.data_menu.setEnabled(True)  # Veri varsa menüyü etkinleştir

        # "Hiçbiri" seçeneği ekle
        none_action_x = QAction("Hiçbiri", self)
        none_action_x.triggered.connect(lambda: self.set_selected_column("X", None))
        self.x_axis_menu.addAction(none_action_x)

        none_action_y = QAction("Hiçbiri", self)
        none_action_y.triggered.connect(lambda: self.set_selected_column("Y", None))
        self.y_axis_menu.addAction(none_action_y)

        none_action_cat = QAction("Hiçbiri", self)
        none_action_cat.triggered.connect(lambda: self.set_selected_column("Category", None))
        self.category_selection_menu.addAction(none_action_cat)

        none_action_color = QAction("Hiçbiri", self)
        none_action_color.triggered.connect(lambda: self.set_selected_column("Color", None))
        self.color_by_data_menu.addAction(none_action_color)

        # Sütunları menülere ekle
        for col in self.column_names:
            # Tüm sütunları X ve Y eksenlerine, kategori ve renk seçim menülerine ekle
            # Tarih/zaman veya string sütunların da seçilebilmesi için Type kontrolü kaldırıldı
            self.x_axis_menu.addAction(self.create_column_selection_action("X", col))
            self.y_axis_menu.addAction(self.create_column_selection_action("Y", col))
            self.category_selection_menu.addAction(self.create_column_selection_action("Category", col))
            self.color_by_data_menu.addAction(self.create_column_selection_action("Color", col))

    def create_column_selection_action(self, axis_type, column_name):
        action = QAction(column_name, self)
        action.triggered.connect(lambda _, t=axis_type, n=column_name: self.set_selected_column(t, n))
        return action

    def set_selected_column(self, axis_type, column_name):
        """Seçilen sütunları günceller ve bilgi etiketini yeniler."""
        if axis_type == "X":
            self.selected_x_column = column_name
        elif axis_type == "Y":
            self.selected_y_column = column_name
        elif axis_type == "Category":
            self.selected_category_column = column_name
        elif axis_type == "Color":
            self.selected_color_column = column_name

        # Bilgi etiketini güncelle
        self.update_plot_info_label()
        # Veri seçimi güncellendiğinde grafiği yeniden çiz (eğer bir grafik türü seçiliyse)
        if self.current_chart_type and self.current_chart_count > 0:
            self.draw_graph(self.current_chart_type, self.current_chart_count)

    def open_file(self, file_filter, file_type_code):
        """Dosya seçim iletişim kutusunu açar ve kullanıcı seçimini işler"""
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Dosya Seç",
            "",  # Başlangıç dizini
            file_filter,
            options=QFileDialog.DontUseNativeDialog
        )

        if not file_name:
            self.reset_data_selection(show_message=False)  # Dosya seçimi iptal edilirse verileri sıfırla
            return

        try:
            ext = os.path.splitext(file_name)[1].lower()
            sheet_name = None
            if ext in ('.xlsx', '.xls'):
                # Excel dosyası ise sayfa adı sor
                sheet_name, ok = QInputDialog.getText(
                    self, "Sayfa Adı", "Lütfen Excel sayfasının adını girin (örneğin: Sayfa1):"
                )
                if not ok or not sheet_name:
                    QMessageBox.warning(self, "Uyarı", "Excel sayfası adı girilmedi. Dosya yükleme iptal edildi.")
                    self.reset_data_selection(show_message=False)
                    return

            self.load_data_from_file(file_name, sheet_name)
            self.update_file_info_label(file_name)
            self.plot_create_menu.setEnabled(True) # Veri yüklendiyse grafik oluştur menüsünü etkinleştir

        except Exception as e:
            QMessageBox.critical(
                self,
                "Yükleme Hatası",
                f"Dosya yüklenirken hata oluştu:\n{str(e)}"
            )
            self.reset_data_selection(show_message=False) # Hata durumunda da sıfırla

    def reset_data_selection(self, show_message=True):
        """Veri seçimlerini sıfırla ve isteğe bağlı mesaj göster"""
        self.update_file_info_label("")
        self.loaded_data_df = None
        self.column_names = []
        self.populate_data_selection_menus()  # Menüleri boşalt ve devre dışı bırak

        # Seçili sütunları sıfırla
        self.selected_x_column = None
        self.selected_y_column = None
        self.selected_category_column = None
        self.selected_color_column = None
        self.current_chart_type = ""
        self.current_chart_count = 0
        self.current_plot_data = []
        self.update_plot_info_label()  # Bilgi etiketini de sıfırla
        self.matplotlib_widget.figure.clear()
        self.matplotlib_widget.canvas.draw_idle()

        self.plot_create_menu.setEnabled(False) # Veri yoksa grafik oluştur menüsünü pasif yap
        self.download_print_menu.setEnabled(False) # Veri yoksa indir/yazdır menüsünü pasif yap


        if show_message:
            QMessageBox.information(
                self,
                "Bilgi",
                "Dosya seçimi iptal edildi veya sıfırlandı."
            )

    def convert_datetime_columns(self, df):
        """DataFrame'deki olası tarih/zaman sütunlarını datetime objelerine dönüştürür."""
        for col in df.columns:
            # Pandas'ın to_datetime fonksiyonunu kullanarak otomatik algılama
            # errors='coerce' geçersiz tarihleri NaT (Not a Time) yapar
            # infer_datetime_format=True formatı otomatik bulmaya çalışır
            converted_series = pd.to_datetime(df[col], errors='coerce', infer_datetime_format=True)

            # Eğer dönüştürme sonucunda orijinalden daha az NaT değeri varsa, dönüştür
            # Yani, orijinal sütunda NaN olanlar hariç, diğer değerlerin çoğu geçerli tarihse
            if converted_series.notna().sum() > (df[col].isna().sum() + (len(df) * 0.5)): # En az %50'si geçerli tarihse
                 df[col] = converted_series
        return df

    def load_data_from_file(self, file_path, sheet_name=None):
        """Dosyadan veri yükler ve arayüzü günceller"""
        try:
            ext = os.path.splitext(file_path)[1].lower()

            # Yaygın boş değerleri NaN olarak okumak için
            na_values = ['', '#N/A', '#N/A N/A', '#NA', '-1.#IND', '-1.#QNAN', '-NaN', '-nan',
                         '1.#IND', '1.#QNAN', '<NA>', 'N/A', 'NA', 'NULL', 'NaN', 'n/a',
                         'nan', 'null', '?', '*', '-', ' ']

            if ext in ('.xlsx', '.xls'):
                df = pd.read_excel(file_path, sheet_name=sheet_name, na_values=na_values)
            elif ext == '.csv':
                df = pd.read_csv(file_path, na_values=na_values)
            else:
                raise ValueError("Desteklenmeyen dosya formatı")

            # Boş (NaN) değerleri 0 ile doldur
            df.fillna(0, inplace=True)

            # Tarih/zaman sütunlarını dönüştürme
            df = self.convert_datetime_columns(df)

            self.loaded_data_df = df
            self.column_names = list(df.columns)

            self.populate_data_selection_menus()
            self.set_default_column_selections()  # Varsayılan sütunları ayarla

            QMessageBox.information(
                self,
                "Başarılı",
                f"Veriler başarıyla yüklendi:\n{os.path.basename(file_path)}\n"
                f"{f'Sayfa: {sheet_name}\n' if sheet_name else ''}"
                f"Toplam {len(df)} satır, {len(df.columns)} sütun yüklendi."
            )

        except Exception as e:
            QMessageBox.critical(
                self,
                "Yükleme Hatası",
                f"Veri yüklenirken hata oluştu:\n{str(e)}"
            )
            self.reset_data_selection(show_message=False)  # Hata durumunda da sıfırla

    def set_default_column_selections(self):
        """Yüklenen DataFrame'e göre varsayılan sütun seçimlerini yapar."""
        if self.loaded_data_df is None or self.loaded_data_df.empty:
            return

        all_cols = self.loaded_data_df.columns.tolist()

        # İlk uygun sütunları varsayılan olarak ata
        self.selected_x_column = all_cols[0] if len(all_cols) >= 1 else None
        self.selected_y_column = all_cols[1] if len(all_cols) >= 2 else None
        self.selected_category_column = all_cols[2] if len(all_cols) >= 3 else None
        self.selected_color_column = all_cols[3] if len(all_cols) >= 4 else None

        self.update_plot_info_label()


    def update_file_info_label(self, file_path):
        if file_path:
            base_name = os.path.basename(file_path)
            # İkon dosyası uzantıdan belirlenir (örn: .xlsx -> excel.png)
            ext = os.path.splitext(file_path)[1].lower()
            icon_name = {
                '.xlsx': 'excel.png', '.xls': 'excel.png',
                '.csv': 'data_file.png'
            }.get(ext, 'data_file.png') # Varsayılan olarak data_file.png

            icon_path = f'icons/{icon_name}'
            icon = QPixmap(icon_path) if os.path.exists(icon_path) else QPixmap()

            if not icon.isNull():
                icon = icon.scaled(24, 24, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                self.file_info_label.setPixmap(icon)
                self.file_info_label.setText(f"  {base_name}")
                self.file_info_label.setContentsMargins(5, 0, 0, 0)
            else:
                self.file_info_label.setPixmap(QPixmap())
                self.file_info_label.setText(f"  {base_name} (İkon yüklenemedi)")
        else:
            self.file_info_label.setPixmap(QPixmap())
            self.file_info_label.setText("Lütfen bir veri dosyası seçin...")
        self.file_info_label.adjustSize()

    def update_plot_info_label(self):
        """Grafik ve seçili veri bilgilerini gösteren etiketi günceller."""
        selected_info_parts = []
        if self.selected_x_column is not None:
            selected_info_parts.append(f"X: {self.selected_x_column}")
        if self.selected_y_column is not None:
            selected_info_parts.append(f"Y: {self.selected_y_column}")
        if self.selected_category_column is not None:
            selected_info_parts.append(f"Kategori: {self.selected_category_column}")
        if self.selected_color_column is not None:
            selected_info_parts.append(f"Renk: {self.selected_color_column}")

        selected_info_str = ", ".join(selected_info_parts) if selected_info_parts else "Yok"

        plot_type_display = self.current_chart_type.split('(')[0].strip() if self.current_chart_type else "Yok"

        if self.current_chart_type and self.current_chart_count > 0:
            new_text = f"Grafik: {plot_type_display} ({self.current_chart_count} adet) | Seçili Veriler: {selected_info_str}"
        else:
            new_text = f"Seçili Veriler: {selected_info_str}"
            if not selected_info_parts: # Hiçbir şey seçili değilse başlangıç metnini göster
                new_text = "Henüz bir grafik oluşturulmadı."

        self.plot_info_label.setText(new_text)
        self.plot_info_label.setContentsMargins(5, 0, 0, 0)
        self.plot_info_label.adjustSize()

    def get_plot_count(self, chart_type):
        if self.loaded_data_df is None:
            QMessageBox.warning(self, "Uyarı", "Lütfen önce bir veri dosyası yükleyin.")
            return

        try:
            num, ok = QInputDialog.getInt(
                self, "Grafik Adedi",
                f"Kaç adet '{chart_type}' grafiği oluşturmak istersiniz?",
                min=1, max=10, value=1
            )
            if ok:
                self.current_chart_type = chart_type
                self.current_chart_count = num
                self.draw_graph(chart_type, num)
                self.download_print_menu.setEnabled(True) # Grafik oluşturulduktan sonra menüyü etkinleştir

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Grafik adedi alınamadı: {e}")

    def draw_graph(self, chart_type_text, count):
        self.matplotlib_widget.figure.clear()
        self.current_plot_data = []

        chart_type = chart_type_text.split('(')[-1][:-1] if '(' in chart_type_text else None
        if not chart_type:
            QMessageBox.warning(self.parent(), "Hata", "Geçersiz grafik türü")
            return

        # Verileri hazırla
        try:
            plot_data = self.prepare_plot_data()
        except Exception as e:
            QMessageBox.critical(self.parent(), "Veri Hazırlama Hatası", f"Veri hazırlama sırasında hata oluştu: {str(e)}")
            return

        # Grafik düzenini ayarla
        rows, cols = self.calculate_grid_layout(count)
        self.matplotlib_widget.figure.set_size_inches(cols * 6, rows * 4.5)

        # Plot info'ya X ve Y ekseni etiketlerini dinamik olarak ata
        default_x_label = self.selected_x_column if self.selected_x_column else "X Değeri (İndeks)"
        default_y_label = self.selected_y_column if self.selected_y_column else "Y Değeri"
        default_cat_label = self.selected_category_column if self.selected_category_column else "Kategori"

        for i in range(count):
            ax = self.matplotlib_widget.figure.add_subplot(rows, cols, i + 1)
            title = f"{chart_type_text.split('(')[0].strip()} {i + 1}"
            ax.set_title(title, pad=20)

            plot_info = {
                'type': chart_type,
                'title': title,
                'xlabel': default_x_label,
                'ylabel': default_y_label,
                'category_label': default_cat_label
            }

            try:
                draw_func = getattr(self, f"draw_{chart_type}", None)
                if draw_func:
                    # plot_data'nın bir kopyasını göndererek her grafiğin kendi üzerinde işlem yapmasını sağla
                    draw_func(ax, plot_data.copy(), plot_info)
                else:
                    ax.text(0.5, 0.5, f"'{chart_type}' için çizim fonksiyonu yok",
                            ha='center', va='center', transform=ax.transAxes)
                    plot_info['error'] = f"Çizim fonksiyonu bulunamadı: {chart_type}"

                self.current_plot_data.append(plot_info)
            except Exception as e:
                error_msg = f"{i + 1}. grafik oluşturulurken hata: {str(e)}"
                ax.text(0.5, 0.5, error_msg, ha='center', va='center',
                        color='red', transform=ax.transAxes)
                self.current_plot_data.append({
                    'type': chart_type,
                    'title': title,
                    'error': error_msg
                })

        self.matplotlib_widget.figure.tight_layout()
        self.matplotlib_widget.canvas.draw()
        self.update_plot_info_label() # Global bilgi etiketini güncelle

    def calculate_grid_layout(self, count):
        if count == 1: return 1, 1
        if count == 2: return 1, 2
        if count == 3: return 1, 3
        if count == 4: return 2, 2
        if count == 5: return 2, 3
        if count == 6: return 2, 3
        if count == 7: return 3, 3
        if count == 8: return 3, 3
        if count == 9: return 3, 3
        if count == 10: return 3, 4
        cols = min(4, count)
        rows = (count + cols - 1) // cols
        return rows, cols

    def prepare_plot_data(self):
        """Yüklenen DataFrame'den seçili sütunlara göre verileri hazırlar."""
        if self.loaded_data_df is None:
            raise ValueError("Veri yüklenmedi.")

        data = {}
        df = self.loaded_data_df.copy() # Orijinal DataFrame'i değiştirmemek için kopya al

        # X Ekseni
        if self.selected_x_column and self.selected_x_column in df.columns:
            data['x'] = df[self.selected_x_column].values
            # Tarih/zaman sütunu ise kontrol et
            if pd.api.types.is_datetime64_any_dtype(data['x']):
                data['x'] = pd.to_datetime(data['x']).values # matplotlib için uygun datetime formatına çevir
            else:
                data['x'] = pd.to_numeric(data['x'], errors='coerce').values # Diğer türleri sayıya çevir
        else:
            data['x'] = np.arange(len(df)) # X seçilmemişse indeks kullan

        # Y Ekseni
        if self.selected_y_column and self.selected_y_column in df.columns:
            data['y'] = df[self.selected_y_column].values
            if pd.api.types.is_datetime64_any_dtype(data['y']):
                data['y'] = pd.to_datetime(data['y']).values
            else:
                data['y'] = pd.to_numeric(data['y'], errors='coerce').values
        else:
            data['y'] = np.zeros(len(df)) # Y seçilmemişse 0'lar dizisi kullan

        # Kategori
        if self.selected_category_column and self.selected_category_column in df.columns:
            data['category'] = df[self.selected_category_column].astype(str).values
        else:
            data['category'] = np.full(len(df), "Genel Kategori")

        # Renk
        if self.selected_color_column and self.selected_color_column in df.columns:
            data['color'] = df[self.selected_color_column].values
            # Renk verisi sayısal ise normalize et, değilse kategorik işlem yapabiliriz
            if pd.api.types.is_numeric_dtype(data['color']):
                # Ortalamayı almak için NaN değerleri temizle
                clean_color_data = data['color'][~np.isnan(data['color'])]
                mean_val = clean_color_data.mean() if len(clean_color_data) > 0 else 0
                data['color'] = pd.to_numeric(data['color'], errors='coerce').fillna(mean_val).values
            else:
                # Kategorik renk için benzersiz değerleri al ve map'le
                unique_colors = pd.Categorical(data['color']).categories.tolist()
                color_map = {val: plt.colormaps['tab10'](i % 10) for i, val in enumerate(unique_colors)}
                data['color_mapped'] = np.array([color_map[c] for c in data['color']])
        else:
            data['color'] = None # Renk sütunu seçilmezse otomatik renk atanacak
            data['color_mapped'] = None


        # Eksik (NaN) verileri temizleme (sadece sayısal eksenler için geçerli)
        # Tarih/zaman değerleri NaT olarak geldiyse de temizlemeliyiz.
        valid_mask = np.ones(len(df), dtype=bool)

        if pd.api.types.is_numeric_dtype(data['x']) or pd.api.types.is_datetime64_any_dtype(data['x']):
            if pd.api.types.is_numeric_dtype(data['x']):
                valid_mask &= ~np.isnan(data['x'])
            else: # datetime
                valid_mask &= ~pd.isna(data['x'])

        if pd.api.types.is_numeric_dtype(data['y']) or pd.api.types.is_datetime64_any_dtype(data['y']):
            if pd.api.types.is_numeric_dtype(data['y']):
                valid_mask &= ~np.isnan(data['y'])
            else: # datetime
                valid_mask &= ~pd.isna(data['y'])

        # Tüm ilgili dizileri aynı maske ile filtrele
        for key in ['x', 'y', 'category', 'color', 'color_mapped']:
            if key in data and data[key] is not None and len(data[key]) == len(valid_mask):
                data[key] = data[key][valid_mask]
            # Eğer x veya y tamamen geçersiz olursa boş array yap
            elif key in ['x', 'y'] and not np.any(valid_mask):
                data[key] = np.array([])


        if data['category'] is not None and len(data['category']) > 0:
            # Kategorik verileri işleme
            data['unique_categories'] = pd.Categorical(data['category']).categories.tolist()
            data['category_indices'] = pd.Categorical(data['category']).codes
        else:
            data['unique_categories'] = []
            data['category_indices'] = np.array([])

        return data

    # Grafik çizim fonksiyonları (mevcut halleriyle bırakıldı, iyi görünüyorlar)
    def draw_plot(self, ax, data, plot_info):
        if len(data['x']) == 0 or len(data['y']) == 0:
            ax.text(0.5, 0.5, "Çizgi grafiği için yeterli veri bulunamadı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        ax.plot(data['x'], data['y'], linewidth=2, marker='o', markersize=5)
        ax.set_xlabel(plot_info['xlabel'])
        ax.set_ylabel(plot_info['ylabel'])
        ax.grid(True, linestyle='--', alpha=0.6)

        # Trend çizgisi ekle
        # Tarih/zaman verileri sayısal trend çizgisi için uygun değildir.
        # Bu yüzden sadece sayısal x ve y değerleri için kontrol yapılır.
        if (len(data['x']) > 1 and np.issubdtype(data['x'].dtype, np.number) and
            len(data['y']) > 1 and np.issubdtype(data['y'].dtype, np.number)):
            # Sadece sayısal veriler için trend çizgisi hesapla
            z = np.polyfit(data['x'], data['y'], 1)
            p = np.poly1d(z)
            ax.plot(data['x'], p(data['x']), "r--", linewidth=1, label='Trend')
            ax.legend()
            plot_info['trend_line'] = p(data['x']).tolist()
        else:
            plot_info['trend_line'] = None

        plot_info.update({
            'data_x': (data['x'].tolist() if not pd.api.types.is_datetime64_any_dtype(data['x']) else
                       [str(d) for d in data['x']]), # Tarihleri string olarak kaydet
            'data_y': (data['y'].tolist() if not pd.api.types.is_datetime64_any_dtype(data['y']) else
                       [str(d) for d in data['y']]),
        })

    def draw_bar(self, ax, data, plot_info):
        # Bar grafikleri için X ekseni genellikle kategorik (string) olmalıdır.
        # Y ekseni sayısal olmalıdır.
        if data['category'] is None or len(data['category']) == 0 or \
           data['y'] is None or len(data['y']) == 0 or \
           not np.issubdtype(data['y'].dtype, np.number):
            ax.text(0.5, 0.5, "Bar grafiği için kategori (X) ve sayısal değer (Y) sütunları seçilmeli ve dolu olmalı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        df_temp = pd.DataFrame({'category': data['category'], 'value': data['y']})
        grouped_data = df_temp.groupby('category')['value'].sum().reset_index()

        categories = grouped_data['category'].tolist()
        values = grouped_data['value'].tolist()

        # Otomatik renk ataması
        colors = plt.colormaps['viridis'](np.linspace(0, 1, len(categories)))
        bars = ax.bar(categories, values, color=colors)

        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2., height,
                    f'{height:.1f}', ha='center', va='bottom')

        ax.set_xlabel(plot_info['category_label'])
        ax.set_ylabel(plot_info['ylabel'])
        ax.tick_params(axis='x', rotation=45)

        plot_info.update({
            'categories': categories,
            'values': values,
            'xlabel': plot_info['category_label'], # Bar grafiği için X ekseni kategori
            'ylabel': plot_info['ylabel']
        })

    def draw_hist(self, ax, data, plot_info):
        if data['y'] is None or len(data['y']) == 0 or not np.issubdtype(data['y'].dtype, np.number):
            ax.text(0.5, 0.5, "Histogram için sayısal değer sütunu seçilmeli ve dolu olmalı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        numeric_y = data['y'][np.issubdtype(data['y'].dtype, np.number)]
        if len(numeric_y) == 0:
            ax.text(0.5, 0.5, "Histogram için sayısal veri bulunamadı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        n, bins, patches = ax.hist(numeric_y, bins='auto', color='#607c8e',
                                   edgecolor='black', alpha=0.7)

        if len(numeric_y) > 1:
            density = stats.gaussian_kde(numeric_y)
            xs = np.linspace(min(numeric_y), max(numeric_y), 200)
            density._compute_covariance()
            ax2 = ax.twinx()
            ax2.plot(xs, density(xs), color='darkred', linewidth=2)
            ax2.set_ylabel('Yoğunluk', color='darkred')
            ax2.tick_params(axis='y', labelcolor='darkred')

        ax.set_xlabel(plot_info['ylabel']) # Histogram için genellikle Y ekseni verisinin kendisi X eksenidir.
        ax.set_ylabel('Frekans')

        plot_info.update({
            'data': numeric_y.tolist(),
            'bins': bins.tolist(),
            'density': density(xs).tolist() if len(numeric_y) > 1 else None,
            'statistics': {
                'mean': float(np.mean(numeric_y)),
                'median': float(np.median(numeric_y)),
                'std': float(np.std(numeric_y)),
                'skewness': float(stats.skew(numeric_y)),
                'kurtosis': float(stats.kurtosis(numeric_y))
            }
        })

    def draw_pie(self, ax, data, plot_info):
        # Pasta grafiği için kategori (etiketler) ve sayısal değer (dilim boyutları) gerekir.
        if data['category'] is None or len(data['category']) == 0 or \
           data['y'] is None or len(data['y']) == 0 or \
           not np.issubdtype(data['y'].dtype, np.number):
            ax.text(0.5, 0.5, "Pasta grafiği için kategori ve sayısal değer sütunları seçilmeli ve dolu olmalı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        df_temp = pd.DataFrame({'category': data['category'], 'value': data['y']})
        grouped_data = df_temp.groupby('category')['value'].sum().reset_index()

        sizes = grouped_data['value'].tolist()
        labels = grouped_data['category'].tolist()

        ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90)
        ax.axis('equal')

        plot_info.update({
            'sizes': sizes,
            'labels': labels,
            'xlabel': plot_info['category_label'], # Pie grafiği için etiketler kategoridir
            'ylabel': plot_info['ylabel']
        })

    def draw_scatter(self, ax, data, plot_info):
        # Dağılım grafiği için hem X hem de Y ekseni sayısal veya tarih/zaman olmalıdır.
        if (len(data['x']) == 0 or len(data['y']) == 0 or
            (not np.issubdtype(data['x'].dtype, np.number) and not pd.api.types.is_datetime64_any_dtype(data['x'])) or
            (not np.issubdtype(data['y'].dtype, np.number) and not pd.api.types.is_datetime64_any_dtype(data['y']))):
            ax.text(0.5, 0.5, "Dağılım grafiği için yeterli sayısal veya tarihsel veri bulunamadı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        # Renklendirme seçeneği
        if data['color_mapped'] is not None and len(data['color_mapped']) == len(data['x']):
            colors = data['color_mapped']
            cmap = None # Renkler zaten maplenmiş
            cbar_label = plot_info.get('color_label', self.selected_color_column) # Renk sütununun adını kullan
        elif data['color'] is not None and len(data['color']) == len(data['x']) and pd.api.types.is_numeric_dtype(data['color']):
            colors = data['color']
            cmap = 'viridis' # Sayısal renk verisi için colormap kullan
            cbar_label = plot_info.get('color_label', self.selected_color_column)
        else:
            colors = None # Otomatik renk döngüsü
            cmap = None
            cbar_label = None


        # Boyutlar için rastgelelik yerine sabit bir boyut veya başka bir sütun kullanılabilir.
        # Eğer veri boyutu 0 ise boş liste atayın
        sizes = np.random.rand(len(data['x'])) * 200 + 20 if len(data['x']) > 0 else []

        scatter = ax.scatter(data['x'], data['y'], c=colors, s=sizes,
                             alpha=0.7, cmap=cmap)
        ax.set_xlabel(plot_info['xlabel'])
        ax.set_ylabel(plot_info['ylabel'])

        if cmap and cbar_label: # Sadece colormap varsa colorbar ekle
            plt.colorbar(scatter, ax=ax, label=cbar_label)

        plot_info.update({
            'data_x': (data['x'].tolist() if not pd.api.types.is_datetime64_any_dtype(data['x']) else
                       [str(d) for d in data['x']]),
            'data_y': (data['y'].tolist() if not pd.api.types.is_datetime64_any_dtype(data['y']) else
                       [str(d) for d in data['y']]),
            'colors': colors.tolist() if hasattr(colors, 'tolist') else (colors if isinstance(colors, str) else None),
            'sizes': sizes.tolist() if hasattr(sizes, 'tolist') else []
        })

    def draw_fill_between(self, ax, data, plot_info):
        # Alan grafiği için X ve Y ekseni sayısal veya tarih/zaman olmalıdır.
        if (len(data['x']) == 0 or len(data['y']) == 0 or
            (not np.issubdtype(data['x'].dtype, np.number) and not pd.api.types.is_datetime64_any_dtype(data['x'])) or
            (not np.issubdtype(data['y'].dtype, np.number) and not pd.api.types.is_datetime64_any_dtype(data['y']))):
            ax.text(0.5, 0.5, "Alan grafiği için yeterli sayısal veya tarihsel veri bulunamadı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        y1 = data['y']
        y2 = y1 - (np.mean(y1) / 2) if len(y1) > 0 and np.issubdtype(y1.dtype, np.number) else np.zeros_like(y1)

        ax.plot(data['x'], y1, label='Veri Hattı')
        ax.fill_between(data['x'], y1, y2, color='skyblue', alpha=0.4)
        ax.set_xlabel(plot_info['xlabel'])
        ax.set_ylabel(plot_info['ylabel'])
        ax.legend()

        plot_info.update({
            'data_x': (data['x'].tolist() if not pd.api.types.is_datetime64_any_dtype(data['x']) else
                       [str(d) for d in data['x']]),
            'data_y1': (y1.tolist() if not pd.api.types.is_datetime64_any_dtype(y1) else
                        [str(d) for d in y1]),
            'data_y2': (y2.tolist() if not pd.api.types.is_datetime64_any_dtype(y2) else
                        [str(d) for d in y2])
        })

    def draw_boxplot(self, ax, data, plot_info):
        if data['y'] is None or len(data['y']) == 0 or not np.issubdtype(data['y'].dtype, np.number):
            ax.text(0.5, 0.5, "Kutu grafiği için sayısal değer sütunu seçilmeli ve dolu olmalı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        numeric_y = data['y']
        if len(numeric_y) == 0:
            ax.text(0.5, 0.5, "Kutu grafiği için sayısal veri bulunamadı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        groups = []
        labels = []

        if self.selected_category_column and data['category'] is not None and len(data['category']) > 0:
            df_temp = pd.DataFrame({'category': data['category'], 'value': numeric_y})
            for cat in df_temp['category'].unique():
                group_data = df_temp[df_temp['category'] == cat]['value'].dropna().tolist()
                if group_data: # Boş grupları ekleme
                    groups.append(group_data)
                    labels.append(str(cat))
        else:
            group_data = numeric_y.dropna().tolist()
            if group_data:
                groups = [group_data]
                labels = [plot_info['ylabel'] if plot_info['ylabel'] else 'Değerler']

        if not groups:
            ax.text(0.5, 0.5, "Kutu grafiği için geçerli grup verisi bulunamadı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        ax.boxplot(groups, patch_artist=True)
        ax.set_xticklabels(labels, rotation=45, ha='right')
        ax.set_ylabel(plot_info['ylabel'])

        plot_info.update({
            'data_groups': groups,
            'labels': labels,
            'xlabel': plot_info['category_label'] if self.selected_category_column else 'Gruplar'
        })

    def draw_violinplot(self, ax, data, plot_info):
        if data['y'] is None or len(data['y']) == 0 or not np.issubdtype(data['y'].dtype, np.number):
            ax.text(0.5, 0.5, "Violin grafiği için sayısal değer sütunu seçilmeli ve dolu olmalı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        numeric_y = data['y']
        if len(numeric_y) == 0:
            ax.text(0.5, 0.5, "Violin grafiği için sayısal veri bulunamadı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        groups = []
        labels = []

        if self.selected_category_column and data['category'] is not None and len(data['category']) > 0:
            df_temp = pd.DataFrame({'category': data['category'], 'value': numeric_y})
            for cat in df_temp['category'].unique():
                group_data = df_temp[df_temp['category'] == cat]['value'].dropna().tolist()
                if group_data:
                    groups.append(group_data)
                    labels.append(str(cat))
        else:
            group_data = numeric_y.dropna().tolist()
            if group_data:
                groups = [group_data]
                labels = [plot_info['ylabel'] if plot_info['ylabel'] else 'Değerler']

        if not groups:
            ax.text(0.5, 0.5, "Violin grafiği için geçerli grup verisi bulunamadı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        ax.violinplot(groups, showmeans=True, showmedians=True)
        ax.set_xticks(np.arange(1, len(groups) + 1))
        ax.set_xticklabels(labels, rotation=45, ha='right')
        ax.set_ylabel(plot_info['ylabel'])

        plot_info.update({
            'data_groups': groups,
            'labels': labels,
            'xlabel': plot_info['category_label'] if self.selected_category_column else 'Gruplar'
        })

    def draw_stem(self, ax, data, plot_info):
        # Stem grafiği için X ve Y ekseni sayısal olmalıdır.
        if (len(data['x']) == 0 or len(data['y']) == 0 or
            not np.issubdtype(data['x'].dtype, np.number) or
            not np.issubdtype(data['y'].dtype, np.number)):
            ax.text(0.5, 0.5, "Stem grafiği için yeterli sayısal veri bulunamadı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        ax.stem(data['x'], data['y'])
        ax.set_xlabel(plot_info['xlabel'])
        ax.set_ylabel(plot_info['ylabel'])

        plot_info.update({
            'data_x': data['x'].tolist(),
            'data_y': data['y'].tolist()
        })

    def draw_errorbar(self, ax, data, plot_info):
        # Hata çubuklu grafik için X ve Y ekseni sayısal olmalıdır.
        if (len(data['x']) == 0 or len(data['y']) == 0 or
            not np.issubdtype(data['x'].dtype, np.number) or
            not np.issubdtype(data['y'].dtype, np.number)):
            ax.text(0.5, 0.5, "Hata çubuklu grafiği için yeterli sayısal veri bulunamadı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        # Y err değerlerini hesaplamadan önce NaN değerleri temizle
        # fillna(0) yapıldığı için burada NaN kalmaması gerekir, ancak yine de kontrol iyi bir pratiktir.
        clean_x = data['x']
        clean_y = data['y']

        if len(clean_x) == 0 or len(clean_y) == 0:
            ax.text(0.5, 0.5, "Hata çubuklu grafiği için geçerli sayısal veri bulunamadı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        y_err = 0.1 * np.abs(clean_y)  # %10 hata payı
        ax.errorbar(clean_x, clean_y, yerr=y_err, fmt='-o',
                    capsize=5, capthick=2)
        ax.set_xlabel(plot_info['xlabel'])
        ax.set_ylabel(plot_info['ylabel'])

        plot_info.update({
            'data_x': clean_x.tolist(),
            'data_y': clean_y.tolist(),
            'y_error': y_err.tolist()
        })

    # Raporlama fonksiyonları
    def generate_pdf_report(self):
        if not self.current_plot_data:
            QMessageBox.warning(self, "Uyarı", "Raporlanacak grafik bulunamadı!")
            return

        pdf_file, _ = QFileDialog.getSaveFileName(
            self, "PDF Olarak Kaydet", "grafik_raporu.pdf",
            "PDF Dosyaları (*.pdf)")

        if not pdf_file:
            return

        progress = QProgressDialog("PDF Raporu Oluşturuluyor...",
                                   "İptal", 0, 0, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.show()
        QApplication.processEvents()

        try:
            pdf = FPDF()
            pdf.add_page()

            # Başlık ve genel bilgiler
            self.add_pdf_header(pdf)

            # Grafik görselleri
            self.add_plot_images_to_pdf(pdf)

            # Detaylı analiz
            self.add_detailed_analysis(pdf)

            # İstatistiksel analizler
            self.add_statistical_analysis(pdf)

            # Korelasyon analizleri
            self.add_correlation_analysis(pdf)

            pdf.output(pdf_file)
            QMessageBox.information(
                self, "Başarılı",
                f"PDF raporu oluşturuldu:\n{pdf_file}"
            )
        except Exception as e:
            QMessageBox.critical(
                self, "Hata",
                f"PDF oluşturma hatası:\n{str(e)}"
            )
        finally:
            progress.close()

    def add_pdf_header(self, pdf):
        try:
            # NotoSans fontunu eklemeye çalış, yoksa Arial kullan
            # Font dosyasının uygulamanın çalıştığı dizinde veya sistem fontlarında olması gerekir.
            pdf.add_font('NotoSans', '', 'NotoSans-Regular.ttf')
            pdf.add_font('NotoSans', 'B', 'NotoSans-Bold.ttf')
            pdf.set_font("NotoSans", "B", 16) # Daha büyük başlık
        except Exception:
            pdf.set_font("Arial", "B", 16)

        pdf.cell(0, 10, "GRAFİK ANALİZ RAPORU", 0, 1, 'C')
        pdf.ln(5)

        try:
            pdf.set_font("NotoSans", "", 10)
        except Exception:
            pdf.set_font("Arial", "", 10)

        pdf.cell(0, 6, f"Oluşturulma Tarihi: {QDateTime.currentDateTime().toString('dd.MM.yyyy HH:mm:ss')}", 0, 1, 'R')
        pdf.ln(10)
        pdf.multi_cell(0, 6,
                       "Bu rapor, uygulama tarafından oluşturulan grafiklerin detaylı analizini içerir. "
                       "Rapor, grafik verilerinden elde edilen istatistiksel bilgileri ve karşılaştırmaları sunar."
                       )
        pdf.ln(10)

    def add_plot_images_to_pdf(self, pdf):
        temp_img = "temp_plot.png"
        try:
            self.matplotlib_widget.figure.savefig(temp_img, dpi=200, bbox_inches='tight')  # Daha yüksek çözünürlük

            try:
                pdf.set_font("NotoSans", "B", 14)
            except Exception:
                pdf.set_font("Arial", "B", 14)

            pdf.cell(0, 10, "1. Grafik Görselleri", 0, 1)
            pdf.ln(5)

            img_width_in = self.matplotlib_widget.figure.get_size_inches()[0]
            img_height_in = self.matplotlib_widget.figure.get_size_inches()[1]

            # PDF sayfa genişliği (210mm) ve kenar boşlukları (toplam 20mm)
            pdf_page_width_mm = 210
            margin_mm = 10 * 2 # 10mm sağ, 10mm sol
            available_width_mm = pdf_page_width_mm - margin_mm

            # Görseli mm cinsinden ölçekle
            img_width_mm = img_width_in * 25.4 # inç'ten mm'ye
            img_height_mm = img_height_in * 25.4

            if img_width_mm > available_width_mm:
                scale_factor = available_width_mm / img_width_mm
                img_width_mm *= scale_factor
                img_height_mm *= scale_factor

            pdf.image(temp_img, x=XPos.CENTER, w=img_width_mm)
            pdf.ln(10)
        except Exception as e:
            try:
                pdf.set_font("NotoSans", "", 10)
            except Exception:
                pdf.set_font("Arial", "", 10)
            pdf.multi_cell(0, 6, f"Grafik görseli eklenemedi: {str(e)}")
        finally:
            if os.path.exists(temp_img):
                os.remove(temp_img)

    def add_detailed_analysis(self, pdf):
        pdf.add_page()
        try:
            pdf.set_font("NotoSans", "B", 14)
        except Exception:
            pdf.set_font("Arial", "B", 14)

        pdf.cell(0, 10, "2. Grafik Detayları ve Analizler", 0, 1)
        pdf.ln(5)

        try:
            pdf.set_font("NotoSans", "", 10)
        except Exception:
            pdf.set_font("Arial", "", 10)

        for i, plot in enumerate(self.current_plot_data):
            try:
                pdf.set_font("NotoSans", "B", 12)
            except Exception:
                pdf.set_font("Arial", "B", 12)

            pdf.cell(0, 8, f"{i + 1}. {plot.get('title', 'Başlıksız Grafik')}", 0, 1)

            try:
                pdf.set_font("NotoSans", "", 10)
            except Exception:
                pdf.set_font("Arial", "", 10)

            if 'error' in plot:
                pdf.multi_cell(0, 6, f"Hata: {plot['error']}")
                pdf.ln(5)
                continue

            # Genel bilgiler
            pdf.cell(0, 6, f"Grafik Türü: {plot.get('type', 'Bilinmiyor')}", 0, 1)
            pdf.cell(0, 6, f"X Ekseni: {plot.get('xlabel', 'Belirtilmedi')}", 0, 1)
            pdf.cell(0, 6, f"Y Ekseni: {plot.get('ylabel', 'Belirtilmedi')}", 0, 1)

            # Grafik türüne özel bilgiler
            if plot['type'] in ['plot', 'scatter', 'stem', 'errorbar', 'fill_between']:
                if 'data_x' in plot and 'data_y' in plot and len(plot['data_x']) > 0:
                    pdf.cell(0, 6, f"Veri Noktası Sayısı: {len(plot['data_x'])}", 0, 1)
                    # Min/Max kontrolü ekle
                    # Tarih/zaman için min/max kontrolü string dönüşümden sonra anlamlı değil,
                    # ancak matplotlib zaten datetime objelerini doğru işleyecektir.
                    if len(plot['data_x']) > 0:
                        # Eğer tarih/zaman ise özel formatla
                        if isinstance(plot['data_x'][0], str) and 'T' in plot['data_x'][0]: # ISO formatında tarih/saat
                            min_x = min(pd.to_datetime(plot['data_x']))
                            max_x = max(pd.to_datetime(plot['data_x']))
                            pdf.cell(0, 6, f"X Aralığı: [{min_x.strftime('%Y-%m-%d %H:%M')}, {max_x.strftime('%Y-%m-%d %H:%M')}]", 0, 1)
                        else:
                            # float kontrolü eklendi
                            if all(isinstance(val, (int, float)) for val in plot['data_x']):
                                pdf.cell(0, 6, f"X Aralığı: [{min(plot['data_x']):.2f}, {max(plot['data_x']):.2f}]", 0, 1)
                            else:
                                pdf.cell(0, 6, f"X Aralığı: [{plot['data_x'][0]}, {plot['data_x'][-1]}] (Sayısal değil)", 0, 1)

                    if len(plot['data_y']) > 0:
                        if isinstance(plot['data_y'][0], str) and 'T' in plot['data_y'][0]:
                            min_y = min(pd.to_datetime(plot['data_y']))
                            max_y = max(pd.to_datetime(plot['data_y']))
                            pdf.cell(0, 6, f"Y Aralığı: [{min_y.strftime('%Y-%m-%d %H:%M')}, {max_y.strftime('%Y-%m-%d %H:%M')}]", 0, 1)
                        else:
                            if all(isinstance(val, (int, float)) for val in plot['data_y']):
                                pdf.cell(0, 6, f"Y Aralığı: [{min(plot['data_y']):.2f}, {max(plot['data_y']):.2f}]", 0, 1)
                            else:
                                pdf.cell(0, 6, f"Y Aralığı: [{plot['data_y'][0]}, {plot['data_y'][-1]}] (Sayısal değil)", 0, 1)

                    if 'trend_line' in plot and plot['trend_line'] is not None:
                        pdf.cell(0, 6, "Trend Çizgisi Eklendi", 0, 1)


            elif plot['type'] == 'bar':
                if 'categories' in plot and 'values' in plot and len(plot['categories']) > 0:
                    pdf.cell(0, 6, f"Kategori Sayısı: {len(plot['categories'])}", 0, 1)
                    pdf.cell(0, 6, f"Toplam Değer: {sum(plot['values']):.2f}", 0, 1)

            elif plot['type'] == 'hist':
                if 'statistics' in plot:
                    stats_info = plot['statistics']
                    pdf.cell(0, 6, f"Veri Sayısı: {len(plot['data'])}", 0, 1)
                    pdf.cell(0, 6, f"Ortalama: {stats_info.get('mean', 0):.2f}", 0, 1)
                    pdf.cell(0, 6, f"Medyan: {stats_info.get('median', 0):.2f}", 0, 1)
                    pdf.cell(0, 6, f"Standart Sapma: {stats_info.get('std', 0):.2f}", 0, 1)
                    pdf.cell(0, 6, f"Çarpıklık: {stats_info.get('skewness', 0):.2f}", 0, 1)
                    pdf.cell(0, 6, f"Basıklık: {stats_info.get('kurtosis', 0):.2f}", 0, 1)

            elif plot['type'] == 'pie':
                if 'labels' in plot and 'sizes' in plot and len(plot['labels']) > 0:
                    pdf.cell(0, 6, f"Dilim Sayısı: {len(plot['labels'])}", 0, 1)
                    pdf.cell(0, 6, f"Toplam Değer: {sum(plot['sizes']):.2f}", 0, 1)

            pdf.ln(5)

    def add_statistical_analysis(self, pdf):
        pdf.add_page()
        try:
            pdf.set_font("NotoSans", "B", 14)
        except Exception:
            pdf.set_font("Arial", "B", 14)

        pdf.cell(0, 10, "3. İstatistiksel Analizler", 0, 1)
        pdf.ln(5)

        try:
            pdf.set_font("NotoSans", "", 10)
        except Exception:
            pdf.set_font("Arial", "", 10)

        all_numeric_data = []
        for plot in self.current_plot_data:
            if 'error' in plot:
                continue

            # Sadece sayısal olan verileri topla
            # Tarih/zaman verileri istatistiksel analizde doğrudan kullanılamaz, bu yüzden sadece sayısal olanlar dahil edilir.
            if 'data_y' in plot and plot['data_y'] is not None:
                all_numeric_data.extend([x for x in plot['data_y'] if isinstance(x, (int, float))])
            elif 'data' in plot and plot['data'] is not None:  # hist için
                all_numeric_data.extend([x for x in plot['data'] if isinstance(x, (int, float))])
            elif 'values' in plot and plot['values'] is not None:  # bar için
                all_numeric_data.extend([x for x in plot['values'] if isinstance(x, (int, float))])
            elif 'sizes' in plot and plot['sizes'] is not None:  # pie için
                all_numeric_data.extend([x for x in plot['sizes'] if isinstance(x, (int, float))])

        if not all_numeric_data:
            pdf.multi_cell(0, 6, "İstatistiksel analiz için yeterli sayısal veri bulunamadı.")
            return

        all_numeric_data_np = np.array(all_numeric_data)

        # Temel istatistikler
        pdf.cell(0, 6, f"Toplam Sayısal Veri Noktası Sayısı: {len(all_numeric_data_np)}", 0, 1)
        if len(all_numeric_data_np) > 0:
            pdf.cell(0, 6, f"Minimum Değer: {np.min(all_numeric_data_np):.2f}", 0, 1)
            pdf.cell(0, 6, f"Maksimum Değer: {np.max(all_numeric_data_np):.2f}", 0, 1)
            pdf.cell(0, 6, f"Ortalama: {np.mean(all_numeric_data_np):.2f}", 0, 1)
            pdf.cell(0, 6, f"Medyan: {np.median(all_numeric_data_np):.2f}", 0, 1)
            pdf.cell(0, 6, f"Standart Sapma: {np.std(all_numeric_data_np):.2f}", 0, 1)
        else:
            pdf.cell(0, 6, "İstatistiksel değerler hesaplanamadı (veri yok).", 0, 1)


        # Çarpıklık ve basıklık için en az 3 veri noktası gerekir
        if len(all_numeric_data_np) >= 3:
            pdf.cell(0, 6, f"Çarpıklık: {stats.skew(all_numeric_data_np):.2f}", 0, 1)
            pdf.cell(0, 6, f"Basıklık: {stats.kurtosis(all_numeric_data_np):.2f}", 0, 1)
        else:
            pdf.cell(0, 6, "Çarpıklık ve Basıklık hesaplamak için yeterli veri yok (en az 3).", 0, 1)
        pdf.ln(5)

        # Normallik testi (en az 8 veri noktası gerekir)
        if len(all_numeric_data_np) >= 8:
            stat, p = stats.normaltest(all_numeric_data_np)
            pdf.cell(0, 6, f"Normallik Testi (p-değeri): {p:.4f}", 0, 1)
            pdf.cell(0, 6, "→ " + ("Veri normal dağılıma uygun" if p > 0.05 else "Veri normal dağılıma uygun değil"), 0,
                     1)
        else:
            pdf.cell(0, 6, "Normallik testi için yeterli veri (en az 8) bulunamadı.", 0, 1)
        pdf.ln(10)

    def add_correlation_analysis(self, pdf):
        # En az iki çizgi veya dağılım grafiği olması gerekiyor
        plot_type_data = [
            plot for plot in self.current_plot_data
            if plot.get('type') in ['plot', 'scatter', 'stem',
                                    'errorbar'] and 'data_x' in plot and 'data_y' in plot and not plot.get('error')
        ]

        # Sadece sayısal veri içeren grafikleri filtrele
        plot_type_data = [
            plot for plot in plot_type_data
            if (isinstance(plot['data_x'], list) and all(isinstance(x, (int, float)) for x in plot['data_x'])) and
               (isinstance(plot['data_y'], list) and all(isinstance(y, (int, float)) for y in plot['data_y']))
        ]


        pdf.add_page()
        try:
            pdf.set_font("NotoSans", "B", 14)
        except Exception:
            pdf.set_font("Arial", "B", 14)

        pdf.cell(0, 10, "4. Grafikler Arası İlişkiler", 0, 1)
        pdf.ln(5)

        try:
            pdf.set_font("NotoSans", "", 10)
        except Exception:
            pdf.set_font("Arial", "", 10)

        if len(plot_type_data) < 2:
            pdf.multi_cell(0, 6, "Korelasyon analizi için en az iki uygun (sayısal X ve Y verisi içeren) grafik bulunamadı (Çizgi, Dağılım vb.).")
            return

        try:
            # Sadece ilk iki uygun grafiğin Y verilerini karşılaştır
            data1_y = np.array(plot_type_data[0]['data_y'])
            data2_y = np.array(plot_type_data[1]['data_y'])
            data1_x = np.array(plot_type_data[0]['data_x'])
            data2_x = np.array(plot_type_data[1]['data_x'])


            # NaN değerleri temizle (fillna(0) yapıldığı için burada NaN kalmaması gerekir)
            # Ancak string'den sayıya çevrimde hata olmuşsa NaN olabilir.
            valid_mask1 = ~np.isnan(data1_y) & ~np.isnan(data1_x)
            valid_mask2 = ~np.isnan(data2_y) & ~np.isnan(data2_x)

            clean_data1_x = data1_x[valid_mask1]
            clean_data1_y = data1_y[valid_mask1]
            clean_data2_x = data2_x[valid_mask2]
            clean_data2_y = data2_y[valid_mask2]


            # Ortak uzunluğa getir
            min_len_x = min(len(clean_data1_x), len(clean_data2_x))
            min_len_y = min(len(clean_data1_y), len(clean_data2_y))

            # Sadece Y değerleri arasındaki korelasyonu hesapla (en yaygın kullanım)
            if min_len_y >= 2:
                corr_y = np.corrcoef(clean_data1_y[:min_len_y], clean_data2_y[:min_len_y])[0, 1]
                pdf.cell(0, 6,
                         f"'{plot_type_data[0]['title']}' ve '{plot_type_data[1]['title']}' verilerinin Y değerleri arasındaki korelasyon: {corr_y:.3f}",
                         0, 1)
                if abs(corr_y) > 0.7:
                    pdf.cell(0, 6, "→ Güçlü bir ilişki var", 0, 1)
                elif abs(corr_y) > 0.3:
                    pdf.cell(0, 6, "→ Orta düzeyde ilişki var", 0, 1)
                else:
                    pdf.cell(0, 6, "→ Zayıf veya ilişki yok", 0, 1)
            else:
                pdf.cell(0, 6, "Y değerleri arasında korelasyon analizi için yeterli ortak sayısal veri bulunamadı.", 0, 1)
            pdf.ln(2)

            # İsteğe bağlı olarak X değerleri arasındaki korelasyon
            if min_len_x >= 2:
                corr_x = np.corrcoef(clean_data1_x[:min_len_x], clean_data2_x[:min_len_x])[0, 1]
                pdf.cell(0, 6,
                         f"'{plot_type_data[0]['title']}' ve '{plot_type_data[1]['title']}' verilerinin X değerleri arasındaki korelasyon: {corr_x:.3f}",
                         0, 1)
                if abs(corr_x) > 0.7:
                    pdf.cell(0, 6, "→ Güçlü bir ilişki var", 0, 1)
                elif abs(corr_x) > 0.3:
                    pdf.cell(0, 6, "→ Orta düzeyde ilişki var", 0, 1)
                else:
                    pdf.cell(0, 6, "→ Zayıf veya ilişki yok", 0, 1)
            else:
                pdf.cell(0, 6, "X değerleri arasında korelasyon analizi için yeterli ortak sayısal veri bulunamadı.", 0, 1)
            pdf.ln(5)

            # Regresyon analizi (ilk grafiğin X'i ile ikinci grafiğin Y'si arasında)
            # Daha anlamlı bir regresyon için X ve Y'nin aynı grafikten gelmesi tercih edilir.
            # Burada farklı grafiklerden X ve Y alınıyor, bu yorumlanırken dikkat edilmeli.
            if len(clean_data1_y) >= 2 and len(clean_data2_y) >= 2:
                min_len_reg = min(len(clean_data1_y), len(clean_data2_y))
                # Regresyonu daha genel hale getirelim: plot1_y ve plot2_y arasında
                if min_len_y >= 2:
                    slope, intercept, r_value, p_value, std_err = stats.linregress(clean_data1_y[:min_len_y], clean_data2_y[:min_len_y])
                    pdf.cell(0, 6,
                             f"'{plot_type_data[0]['title']}' Y ve '{plot_type_data[1]['title']}' Y arasındaki regresyon denklemi: y = {slope:.3f}x + {intercept:.3f}",
                             0, 1)
                    pdf.cell(0, 6, f"R-kare değeri: {r_value**2:.3f}", 0, 1)
                else:
                    pdf.cell(0, 6, "Regresyon analizi için yeterli ortak sayısal veri bulunamadı.", 0, 1)
        except Exception as e:
            pdf.multi_cell(0, 6, f"Korelasyon analizi yapılamadı: {str(e)}")
        pdf.ln(10)

    def save_graph(self, file_format):
        if not self.matplotlib_widget.figure.axes:
            QMessageBox.warning(self, "Uyarı", "Kaydedilecek grafik bulunamadı!")
            return

        file_name, _ = QFileDialog.getSaveFileName(
            self, "Grafiği Kaydet", "grafik",
            f"Grafik Dosyaları (*.{file_format});;Tüm Dosyalar (*)")

        if file_name:
            try:
                self.matplotlib_widget.figure.savefig(file_name, format=file_format,
                                                      facecolor=self.matplotlib_widget.figure.get_facecolor(),
                                                      bbox_inches='tight')
                QMessageBox.information(
                    self, "Başarılı",
                    f"Grafik başarıyla kaydedildi:\n{file_name}"
                )
            except Exception as e:
                QMessageBox.critical(
                    self, "Hata",
                    f"Grafik kaydedilemedi:\n{str(e)}"
                )

    def print_graph(self):
        QMessageBox.information(
            self, "Yazdırma",
            "Grafik yazdırma işlemi başlatılacak. (Henüz tam işlevsel değil)"
        )


if __name__ == '__main__':
    # Gerekli ikon dizinini oluştur (eğer yoksa)
    if not os.path.exists('icons'):
        os.makedirs('icons')
    # Örnek ikon dosyaları için bir not:
    # Bu klasöre 'data_file.png', 'excel.png', 'exit.png',
    # 'report_pdf.png', 'save_jpeg.png', 'save_pdf.png', 'save_png.png', 'save_svg.png', 'print.png'
    # gibi ikonları koymanız gerekmektedir. Aksi takdirde ikonlar görüntülenmez.

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())