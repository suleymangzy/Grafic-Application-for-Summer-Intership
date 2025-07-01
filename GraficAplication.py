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
        self.ai_pipeline = None  # Bu değişken kullanılmıyor, kaldırılabilir veya gelecekteki kullanım için bırakılabilir.

        # Veri yükleme ve seçimi için değişkenler
        self.loaded_data_df = None  # Pandas DataFrame olarak tutalım
        self.column_names = []
        self.selected_x_column = None
        self.selected_y_column = None
        self.selected_category_column = None
        self.selected_color_column = None

        # UI bileşenlerini oluştur
        self.init_ui()
        self.create_menu()
        self.update_plot_info_label("", 0)  # Başlangıçta boş

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
        # Word, Excel, PPTX gibi dosya açma işlevleri veri görselleştirme uygulaması için biraz alakasız.
        # Bunları kaldırıp sadece veri dosyası açma ve uygulamadan çıkış bırakmak daha mantıklı olabilir.
        # Şimdilik yorum satırı yapıldı, isteğe bağlı olarak tamamen kaldırılabilir.
        # file_menu.addAction(create_action(
        #     "Word Dosyası &Aç...", 'icons/word.png', "Ctrl+W",
        #     "Bir Word belgesini açar",
        #     lambda: self.open_file("Word Dosyaları (*.docx *.doc);;Tüm Dosyalar (*)", "word")
        # ))
        # file_menu.addAction(create_action(
        #     "Excel Dosyası &Aç...", 'icons/excel.png', "Ctrl+E",
        #     "Bir Excel çalışma sayfasını açar",
        #     lambda: self.open_file("Excel Dosyaları (*.xlsx *.xls);;Tüm Dosyalar (*)", "excel")
        # ))
        # file_menu.addAction(create_action(
        #     "PPTX Dosyası &Aç...", 'icons/pptx.png', "Ctrl+P",
        #     "Bir PowerPoint sunumunu açar",
        #     lambda: self.open_file("PowerPoint Dosyaları (*.pptx *.ppt);;Tüm Dosyalar (*)", "pptx")
        # ))
        # file_menu.addSeparator()
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
        plot_menu = menubar.addMenu("&Grafik Oluştur")
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
            plot_menu.addAction(action)

        # İndir/Yazdır Menüsü
        download_print_menu = menubar.addMenu("&İndir / Yazdır")
        save_as_menu = QMenu("Farklı Kaydet", self)
        formats = {"PNG G&örseli": "png", "JPEG G&örseli": "jpeg",
                   "PDF &Belgesi": "pdf", "SVG &Vektörü": "svg"}
        for name, ext in formats.items():
            save_action = create_action(
                name, f'icons/save_{ext}.png', "",  # İkonlar için kontrol eklendi
                f"Grafiği .{ext} formatında kaydet",
                lambda checked, fmt=ext: self.save_graph(fmt)
            )
            save_as_menu.addAction(save_action)
        download_print_menu.addMenu(save_as_menu)
        download_print_menu.addSeparator()
        download_print_menu.addAction(create_action(
            "&Yazdır...", 'icons/print.png', "Ctrl+P",
            "Mevcut grafiği yazdır", self.print_graph
        ))

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

        if not self.column_names:
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
            # Numerik sütunları X ve Y eksenlerine ekle
            if pd.api.types.is_numeric_dtype(self.loaded_data_df[col]):
                self.x_axis_menu.addAction(self.create_column_selection_action("X", col))
                self.y_axis_menu.addAction(self.create_column_selection_action("Y", col))

            # Tüm sütunları kategori ve renk seçim menülerine ekle
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
        self.update_selection_info_label()

    def update_selection_info_label(self):
        """Seçili verilerin bilgisini gösteren etiketi günceller."""
        info = [
            f"X: {self.selected_x_column if self.selected_x_column is not None else 'Yok'}",
            f"Y: {self.selected_y_column if self.selected_y_column is not None else 'Yok'}",
            f"Kategori: {self.selected_category_column if self.selected_category_column is not None else 'Yok'}",
            f"Renk: {self.selected_color_column if self.selected_color_column is not None else 'Yok'}"
        ]
        # Eğer zaten bir grafik bilgisi varsa, onu koru
        current_plot_text = self.plot_info_label.text().split(' | ')[0]
        if current_plot_text.startswith("Grafik:"):
            self.plot_info_label.setText(f"{current_plot_text} | Seçili Veriler: {', '.join(info)}")
        else:
            self.plot_info_label.setText(f"Seçili Veriler: {', '.join(info)}")
        self.plot_info_label.adjustSize()

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
            # Sadece veri dosyaları için yükleme yap
            if file_type_code == "data":
                self.load_data_from_file(file_name)
                self.update_file_info_label(file_name, file_type_code)
            else:
                # Diğer dosya türleri için sadece bilgi mesajı
                self.update_file_info_label(file_name, file_type_code)
                QMessageBox.information(
                    self,
                    "Dosya Seçildi",
                    f"Dosya başarıyla seçildi:\n{os.path.basename(file_name)}"
                )

        except Exception as e:
            QMessageBox.critical(
                self,
                "Yükleme Hatası",
                f"Dosya yüklenirken hata oluştu:\n{str(e)}"
            )
            self.reset_data_selection(show_message=False)  # Hata durumunda da sıfırla

    def reset_data_selection(self, show_message=True):
        """Veri seçimlerini sıfırla ve isteğe bağlı mesaj göster"""
        self.update_file_info_label("", "")
        self.loaded_data_df = None
        self.column_names = []
        self.populate_data_selection_menus()  # Menüleri boşalt ve devre dışı bırak

        # Seçili sütunları sıfırla
        self.selected_x_column = None
        self.selected_y_column = None
        self.selected_category_column = None
        self.selected_color_column = None
        self.update_selection_info_label()  # Bilgi etiketini de sıfırla

        if show_message:
            QMessageBox.information(
                self,
                "Bilgi",
                "Dosya seçimi iptal edildi veya sıfırlandı."
            )

    def load_data_from_file(self, file_path):
        """Dosyadan veri yükler ve arayüzü günceller"""
        try:
            ext = os.path.splitext(file_path)[1].lower()

            if ext in ('.xlsx', '.xls'):
                df = pd.read_excel(file_path)
            elif ext == '.csv':
                df = pd.read_csv(file_path)
            else:
                raise ValueError("Desteklenmeyen dosya formatı")

            self.loaded_data_df = df
            self.column_names = list(df.columns)

            self.populate_data_selection_menus()
            self.set_default_column_selections()  # Varsayılan sütunları ayarla

            QMessageBox.information(
                self,
                "Başarılı",
                f"Veriler başarıyla yüklendi:\n{os.path.basename(file_path)}\n\n"
                f"Toplam {len(df)} satır, {len(df.columns)} sütun yüklendi."
            )

        except Exception as e:
            QMessageBox.critical(
                self,
                "Yükleme Hatası",
                f"Veri yüklenirken hata oluştu:\n{str(e)}"
            )
            self.reset_data_selection(show_message=False)  # Hata durumunda sıfırla

    def set_default_column_selections(self):
        """Yüklenen DataFrame'e göre varsayılan sütun seçimlerini yapar."""
        if self.loaded_data_df is None or self.loaded_data_df.empty:
            return

        numeric_cols = self.loaded_data_df.select_dtypes(include=np.number).columns.tolist()
        non_numeric_cols = self.loaded_data_df.select_dtypes(exclude=np.number).columns.tolist()

        # Varsayılan X ve Y ekseni atamaları
        if len(numeric_cols) >= 2:
            self.set_selected_column("X", numeric_cols[0])
            self.set_selected_column("Y", numeric_cols[1])
        elif len(numeric_cols) == 1:
            self.set_selected_column("X", numeric_cols[0])
            self.set_selected_column("Y", None)  # Y için başka numerik yok
        else:
            self.set_selected_column("X", None)
            self.set_selected_column("Y", None)

        # Varsayılan Kategori ve Renk atamaları
        if len(non_numeric_cols) >= 2:
            self.set_selected_column("Category", non_numeric_cols[0])
            self.set_selected_column("Color", non_numeric_cols[1])
        elif len(non_numeric_cols) == 1:
            self.set_selected_column("Category", non_numeric_cols[0])
            self.set_selected_column("Color", None)
        else:
            self.set_selected_column("Category", None)
            self.set_selected_column("Color", None)

        # Eğer numerik sütun yoksa ve kategorik sütunlar varsa, Y eksenine ilk kategorik sütunu atamayı dene
        # Bu, bar veya pie gibi grafikler için kullanışlı olabilir.
        if not numeric_cols and non_numeric_cols:
            if self.selected_x_column is None and self.selected_y_column is None:
                # Y eksenine kategorik bir sütun atamak genellikle sayısal bir değeri temsil etmez,
                # ancak bazı grafik türleri (bar, pie) için sayım veya toplam değeri gösterebilir.
                # Bu durumda, Y eksenini None bırakmak ve kullanıcının seçmesini beklemek daha iyi.
                pass

    def update_file_info_label(self, file_path, file_type):
        if file_path:
            base_name = os.path.basename(file_path)
            icon_path = f'icons/{file_type}.png'
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

    def update_plot_info_label(self, chart_type, count):
        """Grafik ve seçili veri bilgilerini gösteren etiketi günceller."""
        selected_info = f"X: {self.selected_x_column if self.selected_x_column is not None else 'Yok'}, " \
                        f"Y: {self.selected_y_column if self.selected_y_column is not None else 'Yok'}, " \
                        f"Kategori: {self.selected_category_column if self.selected_category_column is not None else 'Yok'}, " \
                        f"Renk: {self.selected_color_column if self.selected_color_column is not None else 'Yok'}"

        if chart_type and count > 0:
            plot_name = chart_type.split('(')[0].strip()
            new_text = f"Grafik: {plot_name} ({count} adet) | Seçili Veriler: {selected_info}"
            self.plot_info_label.setText(new_text)
            self.plot_info_label.setContentsMargins(5, 0, 0, 0)
        else:
            self.plot_info_label.setText(f"Seçili Veriler: {selected_info}")
            self.plot_info_label.setContentsMargins(0, 0, 0, 0)
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
                self.draw_graph(chart_type, num)
                # QMessageBox.information( # Çok fazla pop-up olmaması için kaldırıldı
                #     self, "Başarılı",
                #     f"'{chart_type}' türünde {num} adet grafik oluşturuldu."
                # )
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Grafik adedi alınamadı: {e}")

    def draw_graph(self, chart_type_text, count):
        self.matplotlib_widget.figure.clear()
        self.current_plot_data = []

        chart_type = chart_type_text.split('(')[-1][:-1] if '(' in chart_type_text else None
        if not chart_type:
            QMessageBox.warning(self, "Hata", "Geçersiz grafik türü")
            return

        # Gerekli sütunları kontrol et
        required = self.get_required_columns(chart_type)
        missing = [col for col in required if getattr(self, col) is None]  # None kontrolü

        if missing:
            self.show_missing_columns_error(missing, chart_type_text)
            return

        # Verileri hazırla
        try:
            plot_data = self.prepare_plot_data()
        except Exception as e:
            QMessageBox.critical(self, "Veri Hazırlama Hatası", f"Veri hazırlama sırasında hata oluştu: {str(e)}")
            return

        # Grafik düzenini ayarla
        rows, cols = self.calculate_grid_layout(count)
        self.matplotlib_widget.figure.set_size_inches(cols * 6, rows * 4.5)

        for i in range(count):
            ax = self.matplotlib_widget.figure.add_subplot(rows, cols, i + 1)
            title = f"{chart_type_text.split('(')[0].strip()} {i + 1}"
            ax.set_title(title, pad=20)

            plot_info = {
                'type': chart_type,
                'title': title,
                'xlabel': self.selected_x_column or 'X Değeri',  # Varsayılan değerler
                'ylabel': self.selected_y_column or 'Y Değeri'  # Varsayılan değerler
            }

            try:
                # Grafik türüne göre çizim yap
                draw_func = getattr(self, f"draw_{chart_type}", None)
                if draw_func:
                    draw_func(ax, plot_data.copy(), plot_info)  # plot_data'nın kopyasını gönder
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
        self.update_plot_info_label(chart_type_text, count)

    def calculate_grid_layout(self, count):
        if count == 1: return 1, 1
        if count == 2: return 1, 2
        if count == 3: return 1, 3
        if count == 4: return 2, 2
        if count == 5: return 2, 3  # 2x3 layout
        if count == 6: return 2, 3
        if count == 7: return 3, 3  # 3x3 layout
        if count == 8: return 3, 3
        if count == 9: return 3, 3
        if count == 10: return 3, 4  # Max 10 grafik için
        cols = min(4, count)  # Maksimum 4 sütun
        rows = (count + cols - 1) // cols
        return rows, cols

    def get_required_columns(self, chart_type):
        requirements = {
            'plot': ['selected_x_column', 'selected_y_column'],
            'bar': ['selected_category_column', 'selected_y_column'],  # Bar için y de gerekli
            'hist': ['selected_y_column'],
            'pie': ['selected_category_column', 'selected_y_column'],  # Pie için de y gerekli
            'scatter': ['selected_x_column', 'selected_y_column'],
            'fill_between': ['selected_x_column', 'selected_y_column'],
            'boxplot': ['selected_y_column'],
            'violinplot': ['selected_y_column'],
            'stem': ['selected_x_column', 'selected_y_column'],
            'errorbar': ['selected_x_column', 'selected_y_column']
        }
        return requirements.get(chart_type, [])

    def show_missing_columns_error(self, missing_columns, chart_type):
        missing_names = []
        for col_attr in missing_columns:
            if col_attr == 'selected_x_column':
                missing_names.append('X Ekseni')
            elif col_attr == 'selected_y_column':
                missing_names.append('Y Ekseni')
            elif col_attr == 'selected_category_column':
                missing_names.append('Kategori Verisi')
            elif col_attr == 'selected_color_column':
                missing_names.append('Renk Verisi')

        QMessageBox.critical(
            self, "Eksik Veri",
            f"{chart_type} için gerekli sütunlar seçilmemiş:\n\n" +
            "\n".join(f"- {name}" for name in missing_names) +
            "\n\nLütfen 'Veri Seç' menüsünden gerekli sütunları seçin."
        )

    def prepare_plot_data(self):
        """Yüklenen DataFrame'den seçili sütunlara göre verileri hazırlar."""
        if self.loaded_data_df is None:
            raise ValueError("Veri yüklenmedi.")

        data = {}

        if self.selected_x_column:
            data['x'] = self.loaded_data_df[self.selected_x_column].values
            # Sayısal olmayan değerleri NaN yap
            data['x'] = pd.to_numeric(data['x'], errors='coerce')
        else:
            data['x'] = np.arange(len(self.loaded_data_df))  # X seçilmemişse indeks kullan

        if self.selected_y_column:
            data['y'] = self.loaded_data_df[self.selected_y_column].values
            data['y'] = pd.to_numeric(data['y'], errors='coerce')
        else:
            data['y'] = np.zeros(len(self.loaded_data_df))  # Y seçilmemişse boş veri

        if self.selected_category_column:
            data['category'] = self.loaded_data_df[self.selected_category_column].values
        else:
            data['category'] = np.full(len(self.loaded_data_df), "Genel")  # Kategorik yoksa "Genel" kullan

        if self.selected_color_column:
            data['color'] = self.loaded_data_df[self.selected_color_column].values
            data['color'] = pd.to_numeric(data['color'], errors='ignore')  # Sayısal ise sayısal kalsın
        else:
            data['color'] = None

        # Eksik (NaN) verileri temizleme
        # Sadece x ve y'nin ortak geçerli indekslerini al
        valid_mask = ~np.isnan(data['x'])
        if data['y'] is not None:  # Y ekseni seçili ise onun da NaN kontrolünü yap
            valid_mask &= ~np.isnan(data['y'])

        for key in data:
            if data[key] is not None and len(data[key]) == len(valid_mask):
                data[key] = data[key][valid_mask]

        if data['category'] is not None:
            # Kategorik verileri işleme
            data['unique_categories'] = pd.Categorical(data['category']).categories.tolist()
            data['category_indices'] = pd.Categorical(data['category']).codes

        return data

    # Grafik çizim fonksiyonları (mevcut halleriyle bırakıldı, iyi görünüyorlar)
    def draw_plot(self, ax, data, plot_info):
        ax.plot(data['x'], data['y'], linewidth=2, marker='o', markersize=5)
        ax.set_xlabel(plot_info['xlabel'])
        ax.set_ylabel(plot_info['ylabel'])
        ax.grid(True, linestyle='--', alpha=0.6)

        # Trend çizgisi ekle
        if len(data['x']) > 1 and np.issubdtype(data['x'].dtype, np.number) and np.issubdtype(data['y'].dtype,
                                                                                              np.number):
            # Sadece sayısal veriler için trend çizgisi hesapla
            z = np.polyfit(data['x'], data['y'], 1)
            p = np.poly1d(z)
            ax.plot(data['x'], p(data['x']), "r--", linewidth=1, label='Trend')
            ax.legend()
            plot_info['trend_line'] = p(data['x']).tolist()
        else:
            plot_info['trend_line'] = None

        plot_info.update({
            'data_x': data['x'].tolist(),
            'data_y': data['y'].tolist(),
        })

    def draw_bar(self, ax, data, plot_info):
        if data['category'] is None or data['y'] is None:
            ax.text(0.5, 0.5, "Bar grafiği için kategori ve değer sütunları seçilmeli.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        # Pandas Series kullanarak group_by ve sum yap
        df_temp = pd.DataFrame({'category': data['category'], 'value': data['y']})
        grouped_data = df_temp.groupby('category')['value'].sum().reset_index()

        categories = grouped_data['category'].tolist()
        values = grouped_data['value'].tolist()

        colors = plt.cm.viridis(np.linspace(0, 1, len(categories)))
        bars = ax.bar(categories, values, color=colors)

        # Değerleri çubukların üzerine yaz
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2., height,
                    f'{height:.1f}', ha='center', va='bottom')

        ax.set_xlabel(plot_info['xlabel'])
        ax.set_ylabel(plot_info['ylabel'])
        ax.tick_params(axis='x', rotation=45)

        plot_info.update({
            'categories': categories,
            'values': values,
            'xlabel': self.selected_category_column,
            'ylabel': self.selected_y_column
        })

    def draw_hist(self, ax, data, plot_info):
        if data['y'] is None:
            ax.text(0.5, 0.5, "Histogram için değer sütunu seçilmeli.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        # Sadece sayısal verileri histograma al
        numeric_y = data['y'][np.issubdtype(data['y'].dtype, np.number)]
        if len(numeric_y) == 0:
            ax.text(0.5, 0.5, "Histogram için sayısal veri bulunamadı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        n, bins, patches = ax.hist(numeric_y, bins='auto', color='#607c8e',
                                   edgecolor='black', alpha=0.7)

        # Yoğunluk eğrisi ekle (sadece yeterli veri varsa)
        if len(numeric_y) > 1:
            density = stats.gaussian_kde(numeric_y)
            xs = np.linspace(min(numeric_y), max(numeric_y), 200)
            density._compute_covariance()
            ax2 = ax.twinx()
            ax2.plot(xs, density(xs), color='darkred', linewidth=2)
            ax2.set_ylabel('Yoğunluk', color='darkred')
            ax2.tick_params(axis='y', labelcolor='darkred')

        ax.set_xlabel(plot_info['ylabel'])
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
        if data['category'] is None or data['y'] is None:
            ax.text(0.5, 0.5, "Pasta grafiği için kategori ve değer sütunları seçilmeli.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        df_temp = pd.DataFrame({'category': data['category'], 'value': data['y']})
        # Pie grafiği için kategori bazında toplam değerleri al
        grouped_data = df_temp.groupby('category')['value'].sum().reset_index()

        sizes = grouped_data['value'].tolist()
        labels = grouped_data['category'].tolist()

        ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90)
        ax.axis('equal')  # Oranların daire şeklinde olmasını sağlar

        plot_info.update({
            'sizes': sizes,
            'labels': labels,
            'xlabel': self.selected_category_column,
            'ylabel': self.selected_y_column
        })

    def draw_scatter(self, ax, data, plot_info):
        if data['x'] is None or data['y'] is None:
            ax.text(0.5, 0.5, "Dağılım grafiği için X ve Y sütunları seçilmeli.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        colors = data['color'] if data['color'] is not None else 'blue'
        cmap = 'viridis' if data['color'] is not None else None

        # Boyutlar için rastgelelik yerine sabit bir boyut veya başka bir sütun kullanılabilir.
        # Şimdilik rastgele boyut bırakıldı.
        sizes = np.random.rand(len(data['x'])) * 200 + 20 if len(data['x']) > 0 else []

        scatter = ax.scatter(data['x'], data['y'], c=colors, s=sizes,
                             alpha=0.7, cmap=cmap)
        ax.set_xlabel(plot_info['xlabel'])
        ax.set_ylabel(plot_info['ylabel'])

        if cmap:
            plt.colorbar(scatter, ax=ax, label=self.selected_color_column)

        plot_info.update({
            'data_x': data['x'].tolist(),
            'data_y': data['y'].tolist(),
            'colors': colors.tolist() if hasattr(colors, 'tolist') else (colors if isinstance(colors, str) else None),
            'sizes': sizes.tolist() if hasattr(sizes, 'tolist') else []
        })

    def draw_fill_between(self, ax, data, plot_info):
        if data['x'] is None or data['y'] is None:
            ax.text(0.5, 0.5, "Alan grafiği için X ve Y sütunları seçilmeli.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        y1 = data['y']
        # y2'yi y1'in bir fonksiyonu olarak belirle, negatif değerleri de destekleyecek şekilde
        y2 = y1 - (np.mean(y1) / 2) if len(y1) > 0 else np.zeros_like(y1)

        ax.plot(data['x'], y1, label='Veri Hattı')
        ax.fill_between(data['x'], y1, y2, color='skyblue', alpha=0.4)
        ax.set_xlabel(plot_info['xlabel'])
        ax.set_ylabel(plot_info['ylabel'])
        ax.legend()

        plot_info.update({
            'data_x': data['x'].tolist(),
            'data_y1': y1.tolist(),
            'data_y2': y2.tolist()
        })

    def draw_boxplot(self, ax, data, plot_info):
        if data['y'] is None:
            ax.text(0.5, 0.5, "Kutu grafiği için değer sütunu seçilmeli.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        # Sadece sayısal verileri al
        numeric_y = data['y'][np.issubdtype(data['y'].dtype, np.number)]
        if len(numeric_y) == 0:
            ax.text(0.5, 0.5, "Kutu grafiği için sayısal veri bulunamadı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        groups = []
        labels = []

        if self.selected_category_column and data['category'] is not None:
            df_temp = pd.DataFrame({'category': data['category'], 'value': numeric_y})
            for cat in df_temp['category'].unique():
                groups.append(df_temp[df_temp['category'] == cat]['value'].tolist())
                labels.append(str(cat))  # Label'ları string yap
        else:
            groups = [numeric_y.tolist()]
            labels = [plot_info['ylabel']] if plot_info['ylabel'] else ['Değerler']

        # Boş grupları filtrele
        groups = [g for g in groups if g]
        labels = [l for i, l in enumerate(labels) if groups[i]]

        if not groups:
            ax.text(0.5, 0.5, "Kutu grafiği için geçerli grup verisi bulunamadı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        ax.boxplot(groups, patch_artist=True)
        ax.set_xticklabels(labels, rotation=45, ha='right')  # Etiketleri döndür
        ax.set_ylabel(plot_info['ylabel'])

        plot_info.update({
            'data_groups': groups,
            'labels': labels,
            'xlabel': self.selected_category_column or 'Gruplar'
        })

    def draw_violinplot(self, ax, data, plot_info):
        if data['y'] is None:
            ax.text(0.5, 0.5, "Violin grafiği için değer sütunu seçilmeli.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        # Sadece sayısal verileri al
        numeric_y = data['y'][np.issubdtype(data['y'].dtype, np.number)]
        if len(numeric_y) == 0:
            ax.text(0.5, 0.5, "Violin grafiği için sayısal veri bulunamadı.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        groups = []
        labels = []

        if self.selected_category_column and data['category'] is not None:
            df_temp = pd.DataFrame({'category': data['category'], 'value': numeric_y})
            for cat in df_temp['category'].unique():
                groups.append(df_temp[df_temp['category'] == cat]['value'].tolist())
                labels.append(str(cat))
        else:
            groups = [numeric_y.tolist()]
            labels = [plot_info['ylabel']] if plot_info['ylabel'] else ['Değerler']

        groups = [g for g in groups if g]
        labels = [l for i, l in enumerate(labels) if groups[i]]

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
            'xlabel': self.selected_category_column or 'Gruplar'
        })

    def draw_stem(self, ax, data, plot_info):
        if data['x'] is None or data['y'] is None:
            ax.text(0.5, 0.5, "Stem grafiği için X ve Y sütunları seçilmeli.", ha='center', va='center',
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
        if data['x'] is None or data['y'] is None:
            ax.text(0.5, 0.5, "Hata çubuklu grafiği için X ve Y sütunları seçilmeli.", ha='center', va='center',
                    transform=ax.transAxes, color='red')
            return

        y_err = 0.1 * np.abs(data['y'])  # %10 hata payı
        ax.errorbar(data['x'], data['y'], yerr=y_err, fmt='-o',
                    capsize=5, capthick=2)
        ax.set_xlabel(plot_info['xlabel'])
        ax.set_ylabel(plot_info['ylabel'])

        plot_info.update({
            'data_x': data['x'].tolist(),
            'data_y': data['y'].tolist(),
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
            pdf.set_font("NotoSans", "B", 16)  # Daha büyük başlık
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

            # Resmi PDF'e ekle (en fazla 180mm genişlik)
            img_width_pt = self.matplotlib_widget.figure.get_size_inches()[0] * 72  # inç'ten point'e
            img_height_pt = self.matplotlib_widget.figure.get_size_inches()[1] * 72

            # PDF sayfa genişliği (210mm = 595.276pt)
            pdf_page_width = 210 * 2.83465  # mm'den point'e

            # Orantılı ölçekleme
            if img_width_pt > (pdf_page_width - 20):  # Kenar boşlukları için 20pt bırak
                scale_factor = (pdf_page_width - 20) / img_width_pt
                img_width_pt *= scale_factor
                img_height_pt *= scale_factor

            pdf.image(temp_img, x=XPos.CENTER, w=img_width_pt / 2.83465)  # fpdf w değeri mm cinsinden istiyor
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
                    pdf.cell(0, 6, f"X Aralığı: [{min(plot['data_x']):.2f}, {max(plot['data_x']):.2f}]", 0, 1)
                    pdf.cell(0, 6, f"Y Aralığı: [{min(plot['data_y']):.2f}, {max(plot['data_y']):.2f}]", 0, 1)
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
            if 'data_y' in plot and plot['data_y'] is not None:
                all_numeric_data.extend([x for x in plot['data_y'] if isinstance(x, (int, float)) and not np.isnan(x)])
            elif 'data' in plot and plot['data'] is not None:  # hist için
                all_numeric_data.extend([x for x in plot['data'] if isinstance(x, (int, float)) and not np.isnan(x)])
            elif 'values' in plot and plot['values'] is not None:  # bar için
                all_numeric_data.extend([x for x in plot['values'] if isinstance(x, (int, float)) and not np.isnan(x)])
            elif 'sizes' in plot and plot['sizes'] is not None:  # pie için
                all_numeric_data.extend([x for x in plot['sizes'] if isinstance(x, (int, float)) and not np.isnan(x)])

        if not all_numeric_data:
            pdf.multi_cell(0, 6, "İstatistiksel analiz için yeterli sayısal veri bulunamadı.")
            return

        all_numeric_data_np = np.array(all_numeric_data)

        # Temel istatistikler
        pdf.cell(0, 6, f"Toplam Sayısal Veri Noktası Sayısı: {len(all_numeric_data_np)}", 0, 1)
        pdf.cell(0, 6, f"Minimum Değer: {np.min(all_numeric_data_np):.2f}", 0, 1)
        pdf.cell(0, 6, f"Maksimum Değer: {np.max(all_numeric_data_np):.2f}", 0, 1)
        pdf.cell(0, 6, f"Ortalama: {np.mean(all_numeric_data_np):.2f}", 0, 1)
        pdf.cell(0, 6, f"Medyan: {np.median(all_numeric_data_np):.2f}", 0, 1)
        pdf.cell(0, 6, f"Standart Sapma: {np.std(all_numeric_data_np):.2f}", 0, 1)

        # Çarpıklık ve basıklık için en az 3 veri noktası gerekir
        if len(all_numeric_data_np) >= 3:
            pdf.cell(0, 6, f"Çarpıklık: {stats.skew(all_numeric_data_np):.2f}", 0, 1)
            pdf.cell(0, 6, f"Basıklık: {stats.kurtosis(all_numeric_data_np):.2f}", 0, 1)
        else:
            pdf.cell(0, 6, "Çarpıklık ve Basıklık hesaplamak için yeterli veri yok.", 0, 1)
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

        if len(plot_type_data) < 2:
            try:
                pdf.set_font("NotoSans", "", 10)
            except Exception:
                pdf.set_font("Arial", "", 10)
            pdf.cell(0, 6, "Korelasyon analizi için en az iki uygun grafik bulunamadı (Çizgi, Dağılım vb.).", 0, 1)
            return

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

        try:
            # Sadece ilk iki uygun grafiğin Y verilerini karşılaştır
            # Boyutları eşleşen verileri al

            data1_y = np.array(plot_type_data[0]['data_y'])
            data2_y = np.array(plot_type_data[1]['data_y'])

            # Uzunlukları eşitlemek için en kısa olanı referans al
            min_len = min(len(data1_y), len(data2_y))
            data1_y = data1_y[:min_len]
            data2_y = data2_y[:min_len]

            # Sadece sayısal ve NaN olmayan değerleri al
            valid_mask = ~np.isnan(data1_y) & ~np.isnan(data2_y)
            data1_y_clean = data1_y[valid_mask]
            data2_y_clean = data2_y[valid_mask]

            if len(data1_y_clean) < 2:  # Korelasyon için en az 2 veri noktası gerekir
                pdf.cell(0, 6, "Korelasyon analizi için yeterli sayıda ortak sayısal veri noktası bulunamadı.", 0, 1)
                return

            corr = np.corrcoef(data1_y_clean, data2_y_clean)[0, 1]
            pdf.cell(0, 6,
                     f"{plot_type_data[0]['title']} ve {plot_type_data[1]['title']} verileri arasındaki korelasyon: {corr:.3f}",
                     0, 1)

            if abs(corr) > 0.7:
                pdf.cell(0, 6, "→ Güçlü bir ilişki var", 0, 1)
            elif abs(corr) > 0.3:
                pdf.cell(0, 6, "→ Orta düzeyde ilişki var", 0, 1)
            else:
                pdf.cell(0, 6, "→ Zayıf veya ilişki yok", 0, 1)

            # Regresyon analizi (en az 2 veri noktası gerekir)
            if len(data1_y_clean) >= 2:
                slope, intercept = np.polyfit(data1_y_clean, data2_y_clean, 1)
                pdf.cell(0, 6,
                         f"Regresyon denklemi: y = {slope:.3f}x + {intercept:.3f}",
                         0, 1)
        except Exception as e:
            pdf.cell(0, 6, f"Korelasyon analizi yapılamadı: {str(e)}", 0, 1)
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
                # Arka plan rengini kaydetme sırasında da koru
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
    # Buraya örnek ikon dosyaları eklenmeli veya kullanıcıya bildirilmelidir.
    # Örneğin: 'icons/data_file.png', 'icons/exit.png', 'icons/save_pdf.png' vb.

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())