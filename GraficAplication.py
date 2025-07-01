import sys
import os
import random
import numpy as np

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QAction, QMenu, QInputDialog,
    QMessageBox, QFileDialog, QLabel, QVBoxLayout, QWidget, QScrollArea # QScrollArea eklendi
)
from PyQt5.QtGui import QIcon, QFont, QPixmap
from PyQt5.QtCore import Qt, QSize

# Matplotlib entegrasyonu için gerekli importlar
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar

# Matplotlib varsayılan arka planını daha koyu bir temaya uyduralım
plt.style.use('dark_background') # Matplotlib grafiklerinin arka planını koyu yap
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
    """
    Matplotlib grafiğini içinde barındıracak özel bir QWidget.
    Bu widget bir kaydırma alanının içine yerleştirilecek.
    """
    def __init__(self, parent=None):
        super().__init__(parent)

        self.figure = plt.figure() # Matplotlib figürü oluştur
        self.canvas = FigureCanvas(self.figure) # Figürü bir Qt widget'ına dönüştür
        self.toolbar = NavigationToolbar(self.canvas, self) # Navigasyon araç çubuğu

        # Dikey bir layout oluştur ve canvas ile araç çubuğunu ekle
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.toolbar)
        self.layout.addWidget(self.canvas)
        self.setLayout(self.layout)

        # Plot'u temizle ve ilk boş grafiği çiz
        self.figure.clear()
        self.canvas.draw_idle()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Grafik Uygulaması")
        self.setGeometry(100, 100, 1024, 768)

        # Ana layout ve merkezi widget oluşturma
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        # Dosya bilgisi etiketini oluştur
        self.file_info_label = QLabel("Lütfen bir dosya seçin...", self)
        self.file_info_label.setAlignment(Qt.AlignCenter)
        self.file_info_label.setFont(QFont("Arial", 12))
        self.file_info_label.setObjectName("fileInfoLabel")
        self.file_info_label.setFixedHeight(30)

        # Grafik bilgisi etiketini oluştur
        self.plot_info_label = QLabel("Henüz bir grafik oluşturulmadı.", self)
        self.plot_info_label.setAlignment(Qt.AlignCenter)
        self.plot_info_label.setFont(QFont("Arial", 12))
        self.plot_info_label.setObjectName("plotInfoLabel")
        self.plot_info_label.setFixedHeight(30)

        # QLabel'leri ana layout'a ekle
        self.main_layout.addWidget(self.file_info_label)
        self.main_layout.addWidget(self.plot_info_label)

        # --- QScrollArea ve İçerik Widget'ı ---
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True) # İçindeki widget'ın boyutunu otomatik ayarla

        # İçerik widget'ı oluştur, MplWidget'ı ve onun layout'unu bu widget'a yerleştireceğiz.
        self.scroll_content_widget = QWidget()
        self.scroll_content_layout = QVBoxLayout(self.scroll_content_widget)

        # MplWidget'ı içerik widget'ının layout'una ekle
        self.matplotlib_widget = MplWidget(self)
        self.scroll_content_layout.addWidget(self.matplotlib_widget)

        # QScrollArea'nın widget'ını ayarla
        self.scroll_area.setWidget(self.scroll_content_widget)

        # QScrollArea'yı ana layout'a ekle
        self.main_layout.addWidget(self.scroll_area)
        # -------------------------------------------

        self.create_menu()

        # Uygulama genel stilini ayarla
        self.setStyleSheet("""
            QMainWindow {
                background-color: #202020; /* Ana pencere arka planı */
            }
            QMenuBar {
                background-color: #333;
                color: #FFF;
                font-size: 14px;
                padding: 5px 0px;
            }
            QMenuBar::item {
                padding: 8px 15px;
                background-color: transparent;
            }
            QMenuBar::item:selected {
                background-color: #555;
            }
            QMenu {
                background-color: #444;
                color: #FFF;
                border: 1px solid #666;
            }
            QMenu::item {
                padding: 8px 25px;
                background-color: transparent;
            }
            QMenu::item:selected {
                background-color: #666;
            }
            QMenu::separator {
                height: 1px;
                background-color: #555;
                margin-left: 10px;
                margin-right: 10px;
            }
            QLabel#fileInfoLabel {
                color: #ADD8E6; /* Açık mavi */
                font-weight: bold;
                padding: 5px;
                background-color: #282828; /* Koyu gri */
                border-bottom: 1px solid #444;
            }
            QLabel#plotInfoLabel {
                color: #A0FFA0; /* Açık yeşil */
                font-weight: bold;
                padding: 5px;
                background-color: #3A3A3A; /* Biraz daha açık koyu gri */
                border-bottom: 1px solid #555;
            }
            QScrollArea {
                border: none; /* Kaydırma alanı çerçevesini kaldır */
                background-color: #282828; /* İçeriğin arka planını ayarlayabiliriz */
            }
            /* QScrollBar stilleri */
            QScrollBar:vertical {
                border: 1px solid #555;
                background: #333;
                width: 15px;
                margin: 15px 0 15px 0;
            }
            QScrollBar::handle:vertical {
                background: #666;
                min-height: 20px;
                border-radius: 5px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                background: none;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }
            QScrollBar:horizontal {
                border: 1px solid #555;
                background: #333;
                height: 15px;
                margin: 0 15px 0 15px;
            }
            QScrollBar::handle:horizontal {
                background: #666;
                min-width: 20px;
                border-radius: 5px;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                background: none;
            }
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: none;
            }
        """)

    def create_menu(self):
        menubar = self.menuBar()

        # 1. Dosya Menüsü
        file_menu = menubar.addMenu("&Dosya")

        # İkonları yüklerken hata kontrolü ekleyelim
        def create_action_with_icon(text, shortcut, status_tip, icon_path, connect_func):
            try:
                action = QAction(QIcon(icon_path), text, self)
            except Exception as e:
                print(f"Uyarı: İkon yüklenirken hata oluştu ({icon_path}): {e}")
                action = QAction(text, self) # İkon olmadan oluştur
            action.setShortcut(shortcut)
            action.setStatusTip(status_tip)
            action.triggered.connect(connect_func)
            return action

        file_menu.addAction(create_action_with_icon(
            "Word Dosyası &Aç...", "Ctrl+W", "Bir Word belgesini açar",
            'icons/word.png', lambda: self.open_file("Word Dosyaları (*.docx *.doc);;Tüm Dosyalar (*)", "word")
        ))
        file_menu.addAction(create_action_with_icon(
            "Excel Dosyası &Aç...", "Ctrl+E", "Bir Excel çalışma sayfasını açar",
            'icons/excel.png', lambda: self.open_file("Excel Dosyaları (*.xlsx *.xls);;Tüm Dosyalar (*)", "excel")
        ))
        file_menu.addAction(create_action_with_icon(
            "PPTX Dosyası &Aç...", "Ctrl+P", "Bir PowerPoint sunumunu açar",
            'icons/pptx.png', lambda: self.open_file("PowerPoint Dosyaları (*.pptx *.ppt);;Tüm Dosyalar (*)", "pptx")
        ))

        file_menu.addSeparator()

        file_menu.addAction(create_action_with_icon(
            "Çı&kış", "Ctrl+Q", "Uygulamadan çıkar",
            'icons/exit.png', self.close
        ))

        # 2. Grafik Oluştur Menüsü
        plot_menu = menubar.addMenu("&Grafik Oluştur")

        self.chart_types = [
            "Çizgi Grafiği (plot)", "Bar Grafiği (bar)", "Histogram (hist)",
            "Pasta Grafiği (pie)", "Dağılım Grafiği (scatter)", "Alan Grafiği (fill_between)",
            "Kutu Grafiği (boxplot)", "Violin Grafiği (violinplot)", "Stem Grafiği (stem)",
            "Hata Çubuklu Grafik (errorbar)"
        ]

        for grafik_adı in self.chart_types:
            action = QAction(grafik_adı + "...", self)
            action.setStatusTip(f"'{grafik_adı}' grafiği için adet seçimi")
            action.triggered.connect(lambda checked, name=grafik_adı: self.get_plot_count(name))
            plot_menu.addAction(action)

        # 3. İndir / Yazdır Menüsü
        download_print_menu = menubar.addMenu("&İndir / Yazdır")

        save_as_menu = QMenu("Farklı Kaydet", self)

        formats = {"PNG G&örseli": "png", "JPEG G&örseli": "jpeg", "PDF &Belgesi": "pdf", "SVG &Vektörü": "svg"}
        for name, file_ext in formats.items():
            save_action = create_action_with_icon(
                name, "", f"Grafiği .{file_ext} formatında kaydet",
                f'icons/save_{file_ext}.png', lambda checked, fmt=file_ext: self.save_graph(fmt)
            )
            save_as_menu.addAction(save_action)

        download_print_menu.addMenu(save_as_menu)

        download_print_menu.addSeparator()

        download_print_menu.addAction(create_action_with_icon(
            "&Yazdır...", "Ctrl+P", "Mevcut grafiği yazdır",
            'icons/print.png', self.print_graph
        ))

        # 4. Veri Seç Menüsü
        data_menu = menubar.addMenu("&Veri Seç")

        x_axis_menu = QMenu("X Ekseni &Seç", self)
        x1_action = QAction("X1 Verisi", self)
        x2_action = QAction("X2 Verisi", self)
        x_axis_menu.addAction(x1_action)
        x_axis_menu.addAction(x2_action)
        data_menu.addMenu(x_axis_menu)

        y_axis_menu = QMenu("Y Ekseni &Seç", self)
        y1_action = QAction("Y1 Verisi", self)
        y2_action = QAction("Y2 Verisi", self)
        y_axis_menu.addAction(y1_action)
        y_axis_menu.addAction(y2_action)
        data_menu.addMenu(y_axis_menu)

        data_menu.addSeparator()

        category_selection_action = QAction("Kategori Verisi &Seç", self)
        category_selection_action.setStatusTip("Kategori bazlı grafikler için veri sütunu seçer")
        data_menu.addAction(category_selection_action)

        color_by_data_action = QAction("Veriye Göre &Renklendir", self)
        color_by_data_action.setStatusTip("Dağılım grafiklerinde noktaları bir veri sütununa göre renklendirir")
        data_menu.addAction(color_by_data_action)

    def get_plot_count(self, chart_type):
        try:
            num, ok = QInputDialog.getInt(self, "Grafik Adedi Girin",
                                          f"Kaç adet '{chart_type}' grafiği oluşturmak istersiniz?",
                                          min=1, max=10, step=1)

            if ok:
                self.draw_graph(chart_type, num)
                self.update_plot_info_label(chart_type, num)
                QMessageBox.information(self, "Grafik Oluşturuldu",
                                        f"'{chart_type}' türünde {num} adet grafik başarıyla oluşturuldu.")
            else:
                QMessageBox.information(self, "İptal Edildi", "Grafik oluşturma işlemi iptal edildi.")
                self.update_plot_info_label("", 0)
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Grafik adedi alınırken bir hata oluştu: {e}")
            self.update_plot_info_label("", 0)


    def save_graph(self, file_format):
        if not self.matplotlib_widget.figure.axes:
            QMessageBox.warning(self, "Uyarı", "Kaydedilecek bir grafik bulunamadı. Lütfen önce bir grafik oluşturun.")
            return

        # QFileDialog'ı bir try-except bloğuna alalım
        try:
            file_name, _ = QFileDialog.getSaveFileName(self, "Grafiği Kaydet", "grafik",
                                                       f"Grafik Dosyaları (*.{file_format});;Tüm Dosyalar (*)")

            if file_name:
                # Matplotlib figürünü belirtilen dosya formatında kaydet
                self.matplotlib_widget.figure.tight_layout() # Kaydetmeden önce son düzenleme
                self.matplotlib_widget.figure.savefig(file_name, format=file_format)
                QMessageBox.information(self, "Kaydetme Başarılı",
                                        f"Grafik '{file_name}' olarak kaydedildi.")
            else:
                QMessageBox.information(self, "Kaydetme İptal Edildi", "Kaydetme işlemi iptal edildi.")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Grafik kaydedilirken bir hata oluştu: {e}\n"
                                               f"Lütfen dosya yolunun geçerli olduğundan ve yazma izniniz olduğundan emin olun.")


    def print_graph(self):
        QMessageBox.information(self, "Yazdırma İşlemi",
                                "Grafik yazdırma işlemi başlatılacak. (Henüz tam işlevsel değil)")

    def open_file(self, file_filter, file_type_code):
        try:
            file_name, _ = QFileDialog.getOpenFileName(self, "Dosya Seç", "", file_filter)

            if file_name:
                self.update_file_info_label(file_name, file_type_code)
                self.update_plot_info_label("", 0) # Dosya seçildiğinde grafik bilgisini temizle
                QMessageBox.information(self, "Dosya Seçimi",
                                        f"Seçilen Dosya: {file_name}")
            else:
                self.update_file_info_label("", "")
                self.update_plot_info_label("", 0)
                QMessageBox.information(self, "İptal Edildi", "Dosya seçimi iptal edildi.")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Dosya seçimi sırasında bir hata oluştu: {e}")
            self.update_file_info_label("", "")
            self.update_plot_info_label("", 0)


    def update_file_info_label(self, file_path, file_type_code):
        if file_path:
            base_name = os.path.basename(file_path)
            icon_path = f'icons/{file_type_code}.png'

            pixmap = QPixmap(icon_path)
            if not pixmap.isNull():
                pixmap = pixmap.scaled(24, 24, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                self.file_info_label.setPixmap(pixmap)
                self.file_info_label.setText(f"  {base_name}")
                self.file_info_label.setContentsMargins(5, 0, 0, 0)
            else:
                self.file_info_label.setPixmap(QPixmap())
                # Eğer ikon bulunamazsa uyarı yazısını göster
                self.file_info_label.setText(f"  {base_name} (İkon yüklenemedi)")
                print(f"Uyarı: '{icon_path}' ikonu bulunamadı veya yüklenemedi.") # Konsola da yazdır
        else:
            self.file_info_label.setPixmap(QPixmap())
            self.file_info_label.setText("Lütfen bir dosya seçin...")

        self.file_info_label.adjustSize()
        self.file_info_label.setFixedWidth(self.width())

    def update_plot_info_label(self, chart_type_text, count):
        if chart_type_text and count > 0:
            self.plot_info_label.setText(f"  Grafik: {chart_type_text.split('(')[0].strip()} ({count} adet)")
            self.plot_info_label.setContentsMargins(5, 0, 0, 0)
        else:
            self.plot_info_label.setText("Henüz bir grafik oluşturulmadı.")
            self.plot_info_label.setContentsMargins(0, 0, 0, 0)

        self.plot_info_label.adjustSize()
        self.plot_info_label.setFixedWidth(self.width())

    def draw_graph(self, chart_type_text, count):
        # Önceki grafik figürünü temizle
        self.matplotlib_widget.figure.clear()

        # Grafik adetine göre alt grafikler (subplots) oluşturmak için satır ve sütun hesapla
        figure_height_per_row = 5 # Her satır grafik için yaklaşık yükseklik (inç cinsinden)

        rows = 1
        cols = 1
        if count == 1:
            rows, cols = 1, 1
        elif count == 2:
            rows, cols = 1, 2
        elif count == 3:
            rows, cols = 1, 3
        elif count == 4:
            rows, cols = 2, 2
        elif count == 5 or count == 6:
            rows, cols = 2, 3
        elif count > 6:
            cols = 3
            rows = (count + cols - 1) // cols

        # Figürün genişliği ve yüksekliğini inç cinsinden ayarla
        # Matplotlib'in dpi'ı (dots per inch) genellikle 100'dür.
        # Bu, boyutları piksel cinsine çevirmek için kullanılır.
        self.matplotlib_widget.figure.set_size_inches(cols * 4, rows * figure_height_per_row)

        chart_func_name = chart_type_text.split('(')[-1][:-1] if '(' in chart_type_text else None

        if chart_func_name is None:
            QMessageBox.warning(self, "Hata", "Geçersiz grafik türü adı.")
            return

        for i in range(count):
            try:
                ax = self.matplotlib_widget.figure.add_subplot(rows, cols, i + 1)
                ax.set_title(f"{chart_type_text.split('(')[0].strip()} {i + 1}")

                # Örnek veri üretimi
                x = np.linspace(0, 10, 100)
                y = np.sin(x + i * 0.5) + random.uniform(-0.5, 0.5)

                # Grafik türüne göre çizim
                if chart_func_name == "plot":
                    ax.plot(x, y, label=f'Series {i + 1}')
                    ax.set_xlabel("X Ekseni")
                    ax.set_ylabel("Y Ekseni")
                    ax.legend()
                elif chart_func_name == "bar":
                    categories = [f'Kategori {k + 1}' for k in range(5)]
                    values = np.random.randint(5, 20, 5)
                    colors = [plt.cm.viridis(k / (len(categories) - 1)) for k in range(len(categories))]
                    ax.bar(categories, values, color=colors)
                    ax.set_xlabel("Kategoriler")
                    ax.set_ylabel("Değerler")
                elif chart_func_name == "hist":
                    data = np.random.randn(1000)
                    ax.hist(data, bins=30, color='skyblue', edgecolor='black')
                    ax.set_xlabel("Değer")
                    ax.set_ylabel("Frekans")
                elif chart_func_name == "pie":
                    sizes = [random.randint(10, 30) for _ in range(4)]
                    labels = [f'Dilim {k + 1}' for k in range(4)]
                    ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90)
                    ax.axis('equal')
                elif chart_func_name == "scatter":
                    x_scatter = np.random.rand(50) * 10
                    y_scatter = np.random.rand(50) * 10
                    colors = np.random.rand(50)
                    sizes = np.random.rand(50) * 200 + 20
                    ax.scatter(x_scatter, y_scatter, c=colors, s=sizes, alpha=0.7, cmap='viridis')
                    ax.set_xlabel("X Verisi")
                    ax.set_ylabel("Y Verisi")
                elif chart_func_name == "fill_between":
                    x_area = np.linspace(0, 10, 100)
                    y1_area = np.sin(x_area + i * 0.5) + 2
                    y2_area = np.cos(x_area + i * 0.5) + 1
                    ax.plot(x_area, y1_area, label='Üst Sınır')
                    ax.plot(x_area, y2_area, label='Alt Sınır', linestyle='--')
                    ax.fill_between(x_area, y1_area, y2_area, color='skyblue', alpha=0.4)
                    ax.set_xlabel("X Ekseni")
                    ax.set_ylabel("Değer")
                    ax.legend()
                elif chart_func_name == "boxplot":
                    data = [np.random.normal(0, std, 100) for std in range(1, 4)]
                    ax.boxplot(data, patch_artist=True)
                    ax.set_xticklabels([f'Grup {k + 1}' for k in range(len(data))])
                    ax.set_ylabel("Değer")
                elif chart_func_name == "violinplot":
                    data = [np.random.normal(0, std, 100) for std in range(1, 4)]
                    ax.violinplot(data, showmeans=True, showmedians=True)
                    ax.set_xticks(np.arange(1, len(data) + 1))
                    ax.set_xticklabels([f'Grup {k + 1}' for k in range(len(data))])
                    ax.set_ylabel("Değer")
                elif chart_func_name == "stem":
                    x_stem = np.arange(10)
                    y_stem = np.random.randint(1, 10, 10)
                    ax.stem(x_stem, y_stem)
                    ax.set_xlabel("Dizin")
                    ax.set_ylabel("Değer")
                elif chart_func_name == "errorbar":
                    x_err = np.linspace(0, 10, 10)
                    y_err = np.sin(x_err)
                    y_error = 0.1 + 0.2 * np.random.rand(len(x_err))
                    ax.errorbar(x_err, y_err, yerr=y_error, fmt='-o')
                    ax.set_xlabel("X Ekseni")
                    ax.set_ylabel("Y Ekseni")
                else:
                    ax.text(0.5, 0.5, f"'{chart_type_text}' için çizim kodu yok.",
                            horizontalalignment='center', verticalalignment='center', transform=ax.transAxes)
                    ax.set_title("Bilinmeyen Grafik Türü")

            except Exception as e:
                # Bir grafiğin çizimi sırasında hata oluşursa, bunu kullanıcıya bildir
                print(f"Hata: {i+1}. grafiği çizerken bir sorun oluştu: {e}")
                ax.text(0.5, 0.5, f"Grafik çiziminde hata:\n{e}",
                        horizontalalignment='center', verticalalignment='center',
                        transform=ax.transAxes, color='red', fontsize=12)
                ax.set_title(f"Hata Oluştu ({chart_type_text.split('(')[0].strip()} {i + 1})")


        # Grafiklerin düzenini optimize et
        self.matplotlib_widget.figure.tight_layout()
        # Kaydırma alanının içeriğini güncellemek için canvas'ı yeniden çiz
        self.matplotlib_widget.canvas.draw_idle()
        # Kaydırma alanının widget'ının boyutunu içeriğine göre yeniden ayarla
        # Bu, kaydırma çubuklarının doğru görünmesini sağlar.
        self.scroll_content_widget.setMinimumSize(
            int(self.matplotlib_widget.figure.get_size_inches()[0] * self.matplotlib_widget.figure.dpi),
            int(self.matplotlib_widget.figure.get_size_inches()[1] * self.matplotlib_widget.figure.dpi)
        )


if __name__ == "__main__":
    # Uygulamanın düzgün bir şekilde başlatıldığından emin olmak için try-except bloğu
    try:
        app = QApplication(sys.argv)
        window = MainWindow()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        # Uygulama başlatılamazsa hata mesajı göster
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText("Uygulama başlatılırken kritik bir hata oluştu.")
        msg.setInformativeText(f"Lütfen gerekli kütüphanelerin (PyQt5, Matplotlib, NumPy) kurulu olduğundan ve ikon dosyalarının 'icons/' klasöründe bulunduğundan emin olun.\nHata: {e}")
        msg.setWindowTitle("Kritik Hata")
        msg.exec_()
        sys.exit(1)