import sys
import os  # Dosya yolu işlemleri için
import random  # Örnek veri üretimi için
import numpy as np  # Sayısal işlemler için

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QAction, QMenu, QInputDialog,
    QMessageBox, QFileDialog, QLabel, QVBoxLayout, QWidget
)
from PyQt5.QtGui import QIcon, QFont, QPixmap
from PyQt5.QtCore import Qt, QSize

# Matplotlib entegrasyonu için gerekli importlar
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar


class MplWidget(QWidget):
    """
    Matplotlib grafiğini içinde barındıracak özel bir QWidget.
    """

    def __init__(self, parent=None):
        super().__init__(parent)

        self.figure = plt.figure()  # Matplotlib figürü oluştur
        self.canvas = FigureCanvas(self.figure)  # Figürü bir Qt widget'ına dönüştür
        self.toolbar = NavigationToolbar(self.canvas, self)  # Navigasyon araç çubuğu

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
        self.file_info_label.setFixedHeight(30)  # Sabit yükseklik verelim

        # QLabel'i ana layout'a ekle
        self.main_layout.addWidget(self.file_info_label)

        # Matplotlib widget'ını oluştur
        self.matplotlib_widget = MplWidget(self)
        self.main_layout.addWidget(self.matplotlib_widget)  # Matplotlib widget'ını layout'a ekle

        self.create_menu()

        self.setStyleSheet("""
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
                color: #ADD8E6;
                font-weight: bold;
                padding: 5px;
                background-color: #282828;
                border-bottom: 1px solid #444;
            }
        """)

    # resizeEvent'i artık main_layout yönettiği için kaldırdık, QLabel'in yüksekliğini sabitledik.
    # Ancak yine de QLabel'in genişliğini ayarlamak için gerekebilir.
    # def resizeEvent(self, event):
    #    self.file_info_label.setGeometry(0, self.menuBar().height(), self.width(), 30)
    #    super().resizeEvent(event)

    def create_menu(self):
        menubar = self.menuBar()

        # 1. Dosya Menüsü
        file_menu = menubar.addMenu("&Dosya")

        word_action = QAction(QIcon('icons/word.png'), "Word Dosyası &Aç...", self)
        word_action.setShortcut("Ctrl+W")
        word_action.setStatusTip("Bir Word belgesini açar")
        word_action.triggered.connect(lambda: self.open_file("Word Dosyaları (*.docx *.doc);;Tüm Dosyalar (*)", "word"))
        file_menu.addAction(word_action)

        excel_action = QAction(QIcon('icons/excel.png'), "Excel Dosyası &Aç...", self)
        excel_action.setShortcut("Ctrl+E")
        excel_action.setStatusTip("Bir Excel çalışma sayfasını açar")
        excel_action.triggered.connect(
            lambda: self.open_file("Excel Dosyaları (*.xlsx *.xls);;Tüm Dosyalar (*)", "excel"))
        file_menu.addAction(excel_action)

        pptx_action = QAction(QIcon('icons/pptx.png'), "PPTX Dosyası &Aç...", self)
        pptx_action.setShortcut("Ctrl+P")
        pptx_action.setStatusTip("Bir PowerPoint sunumunu açar")
        pptx_action.triggered.connect(
            lambda: self.open_file("PowerPoint Dosyaları (*.pptx *.ppt);;Tüm Dosyalar (*)", "pptx"))
        file_menu.addAction(pptx_action)

        file_menu.addSeparator()

        exit_action = QAction(QIcon('icons/exit.png'), "Çı&kış", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.setStatusTip("Uygulamadan çıkar")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # 2. Grafik Oluştur Menüsü (Değişiklik burada!)
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
            # draw_graph metodunu bağla
            action.triggered.connect(lambda checked, name=grafik_adı: self.get_plot_count(name))
            plot_menu.addAction(action)

        # 3. İndir / Yazdır Menüsü (Aynı kaldı)
        download_print_menu = menubar.addMenu("&İndir / Yazdır")

        save_as_menu = QMenu("Farklı Kaydet", self)

        formats = {"PNG G&örseli": "png", "JPEG G&örseli": "jpeg", "PDF &Belgesi": "pdf", "SVG &Vektörü": "svg"}
        for name, file_ext in formats.items():
            save_action = QAction(QIcon(f'icons/save_{file_ext}.png'), name, self)
            save_action.setStatusTip(f"Grafiği .{file_ext} formatında kaydet")
            save_action.triggered.connect(lambda checked, fmt=file_ext: self.save_graph(fmt))
            save_as_menu.addAction(save_action)

        download_print_menu.addMenu(save_as_menu)

        download_print_menu.addSeparator()

        print_action = QAction(QIcon('icons/print.png'), "&Yazdır...", self)
        print_action.setShortcut("Ctrl+P")
        print_action.setStatusTip("Mevcut grafiği yazdır")
        print_action.triggered.connect(self.print_graph)
        download_print_menu.addAction(print_action)

        # 4. Veri Seç Menüsü (Aynı kaldı)
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
        num, ok = QInputDialog.getInt(self, "Grafik Adedi Girin",
                                      f"Kaç adet '{chart_type}' grafiği oluşturmak istersiniz?",
                                      min=1, max=5, step=1)  # Maksimum 5 grafik ile sınırlandırdık

        if ok:
            # Burası önemli: draw_graph metodunu çağırıyoruz
            self.draw_graph(chart_type, num)
            QMessageBox.information(self, "Grafik Oluşturuldu",
                                    f"'{chart_type}' türünde {num} adet grafik başarıyla oluşturuldu.")
        else:
            QMessageBox.information(self, "İptal Edildi", "Grafik oluşturma işlemi iptal edildi.")

    def save_graph(self, file_format):
        # Eğer henüz bir grafik çizilmemişse, kaydetme işlemini engelle
        if not self.matplotlib_widget.figure.axes:
            QMessageBox.warning(self, "Uyarı", "Kaydedilecek bir grafik bulunamadı. Lütfen önce bir grafik oluşturun.")
            return

        file_name, _ = QFileDialog.getSaveFileName(self, "Grafiği Kaydet", "grafik",  # Varsayılan dosya adı
                                                   f"Grafik Dosyaları (*.{file_format});;Tüm Dosyalar (*)")

        if file_name:
            try:
                # Matplotlib figürünü belirtilen dosya formatında kaydet
                self.matplotlib_widget.figure.savefig(file_name, format=file_format)
                QMessageBox.information(self, "Kaydetme Başarılı",
                                        f"Grafik '{file_name}' olarak kaydedildi.")
            except Exception as e:
                QMessageBox.critical(self, "Hata", f"Grafik kaydedilirken bir hata oluştu: {e}")
        else:
            QMessageBox.information(self, "Kaydetme İptal Edildi", "Kaydetme işlemi iptal edildi.")

    def print_graph(self):
        # Yazdırma fonksiyonu için QtPrintSupport gerekli.
        # Bu örneği daha sonra detaylandırabiliriz.
        QMessageBox.information(self, "Yazdırma İşlemi",
                                "Grafik yazdırma işlemi başlatılacak. (Henüz tam işlevsel değil)")

    def open_file(self, file_filter, file_type_code):
        file_name, _ = QFileDialog.getOpenFileName(self, "Dosya Seç", "", file_filter)

        if file_name:
            self.update_file_info_label(file_name, file_type_code)
            QMessageBox.information(self, "Dosya Seçimi",
                                    f"Seçilen Dosya: {file_name}")
            # Burada seçilen dosyadan veri okuma ve işleme mantığını ekleyebilirsiniz.
            # Örneğin: self.load_data_from_file(file_name, file_type_code)
        else:
            self.update_file_info_label("", "")
            QMessageBox.information(self, "İptal Edildi", "Dosya seçimi iptal edildi.")

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
                self.file_info_label.setText(f"  {base_name} (İkon bulunamadı)")
        else:
            self.file_info_label.setPixmap(QPixmap())
            self.file_info_label.setText("Lütfen bir dosya seçin...")

        # QLabel'in boyutunu içeriğine göre ayarla
        self.file_info_label.adjustSize()
        # ve genişliğini pencere genişliğine sabitle
        self.file_info_label.setFixedWidth(self.width())

    ## YENİ: Grafik Çizme Metodu
    def draw_graph(self, chart_type_text, count):
        # Matplotlib figürünü ve alt grafiklerini temizle
        self.matplotlib_widget.figure.clear()

        # Grafik adetine göre alt grafikler (subplots) oluştur
        # Basitlik için en fazla 2x2 veya 3x2 düzen kullanabiliriz.
        rows = 1
        cols = 1
        if count == 2:
            rows, cols = 1, 2
        elif count == 3:
            rows, cols = 1, 3
        elif count == 4:
            rows, cols = 2, 2
        elif count > 4:  # Maksimum 5 grafik için
            rows, cols = 2, 3  # 2 satır 3 sütun = 6 adet yer var, 5'ini kullanırız

        # Matplotlib'in parantez içindeki fonksiyon adını al (örneğin 'plot' için 'çizgi grafiği (plot)')
        chart_func_name = chart_type_text.split('(')[-1][:-1] if '(' in chart_type_text else None

        if chart_func_name is None:
            QMessageBox.warning(self, "Hata", "Geçersiz grafik türü adı.")
            return

        for i in range(count):
            ax = self.matplotlib_widget.figure.add_subplot(rows, cols, i + 1)
            ax.set_title(f"{chart_type_text.split('(')[0].strip()} {i + 1}")

            # Örnek veri üretimi
            x = np.linspace(0, 10, 100)
            y = np.sin(x + i * 0.5) + random.uniform(-0.5, 0.5)  # Rastgelelik ekle

            if chart_func_name == "plot":
                ax.plot(x, y, label=f'Series {i + 1}')
                ax.set_xlabel("X Ekseni")
                ax.set_ylabel("Y Ekseni")
                ax.legend()
            elif chart_func_name == "bar":
                categories = [f'Kategori {j + 1}' for j in range(5)]
                values = np.random.randint(5, 20, 5)
                # Düzeltilen kısım: enumerate kullanarak renkleri kategori sayısına göre belirliyoruz
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
                labels = [f'Dilim {j + 1}' for j in range(4)]
                ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90)
                ax.axis('equal')  # Pastanın daire şeklinde olmasını sağla
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
                ax.boxplot(data, patch_artist=True)  # Patch_artist kutuları renklendirmek için
                ax.set_xticklabels([f'Grup {j + 1}' for j in range(len(data))])
                ax.set_ylabel("Değer")
            elif chart_func_name == "violinplot":
                data = [np.random.normal(0, std, 100) for std in range(1, 4)]
                ax.violinplot(data, showmeans=True, showmedians=True)
                ax.set_xticks(np.arange(1, len(data) + 1))
                ax.set_xticklabels([f'Grup {j + 1}' for j in range(len(data))])
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

        # Tüm alt grafikleri sıkıştır
        self.matplotlib_widget.figure.tight_layout()
        # Figürü yenile
        self.matplotlib_widget.canvas.draw_idle()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())