import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QAction, QMenu, QInputDialog, \
    QMessageBox  # QInputDialog ve QMessageBox eklendi
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import Qt


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Grafik Uygulaması")
        self.setGeometry(100, 100, 1024, 768)
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
        """)

    def create_menu(self):
        menubar = self.menuBar()

        # 1. Dosya Menüsü (Aynı kaldı)
        file_menu = menubar.addMenu("&Dosya")

        word_action = QAction(QIcon('icons/word.png'), "Word Dosyası &Aç    ", self)
        word_action.setShortcut("Ctrl+W")
        word_action.setStatusTip("Bir Word belgesini açar")
        file_menu.addAction(word_action)

        excel_action = QAction(QIcon('icons/excel.png'), "Excel Dosyası &Aç    ", self)
        excel_action.setShortcut("Ctrl+E")
        excel_action.setStatusTip("Bir Excel çalışma sayfasını açar")
        file_menu.addAction(excel_action)

        pptx_action = QAction(QIcon('icons/pptx.png'), "PPTX Dosyası &Aç    ", self)
        pptx_action.setShortcut("Ctrl+P")
        pptx_action.setStatusTip("Bir PowerPoint sunumunu açar")
        file_menu.addAction(pptx_action)

        file_menu.addSeparator()

        exit_action = QAction(QIcon('icons/exit.png'), "Çı&kış", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.setStatusTip("Uygulamadan çıkar")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # 2. Grafik Oluştur Menüsü (Değişiklik burada!)
        plot_menu = menubar.addMenu("&Grafik Oluştur")

        self.chart_types = [  # chart_types'ı sınıf özelliği olarak tanımladık
            "Çizgi Grafiği (plot)",
            "Bar Grafiği (bar)",
            "Histogram (hist)",
            "Pasta Grafiği (pie)",
            "Dağılım Grafiği (scatter)",
            "Alan Grafiği (fill_between)",
            "Kutu Grafiği (boxplot)",
            "Violin Grafiği (violinplot)",
            "Stem Grafiği (stem)",
            "Hata Çubuklu Grafik (errorbar)"
        ]

        for grafik_adı in self.chart_types:
            action = QAction(grafik_adı + "...", self)  # "..." ekledik, çünkü bir diyalog açılacak
            action.setStatusTip(f"'{grafik_adı}' grafiği için adet seçimi")
            # Her grafik türü aksiyonuna, sayı girişi isteyen bir metot bağladık
            action.triggered.connect(lambda checked, name=grafik_adı: self.get_plot_count(name))
            plot_menu.addAction(action)

        # 3. Veri Seç Menüsü (Aynı kaldı)
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

    # Yeni metot: Kullanıcıdan grafik adedi al
    def get_plot_count(self, chart_type):
        num, ok = QInputDialog.getInt(self, "Grafik Adedi Girin",
                                      f"Kaç adet '{chart_type}' grafiği oluşturmak istersiniz?",
                                      min=1, max=100, step=1)  # Min, max ve step değerleri belirledik

        if ok:  # Kullanıcı OK tuşuna bastıysa
            QMessageBox.information(self, "Seçim Onayı",
                                    f"'{chart_type}' türünde {num} adet grafik oluşturulacak.")
            # Burada 'chart_type' ve 'num' değerlerini kullanarak grafik çizme fonksiyonunuzu çağırabilirsiniz.
            # Örneğin: self.draw_graph(chart_type, num)
        else:  # Kullanıcı İptal tuşuna bastıysa
            QMessageBox.information(self, "İptal Edildi", "Grafik oluşturma işlemi iptal edildi.")

    # Grafik çizme işini yapacak varsayımsal bir metot (gerçek uygulamanızda doldurulacak)
    # def draw_graph(self, chart_type, count):
    #     print(f"'{chart_type}' türünde {count} adet grafik çiziliyor...")
    #     # Burada Matplotlib veya başka bir kütüphane ile grafik çizim kodunuzu yazın.


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())