import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QAction, QMenu, QGraphicsView, QGraphicsScene
from PyQt5.QtGui import QIcon, QFont  # QFont'u ekledik
from PyQt5.QtCore import Qt


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Grafik Uygulaması")
        self.setGeometry(100, 100, 1024, 768)  # Daha geniş bir başlangıç boyutu verdik
        self.create_menu()

        # Uygulamanın genel stilini iyileştirmek için isteğe bağlı stil
        self.setStyleSheet("""
            QMenuBar {
                background-color: #333; /* Koyu arka plan */
                color: #FFF; /* Beyaz yazı rengi */
                font-size: 14px; /* Menü çubuğu yazıları biraz daha büyük */
                padding: 5px 0px; /* Üst ve altta boşluk */
            }
            QMenuBar::item {
                padding: 8px 15px; /* Menü öğeleri arasında daha fazla boşluk */
                background-color: transparent;
            }
            QMenuBar::item:selected {
                background-color: #555; /* Seçili öğe arka plan rengi */
            }
            QMenu {
                background-color: #444; /* Açılır menü arka plan rengi */
                color: #FFF;
                border: 1px solid #666; /* Kenarlık */
            }
            QMenu::item {
                padding: 8px 25px; /* Alt menü öğeleri için daha fazla dolgu */
                background-color: transparent;
            }
            QMenu::item:selected {
                background-color: #666; /* Açılır menüde seçili öğe rengi */
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
        # macOS'te doğal menü çubuğunu kullanmak için:
        # menubar.setNativeMenuBar(True)

        # 1. Dosya Menüsü
        file_menu = menubar.addMenu("&Dosya")

        # Varsayımsal ikonlar: Gerçek ikon dosyalarını ('icons/word.png' gibi) oluşturmanız veya path'lerini ayarlamanız gerekir.
        # İkonlarınız yoksa QIcon() kısmını silebilirsiniz.

        word_action = QAction(QIcon('icons/word.png'), "Word Dosyası &Aç...", self)
        word_action.setShortcut("Ctrl+W")
        word_action.setStatusTip("Bir Word belgesini açar")
        file_menu.addAction(word_action)

        excel_action = QAction(QIcon('icons/excel.png'), "Excel Dosyası &Aç...", self)
        excel_action.setShortcut("Ctrl+E")
        excel_action.setStatusTip("Bir Excel çalışma sayfasını açar")
        file_menu.addAction(excel_action)

        pptx_action = QAction(QIcon('icons/pptx.png'), "PPTX Dosyası &Aç...", self)
        pptx_action.setShortcut("Ctrl+P")
        pptx_action.setStatusTip("Bir PowerPoint sunumunu açar")
        file_menu.addAction(pptx_action)

        file_menu.addSeparator()

        exit_action = QAction(QIcon('icons/exit.png'), "Çı&kış", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.setStatusTip("Uygulamadan çıkar")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # 2. Grafik Oluştur Menüsü
        plot_menu = menubar.addMenu("&Grafik Oluştur")

        chart_types = [
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

        for grafik in chart_types:
            action = QAction(grafik, self)
            action.setStatusTip(f"{grafik} oluşturur")
            # action.triggered.connect(lambda checked, g=grafik: self.create_plot(g))
            plot_menu.addAction(action)

        # 3. Veri Seç Menüsü
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


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())