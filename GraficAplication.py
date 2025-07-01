import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QAction, QMenu

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyQt5 Menü Çubuğu Örneği")
        self.setGeometry(100, 100, 800, 600)
        self.create_menu()

    def create_menu(self):
        menubar = self.menuBar()

        # 1. Dosya Seç Menüsü
        dosya_menu = menubar.addMenu("Dosya Seç")
        word_action = QAction("Word", self)
        excel_action = QAction("Excel", self)
        pptx_action = QAction("PPTX", self)
        dosya_menu.addAction(word_action)
        dosya_menu.addAction(excel_action)
        dosya_menu.addAction(pptx_action)

        # 2. Grafik Oluştur Menüsü
        grafik_menu = menubar.addMenu("Grafik Oluştur")
        # Matplotlib ile oluşturulabilecek temel grafikler
        grafikler = [
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
        for grafik in grafikler:
            action = QAction(grafik, self)
            grafik_menu.addAction(action)

        # 3. Veri Seç Menüsü
        veri_menu = menubar.addMenu("Veri Seç")
        # Örnek olarak x ve y ekseni seçimleri
        x_ekseni_menu = QMenu("X Ekseni Seç", self)
        y_ekseni_menu = QMenu("Y Ekseni Seç", self)
        # Burada örnek olarak bazı seçenekler ekleniyor, ihtiyaca göre dinamik yapılabilir
        x1_action = QAction("X1", self)
        x2_action = QAction("X2", self)
        y1_action = QAction("Y1", self)
        y2_action = QAction("Y2", self)
        x_ekseni_menu.addAction(x1_action)
        x_ekseni_menu.addAction(x2_action)
        y_ekseni_menu.addAction(y1_action)
        y_ekseni_menu.addAction(y2_action)
        veri_menu.addMenu(x_ekseni_menu)
        veri_menu.addMenu(y_ekseni_menu)

        # Farklı grafik türleri için farklı gereklilikler burada eklenebilir
        # Örneğin: Bar grafiği için kategori seçimi, scatter için renk seçimi gibi

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
