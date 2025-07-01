import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QAction, QMenu, QInputDialog, QMessageBox, \
    QFileDialog  # QFileDialog eklendi
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

        # 2. Grafik Oluştur Menüsü (Aynı kaldı)
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

        # ---
        ## 3. İndir / Yazdır Menüsü (YENİ!)
        # ---
        download_print_menu = menubar.addMenu("&İndir / Yazdır")

        # Kaydetme Alt Menüsü
        save_as_menu = QMenu("Farklı Kaydet", self)

        # Farklı Kaydet seçenekleri
        formats = {"PNG G&örseli": "png", "JPEG G&örseli": "jpeg", "PDF &Belgesi": "pdf", "SVG &Vektörü": "svg"}
        for name, file_ext in formats.items():
            save_action = QAction(QIcon(f'icons/save_{file_ext}.png'), name, self)  # İkonları varsayıyoruz
            save_action.setStatusTip(f"Grafiği .{file_ext} formatında kaydet")
            # Her format için farklı kaydet fonksiyonuna bağlanıyoruz
            save_action.triggered.connect(lambda checked, fmt=file_ext: self.save_graph(fmt))
            save_as_menu.addAction(save_action)

        download_print_menu.addMenu(save_as_menu)

        download_print_menu.addSeparator()  # Ayırıcı

        # Yazdırma Aksiyonu
        print_action = QAction(QIcon('icons/print.png'), "&Yazdır...", self)  # Yazdır ikonu varsayıyoruz
        print_action.setShortcut("Ctrl+P")  # Çıktı menüsü için yeni P kısayolu
        print_action.setStatusTip("Mevcut grafiği yazdır")
        print_action.triggered.connect(self.print_graph)
        download_print_menu.addAction(print_action)

        # 4. Veri Seç Menüsü (Sıra değişti, aynı kaldı)
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
                                      min=1, max=100, step=1)

        if ok:
            QMessageBox.information(self, "Seçim Onayı",
                                    f"'{chart_type}' türünde {num} adet grafik oluşturulacak.")
            # Burada 'chart_type' ve 'num' değerlerini kullanarak grafik çizme fonksiyonunuzu çağırabilirsiniz.
            # self.draw_graph(chart_type, num)
        else:
            QMessageBox.information(self, "İptal Edildi", "Grafik oluşturma işlemi iptal edildi.")

    # Yeni metot: Grafiği farklı formatta kaydet
    def save_graph(self, file_format):
        # QFileDialog.getSaveFileName() kullanıcıya dosya kaydetme penceresi açar.
        # İlk argüman parent widget'tır (self).
        # İkinci argüman pencere başlığıdır.
        # Üçüncü argüman varsayılan dosya adıdır (isteğe bağlı).
        # Dördüncü argüman dosya filtreleridir.
        file_name, _ = QFileDialog.getSaveFileName(self, "Grafiği Kaydet", "",
                                                   f"Grafik Dosyaları (*.{file_format});;Tüm Dosyalar (*)")

        if file_name:  # Kullanıcı bir dosya adı seçip kaydet'e bastıysa
            QMessageBox.information(self, "Kaydetme İşlemi",
                                    f"Grafik '{file_name}' olarak kaydedilecek. (Format: .{file_format})")
            # Burada Matplotlib figürünüzü veya QGraphicsScene içeriğini file_name yoluna kaydetmelisiniz.
            # Örnek: my_figure.savefig(file_name)
            # Eğer bir QGraphicsView kullanıyorsanız:
            # from PyQt5.QtGui import QPixmap
            # pixmap = self.view.grab() # View'daki içeriği QPixmap olarak al
            # pixmap.save(file_name) # QPixmap'i kaydet
        else:
            QMessageBox.information(self, "Kaydetme İptal Edildi", "Kaydetme işlemi iptal edildi.")

    # Yeni metot: Grafiği yazdır
    def print_graph(self):
        QMessageBox.information(self, "Yazdırma İşlemi", "Grafik yazdırma işlemi başlatılacak.")
        # Burada grafik yazdırma diyaloğunu açacak ve grafik içeriğini yazıcıya gönderecek kodunuzu eklemelisiniz.
        # Bu, QtPrintSupport modülü ile yapılır. Örnek:
        # from PyQt5.QtPrintSupport import QPrintDialog, QPrinter
        # printer = QPrinter()
        # dialog = QPrintDialog(printer, self)
        # if dialog.exec_() == QPrintDialog.Accepted:
        #     # self.graphics_scene.render(printer) # QGraphicsScene içeriğini yazdır
        #     pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())