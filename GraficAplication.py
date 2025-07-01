import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QAction, QMenu, QInputDialog, QMessageBox, QFileDialog, QLabel
from PyQt5.QtGui import QIcon, QFont, QPixmap
from PyQt5.QtCore import Qt, QSize  # QSize'ı ekledik


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Grafik Uygulaması")
        self.setGeometry(100, 100, 1024, 768)

        # QLabel'i menü çubuğunun altında gösterecek bir merkezi widget veya layout kullanabiliriz.
        # Basitçe, şimdilik direkt merkezi widget'a yerleştirelim.
        # Gerçek bir uygulamada, daha karmaşık bir layout (örneğin QVBoxLayout) kullanmanız gerekebilir.
        self.file_info_label = QLabel("Henüz bir dosya seçilmedi.", self)
        self.file_info_label.setAlignment(Qt.AlignCenter)  # Ortala hizala
        self.file_info_label.setFont(QFont("Arial", 12))  # Font ayarı
        # Başlangıçta ikonu boş bırakabiliriz veya varsayılan bir ikon atayabiliriz
        self.file_info_label.setPixmap(QPixmap())
        self.file_info_label.setText("Lütfen bir dosya seçin...")
        self.file_info_label.setGeometry(0, 50, self.width(), 30)  # Geçici konumlandırma

        # Bu QLabel'ı pencerenin merkezi widget'ı yapıyoruz.
        # Ancak merkezi widget genellikle grafiklerin çizileceği alandır.
        # Daha iyi bir yaklaşım, bu QLabel'ı bir layout içine alıp menü çubuğunun altına eklemektir.
        # Basitlik adına, şimdilik bu şekilde tutalım, ancak daha sonra bir düzen yöneticisi eklememiz gerekecek.
        # Örneğin, merkezi widget bir QWidget olabilir ve içine QLabel ile QGraphicsView'ı bir QBoxLayout ile yerleştirebiliriz.
        # Şimdilik, sadece görselleştirme için QLabel'in konumunu manuel ayarlayalım.

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
            QLabel#fileInfoLabel { /* QLabel'e özel stil için objectName kullanabiliriz */
                color: #ADD8E6; /* Açık mavi */
                font-weight: bold;
                padding: 5px;
                background-color: #282828; /* Koyu gri arka plan */
                border-bottom: 1px solid #444;
            }
        """)
        # QLabel'e objectName atayarak QSS ile stil uygulayabiliriz
        self.file_info_label.setObjectName("fileInfoLabel")

    # Pencere boyutu değiştiğinde label'ın konumunu güncellemek için
    def resizeEvent(self, event):
        self.file_info_label.setGeometry(0, self.menuBar().height(), self.width(), 30)
        super().resizeEvent(event)

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
        num, ok = QInputDialog.getInt(self, "Grafik Adedi Girin",
                                      f"Kaç adet '{chart_type}' grafiği oluşturmak istersiniz?",
                                      min=1, max=100, step=1)

        if ok:
            QMessageBox.information(self, "Seçim Onayı",
                                    f"'{chart_type}' türünde {num} adet grafik oluşturulacak.")
            # self.draw_graph(chart_type, num)
        else:
            QMessageBox.information(self, "İptal Edildi", "Grafik oluşturma işlemi iptal edildi.")

    def save_graph(self, file_format):
        file_name, _ = QFileDialog.getSaveFileName(self, "Grafiği Kaydet", "",
                                                   f"Grafik Dosyaları (*.{file_format});;Tüm Dosyalar (*)")

        if file_name:
            QMessageBox.information(self, "Kaydetme İşlemi",
                                    f"Grafik '{file_name}' olarak kaydedilecek. (Format: .{file_format})")
        else:
            QMessageBox.information(self, "Kaydetme İptal Edildi", "Kaydetme işlemi iptal edildi.")

    def print_graph(self):
        QMessageBox.information(self, "Yazdırma İşlemi", "Grafik yazdırma işlemi başlatılacak.")

    # open_file metodu güncellendi: dosya_tipi_kodu argümanı eklendi
    def open_file(self, file_filter, file_type_code):
        file_name, _ = QFileDialog.getOpenFileName(self, "Dosya Seç", "", file_filter)

        if file_name:
            # QLabel'i güncelle
            self.update_file_info_label(file_name, file_type_code)

            QMessageBox.information(self, "Dosya Seçimi",
                                    f"Seçilen Dosya: {file_name}")
            # Burada dosya okuma ve işleme kodunuzu yazın.
            # Örneğin: self.load_data_from_file(file_name, file_type_code)
        else:
            self.update_file_info_label("", "")  # Dosya seçimi iptal edilirse etiketi sıfırla
            QMessageBox.information(self, "İptal Edildi", "Dosya seçimi iptal edildi.")

    # Yeni metot: Dosya bilgi etiketini güncelle
    def update_file_info_label(self, file_path, file_type_code):
        if file_path:
            # Dosya adını ve uzantısını al
            import os
            base_name = os.path.basename(file_path)

            # İkona path'i oluştur
            icon_path = f'icons/{file_type_code}.png'  # icons/word.png, icons/excel.png gibi

            # İkonu yükle ve boyutu ayarla
            pixmap = QPixmap(icon_path)
            if not pixmap.isNull():
                pixmap = pixmap.scaled(24, 24, Qt.KeepAspectRatio, Qt.SmoothTransformation)  # İkon boyutu
                self.file_info_label.setPixmap(pixmap)
                self.file_info_label.setText(f"  {base_name}")  # İkonun yanına dosya adını yaz
                self.file_info_label.adjustSize()  # Metin ve ikon boyutuna göre ayarla
                self.file_info_label.setContentsMargins(5, 0, 0, 0)  # Sola biraz boşluk
            else:
                # İkon bulunamazsa sadece metin göster
                self.file_info_label.setPixmap(QPixmap())  # İkonu sıfırla
                self.file_info_label.setText(f"  {base_name} (İkon bulunamadı)")
        else:
            self.file_info_label.setPixmap(QPixmap())
            self.file_info_label.setText("Lütfen bir dosya seçin...")

        # QLabel'i yeniden konumlandır (resizeEvent tetiklenmezse diye)
        self.file_info_label.setGeometry(0, self.menuBar().height(), self.width(), self.file_info_label.height())


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())