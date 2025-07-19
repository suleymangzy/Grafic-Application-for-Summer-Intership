import logging
from pathlib import Path
import pandas as pd
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QFileDialog, QMessageBox
)
from config.constants import REQ_SHEETS


class FileSelectionPage(QWidget):
    """
    Excel dosyası seçimi için kullanılan sayfa bileşeni.
    Kullanıcıdan .xlsx dosyası seçmesini ister,
    seçilen dosyadaki uygun sayfaları kontrol eder ve
    uygun sayfa varsa diğer grafik sayfalarına geçiş düğmelerini etkinleştirir.
    """

    def __init__(self, main_window: "MainWindow") -> None:
        """Ana pencere referansını alır ve kullanıcı arayüzünü başlatır."""
        super().__init__()
        self.main_window = main_window
        self.init_ui()

    def init_ui(self):
        """Sayfa arayüz bileşenlerini oluşturur ve yerleştirir."""
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)  # İçeriği dikeyde ortala

        # Başlık etiketi
        title_label = QLabel("<h2>Dosya Seçimi</h2>")
        title_label.setObjectName("title_label")  # Stil amaçlı
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        # Seçilen dosya yolunu gösteren etiket
        self.lbl_path = QLabel("Henüz dosya seçilmedi")
        self.lbl_path.setAlignment(Qt.AlignCenter)
        self.lbl_path.setStyleSheet(
            "font-size: 11pt; color: #555555; margin-bottom: 15px;"
        )  # Görsel iyileştirme
        layout.addWidget(self.lbl_path)

        # Dosya seçme butonu
        self.btn_browse = QPushButton(".xlsx dosyası seç…")
        self.btn_browse.clicked.connect(self.browse)
        layout.addWidget(self.btn_browse)

        layout.addStretch(1)  # Aşağıya boşluk bırak

        # Grafik türü düğmelerini yatay düzenle
        h_layout_buttons = QHBoxLayout()
        h_layout_buttons.addStretch(1)  # Düğmeleri ortaya hizala

        # Günlük grafikler düğmesi
        self.btn_daily_graphs = QPushButton("Günlük Grafikler")
        self.btn_daily_graphs.clicked.connect(self.go_to_daily_graphs)
        self.btn_daily_graphs.setEnabled(False)  # Başlangıçta devre dışı
        h_layout_buttons.addWidget(self.btn_daily_graphs)

        # Aylık grafikler düğmesi
        self.btn_monthly_graphs = QPushButton("Aylık Grafikler")
        self.btn_monthly_graphs.clicked.connect(self.go_to_monthly_graphs)
        self.btn_monthly_graphs.setEnabled(False)  # Başlangıçta devre dışı
        h_layout_buttons.addWidget(self.btn_monthly_graphs)

        h_layout_buttons.addStretch(1)  # Düğmeleri ortaya hizala
        layout.addLayout(h_layout_buttons)

        layout.addStretch(1)  # Yukarıya boşluk bırak

    def browse(self) -> None:
        """
        Dosya seçim penceresini açar, kullanıcının Excel dosyası seçmesini sağlar.
        Seçilen dosyanın gerekli sayfaları içerip içermediğini kontrol eder,
        uygun değilse kullanıcıyı uyarır ve sayfayı sıfırlar.
        """
        path, _ = QFileDialog.getOpenFileName(
            self, "Excel seç", str(Path.home()), "Excel Files (*.xlsx)"
        )
        if not path:
            return  # Dosya seçilmediyse fonksiyonu bitir

        try:
            xls = pd.ExcelFile(path)
            # Dosyadaki sayfalarla gereken sayfaların kesişimi
            sheets = sorted(list(REQ_SHEETS.intersection(set(xls.sheet_names))))

            if not sheets:
                # Gerekli sayfalar yoksa uyarı göster
                QMessageBox.warning(
                    self,
                    "Uygun sayfa yok",
                    f"Seçilen dosyada istenen ({', '.join(REQ_SHEETS)}) sheet bulunamadı.",
                )
                self.reset_page()  # Sayfayı varsayılana döndür
                return

            # Dosya yolu ve uygun sayfaları ana pencereye bildir
            self.main_window.excel_path = Path(path)
            self.lbl_path.setText(f"Seçilen Dosya: <b>{Path(path).name}</b>")

            self.main_window.available_sheets = sheets
            # Varsayılan sayfa olarak "SMD-OEE" varsa onu seç, yoksa ilkini seç
            if "SMD-OEE" in sheets:
                self.main_window.selected_sheet = "SMD-OEE"
            elif sheets:
                self.main_window.selected_sheet = sheets[0]
            else:
                self.main_window.selected_sheet = None  # Uygun sayfa yok

            # Grafik sayfalarına geçiş düğmelerini etkinleştir
            self.btn_daily_graphs.setEnabled(True)
            self.btn_monthly_graphs.setEnabled(True)

            logging.info("Dosya seçildi: %s", path)

        except Exception as e:
            # Dosya okuma veya işleme hatası durumunda uyarı göster ve sayfayı sıfırla
            QMessageBox.critical(
                self,
                "Okuma hatası",
                f"Dosya okunurken bir hata oluştu: {e}\n"
                "Lütfen dosyanın bozuk olmadığından ve Excel formatında olduğundan emin olun.",
            )
            self.reset_page()

    def go_to_daily_graphs(self) -> None:
        """Günlük grafikler sayfasına geçiş yapar."""
        self.main_window.goto_page(1)

    def go_to_monthly_graphs(self) -> None:
        """
        Aylık grafikler sayfasına geçiş yapar.
        Öncelikle uygun sayfa seçilir ve yüklenir.
        """
        # Aylık grafikler için "SMD-OEE" sayfası tercih edilir
        if "SMD-OEE" in self.main_window.available_sheets:
            self.main_window.selected_sheet = "SMD-OEE"
        elif self.main_window.available_sheets:
            self.main_window.selected_sheet = self.main_window.available_sheets[0]
        else:
            QMessageBox.warning(self, "Uyarı", "Aylık grafikler için uygun sayfa bulunamadı.")
            return

        self.main_window.load_excel()  # Seçilen sayfayı yükle
        self.main_window.goto_page(3)

    def reset_page(self):
        """Sayfayı ilk haline döndürür: dosya seçimini iptal eder ve butonları pasif yapar."""
        self.main_window.excel_path = None
        self.main_window.selected_sheet = None
        self.main_window.available_sheets = []
        self.lbl_path.setText("Henüz dosya seçilmedi")
        self.btn_daily_graphs.setEnabled(False)
        self.btn_monthly_graphs.setEnabled(False)
