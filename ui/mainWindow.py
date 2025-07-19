import logging
from pathlib import Path
from typing import List
import pandas as pd
from PyQt5.QtWidgets import QMainWindow, QStackedWidget, QMessageBox

from ui.fileSelectionPage import FileSelectionPage
from ui.dataSelectionPage import DataSelectionPage
from ui.dailyGraphPage import DailyGraphsPage
from ui.monthlyGraphPage import MonthlyGraphsPage
from utils.helpers import excel_col_to_index

class MainWindow(QMainWindow):
    """Ana uygulama penceresini temsil eder. Sayfalar arası geçişi yönetir ve global verileri tutar."""

    def __init__(self) -> None:
        super().__init__()
        # Uygulama genelinde kullanılacak veri ve durum değişkenleri
        self.excel_path: Path | None = None
        self.selected_sheet: str | None = None
        self.available_sheets: List[str] = []
        self.df: pd.DataFrame = pd.DataFrame()
        self.grouping_col_name: str | None = None
        self.grouped_col_name: str | None = None
        self.oee_col_name: str | None = None
        self.metric_cols: List[str] = []
        self.grouped_values: List[str] = []
        self.selected_metrics: List[str] = []
        self.selected_grouping_val: str = ""

        # Sayfaları yönetmek için QStackedWidget kullanımı
        self.stacked_widget = QStackedWidget()
        self.file_selection_page = FileSelectionPage(self)
        self.data_selection_page = DataSelectionPage(self)
        self.daily_graphs_page = DailyGraphsPage(self)
        self.monthly_graphs_page = MonthlyGraphsPage(self)

        # Sayfaları stacked widget'a ekleme
        self.stacked_widget.addWidget(self.file_selection_page)
        self.stacked_widget.addWidget(self.data_selection_page)
        self.stacked_widget.addWidget(self.daily_graphs_page)
        self.stacked_widget.addWidget(self.monthly_graphs_page)

        # Ana pencerenin merkezi widget'ı olarak stacked widget'ı ayarla
        self.setCentralWidget(self.stacked_widget)
        self.setWindowTitle("OEE ve Durum Grafiği Uygulaması")
        self.setGeometry(100, 100, 1200, 800) # Pencere boyutunu ayarla

        # Uygulama genelinde stil sayfasını uygula
        self.apply_stylesheet()
        # Uygulama başlangıcında ilk sayfaya git
        self.goto_page(0)

    def apply_stylesheet(self):
        """Mavi-siyah-beyaz temalı, sade ve modern bir stil uygular."""
        self.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                color: #000000;
                font-family: 'Segoe UI', 'Helvetica', 'Arial', sans-serif;
                font-size: 10.5pt;
            }

            QLabel#title_label {
                font-size: 28pt;
                font-weight: 700;
                color: #000000;
                padding: 20px;
                margin-bottom: 15px;
                border-bottom: 2px solid #007bff;
            }

            QLabel {
                font-size: 11pt;
                color: #000000;
            }

            QPushButton {
                background-color: #007bff;
                color: #ffffff;
                padding: 10px 24px;
                border: none;
                border-radius: 8px;
                font-weight: 600;
                font-size: 11pt;
                transition: all 0.3s ease;
            }

            QPushButton:hover {
                background-color: #0056b3;
            }

            QPushButton:pressed {
                background-color: #003e80;
            }

            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }

            QComboBox, QListWidget, QLineEdit, QScrollArea, QFrame {
                border: 1px solid #cccccc;
                border-radius: 6px;
                padding: 6px;
                background-color: #ffffff;
                selection-background-color: #007bff;
                selection-color: #ffffff;
            }

            QListWidget::item:selected {
                background-color: #007bff;
                color: white;
                border-radius: 4px;
            }

            QCheckBox {
                spacing: 8px;
                padding: 4px;
            }

            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 3px;
                border: 2px solid #007bff;
                background-color: #ffffff;
            }

            QCheckBox::indicator:checked {
                background-color: #007bff;
                border: 2px solid #0056b3;
            }

            QMessageBox {
                background-color: #ffffff;
                color: #000000;
                font-size: 10pt;
            }

            QMessageBox QPushButton {
                background-color: #007bff;
                color: white;
                padding: 6px 14px;
                border-radius: 5px;
                font-weight: 500;
            }

            QMessageBox QPushButton:hover {
                background-color: #0056b3;
            }

            QProgressBar {
                border: 1px solid #007bff;
                border-radius: 10px;
                background-color: #ffffff;
                text-align: center;
                height: 22px;
            }

            QProgressBar::chunk {
                background-color: #007bff;
                border-radius: 10px;
            }

            QScrollBar:vertical, QScrollBar:horizontal {
                background: #f0f0f0;
                border: none;
                width: 14px;
            }

            QScrollBar::handle:vertical, QScrollBar::handle:horizontal {
                background: #007bff;
                border-radius: 6px;
                min-height: 30px;
                min-width: 30px;
            }

            QScrollBar::add-line, QScrollBar::sub-line,
            QScrollBar::add-page, QScrollBar::sub-page {
                background: none;
                border: none;
            }
        """)

    def goto_page(self, index: int) -> None:
        """
        Belirli bir sayfaya gider ve o sayfayı yeniler.
        Sayfa indeksleri:
        0: Dosya Seçim Sayfası
        1: Günlük Grafik Veri Seçim Sayfası
        2: Günlük Grafikler Sayfası
        3: Aylık Grafikler Sayfası
        """
        self.stacked_widget.setCurrentIndex(index)
        # Her sayfaya geçişte ilgili sayfanın yenileme metodunu çağır
        if index == 1:
            self.data_selection_page.refresh()
        elif index == 2:
            self.daily_graphs_page.enter_page()
        elif index == 3:
            self.monthly_graphs_page.enter_page()

    def load_excel(self) -> None:
        """
        Seçilen Excel dosyasını ve sayfasını yükler.
        Sütun isimlerini dinamik olarak tanımlar ve metrik sütunlarını belirler.
        """
        # Excel yolu veya seçilen sayfa boşsa işlemi durdur
        if not self.excel_path or not self.selected_sheet:
            logging.warning("load_excel: Excel yolu veya seçilen sayfa boş. Veri yüklenemiyor.")
            return

        # Eğer aynı dosya ve sayfa zaten yüklüyse tekrar yüklemeyi önle
        # Bu, gereksiz dosya okuma işlemlerini azaltarak performansı artırır.
        if not self.df.empty and self.df.attrs.get('excel_path') == self.excel_path and \
                self.df.attrs.get('selected_sheet') == self.selected_sheet:
            logging.info(f"'{self.selected_sheet}' sayfasındaki veriler zaten yüklü. Yeniden yüklenmiyor.")
            return

        try:
            # Excel dosyasını belirtilen sayfadan yükle, ilk satırı başlık olarak kullan
            self.df = pd.read_excel(self.excel_path, sheet_name=self.selected_sheet, header=0)
            # Sütun isimlerini string tipine dönüştür
            self.df.columns = self.df.columns.astype(str)

            # Yüklenen dosya ve sayfa bilgilerini DataFrame özniteliklerine kaydet
            self.df.attrs['excel_path'] = self.excel_path
            self.df.attrs['selected_sheet'] = self.selected_sheet

            logging.info("Veri '%s' sayfasından yüklendi. Satır sayısı: %d", self.selected_sheet, len(self.df))

            # Sütun isimlerini dinamik olarak belirle (Excel sütun indekslerine göre)
            # A sütunu gruplama (tarih), B sütunu gruplanan (ürün)
            self.grouping_col_name = self.df.columns[excel_col_to_index('A')]
            self.grouped_col_name = self.df.columns[excel_col_to_index('B')]
            self.oee_col_name = None
            self.metric_cols = []

            # Seçilen sayfaya göre OEE ve metrik sütunlarını belirle
            if self.selected_sheet == "SMD-OEE":
                # OEE sütunu BP'de
                self.oee_col_name = self.df.columns[excel_col_to_index('BP')] if excel_col_to_index('BP') < len(
                    self.df.columns) else None
                # Metrik sütunları H'den BD'ye kadar, AP hariç
                start_col_index = excel_col_to_index('H')
                end_col_index = excel_col_to_index('BD')
                ap_col_index = excel_col_to_index('AP') # Hariç tutulacak sütun
                for i in range(start_col_index, end_col_index + 1):
                    if i < len(self.df.columns) and i != ap_col_index:
                        self.metric_cols.append(self.df.columns[i])
            elif self.selected_sheet == "ROBOT":
                # ROBOT sayfası için OEE sütunu BG olarak belirtildi, ancak mevcut kodda kullanılmıyor.
                # Eğer ROBOT sayfası için de OEE grafiği çizilecekse bu kısım güncellenmeli.
                # Günlük grafiklerde OEE sütunu kullanılmadığı için burada sadece metrikler tanımlanır.
                # Metrik sütunları H'den AU'ya kadar, AO hariç
                start_col_index = excel_col_to_index('H')
                end_col_index = excel_col_to_index('AU')
                ao_col_index = excel_col_to_index('AO') # Hariç tutulacak sütun

                for i in range(start_col_index, end_col_index + 1):
                    if i < len(self.df.columns) and i != ao_col_index:
                        self.metric_cols.append(self.df.columns[i])
            elif self.selected_sheet == "DALGA_LEHİM":
                # OEE sütunu BP'de
                self.oee_col_name = self.df.columns[excel_col_to_index('BP')] if excel_col_to_index('BP') < len(
                    self.df.columns) else None
                # Metrik sütunları H'den BD'ye kadar, AP hariç
                start_col_index = excel_col_to_index('H')
                end_col_index = excel_col_to_index('BD')
                ap_col_index = excel_col_to_index('AP') # Hariç tutulacak sütun
                for i in range(start_col_index, end_col_index + 1):
                    if i < len(self.df.columns) and i != ap_col_index:
                        self.metric_cols.append(self.df.columns[i])
            elif self.selected_sheet == "KAPLAMA-OEE":
                # OEE sütunu BG'de
                self.oee_col_name = self.df.columns[excel_col_to_index('BG')] if excel_col_to_index('BG') < len(
                    self.df.columns) else None
                # KAPLAMA-OEE için özel metrik sütunları tanımlanmadıysa, boş bırakılır veya varsayılan atanır.
                # Bu sayfa için sadece OEE grafiği istendiği için metrikler boş kalabilir.
                self.metric_cols = []

            logging.info("Gruplama sütunu tanımlandı: %s", self.grouping_col_name)
            logging.info("Gruplanan sütun tanımlandı: %s", self.grouped_col_name)
            logging.info("OEE sütunu tanımlandı: %s", self.oee_col_name)
            logging.info("Metrik sütunları tanımlandı: %s", self.metric_cols)

        except Exception as e:
            # Hata durumunda kullanıcıya bilgi ver ve DataFrame'i sıfırla
            QMessageBox.critical(self, "Veri Yükleme Hatası", f"Veri yüklenirken bir hata oluştu: {e}")
            logging.exception("Excel veri yükleme hatası.")
            self.df = pd.DataFrame() # Hata durumunda boş DataFrame ata
