import logging  # Loglama işlemleri için kullanılan modül
import datetime  # Tarih ve saat manipülasyonları için kullanılan modül
from pathlib import Path  # Dosya yolu işlemleri için kullanılan modül

from typing import List, Tuple, Any, Union, Dict  # Tip ipuçları için kullanılan modüller

import pandas as pd  # Veri manipülasyonu ve analizi için kullanılan kütüphane
import numpy as np  # Sayısal işlemler için kullanılan kütüphane

import matplotlib.pyplot as plt  # Grafik çizimi için kullanılan kütüphane
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas  # Matplotlib figürlerini Qt widget'ı olarak gömmek için
from matplotlib.ticker import PercentFormatter  # Y ekseninde yüzde formatı için

from PyQt5.QtCore import Qt  # Qt temel sınıfları ve sabitleri için
from PyQt5.QtWidgets import (  # Qt widget'ları için
    QWidget,  # Temel widget sınıfı
    QFileDialog,  # Dosya iletişim kutusu için
    QPushButton,  # Buton widget'ı için
    QLabel,  # Metin etiketi widget'ı için
    QVBoxLayout,  # Dikey düzen yöneticisi için
    QHBoxLayout,  # Yatay düzen yöneticisi için
    QComboBox,  # Açılır menü (seçim kutusu) widget'ı için
    QMessageBox,  # Mesaj kutusu için
    QProgressBar,  # İlerleme çubuğu widget'ı için
    QScrollArea,  # Kaydırılabilir alan widget'ı için
    QFrame,  # Çerçeve widget'ı için
    QLineEdit,  # Tek satırlık metin giriş kutusu için
    QSizePolicy  # Widget'ların boyutlandırma politikası için
)
from PyQt5 import QtGui  # QtGui modülü (QDoubleValidator için)

from logic.monthlyGraphWorker import MonthlyGraphWorker  # Arka planda grafik oluşturma işlemlerini yürüten worker sınıfı

class MonthlyGraphsPage(QWidget):
    """Aylık grafikler ve veri seçim sayfasını temsil eder."""

    def __init__(self, main_window: "MainWindow") -> None:
        """
        MonthlyGraphsPage sınıfının yapıcı metodu.

        Args:
            main_window: Ana uygulama penceresinin bir referansı.
        """
        super().__init__()
        self.main_window = main_window  # Ana pencere referansını saklar

        # Aylık grafiklerin görüntüleneceği container ve layout
        self.monthly_chart_container = QFrame(objectName="chartContainer")
        self.monthly_chart_layout = QVBoxLayout(self.monthly_chart_container)
        self.monthly_chart_layout.setAlignment(Qt.AlignCenter)  # Ortalamak için hizalama

        self.current_monthly_chart_figure = None  # Mevcut gösterilen Matplotlib figürü
        # Oluşturulan tüm figür verilerini saklayan liste (isim ve grafik verisi)
        self.figures_data_monthly: List[Tuple[str, Union[List[dict[str, Any]], Dict[str, Any]]]] = []
        self.current_page_monthly = 0  # Mevcut grafik sayfasının indeksi
        self.monthly_worker: MonthlyGraphWorker | None = None  # Aylık grafik oluşturma worker'ı
        self.prev_year_oee_for_plot: float | None = None  # Önceki yılın OEE değeri (grafik çizimi için)
        self.prev_month_oee_for_plot: float | None = None  # Önceki ayın OEE değeri (grafik çizimi için)
        self.current_graph_mode: str = "hat"  # Varsayılan grafik modu ("hat" veya "page")

        # Her hat/sayfa için OEE değerlerini saklamak için dictionary
        # Anahtar: Hat/Sayfa adı (string), Değer: (Önceki Yıl OEE, Önceki Ay OEE) tuple'ı
        self.cached_oee_values: Dict[str, Tuple[float | None, float | None]] = {}

        self.init_ui()  # Kullanıcı arayüzünü başlatır

    def init_ui(self):
        """Kullanıcı arayüzünü başlatır ve düzeni ayarlar."""
        # En dıştaki dikey düzen (başlık ve altındaki paneller için)
        outer_layout = QVBoxLayout(self)

        # Sayfa başlığı (en üstte)
        title_label = QLabel("<h2>Aylık Grafikler ve Veri Seçimi</h2>")
        title_label.setObjectName("title_label")  # CSS stil uygulamak için objectName
        title_label.setAlignment(Qt.AlignLeft)  # Sola hizala
        outer_layout.addWidget(title_label)
        outer_layout.addSpacing(10)  # Başlık ile paneller arasına boşluk

        # Sağ ve sol paneller için yatay düzen
        main_content_layout = QHBoxLayout()

        # Sol taraf için dikey düzen (menüler)
        left_panel_layout = QVBoxLayout()
        left_panel_layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)  # Üstte ve solda hizala
        left_panel_layout.setContentsMargins(20, 0, 10, 20)  # Kenar boşlukları (sol, üst, sağ, alt)

        # Grafik tipi seçim bölümü
        graph_type_selection_layout = QVBoxLayout()
        graph_type_selection_layout.addWidget(QLabel("<b>Grafik Tipi:</b>"))
        self.cmb_monthly_graph_type = QComboBox()  # Grafik tipi seçim kutusu
        self.cmb_monthly_graph_type.addItems(["OEE Grafikleri", "Dizgi Duruş Grafiği", "Dizgi Onay Dağılım Grafiği"])
        # Seçim değiştiğinde tetiklenecek sinyal bağlantısı
        self.cmb_monthly_graph_type.currentIndexChanged.connect(self.on_monthly_graph_type_changed)
        graph_type_selection_layout.addWidget(self.cmb_monthly_graph_type)
        left_panel_layout.addLayout(graph_type_selection_layout)
        left_panel_layout.addSpacing(15)

        # OEE Grafikleri için özel seçenekler widget'ı
        self.oee_options_widget = QWidget()
        oee_options_layout = QVBoxLayout(self.oee_options_widget)

        # Önceki Yıl OEE değeri girişi
        prev_year_oee_layout = QHBoxLayout()
        prev_year_oee_layout.addWidget(QLabel("Önceki Yılın OEE Değeri (%):"))
        self.txt_prev_year_oee = QLineEdit()
        self.txt_prev_year_oee.setPlaceholderText("Örn: 85.5")  # Yer tutucu metin
        # Sayısal giriş doğrulaması (0.0 ile 100.0 arası, 2 ondalık basamak)
        self.txt_prev_year_oee.setValidator(QtGui.QDoubleValidator(0.0, 100.0, 2))
        # Değişiklik: Text değiştiğinde değerleri önbelleğe al
        self.txt_prev_year_oee.textChanged.connect(self._cache_current_oee_values)
        prev_year_oee_layout.addWidget(self.txt_prev_year_oee)
        oee_options_layout.addLayout(prev_year_oee_layout)
        oee_options_layout.addSpacing(10)

        # Önceki Ay OEE değeri girişi
        prev_month_oee_layout = QHBoxLayout()
        prev_month_oee_layout.addWidget(QLabel("Önceki Ayın OEE Değeri (%):"))
        self.txt_prev_month_oee = QLineEdit()
        self.txt_prev_month_oee.setPlaceholderText("Örn: 82.0")
        # Sayısal giriş doğrulaması
        self.txt_prev_month_oee.setValidator(QtGui.QDoubleValidator(0.0, 100.0, 2))
        # Değişiklik: Text değiştiğinde değerleri önbelleğe al
        self.txt_prev_month_oee.textChanged.connect(self._cache_current_oee_values)
        prev_month_oee_layout.addWidget(self.txt_prev_month_oee)
        oee_options_layout.addLayout(prev_month_oee_layout)
        oee_options_layout.addSpacing(15)

        # OEE değerlerini grafiğe işlemek için buton
        self.btn_apply_oee_values = QPushButton("OEE Değerlerini Grafiğe Uygula")
        # Butona tıklanınca _apply_oee_values_to_current_graph metodunu çağır
        self.btn_apply_oee_values.clicked.connect(self._apply_oee_values_to_current_graph)
        oee_options_layout.addWidget(self.btn_apply_oee_values)
        oee_options_layout.addSpacing(15)

        # OEE grafik modları için düğmeler (yatayda yan yana)
        oee_buttons_layout = QHBoxLayout()
        self.btn_line_chart = QPushButton("Hat Grafikleri")
        # Hat grafiği modunu başlatmak için sinyal bağlantısı
        self.btn_line_chart.clicked.connect(lambda: self._start_monthly_graph_worker(graph_mode="hat"))
        self.btn_line_chart.setEnabled(False)  # Başlangıçta devre dışı
        oee_buttons_layout.addWidget(self.btn_line_chart)

        self.btn_page_chart = QPushButton("Sayfa Grafikleri")
        # Sayfa grafiği modunu başlatmak için sinyal bağlantısı
        self.btn_page_chart.clicked.connect(lambda: self._start_monthly_graph_worker(graph_mode="page"))
        self.btn_page_chart.setEnabled(False)  # Başlangıçta devre dışı
        oee_buttons_layout.addWidget(self.btn_page_chart)
        oee_options_layout.addLayout(oee_buttons_layout)

        left_panel_layout.addWidget(self.oee_options_widget)
        left_panel_layout.addStretch(1)  # Kalan boşluğu doldurmak için esnek boşluk

        # Diğer grafik türleri için boş widget (şimdilik)
        self.other_graphs_widget = QWidget()
        other_graphs_layout = QVBoxLayout(self.other_graphs_widget)
        left_panel_layout.addWidget(self.other_graphs_widget)
        self.other_graphs_widget.hide()  # Başlangıçta gizli

        # Ana içerik layout'una sol paneli ekle
        main_content_layout.addLayout(left_panel_layout, 1)  # 1 oranıyla sol paneli kapla (toplamda %20)

        # Sağ taraf için dikey düzen (grafik ve alt menü)
        right_panel_layout = QVBoxLayout()
        right_panel_layout.setAlignment(Qt.AlignTop | Qt.AlignRight)  # Üstte ve sağda hizala
        right_panel_layout.setContentsMargins(10, 0, 20, 20)  # Kenar boşlukları

        # Grafiklerin görüntüleneceği kaydırılabilir alan
        self.monthly_chart_scroll_area = QScrollArea()
        self.monthly_chart_scroll_area.setWidgetResizable(True)  # İçindeki widget'ın boyutunu otomatik ayarla
        self.monthly_chart_scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)  # Yatay kaydırma çubuğu gerektiğinde
        self.monthly_chart_scroll_area.setWidget(self.monthly_chart_container)  # Container'ı kaydırılabilir alana ata
        right_panel_layout.addWidget(self.monthly_chart_scroll_area, 1)  # Esnek boyut için stretch faktörü

        # Aylık grafik ilerleme çubuğu
        self.monthly_progress = QProgressBar()
        self.monthly_progress.setAlignment(Qt.AlignCenter)  # Metni ortala
        self.monthly_progress.setTextVisible(True)  # Metni görünür yap
        self.monthly_progress.hide()  # Başlangıçta gizli
        right_panel_layout.addWidget(self.monthly_progress)

        # Alt navigasyon düğmeleri (QHBoxLayout olarak kalacak)
        nav_bottom = QHBoxLayout()
        self.btn_monthly_back = QPushButton("← Geri")
        # Ana pencerede sayfa 0'a dönmek için sinyal bağlantısı
        self.btn_monthly_back.clicked.connect(lambda: self.main_window.goto_page(0))
        nav_bottom.addWidget(self.btn_monthly_back)

        self.lbl_monthly_page = QLabel("Sayfa 0 / 0")  # Sayfa etiketi
        self.lbl_monthly_page.setAlignment(Qt.AlignCenter)
        nav_bottom.addWidget(self.lbl_monthly_page)

        self.btn_prev_monthly = QPushButton("← Önceki Hat")
        self.btn_prev_monthly.clicked.connect(self.prev_monthly_page)  # Önceki sayfaya gitmek için
        self.btn_prev_monthly.setEnabled(False)  # Başlangıçta devre dışı
        nav_bottom.addWidget(self.btn_prev_monthly)

        self.btn_next_monthly = QPushButton("Sonraki Hat →")
        self.btn_next_monthly.clicked.connect(self.next_monthly_page)  # Sonraki sayfaya gitmek için
        self.btn_next_monthly.setEnabled(False)  # Başlangıçta devre dışı
        nav_bottom.addWidget(self.btn_next_monthly)

        self.btn_save_monthly_chart = QPushButton("Grafiği Kaydet (PNG/JPEG)")
        self.btn_save_monthly_chart.clicked.connect(self._save_monthly_chart_as_image)  # Grafiği kaydetmek için
        self.btn_save_monthly_chart.setEnabled(False)  # Başlangıçta devre dışı
        nav_bottom.addStretch(1)  # Butonları sağa yaslamak için esnek boşluk
        nav_bottom.addWidget(self.btn_save_monthly_chart)
        right_panel_layout.addLayout(nav_bottom)

        # Ana içerik layout'una sağ paneli ekle
        main_content_layout.addLayout(right_panel_layout, 4)  # 4 oranıyla sağ paneli kapla (toplamda %80)

        # En dıştaki layout'a ana içerik layout'unu ekle
        outer_layout.addLayout(main_content_layout)

        # Başlangıçta ilk grafik tipine göre UI'ı ayarla
        self.on_monthly_graph_type_changed(0)

    def enter_page(self) -> None:
        """Grafiği temizler ve bu sayfaya girildiğinde düğme durumlarını günceller."""
        self.clear_monthly_chart_canvas()  # Grafik tuvalini temizle
        self.btn_save_monthly_chart.setEnabled(False)  # Kaydet butonunu devre dışı bırak
        # Sayfa etiketini ve navigasyon butonlarını güncelle
        self.update_monthly_page_label(graph_mode=self.current_graph_mode)
        self.update_monthly_navigation_buttons(graph_mode=self.current_graph_mode)

        selected_type = self.cmb_monthly_graph_type.currentText()  # Seçili grafik tipini al
        if selected_type == "OEE Grafikleri":
            self.oee_options_widget.show()  # OEE seçeneklerini göster
            self.btn_line_chart.setEnabled(True)  # Hat grafiği butonunu etkinleştir
            self.btn_page_chart.setEnabled(True)  # Sayfa grafiği butonunu etkinleştir
            self.btn_apply_oee_values.setEnabled(True)  # Apply button enabled for OEE graphs

            # İlk grafik sayfasının OEE değerlerini önbellekten yükle
            if self.figures_data_monthly:
                self._load_cached_oee_values(self.figures_data_monthly[0][0])
            else:
                self.txt_prev_year_oee.clear()  # Boşalt
                self.txt_prev_month_oee.clear()  # Boşalt

            # OEE Grafikleri seçildiğinde varsayılan olarak Hat Grafikleri'ni başlat
            # Sadece bir dosya yüklüyse ve henüz grafik oluşturulmamışsa otomatik başlat
            if self.main_window.excel_path and not self.main_window.df.empty and not self.figures_data_monthly:
                self._start_monthly_graph_worker(graph_mode="hat")

        else:
            self.oee_options_widget.hide()  # OEE seçeneklerini gizle
            self.btn_line_chart.setEnabled(False)  # Hat grafiği butonunu devre dışı bırak
            self.btn_page_chart.setEnabled(False)  # Sayfa grafiği butonunu devre dışı bırak
            self.btn_apply_oee_values.setEnabled(False)  # Apply button disabled for other graph types

    def on_monthly_graph_type_changed(self, index: int):
        """
        Aylık grafik türü seçimi değiştiğinde ilgili seçenekleri gösterir/gizler.

        Args:
            index: Seçilen öğenin indeksi.
        """
        selected_type = self.cmb_monthly_graph_type.currentText()  # Seçili grafik tipini al
        self.clear_monthly_chart_canvas()  # Grafik tuvalini temizle
        self.btn_save_monthly_chart.setEnabled(False)  # Kaydet butonunu devre dışı bırak
        self.figures_data_monthly.clear()  # Önbellekteki grafik verilerini temizle
        self.current_page_monthly = 0  # Sayfa indeksini sıfırla
        self.cached_oee_values.clear()  # Grafik tipi değiştiğinde önbelleği temizle
        self.txt_prev_year_oee.clear()  # OEE seçildiğinde giriş alanlarını temizle
        self.txt_prev_month_oee.clear()

        if selected_type == "OEE Grafikleri":
            self.current_graph_mode = "hat"  # Varsayılan modu "hat" olarak ayarla
            self.oee_options_widget.show()  # OEE seçeneklerini göster
            self.other_graphs_widget.hide()  # Diğer grafik seçeneklerini gizle
            self.btn_line_chart.setEnabled(True)  # Hat grafiği butonunu etkinleştir
            self.btn_page_chart.setEnabled(True)  # Sayfa grafiği butonunu etkinleştir
            self.btn_apply_oee_values.setEnabled(True)
            # OEE Grafikleri seçildiğinde varsayılan olarak Hat Grafikleri'ni başlat
            # Sadece bir dosya yüklüyse ve henüz grafik oluşturulmamışsa otomatik başlat
            if self.main_window.excel_path and not self.main_window.df.empty:
                self._start_monthly_graph_worker(graph_mode="hat")
        elif selected_type in ["Dizgi Onay Dağılım Grafiği", "Dizgi Duruş Grafiği"]:
            self.current_graph_mode = "hat"  # Varsayılan modu "hat" olarak ayarla
            self.oee_options_widget.hide()  # OEE seçeneklerini gizle
            self.other_graphs_widget.show()  # Diğer grafik seçeneklerini göster
            self.btn_line_chart.setEnabled(False)  # Hat grafiği butonunu devre dışı bırak
            self.btn_page_chart.setEnabled(False)  # Sayfa grafiği butonunu devre dışı bırak
            self.btn_apply_oee_values.setEnabled(False)
            # Bu tipler için otomatik başlatma, OEE değerleri gerektirmediği için
            # doğrudan worker'ı başlatırız.
            if self.main_window.excel_path and not self.main_window.df.empty:
                self._start_monthly_graph_worker(graph_mode="hat")
        else:
            self.current_graph_mode = "hat"
            self.oee_options_widget.hide()
            self.other_graphs_widget.show()
            self.btn_line_chart.setEnabled(False)
            self.btn_page_chart.setEnabled(False)
            self.btn_apply_oee_values.setEnabled(False)

        # Sayfa etiketini ve navigasyon butonlarını güncelle
        self.update_monthly_page_label(graph_mode=self.current_graph_mode)
        self.update_monthly_navigation_buttons(graph_mode=self.current_graph_mode)

    def clear_monthly_chart_canvas(self):
        """Aylık grafik tuvallerini temizler."""
        # monthly_chart_layout içindeki tüm widget'ları siler
        while self.monthly_chart_layout.count():
            item = self.monthly_chart_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()  # Widget'ı güvenli bir şekilde sil

    def _start_monthly_graph_worker(self, graph_mode: str):
        """
        Aylık grafik çalışanını başlatır.

        Args:
            graph_mode: "hat" veya "page" olarak grafik modunu belirtir.
        """
        # Excel dosyasının yüklü olup olmadığını kontrol et
        if not self.main_window.excel_path or self.main_window.df.empty:
            QMessageBox.information(self, "Dosya Yüklü Değil",
                                    "Lütfen önce bir Excel dosyası yükleyin.")
            self.monthly_progress.hide()  # İlerleme çubuğunu gizle
            self.btn_save_monthly_chart.setEnabled(False)  # Kaydet butonunu devre dışı bırak
            return

        self.clear_monthly_chart_canvas()  # Grafik tuvalini temizle
        self.btn_save_monthly_chart.setEnabled(False)  # Kaydet butonunu devre dışı bırak
        self.figures_data_monthly.clear()  # Önceki grafik verilerini temizle
        self.current_page_monthly = 0  # Sayfa indeksini sıfırla
        self.monthly_progress.setValue(0)  # İlerleme çubuğunu sıfırla
        self.monthly_progress.show()  # İlerleme çubuğunu göster
        self.current_graph_mode = graph_mode  # Güncel grafik modunu ayarla
        # Sayfa etiketini ve navigasyon butonlarını güncelle
        self.update_monthly_page_label(graph_mode=self.current_graph_mode)
        self.update_monthly_navigation_buttons(graph_mode=self.current_graph_mode)

        # Çalışan bir worker varsa durdur ve bekle
        if self.monthly_worker and self.monthly_worker.isRunning():
            self.monthly_worker.quit()
            self.monthly_worker.wait()

        prev_year_oee = None
        prev_month_oee = None

        # Sadece OEE grafikleri için giriş alanlarından değerleri al
        if self.cmb_monthly_graph_type.currentText() == "OEE Grafikleri":
            try:
                if self.txt_prev_year_oee.text():
                    # Virgülü nokta ile değiştirerek float'a dönüştür
                    prev_year_oee = float(self.txt_prev_year_oee.text().replace(",", "."))
                if self.txt_prev_month_oee.text():
                    # Virgülü nokta ile değiştirerek float'a dönüştür
                    prev_month_oee = float(self.txt_prev_month_oee.text().replace(",", "."))
            except ValueError:
                QMessageBox.warning(self, "Geçersiz Giriş",
                                    "Lütfen Önceki Yıl/Ay OEE değerlerini geçerli sayı olarak girin.")
                self.monthly_progress.hide()  # İlerleme çubuğunu gizle
                return

        # Yeni MonthlyGraphWorker örneği oluştur ve başlat
        self.monthly_worker = MonthlyGraphWorker(
            excel_path=self.main_window.excel_path,
            current_df=self.main_window.df,
            graph_mode=self.current_graph_mode,
            graph_type=self.cmb_monthly_graph_type.currentText(),
            prev_year_oee=prev_year_oee,  # Bu değerler worker'a iletilir
            prev_month_oee=prev_month_oee,  # Bu değerler worker'a iletilir
            main_window=self.main_window
        )
        # Worker'ın finished sinyali _on_monthly_graphs_generated metoduna bağlanır
        self.monthly_worker.finished.connect(self._on_monthly_graphs_generated)
        # Worker'ın progress sinyali ilerleme çubuğunun değerine bağlanır
        self.monthly_worker.progress.connect(self.monthly_progress.setValue)
        # Worker'ın error sinyali _on_monthly_graph_error metoduna bağlanır
        self.monthly_worker.error.connect(self._on_monthly_graph_error)
        self.monthly_worker.start()  # Worker'ı başlat

    def _on_monthly_graphs_generated(self,
                                     figures_data_raw: List[Tuple[str, Union[List[dict[str, Any]], Dict[str, Any]]]],
                                     prev_year_oee: float | None, prev_month_oee: float | None):
        """
        MonthlyGraphWorker'dan gelen sonuçları işler.

        Args:
            figures_data_raw: Oluşturulan grafik verilerinin listesi.
            prev_year_oee: Hesaplamalar için kullanılan önceki yılın OEE değeri.
            prev_month_oee: Hesaplamalar için kullanılan önceki ayın OEE değeri.
        """
        self.monthly_progress.setValue(100)  # İlerleme çubuğunu tamamla
        self.monthly_progress.hide()  # İlerleme çubuğunu gizle

        if not figures_data_raw:
            QMessageBox.information(self, "Veri Yok",
                                    "Aylık grafik oluşturulamadı. Seçilen kriterlere göre veri bulunamadı.")
            self.btn_save_monthly_chart.setEnabled(False)  # Kaydet butonunu devre dışı bırak
            return

        self.figures_data_monthly = figures_data_raw  # Grafik verilerini sakla
        self.prev_year_oee_for_plot = prev_year_oee  # Çizim için önceki yıl OEE'yi sakla
        self.prev_month_oee_for_plot = prev_month_oee  # Çizim için önceki ay OEE'yi sakla

        # İlk grafiğin OEE değerlerini önbelleğe al ve giriş alanlarına yükle
        if self.cmb_monthly_graph_type.currentText() == "OEE Grafikleri" and self.figures_data_monthly:
            # Sadece worker çalıştıktan sonra ve grafikler oluştuğunda önbelleği başlangıç değerleriyle doldur.
            # Her bir grafik için önbelleğe başlangıç değerlerini kaydet.
            # Şu anki mantık, worker'a verdiğiniz prev_year_oee ve prev_month_oee değerlerinin
            # ilk grafik için geçerli olmasıdır. Sonraki grafikler için bu değerlerin
            # otomatik olarak doldurulması istenir.
            for name, _ in self.figures_data_monthly:
                # Eğer daha önce kaydedilmemişse, başlangıçtaki global değerleri kullan
                if name not in self.cached_oee_values:
                    self.cached_oee_values[name] = (self.prev_year_oee_for_plot, self.prev_month_oee_for_plot)
            self._load_cached_oee_values(self.figures_data_monthly[self.current_page_monthly][0])

        self.display_current_page_graphs_monthly()  # Mevcut sayfadaki grafiği göster
        self.btn_save_monthly_chart.setEnabled(True)  # Kaydet butonunu etkinleştir

    def _on_monthly_graph_error(self, message: str):
        """
        MonthlyGraphWorker'dan gelen bir hata mesajını görüntüler.

        Args:
            message: Görüntülenecek hata mesajı.
        """
        QMessageBox.critical(self, "Hata", message)  # Hata mesajı kutusu göster
        self.monthly_progress.setValue(0)  # İlerleme çubuğunu sıfırla
        self.monthly_progress.hide()  # İlerleme çubuğunu gizle
        self.btn_save_monthly_chart.setEnabled(False)  # Kaydet butonunu devre dışı bırak

    def _cache_current_oee_values(self):
        """Mevcut hat/sayfa için OEE değerlerini önbelleğe alır."""
        if not self.figures_data_monthly:
            return

        current_entity_name = self.figures_data_monthly[self.current_page_monthly][0]  # Mevcut varlık adını al
        try:
            # Önceki yıl OEE değerini al (virgülü noktaya çevir)
            prev_year = float(
                self.txt_prev_year_oee.text().replace(",", ".")) if self.txt_prev_year_oee.text() else None
            # Önceki ay OEE değerini al (virgülü noktaya çevir)
            prev_month = float(
                self.txt_prev_month_oee.text().replace(",", ".")) if self.txt_prev_month_oee.text() else None
            # Önbelleğe kaydet
            self.cached_oee_values[current_entity_name] = (prev_year, prev_month)
        except ValueError:
            # Geçersiz giriş durumunda, önbelleği güncelleme (veya varsayılan None yapma)
            self.cached_oee_values[current_entity_name] = (None, None)

    def _load_cached_oee_values(self, entity_name: str):
        """
        Belirtilen hat/sayfa için önbelleğe alınmış OEE değerlerini giriş alanlarına yükler.

        Args:
            entity_name: Yüklenecek varlık (hat/sayfa) adı.
        """
        prev_year, prev_month = self.cached_oee_values.get(entity_name, (None, None))  # Önbellekten değerleri al

        # Signals Blocked: textChanged sinyalini geçici olarak engelle
        # Bu, _cache_current_oee_values'ın tekrar çağrılmasını engeller
        self.txt_prev_year_oee.blockSignals(True)
        self.txt_prev_month_oee.blockSignals(True)

        # Giriş alanlarını güncel değerlerle doldur
        self.txt_prev_year_oee.setText(str(prev_year) if prev_year is not None else "")
        self.txt_prev_month_oee.setText(str(prev_month) if prev_month is not None else "")

        # Signals Blocked: sinyalleri tekrar etkinleştir
        self.txt_prev_year_oee.blockSignals(False)
        self.txt_prev_month_oee.blockSignals(False)

    def _apply_oee_values_to_current_graph(self):
        """Giriş alanlarındaki OEE değerlerini alarak mevcut grafiği yeniden çizer."""
        if self.cmb_monthly_graph_type.currentText() != "OEE Grafikleri":
            QMessageBox.information(self, "Geçersiz İşlem", "OEE değerleri sadece OEE Grafikleri için uygulanabilir.")
            return

        if not self.figures_data_monthly:
            QMessageBox.information(self, "Grafik Yok", "Uygulanacak bir grafik bulunmamaktadır.")
            return

        # Giriş alanlarındaki değerleri al ve önbelleğe kaydet
        self._cache_current_oee_values()

        # Mevcut grafiği tekrar çizerek yeni değerleri uygula
        self.display_current_page_graphs_monthly()

    def display_current_page_graphs_monthly(self) -> None:
        """Mevcut sayfadaki aylık grafiği görüntüler."""
        self.clear_monthly_chart_canvas()  # Grafik tuvalini temizle

        total_pages = len(self.figures_data_monthly)  # Toplam sayfa sayısını al

        # Geçerli sayfa indeksini ayarla (sınırların dışına çıkmayı engelle)
        if self.current_page_monthly >= total_pages and total_pages > 0:
            self.current_page_monthly = total_pages - 1
        elif total_pages == 0:
            self.current_page_monthly = 0

        if not self.figures_data_monthly:
            # Grafik verisi yoksa bir etiket göster
            no_data_label = QLabel("Gösterilecek aylık grafik bulunamadı.", alignment=Qt.AlignCenter)
            self.monthly_chart_layout.addWidget(no_data_label)
            self.current_monthly_chart_figure = None
            self.btn_save_monthly_chart.setEnabled(False)
            # Sayfa etiketini ve navigasyon butonlarını güncelle
            self.update_monthly_page_label(graph_mode=self.current_graph_mode)
            self.update_monthly_navigation_buttons(graph_mode=self.current_graph_mode)
            return

        # Mevcut sayfanın grafik verilerini al
        name, data_container = self.figures_data_monthly[self.current_page_monthly]

        # Değişiklik: OEE grafiği ise, giriş alanlarını güncel OEE değerleriyle doldur
        if self.cmb_monthly_graph_type.currentText() == "OEE Grafikleri":
            self._load_cached_oee_values(name)  # Önbellekteki OEE değerlerini yükle
            # Grafik çizimi için güncel giriş alanlarındaki değerleri kullan
            try:
                self.prev_year_oee_for_plot = float(
                    self.txt_prev_year_oee.text().replace(",", ".")) if self.txt_prev_year_oee.text() else None
                self.prev_month_oee_for_plot = float(
                    self.txt_prev_month_oee.text().replace(",", ".")) if self.txt_prev_month_oee.text() else None
            except ValueError:
                self.prev_year_oee_for_plot = None
                self.prev_month_oee_for_plot = None
                QMessageBox.warning(self, "Geçersiz Giriş",
                                    "Kaydedilmiş OEE değerleri geçersiz. Lütfen doğru formatta girin.")

        # Matplotlib figürü ve eksenleri oluştur
        fig_width_inches = 12.0
        fig_height_inches = 7.0

        # "Dizgi Duruş Grafiği" için yüksekliği özel olarak ayarla
        if self.cmb_monthly_graph_type.currentText() == "Dizgi Duruş Grafiği":
            metric_sums_dict = data_container["metrics"]
            num_bars = len(metric_sums_dict)
            # Çubuk sayısına göre dinamik yükseklik hesapla.
            # Verilerinize ve istenen görünüme göre bu değerleri ayarlamanız gerekebilir.
            # Temel bir yükseklik artı çubuk başına bir artış.
            base_height = 6.0  # Minimum yükseklik
            height_per_bar = 0.5  # Aşırı kalabalığı önlemek için çubuk başına ek yükseklik
            # Etiketler ve ek açıklamalar için yeterli alan olduğundan emin olun ve aşırı uzun grafikleri önleyin
            fig_height_inches = max(base_height, min(25.0, num_bars * height_per_bar + 2.0))
            fig_width_inches = 14.0  # Etiketlerin daha iyi görünmesi için biraz daha geniş

        fig, ax = plt.subplots(figsize=(fig_width_inches, fig_height_inches), dpi=120)
        background_color = 'white'
        fig.patch.set_facecolor(background_color)  # Figür arka plan rengi
        ax.set_facecolor(background_color)  # Eksen arka plan rengi

        # Çerçeve çizgilerini ayarla
        ax.spines['top'].set_visible(False)  # Üst çerçeveyi gizle
        ax.spines['right'].set_visible(False)  # Sağ çerçeveyi gizle
        ax.spines['left'].set_linewidth(1.5)  # Sol çerçevenin kalınlığı
        ax.spines['bottom'].set_linewidth(1.5)  # Alt çerçevenin kalınlığı
        ax.grid(False)  # Izgarayı gizle

        if self.cmb_monthly_graph_type.currentText() == "OEE Grafikleri":
            grouped_oee = pd.DataFrame(data_container)  # Veri çerçevesi oluştur
            grouped_oee['Tarih'] = pd.to_datetime(grouped_oee['Tarih'])  # Tarih sütununu datetime'a çevir

            dates = grouped_oee['Tarih']
            # İşlenmiş OEE değeri varsa onu kullan, yoksa ham OEE değerini kullan
            oee_values = grouped_oee['OEE_Degeri_Processed'] if 'OEE_Degeri_Processed' in grouped_oee.columns else \
                grouped_oee['OEE_Degeri']

            line_color = '#1f77b4'  # Çizgi rengi

            x_indices = np.arange(len(dates))  # X ekseni indeksleri
            # OEE değerlerini çiz
            ax.plot(x_indices, oee_values, marker='o', markersize=8, color=line_color, linewidth=2, label=name)
            # Beyaz içi boş noktalarla çizgiyi vurgula
            ax.plot(x_indices, oee_values, 'o', markersize=6, color='white', markeredgecolor=line_color,
                    markeredgewidth=1.5, zorder=5)

            # Sayfa modunda ve 'OEE_Degeri_Half' sütunu varsa çift vardiya OEE'sini çiz
            if self.current_graph_mode == "page" and 'OEE_Degeri_Half' in grouped_oee.columns:
                half_oee_values = grouped_oee['OEE_Degeri_Half']

                # Türkçe ay isimleri sözlüğü
                month_names_turkish = {
                    1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
                    7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
                }
                month_name_for_half_oee = ""
                if not dates.empty:
                    first_date_in_data = dates.min()
                    month_name_for_half_oee = month_names_turkish.get(first_date_in_data.month,
                                                                      first_date_in_data.strftime('%B')).capitalize()
                else:
                    month_name_for_half_oee = datetime.date.today().strftime('%B').capitalize()

                average_half_oee = half_oee_values.mean() if not half_oee_values.empty else 0.0

                half_oee_label = f"{month_name_for_half_oee} Ayı Çift Vardiya Durumunda OEE ({average_half_oee * 100:.1f}%)"

                ax.plot(x_indices, half_oee_values, color='#ADD8E6', linestyle='--', linewidth=1.5,
                        label=half_oee_label)
                ax.plot(x_indices, half_oee_values, 'o', markersize=6, markerfacecolor='#ADD8E6',
                        markeredgecolor='#ADD8E6', markeredgewidth=1.5, zorder=6)

                if not half_oee_values.empty and len(x_indices) > 0:
                    last_x_index = x_indices[-1]
                    ax.annotate(f'{average_half_oee * 100:.1f}%', (last_x_index, half_oee_values.iloc[-1]),
                                textcoords="offset points", xytext=(5, -5), ha='left', va='center',
                                fontsize=9, fontweight='bold', color='#ADD8E6')

            # Her OEE noktasına değerini etiket olarak ekle
            for i, (x, y) in enumerate(zip(x_indices, oee_values)):
                if pd.notna(y) and y > 0:
                    ax.annotate(f'{y * 100:.1f}%', (x, y), textcoords="offset points", xytext=(0, 10), ha='center',
                                fontsize=8, fontweight='bold')

            overall_calculated_average = np.mean(oee_values) if not oee_values.empty else 0

            # Burada self.prev_year_oee_for_plot ve self.prev_month_oee_for_plot
            # display_current_page_graphs_monthly başında güncellendiği için doğru değerleri içerir.
            # Önceki yıl OEE çizgisi
            if self.prev_year_oee_for_plot is not None:
                y_val = self.prev_year_oee_for_plot / 100
                ax.axhline(y_val, color='red', linestyle='--', linewidth=1.5,
                           label=f'Önceki Yıl OEE ({self.prev_year_oee_for_plot:.1f}%)')
                ax.text(1.01, y_val, f'{self.prev_year_oee_for_plot:.1f}%',
                        transform=ax.transAxes, color='red', va='center', ha='left', fontsize=9, fontweight='bold')

            # Önceki ay OEE çizgisi
            if self.prev_month_oee_for_plot is not None:
                y_val = self.prev_month_oee_for_plot / 100
                ax.axhline(y_val, color='orange', linestyle='--', linewidth=1.5,
                           label=f'Önceki Ay OEE ({self.prev_month_oee_for_plot:.1f}%)')
                ax.text(1.01, y_val, f'{self.prev_month_oee_for_plot:.1f}%',
                        transform=ax.transAxes, color='orange', va='center', ha='left', fontsize=9, fontweight='bold')

            # Bu ayın ortalama OEE çizgisi
            if overall_calculated_average > 0:
                y_val = overall_calculated_average
                # Türkçe ay isimleri
                month_names_turkish = {
                    1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
                    7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
                }
                month_name = ""
                if not dates.empty:
                    first_date_in_data = dates.min()
                    month_name = month_names_turkish.get(first_date_in_data.month,
                                                         first_date_in_data.strftime('%B')).capitalize()
                else:
                    month_name = datetime.date.today().strftime('%B').capitalize()

                ax.axhline(y_val, color='purple', linestyle='--', linewidth=1.5,
                           label=f'{month_name} OEE ({overall_calculated_average * 100:.1f}%)')
                ax.text(1.01, y_val, f'{overall_calculated_average * 100:.1f}%',
                        transform=ax.transAxes, color='purple', va='center', ha='left', fontsize=9, fontweight='bold')

            # X ekseni etiketlerini ayarla
            ax.set_xticks(x_indices)
            ax.set_xticklabels([d.strftime('%d.%m.%Y') for d in dates])
            fig.autofmt_xdate(rotation=45)  # Tarih etiketlerini otomatik döndür

            # Y eksenini yüzde olarak formatla
            ax.yaxis.set_major_formatter(PercentFormatter(xmax=1, decimals=0))
            ax.set_yticks(np.arange(0.0, 1.001, 0.25))  # Y ekseni tick'lerini ayarla
            ax.set_ylim(bottom=-0.05, top=1.05)  # Y ekseni limitlerini ayarla

            ax.set_xlabel("Tarih", fontsize=12, fontweight='bold')  # X ekseni etiketi
            ax.set_ylabel("OEE (%)", fontsize=12, fontweight='bold')  # Y ekseni etiketi

            # Başlık için ay ismini al
            month_names_turkish = {
                1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
                7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
            }
            month_name = ""
            if not dates.empty:
                first_date_in_data = dates.min()
                month_name = month_names_turkish.get(first_date_in_data.month,
                                                     first_date_in_data.strftime('%B')).capitalize()
            else:
                month_name = datetime.date.today().strftime('%B').capitalize()

            # Grafik başlığını ayarla
            if self.current_graph_mode == "page":
                cleaned_name = name.replace('_', ' ').replace('-', ' ')
                if cleaned_name.endswith("OEE"):
                    cleaned_name = cleaned_name.rsplit(' ', 1)[0]
                chart_title = f"{month_name} {cleaned_name} OEE"
            else:
                chart_title = f"{name} {month_name} OEE"

            ax.set_title(chart_title, fontsize=24, color='#2c3e50', fontweight='bold')

            ax.legend(loc='upper left', bbox_to_anchor=(1.02, 0), fontsize=10)  # Lejantı ayarla
            fig.subplots_adjust(right=0.60)  # Lejant için sağda boşluk bırak

        elif self.cmb_monthly_graph_type.currentText() == "Dizgi Onay Dağılım Grafiği":
            labels = [d["label"] for d in data_container]  # Etiketleri al
            values = [d["value"] for d in data_container]  # Değerleri al

            colors = ['#00008B', '#ff7f0e']  # Pasta dilimi renkleri

            total_sum = sum(values)  # Toplam değeri hesapla

            # Pasta dilimi yüzdesini ve süresini formatlayan fonksiyon
            def func(pct, allvals):
                absolute = int(np.round(pct / 100. * total_sum))
                hours = absolute // 3600
                minutes = (absolute % 3600) // 60
                seconds = absolute % 60
                return f"{hours:02d}:{minutes:02d}:{seconds:02d}; {pct:.0f}%"

            # Pasta grafiğini çiz
            wedges, texts, autotexts = ax.pie(
                values,
                autopct=lambda pct: func(pct, values),  # Otomatik yüzde formatı
                startangle=90,  # Başlangıç açısı
                colors=colors,  # Renkler
                wedgeprops=dict(edgecolor='black', linewidth=1.5)  # Dilim özellikleri
            )

            # Otomatik metin etiketlerini ayarla
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontsize(14)
                autotext.set_fontweight('bold')

            ax.axis('equal')  # Pasta grafiğini daire şeklinde tut

            # Lejantı ayarla
            ax.legend(wedges, labels,
                      title="Hatlar",
                      loc="upper right",
                      bbox_to_anchor=(1.2, 1),
                      fontsize=10,
                      title_fontsize=12)

            chart_title = f"Dizgi Onay Dağılımı"
            ax.set_title(chart_title, fontsize=24, color='#2c3e50', fontweight='bold')  # Grafik başlığı
            fig.tight_layout()  # Düzeni sıkılaştır

        elif self.cmb_monthly_graph_type.currentText() == "Dizgi Duruş Grafiği":
            metric_sums_dict = data_container["metrics"]  # Metrik toplamlarını al
            total_overall_sum = data_container["total_overall_sum"]  # Genel toplamı al
            cumulative_percentages_dict = data_container["cumulative_percentages"]  # Kümülatif yüzdeleri al

            metric_sums = pd.Series(metric_sums_dict)  # Metrik toplamlarını Series'e çevir
            cumulative_percentage = pd.Series(cumulative_percentages_dict)  # Kümülatif yüzdeleri Series'e çevir

            ax2 = ax.twinx()  # İkinci bir y ekseni oluştur

            bar_color = '#AECDCB'  # Çubuk rengi
            line_color = '#6B0000'  # Çizgi rengi

            # Çubuk grafiğini çiz (süreleri dakikaya çevirerek)
            bars = ax.bar(metric_sums.index, metric_sums.values / 60, color=bar_color, alpha=0.8, edgecolor='black',
                          linewidth=1.5)

            # Kümülatif yüzde çizgisini çiz
            ax2.plot(metric_sums.index, cumulative_percentage, color=line_color, linestyle='-', linewidth=2, zorder=1)
            # Kümülatif yüzde noktalarını çiz
            ax2.plot(metric_sums.index, cumulative_percentage, 'o', markersize=8, markerfacecolor='white',
                     markeredgecolor=line_color, markeredgewidth=2, zorder=2)

            x_min_data, x_max_data = ax.get_xlim()

            # %80 Pareto çizgisini ekle
            normalized_xmin = (0 - x_min_data) / (x_max_data - x_min_data)
            normalized_xmax = (len(metric_sums.index) - 1 - x_min_data) / (x_max_data - x_min_data)
            ax2.axhline(80, color='#B0B0B0', linestyle='--', linewidth=1.5, xmin=normalized_xmin, xmax=normalized_xmax)

            ax.grid(False)  # Izgarayı gizle
            ax2.grid(False)  # İkinci eksenin ızgarasını gizle

            ax.set_xlabel("")  # X ekseni etiketini boş bırak

            ax.set_xticks(np.arange(len(metric_sums.index)))  # X ekseni tick'lerini ayarla
            # X ekseni etiketlerinin çakışmasını önlemek için döndürme ve hizalama ayarları
            ax.set_xticklabels(metric_sums.index, fontsize=10, fontweight='bold', rotation=60, ha='right')

            # --- DÜZELTME: text_label'ı tanımla ve açıklama mantığını geliştir ---
            # Her çubuğa süre ve yüzde etiketleri ekle
            for i, bar in enumerate(bars):
                value_seconds = metric_sums.values[i]
                percentage = (value_seconds / total_overall_sum) * 100 if total_overall_sum > 0 else 0
                duration_hours = int(value_seconds // 3600)
                duration_minutes = int((value_seconds % 3600) // 60)
                duration_seconds = int(value_seconds % 60)

                # Metin etiketini formatla
                text_label = f"{duration_hours:02d}:{duration_minutes:02d}:{duration_seconds:02d}\n({percentage:.1f}%)"

                # Metin etiketleri için dikey ofset hesapla
                # Tutarlılık için sabit bir ofset kullanın veya çubuk yüksekliğine göre ayarlayın
                # Küçük bir mutlak ofset (örn. 5 nokta) yüzde olarak ayarlamaktan daha güvenilir olabilir
                text_offset_points = 10  # Etiketler çubuklara çok yakınsa bu değeri artırın

                ax.annotate(text_label,
                            (bar.get_x() + bar.get_width() / 2, bar.get_height()),  # Çubuğun üstünde konumlandır
                            textcoords="offset points",  # Konumdan ofset
                            xytext=(0, text_offset_points),  # (x_ofset, y_ofset)
                            ha='center', va='bottom',  # Yatay ve dikey hizalama
                            fontsize=9, fontweight='bold', color='black')  # Okunabilirlik için ayarlanmış yazı tipi boyutu

            # Çizgi üzerindeki kümülatif yüzde etiketlerini ekle
            for i, (x, y) in enumerate(zip(metric_sums.index, cumulative_percentage)):
                ax2.annotate(f'{y:.1f}%', (x, y),
                             textcoords="offset points", xytext=(0, -15),  # Noktanın altında ofset
                             ha='center', va='top', fontsize=9, color=line_color, fontweight='bold')
            # --- DÜZELTME SONU ---

            ax.set_ylabel("Süre (Dakika)", fontsize=12, fontweight='bold', color=bar_color)  # Birincil y ekseni etiketi
            ax2.set_ylabel("Kümülatif Yüzde (%)", fontsize=12, fontweight='bold', color=line_color)  # İkincil y ekseni etiketi

            chart_title = "Genel Dizgi Duruş Pareto Analizi"
            # Tarih aralığına göre başlığı güncelle
            if not self.main_window.df.empty and 'Tarih' in self.main_window.df.columns:
                df_dates = pd.to_datetime(self.main_window.df['Tarih'], errors='coerce').dropna()
                if not df_dates.empty:
                    min_date = df_dates.min()
                    max_date = df_dates.max()
                    month_names_turkish = {
                        1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
                        7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
                    }
                    if min_date.month == max_date.month and min_date.year == max_date.year:
                        month_name = month_names_turkish.get(min_date.month, min_date.strftime('%B')).capitalize()
                        chart_title = f"{month_name} Ayı Dizgi Duruşları"
                    elif min_date.year == max_date.year:
                        first_month_name = month_names_turkish.get(min_date.month, min_date.strftime('%B')).capitalize()
                        last_month_name = month_names_turkish.get(max_date.month, max_date.strftime('%B')).capitalize()
                        chart_title = f"{min_date.year} Yılı {first_month_name}-{last_month_name} Ayları Dizgi Duruşları"
                    else:
                        chart_title = f"{min_date.year}-{max_date.year} Yılları Dizgi Duruşları"

            ax.set_title(chart_title, fontsize=24, color='#363636', fontweight='bold')

            ax.set_ylim(bottom=0)  # Birincil y ekseni alt limiti
            ax2.set_ylim(0, 100)  # İkincil y ekseni limitleri

            ax.spines['top'].set_visible(False)
            ax2.spines['top'].set_visible(False)

            fig.tight_layout()  # Düzeni sıkılaştır

        canvas = FigureCanvas(fig)  # Matplotlib figürünü bir Qt widget'ına dönüştür
        canvas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # Boyut politikasını ayarla
        self.monthly_chart_layout.addWidget(canvas, stretch=1)  # Layout'a ekle
        canvas.draw()  # Tuvali çiz

        self.current_monthly_chart_figure = fig  # Mevcut figürü sakla
        self.btn_save_monthly_chart.setEnabled(True)  # Kaydet butonunu etkinleştir
        # Sayfa etiketini ve navigasyon butonlarını güncelle
        self.update_monthly_page_label(graph_mode=self.current_graph_mode)
        self.update_monthly_navigation_buttons(graph_mode=self.current_graph_mode)

    def update_monthly_page_label(self, graph_mode: str) -> None:
        """
        Aylık grafik sayfa etiketini günceller.

        Args:
            graph_mode: "hat" veya "page" olarak grafik modunu belirtir.
        """
        total_pages = len(self.figures_data_monthly)  # Toplam sayfa sayısını al
        self.lbl_monthly_page.setText(f"Sayfa {self.current_page_monthly + 1} / {total_pages}")  # Etiketi güncelle

    def update_monthly_navigation_buttons(self, graph_mode: str) -> None:
        """
        Aylık grafik gezinme düğmelerinin etkin durumunu günceller.

        Args:
            graph_mode: "hat" veya "page" olarak grafik modunu belirtir.
        """
        total_pages = len(self.figures_data_monthly)  # Toplam sayfa sayısını al
        self.btn_prev_monthly.setEnabled(self.current_page_monthly > 0)  # Önceki butonunu etkinleştir/devre dışı bırak
        # Sonraki butonunu etkinleştir/devre dışı bırak
        self.btn_next_monthly.setEnabled(self.current_page_monthly < total_pages - 1)
        if graph_mode == "hat":
            self.btn_prev_monthly.setText("← Önceki Hat")
            self.btn_next_monthly.setText("Sonraki Hat →")
        elif graph_mode == "page":
            self.btn_prev_monthly.setText("← Önceki Sayfa")
            self.btn_next_monthly.setText("Sonraki Sayfa →")

    def prev_monthly_page(self) -> None:
        """Önceki aylık grafik sayfasına gider."""
        if self.current_page_monthly > 0:
            # Mevcut hat/sayfa değerlerini önbelleğe al
            self._cache_current_oee_values()
            self.current_page_monthly -= 1  # Sayfa indeksini azalt
            self.display_current_page_graphs_monthly()  # Mevcut sayfadaki grafiği göster

    def next_monthly_page(self) -> None:
        """Sonraki aylık grafik sayfasına gider."""
        total_pages = len(self.figures_data_monthly)  # Toplam sayfa sayısını al
        if self.current_page_monthly < total_pages - 1:
            # Mevcut hat/sayfa değerlerini önbelleğe al
            self._cache_current_oee_values()
            self.current_page_monthly += 1  # Sayfa indeksini artır
            self.display_current_page_graphs_monthly()  # Mevcut sayfadaki grafiği göster

    def _save_monthly_chart_as_image(self):
        """Aylık grafiği PNG/JPEG olarak kaydeder."""
        if self.current_monthly_chart_figure is None:
            QMessageBox.warning(self, "Kaydedilecek Grafik Yok", "Görüntülenecek bir aylık grafik bulunmamaktadır.")
            return

        current_name = "grafik"
        # Mevcut grafik adını al ve dosya adı için düzenle
        if self.figures_data_monthly and 0 <= self.current_page_monthly < len(self.figures_data_monthly):
            current_name = self.figures_data_monthly[self.current_page_monthly][0].replace(" ", "_").replace("/",
                                                                                                             "-")

        graph_type_name = self.cmb_monthly_graph_type.currentText().replace(" ", "_").replace("/", "-")
        default_filename = f"{graph_type_name}_{current_name}.png"  # Varsayılan dosya adı

        # Dosya kaydetme iletişim kutusunu aç
        filepath, _ = QFileDialog.getSaveFileName(
            self, "Aylık Grafiği Kaydet", default_filename, "PNG (*.png);;JPEG (*.jpeg);;JPG (*.jpg)"
        )

        if filepath:
            try:
                # Figürü belirtilen yola kaydet
                self.current_monthly_chart_figure.savefig(filepath, dpi=120, bbox_inches='tight',
                                                          facecolor=self.current_monthly_chart_figure.get_facecolor())
                QMessageBox.information(self, "Kaydedildi", f"Aylık grafik başarıyla kaydedildi: {Path(filepath).name}")
                logging.info("Aylık grafik kaydedildi: %s", filepath)  # Loglama
            except Exception as e:
                QMessageBox.critical(self, "Kaydetme Hatası", f"Aylık grafik kaydedilirken bir hata oluştu: {e}")
                logging.exception("Aylık grafik kaydetme hatası.")  # Hata loglama