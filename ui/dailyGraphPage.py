import logging
from pathlib import Path
from typing import List, Tuple
import pandas as pd

import matplotlib
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QWidget,
    QFileDialog,
    QPushButton,
    QLabel,
    QVBoxLayout,
    QHBoxLayout,
    QComboBox,
    QMessageBox,
    QProgressBar,
    QScrollArea
)

from utils.helpers import GRAPHS_PER_PAGE
from logic.graphWorker import GraphWorker
from logic.graphPlotter import GraphPlotter


class DailyGraphsPage(QWidget):
    """Günlük grafiklerin oluşturulup görüntülendiği, sayfa bazlı navigasyon ve grafik kaydetme
    özelliklerinin sunulduğu PyQt5 widget sayfası.

    Attributes:
        main_window: Ana pencere referansı, genel veri ve ayarların erişimi için.
        worker: Arka planda grafik verilerini hazırlayan iş parçacığı (GraphWorker).
        figures_data: Oluşturulan grafiklerin tuple listesi (etiket, Matplotlib Figure, OEE değeri).
        current_page: Şu anda görüntülenen sayfa indeksi (0 tabanlı).
        current_graph_type: Kullanıcının seçtiği grafik türü ("Donut" veya "Bar").
    """

    def __init__(self, main_window: "MainWindow") -> None:
        """Sayfa widget'ını başlatır ve kullanıcı arayüzünü kurar."""
        super().__init__()
        self.main_window = main_window
        self.worker: GraphWorker | None = None
        self.figures_data: List[Tuple[str, Figure, str]] = []
        self.current_page = 0
        self.current_graph_type = "Donut"
        self.init_ui()

    def init_ui(self):
        """Kullanıcı arayüzü bileşenlerini oluşturur ve sayfa düzenini ayarlar."""
        main_layout = QVBoxLayout(self)
        main_layout.setAlignment(Qt.AlignTop)

        # Sayfa başlığı
        title_label = QLabel("<h2>Günlük Grafikler</h2>")
        title_label.setObjectName("title_label")  # Stil için objectName atandı
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # İlerleme çubuğu (grafik oluşturulurken görünür)
        self.progress = QProgressBar()
        self.progress.setAlignment(Qt.AlignCenter)
        self.progress.setTextVisible(True)
        self.progress.hide()  # Başlangıçta gizli
        main_layout.addWidget(self.progress)

        # Üst navigasyon ve kontrol elemanları (geri butonu, grafik tipi seçici, sayfa bilgisi, kaydetme butonu)
        nav_top = QHBoxLayout()
        self.btn_back = QPushButton("← Veri Seçimi")
        self.btn_back.clicked.connect(lambda: self.main_window.goto_page(1))  # Ana pencereye geri döner
        nav_top.addWidget(self.btn_back)

        self.lbl_chart_info = QLabel("Grafikler oluşturuluyor...")  # Başlangıç mesajı
        self.lbl_chart_info.setAlignment(Qt.AlignCenter)
        self.lbl_chart_info.setStyleSheet("font-weight: bold; font-size: 14pt; color: #34495e;")  # Stil
        nav_top.addWidget(self.lbl_chart_info)

        self.cmb_graph_type = QComboBox()
        self.cmb_graph_type.addItems(["Donut", "Bar"])
        self.cmb_graph_type.setCurrentText(self.current_graph_type)
        self.cmb_graph_type.currentIndexChanged.connect(self.on_graph_type_changed)  # Grafik tipi değişince çağrılır
        nav_top.addWidget(self.cmb_graph_type)

        nav_top.addStretch(1)  # Boşluk bırakma için esneme
        self.lbl_page = QLabel("Sayfa 0 / 0")  # Sayfa numarası gösterimi
        self.lbl_page.setAlignment(Qt.AlignCenter)
        nav_top.addWidget(self.lbl_page)
        nav_top.addStretch(1)  # Boşluk bırakma

        self.btn_save_image = QPushButton("Grafiği Kaydet (PNG/JPEG)")
        self.btn_save_image.clicked.connect(self.save_single_graph_as_image)
        self.btn_save_image.setEnabled(False)  # Başlangıçta pasif
        nav_top.addWidget(self.btn_save_image)
        main_layout.addLayout(nav_top)

        # Grafiklerin gösterileceği kaydırılabilir alan
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.canvas_holder = QWidget()
        self.canvas_centered_layout = QHBoxLayout(self.canvas_holder)

        # Grafiklerin dikey olarak yerleştirileceği düzen
        self.vbox_canvases = QVBoxLayout()
        self.canvas_centered_layout.addStretch(1)  # Ortalamak için esneme
        self.canvas_centered_layout.addLayout(self.vbox_canvases)
        self.canvas_centered_layout.addStretch(1)  # Ortalamak için esneme

        self.scroll.setWidget(self.canvas_holder)
        main_layout.addWidget(self.scroll)

        # Alt navigasyon: önceki ve sonraki sayfa düğmeleri
        nav_bottom = QHBoxLayout()
        nav_bottom.addStretch(1)
        self.btn_prev = QPushButton("← Önceki Sayfa")
        self.btn_prev.clicked.connect(self.prev_page)
        self.btn_prev.setEnabled(False)  # Başlangıçta pasif
        nav_bottom.addWidget(self.btn_prev)

        self.btn_next = QPushButton("Sonraki Sayfa →")
        self.btn_next.clicked.connect(self.next_page)
        self.btn_next.setEnabled(False)  # Başlangıçta pasif
        nav_bottom.addWidget(self.btn_next)
        nav_bottom.addStretch(1)
        main_layout.addLayout(nav_bottom)

    def on_graph_type_changed(self, index: int) -> None:
        """Kullanıcı grafik tipini değiştirdiğinde çağrılır.

        Args:
            index: Seçilen combo box indeksi.
        """
        self.current_graph_type = self.cmb_graph_type.currentText()
        self.enter_page()  # Sayfayı yeniden yükleyerek grafikleri güncelle

    def on_results(self, results: List[Tuple[str, pd.Series, str]]) -> None:
        """GraphWorker'dan grafik verisi geldiğinde işleme ve grafik oluşturma.

        Args:
            results: Tuple listesi, her tuple (gruplama değeri, metrik verileri serisi, OEE değeri).
        """
        self.progress.setValue(100)
        self.progress.hide()

        if not results:
            QMessageBox.information(self, "Veri Yok", "Grafik oluşturulamadı. Seçilen kriterlere göre veri bulunamadı.")
            self.btn_save_image.setEnabled(False)
            self.lbl_chart_info.setText("Gösterilecek grafik bulunmadı.")
            return

        self.figures_data.clear()  # Önceki grafikleri temizle

        # Grafik boyutları (700x460 piksel, DPI 100 varsayılır, inç cinsinden)
        fig_width_inches = 700 / 100
        fig_height_inches = 460 / 100

        for grouped_val, metric_sums, oee_display_value in results:
            # Yeni Matplotlib figürü ve ekseni oluştur
            fig, ax = plt.subplots(figsize=(fig_width_inches, fig_height_inches))
            background_color = 'white'
            fig.patch.set_facecolor(background_color)
            ax.set_facecolor(background_color)

            # Metrikleri azalan sırada sıralar
            sorted_metrics_series = metric_sums.sort_values(ascending=False) if not metric_sums.empty else pd.Series()

            num_metrics = len(sorted_metrics_series)
            if num_metrics == 1 and sorted_metrics_series.index[0] == 'HAT ÇALIŞMADI':
                chart_colors = ['#FF9841']  # Özel durum renk
            else:
                # Matplotlib renk paletinden renkler alır
                colors_palette = matplotlib.colormaps.get_cmap('tab20')
                chart_colors = [colors_palette(i % 20) for i in range(num_metrics)] if num_metrics > 0 else []

            # Seçilen grafik türüne göre çizim yapar
            if self.current_graph_type == "Donut":
                GraphPlotter.create_donut_chart(ax, sorted_metrics_series, oee_display_value, chart_colors, fig)
            elif self.current_graph_type == "Bar":
                GraphPlotter.create_bar_chart(ax, sorted_metrics_series, oee_display_value, chart_colors)

            # Toplam duruş süresini saat ve dakika cinsinden hesapla
            total_duration_seconds = sorted_metrics_series.sum()
            total_duration_hours = int(total_duration_seconds // 3600)
            total_duration_minutes = int((total_duration_seconds % 3600) // 60)
            total_duration_text = f"TOPLAM DURUŞ\n{total_duration_hours} SAAT {total_duration_minutes} DAKİKA"

            # Grafik altına toplam duruş metnini ekle
            fig.text(0.01, 0.05, total_duration_text, transform=fig.transFigure,
                     fontsize=14, fontweight='bold', verticalalignment='bottom')

            self.figures_data.append((grouped_val, fig, oee_display_value))
            plt.close(fig)  # Bellek sızıntısını önlemek için figürü kapat

        self.display_current_page_graphs()  # Oluşturulan grafikleri göster
        if self.figures_data:
            self.btn_save_image.setEnabled(True)  # Kaydetme butonunu aktif et

    def enter_page(self) -> None:
        """Sayfaya girildiğinde grafik oluşturma sürecini başlatır."""
        self.figures_data.clear()  # Önceki verileri temizle
        self.clear_canvases()  # Önceki grafik tuvalini temizle
        self.progress.setValue(0)
        self.progress.show()  # İlerleme çubuğunu göster
        self.btn_save_image.setEnabled(False)
        self.lbl_chart_info.setText("Grafikler oluşturuluyor...")
        self.update_page_label()
        self.update_navigation_buttons()

        # Önceki çalışan iş parçacığı varsa durdur
        if self.worker and self.worker.isRunning():
            self.worker.quit()
            self.worker.wait()

        # Yeni GraphWorker nesnesi oluştur ve başlat
        self.worker = GraphWorker(
            df=self.main_window.df,
            grouping_col_name=self.main_window.grouping_col_name,
            grouped_col_name=self.main_window.grouped_col_name,
            grouped_values=self.main_window.grouped_values,
            metric_cols=self.main_window.selected_metrics,
            oee_col_name=self.main_window.oee_col_name,
            selected_grouping_val=self.main_window.selected_grouping_val
        )
        # Sinyalleri slotlara bağla
        self.worker.finished.connect(self.on_results)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_error(self, message: str) -> None:
        """GraphWorker'dan hata mesajı geldiğinde kullanıcıya gösterir.

        Args:
            message: Hata mesajı metni.
        """
        QMessageBox.critical(self, "Hata", message)
        self.progress.setValue(0)
        self.progress.hide()
        self.lbl_chart_info.setText("Grafik oluşturma hatası.")
        self.btn_save_image.setEnabled(False)

    def clear_canvases(self) -> None:
        """Mevcut grafik tuval widgetlarını temizler ve bellekten kaldırır."""
        while self.vbox_canvases.count():
            item = self.vbox_canvases.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()  # Widget'ı bellekten sil

    def display_current_page_graphs(self) -> None:
        """Geçerli sayfadaki grafiklere ait figürleri tuval üzerinde gösterir."""
        self.clear_canvases()

        # Toplam sayfa sayısını hesaplar
        total_pages = (len(self.figures_data) + GRAPHS_PER_PAGE - 1) // GRAPHS_PER_PAGE if self.figures_data else 0

        # Sayfa indeksi sınırlandırması
        if self.current_page >= total_pages and total_pages > 0:
            self.current_page = total_pages - 1
        elif total_pages == 0:
            self.current_page = 0

        # Gösterilecek grafiklerin indeks aralığı
        start_index = self.current_page * GRAPHS_PER_PAGE
        end_index = start_index + GRAPHS_PER_PAGE

        graphs_to_display = self.figures_data[start_index:end_index]

        if not graphs_to_display:
            self.lbl_chart_info.setText("Gösterilecek grafik bulunamadı.")
            self.btn_save_image.setEnabled(False)
            self.update_page_label()
            self.update_navigation_buttons()
            return

        for grouped_val, fig, oee_display_value in graphs_to_display:
            canvas = FigureCanvas(fig)
            canvas.setFixedSize(700, 460)  # Sabit boyutlu grafik tuvali
            self.vbox_canvases.addWidget(canvas)
            display_grouped_val = grouped_val.replace("HAT-#", "").strip()
            self.lbl_chart_info.setText(f"{self.main_window.selected_grouping_val} - {display_grouped_val}")

        self.update_page_label()
        self.update_navigation_buttons()
        self.btn_save_image.setEnabled(True)

    def update_page_label(self) -> None:
        """Sayfa numarası etiketini günceller (örn. 'Sayfa 1 / 5')."""
        total_pages = (len(self.figures_data) + GRAPHS_PER_PAGE - 1) // GRAPHS_PER_PAGE if self.figures_data else 0
        self.lbl_page.setText(f"Sayfa {self.current_page + 1} / {total_pages}")

    def update_navigation_buttons(self) -> None:
        """Önceki ve Sonraki sayfa düğmelerinin etkinlik durumlarını günceller."""
        total_pages = (len(self.figures_data) + GRAPHS_PER_PAGE - 1) // GRAPHS_PER_PAGE if self.figures_data else 0
        self.btn_prev.setEnabled(self.current_page > 0)
        self.btn_next.setEnabled(self.current_page < total_pages - 1)

    def prev_page(self) -> None:
        """Önceki sayfa grafiklerini görüntüler, sayfa numarasını azaltır."""
        if self.current_page > 0:
            self.current_page -= 1
            self.display_current_page_graphs()

    def next_page(self) -> None:
        """Sonraki sayfa grafiklerini görüntüler, sayfa numarasını artırır."""
        total_pages = (len(self.figures_data) + GRAPHS_PER_PAGE - 1) // GRAPHS_PER_PAGE if self.figures_data else 0
        if self.current_page < total_pages - 1:
            self.current_page += 1
            self.display_current_page_graphs()

    def save_single_graph_as_image(self) -> None:
        """Mevcut sayfadaki ilk grafiği kullanıcıya seçtirilen dosya adıyla PNG/JPEG olarak kaydeder."""
        if not self.figures_data:
            QMessageBox.warning(self, "Kaydedilecek Grafik Yok", "Görüntülenecek bir grafik bulunmamaktadır.")
            return

        total_graphs = len(self.figures_data)
        fig_index_on_page = self.current_page * GRAPHS_PER_PAGE

        if not (0 <= fig_index_on_page < total_graphs):
            QMessageBox.warning(self, "Geçersiz Sayfa", "Mevcut sayfada kaydedilecek bir grafik yok.")
            return

        grouped_val, fig, _ = self.figures_data[fig_index_on_page]

        # Varsayılan dosya adını düzenle (boşluk ve / karakterleri kaldırılır)
        default_filename = f"grafik_{grouped_val}_{self.main_window.selected_grouping_val}_{self.current_graph_type}.png".replace(
            " ", "_").replace("/", "-")
        filepath, _ = QFileDialog.getSaveFileName(
            self, "Grafiği Kaydet", default_filename, "PNG (*.png);;JPEG (*.jpeg);;JPG (*.jpg)"
        )

        if filepath:
            try:
                # Grafik dosyasını kaydet
                fig.savefig(filepath, dpi=plt.rcParams['savefig.dpi'], bbox_inches='tight',
                            facecolor=fig.get_facecolor())
                QMessageBox.information(self, "Kaydedildi", f"Grafik başarıyla kaydedildi: {Path(filepath).name}")
                logging.info("Grafik kaydedildi: %s", filepath)
            except Exception as e:
                QMessageBox.critical(self, "Kaydetme Hatası", f"Grafik kaydedilirken bir hata oluştu: {e}")
                logging.exception("Grafik kaydetme hatası.")
