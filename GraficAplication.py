import sys
import os
import random
import numpy as np

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QAction, QMenu, QInputDialog,
    QMessageBox, QFileDialog, QLabel, QVBoxLayout, QWidget, QScrollArea, QProgressDialog
)
from PyQt5.QtGui import QIcon, QFont, QPixmap
from PyQt5.QtCore import Qt, QSize, QTimer, QDateTime

import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar

from fpdf import FPDF
from fpdf.enums import XPos, YPos

# Matplotlib varsayılan arka planını daha koyu bir temaya uyduralım
plt.style.use('dark_background')
plt.rcParams.update({
    'axes.facecolor': '#282828',
    'axes.edgecolor': '#888888',
    'axes.labelcolor': '#E0E0E0',
    'xtick.color': '#E0E0E0',
    'ytick.color': '#E0E0E0',
    'grid.color': '#444444',
    'text.color': '#E0E0E0',
    'figure.facecolor': '#282828',
    'savefig.facecolor': '#282828'
})


class MplWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.figure = plt.figure()
        self.canvas = FigureCanvas(self.figure)
        self.toolbar = NavigationToolbar(self.canvas, self)

        self.layout = QVBoxLayout()
        self.layout.addWidget(self.toolbar)
        self.layout.addWidget(self.canvas)
        self.setLayout(self.layout)

        self.figure.clear()
        self.canvas.draw_idle()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Grafik Uygulaması")
        self.setGeometry(100, 100, 1024, 768)

        self.current_plot_data = []
        self.ai_pipeline = None

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        self.file_info_label = QLabel("Lütfen bir dosya seçin...", self)
        self.file_info_label.setAlignment(Qt.AlignCenter)
        self.file_info_label.setFont(QFont("Arial", 12))
        self.file_info_label.setObjectName("fileInfoLabel")
        self.file_info_label.setFixedHeight(30)

        self.plot_info_label = QLabel("Henüz bir grafik oluşturulmadı.", self)
        self.plot_info_label.setAlignment(Qt.AlignCenter)
        self.plot_info_label.setFont(QFont("Arial", 12))
        self.plot_info_label.setObjectName("plotInfoLabel")
        self.plot_info_label.setFixedHeight(30)

        self.main_layout.addWidget(self.file_info_label)
        self.main_layout.addWidget(self.plot_info_label)

        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)

        self.scroll_content_widget = QWidget()
        self.scroll_content_layout = QVBoxLayout(self.scroll_content_widget)

        self.matplotlib_widget = MplWidget(self)
        self.scroll_content_layout.addWidget(self.matplotlib_widget)

        self.scroll_area.setWidget(self.scroll_content_widget)

        self.main_layout.addWidget(self.scroll_area)

        self.create_menu()

        self.setStyleSheet("""
            QMainWindow {
                background-color: #202020;
            }
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
            QLabel#plotInfoLabel {
                color: #A0FFA0;
                font-weight: bold;
                padding: 5px;
                background-color: #3A3A3A;
                border-bottom: 1px solid #555;
            }
            QScrollArea {
                border: none;
                background-color: #282828;
            }
            QScrollBar:vertical {
                border: 1px solid #555;
                background: #333;
                width: 15px;
                margin: 15px 0 15px 0;
            }
            QScrollBar::handle:vertical {
                background: #666;
                min-height: 20px;
                border-radius: 5px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                background: none;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }
            QScrollBar:horizontal {
                border: 1px solid #555;
                background: #333;
                height: 15px;
                margin: 0 15px 0 15px;
            }
            QScrollBar::handle:horizontal {
                background: #666;
                min-width: 20px;
                border-radius: 5px;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                background: none;
            }
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: none;
            }
        """)

    def create_menu(self):
        menubar = self.menuBar()

        def create_action_with_icon(text, shortcut, status_tip, icon_path, connect_func):
            try:
                action = QAction(QIcon(icon_path), text, self)
            except Exception as e:
                print(f"Uyarı: İkon yüklenirken hata oluştu ({icon_path}): {e}")
                action = QAction(text, self)
            action.setShortcut(shortcut)
            action.setStatusTip(status_tip)
            action.triggered.connect(connect_func)
            return action

        # 1. Dosya Menüsü
        file_menu = menubar.addMenu("&Dosya")
        file_menu.addAction(create_action_with_icon(
            "Word Dosyası &Aç...", "Ctrl+W", "Bir Word belgesini açar",
            'icons/word.png', lambda: self.open_file("Word Dosyaları (*.docx *.doc);;Tüm Dosyalar (*)", "word")
        ))
        file_menu.addAction(create_action_with_icon(
            "Excel Dosyası &Aç...", "Ctrl+E", "Bir Excel çalışma sayfasını açar",
            'icons/excel.png', lambda: self.open_file("Excel Dosyaları (*.xlsx *.xls);;Tüm Dosyalar (*)", "excel")
        ))
        file_menu.addAction(create_action_with_icon(
            "PPTX Dosyası &Aç...", "Ctrl+P", "Bir PowerPoint sunumunu açar",
            'icons/pptx.png', lambda: self.open_file("PowerPoint Dosyaları (*.pptx *.ppt);;Tüm Dosyalar (*)", "pptx")
        ))
        file_menu.addSeparator()
        file_menu.addAction(create_action_with_icon(
            "Çı&kış", "Ctrl+Q", "Uygulamadan çıkar",
            'icons/exit.png', self.close
        ))

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
            save_action = create_action_with_icon(
                name, "", f"Grafiği .{file_ext} formatında kaydet",
                f'icons/save_{file_ext}.png', lambda checked, fmt=file_ext: self.save_graph(fmt)
            )
            save_as_menu.addAction(save_action)
        download_print_menu.addMenu(save_as_menu)
        download_print_menu.addSeparator()
        download_print_menu.addAction(create_action_with_icon(
            "&Yazdır...", "Ctrl+P", "Mevcut grafiği yazdır",
            'icons/print.png', self.print_graph
        ))

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

        # 5. Rapor Menüsü (Sadece PDF Raporu)
        report_menu = menubar.addMenu("&Rapor")
        report_action = create_action_with_icon(
            "&Grafik Raporu Oluştur (PDF)...", "",
            "Oluşturulan grafiklere ve verilere dayalı detaylı bir PDF raporu oluşturur",
            'icons/report.png', self.generate_pdf_report
        )
        report_menu.addAction(report_action)

    def get_plot_count(self, chart_type):
        try:
            num, ok = QInputDialog.getInt(self, "Grafik Adedi Girin",
                                          f"Kaç adet '{chart_type}' grafiği oluşturmak istersiniz?",
                                          min=1, max=10, step=1)

            if ok:
                self.draw_graph(chart_type, num)
                self.update_plot_info_label(chart_type, num)
                QMessageBox.information(self, "Grafik Oluşturuldu",
                                        f"'{chart_type}' türünde {num} adet grafik başarıyla oluşturuldu.")
            else:
                QMessageBox.information(self, "İptal Edildi", "Grafik oluşturma işlemi iptal edildi.")
                self.update_plot_info_label("", 0)
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Grafik adedi alınırken bir hata oluştu: {e}")
            self.update_plot_info_label("", 0)

    def save_graph(self, file_format):
        if not self.matplotlib_widget.figure.axes:
            QMessageBox.warning(self, "Uyarı", "Kaydedilecek bir grafik bulunamadı. Lütfen önce bir grafik oluşturun.")
            return

        try:
            file_name, _ = QFileDialog.getSaveFileName(self, "Grafiği Kaydet", "grafik",
                                                       f"Grafik Dosyaları (*.{file_format});;Tüm Dosyalar (*)")

            if file_name:
                self.matplotlib_widget.figure.tight_layout()
                self.matplotlib_widget.figure.savefig(file_name, format=file_format)
                QMessageBox.information(self, "Kaydetme Başarılı",
                                        f"Grafik '{file_name}' olarak kaydedildi.")
            else:
                QMessageBox.information(self, "Kaydetme İptal Edildi", "Kaydetme işlemi iptal edildi.")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Grafik kaydedilirken bir hata oluştu: {e}\n"
                                               f"Lütfen dosya yolunun geçerli olduğundan ve yazma izniniz olduğundan emin olun.")

    def print_graph(self):
        QMessageBox.information(self, "Yazdırma İşlemi",
                                "Grafik yazdırma işlemi başlatılacak. (Henüz tam işlevsel değil)")

    def open_file(self, file_filter, file_type_code):
        try:
            file_name, _ = QFileDialog.getOpenFileName(self, "Dosya Seç", "", file_filter)

            if file_name:
                self.update_file_info_label(file_name, file_type_code)
                self.update_plot_info_label("", 0)
                QMessageBox.information(self, "Dosya Seçimi",
                                        f"Seçilen Dosya: {file_name}")
            else:
                self.update_file_info_label("", "")
                self.update_plot_info_label("", 0)
                QMessageBox.information(self, "İptal Edildi", "Dosya seçimi iptal edildi.")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Dosya seçimi sırasında bir hata oluştu: {e}")
            self.update_file_info_label("", "")
            self.update_plot_info_label("", 0)

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
                self.file_info_label.setText(f"  {base_name} (İkon yüklenemedi)")
                print(f"Uyarı: '{icon_path}' ikonu bulunamadı veya yüklenemedi.")
        else:
            self.file_info_label.setPixmap(QPixmap())
            self.file_info_label.setText("Lütfen bir dosya seçin...")

        self.file_info_label.adjustSize()
        self.file_info_label.setFixedWidth(self.width())

    def update_plot_info_label(self, chart_type_text, count):
        if chart_type_text and count > 0:
            self.plot_info_label.setText(f"  Grafik: {chart_type_text.split('(')[0].strip()} ({count} adet)")
            self.plot_info_label.setContentsMargins(5, 0, 0, 0)
        else:
            self.plot_info_label.setText("Henüz bir grafik oluşturulmadı.")
            self.plot_info_label.setContentsMargins(0, 0, 0, 0)

        self.plot_info_label.adjustSize()
        self.plot_info_label.setFixedWidth(self.width())

    def draw_graph(self, chart_type_text, count):
        self.matplotlib_widget.figure.clear()
        self.current_plot_data = []  # Her yeni çizimde mevcut verileri temizle

        figure_height_per_row = 5

        rows = 1
        cols = 1
        if count == 1:
            rows, cols = 1, 1
        elif count == 2:
            rows, cols = 1, 2
        elif count == 3:
            rows, cols = 1, 3
        elif count == 4:
            rows, cols = 2, 2
        elif count == 5 or count == 6:
            rows, cols = 2, 3
        elif count > 6:
            cols = 3
            rows = (count + cols - 1) // cols

        self.matplotlib_widget.figure.set_size_inches(cols * 4, rows * figure_height_per_row)

        chart_func_name = chart_type_text.split('(')[-1][:-1] if '(' in chart_type_text else None

        if chart_func_name is None:
            QMessageBox.warning(self, "Hata", "Geçersiz grafik türü adı.")
            return

        for i in range(count):
            try:
                ax = self.matplotlib_widget.figure.add_subplot(rows, cols, i + 1)
                title = f"{chart_type_text.split('(')[0].strip()} {i + 1}"
                ax.set_title(title)

                plot_info = {'type': chart_func_name, 'title': title, 'xlabel': 'N/A',
                             'ylabel': 'N/A'}  # Default values for labels

                x = np.linspace(0, 10, 100)
                y = np.sin(x + i * 0.5) + random.uniform(-0.5, 0.5)

                if chart_func_name == "plot":
                    ax.plot(x, y, label=f'Series {i + 1}')
                    ax.set_xlabel("X Ekseni")
                    ax.set_ylabel("Y Ekseni")
                    ax.legend()
                    plot_info['data_x'] = x.tolist()
                    plot_info['data_y'] = y.tolist()
                    plot_info['xlabel'] = "X Ekseni"
                    plot_info['ylabel'] = "Y Ekseni"
                elif chart_func_name == "bar":
                    categories = [f'Kategori {k + 1}' for k in range(5)]
                    values = np.random.randint(5, 20, 5)
                    colors = [plt.cm.viridis(k / (len(categories) - 1)) for k in range(len(categories))]
                    ax.bar(categories, values, color=colors)
                    ax.set_xlabel("Kategoriler")
                    ax.set_ylabel("Değerler")
                    plot_info['categories'] = categories
                    plot_info['values'] = values.tolist()
                    plot_info['xlabel'] = "Kategoriler"
                    plot_info['ylabel'] = "Değerler"
                elif chart_func_name == "hist":
                    data = np.random.randn(1000)
                    ax.hist(data, bins=30, color='skyblue', edgecolor='black')
                    ax.set_xlabel("Değer")
                    ax.set_ylabel("Frekans")
                    plot_info['data'] = data.tolist()
                    plot_info['bins'] = 30
                    plot_info['xlabel'] = "Değer"
                    plot_info['ylabel'] = "Frekans"
                elif chart_func_name == "pie":
                    sizes = [random.randint(10, 30) for _ in range(4)]
                    labels = [f'Dilim {k + 1}' for k in range(4)]
                    ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90)
                    ax.axis('equal')
                    plot_info['sizes'] = sizes
                    plot_info['labels'] = labels
                    plot_info['xlabel'] = "N/A"  # Pasta grafiğinde X/Y ekseni anlamlı değil
                    plot_info['ylabel'] = "N/A"
                elif chart_func_name == "scatter":
                    x_scatter = np.random.rand(50) * 10
                    y_scatter = np.random.rand(50) * 10
                    colors = np.random.rand(50)
                    sizes = np.random.rand(50) * 200 + 20
                    ax.scatter(x_scatter, y_scatter, c=colors, s=sizes, alpha=0.7, cmap='viridis')
                    ax.set_xlabel("X Verisi")
                    ax.set_ylabel("Y Verisi")
                    plot_info['data_x'] = x_scatter.tolist()
                    plot_info['data_y'] = y_scatter.tolist()
                    plot_info['colors'] = colors.tolist()
                    plot_info['sizes'] = sizes.tolist()
                    plot_info['xlabel'] = "X Verisi"
                    plot_info['ylabel'] = "Y Verisi"
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
                    plot_info['data_x'] = x_area.tolist()
                    plot_info['data_y1'] = y1_area.tolist()
                    plot_info['data_y2'] = y2_area.tolist()
                    plot_info['xlabel'] = "X Ekseni"
                    plot_info['ylabel'] = "Değer"
                elif chart_func_name == "boxplot":
                    data = [np.random.normal(0, std, 100) for std in range(1, 4)]
                    ax.boxplot(data, patch_artist=True)
                    labels = [f'Grup {k + 1}' for k in range(len(data))]
                    ax.set_xticklabels(labels)
                    ax.set_ylabel("Değer")
                    plot_info['data_groups'] = [d.tolist() for d in data]
                    plot_info['labels'] = labels
                    plot_info['xlabel'] = "Gruplar"
                    plot_info['ylabel'] = "Değer"
                elif chart_func_name == "violinplot":
                    data = [np.random.normal(0, std, 100) for std in range(1, 4)]
                    ax.violinplot(data, showmeans=True, showmedians=True)
                    labels = [f'Grup {k + 1}' for k in range(len(data))]
                    ax.set_xticks(np.arange(1, len(data) + 1))
                    ax.set_xticklabels(labels)
                    ax.set_ylabel("Değer")
                    plot_info['data_groups'] = [d.tolist() for d in data]
                    plot_info['labels'] = labels
                    plot_info['xlabel'] = "Gruplar"
                    plot_info['ylabel'] = "Değer"
                elif chart_func_name == "stem":
                    x_stem = np.arange(10)
                    y_stem = np.random.randint(1, 10, 10)
                    ax.stem(x_stem, y_stem)
                    ax.set_xlabel("Dizin")
                    ax.set_ylabel("Değer")
                    plot_info['data_x'] = x_stem.tolist()
                    plot_info['data_y'] = y_stem.tolist()
                    plot_info['xlabel'] = "Dizin"
                    plot_info['ylabel'] = "Değer"
                elif chart_func_name == "errorbar":
                    x_err = np.linspace(0, 10, 10)
                    y_err = np.sin(x_err)
                    y_error = 0.1 + 0.2 * np.random.rand(len(x_err))
                    ax.errorbar(x_err, y_err, yerr=y_error, fmt='-o')
                    ax.set_xlabel("X Ekseni")
                    ax.set_ylabel("Y Ekseni")
                    plot_info['data_x'] = x_err.tolist()
                    plot_info['data_y'] = y_err.tolist()
                    plot_info['y_error'] = y_error.tolist()
                    plot_info['xlabel'] = "X Ekseni"
                    plot_info['ylabel'] = "Y Ekseni"
                else:
                    ax.text(0.5, 0.5, f"'{chart_type_text}' için çizim kodu yok.",
                            horizontalalignment='center', verticalalignment='center', transform=ax.transAxes)
                    ax.set_title("Bilinmeyen Grafik Türü")
                    plot_info = {'type': 'unknown', 'title': title, 'error': f"Çizim kodu yok: {chart_type_text}"}

                self.current_plot_data.append(plot_info)

            except Exception as e:
                print(f"Hata: {i + 1}. grafiği çizerken bir sorun oluştu: {e}")
                ax.text(0.5, 0.5, f"Grafik çiziminde hata:\n{e}",
                        horizontalalignment='center', verticalalignment='center',
                        transform=ax.transAxes, color='red', fontsize=12)
                ax.set_title(f"Hata Oluştu ({chart_type_text.split('(')[0].strip()} {i + 1})")
                self.current_plot_data.append(
                    {'type': chart_func_name, 'title': title, 'error': str(e), 'xlabel': 'Hata', 'ylabel': 'Hata'})

        self.matplotlib_widget.figure.tight_layout()
        self.matplotlib_widget.canvas.draw_idle()
        self.scroll_content_widget.setMinimumSize(
            int(self.matplotlib_widget.figure.get_size_inches()[0] * self.matplotlib_widget.figure.dpi),
            int(self.matplotlib_widget.figure.get_size_inches()[1] * self.matplotlib_widget.figure.dpi)
        )

    def generate_report_content_for_pdf(self):
        """
        Grafik verilerine dayalı, yapay zeka kullanmadan detaylı ve karşılaştırmalı bir rapor içeriği oluşturur.
        PDF için özel formatlama gerektiren metni döndürür.
        """
        if not self.current_plot_data:
            return ["Raporlanacak grafik verisi bulunamadı. Lütfen önce bir grafik oluşturun."]

        report_lines = []

        report_lines.append("## A. Bireysel Grafik Detayları\n")
        plot_type_counts = {}
        all_x_data = []
        all_y_data = []

        for i, plot_info in enumerate(self.current_plot_data):
            plot_type = plot_info.get('type', 'Bilinmiyor').replace('_', ' ').capitalize()
            title = plot_info.get('title', 'Başlıksız')
            xlabel = plot_info.get('xlabel', 'Belirtilmedi')
            ylabel = plot_info.get('ylabel', 'Belirtilmedi')

            report_lines.append(f"### {i + 1}. Grafik Detayları: {title} ({plot_type})\n")
            report_lines.append(f"  Grafik Türü: {plot_type}")
            report_lines.append(f"  Başlık: {title}")
            report_lines.append(f"  X Ekseni Etiketi: {xlabel}")
            report_lines.append(f"  Y Ekseni Etiketi: {ylabel}")

            plot_type_counts[plot_type] = plot_type_counts.get(plot_type, 0) + 1

            if 'error' in plot_info:
                report_lines.append(f"  Hata: Bu grafik çizilirken bir sorun oluştu: {plot_info['error']}\n")
            else:
                if plot_info['type'] in ['plot', 'scatter', 'stem', 'errorbar']:
                    x_data = plot_info.get('data_x', [])
                    y_data = plot_info.get('data_y', [])
                    report_lines.append(
                        f"  X Verisi (ilk 5): {[f'{val:.2f}' for val in x_data[:min(5, len(x_data))]]}...")
                    report_lines.append(
                        f"  Y Verisi (ilk 5): {[f'{val:.2f}' for val in y_data[:min(5, len(y_data))]]}...")
                    if x_data:
                        report_lines.append(f"  X Verisi Aralığı: [{min(x_data):.2f}, {max(x_data):.2f}]")
                    if y_data:
                        report_lines.append(f"  Y Verisi Aralığı: [{min(y_data):.2f}, {max(y_data):.2f}]")
                    all_x_data.extend(x_data)
                    all_y_data.extend(y_data)

                elif plot_info['type'] == 'bar':
                    categories = plot_info.get('categories', [])
                    values = plot_info.get('values', [])
                    report_lines.append(f"  Kategoriler: {', '.join(categories)}")
                    report_lines.append(f"  Değerler: {values}")
                    if values:
                        report_lines.append(
                            f"  Min Değer: {min(values)}, Max Değer: {max(values)}, Ortalama: {np.mean(values):.2f}")

                elif plot_info['type'] == 'hist':
                    data = plot_info.get('data', [])
                    if data:
                        report_lines.append(f"  Veri Sayısı: {len(data)}")
                        report_lines.append(f"  Min Değer: {min(data):.2f}, Max Değer: {max(data):.2f}")
                        report_lines.append(f"  Ortalama: {np.mean(data):.2f}, Medyan: {np.median(data):.2f}")
                        report_lines.append(f"  Standart Sapma: {np.std(data):.2f}")
                    else:
                        report_lines.append("  Veri bulunamadı.")

                elif plot_info['type'] == 'pie':
                    labels = plot_info.get('labels', [])
                    sizes = plot_info.get('sizes', [])
                    report_lines.append(f"  Dilim Etiketleri: {labels}")
                    report_lines.append(f"  Dilim Boyutları: {sizes}")
                    if sizes:
                        report_lines.append(f"  Toplam Boyut: {sum(sizes)}")

                elif plot_info['type'] == 'fill_between':
                    x_data = plot_info.get('data_x', [])
                    y1_data = plot_info.get('data_y1', [])
                    y2_data = plot_info.get('data_y2', [])
                    report_lines.append(
                        f"  X Verisi (ilk 5): {[f'{val:.2f}' for val in x_data[:min(5, len(x_data))]]}...")
                    report_lines.append(
                        f"  Y1 Verisi (ilk 5): {[f'{val:.2f}' for val in y1_data[:min(5, len(y1_data))]]}...")
                    report_lines.append(
                        f"  Y2 Verisi (ilk 5): {[f'{val:.2f}' for val in y2_data[:min(5, len(y2_data))]]}...")
                    if x_data:
                        report_lines.append(f"  X Verisi Aralığı: [{min(x_data):.2f}, {max(x_data):.2f}]")
                    if y1_data:
                        report_lines.append(f"  Y1 Verisi Aralığı: [{min(y1_data):.2f}, {max(y1_data):.2f}]")
                    if y2_data:
                        report_lines.append(f"  Y2 Verisi Aralığı: [{min(y2_data):.2f}, {max(y2_data):.2f}]")
                    all_x_data.extend(x_data)
                    all_y_data.extend(y1_data)
                    all_y_data.extend(y2_data)

                elif plot_info['type'] in ['boxplot', 'violinplot']:
                    data_groups = plot_info.get('data_groups', [])
                    labels = plot_info.get('labels', [])
                    report_lines.append(f"  Gruplar: {labels}")
                    for group_idx, group_data in enumerate(data_groups):
                        if group_data:
                            report_lines.append(
                                f"    Grup {labels[group_idx]} - Sayı: {len(group_data)}, Min: {min(group_data):.2f}, Max: {max(group_data):.2f}, Ort: {np.mean(group_data):.2f}")
                            all_y_data.extend(group_data)

            report_lines.append("\n")

        # Karşılaştırmalı Analiz Bölümü
        report_lines.append("-" * 60 + "\n")
        report_lines.append("## B. Grafikler Arası Karşılaştırmalı Analiz\n")

        if len(self.current_plot_data) > 1:
            report_lines.append("Oluşturulan farklı grafik tipleri ve genel veri dağılımları üzerine kıyaslamalar:\n")

            report_lines.append("### Grafik Tipi Dağılımı:")
            for plot_type, count in plot_type_counts.items():
                report_lines.append(f"- **{plot_type}**: {count} adet")
            report_lines.append("\n")

            if all_x_data:
                report_lines.append("### Tüm Grafiklerdeki X Verisi Genel İstatistikleri:")
                report_lines.append(f"- Toplam Veri Noktası: {len(all_x_data)}")
                report_lines.append(f"- Min Değer: {min(all_x_data):.2f}")
                report_lines.append(f"- Max Değer: {max(all_x_data):.2f}")
                report_lines.append(f"- Ortalama: {np.mean(all_x_data):.2f}")
                report_lines.append(f"- Medyan: {np.median(all_x_data):.2f}")
                report_lines.append(f"- Standart Sapma: {np.std(all_x_data):.2f}\n")
            else:
                report_lines.append("### Tüm Grafikler İçin X Ekseni Verisi Bulunamadı.\n")

            if all_y_data:
                report_lines.append("### Tüm Grafiklerdeki Y Verisi Genel İstatistikleri:")
                report_lines.append(f"- Toplam Veri Noktası: {len(all_y_data)}")
                report_lines.append(f"- Min Değer: {min(all_y_data):.2f}")
                report_lines.append(f"- Max Değer: {max(all_y_data):.2f}")
                report_lines.append(f"- Ortalama: {np.mean(all_y_data):.2f}")
                report_lines.append(f"- Medyan: {np.median(all_y_data):.2f}")
                report_lines.append(f"- Standart Sapma: {np.std(all_y_data):.2f}\n")
            else:
                report_lines.append("### Tüm Grafikler İçin Y Ekseni Verisi Bulunamadı.\n")

            for p_type in plot_type_counts.keys():
                if plot_type_counts[p_type] > 1:
                    report_lines.append(f"### {p_type} Tipi Grafikler Arası Gözlemler:")
                    if p_type == "Bar":
                        bar_graphs = [p for p in self.current_plot_data if p.get('type') == 'bar' and 'values' in p]
                        if bar_graphs:
                            all_bar_values = [val for p in bar_graphs for val in p['values']]
                            if all_bar_values:
                                report_lines.append(
                                    f"  Tüm Bar Grafiklerindeki Değerlerin Ortalaması: {np.mean(all_bar_values):.2f}")
                                report_lines.append(
                                    f"  Tüm Bar Grafiklerindeki Değerlerin Min/Max: [{min(all_bar_values)}, {max(all_bar_values)}]")
                    elif p_type == "Plot":
                        plot_graphs = [p for p in self.current_plot_data if p.get('type') == 'plot' and 'data_y' in p]
                        if plot_graphs:
                            all_plot_y_data = [val for p in plot_graphs for val in p['data_y']]
                            if all_plot_y_data:
                                report_lines.append(
                                    f"  Tüm Çizgi Grafiklerindeki Y Verisi Ortalaması: {np.mean(all_plot_y_data):.2f}")
                                report_lines.append(
                                    f"  Tüm Çizgi Grafiklerindeki Y Verisi Min/Max: [{min(all_plot_y_data):.2f}, {max(all_plot_y_data):.2f}]")
                    else:
                        report_lines.append(f"  Birden fazla {p_type} grafiği oluşturuldu. Detaylı karşılaştırma için "
                                            "bireysel grafik detaylarına bakınız. Genel eğilimler benzer görünmektedir.\n")
                    report_lines.append("\n")

        else:
            report_lines.append("Karşılaştırmalı analiz için birden fazla grafik bulunmamaktadır.\n")

        report_lines.append("-" * 60 + "\n")
        report_lines.append("Rapor Sonu.\n")

        return report_lines

    def generate_pdf_report(self):
        """
        Oluşturulan grafiklere dayalı, hoş görünümlü bir PDF raporu oluşturur.
        """
        if not self.current_plot_data:
            QMessageBox.warning(self, "Rapor Oluşturma", "Raporlanacak bir grafik bulunmamaktadır.")
            return

        pdf_file_name, _ = QFileDialog.getSaveFileName(self, "Grafik Raporunu Kaydet (PDF)", "grafik_raporu",
                                                       "PDF Dosyaları (*.pdf);;Tüm Dosyalar (*)")

        if not pdf_file_name:
            QMessageBox.information(self, "Rapor Oluşturma İptal Edildi", "PDF raporu oluşturma işlemi iptal edildi.")
            return

        progress_dialog = QProgressDialog("PDF Raporu Oluşturuluyor...", "İptal", 0, 0, self)
        progress_dialog.setWindowModality(Qt.WindowModal)
        progress_dialog.setMinimumDuration(0)
        progress_dialog.setValue(0)
        progress_dialog.show()
        QApplication.processEvents()

        temp_image_path = "temp_chart_full.png"  # Tüm grafiği PDF'e eklemek için geçici dosya

        try:
            pdf = FPDF()
            pdf.add_page()  # İlk sayfayı ekle

            current_dir = os.getcwd()
            font_path_regular = os.path.join(current_dir, 'NotoSans-Regular.ttf')
            font_path_bold = os.path.join(current_dir, 'NotoSans-Bold.ttf')

            font_loaded = False
            try:
                if os.path.exists(font_path_regular):
                    pdf.add_font('NotoSans', '', font_path_regular)
                if os.path.exists(font_path_bold):
                    pdf.add_font('NotoSans', 'B', font_path_bold)

                # Test the font setting
                pdf.set_font("NotoSans", "B", 20)
                font_loaded = True
            except Exception as e:
                print(f"HATA: Türkçe font yüklenirken sorun oluştu: {e}. Arial kullanılacak.")
                pdf.set_font("Arial", "B", 20)
                font_loaded = False

            # Rapor Başlığı - Üst Kısım
            header_lines = [
                "###################################################",
                "#             GRAFİK ANALİZ RAPORU                #",
                "#           (Yapay Zeka Desteksiz)                #",
                "###################################################"
            ]
            # Calculate width for centering, excluding margins
            page_width_minus_margins = pdf.w - pdf.l_margin - pdf.r_margin

            for h_line in header_lines:
                if font_loaded:
                    # Adjust font size for main title line
                    if "GRAFİK ANALİZ RAPORU" in h_line:
                        pdf.set_font("NotoSans", "B", 12)
                        line_height = 8  # Increased height for main title
                    else:
                        pdf.set_font("NotoSans", "B", 10)
                        line_height = 6  # Standard height for hash lines
                else:
                    if "GRAFİK ANALİZ RAPORU" in h_line:
                        pdf.set_font("Arial", "B", 12)
                        line_height = 8
                    else:
                        pdf.set_font("Arial", "B", 10)
                        line_height = 6

                # Using cell for each line to ensure centering
                pdf.cell(page_width_minus_margins, line_height, h_line, 0, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                         align="C")
            pdf.ln(5)  # Add extra space after header block

            # Oluşturulma Tarihi
            if font_loaded:
                pdf.set_font("NotoSans", "", 10)
            else:
                pdf.set_font("Arial", "", 10)
            # Use multi_cell for right alignment across full width, or set x/y
            pdf.set_x(pdf.l_margin)  # Reset X to left margin before right-aligned text
            pdf.cell(page_width_minus_margins, 6,
                     f"Oluşturulma Tarihi: {QDateTime.currentDateTime().toString('dd.MM.yyyy HH:mm:ss')}", 0,
                     new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="R")
            pdf.ln(5)

            # Giriş Metni
            if font_loaded:
                pdf.set_font("NotoSans", "", 10)
            else:
                pdf.set_font("Arial", "", 10)
            intro_text = (
                "Bu rapor, uygulama tarafından oluşturulan grafiklerin detaylı analizi ve karşılaştırmalarını sunmaktadır. "
                "Rapor, yapay zeka desteği olmadan, doğrudan grafik verilerinden elde edilmiştir.")
            pdf.multi_cell(page_width_minus_margins, 6, intro_text)  # Use page_width_minus_margins for width
            pdf.ln(8)  # Add a bit more space after intro text

            # Ana Grafik Görseli (Tüm grafikleri içeren ana Matplotlib penceresi)
            if self.matplotlib_widget.figure.axes:
                try:
                    self.matplotlib_widget.figure.tight_layout()
                    self.matplotlib_widget.figure.savefig(temp_image_path, format='png', dpi=300)

                    image_width = 180  # PDF sayfasında kullanılacak resim genişliği
                    fig_w, fig_h = self.matplotlib_widget.figure.get_size_inches()
                    image_height = image_width * (fig_h / fig_w)

                    # Check if enough space remains for image + buffer
                    required_space = image_height + 15  # Image height + buffer for title/spacing
                    if pdf.get_y() + required_space > pdf.h - pdf.b_margin:
                        pdf.add_page()

                    if font_loaded:
                        pdf.set_font("NotoSans", "B", 14)
                    else:
                        pdf.set_font("Arial", "B", 14)
                    pdf.cell(page_width_minus_margins, 10, "1. Oluşturulan Grafik Görselleri", 0, new_x=XPos.LMARGIN,
                             new_y=YPos.NEXT, align="L")
                    pdf.ln(2)
                    pdf.image(temp_image_path, x=(pdf.w - image_width) / 2, w=image_width)  # Resmi ortala
                    pdf.ln(10)  # Space after image

                except Exception as e:
                    QMessageBox.warning(self, "Rapor Hatası", f"Grafik görseli PDF'e eklenemedi: {e}")
                    if font_loaded:
                        pdf.set_font("NotoSans", "", 10)
                    else:
                        pdf.set_font("Arial", "", 10)
                    pdf.multi_cell(page_width_minus_margins, 6, f"Grafik görseli eklenirken hata oluştu: {e}")
            else:
                if font_loaded:
                    pdf.set_font("NotoSans", "", 10)
                else:
                    pdf.set_font("Arial", "", 10)
                pdf.multi_cell(page_width_minus_margins, 6,
                               "Raporlanacak bir grafik görseli bulunamadı. Lütfen önce bir grafik oluşturun.")
                pdf.ln(10)

            # Rapor İçeriği (generate_report_content_for_pdf'den gelen metin)
            report_lines = self.generate_report_content_for_pdf()

            pdf.add_page()  # Yeni sayfada rapor detaylarına geç
            if font_loaded:
                pdf.set_font("NotoSans", "B", 14)
            else:
                pdf.set_font("Arial", "B", 14)
            pdf.cell(page_width_minus_margins, 10, "2. Grafik Verisi ve Analiz Detayları", 0, new_x=XPos.LMARGIN,
                     new_y=YPos.NEXT, align="L")
            pdf.ln(5)  # Space after section title

            for line in report_lines:
                # Calculate estimated height for the current line
                # This is a rough estimate; multi_cell would calculate exact height
                # but this helps with proactive page breaks.
                estimated_line_height = 6  # Default line height
                if line.startswith('## '):
                    estimated_line_height = 7 + 5  # Line height + ln space
                elif line.startswith('### '):
                    estimated_line_height = 6 + 3
                elif line.strip() == '-' * 60:
                    estimated_line_height = 1 + 4  # Line height + ln space

                # Check for page break BEFORE drawing the line
                # Add a small buffer (e.g., 5mm) to trigger page break earlier
                if pdf.get_y() + estimated_line_height + 5 > pdf.h - pdf.b_margin:
                    pdf.add_page()
                    # Reset font on new page for consistent look
                    if font_loaded:
                        pdf.set_font("NotoSans", "", 10)
                    else:
                        pdf.set_font("Arial", "", 10)

                # Set X position to left margin before printing each line
                pdf.set_x(pdf.l_margin)

                # Başlıkları ve içeriği doğru font ve stil ile ekle
                if line.startswith('## '):  # Büyük başlıklar (A. Bireysel Grafik Detayları gibi)
                    pdf.ln(5)  # Extra space before main section title
                    if font_loaded:
                        pdf.set_font("NotoSans", "B", 12)
                    else:
                        pdf.set_font("Arial", "B", 12)
                    pdf.multi_cell(page_width_minus_margins, 7, line.replace('## ', ''), new_x=XPos.LMARGIN,
                                   new_y=YPos.NEXT)
                    pdf.ln(2)  # Space after main section title
                    if font_loaded:  # Revert to normal font after heading
                        pdf.set_font("NotoSans", "", 10)
                    else:
                        pdf.set_font("Arial", "", 10)
                elif line.startswith('### '):  # Orta seviye başlıklar (1. Grafik Detayları gibi)
                    pdf.ln(3)  # Extra space before sub-section title
                    if font_loaded:
                        pdf.set_font("NotoSans", "B", 10)
                    else:
                        pdf.set_font("Arial", "B", 10)
                    pdf.multi_cell(page_width_minus_margins, 6, line.replace('### ', ''), new_x=XPos.LMARGIN,
                                   new_y=YPos.NEXT)
                    pdf.ln(1)  # Space after sub-section title
                    if font_loaded:  # Revert to normal font after heading
                        pdf.set_font("NotoSans", "", 10)
                    else:
                        pdf.set_font("Arial", "", 10)
                elif line.strip().startswith('- **'):  # Kalın liste öğeleri (Grafik Tipi: Bar gibi)
                    if font_loaded:
                        pdf.set_font("NotoSans", "B", 10)
                    else:
                        pdf.set_font("Arial", "B", 10)
                    pdf.multi_cell(page_width_minus_margins, 5, line.replace('- **', '  ').replace('**:', ':'),
                                   new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    if font_loaded:  # Revert to normal font after bold item
                        pdf.set_font("NotoSans", "", 10)
                    else:
                        pdf.set_font("Arial", "", 10)
                elif line.startswith('  '):  # Girintili metin (data detayları gibi)
                    pdf.multi_cell(page_width_minus_margins, 5, line, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                elif line.startswith('- '):  # Normal liste öğeleri
                    pdf.multi_cell(page_width_minus_margins, 5, line, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                elif line.strip() == '-' * 60:  # Ayırıcı çizgiler
                    pdf.ln(2)  # Space before separator
                    pdf.cell(page_width_minus_margins, 1, line, 0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
                    pdf.ln(2)  # Space after separator
                elif line.strip() == '':  # Handle empty lines to create line breaks
                    pdf.ln(5)  # Consistent spacing for blank lines
                else:  # Diğer normal metinler
                    pdf.multi_cell(page_width_minus_margins, 5, line, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(10)  # Rapor sonu boşluk
            pdf.output(pdf_file_name)

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"PDF raporu oluşturulurken veya kaydedilirken bir hata oluştu: {e}\n"
                                               f"Lütfen dosya yolunun geçerli olduğundan ve yazma izniniz olduğundan emin olun. "
                                               "Dosya açık olabilir veya izin sorunu yaşanıyor olabilir.")
            print(f"DETAYLI HATA: {e}")
        finally:
            progress_dialog.close()
            if os.path.exists(temp_image_path):
                os.remove(temp_image_path)

        if os.path.exists(pdf_file_name):
            QMessageBox.information(self, "Rapor Oluşturuldu",
                                    f"Grafik raporu '{pdf_file_name}' (PDF) olarak başarıyla oluşturuldu.")
        else:
            pass


# Uygulama çalıştırma bloğu
if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())