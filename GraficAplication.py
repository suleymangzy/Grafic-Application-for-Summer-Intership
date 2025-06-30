import sys
import matplotlib
matplotlib.use('Qt5Agg')

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog,
    QListWidget, QListWidgetItem, QTextEdit, QMessageBox, QScrollArea, QSizePolicy
)
from PyQt5.QtCore import Qt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.pyplot as plt
import pandas as pd
from datetime import datetime
import mplcursors

class ScrollableCanvas(QScrollArea):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.canvas = FigureCanvas(plt.figure())
        self.setWidget(self.canvas)
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.canvas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SMD-OEE Pasta Grafik Uygulaması")
        self.setGeometry(100, 100, 1200, 800)

        self.layout = QVBoxLayout(self)

        self.label = QLabel("Dosya seçilmedi")
        self.layout.addWidget(self.label)

        self.btn_select = QPushButton("Excel Dosyası Seç")
        self.btn_select.clicked.connect(self.select_file)
        self.layout.addWidget(self.btn_select)

        self.layout.addWidget(QLabel("Tarih(ler) Seç"))
        self.date_list = QListWidget()
        self.date_list.setSelectionMode(QListWidget.MultiSelection)
        self.layout.addWidget(self.date_list)

        self.btn_plot = QPushButton("Grafikleri Oluştur")
        self.btn_plot.clicked.connect(self.plot_graphs)
        self.layout.addWidget(self.btn_plot)

        self.scroll_area = ScrollableCanvas(self)
        self.layout.addWidget(self.scroll_area)

        self.text_area = QTextEdit()
        self.text_area.setReadOnly(True)
        self.layout.addWidget(self.text_area)

        self.data = None
        self.current_selected_dates = []
        self.current_total_graphs = 0

    def resizeEvent(self, event):
        if self.current_total_graphs > 0:
            width = self.scroll_area.width()
            cols = 3
            rows = (self.current_total_graphs + cols - 1) // cols
            height = max(rows * 320, 400)
            dpi = self.scroll_area.canvas.figure.dpi
            self.scroll_area.canvas.figure.set_size_inches(width / dpi, height / dpi)
            self.scroll_area.canvas.draw_idle()
        super().resizeEvent(event)

    def select_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Dosya Seç", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.label.setText(f"Seçilen dosya: {path}")
            try:
                # Önce dosyadaki tüm sayfa isimlerini al
                xls = pd.ExcelFile(path)
                if 'SMD-OEE' not in xls.sheet_names:
                    QMessageBox.critical(self, "Hata", "'SMD-OEE' sayfası bu dosyada bulunamadı!")
                    return
                # Sadece SMD-OEE sayfasını oku
                self.data = pd.read_excel(xls, sheet_name='SMD-OEE')
            except Exception as e:
                QMessageBox.critical(self, "Hata", f"Excel okunamadı: {e}")
                return

            if "Tarih" not in self.data.columns:
                QMessageBox.critical(self, "Hata", "'Tarih' sütunu bulunamadı!")
                return

            if "Ürün" not in self.data.columns:
                QMessageBox.critical(self, "Hata", "'Ürün' sütunu bulunamadı!")
                return

            try:
                self.data["Tarih"] = pd.to_datetime(self.data["Tarih"])
            except Exception as e:
                QMessageBox.critical(self, "Hata", f"Tarih sütunu datetime'a çevrilemedi: {e}")
                return

            self.date_list.clear()
            unique_dates = sorted(self.data["Tarih"].dt.date.dropna().unique())
            for d in unique_dates:
                item = QListWidgetItem(str(d))
                self.date_list.addItem(item)

    def plot_graphs(self):
        try:
            if self.data is None:
                QMessageBox.warning(self, "Hata", "Lütfen önce dosya seçin!")
                return

            selected_dates = [datetime.strptime(item.text(), "%Y-%m-%d").date() for item in self.date_list.selectedItems()]

            if not selected_dates:
                QMessageBox.warning(self, "Hata", "Lütfen en az bir tarih seçin!")
                return

            columns = list(self.data.columns)
            value_columns = [col for col in columns if col not in ["Tarih", "Ürün"]]

            filtered_df = self.data[self.data["Tarih"].dt.date.isin(selected_dates)]

            total_graphs = filtered_df.shape[0]
            if total_graphs == 0:
                QMessageBox.warning(self, "Hata", "Seçilen tarihlerde ürün bulunamadı!")
                return

            self.current_selected_dates = selected_dates
            self.current_total_graphs = total_graphs

            fig = self.scroll_area.canvas.figure
            fig.clear()

            cols = 3  # daha mantıklı bir sütun sayısı
            rows = (total_graphs + cols - 1) // cols


            width = self.scroll_area.width()
            height = max(rows * 320, 400)
            dpi = fig.dpi
            fig.set_size_inches(width / dpi, height / dpi)

            # Alt satır:
            fig.subplots_adjust(hspace=1.0, wspace=0.5, bottom=0.2)  # tight_layout yerine

            self.text_area.clear()

            plot_num = 1
            for idx, row in filtered_df.iterrows():
                urun = row.get("Ürün", "Bilinmeyen Ürün")
                tarih = row["Tarih"].date()

                sureler = row[value_columns]
                sureler = sureler[pd.to_numeric(sureler, errors='coerce').notnull()]
                sureler = sureler[sureler > 0]

                labels = sureler.index.tolist()
                values = sureler.values.tolist()

                if len(values) == 0:
                    continue

                ax = fig.add_subplot(rows, cols, plot_num)

                color_map = plt.colormaps['tab20']
                colors = [color_map(i / len(labels)) for i in range(len(labels))]

                wedges, texts, autotexts = ax.pie(
                    values, labels=None, autopct='%1.1f%%', colors=colors, startangle=90
                )

                ax.set_title(f"{urun}", fontsize=11)

                legend_labels = [f"{label} ({value/sum(values)*100:.1f}%)" for label, value in zip(labels, values)]
                ax.legend(wedges, legend_labels, loc="lower center", bbox_to_anchor=(0.5, -0.15), ncol=2, fontsize=8, frameon=False)

                cursor = mplcursors.cursor(wedges, hover=True)
                @cursor.connect("add")
                def on_add(sel, labels=labels, values=values):
                    label = labels[sel.index]
                    val = values[sel.index]
                    sel.annotation.set_text(f"{label}\nDeğer: {val}")

                aciklama = f"Ürün: {urun} | Tarih: {tarih}\n"
                for label, value in zip(labels, values):
                    aciklama += f"{label}: {value}\n"
                self.text_area.append(aciklama + "\n" + "-"*40 + "\n")

                plot_num += 1

            fig.tight_layout()
            self.scroll_area.canvas.draw()

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Grafik oluşturulurken hata oluştu:\n{e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = App()
    ex.show()
    sys.exit(app.exec_())