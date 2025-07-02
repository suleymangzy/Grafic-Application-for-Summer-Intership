"""
SMD‑ROBOT Pie‑Chart Desktop App
==============================
Tam işlevsel PyQt5 uygulaması – bellek hatalarını ve çoklu sinyal çakışmalarını
engelleyen iyileştirmeler + “Tüm Metrikleri Seç” kutucuğu dâhil.
(Kaynak: ChatGPT, Tem 2025)
"""

import sys
from pathlib import Path
import pandas as pd
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QStackedWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QListWidget,
    QListWidgetItem,
    QCheckBox,
    QComboBox,
)

import matplotlib
matplotlib.use("Agg")  # Çökmeleri azaltmak için arka planda çiz – Canvas yine görüntüler
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

VALID_SHEETS = {"SMD-OEE", "ROBOT", "DALGA_LEHİM"}
METRIC_START, METRIC_END = 7, 55  # H–BD (0‑based dizi)


# -----------------------------------------------------------------------------
# 1) Dosya Seçimi Sayfası
# -----------------------------------------------------------------------------
class FileSelectionPage(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.file_path = None
        self.init_ui()

    def init_ui(self):
        lay = QVBoxLayout(self)

        title = QLabel("1) Dosya Seçimi")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        lay.addWidget(title)

        self.btn_select = QPushButton("Excel Dosyası Seç…")
        self.btn_select.clicked.connect(self.open_dialog)
        lay.addWidget(self.btn_select)

        self.sheet_combo = QComboBox()
        lay.addWidget(self.sheet_combo)

        self.btn_next = QPushButton("İleri ⯈")
        self.btn_next.setEnabled(False)
        self.btn_next.clicked.connect(self.go_next)

        nav = QHBoxLayout()
        nav.addStretch()
        nav.addWidget(self.btn_next)
        lay.addLayout(nav)

    # ----------------------------------------------------------------------
    def open_dialog(self):
        fp, _ = QFileDialog.getOpenFileName(self, "Excel Dosyası Seç", str(Path.home()), "Excel (*.xlsx)")
        if not fp:
            return
        try:
            xls = pd.ExcelFile(fp)
            valid = sorted(set(xls.sheet_names) & VALID_SHEETS)
            if not valid:
                QMessageBox.warning(self, "Geçersiz", "Seçilen dosyada gerekli sayfalardan hiçbiri yok!")
                return
            self.file_path = fp
            self.sheet_combo.clear()
            self.sheet_combo.addItems(valid)
            self.btn_next.setEnabled(True)
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Dosya okunamadı:\n{e}")

    # ----------------------------------------------------------------------
    def go_next(self):
        self.parent.shared.update(
            {
                "file_path": self.file_path,
                "sheet_name": self.sheet_combo.currentText(),
            }
        )
        self.parent.data_page.load_data()
        self.parent.goto(1)


# -----------------------------------------------------------------------------
# 2) Veri Seçimi Sayfası
# -----------------------------------------------------------------------------
class DataSelectionPage(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.df = None
        self.init_ui()

    def init_ui(self):
        lay = QVBoxLayout(self)

        title = QLabel("2) Veri Seçimi")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        lay.addWidget(title)

        # Gruplama
        lay.addWidget(QLabel("Gruplama Değişkeni (Sütun A):"))
        self.grp_combo = QComboBox()
        lay.addWidget(self.grp_combo)

        # Gruplanan liste
        lay.addWidget(QLabel("Gruplanan Değişkenler (Sütun B):"))
        self.sub_list = QListWidget()
        lay.addWidget(self.sub_list)

        # Metrikler
        self.chk_all = QCheckBox("Tüm Metrikleri Seç/Deselect")
        self.chk_all.stateChanged.connect(self.toggle_metrics)
        lay.addWidget(self.chk_all)

        lay.addWidget(QLabel("Metrikler (H–BD):"))
        self.metric_list = QListWidget()
        self.metric_list.setSelectionMode(QListWidget.MultiSelection)
        lay.addWidget(self.metric_list)

        # Nav
        nav = QHBoxLayout()
        self.btn_back = QPushButton("⯇ Geri")
        self.btn_next = QPushButton("İleri ⯈")
        self.btn_back.clicked.connect(lambda: self.parent.goto(0))
        self.btn_next.clicked.connect(self.go_next)
        nav.addWidget(self.btn_back)
        nav.addStretch()
        nav.addWidget(self.btn_next)
        lay.addLayout(nav)

    # ------------------------------------------------------------------
    def load_data(self):
        fp, sheet = self.parent.shared["file_path"], self.parent.shared["sheet_name"]
        self.df = pd.read_excel(fp, sheet_name=sheet, header=0)

        # Disconnect previously connected signal to avoid stacking → stack overflow
        try:
            self.grp_combo.currentTextChanged.disconnect()
        except TypeError:
            pass

        groups = self.df.iloc[:, 0].dropna().unique()
        self.grp_combo.clear()
        self.grp_combo.addItems([str(g) for g in groups])
        self.grp_combo.currentTextChanged.connect(self.populate_subgroups)

        # Metrics
        self.metric_list.clear()
        for col in self.df.columns[METRIC_START : METRIC_END + 1]:
            if self.df[col].notna().any():
                item = QListWidgetItem(col)
                item.setSelected(True)
                self.metric_list.addItem(item)
        self.chk_all.setChecked(True)

        # Sub‑grups
        self.populate_subgroups()

    # ------------------------------------------------------------------
    def toggle_metrics(self, state):
        for i in range(self.metric_list.count()):
            self.metric_list.item(i).setSelected(state == Qt.Checked)

    # ------------------------------------------------------------------
    def populate_subgroups(self):
        sel = self.grp_combo.currentText()
        if not sel:
            return
        mask = self.df.iloc[:, 0] == sel
        subs = self.df.loc[mask, self.df.columns[1]].dropna().unique()
        self.sub_list.clear()
        for s in subs:
            self.sub_list.addItem(str(s))
        # Hepsi seçili varsayılan
        for i in range(self.sub_list.count()):
            self.sub_list.item(i).setSelected(True)

    # ------------------------------------------------------------------
    def go_next(self):
        sel_subs = [self.sub_list.item(i).text() for i in range(self.sub_list.count()) if self.sub_list.item(i).isSelected()]
        sel_mets = [self.metric_list.item(i).text() for i in range(self.metric_list.count()) if self.metric_list.item(i).isSelected()]
        if not sel_subs or not sel_mets:
            QMessageBox.warning(self, "Eksik", "En az bir gruplanan değişken ve metrik seçmelisiniz.")
            return

        self.parent.shared.update(
            {
                "group_val": self.grp_combo.currentText(),
                "subgroups": sel_subs,
                "metrics": sel_mets,
                "dataframe": self.df,
            }
        )
        self.parent.graph_page.render_charts()
        self.parent.goto(2)


# -----------------------------------------------------------------------------
# 3) Grafikler Sayfası
# -----------------------------------------------------------------------------
class GraphPage(QWidget):
    MAX_PLOTS = 120  # güvenlik limiti

    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.init_ui()

    def init_ui(self):
        lay = QVBoxLayout(self)
        title = QLabel("3) Grafikler")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        lay.addWidget(title)

        self.plot_area = QVBoxLayout()
        lay.addLayout(self.plot_area)

        nav = QHBoxLayout()
        self.btn_back = QPushButton("⯇ Geri")
        self.btn_back.clicked.connect(lambda: self.parent.goto(1))
        nav.addWidget(self.btn_back)
        nav.addStretch()
        lay.addLayout(nav)

    # ------------------------------------------------------------------
    def clear_area(self):
        while self.plot_area.count():
            child = self.plot_area.takeAt(0).widget()
            if child:
                child.setParent(None)

    # ------------------------------------------------------------------
    def render_charts(self):
        self.clear_area()
        st = self.parent.shared
        df = st["dataframe"]
        group_val = st["group_val"]
        subgroups = st["subgroups"]
        metrics = st["metrics"]

        total_expected = len(subgroups) * len(metrics)
        if total_expected > self.MAX_PLOTS:
            QMessageBox.warning(
                self,
                "Fazla Grafik",
                f"Oluşturulacak grafik sayısı ({total_expected}) çok yüksek. Lütfen seçimlerinizi azaltın.",
            )
            return

        mask_group = df.iloc[:, 0] == group_val
        for sub in subgroups:
            mask_sub = mask_group & (df.iloc[:, 1] == sub)
            df_sub = df.loc[mask_sub]
            if df_sub.empty:
                continue
            for met in metrics:
                series = df_sub[met].dropna()
                val = series.sum()
                if val <= 0:
                    continue
                # Draw single‑slice pie if only one value
                fig, ax = plt.subplots(figsize=(3.6, 3.6))
                ax.pie([val], labels=[sub] if len(metrics) == 1 else [met], autopct="%1.1f%%")
                ax.set_title(f"{sub} – {met}")

                canvas = FigureCanvas(fig)
                self.plot_area.addWidget(canvas)
                plt.close(fig)  # Belleği serbest bırak


# -----------------------------------------------------------------------------
# Ana Pencere
# -----------------------------------------------------------------------------
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SMD‑ROBOT Grafik Uygulaması")
        self.resize(900, 700)
        self.shared = {}

        self.stack = QStackedWidget()
        self.file_page = FileSelectionPage(self)
        self.data_page = DataSelectionPage(self)
        self.graph_page = GraphPage(self)

        self.stack.addWidget(self.file_page)
        self.stack.addWidget(self.data_page)
        self.stack.addWidget(self.graph_page)

        lay = QVBoxLayout(self)
        lay.addWidget(self.stack)

        # Basit stil
        self.setStyleSheet(
            """
            QWidget {font-family: Arial; font-size: 13px;}
            QPushButton {padding:6px 14px; border-radius:6px; background:#0d6efd; color:white;}
            QPushButton:disabled {background:#888;}
            QCheckBox {margin:4px 0;}
            """
        )

    # ---------------------------------------------
    def goto(self, idx: int):
        self.stack.setCurrentIndex(idx)


# -----------------------------------------------------------------------------
# Uygulama giriş noktası
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())
