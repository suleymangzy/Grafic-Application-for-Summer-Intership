"""GraficApplication.py

Sürüm 3 – 3 Temmuz 2025
───────────────────────
Tamamen baştan yazılmış, stabil ve kullanıcı dostu PyQt5 uygulaması.
• Dosya Seçimi → Veri Seçimi → Grafikler üç‑sayfa akışı.
• Excel dosyası > uygun sheet kontrolü (SMD‑OEE/ROBOT/DALGA_LEHİM).
• A sütunu = Gruplama, B sütunu = Gruplanan.
• H–BD sütunları metrik; AP daima hariç.
• BP sütunu dinamik başlık okuma.
• Grafikler sayfa başına 4 adet, ← Önceki / Sonraki → ile gezilebilir.
• Matplotlib figürleri ana iş parçacığında (backend = "Agg").
• GraphWorker QThread sadece veriyi hazırlar; UI donmaz.
• Ayrıntılı logging ve QMessageBox hata yakalama.
"""

from __future__ import annotations

import sys
import logging
from pathlib import Path
from typing import List, Tuple

import pandas as pd
import matplotlib

matplotlib.use("Agg")  # GUI bağımlı olmayan backend
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QFileDialog,
    QPushButton,
    QLabel,
    QVBoxLayout,
    QHBoxLayout,
    QListWidget,
    QListWidgetItem,
    QComboBox,
    QMessageBox,
    QProgressBar,
    QStackedWidget,
    QScrollArea,
    QFrame,
)

# ────────────────────────────────────────────────────────────────────────────────
# Sabitler & Logging
# ────────────────────────────────────────────────────────────────────────────────
GRAPHS_PER_PAGE = 4
REQ_SHEETS = {"SMD-OEE", "ROBOT", "DALGA_LEHİM"}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)

# ────────────────────────────────────────────────────────────────────────────────
# Yardımcı Fonksiyonlar
# ────────────────────────────────────────────────────────────────────────────────

def excel_col_to_index(col: str) -> int:
    """'AP' ‑> 41 gibi 0‑tabanlı index döndür."""
    idx = 0
    for c in col.upper():
        if not c.isalpha():
            raise ValueError(f"Geçersiz sütun: {col}")
        idx = idx * 26 + (ord(c) - ord("A") + 1)
    return idx - 1


def seconds_from_timedelta(series: pd.Series) -> pd.Series:
    td = pd.to_timedelta(series, errors="coerce")
    return td.dt.total_seconds().fillna(0)

# ────────────────────────────────────────────────────────────────────────────────
# QThread: GraphWorker
# ────────────────────────────────────────────────────────────────────────────────

class GraphWorker(QThread):
    finished = pyqtSignal(list)  # List[Tuple[str, pd.Series]]
    progress = pyqtSignal(int)
    error = pyqtSignal(str)

    def __init__(
        self,
        df: pd.DataFrame,
        grouping_col: str,
        grouped_values: List[str],
        metric_cols: List[str],
    ) -> None:
        super().__init__()
        self.df = df.copy()
        self.grouping_col = grouping_col
        self.grouped_values = grouped_values
        self.metric_cols = metric_cols

    # ------------------------------------------------------------------
    def run(self) -> None:
        try:
            metrics_sec = self.df[self.metric_cols].apply(seconds_from_timedelta)
            df_proc = pd.concat([self.df[[self.grouping_col]], metrics_sec], axis=1)
            results: List[Tuple[str, pd.Series]] = []
            total = len(self.grouped_values)
            for i, val in enumerate(self.grouped_values, 1):
                subset = df_proc[df_proc[self.grouping_col] == val]
                sums = subset[self.metric_cols].sum(); sums = sums[sums > 0]
                if not sums.empty:
                    results.append((val, sums))
                self.progress.emit(int(i / total * 100))
            self.finished.emit(results)
        except Exception as exc:  # noqa: BLE001
            logging.exception("GraphWorker hata")
            self.error.emit(str(exc))

# ────────────────────────────────────────────────────────────────────────────────
# Sayfalar
# ────────────────────────────────────────────────────────────────────────────────

class FileSelectionPage(QWidget):
    def __init__(self, main_window: "MainWindow") -> None:  # noqa: F821
        super().__init__()
        self.main_window = main_window
        # UI elemanları
        self.lbl_path = QLabel("Henüz dosya seçilmedi")
        self.cmb_sheet = QComboBox(); self.cmb_sheet.setEnabled(False)
        self.btn_browse = QPushButton(".xlsx dosyası seç…"); self.btn_browse.clicked.connect(self.browse)
        self.btn_next = QPushButton("İleri →"); self.btn_next.setEnabled(False); self.btn_next.clicked.connect(self.go_next)
        layout = QVBoxLayout(self)
        layout.addWidget(self.lbl_path); layout.addWidget(self.btn_browse)
        layout.addWidget(QLabel("İşlenecek Sayfa:")); layout.addWidget(self.cmb_sheet)
        layout.addStretch(1); layout.addWidget(self.btn_next, alignment=Qt.AlignRight)

    # ------------------------------------------------------------------
    def browse(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Excel seç", str(Path.home()), "Excel Files (*.xlsx)")
        if not path:
            return
        try:
            xls = pd.ExcelFile(path)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Okuma hatası", str(e)); return
        sheets = sorted(REQ_SHEETS & set(xls.sheet_names))
        if not sheets:
            QMessageBox.warning(self, "Uygun sayfa yok", "Seçilen dosyada istenen sheet bulunamadı.")
            return
        self.main_window.excel_path = Path(path)
        self.lbl_path.setText(Path(path).name)
        self.cmb_sheet.clear(); self.cmb_sheet.addItems(sheets); self.cmb_sheet.setEnabled(True)
        self.btn_next.setEnabled(True)
        logging.info("Dosya seçildi: %s", path)

    def go_next(self) -> None:
        self.main_window.selected_sheet = self.cmb_sheet.currentText()
        self.main_window.load_excel()
        self.main_window.goto_page(1)


class DataSelectionPage(QWidget):
    def __init__(self, main_window: "MainWindow") -> None:  # noqa: F821
        super().__init__()
        self.main_window = main_window
        # UI
        self.cmb_grouping = QComboBox()
        self.lst_grouped = QListWidget(); self.lst_grouped.setSelectionMode(QListWidget.MultiSelection)
        self.lst_metrics = QListWidget(); self.lst_metrics.setSelectionMode(QListWidget.MultiSelection)
        self.btn_back = QPushButton("← Geri"); self.btn_back.clicked.connect(lambda: self.main_window.goto_page(0))
        self.btn_next = QPushButton("İleri →"); self.btn_next.clicked.connect(self.go_next)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Gruplama Değişkeni (A):")); layout.addWidget(self.cmb_grouping)
        layout.addWidget(QLabel("Gruplanan Değişkenler (B):")); layout.addWidget(self.lst_grouped)
        layout.addWidget(QLabel("Metrikler (H–BD, AP hariç):")); layout.addWidget(self.lst_metrics)
        nav = QHBoxLayout(); nav.addWidget(self.btn_back); nav.addStretch(1); nav.addWidget(self.btn_next); layout.addLayout(nav)
        self.cmb_grouping.currentIndexChanged.connect(self.populate_grouped)

    # ------------------------------------------------------------------
    def refresh(self) -> None:
        df = self.main_window.df
        # Gruplama değişkeni (A sütunu)
        self.cmb_grouping.clear()
        grouping_vals = sorted(df.iloc[:, 0].dropna().astype(str).unique())
        self.cmb_grouping.addItems(grouping_vals)
        # Metrikler listesi
        self.lst_metrics.clear()
        for m in self.main_window.metric_cols:
            item = QListWidgetItem(m); item.setSelected(True); self.lst_metrics.addItem(item)
        self.populate_grouped()

    def populate_grouped(self) -> None:
        df = self.main_window.df
        val = self.cmb_grouping.currentText()
        if not val:
            return
        subset = df[df.iloc[:, 0].astype(str) == val]
        grouped_vals = sorted(subset.iloc[:, 1].dropna().astype(str).unique())
        self.lst_grouped.clear()
        for gv in grouped_vals:
            item = QListWidgetItem(gv); item.setSelected(True); self.lst_grouped.addItem(item)

    def go_next(self) -> None:
        grouped_sel = [i.text() for i in self.lst_grouped.selectedItems()]
        metric_sel = [i.text() for i in self.lst_metrics.selectedItems()]
        if not grouped_sel or not metric_sel:
            QMessageBox.warning(self, "Seçim eksik", "En az bir gruplama ve metrik seçmelisiniz.")
            return
        self.main_window.grouped_values = grouped_sel
        self.main_window.selected_metrics = metric_sel
        self.main_window.goto_page(2)


class GraphsPage(QWidget):
    def __init__(self, main_window: "MainWindow") -> None:  # noqa: F821
        super().__init__()
        self.main_window = main_window
        self.worker: GraphWorker | None = None
        self.progress = QProgressBar(); self.progress.setAlignment(Qt.AlignCenter)
        # Scrollable container
        self.scroll = QScrollArea(); self.scroll.setWidgetResizable(True)
        self.canvas_holder = QWidget(); self.scroll.setWidget(self.canvas_holder)
        self.vbox_canvases = QVBoxLayout(self.canvas_holder)
        # Navigation
        self.lbl_page = QLabel("Sayfa 0 / 0")
        self.btn_prev = QPushButton("← Önceki"); self.btn_prev.clicked.connect(self.prev_page)
        self.btn_next = QPushButton("Sonraki →"); self.btn_next.clicked.connect(self.next_page)
        self.btn_save = QPushButton("Grafikleri Kaydet…"); self.btn_save.clicked.connect(self.save_graphs)
        self.btn_back = QPushButton("← Veri"); self.btn_back.clicked.connect(lambda: self.main_window.goto_page(1))
        nav_top = QHBoxLayout(); nav_top.addWidget(self.btn_back); nav_top.addStretch(1); nav_top.addWidget(self.lbl_page); nav_top.addStretch(1); nav_top.addWidget(self.btn_save)
        nav_bottom = QHBoxLayout(); nav_bottom.addStretch(1); nav_bottom.addWidget(self.btn_prev); nav_bottom.addWidget(self.btn_next)
        layout = QVBoxLayout(self); layout.addWidget(self.progress); layout.addLayout(nav_top); layout.addWidget(self.scroll); layout.addLayout(nav_bottom)
        # Data holders
        self.figures: List[Tuple[str, Figure]] = []
        self.current_page = 0

    # ------------------------------------------------------------------
    def enter_page(self) -> None:
        self.clear_canvases(); self.figures.clear(); self.current_page = 0; self.update_page_label()
        df = self.main_window.df
        self.worker = GraphWorker(
            df=df,
            grouping_col=df.columns[1],  # B sütunu (gruplanan) bazlı grafikler
            grouped_values=self.main_window.grouped_values,
            metric_cols=self.main_window.selected_metrics,
        )
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished.connect(self.on_results)
        self.worker.error.connect(lambda m: QMessageBox.critical(self, "Hata", m))
        self.worker.start()

    def on_results(self, results: List[Tuple[str, pd.Series]]) -> None:
        self.progress.setValue(100)
        if not results:
            QMessageBox.information(self, "Veri yok", "Grafik oluşturulamadı.")
            return
        df = self.main_window.df; bp_idx = excel_col_to_index("BP"); bp_col = df.columns[bp_idx] if bp_idx < len(df.columns) else None
        for gval, series in results:
            fig = Figure(figsize=(6, 6), dpi=110); ax = fig.add_subplot(111)
            wedges, texts, autotexts = ax.pie(series.values, labels=series.index, autopct="%1.0f%%", startangle=90, counterclock=False)
            ax.axis("equal")
            title = f"{df.columns[1]}: {gval}"  # B sütunu başlığı
            if bp_col:
                bp_val = df[df.iloc[:, 1].astype(str) == gval][bp_col].iloc[0] if not df.empty else ""
                title += f" – {bp_col}: {bp_val}"
            ax.set_title(title, fontweight="bold")
            self.figures.append((gval, fig))
        self.display_page()

    # ------------------------------------------------------------------
    def display_page(self) -> None:
        self.clear_canvases()
        start = self.current_page * GRAPHS_PER_PAGE
        for _, fig in self.figures[start : start + GRAPHS_PER_PAGE]:
            canvas = FigureCanvas(fig)
            frame = QFrame(); frame.setFrameShape(QFrame.Box)
            vb = QVBoxLayout(frame); vb.addWidget(canvas)
            self.vbox_canvases.addWidget(frame)
        self.vbox_canvases.addStretch(1)
        self.update_page_label()

    def clear_canvases(self) -> None:
        while self.vbox_canvases.count():
            item = self.vbox_canvases.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def update_page_label(self) -> None:
        total_pages = max(1, (len(self.figures) + GRAPHS_PER_PAGE - 1) // GRAPHS_PER_PAGE)
        self.lbl_page.setText(f"Sayfa {self.current_page + 1} / {total_pages}")

    def next_page(self) -> None:
        if (self.current_page + 1) * GRAPHS_PER_PAGE >= len(self.figures):
            return
        self.current_page += 1; self.display_page()

    def prev_page(self) -> None:
        if self.current_page == 0:
            return
        self.current_page -= 1; self.display_page()

    def save_graphs(self) -> None:
        if not self.figures:
            QMessageBox.warning(self, "Grafik yok", "Kaydedilecek grafik bulunamadı.")
            return
        out_dir = QFileDialog.getExistingDirectory(self, "Kayıt klasörünü seç")
        if not out_dir:
            return
        for idx, (gval, fig) in enumerate(self.figures, 1):
            out_path = Path(out_dir) / f"pie_{idx:02d}_{gval}.png"
            fig.savefig(out_path, bbox_inches="tight")
        QMessageBox.information(self, "Kaydedildi", "Grafikler kaydedildi.")

# ────────────────────────────────────────────────────────────────────────────────
# Main Window
# ────────────────────────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Pasta Grafik Rapor Uygulaması")
        self.resize(1000, 800)
        # Paylaşılan state
        self.excel_path: Path | None = None
        self.selected_sheet: str = ""
        self.df: pd.DataFrame = pd.DataFrame()
        self.metric_cols: List[str] = []
        self.grouped_values: List[str] = []
        self.selected_metrics: List[str] = []
        # Sayfalar
        self.stack = QStackedWidget(); self.setCentralWidget(self.stack)
        self.page_file = FileSelectionPage(self)
        self.page_data = DataSelectionPage(self)
        self.page_graph = GraphsPage(self)
        self.stack.addWidget(self.page_file)
        self.stack.addWidget(self.page_data)
        self.stack.addWidget(self.page_graph)

    # ------------------------------------------------------------------
    def goto_page(self, index: int) -> None:
        self.stack.setCurrentIndex(index)
        if index == 1:
            self.page_data.refresh()
        elif index == 2:
            self.page_graph.enter_page()

    # ------------------------------------------------------------------
    def load_excel(self) -> None:
        assert self.excel_path and self.selected_sheet
        logging.info("Excel okunuyor: %s | Sheet: %s", self.excel_path, self.selected_sheet)
        try:
            self.df = pd.read_excel(self.excel_path, sheet_name=self.selected_sheet)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Okuma hatası", str(e)); return
        # Sütun kontrolleri & metrik listesi oluşturma
        h_idx, bd_idx, ap_idx = excel_col_to_index("H"), excel_col_to_index("BD"), excel_col_to_index("AP")
        potential_metrics = self.df.columns[h_idx : bd_idx + 1].tolist()
        if ap_idx < len(self.df.columns):
            ap_name = self.df.columns[ap_idx]
            if ap_name in potential_metrics:
                potential_metrics.remove(ap_name)
        # Tamamen boş sütunları dışla
        self.metric_cols = [c for c in potential_metrics if not self.df[c].dropna().empty]
        logging.info("%d geçerli metrik bulundu", len(self.metric_cols))

# ────────────────────────────────────────────────────────────────────────────────
# main()
# ────────────────────────────────────────────────────────────────────────────────

def main() -> None:
    app = QApplication(sys.argv)
    win = MainWindow(); win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    print(">> GraficApplication – Sürüm 3 – 3 Tem 2025 – page 4 grafik")
    main()
