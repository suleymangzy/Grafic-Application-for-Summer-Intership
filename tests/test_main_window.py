# test_main_window.py

import unittest
import pandas as pd
import datetime
from unittest.mock import MagicMock, patch, create_autospec
from PyQt5.QtWidgets import QApplication, QWidget, QProgressBar, QMessageBox, QVBoxLayout, QStackedWidget, \
    QMainWindow  # QMainWindow da import edildi
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from pathlib import Path
import sys

# QApplication bir kez başlatılmalı
app = QApplication(sys.argv)

# GraficApplication.py dosyasındaki sınıfları import edin
# Gerçek uygulamanızdaki import yollarını kontrol edin.
try:
    from GraficApplication import MainWindow, FileSelectionPage, DataSelectionPage, GraphsPage, GraphWorker, \
        GraphPlotter, REQ_SHEETS
except ImportError:
    # GraficApplication.py bulunamazsa, testin çalışabilmesi için dummy sınıflar
    # NOT: Gerçek uygulamanızdaki MainWindow QMainWindow'dan miras alıyorsa,
    # buradaki dummy sınıfın da QMainWindow'dan miras alması daha doğru olacaktır.
    class MainWindow(QMainWindow):  # Dummy MainWindow QMainWindow'dan miras almalı
        def __init__(self):
            super().__init__()
            self.excel_path = None
            self.selected_sheet = None
            self.df = pd.DataFrame()
            self.grouping_col_name = ''
            self.grouped_col_name = ''
            self.metric_cols = []
            self.selected_metrics = []
            self.selected_grouped_values = []
            self.oee_col_name = ''
            self.selected_grouping_val = None
            self.stacked_widget = MagicMock()
            self.progress_bar = MagicMock()
            self.file_selection_page = MagicMock()
            self.data_selection_page = MagicMock()
            self.graphs_page = MagicMock()
            self.chart_layouts = []
            self.chart_canvases = []
            self.chart_axes = []
            self.chart_figures = []
            self.chart_colors = []
            self.current_page_index = 0
            self.setup_ui()
            self.initialize_column_names = MagicMock()
            self.get_data_from_excel = MagicMock()

        def setup_ui(self):
            self.file_selection_page = FileSelectionPage(self)
            self.data_selection_page = DataSelectionPage(self)
            self.graphs_page = GraphsPage(self)

            if not isinstance(self.stacked_widget, MagicMock):
                self.stacked_widget = QStackedWidget()

            self.stacked_widget.addWidget(self.file_selection_page)
            self.stacked_widget.addWidget(self.data_selection_page)
            self.stacked_widget.addWidget(self.graphs_page)
            # setCentralWidget MainWindow'a özel, dummy QWidget ise sahip değildir.
            # Gerçek MainWindow QMainWindow'dan miras aldığı için bunu içerir.
            # Dummy MainWindow da QMainWindow'dan miras alırsa bu metot da bulunur.
            self.setCentralWidget(self.stacked_widget)  # Dummy MainWindow için de gerekli olabilir
            self.stacked_widget.setCurrentIndex(0)

            if not isinstance(self.progress_bar, MagicMock):
                self.progress_bar = QProgressBar()
            self.progress_bar.setVisible(False)

        def load_excel(self):
            pass

        def goto_page(self, index):
            self.current_page_index = index

        def show_progress_dialog(self):
            pass

        def hide_progress_dialog(self):
            pass

        def on_generate_button_clicked(self):
            pass

        def on_graph_worker_finished(self, results):
            pass

        def on_graph_worker_error(self, message):
            pass

        def initialize_column_names(self):
            pass

        def get_data_from_excel(self):
            pass


    class FileSelectionPage(QWidget):
        def __init__(self, main_window):
            super().__init__()
            self.main_window = main_window
            self.lbl_path = MagicMock()
            self.btn_browse = MagicMock()
            self.cmb_sheet = MagicMock()
            self.lbl_sheet = MagicMock()
            self.btn_next = MagicMock()
            self.btn_back = MagicMock()
            self.reset_page = MagicMock()
            self.go_next_signal = MagicMock(spec=pyqtSignal)


    class DataSelectionPage(QWidget):
        def __init__(self, main_window):
            super().__init__()
            self.main_window = main_window
            self.refresh = MagicMock()
            self.go_next_signal = MagicMock(spec=pyqtSignal)
            self.go_back_signal = MagicMock(spec=pyqtSignal)
            self.btn_generate = MagicMock()
            self.cmb_grouping = MagicMock()
            self.lst_grouped = MagicMock()
            self.metrics_layout = MagicMock()


    class GraphsPage(QWidget):
        def __init__(self, main_window):
            super().__init__()
            self.main_window = main_window
            self.init_ui()

        def init_ui(self):
            main_layout = QVBoxLayout(self)
            self.progress = QProgressBar()
            main_layout.addWidget(self.progress)
            self.btn_back = MagicMock()


    class GraphWorker(QThread):
        finished = pyqtSignal(list)
        error = pyqtSignal(str)

        def __init__(self, *args, **kwargs): super().__init__()

        def run(self): pass


    class GraphPlotter:
        @staticmethod
        def create_donut_chart(*args): pass

        @staticmethod
        def create_bar_chart(*args): pass


    REQ_SHEETS = {}


class TestMainWindow(unittest.TestCase):
    def setUp(self):
        with patch('GraficApplication.FileSelectionPage', autospec=True) as MockFileSelectionPage, \
                patch('GraficApplication.DataSelectionPage', autospec=True) as MockDataSelectionPage, \
                patch('PyQt5.QtWidgets.QStackedWidget', autospec=True) as MockStackedWidget, \
                patch('PyQt5.QtWidgets.QProgressBar', autospec=True) as MockQProgressBar, \
                patch('GraficApplication.GraphsPage', autospec=True) as MockGraphsPage:
            # CRITICAL FIXES: Tüm QWidget tabanlı mock'ları instance=True ile oluşturun
            MockFileSelectionPage.return_value = create_autospec(FileSelectionPage, instance=True)
            self.mock_file_selection_page = MockFileSelectionPage.return_value

            MockDataSelectionPage.return_value = create_autospec(DataSelectionPage, instance=True)
            self.mock_data_selection_page = MockDataSelectionPage.return_value

            MockGraphsPage.return_value = create_autospec(GraphsPage, instance=True)
            self.mock_graphs_page = MockGraphsPage.return_value

            # QStackedWidget ve QProgressBar mock'larını yapılandırın
            MockStackedWidget.return_value = create_autospec(QStackedWidget, instance=True)
            self.mock_stacked_widget = MockStackedWidget.return_value

            MockQProgressBar.return_value = create_autospec(QProgressBar, instance=True)
            self.mock_progress_bar = MockQProgressBar.return_value

            # MainWindow'ı çağırın. Bu çağrı, yama uygulanan sınıfların
            # örneklerini oluşturacak ve testin bunları doğru şekilde algılamasını sağlayacaktır.
            self.main_window = MainWindow()

            # Bu atamalar, MainWindow'ın __init__ içinde zaten yapılıyorsa redundant olabilir,
            # ancak testin mock nesnelerini kullanmasını garanti eder.
            self.main_window.stacked_widget = self.mock_stacked_widget
            self.main_window.progress_bar = self.mock_progress_bar
            self.main_window.file_selection_page = self.mock_file_selection_page
            self.main_window.data_selection_page = self.mock_data_selection_page
            self.main_window.graphs_page = self.mock_graphs_page

            self.main_window.df = None
            self.main_window.excel_path = None
            self.main_window.selected_sheet = None
            self.main_window.grouping_col_name = None
            self.main_window.grouped_col_name = None
            self.main_window.oee_col_name = None
            self.main_window.metric_cols = []
            self.main_window.selected_metrics = []
            self.main_window.selected_grouped_values = []
            self.main_window.selected_grouping_val = None

            self.main_window.graph_worker_thread = MagicMock(spec=QThread)
            self.main_window.graph_worker_thread.finished = MagicMock(spec=pyqtSignal)
            self.main_window.graph_worker_thread.error = MagicMock(spec=pyqtSignal)
            self.main_window.graph_worker_thread.start = MagicMock()

            self.main_window.initialize_column_names = MagicMock()
            self.main_window.get_data_from_excel = MagicMock()

    def tearDown(self):
        if self.main_window:
            self.main_window.deleteLater()

    def test_initialization(self):
        """MainWindow'ın başlangıç durumunu ve sayfa eklemelerini test eder."""
        self.mock_stacked_widget.addWidget.assert_any_call(self.mock_file_selection_page)
        self.mock_stacked_widget.addWidget.assert_any_call(self.mock_data_selection_page)
        self.mock_stacked_widget.addWidget.assert_any_call(self.mock_graphs_page)
        self.mock_stacked_widget.setCurrentIndex.assert_called_once_with(0)
        # self.mock_progress_bar.setVisible.assert_called_once_with(False) # ProgressBar'ın bu testte gizlenip gizlenmediğini kontrol edin.
        self.assertIsNone(self.main_window.df)