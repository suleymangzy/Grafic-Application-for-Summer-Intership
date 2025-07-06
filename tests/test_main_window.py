# test_main_window.py
import unittest
import pandas as pd
import numpy as np
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch, call, create_autospec

from PyQt5.QtWidgets import QApplication, QStackedWidget, QProgressBar, QMessageBox
from PyQt5.QtCore import Qt, QThread, pyqtSignal

# QApplication bir kez başlatılmalı
app = QApplication(sys.argv)

# GraficApplication.py dosyasındaki sınıfları import edin
from GraficApplication import MainWindow, GraphWorker, GraphPlotter, FileSelectionPage, DataSelectionPage, REQ_SHEETS, \
    excel_col_to_index, seconds_from_timedelta


class TestMainWindow(unittest.TestCase):
    def setUp(self):
        # MainWindow'ın içindeki bağımlılıkları mock'layarak izole test yapıyoruz
        self.mock_app = QApplication(sys.argv)  # MainWindow başlatılırken QApplication'a ihtiyaç duyabilir

        with patch('GraficApplication.FileSelectionPage', autospec=True) as MockFileSelectionPage, \
                patch('GraficApplication.DataSelectionPage', autospec=True) as MockDataSelectionPage, \
                patch('GraficApplication.QStackedWidget', autospec=True) as MockStackedWidget, \
                patch('GraficApplication.QProgressBar', autospec=True) as MockProgressBar:
            # Mock sayfalar
            self.mock_file_selection_page = MockFileSelectionPage.return_value
            self.mock_data_selection_page = MockDataSelectionPage.return_value

            # Mock stacked widget
            self.mock_stacked_widget = MockStackedWidget.return_value

            # Mock progress bar
            self.mock_progress_bar = MockProgressBar.return_value

            # MainWindow örneğini oluştur
            self.main_window = MainWindow()

            # Doğrudan UI elemanlarına erişimi sağlamak için mock'ları atama
            self.main_window.stacked_widget = self.mock_stacked_widget
            self.main_window.progress_bar = self.mock_progress_bar
            self.main_window.file_selection_page = self.mock_file_selection_page
            self.main_window.data_selection_page = self.mock_data_selection_page

            # Diğer önemli özelliklerin başlangıçta boş/varsayılan olduğundan emin ol
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

            # GraphWorker thread'ini de mock'la
            self.main_window.graph_worker_thread = MagicMock(spec=QThread)
            self.main_window.graph_worker_thread.finished = MagicMock(spec=pyqtSignal)
            self.main_window.graph_worker_thread.error = MagicMock(spec=pyqtSignal)
            self.main_window.graph_worker_thread.start = MagicMock()

    def tearDown(self):
        self.main_window.deleteLater()  # Widget'ı temizle
        self.mock_app.quit()  # QApplication'ı kapat

    def test_initialization(self):
        """MainWindow'ın başlangıç durumunu ve sayfa eklemelerini test eder."""
        self.mock_stacked_widget.addWidget.assert_any_call(self.mock_file_selection_page)
        self.mock_stacked_widget.addWidget.assert_any_call(self.mock_data_selection_page)
        self.mock_stacked_widget.setCurrentIndex.assert_called_once_with(0)
        self.mock_progress_bar.setVisible.assert_called_once_with(False)
        self.assertIsNone(self.main_window.df)

        # Sinyal bağlantılarını kontrol et
        self.mock_file_selection_page.go_next_signal.connect.assert_called_once()
        self.mock_data_selection_page.go_next_signal.connect.assert_called_once()
        self.mock_data_selection_page.go_back_signal.connect.assert_called_once()

    @patch('pandas.read_excel')
    def test_load_excel_success(self, mock_read_excel):
        """Excel dosyasının başarıyla yüklendiğini test eder."""
        mock_df = pd.DataFrame({'A': [1], 'B': [2]})
        mock_read_excel.return_value = mock_df

        self.main_window.excel_path = Path('test.xlsx')
        self.main_window.selected_sheet = 'Sheet1'
        self.main_window.load_excel()

        mock_read_excel.assert_called_once_with('test.xlsx', sheet_name='Sheet1')
        self.assertEqual(self.main_window.df.to_dict(), mock_df.to_dict())
        self.main_window.initialize_column_names.assert_called_once_with('Sheet1')
        self.mock_data_selection_page.refresh.assert_called_once()

    @patch('pandas.read_excel', side_effect=Exception("Read error"))
    @patch('PyQt5.QtWidgets.QMessageBox.critical')
    def test_load_excel_failure(self, mock_qmessagebox_critical, mock_read_excel):
        """Excel dosyasının yüklenmesi sırasında hata oluştuğunu test eder."""
        self.main_window.excel_path = Path('test.xlsx')
        self.main_window.selected_sheet = 'Sheet1'
        self.main_window.load_excel()

        mock_read_excel.assert_called_once()
        mock_qmessagebox_critical.assert_called_once()
        self.assertIn("Excel dosyası okunurken hata oluştu", mock_qmessagebox_critical.call_args[0][1])
        self.assertTrue(self.main_window.df.empty)  # df boşaltılmalı
        self.mock_file_selection_page.reset_page.assert_called_once()  # Hata durumunda sayfa sıfırlanmalı

    def test_get_data_from_excel_smd_oee(self):
        """SMD-OEE sayfası için veri çıkarma mantığını test eder."""
        # Gerçek bir dataframe oluştur
        self.main_window.df = pd.DataFrame({
            'Tarih': pd.to_datetime(['2023-01-01']),
            'Ürün Adı': ['Ürün A'],
            'OEE %': [90],
            'DUR_BEKLENEN': ['00:10:00'],
            'DUR_PLANSİZ': ['00:05:00'],
            'BOŞ_ÇALIŞMA': ['00:02:00']
        })
        self.main_window.selected_sheet = 'SMD-OEE'
        self.main_window.get_data_from_excel()

        self.assertEqual(self.main_window.grouping_col_name, 'Tarih')
        self.assertEqual(self.main_window.grouped_col_name, 'Ürün Adı')
        self.assertEqual(self.main_window.oee_col_name, 'OEE %')
        self.assertListEqual(sorted(self.main_window.metric_cols),
                             sorted(['DUR_BEKLENEN', 'DUR_PLANSİZ', 'BOŞ_ÇALIŞMA']))

    def test_get_data_from_excel_robot(self):
        """ROBOT sayfası için veri çıkarma mantığını test eder."""
        self.main_window.df = pd.DataFrame({
            'Tarih': pd.to_datetime(['2023-01-01']),
            'Robot Adı': ['Robot 1'],
            'OEE %': [85],
            'DURUS': ['00:15:00'],
            'CALISMA': ['00:30:00']
        })
        self.main_window.selected_sheet = 'ROBOT'
        self.main_window.get_data_from_excel()

        self.assertEqual(self.main_window.grouping_col_name, 'Tarih')
        self.assertEqual(self.main_window.grouped_col_name, 'Robot Adı')
        self.assertEqual(self.main_window.oee_col_name, 'OEE %')
        self.assertListEqual(sorted(self.main_window.metric_cols), sorted(['DURUS', 'CALISMA']))

    def test_get_data_from_excel_dalga_lehim(self):
        """DALGA_LEHİM sayfası için veri çıkarma mantığını test eder."""
        self.main_window.df = pd.DataFrame({
            'Tarih': pd.to_datetime(['2023-01-01']),
            'Makine Adı': ['Makine X'],
            'AP Değeri': [92],  # OEE_COL_NAME_DALGA_LEHIM
            'ONAYSIZ_DURUS': ['00:20:00'],
            'PLANLI_DURUS': ['00:10:00'],
            'HATALI_DURUS': ['00:05:00']
        })
        self.main_window.selected_sheet = 'DALGA_LEHİM'
        self.main_window.get_data_from_excel()

        self.assertEqual(self.main_window.grouping_col_name, 'Tarih')
        self.assertEqual(self.main_window.grouped_col_name, 'Makine Adı')
        self.assertEqual(self.main_window.oee_col_name, 'AP Değeri')
        self.assertListEqual(sorted(self.main_window.metric_cols),
                             sorted(['ONAYSIZ_DURUS', 'PLANLI_DURUS', 'HATALI_DURUS']))

    @patch('PyQt5.QtWidgets.QMessageBox.critical')
    def test_get_data_from_excel_missing_grouping_col(self, mock_qmessagebox_critical):
        """Gruplama sütunu eksik olduğunda hata işleme."""
        self.main_window.df = pd.DataFrame({
            'Yanlış Sütun': ['2023-01-01'],
            'Ürün Adı': ['Ürün A'],
            'OEE %': [90],
        })
        self.main_window.selected_sheet = 'SMD-OEE'  # 'Tarih' sütunu eksik
        self.main_window.get_data_from_excel()
        mock_qmessagebox_critical.assert_called_once()
        self.assertIn("Tanımlı sütunlar eksik", mock_qmessagebox_critical.call_args[0][1])
        self.assertIsNone(self.main_window.grouping_col_name)  # Sütun adları sıfırlanmalı
        self.main_window.file_selection_page.reset_page.assert_called_once()

    @patch('GraficApplication.GraphWorker')
    def test_on_generate_button_clicked(self, MockGraphWorker):
        """Grafik oluştur butonuna basıldığında GraphWorker'ın başlatılmasını test eder."""
        # Gerekli ana pencere özelliklerini ayarla
        self.main_window.df = pd.DataFrame({'A': [1], 'B': [2]})
        self.main_window.grouping_col_name = 'A'
        self.main_window.grouped_col_name = 'B'
        self.main_window.metric_cols = ['M1', 'M2']
        self.main_window.oee_col_name = 'OEE'
        self.main_window.selected_grouping_val = 'Group1'
        self.main_window.selected_grouped_values = ['Item1', 'Item2']
        self.main_window.selected_metrics = ['M1']

        # GraphWorker thread'i için mock ayarla
        mock_worker = MockGraphWorker.return_value

        self.main_window.on_generate_button_clicked()

        MockGraphWorker.assert_called_once_with(
            self.main_window.df,
            self.main_window.grouping_col_name,
            self.main_window.grouped_col_name,
            self.main_window.selected_grouped_values,
            self.main_window.selected_metrics,
            self.main_window.oee_col_name,
            self.main_window.selected_grouping_val
        )
        self.main_window.graph_worker_thread.start.assert_called_once()
        self.main_window.show_progress_dialog.assert_called_once()

    @patch('GraficApplication.GraphPlotter.create_donut_chart')
    @patch('GraficApplication.GraphPlotter.create_bar_chart')
    @patch('PyQt5.QtWidgets.QProgressBar.setValue')
    @patch('PyQt5.QtWidgets.QMessageBox.information')
    def test_on_graph_worker_finished_success(self, mock_qmessagebox_info, mock_set_value, mock_create_bar_chart,
                                              mock_create_donut_chart):
        """GraphWorker başarıyla bittiğinde sonuçların işlenmesini test eder."""
        # Başarılı sonuçlar simüle et
        mock_results = [
            ('ProductA', pd.Series({'M1': 600, 'M2': 300}), '90%'),
            ('ProductB', pd.Series({'M1': 120, 'M2': 60}), '80%')
        ]

        # Chart colors mock'u için (GraficApplication'daki sabitleri taklit etmeli)
        self.main_window.chart_colors = ['red', 'blue', 'green', 'purple', 'orange']

        # Grafik alanlarını mock'la (Canvas, Axes)
        self.main_window.chart_layouts = [MagicMock(spec=QVBoxLayout), MagicMock(spec=QVBoxLayout)]
        self.main_window.chart_canvases = [MagicMock(), MagicMock()]  # FigureCanvasQTAgg
        self.main_window.chart_axes = [MagicMock(), MagicMock()]  # Axes
        self.main_window.chart_figures = [MagicMock(), MagicMock()]  # Figure

        self.main_window.on_graph_worker_finished(mock_results)

        # Her bir ürün için grafiklerin çağrıldığını kontrol et
        self.assertEqual(mock_create_donut_chart.call_count, len(mock_results))
        self.assertEqual(mock_create_bar_chart.call_count, len(mock_results))

        # İlk ürün için çağrılan argümanları kontrol et
        donut_calls = mock_create_donut_chart.call_args_list
        bar_calls = mock_create_bar_chart.call_args_list

        # ProductA için donut chart
        self.assertEqual(donut_calls[0].args[0], self.main_window.chart_axes[0])
        pd.testing.assert_series_equal(donut_calls[0].args[1], mock_results[0][1])
        self.assertEqual(donut_calls[0].args[2], mock_results[0][2])
        self.assertListEqual(donut_calls[0].args[3], self.main_window.chart_colors)
        self.assertEqual(donut_calls[0].args[4], self.main_window.chart_figures[0])

        # ProductA için bar chart
        self.assertEqual(bar_calls[0].args[0], self.main_window.chart_axes[0])
        pd.testing.assert_series_equal(bar_calls[0].args[1], mock_results[0][1])
        self.assertEqual(bar_calls[0].args[2], mock_results[0][2])
        self.assertListEqual(bar_calls[0].args[3], self.main_window.chart_colors)

        self.assertEqual(self.main_window.current_page_index, 2)  # Grafik sayfasına geçilmeli
        self.main_window.hide_progress_dialog.assert_called_once()
        self.main_window.stacked_widget.setCurrentIndex.assert_called_with(2)  # Grafik sayfasına geçilmeli

    @patch('PyQt5.QtWidgets.QMessageBox.critical')
    def test_on_graph_worker_error(self, mock_qmessagebox_critical):
        """GraphWorker hata sinyali verdiğinde hata işleme."""
        error_message = "Bir hata oluştu."
        self.main_window.on_graph_worker_error(error_message)
        mock_qmessagebox_critical.assert_called_once_with(self.main_window, "Hata", error_message)
        self.main_window.hide_progress_dialog.assert_called_once()
        self.mock_file_selection_page.reset_page.assert_called_once()  # Hata durumunda ilk sayfaya dönülmeli

    def test_goto_page(self):
        """Sayfa geçişlerini test eder."""
        self.main_window.goto_page(1)
        self.mock_stacked_widget.setCurrentIndex.assert_called_once_with(1)
        self.assertEqual(self.main_window.current_page_index, 1)

        self.main_window.goto_page(0)  # İkinci kez çağrı
        self.assertEqual(self.mock_stacked_widget.setCurrentIndex.call_args_list[1].args[0], 0)
        self.assertEqual(self.main_window.current_page_index, 0)

    def test_show_progress_dialog(self):
        """İlerleme çubuğunun gösterilmesini test eder."""
        self.main_window.show_progress_dialog()
        self.mock_progress_bar.setVisible.assert_called_once_with(True)
        self.mock_progress_bar.setValue.assert_called_once_with(0)

    def test_hide_progress_dialog(self):
        """İlerleme çubuğunun gizlenmesini test eder."""
        self.main_window.hide_progress_dialog()
        self.mock_progress_bar.setVisible.assert_called_once_with(False)
        self.mock_progress_bar.setValue.assert_called_once_with(0)


if __name__ == '__main__':
    unittest.main(argv=['first-arg-is-ignored'], exit=False)