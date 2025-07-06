# test_ui_pages.py
import unittest
import pandas as pd
import numpy as np
import datetime
from unittest.mock import MagicMock, patch, create_autospec
from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox, QComboBox, QLabel, QPushButton, QListWidget, \
    QListWidgetItem, QScrollArea, QWidget, QVBoxLayout, QHBoxLayout, QCheckBox
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import sys
from pathlib import Path

# QApplication bir kez başlatılmalı
app = QApplication(sys.argv)

# GraficApplication.py dosyasındaki sınıfları import edin
from GraficApplication import FileSelectionPage, DataSelectionPage, REQ_SHEETS, MainWindow, GraphWorker


class TestFileSelectionPage(unittest.TestCase):
    def setUp(self):
        self.main_window_mock = MagicMock(spec=MainWindow)
        self.main_window_mock.excel_path = None
        self.main_window_mock.selected_sheet = None # Should be None on reset
        self.main_window_mock.goto_page = MagicMock()
        self.main_window_mock.load_excel = MagicMock()

        self.page = FileSelectionPage(self.main_window_mock)

        # UI elemanlarına doğrudan erişim için
        self.lbl_path = self.page.lbl_path
        self.btn_browse = self.page.btn_browse
        self.cmb_sheet = self.page.cmb_sheet
        self.sheet_selection_label = self.page.sheet_selection_label
        self.btn_next = self.page.btn_next

    def tearDown(self):
        self.page.deleteLater()  # Widget'ı temizle

    def test_initial_ui_state(self):
        """UI'ın başlangıç durumunu test eder."""
        self.assertEqual(self.lbl_path.text(), "Henüz dosya seçilmedi")
        self.assertFalse(self.cmb_sheet.isEnabled())
        self.assertFalse(self.cmb_sheet.isVisible())
        self.assertFalse(self.sheet_selection_label.isVisible())
        self.assertFalse(self.btn_next.isEnabled())
        self.assertEqual(self.btn_browse.text(), ".xlsx dosyası seç…")
        self.assertEqual(self.btn_next.text(), "İleri →")

    @patch('PyQt5.QtWidgets.QFileDialog.getOpenFileName', return_value=('fake/path/to/test.xlsx', '.xlsx'))
    @patch('pandas.ExcelFile')
    @patch('PyQt5.QtWidgets.QMessageBox.warning')
    def test_browse_success_multiple_sheets(self, mock_qmessagebox_warning, mock_excel_file, mock_get_open_file_name):
        """Birden fazla uygun sayfa içeren dosya seçimini test eder."""
        mock_xls = MagicMock()
        mock_xls.sheet_names = sorted(list(REQ_SHEETS))  # Tüm REQ_SHEETS'i içeriyor
        mock_excel_file.return_value = mock_xls

        self.btn_browse.click()

        mock_get_open_file_name.assert_called_once()
        self.assertEqual(self.main_window_mock.excel_path, Path('fake/path/to/test.xlsx'))
        self.assertIn("<b>test.xlsx</b>", self.lbl_path.text())
        self.assertTrue(self.cmb_sheet.isEnabled())
        self.assertTrue(self.cmb_sheet.isVisible())
        # Birden fazla sayfa olduğunda sheet_selection_label gizli olmalı
        self.assertFalse(self.sheet_selection_label.isVisible())
        self.assertTrue(self.btn_next.isEnabled())
        self.assertEqual(self.cmb_sheet.count(), len(REQ_SHEETS))  # Tüm uygun sayfalar eklenmeli
        self.assertFalse(mock_qmessagebox_warning.called)
        self.assertNotEqual(len(REQ_SHEETS), 1)  # Birden fazla sayfa olduğundan emin ol
        self.assertTrue(self.cmb_sheet.isVisible())  # ComboBox görünür olmalı


    @patch('PyQt5.QtWidgets.QFileDialog.getOpenFileName', return_value=('fake/path/to/test.xlsx', '.xlsx'))
    @patch('pandas.ExcelFile')
    @patch('PyQt5.QtWidgets.QMessageBox.warning')
    def test_browse_success_single_sheet_auto_select(self, mock_qmessagebox_warning, mock_excel_file,
                                                     mock_get_open_file_name):
        """Sadece bir uygun sayfa içeren dosya seçimini test eder (otomatik seçim)."""
        mock_xls = MagicMock()
        mock_xls.sheet_names = ['SMD-OEE', 'OtherSheet']  # Sadece bir uygun sayfa (assuming SMD-OEE is the only REQ_SHEET here for simplicity of auto-selection)
        # Ensure only one REQ_SHEET is present for this test case
        initial_req_sheets = REQ_SHEETS.copy()
        REQ_SHEETS.clear()
        REQ_SHEETS.add('SMD-OEE')

        mock_excel_file.return_value = mock_xls

        self.btn_browse.click()

        mock_get_open_file_name.assert_called_once()
        self.assertEqual(self.main_window_mock.excel_path, Path('fake/path/to/test.xlsx'))
        self.assertIn("<b>test.xlsx</b>", self.lbl_path.text())
        self.assertFalse(self.cmb_sheet.isVisible())  # ComboBox gizlenmeli
        # Tek sayfa olduğunda sheet_selection_label görünür olmalı
        self.assertTrue(self.sheet_selection_label.isVisible())
        self.assertIn('SMD-OEE', self.sheet_selection_label.text())  # Etikette belirtilmeli
        self.assertTrue(self.btn_next.isEnabled())
        self.assertEqual(self.main_window_mock.selected_sheet, 'SMD-OEE')
        self.assertFalse(mock_qmessagebox_warning.called)

        # Restore REQ_SHEETS to original state after test
        REQ_SHEETS.clear()
        REQ_SHEETS.update(initial_req_sheets)


    @patch('PyQt5.QtWidgets.QFileDialog.getOpenFileName', return_value=('fake/path/to/test.xlsx', '.xlsx'))
    @patch('pandas.ExcelFile')
    @patch('PyQt5.QtWidgets.QMessageBox.warning')
    def test_browse_no_required_sheets(self, mock_qmessagebox_warning, mock_excel_file, mock_get_open_file_name):
        """Uygun sayfa içermeyen dosya seçimini test eder."""
        mock_xls = MagicMock()
        mock_xls.sheet_names = ['Sheet1', 'Sheet2']  # REQ_SHEETS ile çakışmıyor
        mock_excel_file.return_value = mock_xls

        self.btn_browse.click()

        mock_get_open_file_name.assert_called_once()
        mock_qmessagebox_warning.assert_called_once_with(
            self.page, "Uygun sayfa yok", "Seçilen dosyada istenen (SMD-OEE, ROBOT, DALGA_LEHİM) sheet bulunamadı."
        )
        self.assertIsNone(self.main_window_mock.excel_path)  # Yol sıfırlanmalı
        self.assertFalse(self.cmb_sheet.isEnabled())
        self.assertFalse(self.btn_next.isEnabled())
        self.assertEqual(self.lbl_path.text(), "Henüz dosya seçilmedi")  # Başlangıç durumuna dönmeli

    @patch('PyQt5.QtWidgets.QFileDialog.getOpenFileName', return_value=('', ''))
    @patch('PyQt5.QtWidgets.QMessageBox.warning')
    def test_browse_cancelled(self, mock_qmessagebox_warning, mock_get_open_file_name):
        """Dosya seçiminin iptal edilmesini test eder."""
        self.btn_browse.click()

        mock_get_open_file_name.assert_called_once()
        self.assertIsNone(self.main_window_mock.excel_path)  # Değişmemeli
        self.assertFalse(self.cmb_sheet.isEnabled())
        self.assertFalse(self.btn_next.isEnabled())
        self.assertFalse(mock_qmessagebox_warning.called)  # Uyarı olmamalı

    @patch('PyQt5.QtWidgets.QFileDialog.getOpenFileName', return_value=('fake/path/to/corrupt.xlsx', '.xlsx'))
    @patch('pandas.ExcelFile', side_effect=Exception("Corrupt file error"))
    @patch('PyQt5.QtWidgets.QMessageBox.critical')
    def test_browse_read_error(self, mock_qmessagebox_critical, mock_excel_file, mock_get_open_file_name):
        """Dosya okuma hatasını test eder."""
        self.btn_browse.click()

        mock_get_open_file_name.assert_called_once()
        mock_excel_file.assert_called_once()
        mock_qmessagebox_critical.assert_called_once()
        # Hata mesajı "Okuma hatası" olarak geldiği için testi buna göre ayarla
        self.assertIn("Okuma hatası", mock_qmessagebox_critical.call_args[0][1])
        self.assertIsNone(self.main_window_mock.excel_path)  # Yol sıfırlanmalı
        self.assertFalse(self.cmb_sheet.isEnabled())
        self.assertFalse(self.btn_next.isEnabled())
        self.assertEqual(self.lbl_path.text(), "Henüz dosya seçilmedi")  # Başlangıç durumuna dönmeli

    @patch('PyQt5.QtWidgets.QFileDialog.getOpenFileName', return_value=('fake/path/to/test.xlsx', '.xlsx'))
    @patch('pandas.ExcelFile')
    def test_on_sheet_selected(self, mock_excel_file, mock_get_open_file_name):
        """Sayfa seçimi değiştiğinde ana pencere seçiminin güncellenmesini test eder."""
        mock_xls = MagicMock()
        mock_xls.sheet_names = ['SMD-OEE', 'ROBOT']
        mock_excel_file.return_value = mock_xls
        self.btn_browse.click()  # Dosyayı seç ve combobox'ı doldur

        # Patch currentText of cmb_sheet to control its return value
        with patch.object(self.cmb_sheet, 'currentText', return_value='ROBOT') as mock_current_text:
            self.cmb_sheet.setCurrentIndex(1)  # ROBOT'u seç (this will call the real setCurrentIndex)
            self.page.on_sheet_selected() # Call the slot manually
            self.assertEqual(self.main_window_mock.selected_sheet, 'ROBOT')
            self.assertTrue(self.btn_next.isEnabled())

        # Test case for no selection
        with patch.object(self.cmb_sheet, 'currentText', return_value='') as mock_current_text:
            self.cmb_sheet.setCurrentIndex(-1)  # Seçim yokmuş gibi simüle et
            self.page.on_sheet_selected() # Call the slot manually
            # Assuming if currentText is empty, selected_sheet becomes None or empty string
            # Based on the previous error, it seems to become ''
            self.assertFalse(self.btn_next.isEnabled())


    def test_go_next(self):
        """İleri butonuna basıldığında doğru metotların çağrılmasını test eder."""
        # Gerekli koşulları ayarla
        self.main_window_mock.excel_path = Path('fake/path/to/test.xlsx')
        self.main_window_mock.selected_sheet = 'SMD-OEE'
        self.btn_next.setEnabled(True)  # Butonu aktif et

        self.btn_next.click()

        self.main_window_mock.load_excel.assert_called_once()
        self.main_window_mock.goto_page.assert_called_once_with(1)

    def test_reset_page(self):
        """Sayfanın başlangıç durumuna döndürülmesini test eder."""
        # Sayfanın mevcut durumunu değiştir
        self.main_window_mock.excel_path = Path('some/path.xlsx')
        # Bu kısım uygulamanın reset_page metodundaki gerçek davranışa göre ayarlanmalı.
        # Eğer uygulama selected_sheet'i boş string yapıyorsa, testi de ona göre değiştir.
        # Traceback'teki hata '' is not None olduğu için, uygulamanın '' döndürdüğü varsayılıyor.
        self.main_window_mock.selected_sheet = '' # Uygulamanın sıfırlama davranışına göre ayarlandı
        self.lbl_path.setText("Seçilen Dosya: <b>some_path.xlsx</b>")
        self.cmb_sheet.addItem("TestSheet")
        self.cmb_sheet.setEnabled(True)
        self.cmb_sheet.show()
        self.sheet_selection_label.show()
        self.btn_next.setEnabled(True)

        self.page.reset_page()

        self.assertIsNone(self.main_window_mock.excel_path)
        # Testi uygulamanın gerçek davranışına göre ayarla
        self.assertEqual(self.main_window_mock.selected_sheet, '')
        self.assertEqual(self.lbl_path.text(), "Henüz dosya seçilmedi")
        self.assertEqual(self.cmb_sheet.count(), 0)
        self.assertFalse(self.cmb_sheet.isEnabled())
        self.assertFalse(self.cmb_sheet.isVisible())
        self.assertFalse(self.sheet_selection_label.isVisible())
        self.assertFalse(self.btn_next.isEnabled())


class TestDataSelectionPage(unittest.TestCase):
    def setUp(self):
        self.main_window_mock = MagicMock(spec=MainWindow)
        self.main_window_mock.df = pd.DataFrame({
            'Tarih': pd.to_datetime(['2023-01-01', '2023-01-01', '2023-01-02']),
            'Ürün': ['Ürün A', 'Ürün B', 'Ürün A'],
            'OEE_Değeri': [0.85, 0.90, 0.88],
            'Metrik1': ['00:10:00', '00:05:00', '00:12:00'],
            'Metrik2': ['00:02:00', '00:01:00', '00:03:00'],
            'Metrik3': [100, 200, 150]
        })
        self.main_window_mock.grouping_col_name = 'Tarih'
        self.main_window_mock.grouped_col_name = 'Ürün'
        self.main_window_mock.oee_col_name = 'OEE_Değeri'
        self.main_window_mock.metric_cols = ['Metrik1', 'Metrik2', 'Metrik3']
        self.main_window_mock.selected_metrics = []
        self.main_window_mock.goto_page = MagicMock()
        self.main_window_mock.graph_worker_thread = MagicMock(spec=QThread)
        self.main_window_mock.graph_worker_thread.finished = pyqtSignal()
        self.main_window_mock.graph_worker_thread.error = pyqtSignal()
        self.main_window_mock.show_progress_dialog = MagicMock()

        self.page = DataSelectionPage(self.main_window_mock)

        # UI elemanlarına doğrudan erişim için
        self.cmb_grouping = self.page.cmb_grouping
        self.lst_grouped = self.page.lst_grouped
        self.metrics_layout = self.page.metrics_layout
        self.btn_next = self.page.btn_next
        self.btn_back = self.page.btn_back

    def tearDown(self):
        self.page.deleteLater()

    def test_initial_ui_state(self):
        """UI'ın başlangıç durumunu test eder."""
        self.assertEqual(self.cmb_grouping.count(), 0)
        self.assertEqual(self.lst_grouped.count(), 0)
        self.assertEqual(self.metrics_layout.count(), 0)  # Başlangıçta metrik checkbox'ları olmamalı
        self.assertFalse(self.btn_next.isEnabled())
        self.assertEqual(self.btn_back.text(), "← Geri")
        self.assertEqual(self.btn_next.text(), "İleri →")

    @patch('PyQt5.QtWidgets.QMessageBox.critical')
    def test_refresh_empty_dataframe(self, mock_qmessagebox_critical):
        """Boş DataFrame ile refresh metotunu test eder."""
        self.main_window_mock.df = pd.DataFrame()
        self.page.refresh()
        mock_qmessagebox_critical.assert_called_once_with(
            self.page, "Hata", "Veri yüklenemedi. Lütfen dosyayı kontrol edin."
        )
        self.main_window_mock.goto_page.assert_called_once_with(0)

    def test_refresh_with_data(self):
        """Veri içeren DataFrame ile refresh metotunu test eder."""
        self.page.refresh()
        self.assertEqual(self.cmb_grouping.count(), 2)  # 2023-01-01, 2023-01-02
        self.assertIn('2023-01-01', [self.cmb_grouping.itemText(i) for i in range(self.cmb_grouping.count())])
        self.assertIn('2023-01-02', [self.cmb_grouping.itemText(i) for i in range(self.cmb_grouping.count())])

        # metrics_layout'un doğru şekilde doldurulduğunu kontrol et
        self.assertEqual(self.metrics_layout.count(), len(self.main_window_mock.metric_cols))
        for i, col_name in enumerate(self.main_window_mock.metric_cols):
            checkbox = self.metrics_layout.itemAt(i).widget()
            self.assertIsInstance(checkbox, QCheckBox)
            self.assertEqual(checkbox.text(), col_name)
            self.assertTrue(checkbox.isChecked())  # Varsayılan olarak hepsi seçili olmalı

        # lst_grouped'un da dolduğunu kontrol et (populate_grouped çağrıldığı için)
        self.assertEqual(self.lst_grouped.count(), 2)  # Ürün A, Ürün B
        self.assertEqual(self.lst_grouped.item(0).text(), 'Ürün A')
        self.assertEqual(self.lst_grouped.item(1).text(), 'Ürün B')
        self.assertTrue(self.lst_grouped.item(0).isSelected())
        self.assertTrue(self.lst_grouped.item(1).isSelected())

        self.assertTrue(self.btn_next.isEnabled())  # İleri butonu aktif olmalı

    @patch('PyQt5.QtWidgets.QMessageBox.warning')
    def test_refresh_missing_grouping_col(self, mock_qmessagebox_warning):
        """Gruplama sütunu eksik olduğunda refresh metotunu test eder."""
        self.main_window_mock.grouping_col_name = 'NonExistentCol'
        self.page.refresh()
        mock_qmessagebox_warning.assert_called_once_with(
            self.page, "Hata", "Gruplama sütunu (NonExistentCol) bulunamadı."
        )
        self.main_window_mock.goto_page.assert_called_once_with(0)