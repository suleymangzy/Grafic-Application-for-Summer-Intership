# test_graph_worker.py
import unittest
import pandas as pd
import numpy as np
import datetime
from unittest.mock import MagicMock, patch
from PyQt5.QtCore import QThread, pyqtSignal, QObject

# Assuming GraficApplication.py is in the same directory or accessible via PYTHONPATH
from GraficApplication import GraphWorker, excel_col_to_index, seconds_from_timedelta, REQ_SHEETS


class TestGraphWorker(unittest.TestCase):
    def setUp(self):
        # Örnek DataFrame oluştur
        self.sample_df = pd.DataFrame({
            'Tarih': pd.to_datetime(['2023-01-01', '2023-01-01', '2023-01-02', '2023-01-02']),
            'Ürün Adı': ['Ürün A', 'Ürün B', 'Ürün A', 'Ürün C'],
            'OEE %': [90.5, 88.0, 92.1, 75.0],
            'DUR_BEKLENEN': ['00:10:00', '00:05:00', '00:12:00', '00:08:00'],
            'DUR_PLANSİZ': ['00:02:00', '00:01:00', '00:03:00', '00:02:00'],
            'BOŞ_ÇALIŞMA': ['00:01:00', '00:00:30', '00:01:30', '00:01:00'],
            'Üretim Miktarı': [100, 150, 120, 80],
            'Hurda Miktarı': [5, 10, 8, 3]
        })
        self.grouping_col = 'Tarih'
        self.grouped_col = 'Ürün Adı'
        self.oee_col = 'OEE %'
        self.metric_cols = ['DUR_BEKLENEN', 'DUR_PLANSİZ', 'BOŞ_ÇALIŞMA']
        self.selected_grouping_val = pd.to_datetime('2023-01-01')
        self.selected_grouped_values = ['Ürün A', 'Ürün B']
        self.selected_metrics = ['DUR_BEKLENEN', 'BOŞ_ÇALIŞMA']  # Sadece belirli metrikleri seçelim

    @patch.object(GraphWorker, 'finished', new_callable=pyqtSignal)
    @patch.object(GraphWorker, 'error', new_callable=pyqtSignal)
    def test_run_method_success(self, mock_error_signal, mock_finished_signal):
        """GraphWorker'ın run metodunun başarılı çalışmasını test eder."""
        worker = GraphWorker(
            self.sample_df,
            self.grouping_col,
            self.grouped_col,
            self.selected_grouped_values,
            self.selected_metrics,
            self.oee_col,
            self.selected_grouping_val
        )

        # Sinyalleri dinlemek için bir mock object oluştur
        mock_receiver_finished = MagicMock()
        mock_receiver_error = MagicMock()
        worker.finished.connect(mock_receiver_finished)
        worker.error.connect(mock_receiver_error)

        worker.run()

        mock_error_signal.emit.assert_not_called()
        mock_finished_signal.emit.assert_called_once()

        # finished sinyalinin yayınladığı argümanları kontrol et
        emitted_results = mock_finished_signal.emit.call_args[0][0]
        self.assertIsInstance(emitted_results, list)
        self.assertEqual(len(emitted_results), len(self.selected_grouped_values))

        # İlk ürün için beklenen sonuçları kontrol et ('Ürün A')
        # 'Tarih' 2023-01-01, 'Ürün Adı' 'Ürün A' olan satır
        # 'DUR_BEKLENEN': '00:10:00' -> 600 saniye
        # 'BOŞ_ÇALIŞMA': '00:01:00' -> 60 saniye
        # Toplam metrik: 660 saniye
        # OEE: 90.5%
        product_a_result = emitted_results[0]
        self.assertEqual(product_a_result[0], 'Ürün A')
        self.assertIsInstance(product_a_result[1], pd.Series)
        self.assertEqual(product_a_result[1]['DUR_BEKLENEN'], 600)
        self.assertEqual(product_a_result[1]['BOŞ_ÇALIŞMA'], 60)
        self.assertEqual(product_a_result[2], '90.50%')

        # İkinci ürün için beklenen sonuçları kontrol et ('Ürün B')
        # 'Tarih' 2023-01-01, 'Ürün Adı' 'Ürün B' olan satır
        # 'DUR_BEKLENEN': '00:05:00' -> 300 saniye
        # 'BOŞ_ÇALIŞMA': '00:00:30' -> 30 saniye
        # Toplam metrik: 330 saniye
        # OEE: 88.0%
        product_b_result = emitted_results[1]
        self.assertEqual(product_b_result[0], 'Ürün B')
        self.assertIsInstance(product_b_result[1], pd.Series)
        self.assertEqual(product_b_result[1]['DUR_BEKLENEN'], 300)
        self.assertEqual(product_b_result[1]['BOŞ_ÇALIŞMA'], 30)
        self.assertEqual(product_b_result[2], '88.00%')

    @patch.object(GraphWorker, 'finished', new_callable=pyqtSignal)
    @patch.object(GraphWorker, 'error', new_callable=pyqtSignal)
    def test_run_method_no_grouped_values(self, mock_error_signal, mock_finished_signal):
        """run metodunun seçili gruplanmış değerler olmadığında hata yayınlamasını test eder."""
        worker = GraphWorker(
            self.sample_df,
            self.grouping_col,
            self.grouped_col,
            [],  # Boş gruplanmış değerler
            self.selected_metrics,
            self.oee_col,
            self.selected_grouping_val
        )
        worker.run()
        mock_finished_signal.emit.assert_not_called()
        mock_error_signal.emit.assert_called_once_with("Lütfen en az bir gruplanmış değer seçin.")

    @patch.object(GraphWorker, 'finished', new_callable=pyqtSignal)
    @patch.object(GraphWorker, 'error', new_callable=pyqtSignal)
    def test_run_method_no_metrics_selected(self, mock_error_signal, mock_finished_signal):
        """run metodunun seçili metrikler olmadığında hata yayınlamasını test eder."""
        worker = GraphWorker(
            self.sample_df,
            self.grouping_col,
            self.grouped_col,
            self.selected_grouped_values,
            [],  # Boş metrikler
            self.oee_col,
            self.selected_grouping_val
        )
        worker.run()
        mock_finished_signal.emit.assert_not_called()
        mock_error_signal.emit.assert_called_once_with("Lütfen en az bir metrik seçin.")

    @patch.object(GraphWorker, 'finished', new_callable=pyqtSignal)
    @patch.object(GraphWorker, 'error', new_callable=pyqtSignal)
    def test_run_method_oee_col_missing(self, mock_error_signal, mock_finished_signal):
        """run metodunun OEE sütunu eksik olduğunda hata yayınlamasını test eder."""
        worker = GraphWorker(
            self.sample_df.drop(columns=[self.oee_col]),  # OEE sütununu kaldır
            self.grouping_col,
            self.grouped_col,
            self.selected_grouped_values,
            self.selected_metrics,
            self.oee_col,  # Hala eksik sütun adı belirtiliyor
            self.selected_grouping_val
        )
        worker.run()
        mock_finished_signal.emit.assert_not_called()
        mock_error_signal.emit.assert_called_once_with("OEE sütunu (OEE %) bulunamadı.")

    def test_process_data_calculates_oee_correctly(self):
        """_process_data metodunun OEE'yi doğru hesaplamasını test eder."""
        worker = GraphWorker(
            self.sample_df, self.grouping_col, self.grouped_col,
            self.selected_grouped_values, self.metric_cols, self.oee_col,
            self.selected_grouping_val
        )

        # 'Ürün A' için OEE: 90.5%
        oee_value_a = worker._calculate_oee(self.sample_df[
                                                (self.sample_df[self.grouping_col] == self.selected_grouping_val) &
                                                (self.sample_df[self.grouped_col] == 'Ürün A')
                                                ][self.oee_col])
        self.assertEqual(oee_value_a, "90.50%")

        # 'Ürün B' için OEE: 88.0%
        oee_value_b = worker._calculate_oee(self.sample_df[
                                                (self.sample_df[self.grouping_col] == self.selected_grouping_val) &
                                                (self.sample_df[self.grouped_col] == 'Ürün B')
                                                ][self.oee_col])
        self.assertEqual(oee_value_b, "88.00%")

        # Boş OEE değerleri için
        empty_oee = worker._calculate_oee(pd.Series([]))
        self.assertEqual(empty_oee, "N/A")

        # NaN OEE değerleri için
        nan_df = pd.DataFrame({'OEE %': [np.nan, 80.0, np.nan]})
        nan_oee = worker._calculate_oee(nan_df['OEE %'])
        self.assertEqual(nan_oee, "80.00%")

        # Tamamen NaN OEE değerleri için
        all_nan_df = pd.DataFrame({'OEE %': [np.nan, np.nan]})
        all_nan_oee = worker._calculate_oee(all_nan_df['OEE %'])
        self.assertEqual(all_nan_oee, "N/A")

    def test_process_data_handles_missing_metric_data(self):
        """_process_data metodunun eksik metrik verilerini (NaN veya boş string) işlemesini test eder."""
        df_with_nan_metrics = self.sample_df.copy()
        df_with_nan_metrics.loc[0, 'DUR_BEKLENEN'] = np.nan  # NaN değer
        df_with_nan_metrics.loc[1, 'BOŞ_ÇALIŞMA'] = ''  # Boş string değer

        worker = GraphWorker(
            df_with_nan_metrics, self.grouping_col, self.grouped_col,
            self.selected_grouped_values, self.metric_cols, self.oee_col,
            self.selected_grouping_val
        )

        # Sinyalleri dinlemek için bir mock object oluştur
        mock_receiver_finished = MagicMock()
        worker.finished.connect(mock_receiver_finished)
        worker.run()

        emitted_results = mock_receiver_finished.call_args[0][0]

        # Ürün A için (DUR_BEKLENEN NaN)
        # DUR_BEKLENEN: NaN -> 0
        # DUR_PLANSİZ: 00:02:00 -> 120
        # BOŞ_ÇALIŞMA: 00:01:00 -> 60
        # Toplam: 180
        product_a_result = emitted_results[0]
        self.assertEqual(product_a_result[0], 'Ürün A')
        self.assertEqual(product_a_result[1]['DUR_BEKLENEN'], 0)  # NaN olarak gelir ama toplamda 0 olarak alınmalı
        self.assertEqual(product_a_result[1]['DUR_PLANSİZ'], 120)
        self.assertEqual(product_a_result[1]['BOŞ_ÇALIŞMA'], 60)
        self.assertEqual(product_a_result[1].sum(), 180)  # Sadece seçili metrikleri toplar

        # Ürün B için (BOŞ_ÇALIŞMA boş string)
        # DUR_BEKLENEN: 00:05:00 -> 300
        # DUR_PLANSİZ: 00:01:00 -> 60
        # BOŞ_ÇALIŞMA: '' -> 0
        # Toplam: 360
        product_b_result = emitted_results[1]
        self.assertEqual(product_b_result[0], 'Ürün B')
        self.assertEqual(product_b_result[1]['DUR_BEKLENEN'], 300)
        self.assertEqual(product_b_result[1]['DUR_PLANSİZ'], 60)
        self.assertEqual(product_b_result[1]['BOŞ_ÇALIŞMA'],
                         0)  # Boş string olarak gelir ama toplamda 0 olarak alınmalı
        self.assertEqual(product_b_result[1].sum(), 360)  # Sadece seçili metrikleri toplar

    def test_process_data_handles_timedelta_conversion(self):
        """_process_data metodunun timedelta stringlerini doğru bir şekilde saniyeye çevirdiğini test eder."""
        df_timedelta = pd.DataFrame({
            'Tarih': pd.to_datetime(['2023-01-01']),
            'Ürün Adı': ['Test Ürün'],
            'OEE %': [95.0],
            'Metrik_Saat': ['01:00:00'],  # 3600 saniye
            'Metrik_Dakika': ['00:30:00'],  # 1800 saniye
            'Metrik_Saniye': ['00:00:15']  # 15 saniye
        })

        worker = GraphWorker(
            df_timedelta, 'Tarih', 'Ürün Adı', ['Test Ürün'],
            ['Metrik_Saat', 'Metrik_Dakika', 'Metrik_Saniye'],
            'OEE %', pd.to_datetime('2023-01-01')
        )

        mock_receiver_finished = MagicMock()
        worker.finished.connect(mock_receiver_finished)
        worker.run()

        emitted_results = mock_receiver_finished.call_args[0][0]
        result = emitted_results[0]

        self.assertEqual(result[0], 'Test Ürün')
        self.assertEqual(result[1]['Metrik_Saat'], 3600)
        self.assertEqual(result[1]['Metrik_Dakika'], 1800)
        self.assertEqual(result[1]['Metrik_Saniye'], 15)

    def test_process_data_with_non_timedelta_metrics(self):
        """_process_data metodunun sayısal metrikleri de doğru işlediğini test eder."""
        df_numeric_metrics = pd.DataFrame({
            'Tarih': pd.to_datetime(['2023-01-01']),
            'Ürün Adı': ['Sayisal Ürün'],
            'OEE %': [90.0],
            'Adet': [100],
            'Ağırlık': [50.5]
        })

        worker = GraphWorker(
            df_numeric_metrics, 'Tarih', 'Ürün Adı', ['Sayisal Ürün'],
            ['Adet', 'Ağırlık'],
            'OEE %', pd.to_datetime('2023-01-01')
        )

        mock_receiver_finished = MagicMock()
        worker.finished.connect(mock_receiver_finished)
        worker.run()

        emitted_results = mock_receiver_finished.call_args[0][0]
        result = emitted_results[0]

        self.assertEqual(result[0], 'Sayisal Ürün')
        self.assertEqual(result[1]['Adet'], 100)
        self.assertEqual(result[1]['Ağırlık'], 50.5)
        self.assertEqual(result[2], '90.00%')

    def test_process_data_empty_filtered_data(self):
        """_process_data metodunun filtrelenmiş veri boş olduğunda boş seri döndürmesini test eder."""
        # Seçili gruplama değeri ile eşleşmeyen bir DataFrame oluştur
        empty_filtered_df = self.sample_df.copy()
        empty_filtered_df['Tarih'] = pd.to_datetime('2024-01-01')  # Seçili tarihten farklı

        worker = GraphWorker(
            empty_filtered_df, self.grouping_col, self.grouped_col,
            ['Ürün A'], self.metric_cols, self.oee_col,
            self.selected_grouping_val  # 2023-01-01
        )

        mock_receiver_finished = MagicMock()
        worker.finished.connect(mock_receiver_finished)
        worker.run()

        emitted_results = mock_receiver_finished.call_args[0][0]
        self.assertEqual(len(emitted_results), 0)  # Hiçbir ürün için veri işlenmemeli


if __name__ == '__main__':
    unittest.main(argv=['first-arg-is-ignored'], exit=False)