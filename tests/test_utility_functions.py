# test_utility_functions.py
import unittest
import pandas as pd
import numpy as np
import datetime
from unittest.mock import MagicMock, patch

# Assuming GraficApplication.py is in the same directory or accessible via PYTHONPATH
from GraficApplication import excel_col_to_index, seconds_from_timedelta

class TestUtilityFunctions(unittest.TestCase):
    def test_excel_col_to_index(self):
        """Excel sütun adlarını indekslere dönüştürme fonksiyonunu test eder."""
        self.assertEqual(excel_col_to_index("A"), 0)
        self.assertEqual(excel_col_to_index("B"), 1)
        self.assertEqual(excel_col_to_index("Z"), 25)
        self.assertEqual(excel_col_to_index("AA"), 26)
        self.assertEqual(excel_col_to_index("AZ"), 51)
        self.assertEqual(excel_col_to_index("BA"), 52)
        self.assertEqual(excel_col_to_index("BD"), 55)
        self.assertEqual(excel_col_to_index("a"), 0)  # Küçük harf girişi
        self.assertEqual(excel_col_to_index("zZ"), 701) # Karışık büyük/küçük harf

    def test_excel_col_to_index_invalid_input(self):
        """excel_col_to_index fonksiyonunun geçersiz girdileri işlemesini test eder."""
        with self.assertRaises(ValueError):
            excel_col_to_index("1")  # Sayısal
        with self.assertRaises(ValueError):
            excel_col_to_index("A1") # Karışık harf-sayı
        with self.assertRaises(ValueError):
            excel_col_to_index("!")  # Özel karakter
        with self.assertRaises(ValueError):
            excel_col_to_index("")   # Boş string
        with self.assertRaises(ValueError):
            excel_col_to_index(None) # None değeri
        with self.assertRaises(ValueError):
            excel_col_to_index("ABC1") # Sayı içeren string
        with self.assertRaises(ValueError):
            excel_col_to_index(" AB") # Boşluk içeren string


    def test_seconds_from_timedelta(self):
        """Zaman dilimi stringlerini saniyeye dönüştürme fonksiyonunu test eder."""
        self.assertEqual(seconds_from_timedelta("00:00:00"), 0)
        self.assertEqual(seconds_from_timedelta("00:00:01"), 1)
        self.assertEqual(seconds_from_timedelta("00:01:00"), 60)
        self.assertEqual(seconds_from_timedelta("01:00:00"), 3600)
        self.assertEqual(seconds_from_timedelta("00:10:30"), 630) # 10 dakika 30 saniye
        self.assertEqual(seconds_from_timedelta("23:59:59"), 86399)
        self.assertEqual(seconds_from_timedelta("1:2:3"), 3723) # Tek basamaklı saat, dakika, saniye

        # MM:SS formatı
        self.assertEqual(seconds_from_timedelta("05:30"), 330)
        self.assertEqual(seconds_from_timedelta("1:05"), 65)

        # Farklı ayraçlar
        self.assertEqual(seconds_from_timedelta("01-00-00", separator="-"), 3600)

    def test_seconds_from_timedelta_invalid_input(self):
        """seconds_from_timedelta fonksiyonunun geçersiz girdileri işlemesini test eder."""
        # Geçersiz formatlar
        self.assertEqual(seconds_from_timedelta("100"), 0) # Sadece saniye
        self.assertEqual(seconds_from_timedelta("10:00:00:00"), 0) # Çok fazla bölüm
        self.assertEqual(seconds_from_timedelta("abc:def:ghi"), 0) # Sayısal olmayan karakterler
        self.assertEqual(seconds_from_timedelta("10-20"), 0) # Yanlış ayraç
        self.assertEqual(seconds_from_timedelta(""), 0) # Boş string
        self.assertEqual(seconds_from_timedelta(None), 0) # None
        self.assertEqual(seconds_from_timedelta("00:00:"), 0) # Eksik saniye

    def test_seconds_from_timedelta_nan_or_empty(self):
        """seconds_from_timedelta fonksiyonunun NaN veya boş değerleri işlemesini test eder."""
        self.assertEqual(seconds_from_timedelta(np.nan), 0) # NumPy NaN
        self.assertEqual(seconds_from_timedelta(""), 0) # Boş string
        self.assertEqual(seconds_from_timedelta(" "), 0) # Sadece boşluk içeren string
        self.assertEqual(seconds_from_timedelta(None), 0) # None değeri
        self.assertEqual(seconds_from_timedelta(123), 0) # Sayısal olmayan string girdisi bekleniyor

if __name__ == '__main__':
    unittest.main(argv=['first-arg-is-ignored'], exit=False)