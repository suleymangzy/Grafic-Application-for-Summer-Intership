# test_graph_plotter.py
import unittest
import pandas as pd
import numpy as np
import datetime
from unittest.mock import MagicMock, patch, call

import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.axes import Axes

# Assuming GraficApplication.py is in the same directory or accessible via PYTHONPATH
from GraficApplication import GraphPlotter


class TestGraphPlotter(unittest.TestCase):
    def setUp(self):
        self.figure_mock = MagicMock(spec=Figure)
        self.ax_mock = MagicMock(spec=Axes)
        # Mock ax.pie and ax.barh to return a list of patches (expected by the real methods)
        self.ax_mock.pie.return_value = ([MagicMock()], [MagicMock()], [MagicMock()])  # wedges, texts, autotexts
        self.ax_mock.barh.return_value = [MagicMock()]  # patches

        # OEE değeri
        self.oee_value = "90.50%"
        # Metrik örnekleri (saniye cinsinden)
        self.sample_metrics = pd.Series(
            {
                "Metrik1": 600,  # 10 dakika
                "Metrik2": 300,  # 5 dakika
                "Metrik3": 120,  # 2 dakika
                "Metrik4": 60,  # 1 dakika
            }
        )
        self.chart_colors = ["#FF6347", "#4682B4", "#8A2BE2", "#3CB371"]  # Örnek renkler

    def test_create_donut_chart(self):
        """Halka grafiğin (donut chart) doğru çizildiğini test eder."""
        GraphPlotter.create_donut_chart(
            self.ax_mock, self.sample_metrics, self.oee_value, self.chart_colors, self.figure_mock
        )

        self.ax_mock.pie.assert_called_once()
        args, kwargs = self.ax_mock.pie.call_args

        # values
        np.testing.assert_array_equal(args[0], self.sample_metrics.values)

        # labels
        self.assertListEqual(list(kwargs["labels"]), self.sample_metrics.index.tolist())

        # colors
        self.assertListEqual(kwargs["colors"], self.chart_colors)

        # autopct format
        self.assertEqual(kwargs["autopct"], "%1.1f%%")

        # wedgeprops
        self.assertIn("wedgeprops", kwargs)
        self.assertEqual(kwargs["wedgeprops"]["width"], 0.3)
        self.assertEqual(kwargs["wedgeprops"]["edgecolor"], "w")

        # startangle
        self.assertEqual(kwargs["startangle"], 90)

        self.ax_mock.axis.assert_called_once_with("equal")
        self.ax_mock.set_title.assert_called_once_with(
            f"OEE: {self.oee_value}", fontsize=16, fontweight="bold"
        )

    def test_create_donut_chart_empty_metrics(self):
        """Boş metrik serisi ile halka grafiğin çizimini test eder."""
        empty_metrics = pd.Series({})
        GraphPlotter.create_donut_chart(
            self.ax_mock, empty_metrics, self.oee_value, self.chart_colors, self.figure_mock
        )

        self.ax_mock.pie.assert_called_once()
        args, kwargs = self.ax_mock.pie.call_args
        np.testing.assert_array_equal(args[0], [])  # Boş değerler gönderilmeli
        self.ax_mock.set_title.assert_called_once_with(
            f"OEE: {self.oee_value} (Veri Yok)", fontsize=16, fontweight="bold"
        )
        self.ax_mock.text.assert_called_once()  # "Veri Yok" mesajı olmalı

    def test_create_donut_chart_single_metric(self):
        """Tek metrik ile halka grafiğin çizimini test eder."""
        single_metric = pd.Series({"Tek Metrik": 100})
        GraphPlotter.create_donut_chart(
            self.ax_mock, single_metric, self.oee_value, ["red"], self.figure_mock
        )

        self.ax_mock.pie.assert_called_once()
        args, kwargs = self.ax_mock.pie.call_args
        np.testing.assert_array_equal(args[0], [100])
        self.assertListEqual(list(kwargs["labels"]), ["Tek Metrik"])

    def test_create_bar_chart(self):
        """Çubuk grafiğin doğru çizildiğini test eder."""
        GraphPlotter.create_bar_chart(
            self.ax_mock, self.sample_metrics, self.oee_value, self.chart_colors
        )

        self.ax_mock.barh.assert_called_once()
        args, kwargs = self.ax_mock.barh.call_args

        # y_pos
        self.assertEqual(len(args[0]), len(self.sample_metrics))
        # values_minutes (saniyeden dakikaya çevrilmeli)
        self.assertListEqual([v / 60 for v in self.sample_metrics.values], list(args[1]))
        # colors
        self.assertEqual(kwargs["color"], self.chart_colors)

        self.ax_mock.set_yticks.assert_called_once()
        self.ax_mock.set_yticklabels.assert_called_once_with(
            self.sample_metrics.index.tolist(), fontsize=10
        )
        self.ax_mock.invert_yaxis.assert_called_once()
        self.ax_mock.set_xlabel.assert_called_once_with("Dakika")  # X ekseni etiketi "Dakika" olmalı
        self.ax_mock.set_xticks.assert_called_once_with([])  # X ekseninde tik olmamalı
        self.ax_mock.set_title.assert_called_once_with(
            f"OEE: {self.oee_value}", fontsize=16, fontweight="bold"
        )
        self.ax_mock.grid.assert_called_once_with(False)

        # Çerçeve çizgilerinin gizlenmesi
        self.ax_mock.spines["right"].set_visible.assert_called_once_with(False)
        self.ax_mock.spines["top"].set_visible.assert_called_once_with(False)
        self.ax_mock.spines["left"].set_visible.assert_called_once_with(True)  # Sol görünür kalmalı
        self.ax_mock.spines["bottom"].set_visible.assert_called_once_with(True)  # Alt görünür kalmalı

    def test_create_bar_chart_empty_metrics(self):
        """Boş metrik serisi ile çubuk grafiğin çizimini test eder."""
        empty_metrics = pd.Series({})
        GraphPlotter.create_bar_chart(
            self.ax_mock, empty_metrics, self.oee_value, self.chart_colors
        )

        self.ax_mock.barh.assert_called_once()
        args, kwargs = self.ax_mock.barh.call_args
        np.testing.assert_array_equal(args[0], [])  # Boş değerler gönderilmeli
        np.testing.assert_array_equal(args[1], [])  # Boş değerler gönderilmeli
        self.ax_mock.set_title.assert_called_once_with(
            f"OEE: {self.oee_value} (Veri Yok)", fontsize=16, fontweight="bold"
        )
        self.ax_mock.text.assert_called_once()  # "Veri Yok" mesajı olmalı


if __name__ == '__main__':
    unittest.main(argv=['first-arg-is-ignored'], exit=False)