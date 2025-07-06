import unittest
import pandas as pd
import numpy as np
import datetime
from unittest.mock import MagicMock, patch

# Assuming GraficApplication.py is in the same directory or accessible via PYTHONPATH
from GraficApplication import (
    excel_col_to_index,
    seconds_from_timedelta,
    GraphWorker,
    GraphPlotter,
)


class TestUtilityFunctions(unittest.TestCase):
    def test_excel_col_to_index(self):
        self.assertEqual(excel_col_to_index("A"), 0)
        self.assertEqual(excel_col_to_index("B"), 1)
        self.assertEqual(excel_col_to_index("Z"), 25)
        self.assertEqual(excel_col_to_index("AA"), 26)
        self.assertEqual(excel_col_to_index("AZ"), 51)
        self.assertEqual(excel_col_to_index("BA"), 52)
        self.assertEqual(excel_col_to_index("BD"), 55)

        with self.assertRaises(ValueError):
            excel_col_to_index("1")
        with self.assertRaises(ValueError):
            excel_col_to_index("A1")
        # 'a' test case was removed in the previous step, as it's a valid column after uppercasing.
        # Ensure your test.py no longer contains:
        # with self.assertRaises(ValueError):
        #    excel_col_to_index("a")

    def test_seconds_from_timedelta(self):
        # Test timedelta objects
        series_timedelta = pd.Series(
            [
                pd.Timedelta(hours=1, minutes=30, seconds=15),
                pd.Timedelta(minutes=5),
                pd.Timedelta(seconds=45),
                pd.NaT,
            ]
        )
        expected_timedelta = pd.Series([5415.0, 300.0, 45.0, 0.0])
        pd.testing.assert_series_equal(
            seconds_from_timedelta(series_timedelta), expected_timedelta
        )

        # Test datetime.time objects
        series_time = pd.Series(
            [
                datetime.time(1, 30, 15),
                datetime.time(0, 5, 0),
                datetime.time(0, 0, 45),
                None,
            ]
        )
        expected_time = pd.Series([5415.0, 300.0, 45.0, 0.0])
        pd.testing.assert_series_equal(
            seconds_from_timedelta(series_time), expected_time
        )

        # Test string representations
        series_str = pd.Series(
            ["01:30:15", "0:05:00", "00:00:45", "invalid", "1:2:3", ""]
        )
        expected_str = pd.Series([5415.0, 300.0, 45.0, 0.0, 3723.0, 0.0])
        pd.testing.assert_series_equal(seconds_from_timedelta(series_str), expected_str)

        # Test numeric values (interpreted as days from Excel)
        series_numeric = pd.Series([1.5, 0.25, 0.0, np.nan])
        # 1.5 days = 1.5 * 24 * 3600 = 129600 seconds
        # 0.25 days = 0.25 * 24 * 3600 = 21600 seconds
        expected_numeric = pd.Series([129600.0, 21600.0, 0.0, 0.0])
        pd.testing.assert_series_equal(
            seconds_from_timedelta(series_numeric), expected_numeric
        )

        # Test mixed types
        series_mixed = pd.Series(
            [
                pd.Timedelta(seconds=10),
                "00:00:20",
                datetime.time(0, 0, 30),
                1.0 / 24 / 60,
                np.nan,
                "invalid",
            ]
        )  # 1.0 / 24 / 60 = 1 minute in days
        expected_mixed = pd.Series([10.0, 20.0, 30.0, 60.0, 0.0, 0.0])
        pd.testing.assert_series_equal(
            seconds_from_timedelta(series_mixed), expected_mixed
        )


class TestGraphWorker(unittest.TestCase):
    def setUp(self):
        # Create a sample DataFrame for testing
        self.df = pd.DataFrame(
            {
                "A": ["2023-01-01", "2023-01-01", "2023-01-02", "2023-01-02"],
                "B": ["ProductX", "ProductY", "ProductX", "ProductY"],
                "H": [
                    "00:30:00",
                    "00:15:00",
                    "00:45:00",
                    "00:20:00",
                ],  # Metric 1
                "I": [
                    pd.Timedelta(minutes=10),
                    pd.Timedelta(minutes=5),
                    pd.Timedelta(minutes=20),
                    pd.Timedelta(minutes=10),
                ],  # Metric 2
                "BD": [
                    datetime.time(0, 5, 0),
                    datetime.time(0, 2, 0),
                    datetime.time(0, 10, 0),
                    datetime.time(0, 3, 0),
                ],  # Metric 3
                "AP": ["OEE1", "OEE2", "OEE3", "OEE4"],  # OEE column
                "OEE_Value": ["90%", "80%", "95%", "75%"],
            }
        )
        self.grouping_col_name = "A"
        self.grouped_col_name = "B"
        self.metric_cols = ["H", "I", "BD"]
        self.oee_col_name = "OEE_Value"

    @patch("GraficApplication.GraphWorker.progress")
    @patch("GraficApplication.GraphWorker.finished")
    @patch("GraficApplication.GraphWorker.error")
    def test_run_success(self, mock_error, mock_finished, mock_progress):
        grouped_values = ["ProductX", "ProductY"]
        selected_grouping_val = "2023-01-01"

        worker = GraphWorker(
            self.df,
            self.grouping_col_name,
            self.grouped_col_name,
            grouped_values,
            self.metric_cols,
            self.oee_col_name,
            selected_grouping_val,
        )
        worker.run()

        mock_finished.emit.assert_called_once()
        results = mock_finished.emit.call_args[0][0]

        self.assertEqual(len(results), 2)

        # Check results for ProductX on 2023-01-01
        product_x_result = next(r for r in results if r[0] == "ProductX")
        self.assertIsNotNone(product_x_result)
        self.assertEqual(product_x_result[2], "90%")  # OEE value
        expected_sums_x = pd.Series(
            {"H": 1800.0, "I": 600.0, "BD": 300.0}
        )  # 30min, 10min, 5min in seconds
        pd.testing.assert_series_equal(product_x_result[1], expected_sums_x)

        # Check results for ProductY on 2023-01-01
        product_y_result = next(r for r in results if r[0] == "ProductY")
        self.assertIsNotNone(product_y_result)
        self.assertEqual(product_y_result[2], "80%")  # OEE value
        expected_sums_y = pd.Series(
            {"H": 900.0, "I": 300.0, "BD": 120.0}
        )  # 15min, 5min, 2min in seconds
        pd.testing.assert_series_equal(product_y_result[1], expected_sums_y)

        mock_error.emit.assert_not_called()

    @patch("GraficApplication.GraphWorker.progress")
    @patch("GraficApplication.GraphWorker.finished")
    @patch("GraficApplication.GraphWorker.error")
    def test_run_no_data_for_grouped_value(self, mock_error, mock_finished, mock_progress):
        grouped_values = ["ProductZ"]  # ProductZ does not exist in data
        selected_grouping_val = "2023-01-01"

        worker = GraphWorker(
            self.df,
            self.grouping_col_name,
            self.grouped_col_name,
            grouped_values,
            self.metric_cols,
            self.oee_col_name,
            selected_grouping_val,
        )
        worker.run()

        mock_finished.emit.assert_called_once_with([])  # No results should be emitted
        mock_error.emit.assert_not_called()

    @patch("GraficApplication.GraphWorker.progress")
    @patch("GraficApplication.GraphWorker.finished")
    @patch("GraficApplication.GraphWorker.error")
    def test_run_empty_metrics(self, mock_error, mock_finished, mock_progress):
        df_empty_metrics = self.df.copy()
        df_empty_metrics["H"] = np.nan
        df_empty_metrics["I"] = np.nan
        df_empty_metrics["BD"] = np.nan

        grouped_values = ["ProductX"]
        selected_grouping_val = "2023-01-01"

        worker = GraphWorker(
            df_empty_metrics,
            self.grouping_col_name,
            self.grouped_col_name,
            grouped_values,
            self.metric_cols,
            self.oee_col_name,
            selected_grouping_val,
        )
        worker.run()

        mock_finished.emit.assert_called_once_with([])  # No results because sums are empty
        mock_error.emit.assert_not_called()

    @patch("GraficApplication.GraphWorker.progress")
    @patch("GraficApplication.GraphWorker.finished")
    @patch("GraficApplication.GraphWorker.error")
    def test_run_oee_value_handling(self, mock_error, mock_finished, mock_progress):
        df_oee_test = self.df.copy()
        df_oee_test["OEE_Value"] = ["0.9", "0.85", "105%", "invalid"]

        grouped_values = ["ProductX", "ProductY"]
        selected_grouping_val = "2023-01-01"

        worker = GraphWorker(
            df_oee_test,
            self.grouping_col_name,
            self.grouped_col_name,
            grouped_values,
            self.metric_cols,
            self.oee_col_name,
            selected_grouping_val,
        )
        worker.run()

        results = mock_finished.emit.call_args[0][0]
        product_x_result = next(r for r in results if r[0] == "ProductX")
        self.assertEqual(product_x_result[2], "90%")

        product_y_result = next(r for r in results if r[0] == "ProductY")
        self.assertEqual(product_y_result[2], "85%")

        # Test for "invalid" OEE value for 2023-01-02 ProductY
        df_oee_test_invalid = self.df.copy()
        df_oee_test_invalid.loc[
            (df_oee_test_invalid["A"] == "2023-01-02")
            & (df_oee_test_invalid["B"] == "ProductY"),
            "OEE_Value",
        ] = "invalid"

        worker_invalid_oee = GraphWorker(
            df_oee_test_invalid,
            self.grouping_col_name,
            self.grouped_col_name,
            ["ProductY"],
            self.metric_cols,
            self.oee_col_name,
            "2023-01-02",
        )
        worker_invalid_oee.run()
        results_invalid_oee = mock_finished.emit.call_args[0][0]
        product_y_invalid_oee = next(
            r for r in results_invalid_oee if r[0] == "ProductY"
        )
        self.assertEqual(product_y_invalid_oee[2], "0%")


class TestGraphPlotter(unittest.TestCase):
    def setUp(self):
        self.sample_metrics = pd.Series(
            {"Metric1": 3600, "Metric2": 1800, "Metric3": 900, "Metric4": 300}
        )  # in seconds
        self.oee_value = "90%"
        self.chart_colors = ["red", "blue", "green", "purple"]

    @patch("matplotlib.pyplot.figure")
    @patch("matplotlib.pyplot.Axes.pie")
    @patch("matplotlib.pyplot.Axes.text")
    @patch("matplotlib.pyplot.Figure.text")
    @patch("matplotlib.pyplot.Axes.set_title")
    @patch("matplotlib.pyplot.Axes.axis")
    @patch("matplotlib.pyplot.Figure.tight_layout")
    def test_create_donut_chart(
        self,
        mock_tight_layout,
        mock_axis,
        mock_set_title,
        mock_fig_text,
        mock_ax_text,
        mock_pie,
        mock_figure,
    ):
        fig = MagicMock()
        ax = MagicMock()
        mock_figure.return_value = fig  # Mock the figure creation

        mock_pie.return_value = (
            [MagicMock(), MagicMock()],
            [MagicMock(), MagicMock()],
        )
        ax.pie = mock_pie
        ax.text = mock_ax_text
        fig.text = mock_fig_text
        ax.set_title = mock_set_title
        ax.axis = mock_axis # Assign the mock received from the decorator to the instance's method


        GraphPlotter.create_donut_chart(
            ax, self.sample_metrics, self.oee_value, self.chart_colors, fig
        )

        mock_pie.assert_called_once()
        args, kwargs = mock_pie.call_args
        self.assertEqual(list(args[0]), list(self.sample_metrics.values))
        self.assertEqual(
            kwargs["colors"], self.chart_colors[: len(self.sample_metrics)]
        )

        mock_ax_text.assert_called_once_with(
            0,
            0,
            f"OEE\n{self.oee_value}",
            horizontalalignment="center",
            verticalalignment="center",
            fontsize=24,
            fontweight="bold",
            color="black",
        )

        self.assertEqual(mock_fig_text.call_count, 3)  # For top 3 metrics
        mock_set_title.assert_called_once_with("")
        mock_axis.assert_called_once_with("equal")
        mock_tight_layout.assert_called_once()

    @patch("matplotlib.pyplot.figure")
    @patch("matplotlib.pyplot.Axes.barh")
    @patch("matplotlib.pyplot.Axes.set_yticks")
    @patch("matplotlib.pyplot.Axes.set_yticklabels")
    @patch("matplotlib.pyplot.Axes.invert_yaxis")
    @patch("matplotlib.pyplot.Axes.set_xlabel")
    @patch("matplotlib.pyplot.Axes.set_xticks")
    @patch("matplotlib.pyplot.Axes.set_title")
    @patch("matplotlib.pyplot.Axes.grid")
    @patch("matplotlib.pyplot.Axes.set_xlim")
    @patch("matplotlib.pyplot.Figure.tight_layout")
    def test_create_bar_chart(
            self,
            mock_tight_layout,
            mock_set_xlim,
            mock_grid,
            mock_set_title,
            mock_set_xticks,
            mock_set_xlabel,
            mock_invert_yaxis,
            mock_set_yticklabels,
            mock_set_yticks,
            mock_barh,
            mock_figure,
    ):
        fig = MagicMock()
        ax = MagicMock()
        mock_figure.return_value = fig  # Mock the figure creation
        ax.spines = {
            "right": MagicMock(),
            "top": MagicMock(),
            "left": MagicMock(),
            "bottom": MagicMock(),
        }

        ax.barh = mock_barh
        ax.set_yticks = mock_set_yticks
        ax.set_yticklabels = mock_set_yticklabels
        ax.invert_yaxis = mock_invert_yaxis
        ax.set_xlabel = mock_set_xlabel
        ax.set_xticks = mock_set_xticks
        ax.set_title = mock_set_title
        ax.grid = mock_grid  # Assign the mock received from the decorator to the instance's method
        ax.set_xlim = mock_set_xlim # Assign the mock received from the decorator to the instance's method
        fig.tight_layout = mock_tight_layout # Assign the mock received from the decorator to the instance's method


        GraphPlotter.create_bar_chart(
            ax, self.sample_metrics, self.oee_value, self.chart_colors
        )

        mock_barh.assert_called_once()
        args, kwargs = mock_barh.call_args
        self.assertEqual(len(args[0]), len(self.sample_metrics))  # y_pos
        self.assertEqual(
            [v / 60 for v in self.sample_metrics.values], list(args[1])
        )  # values_minutes
        self.assertEqual(kwargs["color"], self.chart_colors)

        mock_set_yticks.assert_called_once()
        mock_set_yticklabels.assert_called_once_with(
            self.sample_metrics.index.tolist(), fontsize=10
        )
        mock_invert_yaxis.assert_called_once()
        mock_set_xlabel.assert_called_once_with("")
        mock_set_xticks.assert_called_once_with([])
        mock_set_title.assert_called_once_with(
            f"OEE: {self.oee_value}", fontsize=16, fontweight="bold"
        )
        mock_grid.assert_called_once_with(False)

        ax.spines["right"].set_visible.assert_called_once_with(False)
        ax.spines["top"].set_visible.assert_called_once_with(False)
        ax.spines["left"].set_visible.assert_called_once_with(True)
        ax.spines["bottom"].set_visible.assert_called_once_with(True)

        self.assertEqual(ax.text.call_count, len(self.sample_metrics))  # ax.text'in çağrı sayısı kontrol edildi
        mock_set_xlim.assert_called_once_with(left=0)
        mock_tight_layout.assert_called_once()


if __name__ == "__main__":
    unittest.main(argv=["first-arg-is-ignored"], exit=False)