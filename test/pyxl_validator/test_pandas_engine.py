from unittest import TestCase
from pyxl_validator.excel_table_engine import get_pandas_engine, load_pandas_engine
import pandas as pd

class TestTableEnginePandas(TestCase):

    def setUp(self):
        self.df = pd.DataFrame(
            data=[
                ["Alice", 30, True],
                ["Bob", 25, False],
                ["Janusch", 52, False]

            ],
            columns=["Name", "Age", "Active"]
        )

        self.fmt = pd.DataFrame([
            [{"bold": False}, {"italic": False}, {"underline": False}],
            [{"bold": False}, {"italic": False}, {"underline": False}]
        ])
        self.pandas_engine = get_pandas_engine(self.df, self.fmt)

    def test_get_max_row(self):
        self.assertEqual(self.pandas_engine.get_max_row(), 3)

    def test_get_max_col(self):
        self.assertEqual(self.pandas_engine.get_max_col(), 3)

    def test_get_cell_value(self):
        self.assertEqual(self.pandas_engine.get_cell_value(1, 1), "Alice")

    def test_get_row_values(self):
        self.assertEqual(self.pandas_engine.get_row_values(2), ["Bob", 25, False])

    def test_get_cell_format(self):
        fmt = self.pandas_engine.get_cell_format(1, 1)
        self.assertIsInstance(fmt, dict)
        self.assertIn("bold", fmt)
        self.assertFalse(fmt["bold"])

    def test_get_row_formats(self):
        formats = self.pandas_engine.get_row_formats(1)
        self.assertEqual(len(formats), 3)
        self.assertEqual(formats[0]["bold"], False)

    def test_is_readonly(self):
        self.assertFalse(self.pandas_engine.is_readonly())

    def test_is_engine_readonly(self):
        self.assertFalse(self.pandas_engine.is_engine_readonly())

    def test_set_cell_value(self):
        self.pandas_engine.set_cell_value(2, 2, 35)
        self.assertEqual(self.pandas_engine.get_cell_value(2, 2), 35)

    def test_add_row(self):
        initial_rows = self.pandas_engine.get_max_row()
        self.pandas_engine.add_row(initial_rows + 1)
        self.assertEqual(self.pandas_engine.get_max_row(), initial_rows + 1)

    def test_set_row_values(self):
        self.pandas_engine.set_row_values(2, ["Alice", 31, True])
        self.assertEqual(self.pandas_engine.get_row_values(2), ["Alice", 31, True])

    def test_set_cell_format(self):
        self.pandas_engine.set_cell_format(1, 1, {"bold": True})
        fmt = self.pandas_engine.get_cell_format(1, 1)
        self.assertTrue(fmt["bold"])

    def test_set_row_formats(self):
        new_formats = [{"bold": True}, {"italic": True}, {"underline": True}]
        self.pandas_engine.set_row_formats(1, new_formats)
        fmt_row = self.pandas_engine.get_row_formats(1)
        self.assertEqual(fmt_row, new_formats)

    def test_save(self):
        try:
            self.pandas_engine.save_as("test/tmp/test_output_engine_pandas")
        except Exception as e:
            self.fail(f"Saving raised an exception: {e}")
        pandas_engine2=load_pandas_engine("test/tmp/test_output_engine_pandas")
        pandas_engine2.save_as("test/tmp/test_output_engine_pandas2")
        self.assertEqual(pandas_engine2.get_row_values(1), self.pandas_engine.get_row_values(1))
        self.assertEqual(pandas_engine2.get_row_values(2), self.pandas_engine.get_row_values(2))
        self.assertEqual(pandas_engine2.get_cell_format(1,1), self.pandas_engine.get_cell_format(1,1))
        self.assertEqual(pandas_engine2.get_cell_format(2,3), self.pandas_engine.get_cell_format(2,3))