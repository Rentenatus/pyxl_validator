"""
<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Eclipse Public License v2.0 which accompanies this distribution, and is available at
http://www.eclipse.org/legal/epl-v20.html
</copyright>

Unit tests for the Excel differator functionality.

This module tests the row-by-row comparison of Excel tables using
different file formats (.xlsx, .xls, .ods). It verifies that differences
are detected, highlighted, and documented correctly.
"""
import os
import unittest
import pandas as pd
from pyxl_validator.excel_table_engine import load_engine, copy_to_pandas, get_pandas_engine
from pyxl_validator.table_validator import (ExcelValueValidator, TolerantFloatValidator,
                                            OmittedValidator, IntValidator, BoolValidator)
from pyxl_validator.table_validator_registry import ValidatorRegistry
from pyxl_validator.table_comparison_summary import ComparisonSummary
from pyxl_validator.excel_differator import differentiate_sheets_by_ws

class TestExcelPandasDifferator(unittest.TestCase):
    """
    Unit test class for the Excel differator.

    Tests the comparison and update of expected Excel tables
    against input tables in various formats.
    """

    def setUp(self):
        """
        Prepares the test environment.

        Removes temporary output files if they exist.
        Loads input and expected tables using the engine factory.
        Initializes the ValidatorRegistry and ComparisonSummary.
        """
        try:
            os.remove("test/tmp/v-daten2_pd.xlsx")
        except FileNotFoundError:
            pass

        self.df = pd.DataFrame(
            data=[
                [0, "Alice", 30, True],
                [1, "Bob", 25, False],
                [2, "Janusch", 52, False]

            ],
            columns=["id", "Name", "Age", "Active"]
        )

        self.fmt = pd.DataFrame([
            [{}, {"bold": False}, {"italic": False}, {"underline": False}],
            [{}, {"bold": False}, {"italic": False}, {"underline": False}]
        ])
        self.pandas_engine1 = get_pandas_engine(self.df, self.fmt)

        # Load engines via factory
        self.wb1, self.eng_expected1 = load_engine("test/assets/expected/e-daten2_pd.xlsx", sheet_name="Tabelle1")

        # Initialize ValidatorRegistry with ExcelValueValidator as default
        self.registry = ValidatorRegistry()
        self.registry.set_default(ExcelValueValidator())
        self.registry.register_by_index(0, OmittedValidator())
        self.registry.register_by_index(3, BoolValidator())

        # Initialize summary with header row from reference
        self.summary = ComparisonSummary()
        self.summary.set_header_values(self.eng_expected1.get_row_values(1))

    EXPECTED = [
        { },
        { },
        { }
    ]

    def test_compare_and_update_expected(self):
        """
        Compares input tables with the expected reference table.

        Differences are highlighted and documented. The updated reference
        tables are saved to temporary files for further inspection.
        """
        differentiate_sheets_by_ws(self.pandas_engine1, self.eng_expected1, False, self.registry, self.summary)


        # Save the updated reference tables
        self.wb1.save("test/tmp/v-daten2_pd.xlsx")

        result = self.summary.summary_by_header_array()
        for val, expected in zip(result, self.EXPECTED):
            self.assertEqual(val, expected, f"Failed for {val}")

