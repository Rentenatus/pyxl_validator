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
from pyxl_validator.excel_table_engine import load_engine, copy_to_pandas
from pyxl_validator.table_validator import (ExcelValueValidator, TolerantFloatValidator,
                                            OmittedValidator, IntValidator, BoolValidator)
from pyxl_validator.table_validator_registry import ValidatorRegistry
from pyxl_validator.table_comparison_summary import ComparisonSummary
from pyxl_validator.excel_differator import differentiate_sheets_by_ws

class TestExcelDifferator(unittest.TestCase):
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
            os.remove("test/tmp/v-daten1-xlsx.xlsx")
        except FileNotFoundError:
            pass
        try:
            os.remove("test/tmp/v-daten1-xls.xlsx")
        except FileNotFoundError:
            pass
        try:
            os.remove("test/tmp/v-daten1-ods.xlsx")
        except FileNotFoundError:
            pass

        # Load engines via factory
        self.wb1, self.eng_expected1 = load_engine("test/assets/expected/e-daten1.xlsx", sheet_name="Tabelle1")
        self.wb2, self.eng_expected2 = load_engine("test/assets/expected/e-daten1.xlsx", sheet_name="Tabelle1")
        self.wb3, self.eng_expected3 = load_engine("test/assets/expected/e-daten1.xlsx", sheet_name="Tabelle1")
        _, self.eng_input1 = load_engine("test/assets/input/daten1.xlsx", sheet_name="Tabelle1")
        _, self.eng_input2 = load_engine("test/assets/input/daten1.xls", sheet_name="Tabelle1")
        _, self.eng_input3 = load_engine("test/assets/input/daten1.ods", sheet_name="Tabelle1")

        # Initialize ValidatorRegistry with ExcelValueValidator as default
        self.registry = ValidatorRegistry()
        self.registry.set_default(ExcelValueValidator())
        self.registry.register_by_name("id", OmittedValidator())
        self.registry.register_by_name("Bool 2", BoolValidator())
        self.registry.register_by_name("Bool-Text", BoolValidator())
        self.registry.register_by_name("W3", TolerantFloatValidator(0.001, 0.001, 2))
        self.registry.register_by_name("W4", TolerantFloatValidator(0.001, 0.001, 4))
        self.registry.register_by_name("W5", TolerantFloatValidator(0.1, 0.1, 2))
        self.registry.register_by_name("W7", TolerantFloatValidator(20, 10, 2))
        self.registry.register_by_name("W8", IntValidator())
        self.registry.register_by_name("W9", TolerantFloatValidator(30, 10, 2))

        # Initialize summary with header row from reference
        self.summary = ComparisonSummary()
        self.summary.set_header_values(self.eng_expected1.get_row_values(1))

    EXPECTED = [
        {},
        {},
        {},
        {'DIFFERENT': 9},
        {},
        {},
        {'CORRUPTED': 12},
        {'DIFFERENT': 6},
        {'DIFFERENT': 6},
        {'CORRUPTED': 6},
        {'CORRUPTED': 6, 'DIFFERENT': 12},
        {'CORRUPTED': 6},
        {'DIFFERENT': 6},
        {'DIFFERENT': 15},
        {},
        {'DIFFERENT': 4}
    ]

    def test_compare_and_update_expected(self):
        """
        Compares input tables with the expected reference table.

        Differences are highlighted and documented. The updated reference
        tables are saved to temporary files for further inspection.
        """
        differentiate_sheets_by_ws(self.eng_input1, self.eng_expected1, registry = self.registry, summary = self.summary)
        differentiate_sheets_by_ws(self.eng_input2, self.eng_expected2, registry = self.registry, summary = self.summary)
        differentiate_sheets_by_ws(self.eng_input3, self.eng_expected3, registry = self.registry, summary = self.summary)

        # Save the updated reference tables
        self.wb1.save("test/tmp/v-daten1-xlsx.xlsx")
        self.wb2.save("test/tmp/v-daten1-xls.xlsx")
        self.wb2.save("test/tmp/v-daten1-ods.xlsx")

        result = self.summary.summary_by_header_array()
        for val, expected in zip(result, self.EXPECTED):
            self.assertEqual(val, expected, f"Failed for {val}")

        pate = copy_to_pandas(self.eng_expected3)
        pate.save_as("test/tmp/v-daten1-pandas")