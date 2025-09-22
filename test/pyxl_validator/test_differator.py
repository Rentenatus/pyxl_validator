import os
import unittest
from pyxl_validator.excel_table_engine import TableEngine, load_engine
from pyxl_validator.table_validator import EqualValidator, ExcelValueValidator
from pyxl_validator.table_validator_registry import ValidatorRegistry
from pyxl_validator.table_comparison_summary import ComparisonSummary
from pyxl_validator.excel_differator import differentiate_sheets_by_ws





class TestExcelDifferator(unittest.TestCase):
    def setUp(self):
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

        # Lade Engines über Factory
        self.wb1, self.eng_expected1 = load_engine("test/assets/expected/e-daten1.xlsx",sheet_name="Tabelle1")
        self.wb2, self.eng_expected2 = load_engine("test/assets/expected/e-daten1.xlsx",sheet_name="Tabelle1")
        self.wb3, self.eng_expected3 = load_engine("test/assets/expected/e-daten1.xlsx",sheet_name="Tabelle1")
        _, self.eng_input1 = load_engine("test/assets/input/daten1.xlsx",sheet_name="Tabelle1")
        _, self.eng_input2 = load_engine("test/assets/input/daten1.xls",sheet_name="Tabelle1")
        _, self.eng_input3 = load_engine("test/assets/input/daten1.ods",sheet_name="Tabelle1")

        # Initialisiere ValidatorRegistry mit EqualValidator
        self.registry = ValidatorRegistry()
        self.registry.set_default(ExcelValueValidator())

        # Initialisiere Summary mit Headerzeile aus Referenz
        self.summary = ComparisonSummary()
        self.summary.set_header_values(self.eng_expected1.get_row_values(1))

    def test_compare_and_update_expected(self):
        # Vergleiche beide Eingaben mit der Referenz
        differentiate_sheets_by_ws(self.eng_input1, self.eng_expected1, self.registry, self.summary)
        differentiate_sheets_by_ws(self.eng_input2, self.eng_expected2, self.registry, self.summary)
        differentiate_sheets_by_ws(self.eng_input3, self.eng_expected3, self.registry, self.summary)

        # Speichere die aktualisierte Referenz
        self.wb1.save("test/tmp/v-daten1-xlsx.xlsx")
        self.wb2.save("test/tmp/v-daten1-xls.xlsx")
        self.wb2.save("test/tmp/v-daten1-ods.xlsx")