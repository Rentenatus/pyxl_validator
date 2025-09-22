import unittest
from datetime import datetime

import pyxl_validator.table_validator


class TestTypeRecognition(unittest.TestCase):
    def test_is_bool_like(self):
        cases = [
            (True, True),
            (False, True),
            ("Ja", True),
            ("nein", True),
            ("true", True),
            ("WAHR", True),
            ("1", True),
            ("yes", True),
            ("false", True),
            ("FALSCH", True),
            ("0", True),
            ("no", True),
            ("vielleicht", False),
            (123, False),
        ]
        for val, expected in cases:
            self.assertEqual(pyxl_validator.table_validator._is_bool_like(val), expected, f"Failed for {val}")

    def test_is_date_then_normalize(self):
        cases = [
            ("2023-09-15", True),
            ("15.09.2023", False),
            (datetime(2023, 9, 15), True),
            ("invalid", False),
        ]
        for val, expected in cases:
            result, _ = pyxl_validator.table_validator._is_date_then_normalize(val)
            self.assertEqual(result, expected, f"Failed for {val}")

    def test_is_int_then_normalize(self):
        cases = [
            (42, True),
            ("42", True),
            ("42.0", False),
            ("42.5", False),
            ("abc", False),
        ]
        for val, expected in cases:
            result, _ = pyxl_validator.table_validator._is_int_then_normalize(val)
            self.assertEqual(result, expected, f"Failed for {val}")

    def test_is_float_then_normalize(self):
        cases = [
            (3.14, True),
            (42, True),
            ("3,14", True),
            ("1.234,56", True),
            ("1.000,00 Euro", True),
            ("abc", False),
        ]
        for val, expected in cases:
            result, _ = pyxl_validator.table_validator._is_float_then_normalize(val)
            self.assertEqual(result, expected, f"Failed for {val}")

    def test_is_number_then_normalize(self):
        cases = [
            ("42", True),
            ("42.0", True),
            ("3,14", True),
            ("abc", False),
        ]
        for val, expected in cases:
            result, _ = pyxl_validator.table_validator._is_number_then_normalize(val)
            self.assertEqual(result, expected, f"Failed for {val}")


