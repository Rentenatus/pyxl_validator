"""
<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Eclipse Public License v2.0 which accompanies this distribution, and is available at
http://www.eclipse.org/legal/epl-v20.html
</copyright>

Unit tests for type recognition and normalization functions.

This module tests the internal type recognition helpers in
`pyxl_validator.table_validator`, including boolean, date, integer,
float, and number detection and normalization.
"""

from datetime import datetime
import unittest
import pyxl_validator.table_validator

class TestTypeRecognition(unittest.TestCase):
    """
    Unit test class for type recognition and normalization functions.

    Verifies correct detection and normalization of booleans, dates,
    integers, floats, and numbers from various input formats.
    """

    def test_is_bool_like(self):
        """
        Tests the `_is_bool_like` function for various boolean representations.

        Checks if values like True, "Ja", "nein", "1", "yes", etc. are
        correctly recognized as boolean-like.
        """
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
        """
        Tests the `_is_date_then_normalize` function for date recognition.

        Checks if strings and datetime objects are correctly identified
        as dates and normalized.
        """
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
        """
        Tests the `_is_int_then_normalize` function for integer recognition.

        Checks if integers and integer-like strings are correctly detected.
        """
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
        """
        Tests the `_is_float_then_normalize` function for float recognition.

        Checks if floats, integers, and localized float strings are detected.
        """
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
        """
        Tests the `_is_number_then_normalize` function for general number recognition.

        Checks if various numeric strings and values are detected as numbers.
        """
        cases = [
            ("42", True),
            ("42.0", True),
            ("3,14", True),
            ("abc", False),
        ]
        for val, expected in cases:
            result, _ = pyxl_validator.table_validator._is_number_then_normalize(val)
            self.assertEqual(result, expected, f"Failed for {val}")