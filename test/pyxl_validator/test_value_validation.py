import unittest
from datetime import datetime
from pyxl_validator.table_validator import (
    BoolValidator,
    DateValidator,
    IntValidator,
    NumberValidator,
    TolerantFloatValidator,
    ComparisonResult
)


def compare_check(test, v, val1, val2, expected):
    result = v.compare(val1, val2)
    if result != expected:
        print(v)
        print("val1 =",val1,":",type(val1))
        print("val2 =", val2,":",type(val2))
        print("result =", result.name)
        print("expected =", expected.name)
        # Hier der Platz für einen Breackpoint:
        v.compare(val1, val2)
    test.assertEqual(result, expected)

# ------------------------------------------------------------
# BoolValidator Tests
# ------------------------------------------------------------
class TestBoolValidator(unittest.TestCase):

    def setUp(self):
        self.v = BoolValidator()

    def test_cases(self):
        cases = [
            # EQUALS
            (True, True, ComparisonResult.EQUALS),
            (False, False, ComparisonResult.EQUALS),

            # DIFFERENT
            (True, False, ComparisonResult.DIFFERENT),
            (False, True, ComparisonResult.DIFFERENT),

            # MATCHING (verschiedene Formate)
            ("TRUE", "true", ComparisonResult.MATCHING),
            (" Ja ", "yes", ComparisonResult.MATCHING),
            ("0", 0, ComparisonResult.MATCHING),
            ("1", 1, ComparisonResult.MATCHING),

            # CORRUPTED
            ("invalid", True, ComparisonResult.CORRUPTED),
            ("maybe", "nope", ComparisonResult.CORRUPTED),
            (self, False, ComparisonResult.CORRUPTED),
            (self.v, True, ComparisonResult.CORRUPTED),
        ]
        for val1, val2, expected in cases:
            with self.subTest(val1=val1, val2=val2):
                compare_check(self, self.v, val1, val2, expected)




# ------------------------------------------------------------
# DateValidator Tests
# ------------------------------------------------------------
class TestDateValidator(unittest.TestCase):

    def test_cases(self):
        cases = [
            # EQUALS
            ("2023-10-01", "2023-10-01", ComparisonResult.EQUALS, "day"),

            # MATCHING
            ("2023-10-01T00:00:00", datetime(2023, 10, 1), ComparisonResult.MATCHING, "day"),

            # ALMOST
            ("2023-10-01T00:00:00", "2023-10-01T23:59:59", ComparisonResult.ALMOST, "day"),
            ("2023-10-01T12:05:00", "2023-10-01T12:50:00", ComparisonResult.ALMOST, "hour"),
            ("2023-10-01T12:30:00", "2023-10-01T12:30:45", ComparisonResult.ALMOST, "minute"),
            ("2023-10-01T12:30:45.123", "2023-10-01T12:30:45.999", ComparisonResult.ALMOST, "second"),

            # DIFFERENT
            ("2023-10-01", "2023-09-30", ComparisonResult.DIFFERENT, "day"),

            # CORRUPTED
            ("invalid-date", "2023-10-01", ComparisonResult.CORRUPTED, "day"),
            (self, "2023-10-01", ComparisonResult.CORRUPTED, "day"),
        ]
        for val1, val2, expected, precision in cases:
            with self.subTest(val1=val1, val2=val2, precision=precision):
                v = DateValidator(precision=precision)
                compare_check(self, v, val1, val2, expected)


# ------------------------------------------------------------
# NumberValidator Tests
# ------------------------------------------------------------
class TestNumberValidator(unittest.TestCase):

    def test_cases(self):
        cases = [
            # EQUALS
            (5, 5, ComparisonResult.EQUALS, 10),
            (3.14, 3.14, ComparisonResult.EQUALS, 5),
            ("1000", "1000", ComparisonResult.EQUALS, 10),
            ("1.000,20", "1.000,20", ComparisonResult.EQUALS, 4),
            ("1000,20", "1000,20", ComparisonResult.EQUALS, 4),
            ("1000.20", "1000.20", ComparisonResult.EQUALS, 4),

            # MATCHING (verschiedene Formate, gleiche Zahl)
            (5, "5", ComparisonResult.MATCHING, 10),
            ("42", 42, ComparisonResult.MATCHING, 10),
            ("5.00", 5.0, ComparisonResult.MATCHING, 5),
            ("1,23", 1.23, ComparisonResult.MATCHING, 2),  # deutsches Komma
            ("1.234,56", 1234.56, ComparisonResult.MATCHING, 2),  # Tausenderpunkt + Komma
            ("€1.234,56", 1234.56, ComparisonResult.MATCHING, 2),  # Euro-Symbol
            ("1.234,56 Euro", 1234.56, ComparisonResult.MATCHING, 2),  # Euro-Wort
            ("  1.234,56  ", 1234.56, ComparisonResult.MATCHING, 2),  # mit Leerzeichen
            ("0,00", 0.0, ComparisonResult.MATCHING, 2),
            ("0001,50", 1.5, ComparisonResult.MATCHING, 2),
            ("€1.234", 1234, ComparisonResult.MATCHING, 2),  # Euro-Symbol
            (1234.0, 1234, ComparisonResult.MATCHING, 2),

            # ALMOST (Rundungsabweichung innerhalb float_precision)
            (5.123, 5.124, ComparisonResult.ALMOST, 2),
            ("1,234", 1.231, ComparisonResult.ALMOST, 2),  # deutsches Komma, passt mit 2 Kommastellen
            ("1,234", 1.228, ComparisonResult.ALMOST, 2),  # deutsches Komma, passt mit 2 Kommastellen
            ("1234,5678", 1234.5681, ComparisonResult.ALMOST, 3),

            # DIFFERENT (echte Abweichung)
            ("1,23", 1.25, ComparisonResult.DIFFERENT, 2),  # außerhalb Rundung
            ("1.234,56", 1235.56, ComparisonResult.DIFFERENT, 2),
            ("€1.234,56", 1234.00, ComparisonResult.DIFFERENT, 2),
            ("1000", 1001, ComparisonResult.DIFFERENT, 0),
            ("1000", "1001", ComparisonResult.DIFFERENT, 0),

            # CORRUPTED
            ("abc", 5, ComparisonResult.CORRUPTED, 10),
            ("invalid", 10.0, ComparisonResult.CORRUPTED, 8),
            ("1-2-3", "123", ComparisonResult.CORRUPTED, 8),
            (datetime(2023, 10, 1), "123", ComparisonResult.CORRUPTED, 8),
            ("1,2,3", 5, ComparisonResult.CORRUPTED, 2),  # mehrere Kommas
            ("1.2.3", 5, ComparisonResult.CORRUPTED, 2),  # mehrere Punkte
            ("1-2-3", 5, ComparisonResult.CORRUPTED, 2),  # Bindestriche im Wert
            ("€abc", 5, ComparisonResult.CORRUPTED, 2),  # Währung, aber keine Zahl
            (self, 5, ComparisonResult.CORRUPTED, 2),
        ]

        for val1, val2, expected, prec in cases:
            with self.subTest(val1=val1, val2=val2, prec=prec):
                v = NumberValidator(float_precision=prec)
                compare_check(self, v, val1, val2, expected)

# ------------------------------------------------------------
# IntValidator Tests
# ------------------------------------------------------------
class TestIntValidator(unittest.TestCase):

    def setUp(self):
        self.v = IntValidator()

    def test_cases(self):
        cases = [
            # EQUALS
            (5, 5, ComparisonResult.EQUALS),
            ("1000", "1000", ComparisonResult.EQUALS),

            # MATCHING (verschiedene Formate, gleiche Zahl)
            (5, "5", ComparisonResult.MATCHING),
            ("42", 42, ComparisonResult.MATCHING),
            ("   1000   ", 1000, ComparisonResult.MATCHING),
            (1000, "   1000   ", ComparisonResult.MATCHING),
            ("1000", "1000          ", ComparisonResult.MATCHING),

            # DIFFERENT (echte Abweichung)
            ("123", 125, ComparisonResult.DIFFERENT),
            ("1000", 1001, ComparisonResult.DIFFERENT),
            ("1000", "1001", ComparisonResult.DIFFERENT),

            # CORRUPTED
            ("abc", 5, ComparisonResult.CORRUPTED),
            ("invalid", 10, ComparisonResult.CORRUPTED),
            ("1-2-3", "123", ComparisonResult.CORRUPTED),
            (datetime(2023, 10, 1), "123", ComparisonResult.CORRUPTED),
            ("1,2,3", 5, ComparisonResult.CORRUPTED),
            ("1.2.3", 5, ComparisonResult.CORRUPTED),
            ("1-2-3", 5, ComparisonResult.CORRUPTED),
            ("€abc", 5, ComparisonResult.CORRUPTED),
            (self.v, 5, ComparisonResult.CORRUPTED),
        ]

        for val1, val2, expected in cases:
            with self.subTest(val1=val1, val2=val2):
                compare_check(self, self.v, val1, val2, expected)


# ------------------------------------------------------------
# TolerantFloatValidator Tests
# ------------------------------------------------------------
class TestTolerantFloatValidator(unittest.TestCase):

    def test_cases(self):
        cases = [
            # EQUALS
            (1.00, 1.00, ComparisonResult.EQUALS),

            # MATCHING
            ("2.0", 2.0, ComparisonResult.MATCHING),

            # ALMOST
            (3.141, 3.140, ComparisonResult.ALMOST),

            # ACCEPTED
            (4.99, 5.00, ComparisonResult.ACCEPTED),
            (5.02, 5.00, ComparisonResult.ACCEPTED),

            # DIFFERENT
            (4.98, 5.00, ComparisonResult.DIFFERENT),
            (5.03, 5.00, ComparisonResult.DIFFERENT),

            # CORRUPTED
            ("not a number", 1.0, ComparisonResult.CORRUPTED),
        ]
        v = TolerantFloatValidator(delta_up=0.02, delta_down=0.01, float_precision=2)
        for val1, val2, expected in cases:
            with self.subTest(val1=val1, val2=val2):
                compare_check(self, v, val1, val2, expected)


# ------------------------------------------------------------
# Test Runner
# ------------------------------------------------------------
if __name__ == "__main__":
    unittest.main()
