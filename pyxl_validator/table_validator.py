"""
table_validator.py

<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Eclipse Public License v2.0 which accompanies this distribution, and is available at
http://www.eclipse.org/legal/epl-v20.html
</copyright>

Defines the TableValidator interface and concrete validators
for comparing Excel cells, based on ComparisonResult.

Enum:
    ComparisonResult(IntEnum)
"""

from abc import ABC, abstractmethod
from datetime import datetime, date
from enum import IntEnum
from typing import Any
import re

PYTHON_INT_REGEX = re.compile(r"^[+-]?\d+$")
PYTHON_FLOAT_REGEX = re.compile(r"^[+-]?(?:\d+\.\d*|\.\d+|\d+)(?:[eE][+-]?\d+)?$")

GERMAN_DEZIM = True

# ============================================================
# Comparison Results
# ============================================================

class ComparisonResult(IntEnum):
    """
    Enum for cell comparison results.
    """
    EQUALS = 0       # Values are exactly equal (==)
    MATCHING = 1     # Values match, possibly in different formats
    ALMOST = 2       # Values are nearly equal (e.g., rounding errors, date comparison by day)
    ACCEPTED = 3     # Values are acceptably different
    OMITTED = 4      # Comparison was intentionally omitted
    DIFFERENT = 8    # Values are different
    CORRUPTED = 9    # At least one value is invalid/corrupted

    SHORTER = 10     # Array or row in Excel is too short
    LONGER = 11      # Array or row in Excel is longer than expected

    def ok(self) -> bool:
        """
        Returns True if the result is considered acceptable.
        """
        return self in {ComparisonResult.EQUALS, ComparisonResult.MATCHING, ComparisonResult.ALMOST,
                        ComparisonResult.ACCEPTED, ComparisonResult.OMITTED}

    def foul(self) -> bool:
        """
        Returns True if the result is considered not acceptable.
        """
        return self in {ComparisonResult.DIFFERENT, ComparisonResult.CORRUPTED}

    def get_cell_colors(result) -> tuple[str, str]:
        """
        Returns the color pair (measured value, reference) for a comparison type.

        :param result: ComparisonResult enum value.
        :return: Tuple of two RGB color strings ("RRGGBB").
        """
        return COLOR_MAP.get(result, ("DDDDDD", "DDDDDD"))  # Fallback: almost white


# RGB colors for Excel cells (openpyxl compatible)
COLOR_MAP = {
    ComparisonResult.EQUALS:     ("FFFFFF", "FFFFFF"),  # white, white
    ComparisonResult.MATCHING:   ("FFFFFF", "CCFFCC"),  # white, light green
    ComparisonResult.ALMOST:     ("FFFFFF", "CCFFFF"),  # white, light turquoise
    ComparisonResult.ACCEPTED:   ("CCFF99", "FFFF99"),  # light yellow-green, light yellow
    ComparisonResult.OMITTED:    ("CCCCCC", "CCCCCC"),  # gray, gray
    ComparisonResult.DIFFERENT:  ("CCFFCC", "FF9999"),  # light green, light red
    ComparisonResult.CORRUPTED:  ("FF9999", "FF0000"),  # light red, red
    ComparisonResult.SHORTER:    ("E0CCFF", "990000"),  # light purple, dark red
    ComparisonResult.LONGER:     ("660066", "FFFF99"),  # dark purple, light yellow
}

# ============================================================
# Interface
# ============================================================

class TableValidator(ABC):
    """
    Interface for cell comparison validators.
    """

    @abstractmethod
    def compare(self, val1: Any, val2: Any) -> ComparisonResult:
        """
        Compares two cell values and returns a ComparisonResult.
        """
        pass


# ============================================================
# EqualValidator
# ============================================================

class EqualValidator(TableValidator):
    """
    Compares values using simple equality (==).
    """

    def compare(self, val1: Any, val2: Any) -> ComparisonResult:
        return ComparisonResult.EQUALS if val1 == val2 else ComparisonResult.DIFFERENT


# ============================================================
# BoolValidator
# ============================================================

class BoolValidator(TableValidator):
    """
    Compares boolean values including typical Excel strings:
    - "TRUE", "FALSE", "WAHR", "FALSCH", "YES", "NO", "JA", "NEIN"
    - 1, 0
    """

    TRUE_VALUES = {"true", "=true()", "wahr", "1", "yes", "ja"}
    FALSE_VALUES = {"false", "=false()", "falsch", "0", "no", "nein"}
    BOOL_VALUES = TRUE_VALUES.union(FALSE_VALUES)

    def _normalize(self, val: Any) -> bool:
        if isinstance(val, bool):
            return val
        if isinstance(val, (int, float)):
            return bool(val)
        if isinstance(val, str):
            val = val.strip().lower()
            if val in self.TRUE_VALUES:
                return True
            if val in self.FALSE_VALUES:
                return False
        raise ValueError(f"Unknown boolean value: {val}")

    def compare(self, val1: Any, val2: Any) -> ComparisonResult:
        try:
            b1 = self._normalize(val1)
            b2 = self._normalize(val2)
            if val1 == val2:
                return ComparisonResult.EQUALS
            return ComparisonResult.MATCHING if b1 == b2 else ComparisonResult.DIFFERENT
        except Exception:
            return ComparisonResult.CORRUPTED


# ============================================================
# DateValidator
# ============================================================

class DateValidator(TableValidator):
    """
    Compares date values with adjustable precision:
    - "day", "hour", "minute", "second"
    """

    def __init__(self, precision: str = "day"):
        self.precision = precision

    @staticmethod
    def _normalize(val: Any) -> datetime:
        if isinstance(val, datetime):
            return val
        if isinstance(val, str):
            return datetime.fromisoformat(val.strip())
        raise ValueError(f"Unknown date value: {val}")

    def compare(self, val1: Any, val2: Any) -> ComparisonResult:
        try:
            d1 = self._normalize(val1)
            d2 = self._normalize(val2)
            if d1 == d2:
                return ComparisonResult.EQUALS if val1 == val2 else ComparisonResult.MATCHING

            if self.precision == "day":
                match = d1.date() == d2.date()
            elif self.precision == "hour":
                match = d1.replace(minute=0, second=0, microsecond=0) == d2.replace(minute=0, second=0, microsecond=0)
            elif self.precision == "minute":
                match = d1.replace(second=0, microsecond=0) == d2.replace(second=0, microsecond=0)
            elif self.precision == "second":
                match = d1.replace(microsecond=0) == d2.replace(microsecond=0)
            else:
                raise ValueError(f"Unknown precision: {self.precision}")

            return ComparisonResult.ALMOST if match else ComparisonResult.DIFFERENT
        except Exception:
            return ComparisonResult.CORRUPTED


# ============================================================
# OmittedValidator
# ============================================================

class OmittedValidator(TableValidator):
    """
    Marks the cell as intentionally omitted.
    """

    def compare(self, val1: Any, val2: Any) -> ComparisonResult:
        return ComparisonResult.OMITTED


# ============================================================
# IgnoreValidator
# ============================================================

class IgnoreValidator(TableValidator):
    """
    Ignores the comparison and accepts all values as matching.
    """

    def compare(self, val1: Any, val2: Any) -> ComparisonResult:
        return ComparisonResult.MATCHING


# ============================================================
# IntValidator
# ============================================================

class IntValidator(TableValidator):
    """
    Compares integer values:
    - int + int / int + str → integer comparison
    """

    def compare(self, val1: Any, val2: Any) -> ComparisonResult:
        if val1 == val2 and type(val1) == type(val2): return ComparisonResult.EQUALS

        # Integer comparison
        like1, normalized1 = _is_int_then_normalize(val1)
        like2, normalized2 = _is_int_then_normalize(val2)
        if like1 and like2:
            return ComparisonResult.MATCHING if normalized1 == normalized2 else ComparisonResult.DIFFERENT

        # Not a number
        return ComparisonResult.CORRUPTED


# ============================================================
# NumberValidator
# ============================================================

class NumberValidator(TableValidator):
    """
    Compares numeric values:
    - int + int / int + str → integer comparison
    - float + float / float + str / float + int → comparison with rounding
    """

    def __init__(self, float_precision: int = 10):
        self.float_precision = float_precision

    def __str__(self):
        return f"{self.__class__.__name__}(float_precision={self.float_precision})"

    def __repr__(self):
        return f"<{self.__class__.__name__} float_precision={self.float_precision}>"

    def compare(self, val1: Any, val2: Any) -> ComparisonResult:
        if val1 == val2 and type(val1) == type(val2): return ComparisonResult.EQUALS

        # Integer comparison
        like1, normalized1 = _is_int_then_normalize(val1)
        like2, normalized2 = _is_int_then_normalize(val2)
        if like1 and like2:
            return ComparisonResult.MATCHING if normalized1 == normalized2 else ComparisonResult.DIFFERENT

        # Floating point numbers with rounding
        like1, normalized1 = _is_float_then_normalize(val1)
        like2, normalized2 = _is_float_then_normalize(val2)
        if like1 and like2:
            if normalized1 == normalized2: return ComparisonResult.MATCHING
            rounded1 = round(normalized1, self.float_precision)
            rounded2 = round(normalized2, self.float_precision)
            return ComparisonResult.ALMOST if rounded1 == rounded2 else ComparisonResult.DIFFERENT
        # Not a number
        return ComparisonResult.CORRUPTED


# ============================================================
# TolerantFloatValidator
# ============================================================

class TolerantFloatValidator(TableValidator):
    """
    Compares floating point numbers with a tolerance range:
    - accepts deviation within [val2 - delta_down, val2 + delta_up]
    """

    def __init__(self, delta_up: float, delta_down: float, float_precision: int = 10):
        self.delta_up = delta_up
        self.delta_down = delta_down
        self.float_precision = float_precision

    def compare(self, val1: Any, val2: Any) -> ComparisonResult:
        if val1 == val2: return ComparisonResult.EQUALS

        like1, normalized1 = _is_float_then_normalize(val1)
        like2, normalized2 = _is_float_then_normalize(val2)
        if like1 and like2:
            if normalized1 == normalized2: return ComparisonResult.MATCHING
            rounded1 = round(normalized1, self.float_precision)
            rounded2 = round(normalized2, self.float_precision)
            if rounded1 == rounded2: return ComparisonResult.ALMOST
            lower = normalized2 - self.delta_down
            upper = normalized2 + self.delta_up
            if lower <= normalized1 <= upper:
                return ComparisonResult.ACCEPTED
            else:
                return ComparisonResult.DIFFERENT
        # Not a number
        return ComparisonResult.CORRUPTED

# ============================================================
# ExcelValueValidator
# ============================================================

class ExcelValueValidator(TableValidator):
    """
    General validator for Excel cell values.
    Automatically detects the appropriate comparison type.
    """

    def __init__(self):
        self.bool_validator = BoolValidator()
        self.date_validator = DateValidator(precision="day")
        self.number_validator = NumberValidator(float_precision=10)

    def compare(self, val1: Any, val2: Any) -> ComparisonResult:
        # 1. Direct comparison
        if val1 == val2:
            return ComparisonResult.EQUALS

        if val1 is None or val2 is None:
            return ComparisonResult.DIFFERENT

        # 2. Date logic
        like1, normalized1 = _is_date_then_normalize(val1)
        like2, normalized2 = _is_date_then_normalize(val2)
        if like1 and like2:
            return self.date_validator.compare(normalized1, normalized2)

        # 3. Numbers
        like1, normalized1 = _is_number_then_normalize(val1)
        like2, normalized2 = _is_number_then_normalize(val2)
        if like1 and like2:
            return self.number_validator.compare(normalized1, normalized2)

        # 4. Boolean logic
        if _is_bool_like(val1) and _is_bool_like(val2):
            return self.bool_validator.compare(val1, val2)

        # 6. Fallback
        return ComparisonResult.DIFFERENT

# ------------------------------------------------------------
# Type detection functions
# ------------------------------------------------------------

def _is_bool_like(val: Any) -> bool:
    """
    Checks if a value is boolean-like.
    """
    if isinstance(val, bool):
        return True
    if isinstance(val, str):
        return val.strip().lower() in BoolValidator.BOOL_VALUES
    if isinstance(val, int):
        return val == 0 or val == 1
    if isinstance(val, float):
        return val == 0.0 or val == 1.0
    return False

def _is_date_then_normalize(val: Any) -> tuple[bool, datetime] | tuple[bool, None]:
    """
    Checks if a value is date-like and normalizes it to datetime.
    """
    if isinstance(val, datetime):
        return True, val
    if isinstance(val, date):
        return True, datetime.combine(val, datetime.min.time())
    if isinstance(val, str):
        try:
            normalized = datetime.fromisoformat(val.strip())
            return True, normalized
        except Exception:
            return False, None
    return False, None

def _is_int_then_normalize(val: Any) -> tuple[bool, int]:
    """
    Checks if a value is integer-like and normalizes it to int.
    """
    if isinstance(val, int):
        return True, val
    if isinstance(val, float):
        return val.is_integer(), int(val)
    if isinstance(val, str):
        val=val.strip()
        if PYTHON_INT_REGEX.match(val):
            return True,int(val)
    return False, 0

def _is_float_then_normalize(val: Any) -> tuple[bool, float]:
    """
    Checks if a value is float-like and normalizes it to float.
    Handles German decimal and currency formats.
    """
    if isinstance(val, float):
        return True, val
    if isinstance(val, int):
        return True, float(val)
    if isinstance(val, str):
        val = val.strip().lower()
        waehrung = GERMAN_DEZIM and (("€" in val) or ("euro" in val) or ("," in val))

        # Step 1: Remove currency symbols and spaces
        cleaned = (val.replace("€", "").replace("euro", "").
                   replace(" ", ""))

        # Step 2: Replace thousand separator and decimal comma
        # e.g. "1.234,56" → "1234.56"
        if waehrung:
            cleaned = cleaned.replace(".", "").replace(",", ".")

        if PYTHON_FLOAT_REGEX.match(cleaned):
            return True, float(cleaned)
    return False, 0.0


def _is_number_then_normalize(val: Any) -> tuple[bool, int] | tuple[bool, float]:
    """
    Checks if a value is number-like and normalizes it.
    """
    like, normalized  = _is_int_then_normalize(val)
    if like:
        return True, normalized
    return _is_float_then_normalize(val)


