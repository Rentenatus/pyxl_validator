"""
table_validator.py

Definiert das TableValidator-Interface und konkrete Validatoren
für den Vergleich von Excel-Zellen, basierend auf ComparisonResult.

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
# Vergleichsresultate
# ============================================================

class ComparisonResult(IntEnum):
    EQUALS = 0       # Werte sind gleich mit ==
    MATCHING = 1     # Werte sind gleich, wenn auch in unterschiedlichen formaten
    ALMOST = 2       # Werte sind in vorgegebenen Scope gleich (z.B. Rundungsfehler bei Ganzzahlvergleich,
                     #                                          Datumsvergleich mit Tag-Genauigkeit)
    ACCEPTED = 3     # Werte sind akzeptabel unterschiedlich
    OMITTED = 4      # Vergleich wurde bewusst ausgelassen
    DIFFERENT = 8    # Werte sind unterschiedlich
    CORRUPTED = 9    # Mindestens einer der Werte ist ungültig/korrupt (Fehler bei der Verarbeitung)

    SHORTER = 10     # Array oder Zeile in Excel ist zu kurz.
    LONGER = 11      # Array oder Zeile in Excel ist zu länger als erwartet.

    def ok(self) -> bool:
        """Gibt zurück, ob das Ergebnis als akzeptabel gilt."""
        return self in {ComparisonResult.EQUALS, ComparisonResult.MATCHING, ComparisonResult.ALMOST,
                        ComparisonResult.ACCEPTED, ComparisonResult.OMITTED}

    def foul(self) -> bool:
        """Gibt zurück, ob das Ergebnis als nicht akzeptabel gilt."""
        return self in {ComparisonResult.DIFFERENT, ComparisonResult.CORRUPTED}

    def get_cell_colors(result) -> tuple[str, str]:
        """
        Gibt das Farbpaar (Messwert, Referenz) für einen Vergleichstyp zurück.

        :param result: Vergleichsergebnis als ComparisonResult-Enum.
        :return: Tupel mit zwei RGB-Farben im Format "RRGGBB".
        """
        return COLOR_MAP.get(result, ("DDDDDD", "DDDDDD"))  # Fallback: fast weiß


# RGB-Farben im Format "RRGGBB" für Excel-Zellen (openpyxl-kompatibel)
COLOR_MAP = {
    ComparisonResult.EQUALS:     ("FFFFFF", "FFFFFF"),  # weiß, weiß
    ComparisonResult.MATCHING:   ("FFFFFF", "CCFFCC"),  # weiß, hell grün
    ComparisonResult.ALMOST:     ("FFFFFF", "CCFFFF"),  # weiß, hell türkis
    ComparisonResult.ACCEPTED:   ("CCFF99", "FFFF99"),  # hell gelb-grün, hell gelb
    ComparisonResult.OMITTED:    ("CCCCCC", "CCCCCC"),  # grau, grau
    ComparisonResult.DIFFERENT:  ("CCFFCC", "FF9999"),  # hell grün, hell rot
    ComparisonResult.CORRUPTED:  ("FF9999", "FF0000"),  # hell rot, rot
    ComparisonResult.SHORTER:    ("E0CCFF", "990000"),  # hell lila, dunkel rot
    ComparisonResult.LONGER:     ("660066", "FFFF99"),  # dunkel lila, hell gelb
}

# ============================================================
# Interface
# ============================================================

class TableValidator(ABC):
    """Interface für Zellvergleichs-Validatoren."""

    @abstractmethod
    def compare(self, val1: Any, val2: Any) -> ComparisonResult:
        """
        Vergleicht zwei Zellwerte und gibt ein ComparisonResult zurück.
        """
        pass


# ============================================================
# EqualValidator
# ============================================================

class EqualValidator(TableValidator):
    """Vergleicht Werte mit einfachem ==."""

    def compare(self, val1: Any, val2: Any) -> ComparisonResult:
        return ComparisonResult.EQUALS if val1 == val2 else ComparisonResult.DIFFERENT


# ============================================================
# BoolValidator
# ============================================================

class BoolValidator(TableValidator):
    """
    Vergleicht boolesche Werte inkl. typischer Excel-Strings:
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
        raise ValueError(f"Unbekannter Bool-Wert: {val}")

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
    Vergleicht Datumswerte mit einstellbarer Genauigkeit:
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
        raise ValueError(f"Unbekanntes Datum: {val}")

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
                raise ValueError(f"Unbekannte Genauigkeit: {self.precision}")

            return ComparisonResult.ALMOST if match else ComparisonResult.DIFFERENT
        except Exception:
            return ComparisonResult.CORRUPTED


# ============================================================
# OmittedValidator
# ============================================================

class OmittedValidator(TableValidator):
    """Kennzeichnet die Zelle als absichtlich ausgelassen."""

    def compare(self, val1: Any, val2: Any) -> ComparisonResult:
        return ComparisonResult.OMITTED


# ============================================================
# IgnoreValidator
# ============================================================

class IgnoreValidator(TableValidator):
    """Ignoriert den Vergleich und akzeptiert alles als gleich."""

    def compare(self, val1: Any, val2: Any) -> ComparisonResult:
        return ComparisonResult.MATCHING




# ============================================================
# IntValidator
# ============================================================

class IntValidator(TableValidator):
    """
    Vergleicht numerische ganzzahlige Werte:
    - int + int / int + str → Ganzzahlvergleich
    """

    def compare(self, val1: Any, val2: Any) -> ComparisonResult:
        if val1 == val2 and type(val1) == type (val2): return ComparisonResult.EQUALS

        # Ganzzahlvergleich
        like1, normalized1 = _is_int_then_normalize(val1)
        like2, normalized2 = _is_int_then_normalize(val2)
        if like1 and like2:
            return ComparisonResult.MATCHING if normalized1 == normalized2 else ComparisonResult.DIFFERENT

        # Keine Zahl
        return ComparisonResult.CORRUPTED


# ============================================================
# NumberValidator
# ============================================================

class NumberValidator(TableValidator):
    """
    Vergleicht numerische Werte:
    - int + int / int + str → Ganzzahlvergleich
    - float + float / float + str / float + int → Vergleich mit Rundung
    """

    def __init__(self, float_precision: int = 10):
        self.float_precision = float_precision

    def __str__(self):
        return f"{self.__class__.__name__}(float_precision={self.float_precision})"

    def __repr__(self):
        # repr kann etwas technischer sein, z. B. für Debug-Ausgaben
        return f"<{self.__class__.__name__} float_precision={self.float_precision}>"

    def compare(self, val1: Any, val2: Any) -> ComparisonResult:
        if val1 == val2 and type(val1) == type (val2): return ComparisonResult.EQUALS

        # Ganzzahlvergleich
        like1, normalized1 = _is_int_then_normalize(val1)
        like2, normalized2 = _is_int_then_normalize(val2)
        if like1 and like2:
            return ComparisonResult.MATCHING if normalized1 == normalized2 else ComparisonResult.DIFFERENT

        # Fließkommazahlen mit Rundung
        like1, normalized1 = _is_float_then_normalize(val1)
        like2, normalized2 = _is_float_then_normalize(val2)
        if like1 and like2:
            if normalized1 == normalized2: return ComparisonResult.MATCHING
            rounded1 = round(normalized1, self.float_precision)
            rounded2 = round(normalized2, self.float_precision)
            return ComparisonResult.ALMOST if rounded1 == rounded2 else ComparisonResult.DIFFERENT
        # Keine Zahl
        return ComparisonResult.CORRUPTED



# ============================================================
# TolerantFloatValidator
# ============================================================

class TolerantFloatValidator(TableValidator):
    """
    Vergleicht Fließkommazahlen mit Toleranzbereich:
    - akzeptiert Abweichung innerhalb [val2 - delta_down, val2 + delta_up]
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
        # Keine Zahl
        return ComparisonResult.CORRUPTED

# ============================================================
# ExcelValueValidator
# ============================================================

class ExcelValueValidator(TableValidator):
    """
    Allgemeiner Validator für Excel-Zellwerte.
    Erkennt automatisch den passenden Vergleichstyp.
    """

    def __init__(self):
        self.bool_validator = BoolValidator()
        self.date_validator = DateValidator(precision="day")
        self.number_validator = NumberValidator(float_precision=10)

    def compare(self, val1: Any, val2: Any) -> ComparisonResult:
        # 1. Direkter Vergleich
        if val1 == val2:
            return ComparisonResult.EQUALS

        if val1 is None or val2 is None:
            return ComparisonResult.DIFFERENT

        # 2. Datumslogik
        like1, normalized1 = _is_date_then_normalize(val1)
        like2, normalized2 = _is_date_then_normalize(val2)
        if like1 and like2:
            return self.date_validator.compare(normalized1, normalized2)

        # 3. Zahlen
        like1, normalized1 = _is_number_then_normalize(val1)
        like2, normalized2 = _is_number_then_normalize(val2)
        if like1 and like2:
            return self.number_validator.compare(normalized1, normalized2)

        # 4. Boolesche Logik
        if _is_bool_like(val1) and _is_bool_like(val2):
            return self.bool_validator.compare(val1, val2)

        # 6. Fallback
        return ComparisonResult.DIFFERENT

# ------------------------------------------------------------
# Typ-Erkennungsfunktionen
# ------------------------------------------------------------

def _is_bool_like(val: Any) -> bool:
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
    if isinstance(val, float):
        return True, val
    if isinstance(val, int):
        return True, float(val)
    if isinstance(val, str):
        val = val.strip().lower()
        waehrung = GERMAN_DEZIM and (("€" in val) or ("euro" in val) or ("," in val))

        # Schritt 1: Entferne Währungszeichen und Leerzeichen
        cleaned = (val.replace("€", "").replace("euro", "").
                   replace(" ", ""))

        # Schritt 2: Ersetze Tausenderpunkt und Dezimalkomma
        # z.B. "1.234,56" → "1234.56"
        if waehrung:
            cleaned = cleaned.replace(".", "").replace(",", ".")

        if PYTHON_FLOAT_REGEX.match(cleaned):
            return True, float(cleaned)
    return False, 0.0


def _is_number_then_normalize(val: Any) -> tuple[bool, int] | tuple[bool, float]:
    like, normalized  = _is_int_then_normalize(val)
    if like:
        return True, normalized
    return _is_float_then_normalize(val)


# ------------------------------------------------------------
# Registry:
# ------------------------------------------------------------

VALIDATORS = {
    "equal": EqualValidator,
    "bool": BoolValidator,
    "date": DateValidator,
    "number": NumberValidator,
    "int": IntValidator,
    "tolerant_float": TolerantFloatValidator,
    "omit": OmittedValidator,
    "ignore": IgnoreValidator,
    "auto": ExcelValueValidator,
    "excel": ExcelValueValidator,
}
