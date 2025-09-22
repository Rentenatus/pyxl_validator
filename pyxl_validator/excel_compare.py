"""
excel_compare.py

Vergleicht zwei Excel-Worksheets (.xlsx via openpyxl, .xls via xlrd)
unter Verwendung eines Engine-Interfaces, um if/elif-Kaskaden zu vermeiden.

Struktur:
    - TableEngine (Interface)
    - TableEnginePyxl (Implementierung für openpyxl)
    - TableEngineXlrd (Implementierung für xlrd)
    - Factory-Funktion: load_engine(file_path, sheet_name)
    - Vergleichsfunktionen: compare_sheets_by_file, compare_sheets_by_ws
"""


from pyxl_validator.excel_table_engine import TableEngine, TableRowEnumerator, load_engine
from pyxl_validator.table_validator import ComparisonResult, EqualValidator
from typing import Any

# ============================================================
# Vergleichslogik
# ============================================================

def compare_sheets_by_file(file1, sheet1, file2, sheet2,
                           validator_arr=None, validator_dict=None, default_validator=None):
    """
    Vergleicht zwei Excel-Worksheets anhand ihrer Dateipfade und Sheetnamen.

    Lädt die Tabellen über die Engine-Fabrik und delegiert den Vergleich
    an compare_sheets_by_ws. Optional können Validatoren spaltenweise
    übergeben werden – entweder als Liste (validator_arr), als Dictionary
    (validator_dict), oder als Fallback (default_validator).

    :param file1: Pfad zur Datei mit den gemessenen Werten.
    :param sheet1: Name des zu vergleichenden Sheets in Datei 1.
    :param file2: Pfad zur Datei mit den erwarteten Werten.
    :param sheet2: Name des zu vergleichenden Sheets in Datei 2.
    :param validator_arr: Liste von Validatoren pro Spalte (Index-basiert).
    :param validator_dict: Dictionary mit Validatoren pro Spalte (Index oder Spaltenname).
    :param default_validator: Fallback-Validator für nicht spezifizierte Spalten.
    :return: Liste von Vergleichsergebnissen (ComparisonResult).
    """
    _, eng1 = load_engine(file1, sheet1)
    _, eng2 = load_engine(file2, sheet2)
    return compare_sheets_by_ws(eng1, eng2, validator_arr, validator_dict, default_validator)


def compare_sheets_by_ws(eng1: TableEngine, eng2: TableEngine,
                         validator_arr=None, validator_dict=None, default_validator=None,
                         consumer=None):
    """
    Vergleicht zwei Excel-Worksheets zeilenweise und spaltenweise.

    Nutzt TableRowEnumerator für zeilenweises Durchlaufen.
    Validatoren werden pro Spalte angewendet. Unterschiede werden als
    ComparisonResult gesammelt. Strukturelle Differenzen wie längere oder
    kürzere Zeilen werden ebenfalls dokumentiert.

    Falls ein `consumer`-Objekt übergeben wird, wird jede Zeile direkt nach
    dem Vergleich an `consumer.diff(...)` übergeben. Dies ermöglicht
    speicherschonende Verarbeitung großer Tabellen oder Live-Streaming
    von Vergleichsergebnissen.

    :param eng1: TableEngine mit den gemessenen Werten.
    :param eng2: TableEngine mit den erwarteten Werten.
    :param validator_arr: Liste von Validatoren pro Spalte (Index-basiert).
    :param validator_dict: Dictionary mit Validatoren pro Spalte (Index oder Spaltenname).
    :param default_validator: Fallback-Validator für nicht spezifizierte Spalten.
    :param consumer: Optionales Objekt mit diff()-Methode zur zeilenweisen Verarbeitung.
    :return: Falls kein consumer gesetzt ist, Liste aller Vergleichsergebnisse.
    """
    validator_arr = calculate_validator_array(eng2, validator_arr, validator_dict, default_validator)
    enum1 = TableRowEnumerator(eng1)
    enum2 = TableRowEnumerator(eng2)
    return compare_sheets_by_enum(enum1, enum2, validator_arr=validator_arr, consumer=consumer)


def compare_sheets_by_enum(enum1: TableRowEnumerator, enum2: TableRowEnumerator,
                           validator_arr=None, consumer=None):
    """
    Vergleicht zwei TableRowEnumerator zeilenweise und spaltenweise.

    Jede Zeile wird mit compare_a_row verglichen. Falls ein consumer gesetzt ist,
    wird jede Vergleichszeile direkt übergeben. Andernfalls wird eine Gesamtliste zurückgegeben.

    :param enum1: Enumerator über die gemessenen Werte.
    :param enum2: Enumerator über die erwarteten Werte.
    :param validator_arr: Liste von Validatoren pro Spalte.
    :param consumer: Optionales Objekt mit diff()-Methode zur zeilenweisen Verarbeitung.
    :return: Liste aller Vergleichsergebnisse oder None bei consumer-Modus.
    """
    max_rows = max(enum1.get_max_row(), enum2.get_max_row())
    all_differences = [] if consumer is None else None

    validator_arr_nur_str = calculate_validator_array(enum2.engine, None, None, EqualValidator())
    compare_next(1, enum1, enum2, validator_arr_nur_str, consumer, all_differences)

    for r in range(2, max_rows + 1):
        compare_next(r, enum1, enum2, validator_arr, consumer, all_differences)

    return all_differences


def compare_next(r: int, enum1: TableRowEnumerator, enum2: TableRowEnumerator, validator_arr, consumer: Any | None,
                 all_differences: list[Any] | None):
    try:
        index1, row1 = next(enum1)
    except StopIteration:
        index1, row1 = -1, []

    try:
        index2, row2 = next(enum2)
    except StopIteration:
        index2, row2 = -1, []

    differences = compare_a_row(row1, row2, validator_arr)

    if consumer:
        consumer.diff(r, index1, row1, index2, row2, differences)
    else:
        all_differences.append((index1, row1, index2, row2, differences))


def compare_a_row(row1: list[Any], row2: list[Any], validator_arr: list) -> list[ComparisonResult]:
    """
    Vergleicht zwei Zeilen zellenweise mit Hilfe der Validatoren.

    :param row1: Wertezeile aus Messdaten.
    :param row2: Wertezeile aus Referenzdaten.
    :param validator_arr: Liste von Validatoren pro Spalte.
    :return: Liste von ComparisonResult-Werten pro Zelle.
    """
    differences = []
    for c, (val1, val2) in enumerate(zip(row1, row2)):
        v = validator_arr[c]
        differences.append(v.compare(val1, val2) if v else ComparisonResult.OMITTED)

    if len(row1) > len(row2):
        differences.extend([ComparisonResult.LONGER] * (len(row1) - len(row2)))
    elif len(row2) > len(row1):
        differences.extend([ComparisonResult.SHORTER] * (len(row2) - len(row1)))

    return differences


# ============================================================
# Utils
# ============================================================

def calculate_validator_array(eng2: TableEngine, validator_arr,
                              validator_dict, default_validator) -> list:
    """
    Erzeugt ein vollständiges Validator-Array für den Spaltenvergleich.

    Kombiniert die übergebene Validator-Liste mit einem Dictionary,
    das entweder Spaltenindizes oder Spaltennamen referenziert.
    Fehlende Einträge werden mit dem Default-Validator aufgefüllt.

    :param eng2: Engine mit den erwarteten Werten (für Spaltennamenzugriff).
    :param validator_arr: Liste von Validatoren pro Spalte (Index-basiert).
    :param validator_dict: Dictionary mit Validatoren pro Spalte (Index oder Spaltenname).
    :param default_validator: Fallback-Validator für nicht spezifizierte Spalten.
    :return: Vollständige Liste von Validatoren pro Spalte.
    """
    max_cols = eng2.get_max_col()

    if validator_arr is None:
        validator_arr = [default_validator] * max_cols
    elif len(validator_arr) < max_cols:
        validator_arr.extend([default_validator] * (max_cols - len(validator_arr)))

    if validator_dict:
        row2 = eng2.get_row_values(1)
        for key, value in validator_dict.items():
            if isinstance(key, int) and key < max_cols:
                validator_arr[key] = value
            elif isinstance(key, str):
                try:
                    index = row2.index(key)
                    validator_arr[index] = value
                except ValueError:
                    pass  # Spaltenname nicht gefunden

    return validator_arr
