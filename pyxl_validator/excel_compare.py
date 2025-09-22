"""
excel_compare.py

<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Eclipse Public License v2.0 which accompanies this distribution, and is available at
http://www.eclipse.org/legal/epl-v20.html
</copyright>

Compares two Excel worksheets (.xlsx via openpyxl, .xls via xlrd)
using an engine interface to avoid if/elif cascades.

Structure:
    - TableEngine (interface)
    - TableEnginePyxl (implementation for openpyxl)
    - TableEngineXlrd (implementation for xlrd)
    - Factory function: load_engine(file_path, sheet_name)
    - Comparison functions: compare_sheets_by_file, compare_sheets_by_ws
"""

from pyxl_validator.excel_table_engine import TableEngine, TableRowEnumerator, load_engine
from pyxl_validator.table_validator import ComparisonResult, EqualValidator
from typing import Any

# ============================================================
# Comparison Logic
# ============================================================

def compare_sheets_by_file(file1, sheet1, file2, sheet2,
                           validator_arr=None, validator_dict=None, default_validator=None):
    """
    Compares two Excel worksheets by file path and sheet name.

    Loads the tables using the engine factory and delegates the comparison
    to compare_sheets_by_ws. Optionally, validators can be provided per column
    as a list (validator_arr), a dictionary (validator_dict), or a fallback (default_validator).

    Args:
        file1 (str): Path to the file with measured values.
        sheet1 (str): Name of the sheet in file1 to compare.
        file2 (str): Path to the file with expected values.
        sheet2 (str): Name of the sheet in file2 to compare.
        validator_arr (list, optional): List of validators per column (index-based).
        validator_dict (dict, optional): Dictionary of validators per column (index or name).
        default_validator (TableValidator, optional): Fallback validator for unspecified columns.

    Returns:
        list: List of comparison results (ComparisonResult).
    """
    _, eng1 = load_engine(file1, sheet1)
    _, eng2 = load_engine(file2, sheet2)
    return compare_sheets_by_ws(eng1, eng2, validator_arr, validator_dict, default_validator)


def compare_sheets_by_ws(eng1: TableEngine, eng2: TableEngine,
                         validator_arr=None, validator_dict=None, default_validator=None,
                         consumer=None):
    """
    Compares two Excel worksheets row by row and column by column.

    Uses TableRowEnumerator for row-wise iteration.
    Validators are applied per column. Differences are collected as ComparisonResult.
    Structural differences such as longer or shorter rows are also documented.

    If a `consumer` object is provided, each row is passed directly to `consumer.diff(...)`
    after comparison. This enables memory-efficient processing of large tables or live streaming
    of comparison results.

    Args:
        eng1 (TableEngine): TableEngine with measured values.
        eng2 (TableEngine): TableEngine with expected values.
        validator_arr (list, optional): List of validators per column (index-based).
        validator_dict (dict, optional): Dictionary of validators per column (index or name).
        default_validator (TableValidator, optional): Fallback validator for unspecified columns.
        consumer (object, optional): Object with a diff() method for row-wise processing.

    Returns:
        list: List of all comparison results if no consumer is set.
    """
    validator_arr = calculate_validator_array(eng2, validator_arr, validator_dict, default_validator)
    enum1 = TableRowEnumerator(eng1)
    enum2 = TableRowEnumerator(eng2)
    return compare_sheets_by_enum(enum1, enum2, validator_arr=validator_arr, consumer=consumer)


def compare_sheets_by_enum(enum1: TableRowEnumerator, enum2: TableRowEnumerator,
                           validator_arr=None, consumer=None):
    """
    Compares two TableRowEnumerators row by row and column by column.

    Each row is compared using compare_a_row. If a consumer is set,
    each comparison row is passed directly. Otherwise, a complete list is returned.

    Args:
        enum1 (TableRowEnumerator): Enumerator over measured values.
        enum2 (TableRowEnumerator): Enumerator over expected values.
        validator_arr (list, optional): List of validators per column.
        consumer (object, optional): Object with a diff() method for row-wise processing.

    Returns:
        list: List of all comparison results, or None in consumer mode.
    """
    max_rows = max(enum1.get_max_row(), enum2.get_max_row())
    all_differences = [] if consumer is None else None

    # Prepare default validators for the first row comparison
    validator_arr_nur_str = calculate_validator_array(enum2.engine, None, None, EqualValidator())
    compare_next(1, enum1, enum2, validator_arr_nur_str, consumer, all_differences)

    # Compare remaining rows
    for r in range(2, max_rows + 1):
        compare_next(r, enum1, enum2, validator_arr, consumer, all_differences)

    return all_differences


def compare_next(r: int, enum1: TableRowEnumerator, enum2: TableRowEnumerator, validator_arr, consumer: Any | None,
                 all_differences: list[Any] | None):
    """
    Compares the next row from two enumerators.

    If a row is missing, it is treated as empty. Differences are either collected
    or passed to a consumer.

    Args:
        r (int): Row number.
        enum1 (TableRowEnumerator): Enumerator over measured values.
        enum2 (TableRowEnumerator): Enumerator over expected values.
        validator_arr (list): List of validators per column.
        consumer (object, optional): Object with a diff() method for row-wise processing.
        all_differences (list, optional): List to collect all differences (if no consumer is set).
    """
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
    Compares two rows cell by cell using the provided validators.

    Args:
        row1 (list): Row of measured values.
        row2 (list): Row of reference values.
        validator_arr (list): List of validators per column.

    Returns:
        list: List of ComparisonResult values per cell.
    """
    differences = []
    for c, (val1, val2) in enumerate(zip(row1, row2)):
        v = validator_arr[c]
        differences.append(v.compare(val1, val2) if v else ComparisonResult.OMITTED)

    # Handle row length differences
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
    Creates a complete validator array for column comparison.

    Combines the provided validator list with a dictionary that references
    either column indices or column names. Missing entries are filled with the default validator.

    Args:
        eng2 (TableEngine): Engine with expected values (for column name access).
        validator_arr (list, optional): List of validators per column (index-based).
        validator_dict (dict, optional): Dictionary of validators per column (index or name).
        default_validator (TableValidator, optional): Fallback validator for unspecified columns.

    Returns:
        list: Complete list of validators per column.
    """
    max_cols = eng2.get_max_col()

    # Initialize validator array
    if validator_arr is None:
        validator_arr = [default_validator] * max_cols
    elif len(validator_arr) < max_cols:
        validator_arr.extend([default_validator] * (max_cols - len(validator_arr)))

    # Override with dictionary values
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
                    pass  # Column name not found

    return validator_arr