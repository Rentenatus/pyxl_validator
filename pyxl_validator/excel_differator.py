"""
excel_differator.py

<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Eclipse Public License v2.0 which accompanies this distribution, and is available at
http://www.eclipse.org/legal/epl-v20.html
</copyright>

Compares two Excel tables row by row and documents differences both visually and statistically.

Functionality:
    - DiffConsumer: Processes comparison rows, highlights cells with colors,
      and inserts measured values into the reference table for discrepancies.
    - Optional: Collects all erroneous cells in a ComparisonSummary.

Used components:
    - TableEngine: Abstract table interface
    - TableRowEnumerator: Iterator over table rows
    - ComparisonResult: Enum for comparison types
    - ComparisonSummary: Collection and evaluation of comparison results
"""

from typing import Any

from pyxl_validator.table_validator import ComparisonResult
from pyxl_validator.table_validator_registry import ValidatorRegistry
from pyxl_validator.excel_compare import compare_sheets_by_enum
from pyxl_validator.excel_table_engine import TableEngine, TableRowEnumerator
from pyxl_validator.table_comparison_summary import ComparisonSummary


def differentiate_sheets_by_ws(eng1: TableEngine, eng2: TableEngine,
                               registry: ValidatorRegistry, summary: ComparisonSummary):
    """
    Performs a row-by-row comparison of two tables.

    Uses the ValidatorRegistry for column validation and passes
    the comparison results to a DiffConsumer. The consumer highlights cells
    and documents erroneous cells in the summary.

    Args:
        eng1 (TableEngine): Table with measured values.
        eng2 (TableEngine): Table with reference values (will be highlighted and possibly extended).
        registry (ValidatorRegistry): Registry for assigning validators to columns.
        summary (ComparisonSummary): Optional object for collecting erroneous cells.

    Returns:
        Any: Result of the comparison (e.g., None in consumer mode).
    """
    values_row1 = eng2.get_row_values(1)
    validator_arr = registry.resolve_validators(values_row1)
    if summary:
        summary.set_header_values(values_row1)
    return DiffConsumer(eng1, eng2, summary).compare_sheets_consume_diff(validator_arr)

# ------------------------------------------------------------
# DiffConsumer
# ------------------------------------------------------------

class DiffConsumer:
    """
    Consumer for comparison results – processes each comparison row.

    - Highlights cells in eng2 according to ComparisonResult.
    - Inserts measured values from eng1 into eng2 for discrepancies.
    - Documents erroneous cells in a ComparisonSummary.
    """

    def __init__(self, eng1: TableEngine, eng2: TableEngine, summary: ComparisonSummary):
        """
        Initializes the consumer with two tables and an optional summary.

        Args:
            eng1 (TableEngine): Table with measured values.
            eng2 (TableEngine): Table with reference values.
            summary (ComparisonSummary): Object for collecting erroneous cells.
        """
        self.enum1 = TableRowEnumerator(eng1)
        self.enum2 = TableRowEnumerator(eng2)
        self.eng2 = eng2
        self.summary = summary

    def compare_sheets_consume_diff(self, validator_arr):
        """
        Starts the comparison using compare_sheets_by_enum.

        Args:
            validator_arr (list): List of validators per column.

        Returns:
            Any: Comparison result (e.g., None in consumer mode).
        """
        return compare_sheets_by_enum(self.enum1, self.enum2,
                                      validator_arr=validator_arr, consumer=self)

    def diff(self, r: int, index1: int, row1: list[Any], index2: int, row2: list[Any],
             differences: list[ComparisonResult]):
        """
        Processes a comparison row.

        - Highlights the reference row with colors.
        - Inserts the measured row if there are discrepancies.
        - Documents erroneous cells in the summary.

        Args:
            r (int): Current comparison row (1-based).
            index1 (int): Row index in eng1.
            row1 (list): Values from eng1.
            index2 (int): Row index in eng2.
            row2 (list): Values from eng2.
            differences (list[ComparisonResult]): List of ComparisonResult per cell.
        """
        okay = True
        formats_ref = []
        formats_mess = []

        for c, result in enumerate(differences):
            okay = okay and result.ok()
            fg_ref, fg_mess = result.get_cell_colors()
            formats_ref.append({"fill_color": fg_ref})
            formats_mess.append({"fill_color": fg_mess})

            if self.summary and result.foul():
                val1 = row1[c] if c < len(row1) else None
                val2 = row2[c] if c < len(row2) else None
                self.summary.add(r, c + 1, val1, val2, result)

        if index2 > 0:
            self.eng2.set_row_formats(index2, formats_ref)

        if not okay and row1:
            new_row = self.enum2.add_row(row1)
            self.eng2.set_row_formats(new_row, formats_mess)

