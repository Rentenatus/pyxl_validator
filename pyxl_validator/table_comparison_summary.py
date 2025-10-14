"""
table_comparison_summary.py

<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Eclipse Public License v2.0 which accompanies this distribution, and is available at
http://www.eclipse.org/legal/epl-v20.html
</copyright>

Defines the ComparisonSummary class for evaluating cell comparisons.
Uses ComparisonResult from table_validator.py.
"""

from collections import defaultdict
from typing import Any, Tuple, List, Dict
from pyxl_validator.table_validator import ComparisonResult


class ComparisonSummary:
    """
    Collects and analyzes cell comparison results.

    This class stores the results of cell-by-cell comparisons between two tables.
    It provides methods to add results, count occurrences of specific result types,
    generate summaries, and retrieve detailed information about compared cells.
    """

    def __init__(self):
        """
        Initializes a new ComparisonSummary instance.

        Attributes:
            results (Dict[ComparisonResult, List[Tuple[int, int, Any, Any]]]):
                Maps each ComparisonResult to a list of tuples containing row, column, and cell values.
            header_values (List[Any]):
                Stores header values for column-wise summaries.
        """
        self.results: Dict[ComparisonResult, List[Tuple[int, int, Any, Any]]] = defaultdict(list)
        self.header_values = []

    def add(self, row: int, col: int, val1: Any, val2: Any, result: ComparisonResult):
        """
        Adds a cell comparison result.

        Args:
            row (int): Row number (1-based).
            col (int): Column number (1-based).
            val1 (Any): Value from the first table.
            val2 (Any): Value from the second table.
            result (ComparisonResult): The result of the comparison.
        """
        self.results[result].append((row, col, val1, val2))

    def count(self, result_type: ComparisonResult) -> int:
        """
        Returns the number of results of a specific type.

        Args:
            result_type (ComparisonResult): The type of comparison result to count.

        Returns:
            int: Number of occurrences of the specified result type.
        """
        return len(self.results[result_type])

    def total(self) -> int:
        """
        Returns the total number of compared cells.

        Returns:
            int: Total count of all compared cells.
        """
        return sum(len(lst) for lst in self.results.values())

    def summary(self) -> Dict[str, int]:
        """
        Returns a summary as a dictionary:
        { "MATCHING": 42, "DIFFERENT": 7, ... }

        Returns:
            Dict[str, int]: Mapping of result type names to their counts.
        """
        return {res.name: len(lst) for res, lst in self.results.items()}

    def get_cells(self, result_type: ComparisonResult) -> List[Tuple[int, int, Any, Any]]:
        """
        Returns all cells of a specific comparison result type.

        Args:
            result_type (ComparisonResult): The type of comparison result.

        Returns:
            List[Tuple[int, int, Any, Any]]: List of tuples with row, column, and cell values.
        """
        return self.results.get(result_type, [])

    def __str__(self):
        """
        Returns a string representation of the summary.

        Returns:
            str: Human-readable summary of comparison results.
        """
        lines = [f"{res.name}: {len(lst)}" for res, lst in self.results.items()]
        return "Comparison Summary:\n" + "\n".join(lines)

    def set_header_values(self, header_values: List[Any]):
        """
        Sets the header values for column-wise summaries.

        Args:
            header_values (List[Any]): List of header values.
        """
        self.header_values = header_values

    def summary_by_header_array(self) -> List[Dict[str, int]]:
        """
        Returns a column-wise summary as an array.

        Return format:
            [
                { "MATCHING": 12, "DIFFERENT": 3, ... },  # Column 1
                { "MATCHING": 7, "CORRUPTED": 2, ... },   # Column 2
                ...
            ]

        The length of the array matches the number of header_values.
        Prerequisite: header_values must be set.

        Returns:
            List[Dict[str, int]]: List of dictionaries, each representing a column summary.

        Raises:
            ValueError: If header_values are not set.
        """
        if not self.header_values:
            raise ValueError("header_values must be set before calling summary_by_header_array.")

        summary_array = [defaultdict(int) for _ in self.header_values]

        for result, cells in self.results.items():
            for row, col, _, _ in cells:
                if 1 <= col <= len(summary_array):
                    summary_array[col - 1][result.name] += 1

        return [dict(counts) for counts in summary_array]