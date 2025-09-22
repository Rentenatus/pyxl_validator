"""
table_comparison_summary.py

<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Eclipse Public License v2.0 which accompanies this distribution, and is available at
http://www.eclipse.org/legal/epl-v20.html
</copyright>

Definiert die Klasse ComparisonSummary zur Auswertung von Zellvergleichen.
Verwendet ComparisonResult aus table_validator.py.
"""

from collections import defaultdict
from typing import Any, Tuple, List, Dict
from pyxl_validator.table_validator import ComparisonResult


class ComparisonSummary:
    """
    Sammelt und analysiert Vergleichsergebnisse von Zellen.
    """

    def __init__(self):
        # Dict: ComparisonResult → Liste von (row, col, val1, val2)
        self.results: Dict[ComparisonResult, List[Tuple[int, int, Any, Any]]] = defaultdict(list)
        self.header_values = []

    def add(self, row: int, col: int, val1: Any, val2: Any, result: ComparisonResult):
        """
        Fügt ein Vergleichsergebnis hinzu.

        Args:
            row (int): Zeilennummer (1-basiert)
            col (int): Spaltennummer (1-basiert)
            val1, val2: Zellwerte
            result (ComparisonResult): Vergleichsergebnis
        """
        self.results[result].append((row, col, val1, val2))

    def count(self, result_type: ComparisonResult) -> int:
        """Gibt die Anzahl der Ergebnisse eines bestimmten Typs zurück."""
        return len(self.results[result_type])

    def total(self) -> int:
        """Gibt die Gesamtanzahl aller verglichenen Zellen zurück."""
        return sum(len(lst) for lst in self.results.values())

    def summary(self) -> Dict[str, int]:
        """
        Gibt eine Zusammenfassung als Dictionary zurück:
        { "MATCHING": 42, "DIFFERENT": 7, ... }
        """
        return {res.name: len(lst) for res, lst in self.results.items()}

    def get_cells(self, result_type: ComparisonResult) -> List[Tuple[int, int, Any, Any]]:
        """
        Gibt alle Zellen eines bestimmten Vergleichstyps zurück.
        """
        return self.results.get(result_type, [])

    def __str__(self):
        lines = [f"{res.name}: {len(lst)}" for res, lst in self.results.items()]
        return "Comparison Summary:\n" + "\n".join(lines)

    def set_header_values(self, header_values):
        self.header_values = header_values

    def summary_by_header_array(self) -> List[Dict[str, int]]:
        """
        Gibt eine spaltenweise Zusammenfassung als Array zurück.

        Rückgabeformat:
            [
                { "MATCHING": 12, "DIFFERENT": 3, ... },  # Spalte 1
                { "MATCHING": 7, "CORRUPTED": 2, ... },   # Spalte 2
                ...
            ]

        Die Länge des Arrays entspricht der Anzahl der header_values.
        Voraussetzung: header_values wurden gesetzt.
        """
        if not self.header_values:
            raise ValueError("header_values wurden nicht gesetzt.")

        summary_array = [defaultdict(int) for _ in self.header_values]

        for result, cells in self.results.items():
            for row, col, _, _ in cells:
                if 1 <= col <= len(summary_array):
                    summary_array[col - 1][result.name] += 1

        return [dict(counts) for counts in summary_array]


