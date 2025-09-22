"""
excel_differator.py

Vergleicht zwei Excel-Tabellen zeilenweise und dokumentiert Unterschiede visuell und statistisch.

Funktionalität:
    - DiffConsumer: verarbeitet Vergleichszeilen, markiert Zellen farblich
      und fügt bei Abweichungen Messwerte in die Referenztabelle ein
    - Optional: Erfassung aller fehlerhaften Zellen in einer ComparisonSummary

Verwendete Komponenten:
    - TableEngine: abstraktes Tabellen-Interface
    - TableRowEnumerator: Iterator über Tabellenzeilen
    - ComparisonResult: Enum für Vergleichstypen
    - ComparisonSummary: Sammlung und Auswertung von Vergleichsergebnissen
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
    Führt einen zeilenweisen Vergleich zweier Tabellen durch.

    Verwendet die ValidatorRegistry zur Spaltenvalidierung und übergibt
    die Vergleichsergebnisse an einen DiffConsumer. Dieser markiert Zellen
    farblich und dokumentiert fehlerhafte Zellen in der Summary.

    :param eng1: Tabelle mit Messwerten.
    :param eng2: Tabelle mit Referenzwerten (wird farblich markiert und ggf. ergänzt).
    :param registry: Registry zur Zuordnung von Validatoren zu Spalten.
    :param summary: Optionales Objekt zur Sammlung fehlerhafter Zellen.
    :return: Ergebnis des Vergleichs (z. B. None bei Consumer-Modus).
    """
    values_row1 = eng2.get_row_values(1)
    validator_arr = registry.resolve_validators(values_row1)
    if summary:
        summary.set_header_values(values_row1)
    return DiffConsumer(eng1, eng2, summary).compare_sheets_consume_diff(validator_arr)


class DiffConsumer:
    """
    Konsument für Vergleichsergebnisse – verarbeitet jede Vergleichszeile.

    - Markiert Zellen in eng2 farblich gemäß ComparisonResult
    - Fügt bei Abweichungen die Messwerte aus eng1 zusätzlich in eng2 ein
    - Dokumentiert fehlerhafte Zellen in einer ComparisonSummary
    """

    def __init__(self, eng1: TableEngine, eng2: TableEngine, summary: ComparisonSummary):
        """
        Initialisiert den Consumer mit zwei Tabellen und einer optionalen Summary.

        :param eng1: Tabelle mit Messwerten.
        :param eng2: Tabelle mit Referenzwerten.
        :param summary: Objekt zur Sammlung fehlerhafter Zellen.
        """
        self.enum1 = TableRowEnumerator(eng1)
        self.enum2 = TableRowEnumerator(eng2)
        self.eng2 = eng2
        self.summary = summary

    def compare_sheets_consume_diff(self, validator_arr):
        """
        Startet den Vergleich über compare_sheets_by_enum.

        :param validator_arr: Liste von Validatoren pro Spalte.
        :return: Vergleichsergebnis (z.B. None bei Consumer-Modus).
        """
        return compare_sheets_by_enum(self.enum1, self.enum2,
                                      validator_arr=validator_arr, consumer=self)

    def diff(self, r: int, index1: int, row1: list[Any], index2: int, row2: list[Any],
             differences: list[ComparisonResult]):
        """
        Verarbeitet eine Vergleichszeile.

        - Markiert die Referenzzeile farblich.
        - Fügt bei Abweichung die Messzeile zusätzlich ein.
        - Dokumentiert fehlerhafte Zellen in der Summary.

        :param r: Laufende Vergleichszeile (1-basiert).
        :param index1: Zeilenindex in eng1.
        :param row1: Werte aus eng1.
        :param index2: Zeilenindex in eng2.
        :param row2: Werte aus eng2.
        :param differences: Liste von ComparisonResult pro Zelle.
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





