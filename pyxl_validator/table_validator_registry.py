from typing import Dict, List
from typing import Dict
from pyxl_validator.table_validator import TableValidator


# ------------------------------------------------------------
# ValidatorRegistry für eine Tabelle
# ------------------------------------------------------------

class ValidatorRegistry:
    """
    Verwaltet die Zuordnung von Validatoren zu Spaltennamen oder Indizes.
    """

    def __init__(self):
        self.by_column_name: Dict[str, TableValidator] = {}
        self.by_column_index: Dict[int, TableValidator] = {}
        self.default_validator: TableValidator = None

    def register_by_name(self, column_name: str, validator: TableValidator):
        self.by_column_name[column_name] = validator

    def register_by_index(self, column_index: int, validator: TableValidator):
        self.by_column_index[column_index] = validator

    def set_default(self, validator: TableValidator):
        self.default_validator = validator

    def get_validator(self, column_name: str = None, column_index: int = None) -> TableValidator:
        if column_name in self.by_column_name:
            return self.by_column_name[column_name]
        if column_index in self.by_column_index:
            return self.by_column_index[column_index]
        return self.default_validator

    def resolve_validators(self, header_row: List[str]) -> List[TableValidator]:
        """
        Wandelt eine Kopfzeile in eine vollständige Liste von Validatoren um.

        Zuerst werden Validatoren anhand von Spaltennamen oder Indizes aus der Kopfzeile zugewiesen.
        Anschließend werden alle explizit registrierten Index-Validatoren ergänzt – auch für Spalten,
        die nicht in der Kopfzeile stehen. Die Liste wird bei Bedarf verlängert.

        :param header_row: Liste der Spaltennamen aus Zeile 1.
        :return: Vollständige Liste von Validatoren pro Spalte.
        """
        validators = []
        for i, col_name in enumerate(header_row):
            validator = self.get_validator(column_name=col_name, column_index=i)
            validators.append(validator)

        # Ergänze alle explizit registrierten Index-Validatoren
        for index, validator in self.by_column_index.items():
            if index < len(validators):
                validators[index] = validator
            else:
                # Liste verlängern mit Default-Validatoren, falls nötig
                while len(validators) < index:
                    validators.append(self.default_validator)
                validators.append(validator)

        return validators


# ------------------------------------------------------------
# RegistryStore für mehrere Sheets
# ------------------------------------------------------------

class RegistryStore:
    """
    Verwaltet mehrere ValidatorRegistries, jeweils zugeordnet zu einem Worksheet-Namen.
    """

    def __init__(self):
        self.by_sheet_name: Dict[str, ValidatorRegistry] = {}
        self.default_registry: ValidatorRegistry = None

    def register(self, sheet_name: str, registry: ValidatorRegistry):
        """Registriert eine Registry für ein bestimmtes Tabellenblatt."""
        self.by_sheet_name[sheet_name] = registry

    def set_default(self, registry: ValidatorRegistry):
        """Setzt die Standard-Registry für nicht explizit registrierte Sheets."""
        self.default_registry = registry

    def get_registry(self, sheet_name: str) -> ValidatorRegistry:
        """Gibt die passende Registry für das Tabellenblatt zurück."""
        return self.by_sheet_name.get(sheet_name, self.default_registry)
