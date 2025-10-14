"""
table_validator_registry.py

<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Eclipse Public License v2.0 which accompanies this distribution, and is available at
http://www.eclipse.org/legal/epl-v20.html
</copyright>
"""

from typing import Dict, List
from pyxl_validator.table_validator import TableValidator

# ------------------------------------------------------------
# ValidatorRegistry for a single table
# ------------------------------------------------------------

class ValidatorRegistry:
    """
    Manages the assignment of validators to column names or indices for a worksheet.

    This registry allows registering validators for specific columns by name or index,
    as well as setting a default validator for columns without explicit assignment.
    """

    def __init__(self):
        self.by_column_name: Dict[str, TableValidator] = {}
        self.by_column_index: Dict[int, TableValidator] = {}
        self.default_validator: TableValidator = None

    def register_by_name(self, column_name: str, validator: TableValidator):
        """
        Registers a validator for a column by its name.

        :param column_name: The name of the column.
        :param validator: The TableValidator instance to use for this column.
        """
        self.by_column_name[column_name] = validator

    def register_by_index(self, column_index: int, validator: TableValidator):
        """
        Registers a validator for a column by its index.

        :param column_index: The zero-based index of the column.
        :param validator: The TableValidator instance to use for this column.
        """
        self.by_column_index[column_index] = validator

    def set_default(self, validator: TableValidator):
        """
        Sets the default validator to use for columns without explicit assignment.

        :param validator: The TableValidator instance to use as default.
        """
        self.default_validator = validator

    def get_validator(self, column_name: str = None, column_index: int = None) -> TableValidator:
        """
        Returns the validator for a given column name or index.

        :param column_name: The name of the column (optional).
        :param column_index: The index of the column (optional).
        :return: The assigned TableValidator, or the default if none is registered.
        """
        if column_name in self.by_column_name:
            return self.by_column_name[column_name]
        if column_index in self.by_column_index:
            return self.by_column_index[column_index]
        return self.default_validator

    def resolve_validators(self, header_row: List[str], max_col: int = 1) -> List[TableValidator]:
        """
        Converts a header row into a complete list of validators for each column.

        Validators are assigned by column name or index from the header row.
        Explicitly registered index-based validators are added, extending the list if necessary.

        :param header_row: List of column names from the first row of the worksheet.
        :return: List of TableValidator instances for each column.
        """
        validators = [self.default_validator]*max_col
        for index, col_name in enumerate(header_row):
            validator = self.get_validator(column_name=col_name, column_index=index)
            validators[index] = validator

        # Add all explicitly registered index-based validators
        for index, validator in self.by_column_index.items():
            if index < len(validators):
                validators[index] = validator
            else:
                # Extend the list with default validators if needed
                while len(validators) < index:
                    validators.append(self.default_validator)
                validators.append(validator)

        return validators

# ------------------------------------------------------------
# RegistryStore for multiple sheets
# ------------------------------------------------------------

class RegistryStore:
    """
    Manages multiple ValidatorRegistry instances, each assigned to a worksheet name.

    This store allows registering and retrieving validator registries for different sheets,
    and setting a default registry for sheets without explicit assignment.
    """

    def __init__(self):
        self.by_sheet_name: Dict[str, ValidatorRegistry] = {}
        self.default_registry: ValidatorRegistry = None

    def register(self, sheet_name: str, registry: ValidatorRegistry):
        """
        Registers a ValidatorRegistry for a specific worksheet.

        :param sheet_name: The name of the worksheet.
        :param registry: The ValidatorRegistry instance to assign.
        """
        self.by_sheet_name[sheet_name] = registry

    def set_default(self, registry: ValidatorRegistry):
        """
        Sets the default ValidatorRegistry for worksheets without explicit assignment.

        :param registry: The ValidatorRegistry instance to use as default.
        """
        self.default_registry = registry

    def get_registry(self, sheet_name: str) -> ValidatorRegistry:
        """
        Returns the ValidatorRegistry for the given worksheet name.

        If no registry is registered for the sheet, the default registry is returned.

        :param sheet_name: The name of the worksheet.
        :return: The assigned ValidatorRegistry, or the default if none is registered.
        """
        return self.by_sheet_name.get(sheet_name, self.default_registry)