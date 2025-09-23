"""
__init__.py

pyxl_validator – Type-safe, semantically extensible Excel validator.

<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Eclipse Public License v2.0 which accompanies this distribution, and is available at
http://www.eclipse.org/legal/epl-v20.html
</copyright>

This package provides a modular architecture for processing and validating
Excel files (.xls, .xlsx, .ods) with a focus on:
- Type safety
- Semantic enums
- Sophisticated error classification
- Sustainable data migration

Developed for use in insurance-related, public, and collaborative contexts.
"""

__doc__ = "pyxl_validator – Type-safe, semantically extensible Excel validator."
__version__ = "0.1.0"
__author__ = "Janusch Rentenatus"
__license__ = "Apache-2.0"

from .excel_table_engine import TableEngine, TableEnginePyxl, TableEnginePyexcel, TableRowEnumerator, load_engine
from .table_validator_registry import ValidatorRegistry
from .table_validator import  (TableValidator, EqualValidator, BoolValidator, IntValidator, NumberValidator,
                               TolerantFloatValidator, ExcelValueValidator, OmittedValidator, ComparisonResult,
                               _is_bool_like, _is_int_then_normalize, _is_float_then_normalize,
                               _is_date_then_normalize, _is_number_then_normalize)
from .table_comparison_summary import ComparisonSummary
from .excel_compare import compare_sheets_by_file, compare_sheets_by_ws, compare_sheets_by_enum, calculate_validator_array
from .excel_differator import differentiate_sheets_by_ws, DiffConsumer

__all__ = [
    "TableEngine",
    "TableEnginePyxl",
    "TableEnginePyexcel",
    "TableRowEnumerator",
    "load_engine",

    "ValidatorRegistry",

    "TableValidator",
    "EqualValidator",
    "BoolValidator",
    "IntValidator",
    "NumberValidator",
    "TolerantFloatValidator",
    "ExcelValueValidator",
    "OmittedValidator",
    "ComparisonResult",
    "_is_bool_like",
    "_is_int_then_normalize",
    "_is_float_then_normalize",
    "_is_date_then_normalize",
    "_is_number_then_normalize",

    "ComparisonSummary",

    "compare_sheets_by_file",
    "compare_sheets_by_ws",
    "compare_sheets_by_enum",
    "calculate_validator_array",

    "differentiate_sheets_by_ws",
    "DiffConsumer",
]