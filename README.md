# pyxl_validator

**Excel Validation via openpyxl — Compare and Highlight Differences Between Expected and Calculated Tables**


`pyxl_validator` is a compact Python toolkit designed to compare Excel worksheets and detect discrepancies quickly and reliably. It automatically recognizes common Excel value types such as numbers, dates, and booleans, applies configurable validators per column, and compares tables row by row. Differences are classified into meaningful result types and can be highlighted directly within the reference workbook.

Particularly helpful for migrating code (e.g. VBA to Python) and creating automatic unit tests.

## Key Features

- Automatic recognition of value types (numbers, dates, booleans)
- Configurable column-specific validators
- Row-by-row and cell-by-cell comparison
- Tolerant floating-point comparisons with adjustable rounding and tolerance
- Locale-aware parsing (e.g., German number and currency formats)
- Configurable date precision (day/hour/minute/second)
- Color-coded results for immediate visual feedback
- Summary collector for reporting and automated checks
- Modular architecture for custom table engines and validator registries

Use `pyxl_validator` when you need reliable, programmatic validation of measured data against reference tables, visual feedback in Excel files, and compact summaries for QA or integration pipelines.


## comparsion

pyxl_validator's comparison component lets you compare two worksheets row-by-row and cell-by-cell with predictable, configurable rules. It automatically applies column-specific validators (or a default validator) to each cell, classifies differences into meaningful result types, and supports tolerant numeric comparisons, configurable date precision, and handling of locale formats (e.g. German decimals/currencies). 

A concise list of available `TableValidator` classes in this project and what they do.

- `TableValidator`  
  Abstract base interface for all cell comparison validators.

- `EqualValidator`  
  Simple equality check using `==`. Returns `EQUALS` or `DIFFERENT`.

- `BoolValidator`  
  Compares boolean-like values (e.g. `TRUE`/`FALSE`, `WAHR`/`FALSCH`, `1`/`0`) with normalization.

- `DateValidator`  
  Compares date/time values with configurable precision (`day`, `hour`, `minute`, `second`).

- `OmittedValidator`  
  Marks a cell comparison as intentionally omitted (`OMITTED`).

- `IgnoreValidator`  
  Accepts all values as matching (`MATCHING`) and effectively ignores the cell.

- `IntValidator`  
  Integer-aware comparison (int, float-as-int, numeric strings) returning `EQUALS`, `MATCHING` or `DIFFERENT`.

- `NumberValidator`  
  General numeric comparison with integer and floating point handling, configurable rounding for `ALMOST`.

- `TolerantFloatValidator`  
  Floating point comparison with asymmetric tolerance window (`delta_up`, `delta_down`) and precision control.

- `ExcelValueValidator`  
  High-level validator that auto-detects type (date, number, bool) and delegates to the appropriate validator.


## Supported Table Engines

| Engine              | Description                                                                 |
|---------------------|-----------------------------------------------------------------------------|
| `TableEnginePyxl`   | Full read/write support for `.xlsx` via `openpyxl`, including styles/formats |
| `TableEnginePyexcel`| Read-only access for legacy `.xls` and `.ods` via `pyexcel`; no formatting   |
| `TableEnginePandas` | In-memory read/write via `pandas` DataFrame with optional format storage     |

### Utilities and Entry Points

- `load_engine`
- `get_pandas_engine`
- `load_pandas_engine`
- `copy_to_pandas`
- `TableRowEnumerator` — supports row iteration and dynamic row insertion

## Example

Compare a table of measured values against a reference and highlight differences in the reference sheet:

```python
wb1, eng_reference = load_engine("test/assets/expected/e-daten1.xlsx", sheet_name="Tabelle1")
_, eng_measured = load_engine("test/assets/input/daten1.xlsx", sheet_name="Tabelle1")

# wb1 will be changed and will become a diff file:
differentiate_sheets_by_ws(
    eng_measured,
    eng_reference,
    has_header=True,
    registry=...,
    summary=...
)
wb1.save("test/tmp/v-daten1-xlsx.xlsx")
```

# License
Apache-2.0 license.

# Notes
- The API is intentionally modular; custom TableEngine implementations or registry configurations are possible.
- For problems with German formats: The constant GERMAN_DEZIM controls the recognition behavior of numbers.


