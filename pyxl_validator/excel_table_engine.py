"""
excel_table_engine.py

<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Eclipse Public License v2.0 which accompanies this distribution, and is available at
http://www.eclipse.org/legal/epl-v20.html
</copyright>

Provides abstract and concrete table engines for reading and writing Excel files (.xlsx, .xls, .ods).
Includes row iteration and a factory for engine instantiation.
"""
from abc import ABC, abstractmethod
import os
from typing import Any

import openpyxl

class TableEngine(ABC):
    """
    Abstract interface for table engines supporting Excel-like worksheets.

    Defines methods for reading, writing, and formatting cell and row data.
    """

    # Getter
    @abstractmethod
    def get_max_row(self) -> int:
        """
        Returns the maximum number of rows in the worksheet.
        """
        pass

    @abstractmethod
    def get_max_col(self) -> int:
        """
        Returns the maximum number of columns in the worksheet.
        """
        pass

    @abstractmethod
    def get_cell_value(self, row: int, col: int):
        """
        Returns the value of the cell at the given row and column.
        """
        pass

    @abstractmethod
    def get_row_values(self, row: int) -> list:
        """
        Returns a list of all cell values in the given row.
        """
        pass

    @abstractmethod
    def get_cell_format(self, row: int, col: int) -> dict:
        """
        Returns the format of the cell at the given row and column as a dictionary.
        """
        pass

    @abstractmethod
    def get_row_formats(self, row: int) -> list:
        """
        Returns a list of format dictionaries for all cells in the given row.
        """
        pass

    # Tester
    @abstractmethod
    def is_readonly(self) -> bool:
        """
        Returns True if the underlying file is read-only.
        """
        pass

    def is_engine_readonly(self) -> bool:
        """
        Returns True if the engine implementation is read-only.
        """
        pass

    # Setter
    @abstractmethod
    def set_cell_value(self, row: int, col: int, value):
        """
        Sets the value of the cell at the given row and column.
        """
        pass

    @abstractmethod
    def add_row(self, row: int):
        """
        Inserts a new row at the given position.
        """
        pass

    @abstractmethod
    def set_row_values(self, row: int, values: list):
        """
        Sets the values of all cells in the given row.
        """
        pass

    @abstractmethod
    def set_cell_format(self, row: int, col: int, fmt: dict):
        """
        Sets the format of the cell at the given row and column.
        """
        pass

    @abstractmethod
    def set_row_formats(self, row: int, formats: list):
        """
        Sets the formats of all cells in the given row.
        """
        pass


class TableEnginePyxl(TableEngine):
    """
    Table engine implementation for .xlsx files using openpyxl.

    Supports reading and writing cell values and formats.
    """

    def __init__(self, ws):
        self.ws = ws

    def get_max_row(self) -> int:
        return self.ws.max_row

    def get_max_col(self) -> int:
        return self.ws.max_column

    def get_cell_value(self, row: int, col: int):
        return self.ws.cell(row=row, column=col).value

    def get_row_values(self, row: int) -> list:
        return [self.ws.cell(row=row, column=c).value for c in range(1, self.get_max_col() + 1)]

    def get_cell_format(self, row: int, col: int) -> dict:
        """
        Returns font and fill information for the cell as a dictionary.
        """
        cell = self.ws.cell(row=row, column=col)
        font = cell.font
        fill = cell.fill
        return {
            "font_name": font.name,
            "font_size": font.size,
            "bold": font.bold,
            "italic": font.italic,
            "font_color": font.color.rgb if font.color and font.color.type == "rgb" else None,
            "fill_color": fill.fgColor.rgb if fill and fill.fgColor.type == "rgb" else None
        }

    def get_row_formats(self, row: int) -> list:
        return [self.get_cell_format(row, c) for c in range(1, self.get_max_col() + 1)]

    def is_readonly(self) -> bool:
        return self.ws.parent.read_only

    def is_engine_readonly(self) -> bool:
        return False

    def set_cell_value(self, row: int, col: int, value):
        self.ws.cell(row=row, column=col).value = value

    def add_row(self, row: int):
        self.ws.insert_rows(row)

    def set_row_values(self, row: int, values: list):
        for c, val in enumerate(values, start=1):
            self.set_cell_value(row, c, val)

    def set_cell_format(self, row: int, col: int, fmt: dict):
        """
        Sets font and fill for the cell using openpyxl styles.
        """
        from openpyxl.styles import Font, PatternFill
        cell = self.ws.cell(row=row, column=col)
        cell.font = Font(
            name=fmt.get("font_name", cell.font.name),
            size=fmt.get("font_size", cell.font.size),
            bold=fmt.get("bold", cell.font.bold),
            italic=fmt.get("italic", cell.font.italic),
            color=fmt.get("font_color", cell.font.color)
        )
        if fmt.get("fill_color"):
            cell.fill = PatternFill(start_color=fmt["fill_color"], end_color=fmt["fill_color"], fill_type="solid")

    def set_row_formats(self, row: int, formats: list):
        for c, fmt in enumerate(formats, start=1):
            self.set_cell_format(row, c, fmt)


class TableEnginePyexcel(TableEngine):
    """
    Table engine implementation for .xls and .ods files using pyexcel.

    Only supports reading. Writing and formatting are not supported and will raise NotImplementedError.
    """

    def __init__(self, sheet):
        self.sheet = sheet

    def get_max_row(self) -> int:
        return self.sheet.number_of_rows()

    def get_max_col(self) -> int:
        return self.sheet.number_of_columns()

    def get_cell_value(self, row: int, col: int):
        """
        Returns the value of the cell at the given row and column (1-based).
        """
        try:
            return self.sheet[row - 1, col - 1]
        except (IndexError, KeyError):
            return None

    def get_row_values(self, row: int) -> list:
        """
        Returns a list of all cell values in the given row (1-based).
        """
        try:
            raw_row = self.sheet.row[row - 1]
            return list(raw_row)
        except (IndexError, KeyError):
            return []

    def get_cell_format(self, row: int, col: int) -> dict:
        """
        Returns a default format dictionary, as pyexcel does not support formatting.
        """
        return {"number_format": "General"}

    def get_row_formats(self, row: int) -> list:
        return [{"number_format": "General"} for _ in range(self.get_max_col())]

    def is_readonly(self) -> bool:
        return True

    def is_engine_readonly(self) -> bool:
        return True

    def set_cell_value(self, row: int, col: int, value):
        raise NotImplementedError("pyexcel does not support writing cell values.")

    def add_row(self, row: int):
        raise NotImplementedError("pyexcel does not support adding rows.")

    def set_row_values(self, row: int, values: list):
        raise NotImplementedError("pyexcel does not support writing row values.")

    def set_cell_format(self, row: int, col: int, fmt: dict):
        raise NotImplementedError("pyexcel does not support cell formatting.")

    def set_row_formats(self, row: int, formats: list):
        raise NotImplementedError("pyexcel does not support row formatting.")

# ============================================================
# Iteration
# ============================================================

class TableRowEnumerator:
    """
    Iterator for TableEngine rows.

    Allows row-wise iteration over a worksheet using `next()` or `for row in ...`.
    Supports inserting a new row at the current position with `add_row(values)`.

    Each iteration yields a tuple `(row_index, row_values)`.

    Example:
        engine = load_engine("data.xlsx", "Measurements")
        for row_index, row_values in TableRowEnumerator(engine):
            print(f"Row {row_index}: {row_values}")

        enumerator = TableRowEnumerator(engine)
        while True:
            try:
                r, values = next(enumerator)
                # Processing...
                row_index = enum.add_row(["Measurement A", 42.0, True])
                print(f"Row {row_index} inserted.")
            except StopIteration:
                break
    """

    def __init__(self, engine: TableEngine, start_row: int = 1):
        self.engine = engine
        self.current = start_row
        self.max_row = engine.get_max_row()

    def __iter__(self):
        return self

    def __next__(self):
        if self.current > self.max_row:
            raise StopIteration
        row_values = self.engine.get_row_values(self.current)
        result = (self.current, row_values)
        self.current += 1
        return result

    def add_row(self, values: list) -> int:
        """
        Inserts a new row at the current position.

        Args:
            values (list): List of cell values for the new row.

        Returns:
            int: The index of the inserted row.
        """
        self.engine.add_row(self.current)
        self.engine.set_row_values(self.current, values)
        inserted_row = self.current
        self.current += 1
        self.max_row += 1
        return inserted_row

    def get_max_row(self):
        """
        Returns the current maximum row index.
        """
        return self.max_row

# ============================================================
# Factory
# ============================================================

def load_engine(file_path: str, sheet_name: str) -> tuple[Any, TableEngine]:
    """
    Loads an Excel file (.xlsx, .xls, .ods) and returns the workbook and the corresponding TableEngine.

    Args:
        file_path (str): Path to the Excel file.
        sheet_name (str): Name of the sheet to load.

    Returns:
        tuple: (Workbook, TableEngine) for the loaded file and sheet.

    Raises:
        ImportError: If required packages for .xls or .ods are not installed.
        ValueError: If the file extension is not supported.
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".xlsx":
        wb = openpyxl.load_workbook(file_path, data_only=False)
        return wb, TableEnginePyxl(wb[sheet_name])
    elif ext == ".xls":
        try:
            import pyexcel
        except ImportError:
            raise ImportError("The package 'pyexcel' is not installed. Install it with 'pip install pyexcel'.")
        try:
            import pyexcel_xls
        except ImportError:
            raise ImportError("The package 'pyexcel_xls' is not installed. Install it with 'pip install pyexcel_xls'.")
        wb = pyexcel.get_book(file_name=file_path)
        return wb, TableEnginePyexcel(wb.sheet_by_name(sheet_name))
    elif ext == ".ods":
        try:
            import pyexcel
        except ImportError:
            raise ImportError("The package 'pyexcel' is not installed. Install it with 'pip install pyexcel'.")
        try:
            import pyexcel_ods
        except ImportError:
            raise ImportError("The package 'pyexcel_ods' is not installed. Install it with 'pip install pyexcel_ods'.")
        wb = pyexcel.get_book(file_name=file_path)
        return wb, TableEnginePyexcel(wb.sheet_by_name(sheet_name))
    else:
        raise ValueError(f"Unsupported file extension: {ext}")