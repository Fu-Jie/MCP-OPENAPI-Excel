"""
XlsxWriter adapter for high-performance Excel writing.

This module provides the XlsxWriterAdapter class that wraps XlsxWriter
for writing Excel files. XlsxWriter is optimized for creating large files
with minimal memory usage through streaming writes.

Features:
    - Streaming write for large files
    - Auto-formatting (column widths, data types)
    - Multiple sheet support
    - Formatting support (bold, colors, etc.)

Example:
    adapter = XlsxWriterAdapter()
    adapter.write_sheet(
        "/path/to/output.xlsx",
        rows=[["Name", "Age"], ["Alice", 30], ["Bob", 25]],
        sheet_name="Users",
        headers=["Name", "Age"]
    )
"""

import os
from datetime import datetime
from pathlib import Path
from typing import Any

import xlsxwriter
from xlsxwriter.workbook import Workbook

from src.exceptions.excel_exceptions import PermissionError as ExcelPermissionError
from src.exceptions.excel_exceptions import WriteError


class XlsxWriterAdapter:
    """
    Adapter for XlsxWriter Excel writing operations.

    This adapter provides a clean interface for writing Excel files using
    the high-performance XlsxWriter library. It handles format conversion,
    error mapping, and provides auto-formatting capabilities.

    Attributes:
        DEFAULT_COLUMN_WIDTH: Default column width in characters.
        MAX_COLUMN_WIDTH: Maximum column width in characters.

    Example:
        adapter = XlsxWriterAdapter()

        # Write simple data
        result = adapter.write_sheet(
            "/path/to/output.xlsx",
            rows=[["Name", "Age"], ["Alice", 30]],
            sheet_name="Users"
        )

        # Write with headers and auto-formatting
        result = adapter.write_sheet(
            "/path/to/output.xlsx",
            rows=[["Alice", 30], ["Bob", 25]],
            headers=["Name", "Age"],
            auto_format=True
        )
    """

    DEFAULT_COLUMN_WIDTH = 10
    MAX_COLUMN_WIDTH = 50

    def __init__(self) -> None:
        """Initialize the XlsxWriterAdapter."""
        pass

    def _validate_output_path(
        self,
        file_path: str,
        overwrite: bool = False,
    ) -> Path:
        """
        Validate and prepare the output file path.

        Args:
            file_path: Path where the file will be written.
            overwrite: Whether to overwrite if the file exists.

        Returns:
            Path object for the output file.

        Raises:
            WriteError: If the file exists and overwrite is False.
            ExcelPermissionError: If the directory is not writable.
        """
        path = Path(file_path)

        if path.exists() and not overwrite:
            raise WriteError(
                file_path=file_path,
                operation="create",
                reason="File already exists and overwrite is False",
            )

        parent = path.parent
        if not parent.exists():
            try:
                parent.mkdir(parents=True, exist_ok=True)
            except PermissionError as e:
                raise ExcelPermissionError(
                    file_path=str(parent),
                    operation="create directory",
                ) from e
            except Exception as e:
                raise WriteError(
                    file_path=file_path,
                    operation="create directory",
                    reason=str(e),
                ) from e

        if parent.exists() and not os.access(str(parent), os.W_OK):
            raise ExcelPermissionError(
                file_path=file_path,
                operation="write",
            )

        if not file_path.lower().endswith(".xlsx"):
            path = Path(file_path + ".xlsx")

        return path

    def _calculate_column_widths(
        self,
        rows: list[list[Any]],
        headers: list[str] | None = None,
    ) -> list[int]:
        """
        Calculate optimal column widths based on content.

        Args:
            rows: Data rows.
            headers: Optional header row.

        Returns:
            List of column widths.
        """
        all_rows = rows.copy()
        if headers:
            all_rows.insert(0, headers)

        if not all_rows:
            return []

        max_cols = max(len(row) for row in all_rows)
        widths = [self.DEFAULT_COLUMN_WIDTH] * max_cols

        for row in all_rows:
            for col_idx, cell in enumerate(row):
                if cell is not None:
                    cell_str = str(cell)
                    cell_width = min(len(cell_str) + 2, self.MAX_COLUMN_WIDTH)
                    widths[col_idx] = max(widths[col_idx], cell_width)

        return widths

    def _write_cell(
        self,
        worksheet: xlsxwriter.worksheet.Worksheet,
        row: int,
        col: int,
        value: Any,
        cell_format: xlsxwriter.format.Format | None = None,
    ) -> None:
        """
        Write a value to a cell with appropriate type handling.

        Args:
            worksheet: The worksheet to write to.
            row: Row index (0-based).
            col: Column index (0-based).
            value: Value to write.
            cell_format: Optional format to apply.
        """
        if value is None:
            worksheet.write_blank(row, col, None, cell_format)
        elif isinstance(value, bool):
            worksheet.write_boolean(row, col, value, cell_format)
        elif isinstance(value, (int, float)):
            worksheet.write_number(row, col, value, cell_format)
        elif isinstance(value, datetime):
            worksheet.write_datetime(row, col, value, cell_format)
        elif isinstance(value, str):
            if value.startswith("="):
                worksheet.write_formula(row, col, value, cell_format)
            else:
                worksheet.write_string(row, col, value, cell_format)
        else:
            worksheet.write(row, col, str(value), cell_format)

    def _parse_start_cell(self, start_cell: str) -> tuple[int, int]:
        """
        Parse A1 notation to row and column indices.

        Args:
            start_cell: Cell reference in A1 notation (e.g., "A1", "B5").

        Returns:
            Tuple of (row_index, col_index), both 0-based.
        """
        import re

        match = re.match(r"^([A-Za-z]+)(\d+)$", start_cell.strip())
        if not match:
            return (0, 0)

        col_letters = match.group(1).upper()
        row_num = int(match.group(2))

        col_index = 0
        for char in col_letters:
            col_index = col_index * 26 + (ord(char) - ord("A") + 1)
        col_index -= 1

        row_index = row_num - 1

        return (row_index, col_index)

    def write_sheet(
        self,
        file_path: str,
        rows: list[list[Any]],
        sheet_name: str = "Sheet1",
        headers: list[str] | None = None,
        start_cell: str = "A1",
        overwrite: bool = False,
        auto_format: bool = True,
    ) -> dict[str, Any]:
        """
        Write data to an Excel file.

        Creates a new Excel file with the specified data. If headers are
        provided, they will be written first with bold formatting.

        Args:
            file_path: Path where the file will be written.
            rows: List of rows, where each row is a list of values.
            sheet_name: Name of the sheet. Defaults to "Sheet1".
            headers: Optional list of column headers.
            start_cell: Starting cell for data (A1 notation). Defaults to "A1".
            overwrite: Whether to overwrite an existing file.
            auto_format: Whether to auto-format column widths.

        Returns:
            Dictionary containing:
                - file_path: Path to the written file
                - rows_written: Number of data rows written
                - file_size_bytes: Size of the file in bytes

        Raises:
            WriteError: If writing fails.
            ExcelPermissionError: If the file cannot be written due to permissions.

        Example:
            result = adapter.write_sheet(
                "/path/to/output.xlsx",
                rows=[["Alice", 30], ["Bob", 25]],
                headers=["Name", "Age"],
                sheet_name="Users"
            )
            print(f"Wrote {result['rows_written']} rows")
        """
        path = self._validate_output_path(file_path, overwrite)

        workbook: Workbook | None = None
        try:
            workbook = xlsxwriter.Workbook(str(path))
            worksheet = workbook.add_worksheet(sheet_name)

            header_format = workbook.add_format({
                "bold": True,
                "bg_color": "#4F81BD",
                "font_color": "white",
                "border": 1,
            })
            date_format = workbook.add_format({
                "num_format": "yyyy-mm-dd hh:mm:ss",
            })

            start_row, start_col = self._parse_start_cell(start_cell)

            current_row = start_row
            rows_written = 0

            if headers:
                for col_idx, header in enumerate(headers):
                    self._write_cell(
                        worksheet,
                        current_row,
                        start_col + col_idx,
                        header,
                        header_format,
                    )
                current_row += 1

            for row_data in rows:
                for col_idx, value in enumerate(row_data):
                    cell_format = None
                    if isinstance(value, datetime):
                        cell_format = date_format

                    self._write_cell(
                        worksheet,
                        current_row,
                        start_col + col_idx,
                        value,
                        cell_format,
                    )
                current_row += 1
                rows_written += 1

            if auto_format:
                column_widths = self._calculate_column_widths(rows, headers)
                for col_idx, width in enumerate(column_widths):
                    worksheet.set_column(
                        start_col + col_idx,
                        start_col + col_idx,
                        width,
                    )

            workbook.close()
            workbook = None

            file_size = path.stat().st_size if path.exists() else 0

            return {
                "file_path": str(path.absolute()),
                "rows_written": rows_written,
                "file_size_bytes": file_size,
            }

        except xlsxwriter.exceptions.FileCreateError as e:
            raise ExcelPermissionError(
                file_path=file_path,
                operation="write",
            ) from e
        except Exception as e:
            raise WriteError(
                file_path=file_path,
                operation="write",
                reason=str(e),
            ) from e
        finally:
            if workbook is not None:
                try:
                    workbook.close()
                except Exception:
                    pass

    def write_multiple_sheets(
        self,
        file_path: str,
        sheets_data: dict[str, dict[str, Any]],
        overwrite: bool = False,
        auto_format: bool = True,
    ) -> dict[str, Any]:
        """
        Write multiple sheets to an Excel file.

        Creates a new Excel file with multiple sheets. Each sheet is specified
        by a dictionary entry with sheet name as key and configuration as value.

        Args:
            file_path: Path where the file will be written.
            sheets_data: Dictionary mapping sheet names to sheet configurations.
                Each configuration should contain:
                    - rows: List of data rows (required)
                    - headers: Optional list of headers
                    - start_cell: Starting cell (default: "A1")
            overwrite: Whether to overwrite an existing file.
            auto_format: Whether to auto-format column widths.

        Returns:
            Dictionary containing:
                - file_path: Path to the written file
                - sheets_written: Number of sheets written
                - total_rows_written: Total number of data rows written
                - file_size_bytes: Size of the file in bytes

        Raises:
            WriteError: If writing fails.
            ExcelPermissionError: If the file cannot be written due to permissions.

        Example:
            result = adapter.write_multiple_sheets(
                "/path/to/output.xlsx",
                sheets_data={
                    "Users": {
                        "rows": [["Alice", 30], ["Bob", 25]],
                        "headers": ["Name", "Age"]
                    },
                    "Products": {
                        "rows": [["Widget", 10.99], ["Gadget", 24.99]],
                        "headers": ["Name", "Price"]
                    }
                }
            )
        """
        path = self._validate_output_path(file_path, overwrite)

        workbook: Workbook | None = None
        try:
            workbook = xlsxwriter.Workbook(str(path))

            header_format = workbook.add_format({
                "bold": True,
                "bg_color": "#4F81BD",
                "font_color": "white",
                "border": 1,
            })
            date_format = workbook.add_format({
                "num_format": "yyyy-mm-dd hh:mm:ss",
            })

            total_rows_written = 0
            sheets_written = 0

            for sheet_name, config in sheets_data.items():
                rows = config.get("rows", [])
                headers = config.get("headers")
                start_cell = config.get("start_cell", "A1")

                worksheet = workbook.add_worksheet(sheet_name)

                start_row, start_col = self._parse_start_cell(start_cell)
                current_row = start_row

                if headers:
                    for col_idx, header in enumerate(headers):
                        self._write_cell(
                            worksheet,
                            current_row,
                            start_col + col_idx,
                            header,
                            header_format,
                        )
                    current_row += 1

                for row_data in rows:
                    for col_idx, value in enumerate(row_data):
                        cell_format = None
                        if isinstance(value, datetime):
                            cell_format = date_format

                        self._write_cell(
                            worksheet,
                            current_row,
                            start_col + col_idx,
                            value,
                            cell_format,
                        )
                    current_row += 1
                    total_rows_written += 1

                if auto_format:
                    column_widths = self._calculate_column_widths(rows, headers)
                    for col_idx, width in enumerate(column_widths):
                        worksheet.set_column(
                            start_col + col_idx,
                            start_col + col_idx,
                            width,
                        )

                sheets_written += 1

            workbook.close()
            workbook = None

            file_size = path.stat().st_size if path.exists() else 0

            return {
                "file_path": str(path.absolute()),
                "sheets_written": sheets_written,
                "total_rows_written": total_rows_written,
                "file_size_bytes": file_size,
            }

        except xlsxwriter.exceptions.FileCreateError as e:
            raise ExcelPermissionError(
                file_path=file_path,
                operation="write",
            ) from e
        except Exception as e:
            raise WriteError(
                file_path=file_path,
                operation="write",
                reason=str(e),
            ) from e
        finally:
            if workbook is not None:
                try:
                    workbook.close()
                except Exception:
                    pass
