"""
Calamine adapter for high-performance Excel reading.

This module provides the CalamineAdapter class that wraps python-calamine
for reading Excel files. python-calamine is a Rust-based library that provides
exceptional performance for large files (100MB+) with minimal memory footprint.

Supported formats:
    - .xlsx (Excel 2007+)
    - .xls (Excel 97-2003)
    - .xlsb (Excel Binary)
    - .xlsm (Excel Macro-Enabled)
    - .ods (OpenDocument Spreadsheet)

Example:
    adapter = CalamineAdapter()
    sheet_names = adapter.get_sheet_names("/path/to/file.xlsx")
    data = adapter.read_sheet("/path/to/file.xlsx", "Sheet1")
"""

import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from python_calamine import CalamineWorkbook

from src.exceptions.excel_exceptions import (
    CellRangeError,
    FileNotFoundError,
    InvalidFileFormatError,
    ReadError,
    SheetNotFoundError,
)
from src.models.excel_models import CellRange, SheetData, SheetInfo, WorkbookInfo


class CalamineAdapter:
    """
    Adapter for python-calamine Excel reading operations.

    This adapter provides a clean interface for reading Excel files using
    the high-performance python-calamine library. It handles format conversion,
    error mapping, and data normalization.

    Attributes:
        SUPPORTED_EXTENSIONS: Tuple of supported file extensions.
        EXCEL_EPOCH: Excel's date epoch (1899-12-30).

    Example:
        adapter = CalamineAdapter()

        # Get workbook info
        info = adapter.get_workbook_info("/path/to/file.xlsx")

        # Read entire sheet
        data = adapter.read_sheet("/path/to/file.xlsx", sheet_name="Data")

        # Read specific range
        data = adapter.read_range("/path/to/file.xlsx", "A1:C10", sheet_name="Data")
    """

    SUPPORTED_EXTENSIONS = (".xlsx", ".xls", ".xlsb", ".xlsm", ".ods")
    # Excel's date system epoch. Excel incorrectly treats 1900 as a leap year
    # for compatibility with Lotus 1-2-3, so the epoch is December 30, 1899.
    # All Excel serial dates are counted from this date.
    EXCEL_EPOCH = datetime(1899, 12, 30)

    def __init__(self) -> None:
        """Initialize the CalamineAdapter."""
        pass

    def _validate_file_path(self, file_path: str) -> Path:
        """
        Validate that the file exists and has a supported extension.

        Args:
            file_path: Path to the Excel file.

        Returns:
            Path object for the validated file.

        Raises:
            FileNotFoundError: If the file does not exist.
            InvalidFileFormatError: If the file extension is not supported.
        """
        path = Path(file_path)

        if not path.exists():
            raise FileNotFoundError(file_path)

        if path.suffix.lower() not in self.SUPPORTED_EXTENSIONS:
            raise InvalidFileFormatError(
                file_path=file_path,
                expected_formats=list(self.SUPPORTED_EXTENSIONS),
                reason=f"Unsupported file extension: {path.suffix}",
            )

        return path

    def _open_workbook(self, file_path: str) -> CalamineWorkbook:
        """
        Open an Excel workbook using calamine.

        Args:
            file_path: Path to the Excel file.

        Returns:
            CalamineWorkbook instance.

        Raises:
            InvalidFileFormatError: If the file cannot be parsed.
            ReadError: If an unexpected error occurs during opening.
        """
        path = self._validate_file_path(file_path)

        try:
            return CalamineWorkbook.from_path(str(path))
        except Exception as e:
            error_msg = str(e).lower()
            if "invalid" in error_msg or "corrupt" in error_msg or "format" in error_msg:
                raise InvalidFileFormatError(
                    file_path=file_path,
                    reason=str(e),
                ) from e
            raise ReadError(
                file_path=file_path,
                operation="open",
                reason=str(e),
            ) from e

    def _convert_excel_date(self, serial_date: float) -> datetime:
        """
        Convert Excel serial date number to Python datetime.

        Excel stores dates as floating-point numbers representing
        the number of days since December 30, 1899.

        Args:
            serial_date: Excel serial date number.

        Returns:
            Python datetime object.
        """
        return self.EXCEL_EPOCH + timedelta(days=serial_date)

    def _normalize_cell_value(self, value: Any) -> Any:
        """
        Normalize a cell value from calamine to Python types.

        Handles conversion of Excel-specific types (dates, errors, etc.)
        to standard Python types.

        Args:
            value: Raw cell value from calamine.

        Returns:
            Normalized Python value.
        """
        if value is None:
            return None

        if isinstance(value, float):
            if value == int(value):
                return int(value)
            return value

        if isinstance(value, (str, int, bool)):
            return value

        if isinstance(value, datetime):
            return value

        return str(value)

    def _parse_a1_notation(self, a1_range: str) -> CellRange:
        """
        Parse A1 notation range string into CellRange.

        Supports formats like:
        - "A1" (single cell)
        - "A1:C10" (range)
        - "A:C" (full columns)
        - "1:10" (full rows)

        Args:
            a1_range: A1 notation string.

        Returns:
            CellRange object.

        Raises:
            CellRangeError: If the range notation is invalid.
        """
        a1_range = a1_range.strip().upper()

        single_cell_pattern = r"^([A-Z]+)(\d+)$"
        range_pattern = r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$"

        single_match = re.match(single_cell_pattern, a1_range)
        if single_match:
            col = self._column_letter_to_index(single_match.group(1))
            row = int(single_match.group(2)) - 1
            return CellRange(
                start_row=row,
                end_row=row,
                start_col=col,
                end_col=col,
                a1_notation=a1_range,
            )

        range_match = re.match(range_pattern, a1_range)
        if range_match:
            start_col = self._column_letter_to_index(range_match.group(1))
            start_row = int(range_match.group(2)) - 1
            end_col = self._column_letter_to_index(range_match.group(3))
            end_row = int(range_match.group(4)) - 1

            if start_row > end_row or start_col > end_col:
                raise CellRangeError(
                    cell_range=a1_range,
                    reason="Start position must be before end position",
                )

            return CellRange(
                start_row=start_row,
                end_row=end_row,
                start_col=start_col,
                end_col=end_col,
                a1_notation=a1_range,
            )

        raise CellRangeError(
            cell_range=a1_range,
            reason="Invalid A1 notation format. Expected format: 'A1' or 'A1:C10'",
        )

    def _column_letter_to_index(self, column_letter: str) -> int:
        """
        Convert Excel column letter(s) to 0-based index.

        Args:
            column_letter: Column letter(s) like "A", "B", "AA", "AB".

        Returns:
            0-based column index.
        """
        result = 0
        for char in column_letter.upper():
            result = result * 26 + (ord(char) - ord("A") + 1)
        return result - 1

    def get_sheet_names(self, file_path: str) -> list[str]:
        """
        Get the list of sheet names in the workbook.

        Args:
            file_path: Path to the Excel file.

        Returns:
            List of sheet names.

        Raises:
            FileNotFoundError: If the file does not exist.
            InvalidFileFormatError: If the file format is not supported.
        """
        workbook = self._open_workbook(file_path)
        return workbook.sheet_names

    def get_workbook_info(self, file_path: str) -> WorkbookInfo:
        """
        Get metadata about an Excel workbook.

        Args:
            file_path: Path to the Excel file.

        Returns:
            WorkbookInfo containing workbook metadata.

        Raises:
            FileNotFoundError: If the file does not exist.
            InvalidFileFormatError: If the file format is not supported.
        """
        path = self._validate_file_path(file_path)
        workbook = self._open_workbook(file_path)

        sheet_names = workbook.sheet_names
        sheets: list[SheetInfo] = []

        for index, name in enumerate(sheet_names):
            try:
                data = workbook.get_sheet_by_name(name).to_python()
                row_count = len(data)
                column_count = max(len(row) for row in data) if data else 0
            except Exception:
                row_count = None
                column_count = None

            sheets.append(
                SheetInfo(
                    name=name,
                    index=index,
                    visible=True,
                    row_count=row_count,
                    column_count=column_count,
                )
            )

        stat = path.stat()

        return WorkbookInfo(
            file_path=str(path.absolute()),
            file_size_bytes=stat.st_size,
            sheet_count=len(sheet_names),
            sheets=sheets,
            modified_at=datetime.fromtimestamp(stat.st_mtime),
        )

    def read_sheet(
        self,
        file_path: str,
        sheet_name: str | None = None,
        sheet_index: int | None = None,
        skip_empty_rows: bool = False,
    ) -> SheetData:
        """
        Read data from a specific sheet.

        Args:
            file_path: Path to the Excel file.
            sheet_name: Name of the sheet to read. If None and sheet_index is None,
                       reads the first sheet.
            sheet_index: Index of the sheet to read (0-based). Used if sheet_name is None.
            skip_empty_rows: Whether to skip empty rows.

        Returns:
            SheetData containing the sheet contents.

        Raises:
            FileNotFoundError: If the file does not exist.
            InvalidFileFormatError: If the file format is not supported.
            SheetNotFoundError: If the specified sheet does not exist.
        """
        workbook = self._open_workbook(file_path)
        available_sheets = workbook.sheet_names

        if sheet_name is not None:
            if sheet_name not in available_sheets:
                raise SheetNotFoundError(
                    sheet_name=sheet_name,
                    available_sheets=available_sheets,
                )
            target_sheet_name = sheet_name
        elif sheet_index is not None:
            if sheet_index < 0 or sheet_index >= len(available_sheets):
                raise SheetNotFoundError(
                    sheet_name=f"index {sheet_index}",
                    available_sheets=available_sheets,
                )
            target_sheet_name = available_sheets[sheet_index]
        else:
            if not available_sheets:
                raise SheetNotFoundError(
                    sheet_name="(first sheet)",
                    available_sheets=[],
                )
            target_sheet_name = available_sheets[0]

        try:
            sheet = workbook.get_sheet_by_name(target_sheet_name)
            raw_data = sheet.to_python()
        except Exception as e:
            raise ReadError(
                file_path=file_path,
                operation="read sheet",
                reason=str(e),
            ) from e

        rows: list[list[Any]] = []
        for row in raw_data:
            normalized_row = [self._normalize_cell_value(cell) for cell in row]

            if skip_empty_rows:
                if all(cell is None or cell == "" for cell in normalized_row):
                    continue

            rows.append(normalized_row)

        column_count = max(len(row) for row in rows) if rows else 0

        return SheetData(
            sheet_name=target_sheet_name,
            rows=rows,
            row_count=len(rows),
            column_count=column_count,
        )

    def read_range(
        self,
        file_path: str,
        cell_range: str,
        sheet_name: str | None = None,
        sheet_index: int | None = None,
    ) -> SheetData:
        """
        Read a specific cell range from a sheet.

        Args:
            file_path: Path to the Excel file.
            cell_range: Cell range in A1 notation (e.g., "A1:C10").
            sheet_name: Name of the sheet to read. If None, reads the first sheet.
            sheet_index: Index of the sheet to read (0-based). Used if sheet_name is None.

        Returns:
            SheetData containing the range contents.

        Raises:
            FileNotFoundError: If the file does not exist.
            InvalidFileFormatError: If the file format is not supported.
            SheetNotFoundError: If the specified sheet does not exist.
            CellRangeError: If the cell range is invalid.
        """
        parsed_range = self._parse_a1_notation(cell_range)

        full_sheet = self.read_sheet(
            file_path=file_path,
            sheet_name=sheet_name,
            sheet_index=sheet_index,
        )

        start_row = parsed_range.start_row
        end_row = min(parsed_range.end_row, full_sheet.row_count - 1)
        start_col = parsed_range.start_col
        end_col = parsed_range.end_col

        if start_row > full_sheet.row_count - 1:
            return SheetData(
                sheet_name=full_sheet.sheet_name,
                rows=[],
                row_count=0,
                column_count=0,
                cell_range=parsed_range,
            )

        extracted_rows: list[list[Any]] = []
        for row_idx in range(start_row, end_row + 1):
            if row_idx < len(full_sheet.rows):
                row = full_sheet.rows[row_idx]
                row_slice = row[start_col : end_col + 1]

                while len(row_slice) < (end_col - start_col + 1):
                    row_slice.append(None)

                extracted_rows.append(row_slice)
            else:
                extracted_rows.append([None] * (end_col - start_col + 1))

        return SheetData(
            sheet_name=full_sheet.sheet_name,
            rows=extracted_rows,
            row_count=len(extracted_rows),
            column_count=end_col - start_col + 1,
            cell_range=parsed_range,
        )

    def get_cell_value(
        self,
        file_path: str,
        cell: str,
        sheet_name: str | None = None,
        sheet_index: int | None = None,
    ) -> Any:
        """
        Get the value of a single cell.

        Args:
            file_path: Path to the Excel file.
            cell: Cell reference in A1 notation (e.g., "A1").
            sheet_name: Name of the sheet. If None, uses the first sheet.
            sheet_index: Index of the sheet (0-based). Used if sheet_name is None.

        Returns:
            The cell value (can be None, str, int, float, bool, or datetime).

        Raises:
            FileNotFoundError: If the file does not exist.
            InvalidFileFormatError: If the file format is not supported.
            SheetNotFoundError: If the specified sheet does not exist.
            CellRangeError: If the cell reference is invalid.
        """
        range_data = self.read_range(
            file_path=file_path,
            cell_range=cell,
            sheet_name=sheet_name,
            sheet_index=sheet_index,
        )

        if range_data.rows and range_data.rows[0]:
            return range_data.rows[0][0]

        return None
