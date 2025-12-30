"""
Tests for the CalamineAdapter.

Tests the Excel reading functionality using python-calamine.
"""

from pathlib import Path

import pytest

from src.adapters.calamine_adapter import CalamineAdapter
from src.exceptions.excel_exceptions import (
    CellRangeError,
    InvalidFileFormatError,
    SheetNotFoundError,
)
from src.exceptions.excel_exceptions import FileNotFoundError as ExcelFileNotFoundError


class TestCalamineAdapterFileValidation:
    """Tests for file validation in CalamineAdapter."""

    def test_file_not_found_raises_error(
        self,
        calamine_adapter: CalamineAdapter,
    ) -> None:
        """Test that reading a non-existent file raises FileNotFoundError."""
        with pytest.raises(ExcelFileNotFoundError) as exc_info:
            calamine_adapter.get_sheet_names("/nonexistent/file.xlsx")

        assert exc_info.value.error_code == "FILE_NOT_FOUND"
        assert "/nonexistent/file.xlsx" in exc_info.value.message

    def test_invalid_extension_raises_error(
        self,
        calamine_adapter: CalamineAdapter,
        temp_dir: Path,
    ) -> None:
        """Test that reading a file with invalid extension raises error."""
        invalid_file = temp_dir / "test.txt"
        invalid_file.write_text("This is not an Excel file")

        with pytest.raises(InvalidFileFormatError) as exc_info:
            calamine_adapter.get_sheet_names(str(invalid_file))

        assert exc_info.value.error_code == "INVALID_FILE_FORMAT"
        assert ".txt" in exc_info.value.message


class TestCalamineAdapterSheetOperations:
    """Tests for sheet operations in CalamineAdapter."""

    def test_get_sheet_names(
        self,
        calamine_adapter: CalamineAdapter,
        sample_excel_file: Path,
    ) -> None:
        """Test getting sheet names from a workbook."""
        sheet_names = calamine_adapter.get_sheet_names(str(sample_excel_file))

        assert isinstance(sheet_names, list)
        assert len(sheet_names) == 1
        assert "Users" in sheet_names

    def test_get_sheet_names_multiple_sheets(
        self,
        calamine_adapter: CalamineAdapter,
        multi_sheet_excel_file: Path,
    ) -> None:
        """Test getting sheet names from a multi-sheet workbook."""
        sheet_names = calamine_adapter.get_sheet_names(str(multi_sheet_excel_file))

        assert len(sheet_names) == 3
        assert "Users" in sheet_names
        assert "Products" in sheet_names
        assert "Orders" in sheet_names

    def test_sheet_not_found_raises_error(
        self,
        calamine_adapter: CalamineAdapter,
        sample_excel_file: Path,
    ) -> None:
        """Test that reading a non-existent sheet raises SheetNotFoundError."""
        with pytest.raises(SheetNotFoundError) as exc_info:
            calamine_adapter.read_sheet(
                str(sample_excel_file),
                sheet_name="NonExistent",
            )

        assert exc_info.value.error_code == "SHEET_NOT_FOUND"
        assert "NonExistent" in exc_info.value.message
        assert "Users" in exc_info.value.available_sheets


class TestCalamineAdapterReadOperations:
    """Tests for read operations in CalamineAdapter."""

    def test_read_sheet_by_name(
        self,
        calamine_adapter: CalamineAdapter,
        sample_excel_file: Path,
    ) -> None:
        """Test reading a sheet by name."""
        data = calamine_adapter.read_sheet(
            str(sample_excel_file),
            sheet_name="Users",
        )

        assert data.sheet_name == "Users"
        assert data.row_count == 4
        assert data.column_count == 3
        assert data.rows[0] == ["Name", "Age", "Email"]
        assert data.rows[1][0] == "Alice"

    def test_read_sheet_by_index(
        self,
        calamine_adapter: CalamineAdapter,
        sample_excel_file: Path,
    ) -> None:
        """Test reading a sheet by index."""
        data = calamine_adapter.read_sheet(
            str(sample_excel_file),
            sheet_index=0,
        )

        assert data.sheet_name == "Users"
        assert data.row_count == 4

    def test_read_first_sheet_by_default(
        self,
        calamine_adapter: CalamineAdapter,
        sample_excel_file: Path,
    ) -> None:
        """Test that first sheet is read when no sheet is specified."""
        data = calamine_adapter.read_sheet(str(sample_excel_file))

        assert data.sheet_name == "Users"

    def test_read_sheet_skip_empty_rows(
        self,
        calamine_adapter: CalamineAdapter,
        sample_excel_file: Path,
    ) -> None:
        """Test skipping empty rows when reading."""
        data = calamine_adapter.read_sheet(
            str(sample_excel_file),
            skip_empty_rows=True,
        )

        assert data.row_count == 4

    def test_get_workbook_info(
        self,
        calamine_adapter: CalamineAdapter,
        multi_sheet_excel_file: Path,
    ) -> None:
        """Test getting workbook information."""
        info = calamine_adapter.get_workbook_info(str(multi_sheet_excel_file))

        assert info.sheet_count == 3
        assert len(info.sheets) == 3
        assert info.file_size_bytes > 0
        assert info.file_path == str(multi_sheet_excel_file.absolute())


class TestCalamineAdapterRangeOperations:
    """Tests for range operations in CalamineAdapter."""

    def test_read_range(
        self,
        calamine_adapter: CalamineAdapter,
        sample_excel_file: Path,
    ) -> None:
        """Test reading a cell range."""
        data = calamine_adapter.read_range(
            str(sample_excel_file),
            cell_range="A1:B2",
            sheet_name="Users",
        )

        assert data.row_count == 2
        assert data.column_count == 2
        assert data.rows[0] == ["Name", "Age"]
        assert data.rows[1][0] == "Alice"

    def test_read_single_cell_range(
        self,
        calamine_adapter: CalamineAdapter,
        sample_excel_file: Path,
    ) -> None:
        """Test reading a single cell as a range."""
        data = calamine_adapter.read_range(
            str(sample_excel_file),
            cell_range="A1",
            sheet_name="Users",
        )

        assert data.row_count == 1
        assert data.column_count == 1
        assert data.rows[0] == ["Name"]

    def test_get_cell_value(
        self,
        calamine_adapter: CalamineAdapter,
        sample_excel_file: Path,
    ) -> None:
        """Test getting a single cell value."""
        value = calamine_adapter.get_cell_value(
            str(sample_excel_file),
            cell="A2",
            sheet_name="Users",
        )

        assert value == "Alice"

    def test_get_numeric_cell_value(
        self,
        calamine_adapter: CalamineAdapter,
        sample_excel_file: Path,
    ) -> None:
        """Test getting a numeric cell value."""
        value = calamine_adapter.get_cell_value(
            str(sample_excel_file),
            cell="B2",
            sheet_name="Users",
        )

        assert value == 30
        assert isinstance(value, int)

    def test_invalid_range_format(
        self,
        calamine_adapter: CalamineAdapter,
        sample_excel_file: Path,
    ) -> None:
        """Test that invalid range format raises CellRangeError."""
        with pytest.raises(CellRangeError) as exc_info:
            calamine_adapter.read_range(
                str(sample_excel_file),
                cell_range="invalid",
            )

        assert exc_info.value.error_code == "INVALID_CELL_RANGE"

    def test_range_out_of_bounds_returns_empty(
        self,
        calamine_adapter: CalamineAdapter,
        sample_excel_file: Path,
    ) -> None:
        """Test that reading beyond data returns empty/None values."""
        data = calamine_adapter.read_range(
            str(sample_excel_file),
            cell_range="Z100:AA101",
            sheet_name="Users",
        )

        assert data.row_count == 0


class TestCalamineAdapterColumnConversion:
    """Tests for column letter to index conversion."""

    def test_single_letter_columns(
        self,
        calamine_adapter: CalamineAdapter,
    ) -> None:
        """Test converting single letter columns."""
        assert calamine_adapter._column_letter_to_index("A") == 0
        assert calamine_adapter._column_letter_to_index("B") == 1
        assert calamine_adapter._column_letter_to_index("Z") == 25

    def test_double_letter_columns(
        self,
        calamine_adapter: CalamineAdapter,
    ) -> None:
        """Test converting double letter columns."""
        assert calamine_adapter._column_letter_to_index("AA") == 26
        assert calamine_adapter._column_letter_to_index("AB") == 27
        assert calamine_adapter._column_letter_to_index("AZ") == 51
        assert calamine_adapter._column_letter_to_index("BA") == 52

    def test_lowercase_columns(
        self,
        calamine_adapter: CalamineAdapter,
    ) -> None:
        """Test that lowercase columns are converted correctly."""
        assert calamine_adapter._column_letter_to_index("a") == 0
        assert calamine_adapter._column_letter_to_index("z") == 25
