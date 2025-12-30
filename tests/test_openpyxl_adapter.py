"""
Tests for the OpenpyxlAdapter.

Tests the Excel reading and writing functionality using openpyxl.
"""

import os
from datetime import datetime
from pathlib import Path

import pytest

from src.adapters.openpyxl_adapter import OpenpyxlAdapter
from src.exceptions.excel_exceptions import (
    CellRangeError,
    InvalidFileFormatError,
    SheetNotFoundError,
    WriteError,
)
from src.exceptions.excel_exceptions import FileNotFoundError as ExcelFileNotFoundError


@pytest.fixture
def openpyxl_adapter() -> OpenpyxlAdapter:
    """Create an OpenpyxlAdapter instance for testing."""
    return OpenpyxlAdapter()


@pytest.fixture
def sample_openpyxl_file(temp_dir: Path, openpyxl_adapter: OpenpyxlAdapter) -> Path:
    """Create a sample Excel file using openpyxl for testing."""
    file_path = temp_dir / "openpyxl_sample.xlsx"

    openpyxl_adapter.write_sheet(
        file_path=str(file_path),
        rows=[
            ["Alice", 30, "alice@example.com"],
            ["Bob", 25, "bob@example.com"],
            ["Charlie", 35, "charlie@example.com"],
        ],
        headers=["Name", "Age", "Email"],
        sheet_name="Users",
    )

    return file_path


@pytest.fixture
def multi_sheet_openpyxl_file(temp_dir: Path, openpyxl_adapter: OpenpyxlAdapter) -> Path:
    """Create an Excel file with multiple sheets using openpyxl for testing."""
    file_path = temp_dir / "openpyxl_multi_sheet.xlsx"

    openpyxl_adapter.write_multiple_sheets(
        file_path=str(file_path),
        sheets_data={
            "Users": {
                "rows": [
                    ["Alice", 30],
                    ["Bob", 25],
                ],
                "headers": ["Name", "Age"],
            },
            "Products": {
                "rows": [
                    ["Widget", 10.99],
                    ["Gadget", 24.99],
                ],
                "headers": ["Name", "Price"],
            },
            "Orders": {
                "rows": [
                    [1, "Alice", "Widget", 2],
                    [2, "Bob", "Gadget", 1],
                ],
                "headers": ["OrderID", "Customer", "Product", "Quantity"],
            },
        },
    )

    return file_path


class TestOpenpyxlAdapterFileValidation:
    """Tests for file validation in OpenpyxlAdapter."""

    def test_file_not_found_raises_error(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
    ) -> None:
        """Test that reading a non-existent file raises FileNotFoundError."""
        with pytest.raises(ExcelFileNotFoundError) as exc_info:
            openpyxl_adapter.get_sheet_names("/nonexistent/file.xlsx")

        assert exc_info.value.error_code == "FILE_NOT_FOUND"
        assert "/nonexistent/file.xlsx" in exc_info.value.message

    def test_invalid_extension_raises_error(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        temp_dir: Path,
    ) -> None:
        """Test that reading a file with invalid extension raises error."""
        invalid_file = temp_dir / "test.txt"
        invalid_file.write_text("This is not an Excel file")

        with pytest.raises(InvalidFileFormatError) as exc_info:
            openpyxl_adapter.get_sheet_names(str(invalid_file))

        assert exc_info.value.error_code == "INVALID_FILE_FORMAT"
        assert ".txt" in exc_info.value.message


class TestOpenpyxlAdapterSheetOperations:
    """Tests for sheet operations in OpenpyxlAdapter."""

    def test_get_sheet_names(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        sample_openpyxl_file: Path,
    ) -> None:
        """Test getting sheet names from a workbook."""
        sheet_names = openpyxl_adapter.get_sheet_names(str(sample_openpyxl_file))

        assert isinstance(sheet_names, list)
        assert len(sheet_names) == 1
        assert "Users" in sheet_names

    def test_get_sheet_names_multiple_sheets(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        multi_sheet_openpyxl_file: Path,
    ) -> None:
        """Test getting sheet names from a multi-sheet workbook."""
        sheet_names = openpyxl_adapter.get_sheet_names(str(multi_sheet_openpyxl_file))

        assert len(sheet_names) == 3
        assert "Users" in sheet_names
        assert "Products" in sheet_names
        assert "Orders" in sheet_names

    def test_sheet_not_found_raises_error(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        sample_openpyxl_file: Path,
    ) -> None:
        """Test that reading a non-existent sheet raises SheetNotFoundError."""
        with pytest.raises(SheetNotFoundError) as exc_info:
            openpyxl_adapter.read_sheet(
                str(sample_openpyxl_file),
                sheet_name="NonExistent",
            )

        assert exc_info.value.error_code == "SHEET_NOT_FOUND"
        assert "NonExistent" in exc_info.value.message


class TestOpenpyxlAdapterReadOperations:
    """Tests for read operations in OpenpyxlAdapter."""

    def test_read_sheet_by_name(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        sample_openpyxl_file: Path,
    ) -> None:
        """Test reading a sheet by name."""
        data = openpyxl_adapter.read_sheet(
            str(sample_openpyxl_file),
            sheet_name="Users",
        )

        assert data.sheet_name == "Users"
        assert data.row_count == 4
        assert data.column_count == 3
        assert data.rows[0] == ["Name", "Age", "Email"]
        assert data.rows[1][0] == "Alice"

    def test_read_sheet_by_index(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        sample_openpyxl_file: Path,
    ) -> None:
        """Test reading a sheet by index."""
        data = openpyxl_adapter.read_sheet(
            str(sample_openpyxl_file),
            sheet_index=0,
        )

        assert data.sheet_name == "Users"
        assert data.row_count == 4

    def test_read_first_sheet_by_default(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        sample_openpyxl_file: Path,
    ) -> None:
        """Test that first sheet is read when no sheet is specified."""
        data = openpyxl_adapter.read_sheet(str(sample_openpyxl_file))

        assert data.sheet_name == "Users"

    def test_read_sheet_skip_empty_rows(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        sample_openpyxl_file: Path,
    ) -> None:
        """Test skipping empty rows when reading."""
        data = openpyxl_adapter.read_sheet(
            str(sample_openpyxl_file),
            skip_empty_rows=True,
        )

        assert data.row_count == 4

    def test_get_workbook_info(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        multi_sheet_openpyxl_file: Path,
    ) -> None:
        """Test getting workbook information."""
        info = openpyxl_adapter.get_workbook_info(str(multi_sheet_openpyxl_file))

        assert info.sheet_count == 3
        assert len(info.sheets) == 3
        assert info.file_size_bytes > 0
        assert info.file_path == str(multi_sheet_openpyxl_file.absolute())


class TestOpenpyxlAdapterRangeOperations:
    """Tests for range operations in OpenpyxlAdapter."""

    def test_read_range(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        sample_openpyxl_file: Path,
    ) -> None:
        """Test reading a cell range."""
        data = openpyxl_adapter.read_range(
            str(sample_openpyxl_file),
            cell_range="A1:B2",
            sheet_name="Users",
        )

        assert data.row_count == 2
        assert data.column_count == 2
        assert data.rows[0] == ["Name", "Age"]
        assert data.rows[1][0] == "Alice"

    def test_read_single_cell_range(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        sample_openpyxl_file: Path,
    ) -> None:
        """Test reading a single cell as a range."""
        data = openpyxl_adapter.read_range(
            str(sample_openpyxl_file),
            cell_range="A1",
            sheet_name="Users",
        )

        assert data.row_count == 1
        assert data.column_count == 1
        assert data.rows[0] == ["Name"]

    def test_get_cell_value(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        sample_openpyxl_file: Path,
    ) -> None:
        """Test getting a single cell value."""
        value = openpyxl_adapter.get_cell_value(
            str(sample_openpyxl_file),
            cell="A2",
            sheet_name="Users",
        )

        assert value == "Alice"

    def test_get_numeric_cell_value(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        sample_openpyxl_file: Path,
    ) -> None:
        """Test getting a numeric cell value."""
        value = openpyxl_adapter.get_cell_value(
            str(sample_openpyxl_file),
            cell="B2",
            sheet_name="Users",
        )

        assert value == 30
        assert isinstance(value, int)

    def test_invalid_range_format(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        sample_openpyxl_file: Path,
    ) -> None:
        """Test that invalid range format raises CellRangeError."""
        with pytest.raises(CellRangeError) as exc_info:
            openpyxl_adapter.read_range(
                str(sample_openpyxl_file),
                cell_range="invalid",
            )

        assert exc_info.value.error_code == "INVALID_CELL_RANGE"


class TestOpenpyxlAdapterBasicWrite:
    """Tests for basic write operations in OpenpyxlAdapter."""

    def test_write_simple_data(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        temp_dir: Path,
    ) -> None:
        """Test writing simple data to an Excel file."""
        file_path = temp_dir / "openpyxl_output.xlsx"
        sample_data = [
            ["Alice", 30, "Engineering"],
            ["Bob", 25, "Marketing"],
        ]

        result = openpyxl_adapter.write_sheet(
            file_path=str(file_path),
            rows=sample_data,
        )

        assert result["rows_written"] == 2
        assert os.path.exists(result["file_path"])
        assert result["file_size_bytes"] > 0

    def test_write_with_headers(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        temp_dir: Path,
    ) -> None:
        """Test writing data with headers."""
        file_path = temp_dir / "openpyxl_headers.xlsx"
        sample_data = [["Alice", 30], ["Bob", 25]]
        sample_headers = ["Name", "Age"]

        result = openpyxl_adapter.write_sheet(
            file_path=str(file_path),
            rows=sample_data,
            headers=sample_headers,
        )

        assert result["rows_written"] == 2
        assert os.path.exists(result["file_path"])

    def test_write_custom_sheet_name(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        temp_dir: Path,
    ) -> None:
        """Test writing data with a custom sheet name."""
        file_path = temp_dir / "openpyxl_custom_sheet.xlsx"
        sample_data = [["Data"]]

        result = openpyxl_adapter.write_sheet(
            file_path=str(file_path),
            rows=sample_data,
            sheet_name="CustomData",
        )

        assert result["rows_written"] == 1

    def test_write_with_start_cell(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        temp_dir: Path,
    ) -> None:
        """Test writing data starting from a specific cell."""
        file_path = temp_dir / "openpyxl_offset_start.xlsx"
        sample_data = [["Data1"], ["Data2"]]

        result = openpyxl_adapter.write_sheet(
            file_path=str(file_path),
            rows=sample_data,
            start_cell="B5",
        )

        assert result["rows_written"] == 2


class TestOpenpyxlAdapterDataTypes:
    """Tests for writing different data types with openpyxl."""

    def test_write_mixed_types(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        temp_dir: Path,
    ) -> None:
        """Test writing data with mixed types."""
        file_path = temp_dir / "openpyxl_mixed_types.xlsx"

        rows = [
            ["String", 42, 3.14, True, None],
            ["Another", 100, 2.718, False, "Value"],
        ]

        result = openpyxl_adapter.write_sheet(
            file_path=str(file_path),
            rows=rows,
            headers=["Text", "Integer", "Float", "Boolean", "Nullable"],
        )

        assert result["rows_written"] == 2

    def test_write_datetime(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        temp_dir: Path,
    ) -> None:
        """Test writing datetime values."""
        file_path = temp_dir / "openpyxl_datetime.xlsx"

        now = datetime.now()
        rows = [
            ["Event 1", now],
            ["Event 2", datetime(2024, 1, 15, 10, 30)],
        ]

        result = openpyxl_adapter.write_sheet(
            file_path=str(file_path),
            rows=rows,
            headers=["Event", "Timestamp"],
        )

        assert result["rows_written"] == 2


class TestOpenpyxlAdapterOverwrite:
    """Tests for overwrite behavior in openpyxl."""

    def test_write_fails_without_overwrite(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        temp_dir: Path,
    ) -> None:
        """Test that writing to existing file fails without overwrite flag."""
        file_path = temp_dir / "openpyxl_existing.xlsx"
        sample_data = [["Data"]]

        openpyxl_adapter.write_sheet(
            file_path=str(file_path),
            rows=sample_data,
        )

        with pytest.raises(WriteError) as exc_info:
            openpyxl_adapter.write_sheet(
                file_path=str(file_path),
                rows=sample_data,
                overwrite=False,
            )

        assert exc_info.value.error_code == "WRITE_ERROR"
        assert "already exists" in exc_info.value.message

    def test_write_succeeds_with_overwrite(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        temp_dir: Path,
    ) -> None:
        """Test that writing to existing file succeeds with overwrite flag."""
        file_path = temp_dir / "openpyxl_overwrite.xlsx"
        sample_data = [["Original"]]

        openpyxl_adapter.write_sheet(
            file_path=str(file_path),
            rows=sample_data,
        )

        new_data = [["New", "Data"]]
        result = openpyxl_adapter.write_sheet(
            file_path=str(file_path),
            rows=new_data,
            overwrite=True,
        )

        assert result["rows_written"] == 1


class TestOpenpyxlAdapterMultipleSheets:
    """Tests for writing multiple sheets with openpyxl."""

    def test_write_multiple_sheets(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        temp_dir: Path,
    ) -> None:
        """Test writing multiple sheets to a single file."""
        file_path = temp_dir / "openpyxl_multi_sheet.xlsx"

        sheets_data = {
            "Users": {
                "rows": [["Alice", 30], ["Bob", 25]],
                "headers": ["Name", "Age"],
            },
            "Products": {
                "rows": [["Widget", 10.99], ["Gadget", 24.99]],
                "headers": ["Name", "Price"],
            },
        }

        result = openpyxl_adapter.write_multiple_sheets(
            file_path=str(file_path),
            sheets_data=sheets_data,
        )

        assert result["sheets_written"] == 2
        assert result["total_rows_written"] == 4

    def test_write_empty_sheet(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        temp_dir: Path,
    ) -> None:
        """Test writing an empty sheet."""
        file_path = temp_dir / "openpyxl_empty_sheet.xlsx"

        sheets_data = {
            "EmptySheet": {
                "rows": [],
            },
            "DataSheet": {
                "rows": [["Data"]],
            },
        }

        result = openpyxl_adapter.write_multiple_sheets(
            file_path=str(file_path),
            sheets_data=sheets_data,
        )

        assert result["sheets_written"] == 2
        assert result["total_rows_written"] == 1


class TestOpenpyxlAdapterModifyExisting:
    """Tests for modifying existing workbooks with openpyxl."""

    def test_modify_existing_sheet(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        sample_openpyxl_file: Path,
    ) -> None:
        """Test modifying an existing sheet."""
        result = openpyxl_adapter.modify_existing_workbook(
            file_path=str(sample_openpyxl_file),
            sheet_name="Users",
            rows=[["NewData1"], ["NewData2"]],
            start_cell="E1",
        )

        assert result["rows_written"] == 2
        assert result["sheet_created"] is False

    def test_create_new_sheet_in_existing_workbook(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        sample_openpyxl_file: Path,
    ) -> None:
        """Test creating a new sheet in an existing workbook."""
        result = openpyxl_adapter.modify_existing_workbook(
            file_path=str(sample_openpyxl_file),
            sheet_name="NewSheet",
            rows=[["Data1"], ["Data2"]],
            create_sheet_if_missing=True,
        )

        assert result["rows_written"] == 2
        assert result["sheet_created"] is True

        # Verify the new sheet exists
        sheets = openpyxl_adapter.get_sheet_names(str(sample_openpyxl_file))
        assert "NewSheet" in sheets

    def test_modify_fails_if_sheet_missing_and_create_disabled(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        sample_openpyxl_file: Path,
    ) -> None:
        """Test that modifying fails if sheet doesn't exist and create is disabled."""
        with pytest.raises(SheetNotFoundError):
            openpyxl_adapter.modify_existing_workbook(
                file_path=str(sample_openpyxl_file),
                sheet_name="NonExistentSheet",
                rows=[["Data"]],
                create_sheet_if_missing=False,
            )


class TestOpenpyxlAdapterRoundTrip:
    """Tests for read/write round-trip operations."""

    def test_write_then_read(
        self,
        openpyxl_adapter: OpenpyxlAdapter,
        temp_dir: Path,
    ) -> None:
        """Test writing and then reading back the same data."""
        file_path = temp_dir / "openpyxl_roundtrip.xlsx"
        sample_data = [
            ["Alice", 30],
            ["Bob", 25],
        ]
        sample_headers = ["Name", "Age"]

        # Write
        write_result = openpyxl_adapter.write_sheet(
            file_path=str(file_path),
            rows=sample_data,
            headers=sample_headers,
        )
        assert write_result["rows_written"] == 2

        # Read back
        data = openpyxl_adapter.read_sheet(str(file_path))
        assert data.row_count == 3  # 1 header + 2 data rows
        assert data.rows[0] == sample_headers
        assert data.rows[1] == sample_data[0]
        assert data.rows[2] == sample_data[1]
