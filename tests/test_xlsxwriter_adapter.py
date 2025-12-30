"""
Tests for the XlsxWriterAdapter.

Tests the Excel writing functionality using XlsxWriter.
"""

import os
from datetime import datetime
from pathlib import Path

import pytest

from src.adapters.xlsxwriter_adapter import XlsxWriterAdapter
from src.exceptions.excel_exceptions import WriteError


class TestXlsxWriterAdapterBasicWrite:
    """Tests for basic write operations in XlsxWriterAdapter."""

    def test_write_simple_data(
        self,
        xlsxwriter_adapter: XlsxWriterAdapter,
        temp_dir: Path,
        sample_data: list[list],
    ) -> None:
        """Test writing simple data to an Excel file."""
        file_path = temp_dir / "output.xlsx"

        result = xlsxwriter_adapter.write_sheet(
            file_path=str(file_path),
            rows=sample_data,
        )

        assert result["rows_written"] == 5
        assert os.path.exists(result["file_path"])
        assert result["file_size_bytes"] > 0

    def test_write_with_headers(
        self,
        xlsxwriter_adapter: XlsxWriterAdapter,
        temp_dir: Path,
        sample_data: list[list],
        sample_headers: list[str],
    ) -> None:
        """Test writing data with headers."""
        file_path = temp_dir / "output_headers.xlsx"

        result = xlsxwriter_adapter.write_sheet(
            file_path=str(file_path),
            rows=sample_data,
            headers=sample_headers,
        )

        assert result["rows_written"] == 5
        assert os.path.exists(result["file_path"])

    def test_write_custom_sheet_name(
        self,
        xlsxwriter_adapter: XlsxWriterAdapter,
        temp_dir: Path,
        sample_data: list[list],
    ) -> None:
        """Test writing data with a custom sheet name."""
        file_path = temp_dir / "custom_sheet.xlsx"

        result = xlsxwriter_adapter.write_sheet(
            file_path=str(file_path),
            rows=sample_data,
            sheet_name="CustomData",
        )

        assert result["rows_written"] == 5

    def test_write_with_start_cell(
        self,
        xlsxwriter_adapter: XlsxWriterAdapter,
        temp_dir: Path,
        sample_data: list[list],
    ) -> None:
        """Test writing data starting from a specific cell."""
        file_path = temp_dir / "offset_start.xlsx"

        result = xlsxwriter_adapter.write_sheet(
            file_path=str(file_path),
            rows=sample_data,
            start_cell="B5",
        )

        assert result["rows_written"] == 5


class TestXlsxWriterAdapterDataTypes:
    """Tests for writing different data types."""

    def test_write_mixed_types(
        self,
        xlsxwriter_adapter: XlsxWriterAdapter,
        temp_dir: Path,
    ) -> None:
        """Test writing data with mixed types."""
        file_path = temp_dir / "mixed_types.xlsx"

        rows = [
            ["String", 42, 3.14, True, None],
            ["Another", 100, 2.718, False, "Value"],
        ]

        result = xlsxwriter_adapter.write_sheet(
            file_path=str(file_path),
            rows=rows,
            headers=["Text", "Integer", "Float", "Boolean", "Nullable"],
        )

        assert result["rows_written"] == 2

    def test_write_datetime(
        self,
        xlsxwriter_adapter: XlsxWriterAdapter,
        temp_dir: Path,
    ) -> None:
        """Test writing datetime values."""
        file_path = temp_dir / "datetime.xlsx"

        now = datetime.now()
        rows = [
            ["Event 1", now],
            ["Event 2", datetime(2024, 1, 15, 10, 30)],
        ]

        result = xlsxwriter_adapter.write_sheet(
            file_path=str(file_path),
            rows=rows,
            headers=["Event", "Timestamp"],
        )

        assert result["rows_written"] == 2

    def test_write_formula(
        self,
        xlsxwriter_adapter: XlsxWriterAdapter,
        temp_dir: Path,
    ) -> None:
        """Test writing formulas."""
        file_path = temp_dir / "formula.xlsx"

        rows = [
            [10, 20, "=A1+B1"],
            [15, 25, "=A2+B2"],
        ]

        result = xlsxwriter_adapter.write_sheet(
            file_path=str(file_path),
            rows=rows,
            headers=["A", "B", "Sum"],
        )

        assert result["rows_written"] == 2


class TestXlsxWriterAdapterOverwrite:
    """Tests for overwrite behavior."""

    def test_write_fails_without_overwrite(
        self,
        xlsxwriter_adapter: XlsxWriterAdapter,
        temp_dir: Path,
        sample_data: list[list],
    ) -> None:
        """Test that writing to existing file fails without overwrite flag."""
        file_path = temp_dir / "existing.xlsx"

        xlsxwriter_adapter.write_sheet(
            file_path=str(file_path),
            rows=sample_data,
        )

        with pytest.raises(WriteError) as exc_info:
            xlsxwriter_adapter.write_sheet(
                file_path=str(file_path),
                rows=sample_data,
                overwrite=False,
            )

        assert exc_info.value.error_code == "WRITE_ERROR"
        assert "already exists" in exc_info.value.message

    def test_write_succeeds_with_overwrite(
        self,
        xlsxwriter_adapter: XlsxWriterAdapter,
        temp_dir: Path,
        sample_data: list[list],
    ) -> None:
        """Test that writing to existing file succeeds with overwrite flag."""
        file_path = temp_dir / "overwrite.xlsx"

        xlsxwriter_adapter.write_sheet(
            file_path=str(file_path),
            rows=sample_data,
        )

        new_data = [["New", "Data"]]
        result = xlsxwriter_adapter.write_sheet(
            file_path=str(file_path),
            rows=new_data,
            overwrite=True,
        )

        assert result["rows_written"] == 1


class TestXlsxWriterAdapterMultipleSheets:
    """Tests for writing multiple sheets."""

    def test_write_multiple_sheets(
        self,
        xlsxwriter_adapter: XlsxWriterAdapter,
        temp_dir: Path,
    ) -> None:
        """Test writing multiple sheets to a single file."""
        file_path = temp_dir / "multi_sheet.xlsx"

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

        result = xlsxwriter_adapter.write_multiple_sheets(
            file_path=str(file_path),
            sheets_data=sheets_data,
        )

        assert result["sheets_written"] == 2
        assert result["total_rows_written"] == 4

    def test_write_empty_sheet(
        self,
        xlsxwriter_adapter: XlsxWriterAdapter,
        temp_dir: Path,
    ) -> None:
        """Test writing an empty sheet."""
        file_path = temp_dir / "empty_sheet.xlsx"

        sheets_data = {
            "EmptySheet": {
                "rows": [],
            },
            "DataSheet": {
                "rows": [["Data"]],
            },
        }

        result = xlsxwriter_adapter.write_multiple_sheets(
            file_path=str(file_path),
            sheets_data=sheets_data,
        )

        assert result["sheets_written"] == 2
        assert result["total_rows_written"] == 1


class TestXlsxWriterAdapterColumnWidth:
    """Tests for column width calculation."""

    def test_auto_format_column_widths(
        self,
        xlsxwriter_adapter: XlsxWriterAdapter,
        temp_dir: Path,
    ) -> None:
        """Test that auto_format adjusts column widths."""
        file_path = temp_dir / "auto_format.xlsx"

        rows = [
            ["Short", "This is a much longer text value"],
        ]

        result = xlsxwriter_adapter.write_sheet(
            file_path=str(file_path),
            rows=rows,
            auto_format=True,
        )

        assert result["rows_written"] == 1
        assert os.path.exists(result["file_path"])

    def test_no_auto_format(
        self,
        xlsxwriter_adapter: XlsxWriterAdapter,
        temp_dir: Path,
    ) -> None:
        """Test writing without auto formatting."""
        file_path = temp_dir / "no_format.xlsx"

        rows = [["Data1", "Data2"]]

        result = xlsxwriter_adapter.write_sheet(
            file_path=str(file_path),
            rows=rows,
            auto_format=False,
        )

        assert result["rows_written"] == 1


class TestXlsxWriterAdapterDirectoryCreation:
    """Tests for directory creation behavior."""

    def test_create_parent_directories(
        self,
        xlsxwriter_adapter: XlsxWriterAdapter,
        temp_dir: Path,
    ) -> None:
        """Test that parent directories are created if they don't exist."""
        file_path = temp_dir / "nested" / "path" / "output.xlsx"

        result = xlsxwriter_adapter.write_sheet(
            file_path=str(file_path),
            rows=[["Data"]],
        )

        assert os.path.exists(result["file_path"])
        assert os.path.exists(temp_dir / "nested" / "path")
