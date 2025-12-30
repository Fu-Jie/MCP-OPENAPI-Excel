"""
Tests for the ExcelService.

Tests the core service layer that coordinates between adapters.
"""

from pathlib import Path

import pytest

from src.exceptions.excel_exceptions import SheetNotFoundError
from src.models.excel_models import ReadExcelRequest, WriteExcelRequest
from src.services.excel_service import ExcelService


class TestExcelServiceReadOperations:
    """Tests for ExcelService read operations."""

    def test_get_sheet_names(
        self,
        excel_service: ExcelService,
        sample_excel_file: Path,
    ) -> None:
        """Test getting sheet names through the service."""
        sheet_names = excel_service.get_sheet_names(str(sample_excel_file))

        assert isinstance(sheet_names, list)
        assert len(sheet_names) == 1
        assert "Users" in sheet_names

    def test_get_workbook_info(
        self,
        excel_service: ExcelService,
        multi_sheet_excel_file: Path,
    ) -> None:
        """Test getting workbook info through the service."""
        info = excel_service.get_workbook_info(str(multi_sheet_excel_file))

        assert info.sheet_count == 3
        assert len(info.sheets) == 3
        assert info.file_size_bytes > 0

    def test_get_sheet_info_by_name(
        self,
        excel_service: ExcelService,
        multi_sheet_excel_file: Path,
    ) -> None:
        """Test getting sheet info by name."""
        info = excel_service.get_sheet_info(
            str(multi_sheet_excel_file),
            sheet_name="Products",
        )

        assert info.name == "Products"

    def test_get_sheet_info_by_index(
        self,
        excel_service: ExcelService,
        multi_sheet_excel_file: Path,
    ) -> None:
        """Test getting sheet info by index."""
        info = excel_service.get_sheet_info(
            str(multi_sheet_excel_file),
            sheet_index=1,
        )

        assert info.index == 1

    def test_get_sheet_info_not_found(
        self,
        excel_service: ExcelService,
        sample_excel_file: Path,
    ) -> None:
        """Test that getting non-existent sheet raises error."""
        with pytest.raises(SheetNotFoundError):
            excel_service.get_sheet_info(
                str(sample_excel_file),
                sheet_name="NonExistent",
            )

    def test_read_sheet(
        self,
        excel_service: ExcelService,
        sample_excel_file: Path,
    ) -> None:
        """Test reading a sheet through the service."""
        data = excel_service.read_sheet(str(sample_excel_file))

        assert data.sheet_name == "Users"
        assert data.row_count == 4
        assert data.column_count == 3

    def test_read_range(
        self,
        excel_service: ExcelService,
        sample_excel_file: Path,
    ) -> None:
        """Test reading a range through the service."""
        data = excel_service.read_range(
            str(sample_excel_file),
            cell_range="A1:B2",
        )

        assert data.row_count == 2
        assert data.column_count == 2

    def test_get_cell_value(
        self,
        excel_service: ExcelService,
        sample_excel_file: Path,
    ) -> None:
        """Test getting a cell value through the service."""
        value = excel_service.get_cell_value(
            str(sample_excel_file),
            cell="A2",
        )

        assert value == "Alice"


class TestExcelServiceReadExcel:
    """Tests for the read_excel method."""

    def test_read_excel_basic(
        self,
        excel_service: ExcelService,
        sample_excel_file: Path,
    ) -> None:
        """Test basic read_excel operation."""
        request = ReadExcelRequest(file_path=str(sample_excel_file))
        response = excel_service.read_excel(request)

        assert response.success is True
        assert response.workbook_info is not None
        assert response.sheet_data is not None
        assert response.processing_time_ms is not None
        assert response.processing_time_ms >= 0

    def test_read_excel_with_headers(
        self,
        excel_service: ExcelService,
        sample_excel_file: Path,
    ) -> None:
        """Test read_excel with header extraction."""
        request = ReadExcelRequest(
            file_path=str(sample_excel_file),
            include_headers=True,
        )
        response = excel_service.read_excel(request)

        assert response.success is True
        assert response.sheet_data.headers is not None
        assert response.sheet_data.headers == ["Name", "Age", "Email"]
        assert response.sheet_data.row_count == 3

    def test_read_excel_with_range(
        self,
        excel_service: ExcelService,
        sample_excel_file: Path,
    ) -> None:
        """Test read_excel with cell range."""
        request = ReadExcelRequest(
            file_path=str(sample_excel_file),
            cell_range="A1:B2",
        )
        response = excel_service.read_excel(request)

        assert response.success is True
        assert response.sheet_data.row_count == 2
        assert response.sheet_data.column_count == 2

    def test_read_excel_specific_sheet(
        self,
        excel_service: ExcelService,
        multi_sheet_excel_file: Path,
    ) -> None:
        """Test read_excel with specific sheet."""
        request = ReadExcelRequest(
            file_path=str(multi_sheet_excel_file),
            sheet_name="Products",
        )
        response = excel_service.read_excel(request)

        assert response.success is True
        assert response.sheet_data.sheet_name == "Products"


class TestExcelServiceWriteOperations:
    """Tests for ExcelService write operations."""

    def test_write_excel_basic(
        self,
        excel_service: ExcelService,
        temp_dir: Path,
        sample_data: list[list],
    ) -> None:
        """Test basic write_excel operation."""
        file_path = temp_dir / "write_test.xlsx"

        request = WriteExcelRequest(
            file_path=str(file_path),
            rows=sample_data,
        )
        response = excel_service.write_excel(request)

        assert response.success is True
        assert response.rows_written == 5
        assert response.file_size_bytes > 0
        assert response.processing_time_ms >= 0

    def test_write_excel_with_headers(
        self,
        excel_service: ExcelService,
        temp_dir: Path,
        sample_data: list[list],
        sample_headers: list[str],
    ) -> None:
        """Test write_excel with headers."""
        file_path = temp_dir / "write_headers.xlsx"

        request = WriteExcelRequest(
            file_path=str(file_path),
            rows=sample_data,
            headers=sample_headers,
        )
        response = excel_service.write_excel(request)

        assert response.success is True
        assert response.rows_written == 5

    def test_write_excel_custom_sheet(
        self,
        excel_service: ExcelService,
        temp_dir: Path,
        sample_data: list[list],
    ) -> None:
        """Test write_excel with custom sheet name."""
        file_path = temp_dir / "custom_sheet.xlsx"

        request = WriteExcelRequest(
            file_path=str(file_path),
            rows=sample_data,
            sheet_name="MyData",
        )
        response = excel_service.write_excel(request)

        assert response.success is True

    def test_write_multiple_sheets(
        self,
        excel_service: ExcelService,
        temp_dir: Path,
    ) -> None:
        """Test writing multiple sheets through the service."""
        file_path = temp_dir / "multi_sheet.xlsx"

        sheets_data = {
            "Sheet1": {
                "rows": [["A", "B"], [1, 2]],
                "headers": ["Col1", "Col2"],
            },
            "Sheet2": {
                "rows": [["X", "Y"], [3, 4]],
            },
        }

        result = excel_service.write_multiple_sheets(
            file_path=str(file_path),
            sheets_data=sheets_data,
        )

        assert result["sheets_written"] == 2


class TestExcelServiceRoundTrip:
    """Tests for read/write round-trip operations."""

    def test_write_then_read(
        self,
        excel_service: ExcelService,
        temp_dir: Path,
        sample_data: list[list],
        sample_headers: list[str],
    ) -> None:
        """Test writing and then reading back the same data."""
        file_path = temp_dir / "roundtrip.xlsx"

        write_request = WriteExcelRequest(
            file_path=str(file_path),
            rows=sample_data,
            headers=sample_headers,
        )
        write_response = excel_service.write_excel(write_request)
        assert write_response.success is True

        read_request = ReadExcelRequest(
            file_path=str(file_path),
            include_headers=True,
        )
        read_response = excel_service.read_excel(read_request)
        assert read_response.success is True

        assert read_response.sheet_data.headers == sample_headers
        assert read_response.sheet_data.row_count == len(sample_data)

        for i, row in enumerate(read_response.sheet_data.rows):
            for j, value in enumerate(row):
                if value is not None:
                    assert str(value) == str(sample_data[i][j]) or value == sample_data[i][j]
