"""
Core Excel service layer.

This module provides the ExcelService class which encapsulates all Excel
operations and serves as the single entry point for both FastAPI and MCP
interfaces. It implements the Service Layer pattern to ensure clean
separation between business logic and transport layers.

The service coordinates between the CalamineAdapter (reading) and
XlsxWriterAdapter (writing) to provide a unified interface for all
Excel operations.

Example:
    service = ExcelService()

    # Read operations
    info = service.get_workbook_info("/path/to/file.xlsx")
    data = service.read_excel(ReadExcelRequest(file_path="/path/to/file.xlsx"))

    # Write operations
    result = service.write_excel(WriteExcelRequest(
        file_path="/path/to/output.xlsx",
        rows=[["Name", "Age"], ["Alice", 30]]
    ))
"""

import time
from typing import Any

from src.adapters.calamine_adapter import CalamineAdapter
from src.adapters.xlsxwriter_adapter import XlsxWriterAdapter
from src.exceptions.excel_exceptions import ExcelServiceError
from src.models.excel_models import (
    ReadExcelRequest,
    ReadExcelResponse,
    SheetData,
    SheetInfo,
    WorkbookInfo,
    WriteExcelRequest,
    WriteExcelResponse,
)


class ExcelService:
    """
    Core service layer for Excel operations.

    This service provides a unified interface for all Excel read and write
    operations. It is designed to be consumed by both the FastAPI REST API
    and the MCP protocol adapter.

    The service uses:
        - CalamineAdapter: For high-performance reading (Rust-based)
        - XlsxWriterAdapter: For high-performance writing (streaming)

    All methods are designed to be transport-agnostic and return strongly
    typed Pydantic models for serialization.

    Attributes:
        read_adapter: CalamineAdapter instance for read operations.
        write_adapter: XlsxWriterAdapter instance for write operations.

    Example:
        service = ExcelService()

        # Get workbook metadata
        info = service.get_workbook_info("/path/to/file.xlsx")
        print(f"Workbook has {info.sheet_count} sheets")

        # Read specific sheet
        data = service.read_sheet("/path/to/file.xlsx", sheet_name="Data")
        for row in data.rows:
            print(row)

        # Write new file
        result = service.write_excel(WriteExcelRequest(
            file_path="/path/to/output.xlsx",
            rows=[["Name", "Age"], ["Alice", 30], ["Bob", 25]]
        ))
    """

    def __init__(
        self,
        read_adapter: CalamineAdapter | None = None,
        write_adapter: XlsxWriterAdapter | None = None,
    ) -> None:
        """
        Initialize the ExcelService.

        Args:
            read_adapter: Optional CalamineAdapter instance.
                         If None, creates a new instance.
            write_adapter: Optional XlsxWriterAdapter instance.
                          If None, creates a new instance.
        """
        self.read_adapter = read_adapter or CalamineAdapter()
        self.write_adapter = write_adapter or XlsxWriterAdapter()

    def get_sheet_names(self, file_path: str) -> list[str]:
        """
        Get the list of sheet names in an Excel workbook.

        Args:
            file_path: Path to the Excel file.

        Returns:
            List of sheet names.

        Raises:
            FileNotFoundError: If the file does not exist.
            InvalidFileFormatError: If the file format is not supported.
        """
        return self.read_adapter.get_sheet_names(file_path)

    def get_workbook_info(self, file_path: str) -> WorkbookInfo:
        """
        Get metadata about an Excel workbook.

        Retrieves comprehensive information about the workbook including
        file size, sheet count, and detailed metadata for each sheet.

        Args:
            file_path: Path to the Excel file.

        Returns:
            WorkbookInfo containing workbook metadata.

        Raises:
            FileNotFoundError: If the file does not exist.
            InvalidFileFormatError: If the file format is not supported.
        """
        return self.read_adapter.get_workbook_info(file_path)

    def get_sheet_info(
        self,
        file_path: str,
        sheet_name: str | None = None,
        sheet_index: int | None = None,
    ) -> SheetInfo:
        """
        Get metadata about a specific sheet.

        Args:
            file_path: Path to the Excel file.
            sheet_name: Name of the sheet. If None and sheet_index is None,
                       uses the first sheet.
            sheet_index: Index of the sheet (0-based). Used if sheet_name is None.

        Returns:
            SheetInfo containing sheet metadata.

        Raises:
            FileNotFoundError: If the file does not exist.
            InvalidFileFormatError: If the file format is not supported.
            SheetNotFoundError: If the specified sheet does not exist.
        """
        workbook_info = self.read_adapter.get_workbook_info(file_path)

        if sheet_name is not None:
            for sheet in workbook_info.sheets:
                if sheet.name == sheet_name:
                    return sheet
            from src.exceptions.excel_exceptions import SheetNotFoundError

            raise SheetNotFoundError(
                sheet_name=sheet_name,
                available_sheets=[s.name for s in workbook_info.sheets],
            )

        if sheet_index is not None:
            if 0 <= sheet_index < len(workbook_info.sheets):
                return workbook_info.sheets[sheet_index]
            from src.exceptions.excel_exceptions import SheetNotFoundError

            raise SheetNotFoundError(
                sheet_name=f"index {sheet_index}",
                available_sheets=[s.name for s in workbook_info.sheets],
            )

        if workbook_info.sheets:
            return workbook_info.sheets[0]

        from src.exceptions.excel_exceptions import SheetNotFoundError

        raise SheetNotFoundError(
            sheet_name="(first sheet)",
            available_sheets=[],
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
            sheet_name: Name of the sheet to read. If None and sheet_index
                       is None, reads the first sheet.
            sheet_index: Index of the sheet (0-based). Used if sheet_name is None.
            skip_empty_rows: Whether to skip empty rows.

        Returns:
            SheetData containing the sheet contents.

        Raises:
            FileNotFoundError: If the file does not exist.
            InvalidFileFormatError: If the file format is not supported.
            SheetNotFoundError: If the specified sheet does not exist.
        """
        return self.read_adapter.read_sheet(
            file_path=file_path,
            sheet_name=sheet_name,
            sheet_index=sheet_index,
            skip_empty_rows=skip_empty_rows,
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
            sheet_name: Name of the sheet. If None, uses the first sheet.
            sheet_index: Index of the sheet (0-based). Used if sheet_name is None.

        Returns:
            SheetData containing the range contents.

        Raises:
            FileNotFoundError: If the file does not exist.
            InvalidFileFormatError: If the file format is not supported.
            SheetNotFoundError: If the specified sheet does not exist.
            CellRangeError: If the cell range is invalid.
        """
        return self.read_adapter.read_range(
            file_path=file_path,
            cell_range=cell_range,
            sheet_name=sheet_name,
            sheet_index=sheet_index,
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
        return self.read_adapter.get_cell_value(
            file_path=file_path,
            cell=cell,
            sheet_name=sheet_name,
            sheet_index=sheet_index,
        )

    def read_excel(self, request: ReadExcelRequest) -> ReadExcelResponse:
        """
        Read Excel file with comprehensive options.

        This is the main read method that handles all read options including
        cell ranges, headers, and empty row filtering.

        Args:
            request: ReadExcelRequest containing all read parameters.

        Returns:
            ReadExcelResponse containing workbook info and sheet data.

        Raises:
            FileNotFoundError: If the file does not exist.
            InvalidFileFormatError: If the file format is not supported.
            SheetNotFoundError: If the specified sheet does not exist.
            CellRangeError: If the cell range is invalid.
        """
        start_time = time.time()

        try:
            workbook_info = self.read_adapter.get_workbook_info(request.file_path)

            if request.cell_range:
                sheet_data = self.read_adapter.read_range(
                    file_path=request.file_path,
                    cell_range=request.cell_range,
                    sheet_name=request.sheet_name,
                    sheet_index=request.sheet_index,
                )
            else:
                sheet_data = self.read_adapter.read_sheet(
                    file_path=request.file_path,
                    sheet_name=request.sheet_name,
                    sheet_index=request.sheet_index,
                    skip_empty_rows=request.skip_empty_rows,
                )

            if request.include_headers and sheet_data.rows:
                first_row = sheet_data.rows[0]
                sheet_data.headers = [str(cell) if cell is not None else "" for cell in first_row]
                sheet_data.rows = sheet_data.rows[1:]
                sheet_data.row_count = len(sheet_data.rows)

            processing_time = (time.time() - start_time) * 1000

            return ReadExcelResponse(
                success=True,
                workbook_info=workbook_info,
                sheet_data=sheet_data,
                processing_time_ms=round(processing_time, 2),
            )

        except ExcelServiceError:
            raise
        except Exception as e:
            from src.exceptions.excel_exceptions import ReadError

            raise ReadError(
                file_path=request.file_path,
                operation="read",
                reason=str(e),
            ) from e

    def write_excel(self, request: WriteExcelRequest) -> WriteExcelResponse:
        """
        Write data to an Excel file.

        This is the main write method that handles all write options including
        headers, formatting, and overwrite behavior.

        Args:
            request: WriteExcelRequest containing all write parameters.

        Returns:
            WriteExcelResponse containing write results.

        Raises:
            WriteError: If writing fails.
            ExcelPermissionError: If the file cannot be written due to permissions.
        """
        start_time = time.time()

        try:
            result = self.write_adapter.write_sheet(
                file_path=request.file_path,
                rows=request.rows,
                sheet_name=request.sheet_name,
                headers=request.headers,
                start_cell=request.start_cell,
                overwrite=request.overwrite,
                auto_format=request.auto_format,
            )

            processing_time = (time.time() - start_time) * 1000

            return WriteExcelResponse(
                success=True,
                file_path=result["file_path"],
                rows_written=result["rows_written"],
                file_size_bytes=result["file_size_bytes"],
                processing_time_ms=round(processing_time, 2),
            )

        except ExcelServiceError:
            raise
        except Exception as e:
            from src.exceptions.excel_exceptions import WriteError

            raise WriteError(
                file_path=request.file_path,
                operation="write",
                reason=str(e),
            ) from e

    def write_multiple_sheets(
        self,
        file_path: str,
        sheets_data: dict[str, dict[str, Any]],
        overwrite: bool = False,
        auto_format: bool = True,
    ) -> dict[str, Any]:
        """
        Write multiple sheets to an Excel file.

        Args:
            file_path: Path where the file will be written.
            sheets_data: Dictionary mapping sheet names to configurations.
                Each configuration should contain:
                    - rows: List of data rows (required)
                    - headers: Optional list of headers
                    - start_cell: Starting cell (default: "A1")
            overwrite: Whether to overwrite an existing file.
            auto_format: Whether to auto-format column widths.

        Returns:
            Dictionary containing write results.

        Raises:
            WriteError: If writing fails.
            ExcelPermissionError: If the file cannot be written due to permissions.
        """
        return self.write_adapter.write_multiple_sheets(
            file_path=file_path,
            sheets_data=sheets_data,
            overwrite=overwrite,
            auto_format=auto_format,
        )
