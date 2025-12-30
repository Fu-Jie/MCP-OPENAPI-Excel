"""
Pydantic models for Excel service operations.

This module contains all data models used for request/response validation
and serialization in both the FastAPI and MCP interfaces.

All models use Pydantic v2 for validation, serialization, and
JSON Schema generation for OpenAPI documentation.
"""

from datetime import datetime
from enum import Enum
from typing import Any

from pydantic import BaseModel, Field, field_validator


class CellValueType(str, Enum):
    """
    Enumeration of supported cell value types.

    Used to explicitly communicate the type of data stored in a cell,
    enabling proper type handling on the client side.
    """

    STRING = "string"
    INTEGER = "integer"
    FLOAT = "float"
    BOOLEAN = "boolean"
    DATETIME = "datetime"
    DATE = "date"
    TIME = "time"
    EMPTY = "empty"
    ERROR = "error"
    FORMULA = "formula"


class CellValue(BaseModel):
    """
    Represents a single cell value with type information.

    This model provides explicit type information along with the value,
    enabling proper serialization and client-side type handling.

    Attributes:
        value: The actual cell value (can be any JSON-serializable type).
        value_type: The semantic type of the value.
        raw_value: The original raw value before any type conversion.
        formatted: The formatted string representation (if available).
    """

    value: str | int | float | bool | datetime | None = Field(
        default=None,
        description="The cell value, can be string, number, boolean, or datetime",
    )
    value_type: CellValueType = Field(
        default=CellValueType.EMPTY,
        description="The semantic type of the cell value",
    )
    raw_value: str | None = Field(
        default=None,
        description="Original raw string value before type conversion",
    )
    formatted: str | None = Field(
        default=None,
        description="Formatted string representation of the value",
    )

    model_config = {
        "json_schema_extra": {
            "examples": [
                {"value": "Hello World", "value_type": "string"},
                {"value": 42, "value_type": "integer"},
                {"value": 3.14159, "value_type": "float"},
                {"value": True, "value_type": "boolean"},
            ]
        }
    }


class CellRange(BaseModel):
    """
    Represents a cell range specification.

    Supports both A1 notation (e.g., "A1:C10") and numeric indexing.

    Attributes:
        start_row: Starting row index (0-based).
        end_row: Ending row index (0-based, inclusive).
        start_col: Starting column index (0-based).
        end_col: Ending column index (0-based, inclusive).
        a1_notation: Optional A1 notation string (e.g., "A1:C10").
    """

    start_row: int = Field(
        ge=0,
        description="Starting row index (0-based)",
    )
    end_row: int = Field(
        ge=0,
        description="Ending row index (0-based, inclusive)",
    )
    start_col: int = Field(
        ge=0,
        description="Starting column index (0-based)",
    )
    end_col: int = Field(
        ge=0,
        description="Ending column index (0-based, inclusive)",
    )
    a1_notation: str | None = Field(
        default=None,
        description="Optional A1 notation string (e.g., 'A1:C10')",
    )

    @field_validator("end_row")
    @classmethod
    def validate_end_row(cls, v: int, info) -> int:
        """Ensure end_row is greater than or equal to start_row."""
        if "start_row" in info.data and v < info.data["start_row"]:
            raise ValueError("end_row must be >= start_row")
        return v

    @field_validator("end_col")
    @classmethod
    def validate_end_col(cls, v: int, info) -> int:
        """Ensure end_col is greater than or equal to start_col."""
        if "start_col" in info.data and v < info.data["start_col"]:
            raise ValueError("end_col must be >= start_col")
        return v


class SheetInfo(BaseModel):
    """
    Metadata about a single worksheet.

    Attributes:
        name: The name of the sheet.
        index: The 0-based index of the sheet in the workbook.
        visible: Whether the sheet is visible.
        row_count: Number of rows with data (if known).
        column_count: Number of columns with data (if known).
    """

    name: str = Field(
        description="The name of the sheet",
    )
    index: int = Field(
        ge=0,
        description="The 0-based index of the sheet in the workbook",
    )
    visible: bool = Field(
        default=True,
        description="Whether the sheet is visible",
    )
    row_count: int | None = Field(
        default=None,
        ge=0,
        description="Number of rows with data (if known)",
    )
    column_count: int | None = Field(
        default=None,
        ge=0,
        description="Number of columns with data (if known)",
    )


class WorkbookInfo(BaseModel):
    """
    Metadata about an Excel workbook.

    Attributes:
        file_path: Path to the Excel file.
        file_size_bytes: Size of the file in bytes.
        sheet_count: Number of sheets in the workbook.
        sheets: List of sheet metadata.
        created_at: File creation timestamp (if available).
        modified_at: File modification timestamp (if available).
    """

    file_path: str = Field(
        description="Path to the Excel file",
    )
    file_size_bytes: int | None = Field(
        default=None,
        ge=0,
        description="Size of the file in bytes",
    )
    sheet_count: int = Field(
        ge=0,
        description="Number of sheets in the workbook",
    )
    sheets: list[SheetInfo] = Field(
        default_factory=list,
        description="List of sheet metadata",
    )
    created_at: datetime | None = Field(
        default=None,
        description="File creation timestamp (if available)",
    )
    modified_at: datetime | None = Field(
        default=None,
        description="File modification timestamp (if available)",
    )


class SheetData(BaseModel):
    """
    Data from a single worksheet.

    Contains the actual cell data along with metadata about the sheet.

    Attributes:
        sheet_name: Name of the sheet.
        rows: List of rows, where each row is a list of cell values.
        row_count: Number of rows in the data.
        column_count: Number of columns in the data.
        headers: Optional list of column headers (first row if has_headers=True).
        cell_range: The cell range that was read (if specified).
    """

    sheet_name: str = Field(
        description="Name of the sheet",
    )
    rows: list[list[Any]] = Field(
        default_factory=list,
        description="List of rows, where each row is a list of cell values",
    )
    row_count: int = Field(
        ge=0,
        description="Number of rows in the data",
    )
    column_count: int = Field(
        ge=0,
        description="Number of columns in the data",
    )
    headers: list[str] | None = Field(
        default=None,
        description="Optional list of column headers",
    )
    cell_range: CellRange | None = Field(
        default=None,
        description="The cell range that was read (if specified)",
    )


class ReadExcelRequest(BaseModel):
    """
    Request model for reading Excel files.

    Attributes:
        file_path: Path to the Excel file to read.
        sheet_name: Name of the sheet to read. If None, reads the first sheet.
        sheet_index: Index of the sheet to read (0-based). Used if sheet_name is None.
        cell_range: Optional cell range to read (e.g., "A1:C10").
        include_headers: Whether to treat the first row as headers.
        skip_empty_rows: Whether to skip empty rows in the output.
        date_format: Format string for parsing date values.
    """

    file_path: str = Field(
        description="Path to the Excel file to read",
    )
    sheet_name: str | None = Field(
        default=None,
        description="Name of the sheet to read. If None, reads the first sheet.",
    )
    sheet_index: int | None = Field(
        default=None,
        ge=0,
        description="Index of the sheet to read (0-based). Used if sheet_name is None.",
    )
    cell_range: str | None = Field(
        default=None,
        description="Optional cell range to read in A1 notation (e.g., 'A1:C10')",
    )
    include_headers: bool = Field(
        default=False,
        description="Whether to treat the first row as headers",
    )
    skip_empty_rows: bool = Field(
        default=False,
        description="Whether to skip empty rows in the output",
    )
    date_format: str | None = Field(
        default=None,
        description="Format string for parsing date values (e.g., '%Y-%m-%d')",
    )


class ReadExcelResponse(BaseModel):
    """
    Response model for Excel read operations.

    Attributes:
        success: Whether the read operation was successful.
        workbook_info: Metadata about the workbook.
        sheet_data: Data from the requested sheet.
        processing_time_ms: Time taken to process the request in milliseconds.
        message: Optional message (e.g., warnings or info).
    """

    success: bool = Field(
        default=True,
        description="Whether the read operation was successful",
    )
    workbook_info: WorkbookInfo = Field(
        description="Metadata about the workbook",
    )
    sheet_data: SheetData = Field(
        description="Data from the requested sheet",
    )
    processing_time_ms: float | None = Field(
        default=None,
        ge=0,
        description="Time taken to process the request in milliseconds",
    )
    message: str | None = Field(
        default=None,
        description="Optional message (e.g., warnings or info)",
    )


class WriteExcelRequest(BaseModel):
    """
    Request model for writing Excel files.

    Attributes:
        file_path: Path where the Excel file will be written.
        sheet_name: Name of the sheet to write. Defaults to "Sheet1".
        rows: List of rows to write, where each row is a list of values.
        headers: Optional list of column headers.
        start_cell: Starting cell for writing data (A1 notation). Defaults to "A1".
        overwrite: Whether to overwrite an existing file.
        auto_format: Whether to auto-format column widths and data types.
    """

    file_path: str = Field(
        description="Path where the Excel file will be written",
    )
    sheet_name: str = Field(
        default="Sheet1",
        description="Name of the sheet to write. Defaults to 'Sheet1'.",
    )
    rows: list[list[Any]] = Field(
        description="List of rows to write, where each row is a list of values",
    )
    headers: list[str] | None = Field(
        default=None,
        description="Optional list of column headers",
    )
    start_cell: str = Field(
        default="A1",
        description="Starting cell for writing data (A1 notation). Defaults to 'A1'.",
    )
    overwrite: bool = Field(
        default=False,
        description="Whether to overwrite an existing file",
    )
    auto_format: bool = Field(
        default=True,
        description="Whether to auto-format column widths and data types",
    )

    @field_validator("rows")
    @classmethod
    def validate_rows_not_empty(cls, v: list[list[Any]]) -> list[list[Any]]:
        """Ensure at least one row is provided."""
        if len(v) == 0:
            raise ValueError("At least one row of data is required")
        return v


class WriteExcelResponse(BaseModel):
    """
    Response model for Excel write operations.

    Attributes:
        success: Whether the write operation was successful.
        file_path: Path to the written file.
        rows_written: Number of rows written.
        file_size_bytes: Size of the written file in bytes.
        processing_time_ms: Time taken to process the request in milliseconds.
        message: Optional message (e.g., warnings or info).
    """

    success: bool = Field(
        default=True,
        description="Whether the write operation was successful",
    )
    file_path: str = Field(
        description="Path to the written file",
    )
    rows_written: int = Field(
        ge=0,
        description="Number of rows written",
    )
    file_size_bytes: int | None = Field(
        default=None,
        ge=0,
        description="Size of the written file in bytes",
    )
    processing_time_ms: float | None = Field(
        default=None,
        ge=0,
        description="Time taken to process the request in milliseconds",
    )
    message: str | None = Field(
        default=None,
        description="Optional message (e.g., warnings or info)",
    )


class ExcelErrorResponse(BaseModel):
    """
    Standard error response model for the API.

    Attributes:
        success: Always False for error responses.
        error_code: Machine-readable error code.
        message: Human-readable error description.
        details: Additional error context.
    """

    success: bool = Field(
        default=False,
        description="Always False for error responses",
    )
    error_code: str = Field(
        description="Machine-readable error code",
    )
    message: str = Field(
        description="Human-readable error description",
    )
    details: dict | None = Field(
        default=None,
        description="Additional error context",
    )
