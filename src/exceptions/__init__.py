"""
Custom exceptions for the Excel service.

Provides type-safe, descriptive exceptions for error handling throughout
the application.
"""

from src.exceptions.excel_exceptions import (
    CellRangeError,
    ExcelServiceError,
    InvalidFileFormatError,
    ReadError,
    SheetNotFoundError,
    WriteError,
)
from src.exceptions.excel_exceptions import (
    FileNotFoundError as ExcelFileNotFoundError,
)
from src.exceptions.excel_exceptions import (
    PermissionError as ExcelPermissionError,
)

__all__ = [
    "ExcelServiceError",
    "ExcelFileNotFoundError",
    "InvalidFileFormatError",
    "SheetNotFoundError",
    "CellRangeError",
    "WriteError",
    "ReadError",
    "ExcelPermissionError",
]
