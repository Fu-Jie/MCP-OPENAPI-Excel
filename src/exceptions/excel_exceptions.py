"""
Custom exceptions for Excel service operations.

This module defines a hierarchy of exceptions for handling various
error conditions during Excel file processing. All exceptions inherit
from ExcelServiceError for consistent error handling.

Example:
    try:
        service.read_excel("nonexistent.xlsx")
    except SheetNotFoundError as e:
        logger.error(f"Sheet error: {e.sheet_name}")
    except ExcelServiceError as e:
        logger.error(f"General error: {e}")
"""



class ExcelServiceError(Exception):
    """
    Base exception for all Excel service errors.

    All custom exceptions in this module inherit from this class,
    allowing consumers to catch all Excel-related errors with a single
    except clause.

    Attributes:
        message: Human-readable error description.
        error_code: Machine-readable error code for API responses.
        details: Optional additional context about the error.
    """

    def __init__(
        self,
        message: str,
        error_code: str = "EXCEL_ERROR",
        details: dict | None = None,
    ) -> None:
        """
        Initialize the ExcelServiceError.

        Args:
            message: Human-readable error description.
            error_code: Machine-readable error code for API responses.
            details: Optional additional context about the error.
        """
        super().__init__(message)
        self.message = message
        self.error_code = error_code
        self.details = details or {}

    def to_dict(self) -> dict:
        """
        Convert exception to a dictionary for API responses.

        Returns:
            Dictionary containing error_code, message, and details.
        """
        return {
            "error_code": self.error_code,
            "message": self.message,
            "details": self.details,
        }


class FileNotFoundError(ExcelServiceError):
    """
    Raised when the specified Excel file does not exist.

    Attributes:
        file_path: Path to the file that was not found.
    """

    def __init__(self, file_path: str) -> None:
        """
        Initialize the FileNotFoundError.

        Args:
            file_path: Path to the file that was not found.
        """
        self.file_path = file_path
        super().__init__(
            message=f"Excel file not found: {file_path}",
            error_code="FILE_NOT_FOUND",
            details={"file_path": file_path},
        )


class InvalidFileFormatError(ExcelServiceError):
    """
    Raised when the file is not a valid Excel format.

    This exception is raised when python-calamine cannot parse the file
    due to corruption, wrong format, or unsupported file type.

    Attributes:
        file_path: Path to the invalid file.
        expected_formats: List of expected/supported formats.
    """

    def __init__(
        self,
        file_path: str,
        expected_formats: list[str] | None = None,
        reason: str | None = None,
    ) -> None:
        """
        Initialize the InvalidFileFormatError.

        Args:
            file_path: Path to the invalid file.
            expected_formats: List of expected/supported formats.
            reason: Specific reason for the format error.
        """
        self.file_path = file_path
        self.expected_formats = expected_formats or [".xlsx", ".xls", ".xlsb", ".xlsm", ".ods"]
        self.reason = reason

        message = f"Invalid Excel file format: {file_path}"
        if reason:
            message += f" - {reason}"

        super().__init__(
            message=message,
            error_code="INVALID_FILE_FORMAT",
            details={
                "file_path": file_path,
                "expected_formats": self.expected_formats,
                "reason": reason,
            },
        )


class SheetNotFoundError(ExcelServiceError):
    """
    Raised when the specified sheet does not exist in the workbook.

    Attributes:
        sheet_name: Name of the sheet that was not found.
        available_sheets: List of sheets available in the workbook.
    """

    def __init__(
        self,
        sheet_name: str,
        available_sheets: list[str] | None = None,
    ) -> None:
        """
        Initialize the SheetNotFoundError.

        Args:
            sheet_name: Name of the sheet that was not found.
            available_sheets: List of sheets available in the workbook.
        """
        self.sheet_name = sheet_name
        self.available_sheets = available_sheets or []

        message = f"Sheet not found: {sheet_name}"
        if available_sheets:
            message += f". Available sheets: {', '.join(available_sheets)}"

        super().__init__(
            message=message,
            error_code="SHEET_NOT_FOUND",
            details={
                "sheet_name": sheet_name,
                "available_sheets": self.available_sheets,
            },
        )


class CellRangeError(ExcelServiceError):
    """
    Raised when an invalid cell range is specified.

    This covers invalid range syntax (e.g., "A1:"), out-of-bounds ranges,
    and ranges that don't make sense (e.g., end before start).

    Attributes:
        cell_range: The invalid cell range string.
        reason: Specific reason why the range is invalid.
    """

    def __init__(
        self,
        cell_range: str,
        reason: str | None = None,
    ) -> None:
        """
        Initialize the CellRangeError.

        Args:
            cell_range: The invalid cell range string.
            reason: Specific reason why the range is invalid.
        """
        self.cell_range = cell_range
        self.reason = reason

        message = f"Invalid cell range: {cell_range}"
        if reason:
            message += f" - {reason}"

        super().__init__(
            message=message,
            error_code="INVALID_CELL_RANGE",
            details={
                "cell_range": cell_range,
                "reason": reason,
            },
        )


class ReadError(ExcelServiceError):
    """
    Raised when an error occurs during Excel file reading.

    This is a general exception for read operations that fail
    for reasons not covered by more specific exceptions.

    Attributes:
        file_path: Path to the file being read.
        operation: The specific read operation that failed.
    """

    def __init__(
        self,
        file_path: str,
        operation: str = "read",
        reason: str | None = None,
    ) -> None:
        """
        Initialize the ReadError.

        Args:
            file_path: Path to the file being read.
            operation: The specific read operation that failed.
            reason: Specific reason for the read failure.
        """
        self.file_path = file_path
        self.operation = operation
        self.reason = reason

        message = f"Failed to {operation} Excel file: {file_path}"
        if reason:
            message += f" - {reason}"

        super().__init__(
            message=message,
            error_code="READ_ERROR",
            details={
                "file_path": file_path,
                "operation": operation,
                "reason": reason,
            },
        )


class WriteError(ExcelServiceError):
    """
    Raised when an error occurs during Excel file writing.

    This covers disk write failures, permission issues during write,
    and XlsxWriter-specific errors.

    Attributes:
        file_path: Path to the file being written.
        operation: The specific write operation that failed.
    """

    def __init__(
        self,
        file_path: str,
        operation: str = "write",
        reason: str | None = None,
    ) -> None:
        """
        Initialize the WriteError.

        Args:
            file_path: Path to the file being written.
            operation: The specific write operation that failed.
            reason: Specific reason for the write failure.
        """
        self.file_path = file_path
        self.operation = operation
        self.reason = reason

        message = f"Failed to {operation} Excel file: {file_path}"
        if reason:
            message += f" - {reason}"

        super().__init__(
            message=message,
            error_code="WRITE_ERROR",
            details={
                "file_path": file_path,
                "operation": operation,
                "reason": reason,
            },
        )


class PermissionError(ExcelServiceError):
    """
    Raised when file access is denied due to permissions.

    Attributes:
        file_path: Path to the file with permission issues.
        operation: The operation that was denied (read/write).
    """

    def __init__(
        self,
        file_path: str,
        operation: str = "access",
    ) -> None:
        """
        Initialize the PermissionError.

        Args:
            file_path: Path to the file with permission issues.
            operation: The operation that was denied (read/write).
        """
        self.file_path = file_path
        self.operation = operation

        super().__init__(
            message=f"Permission denied for {operation} on: {file_path}",
            error_code="PERMISSION_DENIED",
            details={
                "file_path": file_path,
                "operation": operation,
            },
        )
