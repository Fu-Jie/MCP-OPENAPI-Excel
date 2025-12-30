"""
Test fixtures and utilities for the Excel service tests.

This module provides shared fixtures including temporary files,
mock data, and service instances.
"""

import tempfile
from collections.abc import Generator
from pathlib import Path

import pytest

from src.adapters.calamine_adapter import CalamineAdapter
from src.adapters.xlsxwriter_adapter import XlsxWriterAdapter
from src.services.excel_service import ExcelService


@pytest.fixture
def excel_service() -> ExcelService:
    """
    Create an ExcelService instance for testing.

    Returns:
        ExcelService instance.
    """
    return ExcelService()


@pytest.fixture
def calamine_adapter() -> CalamineAdapter:
    """
    Create a CalamineAdapter instance for testing.

    Returns:
        CalamineAdapter instance.
    """
    return CalamineAdapter()


@pytest.fixture
def xlsxwriter_adapter() -> XlsxWriterAdapter:
    """
    Create an XlsxWriterAdapter instance for testing.

    Returns:
        XlsxWriterAdapter instance.
    """
    return XlsxWriterAdapter()


@pytest.fixture
def temp_dir() -> Generator[Path, None, None]:
    """
    Create a temporary directory for test files.

    Yields:
        Path to the temporary directory.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        yield Path(tmpdir)


@pytest.fixture
def sample_excel_file(temp_dir: Path, xlsxwriter_adapter: XlsxWriterAdapter) -> Path:
    """
    Create a sample Excel file for testing.

    Args:
        temp_dir: Temporary directory path.
        xlsxwriter_adapter: XlsxWriterAdapter instance.

    Returns:
        Path to the sample Excel file.
    """
    file_path = temp_dir / "sample.xlsx"

    xlsxwriter_adapter.write_sheet(
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
def multi_sheet_excel_file(temp_dir: Path, xlsxwriter_adapter: XlsxWriterAdapter) -> Path:
    """
    Create an Excel file with multiple sheets for testing.

    Args:
        temp_dir: Temporary directory path.
        xlsxwriter_adapter: XlsxWriterAdapter instance.

    Returns:
        Path to the multi-sheet Excel file.
    """
    file_path = temp_dir / "multi_sheet.xlsx"

    xlsxwriter_adapter.write_multiple_sheets(
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


@pytest.fixture
def sample_data() -> list[list]:
    """
    Return sample data for write tests.

    Returns:
        List of rows with sample data.
    """
    return [
        ["Alice", 30, "Engineering"],
        ["Bob", 25, "Marketing"],
        ["Charlie", 35, "Sales"],
        ["Diana", 28, "Engineering"],
        ["Eve", 32, "Marketing"],
    ]


@pytest.fixture
def sample_headers() -> list[str]:
    """
    Return sample headers for write tests.

    Returns:
        List of column headers.
    """
    return ["Name", "Age", "Department"]
