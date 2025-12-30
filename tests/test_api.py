"""
Tests for the FastAPI REST API.

Tests the HTTP endpoints for Excel operations.
"""

import os
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from src.main import app
from src.services.excel_service import ExcelService


@pytest.fixture(scope="module")
def client():
    """Create a test client for the FastAPI app."""
    global excel_service
    from src import main

    main.excel_service = ExcelService()
    with TestClient(app) as c:
        yield c
    main.excel_service = None


@pytest.fixture
def temp_excel_file(xlsxwriter_adapter, temp_dir: Path) -> Path:
    """Create a temporary Excel file for API tests."""
    file_path = temp_dir / "api_test.xlsx"

    xlsxwriter_adapter.write_sheet(
        file_path=str(file_path),
        rows=[
            ["Alice", 30, "alice@example.com"],
            ["Bob", 25, "bob@example.com"],
        ],
        headers=["Name", "Age", "Email"],
        sheet_name="Users",
    )

    return file_path


class TestHealthEndpoint:
    """Tests for the health check endpoint."""

    def test_health_check(self, client: TestClient) -> None:
        """Test the health check endpoint."""
        response = client.get("/health")

        assert response.status_code == 200
        data = response.json()
        assert data["status"] == "healthy"
        assert "timestamp" in data


class TestWorkbookEndpoints:
    """Tests for workbook metadata endpoints."""

    def test_get_workbook_info(
        self,
        client: TestClient,
        temp_excel_file: Path,
    ) -> None:
        """Test getting workbook information."""
        response = client.get(
            "/workbook/info",
            params={"file_path": str(temp_excel_file)},
        )

        assert response.status_code == 200
        data = response.json()
        assert data["sheet_count"] == 1
        assert len(data["sheets"]) == 1

    def test_get_workbook_info_not_found(self, client: TestClient) -> None:
        """Test getting info for non-existent file."""
        response = client.get(
            "/workbook/info",
            params={"file_path": "/nonexistent/file.xlsx"},
        )

        assert response.status_code == 404

    def test_get_sheet_names(
        self,
        client: TestClient,
        temp_excel_file: Path,
    ) -> None:
        """Test getting sheet names."""
        response = client.get(
            "/workbook/sheets",
            params={"file_path": str(temp_excel_file)},
        )

        assert response.status_code == 200
        data = response.json()
        assert isinstance(data, list)
        assert "Users" in data

    def test_get_sheet_info(
        self,
        client: TestClient,
        temp_excel_file: Path,
    ) -> None:
        """Test getting sheet information."""
        response = client.get(
            "/workbook/sheet",
            params={
                "file_path": str(temp_excel_file),
                "sheet_name": "Users",
            },
        )

        assert response.status_code == 200
        data = response.json()
        assert data["name"] == "Users"


class TestReadEndpoints:
    """Tests for Excel read endpoints."""

    def test_read_excel_post(
        self,
        client: TestClient,
        temp_excel_file: Path,
    ) -> None:
        """Test the POST read endpoint."""
        response = client.post(
            "/excel/read",
            json={"file_path": str(temp_excel_file)},
        )

        assert response.status_code == 200
        data = response.json()
        assert data["success"] is True
        assert "sheet_data" in data
        assert "workbook_info" in data

    def test_read_excel_with_headers(
        self,
        client: TestClient,
        temp_excel_file: Path,
    ) -> None:
        """Test reading with header extraction."""
        response = client.post(
            "/excel/read",
            json={
                "file_path": str(temp_excel_file),
                "include_headers": True,
            },
        )

        assert response.status_code == 200
        data = response.json()
        assert data["sheet_data"]["headers"] == ["Name", "Age", "Email"]

    def test_read_sheet_get(
        self,
        client: TestClient,
        temp_excel_file: Path,
    ) -> None:
        """Test the GET read sheet endpoint."""
        response = client.get(
            "/excel/read/sheet",
            params={"file_path": str(temp_excel_file)},
        )

        assert response.status_code == 200
        data = response.json()
        assert data["sheet_name"] == "Users"
        assert data["row_count"] == 3

    def test_read_range_get(
        self,
        client: TestClient,
        temp_excel_file: Path,
    ) -> None:
        """Test the GET read range endpoint."""
        response = client.get(
            "/excel/read/range",
            params={
                "file_path": str(temp_excel_file),
                "cell_range": "A1:B2",
            },
        )

        assert response.status_code == 200
        data = response.json()
        assert data["row_count"] == 2
        assert data["column_count"] == 2

    def test_read_cell_get(
        self,
        client: TestClient,
        temp_excel_file: Path,
    ) -> None:
        """Test the GET read cell endpoint."""
        response = client.get(
            "/excel/read/cell",
            params={
                "file_path": str(temp_excel_file),
                "cell": "A2",
            },
        )

        assert response.status_code == 200
        data = response.json()
        assert data["cell"] == "A2"
        assert data["value"] == "Alice"


class TestWriteEndpoints:
    """Tests for Excel write endpoints."""

    def test_write_excel(
        self,
        client: TestClient,
        temp_dir: Path,
    ) -> None:
        """Test the write endpoint."""
        file_path = temp_dir / "api_write.xlsx"

        response = client.post(
            "/excel/write",
            json={
                "file_path": str(file_path),
                "rows": [["A", "B"], [1, 2]],
                "headers": ["Col1", "Col2"],
            },
        )

        assert response.status_code == 200
        data = response.json()
        assert data["success"] is True
        assert data["rows_written"] == 2
        assert os.path.exists(file_path)

    def test_write_excel_with_sheet_name(
        self,
        client: TestClient,
        temp_dir: Path,
    ) -> None:
        """Test writing with custom sheet name."""
        file_path = temp_dir / "custom_sheet.xlsx"

        response = client.post(
            "/excel/write",
            json={
                "file_path": str(file_path),
                "rows": [["Data"]],
                "sheet_name": "CustomSheet",
            },
        )

        assert response.status_code == 200
        data = response.json()
        assert data["success"] is True


class TestUploadEndpoint:
    """Tests for the file upload endpoint."""

    def test_upload_and_read(
        self,
        client: TestClient,
        temp_excel_file: Path,
    ) -> None:
        """Test uploading and reading an Excel file."""
        with open(temp_excel_file, "rb") as f:
            response = client.post(
                "/excel/upload",
                files={"file": ("test.xlsx", f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            )

        assert response.status_code == 200
        data = response.json()
        assert data["success"] is True
        assert "sheet_data" in data

    def test_upload_with_options(
        self,
        client: TestClient,
        temp_excel_file: Path,
    ) -> None:
        """Test uploading with read options."""
        with open(temp_excel_file, "rb") as f:
            response = client.post(
                "/excel/upload",
                files={"file": ("test.xlsx", f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
                params={"include_headers": True},
            )

        assert response.status_code == 200
        data = response.json()
        assert data["sheet_data"]["headers"] is not None

    def test_upload_invalid_extension(self, client: TestClient) -> None:
        """Test uploading a file with invalid extension."""
        response = client.post(
            "/excel/upload",
            files={"file": ("test.txt", b"not excel", "text/plain")},
        )

        assert response.status_code == 400


class TestErrorHandling:
    """Tests for API error handling."""

    def test_file_not_found_error(self, client: TestClient) -> None:
        """Test error response for non-existent file."""
        response = client.post(
            "/excel/read",
            json={"file_path": "/nonexistent/file.xlsx"},
        )

        assert response.status_code == 404
        data = response.json()
        assert "detail" in data

    def test_invalid_range_error(
        self,
        client: TestClient,
        temp_excel_file: Path,
    ) -> None:
        """Test error response for invalid cell range."""
        response = client.get(
            "/excel/read/range",
            params={
                "file_path": str(temp_excel_file),
                "cell_range": "invalid",
            },
        )

        assert response.status_code == 400

    def test_sheet_not_found_error(
        self,
        client: TestClient,
        temp_excel_file: Path,
    ) -> None:
        """Test error response for non-existent sheet."""
        response = client.get(
            "/excel/read/sheet",
            params={
                "file_path": str(temp_excel_file),
                "sheet_name": "NonExistent",
            },
        )

        assert response.status_code == 404
