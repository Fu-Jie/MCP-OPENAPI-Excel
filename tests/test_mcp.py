"""
Tests for the MCP server.

Tests the MCP protocol implementation for Excel operations.
"""

from pathlib import Path

import pytest

from src.mcp_server import MCPExcelServer
from src.services.excel_service import ExcelService


@pytest.fixture
def mcp_server(excel_service: ExcelService) -> MCPExcelServer:
    """Create an MCPExcelServer instance for testing."""
    return MCPExcelServer(service=excel_service)


class TestMCPServerTools:
    """Tests for MCP tool listing."""

    def test_get_tools_returns_all_tools(
        self,
        mcp_server: MCPExcelServer,
    ) -> None:
        """Test that all expected tools are returned."""
        tools = mcp_server._get_tools()

        tool_names = [tool.name for tool in tools]

        assert "get_workbook_info" in tool_names
        assert "list_sheets" in tool_names
        assert "read_sheet" in tool_names
        assert "read_range" in tool_names
        assert "read_cell" in tool_names
        assert "read_excel" in tool_names
        assert "write_excel" in tool_names

    def test_tools_have_required_fields(
        self,
        mcp_server: MCPExcelServer,
    ) -> None:
        """Test that all tools have required fields."""
        tools = mcp_server._get_tools()

        for tool in tools:
            assert tool.name is not None
            assert tool.description is not None
            assert tool.inputSchema is not None
            assert "properties" in tool.inputSchema
            assert "required" in tool.inputSchema


class TestMCPServerToolExecution:
    """Tests for MCP tool execution."""

    @pytest.mark.asyncio
    async def test_execute_get_workbook_info(
        self,
        mcp_server: MCPExcelServer,
        sample_excel_file: Path,
    ) -> None:
        """Test executing get_workbook_info tool."""
        result = await mcp_server._execute_tool(
            "get_workbook_info",
            {"file_path": str(sample_excel_file)},
        )

        assert result["success"] is True
        assert "data" in result
        assert result["data"]["sheet_count"] == 1

    @pytest.mark.asyncio
    async def test_execute_list_sheets(
        self,
        mcp_server: MCPExcelServer,
        sample_excel_file: Path,
    ) -> None:
        """Test executing list_sheets tool."""
        result = await mcp_server._execute_tool(
            "list_sheets",
            {"file_path": str(sample_excel_file)},
        )

        assert result["success"] is True
        assert "sheets" in result["data"]
        assert "Users" in result["data"]["sheets"]

    @pytest.mark.asyncio
    async def test_execute_read_sheet(
        self,
        mcp_server: MCPExcelServer,
        sample_excel_file: Path,
    ) -> None:
        """Test executing read_sheet tool."""
        result = await mcp_server._execute_tool(
            "read_sheet",
            {"file_path": str(sample_excel_file)},
        )

        assert result["success"] is True
        assert result["data"]["sheet_name"] == "Users"
        assert result["data"]["row_count"] == 4

    @pytest.mark.asyncio
    async def test_execute_read_sheet_by_name(
        self,
        mcp_server: MCPExcelServer,
        multi_sheet_excel_file: Path,
    ) -> None:
        """Test executing read_sheet with sheet_name."""
        result = await mcp_server._execute_tool(
            "read_sheet",
            {
                "file_path": str(multi_sheet_excel_file),
                "sheet_name": "Products",
            },
        )

        assert result["success"] is True
        assert result["data"]["sheet_name"] == "Products"

    @pytest.mark.asyncio
    async def test_execute_read_range(
        self,
        mcp_server: MCPExcelServer,
        sample_excel_file: Path,
    ) -> None:
        """Test executing read_range tool."""
        result = await mcp_server._execute_tool(
            "read_range",
            {
                "file_path": str(sample_excel_file),
                "cell_range": "A1:B2",
            },
        )

        assert result["success"] is True
        assert result["data"]["row_count"] == 2
        assert result["data"]["column_count"] == 2

    @pytest.mark.asyncio
    async def test_execute_read_cell(
        self,
        mcp_server: MCPExcelServer,
        sample_excel_file: Path,
    ) -> None:
        """Test executing read_cell tool."""
        result = await mcp_server._execute_tool(
            "read_cell",
            {
                "file_path": str(sample_excel_file),
                "cell": "A2",
            },
        )

        assert result["success"] is True
        assert result["data"]["cell"] == "A2"
        assert result["data"]["value"] == "Alice"

    @pytest.mark.asyncio
    async def test_execute_read_excel(
        self,
        mcp_server: MCPExcelServer,
        sample_excel_file: Path,
    ) -> None:
        """Test executing read_excel tool."""
        result = await mcp_server._execute_tool(
            "read_excel",
            {
                "file_path": str(sample_excel_file),
                "include_headers": True,
            },
        )

        assert result["success"] is True
        assert result["data"]["success"] is True
        assert result["data"]["sheet_data"]["headers"] == ["Name", "Age", "Email"]

    @pytest.mark.asyncio
    async def test_execute_write_excel(
        self,
        mcp_server: MCPExcelServer,
        temp_dir: Path,
    ) -> None:
        """Test executing write_excel tool."""
        file_path = temp_dir / "mcp_write.xlsx"

        result = await mcp_server._execute_tool(
            "write_excel",
            {
                "file_path": str(file_path),
                "rows": [["A", "B"], [1, 2]],
                "headers": ["Col1", "Col2"],
            },
        )

        assert result["success"] is True
        assert result["data"]["success"] is True
        assert result["data"]["rows_written"] == 2


class TestMCPServerErrorHandling:
    """Tests for MCP error handling."""

    @pytest.mark.asyncio
    async def test_unknown_tool_error(
        self,
        mcp_server: MCPExcelServer,
    ) -> None:
        """Test error for unknown tool name."""
        result = await mcp_server._execute_tool(
            "unknown_tool",
            {},
        )

        assert result["success"] is False
        assert result["error"]["error_code"] == "UNKNOWN_TOOL"

    @pytest.mark.asyncio
    async def test_file_not_found_error(
        self,
        mcp_server: MCPExcelServer,
    ) -> None:
        """Test error for non-existent file."""
        result = await mcp_server._execute_tool(
            "read_sheet",
            {"file_path": "/nonexistent/file.xlsx"},
        )

        assert result["success"] is False
        assert result["error"]["error_code"] == "FILE_NOT_FOUND"

    @pytest.mark.asyncio
    async def test_sheet_not_found_error(
        self,
        mcp_server: MCPExcelServer,
        sample_excel_file: Path,
    ) -> None:
        """Test error for non-existent sheet."""
        result = await mcp_server._execute_tool(
            "read_sheet",
            {
                "file_path": str(sample_excel_file),
                "sheet_name": "NonExistent",
            },
        )

        assert result["success"] is False
        assert result["error"]["error_code"] == "SHEET_NOT_FOUND"

    @pytest.mark.asyncio
    async def test_invalid_range_error(
        self,
        mcp_server: MCPExcelServer,
        sample_excel_file: Path,
    ) -> None:
        """Test error for invalid cell range."""
        result = await mcp_server._execute_tool(
            "read_range",
            {
                "file_path": str(sample_excel_file),
                "cell_range": "invalid",
            },
        )

        assert result["success"] is False
        assert result["error"]["error_code"] == "INVALID_CELL_RANGE"


class TestMCPServerIntegration:
    """Integration tests for MCP server."""

    @pytest.mark.asyncio
    async def test_write_then_read(
        self,
        mcp_server: MCPExcelServer,
        temp_dir: Path,
    ) -> None:
        """Test writing and then reading through MCP."""
        file_path = temp_dir / "mcp_roundtrip.xlsx"

        write_result = await mcp_server._execute_tool(
            "write_excel",
            {
                "file_path": str(file_path),
                "rows": [["Alice", 30], ["Bob", 25]],
                "headers": ["Name", "Age"],
            },
        )
        assert write_result["success"] is True

        read_result = await mcp_server._execute_tool(
            "read_sheet",
            {"file_path": str(file_path)},
        )
        assert read_result["success"] is True
        assert read_result["data"]["row_count"] == 3
