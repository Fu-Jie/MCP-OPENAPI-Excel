"""
MCP (Model Context Protocol) server for Excel operations.

This module implements an MCP server that exposes Excel operations as tools
that can be called by AI agents. It provides the same functionality as the
REST API but through the MCP protocol.

MCP Tools:
    - get_workbook_info: Get metadata about an Excel workbook
    - list_sheets: Get sheet names in a workbook
    - read_sheet: Read data from a specific sheet
    - read_range: Read a specific cell range
    - read_cell: Read a single cell value
    - write_excel: Write data to an Excel file

Example:
    To run the MCP server:
        python -m src.mcp_server

    Or programmatically:
        from src.mcp_server import run_mcp_server
        run_mcp_server()
"""

import asyncio
import json
from typing import Any

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import (
    TextContent,
    Tool,
)

from src.exceptions.excel_exceptions import ExcelServiceError
from src.models.excel_models import ReadExcelRequest, WriteExcelRequest
from src.services.excel_service import ExcelService


class MCPExcelServer:
    """
    MCP server implementation for Excel operations.

    This class wraps the ExcelService and exposes it through the MCP protocol,
    allowing AI agents to perform Excel operations using standardized tool calls.

    The server implements:
        - list_tools: Returns available Excel operations as MCP tools
        - call_tool: Executes a specific Excel operation

    Attributes:
        service: The underlying ExcelService instance.
        server: The MCP Server instance.

    Example:
        mcp_server = MCPExcelServer()
        await mcp_server.run()
    """

    def __init__(self, service: ExcelService | None = None) -> None:
        """
        Initialize the MCP Excel Server.

        Args:
            service: Optional ExcelService instance. If None, creates a new one.
        """
        self.service = service or ExcelService()
        self.server = Server("excel-mcp-server")
        self._setup_handlers()

    def _setup_handlers(self) -> None:
        """Set up MCP request handlers."""

        @self.server.list_tools()
        async def list_tools() -> list[Tool]:
            """Return the list of available Excel tools."""
            return self._get_tools()

        @self.server.call_tool()
        async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
            """Execute a tool and return the result."""
            result = await self._execute_tool(name, arguments)
            return [TextContent(type="text", text=json.dumps(result, default=str, indent=2))]

    def _get_tools(self) -> list[Tool]:
        """
        Get the list of available Excel tools.

        Returns:
            List of MCP Tool definitions.
        """
        return [
            Tool(
                name="get_workbook_info",
                description=(
                    "Get metadata about an Excel workbook including file size, "
                    "sheet count, and detailed information about each sheet."
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the Excel file",
                        },
                    },
                    "required": ["file_path"],
                },
            ),
            Tool(
                name="list_sheets",
                description="Get the list of sheet names in an Excel workbook.",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the Excel file",
                        },
                    },
                    "required": ["file_path"],
                },
            ),
            Tool(
                name="read_sheet",
                description=(
                    "Read all data from a specific sheet in an Excel workbook. "
                    "Returns the sheet contents as a list of rows."
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the Excel file",
                        },
                        "sheet_name": {
                            "type": "string",
                            "description": "Name of the sheet to read (optional, defaults to first sheet)",
                        },
                        "sheet_index": {
                            "type": "integer",
                            "description": "Index of the sheet (0-based, used if sheet_name not provided)",
                        },
                        "skip_empty_rows": {
                            "type": "boolean",
                            "description": "Whether to skip empty rows (default: false)",
                        },
                    },
                    "required": ["file_path"],
                },
            ),
            Tool(
                name="read_range",
                description=(
                    "Read a specific cell range from an Excel sheet. "
                    "The range should be in A1 notation (e.g., 'A1:C10')."
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the Excel file",
                        },
                        "cell_range": {
                            "type": "string",
                            "description": "Cell range in A1 notation (e.g., 'A1:C10')",
                        },
                        "sheet_name": {
                            "type": "string",
                            "description": "Name of the sheet (optional, defaults to first sheet)",
                        },
                        "sheet_index": {
                            "type": "integer",
                            "description": "Index of the sheet (0-based)",
                        },
                    },
                    "required": ["file_path", "cell_range"],
                },
            ),
            Tool(
                name="read_cell",
                description=(
                    "Read the value of a single cell in an Excel sheet. "
                    "The cell should be in A1 notation (e.g., 'A1', 'B5')."
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the Excel file",
                        },
                        "cell": {
                            "type": "string",
                            "description": "Cell reference in A1 notation (e.g., 'A1')",
                        },
                        "sheet_name": {
                            "type": "string",
                            "description": "Name of the sheet (optional, defaults to first sheet)",
                        },
                        "sheet_index": {
                            "type": "integer",
                            "description": "Index of the sheet (0-based)",
                        },
                    },
                    "required": ["file_path", "cell"],
                },
            ),
            Tool(
                name="read_excel",
                description=(
                    "Read Excel file with comprehensive options including cell range, "
                    "header extraction, and empty row filtering."
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the Excel file",
                        },
                        "sheet_name": {
                            "type": "string",
                            "description": "Name of the sheet to read",
                        },
                        "sheet_index": {
                            "type": "integer",
                            "description": "Index of the sheet (0-based)",
                        },
                        "cell_range": {
                            "type": "string",
                            "description": "Cell range in A1 notation (e.g., 'A1:C10')",
                        },
                        "include_headers": {
                            "type": "boolean",
                            "description": "Treat first row as headers (default: false)",
                        },
                        "skip_empty_rows": {
                            "type": "boolean",
                            "description": "Skip empty rows (default: false)",
                        },
                    },
                    "required": ["file_path"],
                },
            ),
            Tool(
                name="write_excel",
                description=(
                    "Write data to an Excel file. Creates a new file or overwrites "
                    "an existing one. Supports headers and auto-formatting."
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path where the Excel file will be written",
                        },
                        "rows": {
                            "type": "array",
                            "items": {
                                "type": "array",
                            },
                            "description": "List of rows to write, where each row is a list of values",
                        },
                        "sheet_name": {
                            "type": "string",
                            "description": "Name of the sheet (default: 'Sheet1')",
                        },
                        "headers": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "Optional list of column headers",
                        },
                        "start_cell": {
                            "type": "string",
                            "description": "Starting cell for data (default: 'A1')",
                        },
                        "overwrite": {
                            "type": "boolean",
                            "description": "Whether to overwrite existing file (default: false)",
                        },
                        "auto_format": {
                            "type": "boolean",
                            "description": "Auto-format column widths (default: true)",
                        },
                    },
                    "required": ["file_path", "rows"],
                },
            ),
        ]

    async def _execute_tool(self, name: str, arguments: dict[str, Any]) -> dict[str, Any]:
        """
        Execute a tool by name with the given arguments.

        Args:
            name: The name of the tool to execute.
            arguments: The arguments to pass to the tool.

        Returns:
            Dictionary containing the tool execution result.

        Raises:
            ValueError: If the tool name is not recognized.
        """
        try:
            if name == "get_workbook_info":
                result = self.service.get_workbook_info(arguments["file_path"])
                return {"success": True, "data": result.model_dump()}

            elif name == "list_sheets":
                result = self.service.get_sheet_names(arguments["file_path"])
                return {"success": True, "data": {"sheets": result}}

            elif name == "read_sheet":
                result = self.service.read_sheet(
                    file_path=arguments["file_path"],
                    sheet_name=arguments.get("sheet_name"),
                    sheet_index=arguments.get("sheet_index"),
                    skip_empty_rows=arguments.get("skip_empty_rows", False),
                )
                return {"success": True, "data": result.model_dump()}

            elif name == "read_range":
                result = self.service.read_range(
                    file_path=arguments["file_path"],
                    cell_range=arguments["cell_range"],
                    sheet_name=arguments.get("sheet_name"),
                    sheet_index=arguments.get("sheet_index"),
                )
                return {"success": True, "data": result.model_dump()}

            elif name == "read_cell":
                value = self.service.get_cell_value(
                    file_path=arguments["file_path"],
                    cell=arguments["cell"],
                    sheet_name=arguments.get("sheet_name"),
                    sheet_index=arguments.get("sheet_index"),
                )
                return {
                    "success": True,
                    "data": {
                        "cell": arguments["cell"],
                        "value": value,
                        "value_type": type(value).__name__ if value is not None else "null",
                    },
                }

            elif name == "read_excel":
                request = ReadExcelRequest(
                    file_path=arguments["file_path"],
                    sheet_name=arguments.get("sheet_name"),
                    sheet_index=arguments.get("sheet_index"),
                    cell_range=arguments.get("cell_range"),
                    include_headers=arguments.get("include_headers", False),
                    skip_empty_rows=arguments.get("skip_empty_rows", False),
                )
                result = self.service.read_excel(request)
                return {"success": True, "data": result.model_dump()}

            elif name == "write_excel":
                request = WriteExcelRequest(
                    file_path=arguments["file_path"],
                    rows=arguments["rows"],
                    sheet_name=arguments.get("sheet_name", "Sheet1"),
                    headers=arguments.get("headers"),
                    start_cell=arguments.get("start_cell", "A1"),
                    overwrite=arguments.get("overwrite", False),
                    auto_format=arguments.get("auto_format", True),
                )
                result = self.service.write_excel(request)
                return {"success": True, "data": result.model_dump()}

            else:
                return {
                    "success": False,
                    "error": {
                        "error_code": "UNKNOWN_TOOL",
                        "message": f"Unknown tool: {name}",
                    },
                }

        except ExcelServiceError as e:
            return {
                "success": False,
                "error": e.to_dict(),
            }
        except Exception as e:
            return {
                "success": False,
                "error": {
                    "error_code": "INTERNAL_ERROR",
                    "message": str(e),
                },
            }

    async def run(self) -> None:
        """
        Run the MCP server using stdio transport.

        This method starts the server and blocks until it is terminated.
        It uses stdin/stdout for communication with the MCP client.
        """
        async with stdio_server() as (read_stream, write_stream):
            await self.server.run(
                read_stream,
                write_stream,
                self.server.create_initialization_options(),
            )


def run_mcp_server() -> None:
    """
    Run the MCP Excel server.

    This is the entry point for running the MCP server from the command line.
    It creates an MCPExcelServer instance and runs it.

    Example:
        python -m src.mcp_server
    """
    server = MCPExcelServer()
    asyncio.run(server.run())


if __name__ == "__main__":
    run_mcp_server()
