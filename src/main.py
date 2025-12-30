"""
FastAPI application for Excel service.

This module provides the REST API endpoints for Excel operations using FastAPI.
It exposes both file-based operations (for local files) and file upload
operations (for processing uploaded files).

API Endpoints:
    - GET /health: Health check
    - GET /workbook/info: Get workbook metadata
    - GET /workbook/sheets: List sheet names
    - POST /excel/read: Read Excel data
    - POST /excel/write: Write Excel data
    - POST /excel/upload: Upload and read Excel file

Example:
    To run the server:
        uvicorn src.main:app --reload

    Or programmatically:
        from src.main import run_server
        run_server()
"""

import os
import tempfile
from contextlib import asynccontextmanager
from datetime import datetime, timezone
from typing import Annotated, Any

import uvicorn
from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from src.exceptions.excel_exceptions import (
    CellRangeError,
    ExcelServiceError,
    InvalidFileFormatError,
    SheetNotFoundError,
)
from src.exceptions.excel_exceptions import FileNotFoundError as ExcelFileNotFoundError
from src.models.excel_models import (
    ExcelErrorResponse,
    ReadExcelRequest,
    ReadExcelResponse,
    SheetData,
    SheetInfo,
    WorkbookInfo,
    WriteExcelRequest,
    WriteExcelResponse,
)
from src.services.excel_service import ExcelService

excel_service: ExcelService | None = None


@asynccontextmanager
async def lifespan(app: FastAPI):
    """
    Application lifespan manager.

    Initializes the Excel service on startup and cleans up on shutdown.

    Args:
        app: The FastAPI application instance.
    """
    global excel_service
    excel_service = ExcelService()
    yield
    excel_service = None


app = FastAPI(
    title="Excel Dual-Protocol Service",
    description="""
    High-performance Excel service supporting both REST API and MCP (Model Context Protocol).

    ## Features

    - **Read Excel files**: Using python-calamine (Rust-based) for exceptional performance
    - **Write Excel files**: Using XlsxWriter for streaming large file support
    - **Multiple formats**: .xlsx, .xls, .xlsb, .xlsm, .ods
    - **Range operations**: Read specific cell ranges using A1 notation

    ## Architecture

    This service uses a hexagonal architecture pattern with:
    - **Service Layer**: Core business logic decoupled from transport
    - **Adapters**: python-calamine for reading, XlsxWriter for writing
    - **Dual Protocol**: Same service exposed via REST and MCP
    """,
    version="0.1.0",
    lifespan=lifespan,
    docs_url="/docs",
    redoc_url="/redoc",
    openapi_url="/openapi.json",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def get_service() -> ExcelService:
    """
    Get the Excel service instance.

    Returns:
        The global ExcelService instance.

    Raises:
        HTTPException: If the service is not initialized.
    """
    if excel_service is None:
        raise HTTPException(
            status_code=503,
            detail="Excel service is not initialized",
        )
    return excel_service


def handle_excel_error(error: ExcelServiceError) -> JSONResponse:
    """
    Convert ExcelServiceError to appropriate HTTP response.

    Args:
        error: The ExcelServiceError to convert.

    Returns:
        JSONResponse with appropriate status code and error details.
    """
    status_code_map = {
        "FILE_NOT_FOUND": 404,
        "INVALID_FILE_FORMAT": 400,
        "SHEET_NOT_FOUND": 404,
        "INVALID_CELL_RANGE": 400,
        "READ_ERROR": 500,
        "WRITE_ERROR": 500,
        "PERMISSION_DENIED": 403,
    }

    status_code = status_code_map.get(error.error_code, 500)

    return JSONResponse(
        status_code=status_code,
        content=ExcelErrorResponse(
            success=False,
            error_code=error.error_code,
            message=error.message,
            details=error.details,
        ).model_dump(),
    )


@app.get(
    "/health",
    tags=["System"],
    summary="Health check",
    response_model=dict,
)
async def health_check() -> dict[str, Any]:
    """
    Check the health status of the service.

    Returns:
        Dictionary containing status and timestamp.
    """
    return {
        "status": "healthy",
        "service": "Excel Dual-Protocol Service",
        "version": "0.1.0",
        "timestamp": datetime.now(timezone.utc).isoformat(),
    }


@app.get(
    "/workbook/info",
    tags=["Workbook"],
    summary="Get workbook information",
    response_model=WorkbookInfo,
    responses={
        404: {"model": ExcelErrorResponse, "description": "File not found"},
        400: {"model": ExcelErrorResponse, "description": "Invalid file format"},
    },
)
async def get_workbook_info(
    file_path: Annotated[str, Query(description="Path to the Excel file")],
) -> WorkbookInfo:
    """
    Get metadata about an Excel workbook.

    This endpoint retrieves comprehensive information about the workbook including:
    - File size
    - Number of sheets
    - Sheet names and metadata
    - File timestamps

    Args:
        file_path: Path to the Excel file on the server.

    Returns:
        WorkbookInfo containing workbook metadata.

    Raises:
        HTTPException: If the file is not found or invalid.
    """
    service = get_service()

    try:
        return service.get_workbook_info(file_path)
    except ExcelServiceError as e:
        raise HTTPException(
            status_code=404 if isinstance(e, ExcelFileNotFoundError) else 400,
            detail=e.to_dict(),
        ) from e


@app.get(
    "/workbook/sheets",
    tags=["Workbook"],
    summary="List sheet names",
    response_model=list[str],
    responses={
        404: {"model": ExcelErrorResponse, "description": "File not found"},
        400: {"model": ExcelErrorResponse, "description": "Invalid file format"},
    },
)
async def get_sheet_names(
    file_path: Annotated[str, Query(description="Path to the Excel file")],
) -> list[str]:
    """
    Get the list of sheet names in an Excel workbook.

    Args:
        file_path: Path to the Excel file on the server.

    Returns:
        List of sheet names.

    Raises:
        HTTPException: If the file is not found or invalid.
    """
    service = get_service()

    try:
        return service.get_sheet_names(file_path)
    except ExcelServiceError as e:
        raise HTTPException(
            status_code=404 if isinstance(e, ExcelFileNotFoundError) else 400,
            detail=e.to_dict(),
        ) from e


@app.get(
    "/workbook/sheet",
    tags=["Workbook"],
    summary="Get sheet information",
    response_model=SheetInfo,
    responses={
        404: {"model": ExcelErrorResponse, "description": "File or sheet not found"},
        400: {"model": ExcelErrorResponse, "description": "Invalid file format"},
    },
)
async def get_sheet_info(
    file_path: Annotated[str, Query(description="Path to the Excel file")],
    sheet_name: Annotated[str | None, Query(description="Sheet name")] = None,
    sheet_index: Annotated[int | None, Query(description="Sheet index (0-based)")] = None,
) -> SheetInfo:
    """
    Get metadata about a specific sheet.

    Args:
        file_path: Path to the Excel file on the server.
        sheet_name: Name of the sheet. If None, uses sheet_index or first sheet.
        sheet_index: Index of the sheet (0-based). Used if sheet_name is None.

    Returns:
        SheetInfo containing sheet metadata.

    Raises:
        HTTPException: If the file or sheet is not found.
    """
    service = get_service()

    try:
        return service.get_sheet_info(
            file_path=file_path,
            sheet_name=sheet_name,
            sheet_index=sheet_index,
        )
    except ExcelServiceError as e:
        status_code = 404 if isinstance(e, (ExcelFileNotFoundError, SheetNotFoundError)) else 400
        raise HTTPException(status_code=status_code, detail=e.to_dict()) from e


@app.post(
    "/excel/read",
    tags=["Excel Operations"],
    summary="Read Excel data",
    response_model=ReadExcelResponse,
    responses={
        404: {"model": ExcelErrorResponse, "description": "File or sheet not found"},
        400: {"model": ExcelErrorResponse, "description": "Invalid request"},
    },
)
async def read_excel(request: ReadExcelRequest) -> ReadExcelResponse:
    """
    Read data from an Excel file.

    This endpoint supports reading entire sheets or specific cell ranges.
    Uses python-calamine for high-performance reading, especially optimized
    for large files (100MB+).

    Args:
        request: ReadExcelRequest containing file path and read options.

    Returns:
        ReadExcelResponse containing workbook info and sheet data.

    Raises:
        HTTPException: If reading fails.
    """
    service = get_service()

    try:
        return service.read_excel(request)
    except ExcelServiceError as e:
        status_code = 404 if isinstance(e, (ExcelFileNotFoundError, SheetNotFoundError)) else 400
        raise HTTPException(status_code=status_code, detail=e.to_dict()) from e


@app.get(
    "/excel/read/sheet",
    tags=["Excel Operations"],
    summary="Read sheet data",
    response_model=SheetData,
    responses={
        404: {"model": ExcelErrorResponse, "description": "File or sheet not found"},
        400: {"model": ExcelErrorResponse, "description": "Invalid request"},
    },
)
async def read_sheet(
    file_path: Annotated[str, Query(description="Path to the Excel file")],
    sheet_name: Annotated[str | None, Query(description="Sheet name")] = None,
    sheet_index: Annotated[int | None, Query(description="Sheet index (0-based)")] = None,
    skip_empty_rows: Annotated[bool, Query(description="Skip empty rows")] = False,
) -> SheetData:
    """
    Read data from a specific sheet.

    Simpler endpoint for reading a single sheet without full workbook metadata.

    Args:
        file_path: Path to the Excel file.
        sheet_name: Name of the sheet. If None, uses sheet_index or first sheet.
        sheet_index: Index of the sheet (0-based). Used if sheet_name is None.
        skip_empty_rows: Whether to skip empty rows.

    Returns:
        SheetData containing the sheet contents.

    Raises:
        HTTPException: If reading fails.
    """
    service = get_service()

    try:
        return service.read_sheet(
            file_path=file_path,
            sheet_name=sheet_name,
            sheet_index=sheet_index,
            skip_empty_rows=skip_empty_rows,
        )
    except ExcelServiceError as e:
        status_code = 404 if isinstance(e, (ExcelFileNotFoundError, SheetNotFoundError)) else 400
        raise HTTPException(status_code=status_code, detail=e.to_dict()) from e


@app.get(
    "/excel/read/range",
    tags=["Excel Operations"],
    summary="Read cell range",
    response_model=SheetData,
    responses={
        404: {"model": ExcelErrorResponse, "description": "File or sheet not found"},
        400: {"model": ExcelErrorResponse, "description": "Invalid request or range"},
    },
)
async def read_range(
    file_path: Annotated[str, Query(description="Path to the Excel file")],
    cell_range: Annotated[str, Query(description="Cell range in A1 notation (e.g., 'A1:C10')")],
    sheet_name: Annotated[str | None, Query(description="Sheet name")] = None,
    sheet_index: Annotated[int | None, Query(description="Sheet index (0-based)")] = None,
) -> SheetData:
    """
    Read a specific cell range from a sheet.

    Args:
        file_path: Path to the Excel file.
        cell_range: Cell range in A1 notation (e.g., "A1:C10").
        sheet_name: Name of the sheet. If None, uses sheet_index or first sheet.
        sheet_index: Index of the sheet (0-based). Used if sheet_name is None.

    Returns:
        SheetData containing the range contents.

    Raises:
        HTTPException: If reading fails or range is invalid.
    """
    service = get_service()

    try:
        return service.read_range(
            file_path=file_path,
            cell_range=cell_range,
            sheet_name=sheet_name,
            sheet_index=sheet_index,
        )
    except ExcelServiceError as e:
        status_code = 404 if isinstance(e, (ExcelFileNotFoundError, SheetNotFoundError)) else 400
        if isinstance(e, CellRangeError):
            status_code = 400
        raise HTTPException(status_code=status_code, detail=e.to_dict()) from e


@app.get(
    "/excel/read/cell",
    tags=["Excel Operations"],
    summary="Read single cell value",
    responses={
        404: {"model": ExcelErrorResponse, "description": "File or sheet not found"},
        400: {"model": ExcelErrorResponse, "description": "Invalid cell reference"},
    },
)
async def read_cell(
    file_path: Annotated[str, Query(description="Path to the Excel file")],
    cell: Annotated[str, Query(description="Cell reference in A1 notation (e.g., 'A1')")],
    sheet_name: Annotated[str | None, Query(description="Sheet name")] = None,
    sheet_index: Annotated[int | None, Query(description="Sheet index (0-based)")] = None,
) -> dict[str, Any]:
    """
    Get the value of a single cell.

    Args:
        file_path: Path to the Excel file.
        cell: Cell reference in A1 notation (e.g., "A1").
        sheet_name: Name of the sheet. If None, uses sheet_index or first sheet.
        sheet_index: Index of the sheet (0-based). Used if sheet_name is None.

    Returns:
        Dictionary containing the cell value and metadata.

    Raises:
        HTTPException: If reading fails or cell reference is invalid.
    """
    service = get_service()

    try:
        value = service.get_cell_value(
            file_path=file_path,
            cell=cell,
            sheet_name=sheet_name,
            sheet_index=sheet_index,
        )

        value_type = type(value).__name__ if value is not None else "null"

        return {
            "cell": cell,
            "value": value,
            "value_type": value_type,
        }
    except ExcelServiceError as e:
        status_code = 404 if isinstance(e, (ExcelFileNotFoundError, SheetNotFoundError)) else 400
        raise HTTPException(status_code=status_code, detail=e.to_dict()) from e


@app.post(
    "/excel/write",
    tags=["Excel Operations"],
    summary="Write Excel data",
    response_model=WriteExcelResponse,
    responses={
        400: {"model": ExcelErrorResponse, "description": "Invalid request"},
        403: {"model": ExcelErrorResponse, "description": "Permission denied"},
        500: {"model": ExcelErrorResponse, "description": "Write error"},
    },
)
async def write_excel(request: WriteExcelRequest) -> WriteExcelResponse:
    """
    Write data to an Excel file.

    Creates a new Excel file with the specified data. Uses XlsxWriter
    for high-performance streaming writes, optimal for large datasets.

    Args:
        request: WriteExcelRequest containing file path, data, and options.

    Returns:
        WriteExcelResponse containing write results.

    Raises:
        HTTPException: If writing fails.
    """
    service = get_service()

    try:
        return service.write_excel(request)
    except ExcelServiceError as e:
        status_code_map = {
            "WRITE_ERROR": 500,
            "PERMISSION_DENIED": 403,
        }
        status_code = status_code_map.get(e.error_code, 400)
        raise HTTPException(status_code=status_code, detail=e.to_dict()) from e


@app.post(
    "/excel/upload",
    tags=["Excel Operations"],
    summary="Upload and read Excel file",
    response_model=ReadExcelResponse,
    responses={
        400: {"model": ExcelErrorResponse, "description": "Invalid file"},
        500: {"model": ExcelErrorResponse, "description": "Processing error"},
    },
)
async def upload_and_read_excel(
    file: Annotated[UploadFile, File(description="Excel file to upload")],
    sheet_name: Annotated[str | None, Query(description="Sheet name")] = None,
    sheet_index: Annotated[int | None, Query(description="Sheet index (0-based)")] = None,
    cell_range: Annotated[str | None, Query(description="Cell range (e.g., 'A1:C10')")] = None,
    include_headers: Annotated[bool, Query(description="Treat first row as headers")] = False,
    skip_empty_rows: Annotated[bool, Query(description="Skip empty rows")] = False,
) -> ReadExcelResponse:
    """
    Upload an Excel file and read its contents.

    This endpoint handles file uploads for scenarios where the Excel file
    is not stored on the server. The file is temporarily saved and processed.

    Args:
        file: The uploaded Excel file.
        sheet_name: Name of the sheet to read.
        sheet_index: Index of the sheet (0-based).
        cell_range: Optional cell range to read.
        include_headers: Whether to treat the first row as headers.
        skip_empty_rows: Whether to skip empty rows.

    Returns:
        ReadExcelResponse containing workbook info and sheet data.

    Raises:
        HTTPException: If upload or reading fails.
    """
    if not file.filename:
        raise HTTPException(
            status_code=400,
            detail={"error_code": "INVALID_FILE", "message": "No file provided"},
        )

    valid_extensions = (".xlsx", ".xls", ".xlsb", ".xlsm", ".ods")
    if not file.filename.lower().endswith(valid_extensions):
        raise HTTPException(
            status_code=400,
            detail={
                "error_code": "INVALID_FILE_FORMAT",
                "message": f"Invalid file extension. Supported: {', '.join(valid_extensions)}",
            },
        )

    service = get_service()
    temp_path = None

    try:
        suffix = os.path.splitext(file.filename)[1]
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
            content = await file.read()
            temp_file.write(content)
            temp_path = temp_file.name

        request = ReadExcelRequest(
            file_path=temp_path,
            sheet_name=sheet_name,
            sheet_index=sheet_index,
            cell_range=cell_range,
            include_headers=include_headers,
            skip_empty_rows=skip_empty_rows,
        )

        response = service.read_excel(request)

        response.workbook_info.file_path = file.filename

        return response

    except ExcelServiceError as e:
        raise HTTPException(
            status_code=400 if isinstance(e, InvalidFileFormatError) else 500,
            detail=e.to_dict(),
        ) from e
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail={"error_code": "PROCESSING_ERROR", "message": str(e)},
        ) from e
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.unlink(temp_path)
            except Exception:
                pass


def run_server(
    host: str = "0.0.0.0",
    port: int = 8000,
    reload: bool = False,
) -> None:
    """
    Run the FastAPI server.

    Args:
        host: Host to bind to. Defaults to "0.0.0.0".
        port: Port to listen on. Defaults to 8000.
        reload: Whether to enable auto-reload. Defaults to False.

    Example:
        from src.main import run_server
        run_server(host="127.0.0.1", port=8080)
    """
    uvicorn.run(
        "src.main:app",
        host=host,
        port=port,
        reload=reload,
    )


if __name__ == "__main__":
    run_server()
