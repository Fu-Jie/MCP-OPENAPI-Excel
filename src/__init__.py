"""
MCP-OPENAPI-Excel: High-performance Excel dual-protocol service.

This package provides Excel file reading and writing capabilities through
both OpenAPI (REST via FastAPI) and MCP (Model Context Protocol) interfaces.

Architecture:
    - Hexagonal/Service Layer pattern for clean separation of concerns
    - python-calamine for high-performance reading (Rust-based)
    - XlsxWriter for high-performance streaming writes
"""

__version__ = "0.1.0"
__author__ = "Jeff"
