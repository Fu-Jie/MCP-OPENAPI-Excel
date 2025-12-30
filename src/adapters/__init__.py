"""
Adapters for Excel file operations.

Implements the adapter pattern for different Excel engines:
- CalamineAdapter: High-performance reading using python-calamine (Rust-based)
- XlsxWriterAdapter: High-performance streaming writes using XlsxWriter
"""

from src.adapters.calamine_adapter import CalamineAdapter
from src.adapters.xlsxwriter_adapter import XlsxWriterAdapter

__all__ = [
    "CalamineAdapter",
    "XlsxWriterAdapter",
]
