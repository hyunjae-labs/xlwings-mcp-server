import logging
from pathlib import Path
from typing import Any

from .exceptions import WorkbookError
from .xlwings_impl.workbook_xlw import create_workbook_xlw, get_workbook_metadata_xlw
from .xlwings_impl.sheet_xlw import create_worksheet_xlw

logger = logging.getLogger(__name__)

# xlwings 전환 완료 - openpyxl fallback 제거

def create_workbook(filepath: str, sheet_name: str = "Sheet1") -> dict[str, Any]:
    """Create a new Excel workbook with optional custom sheet name"""
    result = create_workbook_xlw(filepath, sheet_name)
    if "error" in result:
        raise WorkbookError(result["error"])
    return result

def create_sheet(filepath: str, sheet_name: str) -> dict:
    """Create a new worksheet in the workbook if it doesn't exist."""
    result = create_worksheet_xlw(filepath, sheet_name)
    if "error" in result:
        raise WorkbookError(result["error"])
    return result

def get_workbook_info(filepath: str, include_ranges: bool = False) -> dict[str, Any]:
    """Get metadata about workbook including sheets, ranges, etc."""
    result = get_workbook_metadata_xlw(filepath, include_ranges)
    if "error" in result:
        raise WorkbookError(result["error"])
    return result
