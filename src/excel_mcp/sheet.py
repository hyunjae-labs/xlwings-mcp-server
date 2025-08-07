import logging
from typing import Any, Dict

from .exceptions import SheetError
from .xlwings_impl.sheet_xlw import (
    delete_worksheet_xlw, 
    rename_worksheet_xlw, 
    copy_worksheet_xlw
)
from .xlwings_impl.range_xlw import (
    merge_cells_xlw,
    unmerge_cells_xlw,
    get_merged_cells_xlw,
    copy_range_xlw,
    delete_range_xlw
)

logger = logging.getLogger(__name__)

def copy_sheet(filepath: str, source_sheet: str, target_sheet: str) -> Dict[str, Any]:
    """Copy a worksheet within the same workbook."""
    result = copy_worksheet_xlw(filepath, source_sheet, target_sheet)
    if "error" in result:
        raise SheetError(result["error"])
    return result

def delete_sheet(filepath: str, sheet_name: str) -> Dict[str, Any]:
    """Delete a worksheet from the workbook."""
    result = delete_worksheet_xlw(filepath, sheet_name)
    if "error" in result:
        raise SheetError(result["error"])
    return result

def rename_sheet(filepath: str, old_name: str, new_name: str) -> Dict[str, Any]:
    """Rename a worksheet."""
    result = rename_worksheet_xlw(filepath, old_name, new_name)
    if "error" in result:
        raise SheetError(result["error"])
    return result

def merge_range(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> Dict[str, Any]:
    """Merge a range of cells."""
    result = merge_cells_xlw(filepath, sheet_name, start_cell, end_cell)
    if "error" in result:
        raise SheetError(result["error"])
    return result

def unmerge_range(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> Dict[str, Any]:
    """Unmerge a range of cells."""
    result = unmerge_cells_xlw(filepath, sheet_name, start_cell, end_cell)
    if "error" in result:
        raise SheetError(result["error"])
    return result

def get_merged_ranges(filepath: str, sheet_name: str) -> str:
    """Get merged cells in a worksheet."""
    result = get_merged_cells_xlw(filepath, sheet_name)
    if "error" in result:
        raise SheetError(result["error"])
    return str(result)

def copy_range_operation(
    filepath: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    target_sheet: str
) -> Dict[str, Any]:
    """Copy a range of cells to another location."""
    result = copy_range_xlw(filepath, sheet_name, source_start, source_end, target_start, target_sheet)
    if "error" in result:
        raise SheetError(result["error"])
    return result

def delete_range_operation(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    shift_direction: str
) -> Dict[str, Any]:
    """Delete a range of cells and shift remaining cells."""
    result = delete_range_xlw(filepath, sheet_name, start_cell, end_cell, shift_direction)
    if "error" in result:
        raise SheetError(result["error"])
    return result

# 행/열 관련 함수들은 server.py에서 직접 구현됨