"""
xlwings implementation for range operations.
Focuses on batch processing efficiency and native Excel capabilities.
"""

import xlwings as xw
from typing import List, Dict, Any, Optional, Tuple
import logging
import os

logger = logging.getLogger(__name__)

def merge_cells_xlw(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> Dict[str, Any]:
    """
    Merge cells in Excel using xlwings.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        start_cell: Top-left cell of merge range
        end_cell: Bottom-right cell of merge range
        
    Returns:
        Dict with success message or error
    """
    app = None
    wb = None
    
    try:
        logger.info(f"🔗 Merging cells {start_cell}:{end_cell} in {sheet_name}")
        
        # Check if file exists
        if not os.path.exists(filepath):
            return {"error": f"File not found: {filepath}"}
        
        # Open Excel app and workbook
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(filepath)
        
        # Check if sheet exists
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        sheet = wb.sheets[sheet_name]
        
        # Get the range to merge
        merge_range = sheet.range(f"{start_cell}:{end_cell}")
        
        # Check if range is already merged
        if merge_range.merge_cells:
            return {"error": f"Range {start_cell}:{end_cell} is already merged"}
        
        # Merge the cells
        merge_range.merge()
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully merged cells {start_cell}:{end_cell}")
        return {
            "message": f"Successfully merged cells {start_cell}:{end_cell}",
            "range": f"{start_cell}:{end_cell}",
            "sheet": sheet_name
        }
        
    except Exception as e:
        logger.error(f"❌ Error merging cells: {str(e)}")
        return {"error": str(e)}
        
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()


def unmerge_cells_xlw(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> Dict[str, Any]:
    """
    Unmerge cells in Excel using xlwings.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        start_cell: Top-left cell of merge range
        end_cell: Bottom-right cell of merge range
        
    Returns:
        Dict with success message or error
    """
    app = None
    wb = None
    
    try:
        logger.info(f"🔓 Unmerging cells {start_cell}:{end_cell} in {sheet_name}")
        
        # Check if file exists
        if not os.path.exists(filepath):
            return {"error": f"File not found: {filepath}"}
        
        # Open Excel app and workbook
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(filepath)
        
        # Check if sheet exists
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        sheet = wb.sheets[sheet_name]
        
        # Get the range to unmerge
        unmerge_range = sheet.range(f"{start_cell}:{end_cell}")
        
        # Check if range is merged
        if not unmerge_range.merge_cells:
            return {"error": f"Range {start_cell}:{end_cell} is not merged"}
        
        # Unmerge the cells
        unmerge_range.unmerge()
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully unmerged cells {start_cell}:{end_cell}")
        return {
            "message": f"Successfully unmerged cells {start_cell}:{end_cell}",
            "range": f"{start_cell}:{end_cell}",
            "sheet": sheet_name
        }
        
    except Exception as e:
        logger.error(f"❌ Error unmerging cells: {str(e)}")
        return {"error": str(e)}
        
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()


def get_merged_cells_xlw(filepath: str, sheet_name: str) -> Dict[str, Any]:
    """
    Get all merged cell ranges in a worksheet using xlwings.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        
    Returns:
        Dict with list of merged ranges or error
    """
    app = None
    wb = None
    
    try:
        logger.info(f"📊 Getting merged cells in {sheet_name}")
        
        # Check if file exists
        if not os.path.exists(filepath):
            return {"error": f"File not found: {filepath}"}
        
        # Open Excel app and workbook
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(filepath)
        
        # Check if sheet exists
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        sheet = wb.sheets[sheet_name]
        
        # Get all merged ranges
        merged_ranges = []
        
        # Use a simpler approach with COM API
        try:
            # Access the worksheet's COM object directly
            ws_com = sheet.api
            
            # Check if there are any merged cells in the sheet
            used_range = sheet.used_range
            if used_range:
                # Track processed merged areas to avoid duplicates
                processed_areas = set()
                
                # Get dimensions of used range
                max_row = used_range.last_cell.row
                max_col = used_range.last_cell.column
                
                # Iterate through the used range
                for row in range(1, max_row + 1):
                    for col in range(1, max_col + 1):
                        try:
                            # Get the cell using COM API
                            cell_com = ws_com.Cells(row, col)
                            
                            # Check if this cell is part of a merged range
                            if cell_com.MergeCells:
                                # Get the merge area
                                merge_area = cell_com.MergeArea
                                merge_address = merge_area.Address.replace("$", "")
                                
                                # Skip if we've already processed this merged area
                                if merge_address in processed_areas:
                                    continue
                                
                                # Add to processed areas
                                processed_areas.add(merge_address)
                                
                                # Get details about the merged range
                                first_row = merge_area.Row
                                first_col = merge_area.Column
                                row_count = merge_area.Rows.Count
                                col_count = merge_area.Columns.Count
                                
                                # Calculate last cell
                                last_row = first_row + row_count - 1
                                last_col = first_col + col_count - 1
                                
                                # Create cell addresses
                                def get_column_letter(col_idx):
                                    """Convert column index to Excel column string"""
                                    result = ""
                                    while col_idx > 0:
                                        col_idx -= 1
                                        result = chr(col_idx % 26 + ord('A')) + result
                                        col_idx //= 26
                                    return result
                                
                                start_addr = f"{get_column_letter(first_col)}{first_row}"
                                end_addr = f"{get_column_letter(last_col)}{last_row}"
                                
                                merged_ranges.append({
                                    "range": merge_address,
                                    "start": start_addr,
                                    "end": end_addr,
                                    "rows": row_count,
                                    "columns": col_count
                                })
                        except:
                            # Skip cells that cause errors
                            continue
        except Exception as e:
            logger.warning(f"Could not get merged cells: {e}")
        
        logger.info(f"✅ Found {len(merged_ranges)} merged ranges")
        return {
            "merged_ranges": merged_ranges,
            "count": len(merged_ranges),
            "sheet": sheet_name
        }
        
    except Exception as e:
        logger.error(f"❌ Error getting merged cells: {str(e)}")
        return {"error": str(e)}
        
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()


def copy_range_xlw(
    filepath: str, 
    sheet_name: str, 
    source_start: str, 
    source_end: str, 
    target_start: str,
    target_sheet: Optional[str] = None
) -> Dict[str, Any]:
    """
    Copy a range of cells to another location using xlwings.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of source worksheet
        source_start: Top-left cell of source range
        source_end: Bottom-right cell of source range
        target_start: Top-left cell of target location
        target_sheet: Name of target worksheet (optional, defaults to source sheet)
        
    Returns:
        Dict with success message or error
    """
    app = None
    wb = None
    
    try:
        # Use target_sheet if provided, otherwise use source sheet
        target_sheet = target_sheet or sheet_name
        
        logger.info(f"📋 Copying range {source_start}:{source_end} to {target_start} in {target_sheet}")
        
        # Check if file exists
        if not os.path.exists(filepath):
            return {"error": f"File not found: {filepath}"}
        
        # Open Excel app and workbook
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(filepath)
        
        # Check if sheets exist
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Source sheet '{sheet_name}' not found"}
        if target_sheet not in sheet_names:
            return {"error": f"Target sheet '{target_sheet}' not found"}
        
        source_sheet = wb.sheets[sheet_name]
        dest_sheet = wb.sheets[target_sheet]
        
        # Get source range
        source_range = source_sheet.range(f"{source_start}:{source_end}")
        
        # Copy to target location
        # xlwings copy method preserves formatting and formulas
        source_range.copy(destination=dest_sheet.range(target_start))
        
        # Calculate target end cell
        rows = source_range.rows.count
        cols = source_range.columns.count
        target_end_row = dest_sheet.range(target_start).row + rows - 1
        target_end_col = dest_sheet.range(target_start).column + cols - 1
        target_end = dest_sheet.cells(target_end_row, target_end_col).address.replace("$", "")
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully copied range to {target_start}:{target_end}")
        return {
            "message": f"Successfully copied range {source_start}:{source_end} to {target_start}",
            "source_range": f"{source_start}:{source_end}",
            "target_range": f"{target_start}:{target_end}",
            "source_sheet": sheet_name,
            "target_sheet": target_sheet
        }
        
    except Exception as e:
        logger.error(f"❌ Error copying range: {str(e)}")
        return {"error": str(e)}
        
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()


def delete_range_xlw(
    filepath: str, 
    sheet_name: str, 
    start_cell: str, 
    end_cell: str,
    shift_direction: str = "up"
) -> Dict[str, Any]:
    """
    Delete a range of cells and shift remaining cells using xlwings.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        start_cell: Top-left cell of range to delete
        end_cell: Bottom-right cell of range to delete
        shift_direction: Direction to shift cells ("up" or "left")
        
    Returns:
        Dict with success message or error
    """
    app = None
    wb = None
    
    try:
        logger.info(f"🗑️ Deleting range {start_cell}:{end_cell} in {sheet_name}, shift {shift_direction}")
        
        # Validate shift direction
        if shift_direction not in ["up", "left"]:
            return {"error": f"Invalid shift direction: {shift_direction}. Must be 'up' or 'left'"}
        
        # Check if file exists
        if not os.path.exists(filepath):
            return {"error": f"File not found: {filepath}"}
        
        # Open Excel app and workbook
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(filepath)
        
        # Check if sheet exists
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        sheet = wb.sheets[sheet_name]
        
        # Get the range to delete
        delete_range = sheet.range(f"{start_cell}:{end_cell}")
        
        # Delete and shift cells
        if shift_direction == "up":
            # Shift cells up (xlShiftUp = -4162)
            delete_range.api.Delete(Shift=-4162)
        else:  # shift_direction == "left"
            # Shift cells left (xlShiftToLeft = -4159)
            delete_range.api.Delete(Shift=-4159)
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully deleted range {start_cell}:{end_cell}")
        return {
            "message": f"Successfully deleted range {start_cell}:{end_cell} and shifted cells {shift_direction}",
            "deleted_range": f"{start_cell}:{end_cell}",
            "shift_direction": shift_direction,
            "sheet": sheet_name
        }
        
    except Exception as e:
        logger.error(f"❌ Error deleting range: {str(e)}")
        return {"error": str(e)}
        
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()


# Batch operation helper for efficiency
def batch_range_operations_xlw(
    filepath: str,
    operations: List[Dict[str, Any]]
) -> Dict[str, Any]:
    """
    Execute multiple range operations in a single Excel session for efficiency.
    
    Args:
        filepath: Path to Excel file
        operations: List of operations to perform, each with 'type' and parameters
        
    Returns:
        Dict with results of all operations
    """
    app = None
    wb = None
    results = []
    
    try:
        logger.info(f"⚡ Executing {len(operations)} batch operations")
        
        # Check if file exists
        if not os.path.exists(filepath):
            return {"error": f"File not found: {filepath}"}
        
        # Open Excel app and workbook once
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(filepath)
        
        # Execute each operation
        for idx, op in enumerate(operations):
            op_type = op.get("type")
            result = {"operation": idx + 1, "type": op_type}
            
            try:
                if op_type == "merge":
                    sheet = wb.sheets[op["sheet_name"]]
                    merge_range = sheet.range(f"{op['start_cell']}:{op['end_cell']}")
                    merge_range.merge()
                    result["status"] = "success"
                    result["message"] = f"Merged {op['start_cell']}:{op['end_cell']}"
                    
                elif op_type == "unmerge":
                    sheet = wb.sheets[op["sheet_name"]]
                    unmerge_range = sheet.range(f"{op['start_cell']}:{op['end_cell']}")
                    unmerge_range.unmerge()
                    result["status"] = "success"
                    result["message"] = f"Unmerged {op['start_cell']}:{op['end_cell']}"
                    
                elif op_type == "copy":
                    source_sheet = wb.sheets[op["source_sheet"]]
                    target_sheet = wb.sheets[op.get("target_sheet", op["source_sheet"])]
                    source_range = source_sheet.range(f"{op['source_start']}:{op['source_end']}")
                    source_range.copy(destination=target_sheet.range(op["target_start"]))
                    result["status"] = "success"
                    result["message"] = f"Copied range to {op['target_start']}"
                    
                elif op_type == "delete":
                    sheet = wb.sheets[op["sheet_name"]]
                    delete_range = sheet.range(f"{op['start_cell']}:{op['end_cell']}")
                    shift = -4162 if op.get("shift_direction", "up") == "up" else -4159
                    delete_range.api.Delete(Shift=shift)
                    result["status"] = "success"
                    result["message"] = f"Deleted {op['start_cell']}:{op['end_cell']}"
                    
                else:
                    result["status"] = "error"
                    result["message"] = f"Unknown operation type: {op_type}"
                    
            except Exception as e:
                result["status"] = "error"
                result["message"] = str(e)
            
            results.append(result)
        
        # Save once after all operations
        wb.save()
        
        # Count successes and failures
        successes = sum(1 for r in results if r["status"] == "success")
        failures = sum(1 for r in results if r["status"] == "error")
        
        logger.info(f"✅ Batch operations complete: {successes} succeeded, {failures} failed")
        return {
            "total_operations": len(operations),
            "successes": successes,
            "failures": failures,
            "results": results
        }
        
    except Exception as e:
        logger.error(f"❌ Error in batch operations: {str(e)}")
        return {"error": str(e)}
        
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()

def merge_cells_xlw_with_wb(wb, sheet_name: str, start_cell: str, end_cell: str) -> Dict[str, Any]:
    """
    Session-based cell merging using existing workbook object.
    
    Args:
        wb: Workbook object from session
        sheet_name: Name of worksheet
        start_cell: Top-left cell of merge range
        end_cell: Bottom-right cell of merge range
        
    Returns:
        Dict with success message or error
    """
    try:
        logger.info(f"🔗 Merging cells {start_cell}:{end_cell} in {sheet_name}")
        
        # Check if sheet exists
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        sheet = wb.sheets[sheet_name]
        
        # Get the range to merge
        merge_range = sheet.range(f"{start_cell}:{end_cell}")
        
        # Check if range is already merged
        if merge_range.merge_cells:
            return {"error": f"Range {start_cell}:{end_cell} is already merged"}
        
        # Merge the cells
        merge_range.merge()
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully merged cells {start_cell}:{end_cell}")
        return {
            "message": f"Successfully merged cells {start_cell}:{end_cell}",
            "range": f"{start_cell}:{end_cell}",
            "sheet": sheet_name
        }
        
    except Exception as e:
        logger.error(f"❌ Error merging cells: {str(e)}")
        return {"error": str(e)}

def unmerge_cells_xlw_with_wb(wb, sheet_name: str, start_cell: str, end_cell: str) -> Dict[str, Any]:
    """
    Session-based cell unmerging using existing workbook object.
    
    Args:
        wb: Workbook object from session
        sheet_name: Name of worksheet
        start_cell: Top-left cell of merge range
        end_cell: Bottom-right cell of merge range
        
    Returns:
        Dict with success message or error
    """
    try:
        logger.info(f"🔓 Unmerging cells {start_cell}:{end_cell} in {sheet_name}")
        
        # Check if sheet exists
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        sheet = wb.sheets[sheet_name]
        
        # Get the range to unmerge
        unmerge_range = sheet.range(f"{start_cell}:{end_cell}")
        
        # Check if range is merged
        if not unmerge_range.merge_cells:
            return {"error": f"Range {start_cell}:{end_cell} is not merged"}
        
        # Unmerge the cells
        unmerge_range.unmerge()
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully unmerged cells {start_cell}:{end_cell}")
        return {
            "message": f"Successfully unmerged cells {start_cell}:{end_cell}",
            "range": f"{start_cell}:{end_cell}",
            "sheet": sheet_name
        }
        
    except Exception as e:
        logger.error(f"❌ Error unmerging cells: {str(e)}")
        return {"error": str(e)}

def get_merged_cells_xlw_with_wb(wb, sheet_name: str) -> Dict[str, Any]:
    """
    Session-based merged cells retrieval using existing workbook object.
    
    Args:
        wb: Workbook object from session
        sheet_name: Name of worksheet
        
    Returns:
        Dict with list of merged ranges or error
    """
    try:
        logger.info(f"📊 Getting merged cells in {sheet_name}")
        
        # Check if sheet exists
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        sheet = wb.sheets[sheet_name]
        
        # Get all merged ranges
        merged_ranges = []
        
        # Use a simpler approach with COM API
        try:
            # Access the worksheet's COM object directly
            ws_com = sheet.api
            
            # Check if there are any merged cells in the sheet
            used_range = sheet.used_range
            if used_range:
                # Track processed merged areas to avoid duplicates
                processed_areas = set()
                
                # Get dimensions of used range
                max_row = used_range.last_cell.row
                max_col = used_range.last_cell.column
                
                # Iterate through the used range
                for row in range(1, max_row + 1):
                    for col in range(1, max_col + 1):
                        try:
                            # Get the cell using COM API
                            cell_com = ws_com.Cells(row, col)
                            
                            # Check if this cell is part of a merged range
                            if cell_com.MergeCells:
                                # Get the merge area
                                merge_area = cell_com.MergeArea
                                merge_address = merge_area.Address.replace("$", "")
                                
                                # Skip if we've already processed this merged area
                                if merge_address in processed_areas:
                                    continue
                                
                                # Add to processed areas
                                processed_areas.add(merge_address)
                                
                                # Get details about the merged range
                                first_row = merge_area.Row
                                first_col = merge_area.Column
                                row_count = merge_area.Rows.Count
                                col_count = merge_area.Columns.Count
                                
                                # Calculate last cell
                                last_row = first_row + row_count - 1
                                last_col = first_col + col_count - 1
                                
                                # Create cell addresses
                                def get_column_letter(col_idx):
                                    """Convert column index to Excel column string"""
                                    result = ""
                                    while col_idx > 0:
                                        col_idx -= 1
                                        result = chr(col_idx % 26 + ord('A')) + result
                                        col_idx //= 26
                                    return result
                                
                                start_addr = f"{get_column_letter(first_col)}{first_row}"
                                end_addr = f"{get_column_letter(last_col)}{last_row}"
                                
                                merged_ranges.append({
                                    "range": merge_address,
                                    "start": start_addr,
                                    "end": end_addr,
                                    "rows": row_count,
                                    "columns": col_count
                                })
                        except:
                            # Skip cells that cause errors
                            continue
        except Exception as e:
            logger.warning(f"Could not get merged cells: {e}")
        
        logger.info(f"✅ Found {len(merged_ranges)} merged ranges")
        return {
            "merged_ranges": merged_ranges,
            "count": len(merged_ranges),
            "sheet": sheet_name
        }
        
    except Exception as e:
        logger.error(f"❌ Error getting merged cells: {str(e)}")
        return {"error": str(e)}

def copy_range_xlw_with_wb(
    wb,
    sheet_name: str, 
    source_start: str, 
    source_end: str, 
    target_start: str,
    target_sheet: Optional[str] = None
) -> Dict[str, Any]:
    """
    Session-based range copying using existing workbook object.
    
    Args:
        wb: Workbook object from session
        sheet_name: Name of source worksheet
        source_start: Top-left cell of source range
        source_end: Bottom-right cell of source range
        target_start: Top-left cell of target location
        target_sheet: Name of target worksheet (optional, defaults to source sheet)
        
    Returns:
        Dict with success message or error
    """
    try:
        # Use target_sheet if provided, otherwise use source sheet
        target_sheet = target_sheet or sheet_name
        
        logger.info(f"📋 Copying range {source_start}:{source_end} to {target_start} in {target_sheet}")
        
        # Check if sheets exist
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Source sheet '{sheet_name}' not found"}
        if target_sheet not in sheet_names:
            return {"error": f"Target sheet '{target_sheet}' not found"}
        
        source_sheet = wb.sheets[sheet_name]
        dest_sheet = wb.sheets[target_sheet]
        
        # Get source range
        source_range = source_sheet.range(f"{source_start}:{source_end}")
        
        # Copy to target location
        # xlwings copy method preserves formatting and formulas
        source_range.copy(destination=dest_sheet.range(target_start))
        
        # Calculate target end cell
        rows = source_range.rows.count
        cols = source_range.columns.count
        target_end_row = dest_sheet.range(target_start).row + rows - 1
        target_end_col = dest_sheet.range(target_start).column + cols - 1
        target_end = dest_sheet.cells(target_end_row, target_end_col).address.replace("$", "")
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully copied range to {target_start}:{target_end}")
        return {
            "message": f"Successfully copied range {source_start}:{source_end} to {target_start}",
            "source_range": f"{source_start}:{source_end}",
            "target_range": f"{target_start}:{target_end}",
            "source_sheet": sheet_name,
            "target_sheet": target_sheet
        }
        
    except Exception as e:
        logger.error(f"❌ Error copying range: {str(e)}")
        return {"error": str(e)}

def delete_range_xlw_with_wb(
    wb,
    sheet_name: str, 
    start_cell: str, 
    end_cell: str,
    shift_direction: str = "up"
) -> Dict[str, Any]:
    """
    Session-based range deletion using existing workbook object.
    
    Args:
        wb: Workbook object from session
        sheet_name: Name of worksheet
        start_cell: Top-left cell of range to delete
        end_cell: Bottom-right cell of range to delete
        shift_direction: Direction to shift cells ("up" or "left")
        
    Returns:
        Dict with success message or error
    """
    try:
        logger.info(f"🗑️ Deleting range {start_cell}:{end_cell} in {sheet_name}, shift {shift_direction}")
        
        # Validate shift direction
        if shift_direction not in ["up", "left"]:
            return {"error": f"Invalid shift direction: {shift_direction}. Must be 'up' or 'left'"}
        
        # Check if sheet exists
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        sheet = wb.sheets[sheet_name]
        
        # Get the range to delete
        delete_range = sheet.range(f"{start_cell}:{end_cell}")
        
        # Delete and shift cells
        if shift_direction == "up":
            # Shift cells up (xlShiftUp = -4162)
            delete_range.api.Delete(Shift=-4162)
        else:  # shift_direction == "left"
            # Shift cells left (xlShiftToLeft = -4159)
            delete_range.api.Delete(Shift=-4159)
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully deleted range {start_cell}:{end_cell}")
        return {
            "message": f"Successfully deleted range {start_cell}:{end_cell} and shifted cells {shift_direction}",
            "deleted_range": f"{start_cell}:{end_cell}",
            "shift_direction": shift_direction,
            "sheet": sheet_name
        }
        
    except Exception as e:
        logger.error(f"❌ Error deleting range: {str(e)}")
        return {"error": str(e)}

def validate_excel_range_xlw_with_wb(
    wb,
    sheet_name: str,
    start_cell: str,
    end_cell: str = None
) -> Dict[str, Any]:
    """
    Session-based range validation using existing workbook object.
    
    Args:
        wb: Workbook object from session
        sheet_name: Name of worksheet
        start_cell: Starting cell address
        end_cell: Ending cell address (optional)
        
    Returns:
        Dict containing validation result and range information
    """
    try:
        logger.info(f"🔍 Validating range {start_cell}:{end_cell or start_cell} in {sheet_name}")
        
        # Check if sheet exists
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found", "valid": False}
        
        sheet = wb.sheets[sheet_name]
        
        # Validate the range
        try:
            if end_cell:
                range_obj = sheet.range(f"{start_cell}:{end_cell}")
            else:
                range_obj = sheet.range(start_cell)
            
            # Get range information
            range_info = {
                "message": f"Range validation successful: {range_obj.address}",
                "valid": True,
                "range": range_obj.address,
                "start_cell": start_cell,
                "end_cell": end_cell,
                "rows": range_obj.rows.count,
                "columns": range_obj.columns.count,
                "size": range_obj.rows.count * range_obj.columns.count,
                "sheet": sheet_name,
                "has_data": bool(range_obj.value is not None)
            }
            
            # Check if range has any data
            if range_obj.value:
                if isinstance(range_obj.value, (list, tuple)):
                    non_empty_count = sum(1 for row in range_obj.value 
                                        if row and any(cell for cell in (row if isinstance(row, (list, tuple)) else [row]) if cell is not None))
                else:
                    non_empty_count = 1 if range_obj.value is not None else 0
                range_info["non_empty_cells"] = non_empty_count
            else:
                range_info["non_empty_cells"] = 0
            
            logger.info(f"✅ Range validation successful: {range_obj.address}")
            return range_info
            
        except Exception as range_error:
            return {
                "error": f"Invalid range: {range_error}",
                "valid": False,
                "start_cell": start_cell,
                "end_cell": end_cell,
                "sheet": sheet_name
            }
        
    except Exception as e:
        logger.error(f"Error validating range: {e}")
        return {"error": str(e), "valid": False}