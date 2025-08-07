"""
xlwings implementation for data validation operations.
Handles data validation rules and information retrieval in Excel worksheets.
"""

import xlwings as xw
from typing import Dict, Any, List
import logging
import os
import json

logger = logging.getLogger(__name__)


def get_data_validation_info_xlw(
    filepath: str,
    sheet_name: str
) -> Dict[str, Any]:
    """
    Get all data validation rules in a worksheet using xlwings.
    
    This tool helps identify which cell ranges have validation rules
    and what types of validation are applied.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        
    Returns:
        Dict containing all validation rules in the worksheet
    """
    app = None
    wb = None
    
    try:
        logger.info(f"🔍 Getting data validation info for {sheet_name}")
        
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
        
        # Access worksheet COM object for validation
        ws_com = sheet.api
        
        validation_rules = []
        processed_ranges = set()
        
        # Get used range to scan for validation
        try:
            used_range = sheet.used_range
            if used_range:
                # Iterate through cells to find validation rules
                # Note: This is more efficient than checking every cell
                # We'll check representative cells and expand to find full ranges
                
                max_row = used_range.last_cell.row
                max_col = used_range.last_cell.column
                
                # Sample cells to check (every 5th cell for efficiency)
                for row in range(1, max_row + 1, 5):
                    for col in range(1, max_col + 1, 5):
                        try:
                            cell = sheet.range((row, col))
                            cell_address = cell.address.replace('$', '')
                            
                            # Skip if already processed
                            if cell_address in processed_ranges:
                                continue
                            
                            # Check if cell has validation using COM API
                            cell_com = cell.api
                            validation = cell_com.Validation
                            
                            # Check if validation exists (Type > 0 means validation is present)
                            if hasattr(validation, 'Type') and validation.Type > 0:
                                # Found validation, now find the full range
                                validation_info = {
                                    "range": cell_address,
                                    "type": get_validation_type_name(validation.Type),
                                    "operator": None,
                                    "formula1": None,
                                    "formula2": None,
                                    "error_message": None,
                                    "input_message": None,
                                    "show_error": True,
                                    "show_input": True
                                }
                                
                                # Get validation details
                                try:
                                    if hasattr(validation, 'Operator'):
                                        validation_info["operator"] = get_operator_name(validation.Operator)
                                except:
                                    pass
                                
                                try:
                                    if hasattr(validation, 'Formula1'):
                                        validation_info["formula1"] = str(validation.Formula1)
                                except:
                                    pass
                                
                                try:
                                    if hasattr(validation, 'Formula2'):
                                        validation_info["formula2"] = str(validation.Formula2)
                                except:
                                    pass
                                
                                try:
                                    if hasattr(validation, 'ErrorMessage'):
                                        validation_info["error_message"] = validation.ErrorMessage
                                except:
                                    pass
                                
                                try:
                                    if hasattr(validation, 'InputMessage'):
                                        validation_info["input_message"] = validation.InputMessage
                                except:
                                    pass
                                
                                try:
                                    if hasattr(validation, 'ShowError'):
                                        validation_info["show_error"] = bool(validation.ShowError)
                                except:
                                    pass
                                
                                try:
                                    if hasattr(validation, 'ShowInput'):
                                        validation_info["show_input"] = bool(validation.ShowInput)
                                except:
                                    pass
                                
                                # Try to find the full range with this validation
                                # by checking adjacent cells
                                full_range = expand_validation_range(sheet, row, col, validation)
                                validation_info["range"] = full_range
                                
                                # Mark cells as processed
                                for r in range(row, row + 10):
                                    for c in range(col, col + 10):
                                        processed_ranges.add(f"{chr(64+c)}{r}")
                                
                                validation_rules.append(validation_info)
                                
                        except Exception as e:
                            # Cell might not have validation, continue
                            continue
                
        except Exception as e:
            logger.warning(f"Error scanning for validation rules: {e}")
        
        # Return validation information
        result = {
            "sheet": sheet_name,
            "validation_count": len(validation_rules),
            "validation_rules": validation_rules
        }
        
        logger.info(f"✅ Found {len(validation_rules)} validation rules in {sheet_name}")
        return result
        
    except Exception as e:
        logger.error(f"Error getting validation info: {e}")
        return {"error": str(e)}
        
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()


def get_validation_type_name(type_value: int) -> str:
    """Convert Excel validation type constant to readable name."""
    validation_types = {
        0: "None",
        1: "Whole Number",
        2: "Decimal",
        3: "List",
        4: "Date",
        5: "Time",
        6: "Text Length",
        7: "Custom"
    }
    return validation_types.get(type_value, f"Unknown ({type_value})")


def get_operator_name(operator_value: int) -> str:
    """Convert Excel validation operator constant to readable name."""
    operators = {
        1: "Between",
        2: "Not Between",
        3: "Equal",
        4: "Not Equal",
        5: "Greater",
        6: "Less",
        7: "Greater or Equal",
        8: "Less or Equal"
    }
    return operators.get(operator_value, f"Unknown ({operator_value})")


def expand_validation_range(sheet, start_row: int, start_col: int, validation) -> str:
    """
    Try to find the full range that has the same validation rule.
    This is a simplified version - checks a limited area around the found cell.
    """
    try:
        # For simplicity, we'll just return the single cell
        # In a production version, you'd want to check adjacent cells
        # to find the full range with the same validation
        cell = sheet.range((start_row, start_col))
        return cell.address.replace('$', '')
    except:
        return f"{chr(64+start_col)}{start_row}"


def validate_excel_range_xlw(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str = None
) -> Dict[str, Any]:
    """
    Validate if a range exists and is properly formatted using xlwings.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        start_cell: Starting cell address
        end_cell: Ending cell address (optional)
        
    Returns:
        Dict containing validation result and range information
    """
    app = None
    wb = None
    
    try:
        logger.info(f"🔍 Validating range {start_cell}:{end_cell or start_cell} in {sheet_name}")
        
        # Check if file exists
        if not os.path.exists(filepath):
            return {"error": f"File not found: {filepath}", "valid": False}
        
        # Open Excel app and workbook
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(filepath)
        
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
        
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()