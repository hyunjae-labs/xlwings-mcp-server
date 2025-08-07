"""
Helper functions for xlwings implementation.
Provides abstraction layer for common operations and better error handling.
"""

import xlwings as xw
from typing import Optional, Dict, Any, Tuple
import logging

logger = logging.getLogger(__name__)


class ExcelHelper:
    """Helper class for Excel operations with better error handling and abstraction."""
    
    @staticmethod
    def find_empty_cell(sheet: xw.Sheet, start_row: int = 1, start_col: int = 1) -> str:
        """
        Find the next empty cell in a worksheet.
        
        Args:
            sheet: xlwings sheet object
            start_row: Starting row to search from
            start_col: Starting column to search from
            
        Returns:
            Cell address (e.g., "A3")
        """
        try:
            used_range = sheet.used_range
            if not used_range:
                return "A1"
            
            # Find first empty row after used range
            last_row = used_range.last_cell.row
            empty_row = last_row + 2  # Add some spacing
            
            # Convert column number to letter
            col_letter = ExcelHelper.get_column_letter(start_col)
            
            return f"{col_letter}{empty_row}"
        except Exception as e:
            logger.warning(f"Could not find empty cell: {e}, defaulting to A1")
            return "A1"
    
    @staticmethod
    def get_column_letter(col_idx: int) -> str:
        """
        Convert column index to Excel column letter.
        
        Args:
            col_idx: Column index (1-based)
            
        Returns:
            Column letter (e.g., "A", "B", "AA")
        """
        result = ""
        while col_idx > 0:
            col_idx -= 1
            result = chr(col_idx % 26 + ord('A')) + result
            col_idx //= 26
        return result
    
    @staticmethod
    def parse_range_with_sheet(range_str: str, wb: xw.Book, default_sheet_name: str) -> Tuple[xw.Sheet, xw.Range]:
        """
        Parse a range string that may include sheet reference.
        
        Args:
            range_str: Range string (e.g., "A1:B10" or "Sheet2!A1:B10")
            wb: Workbook object
            default_sheet_name: Default sheet name if not specified
            
        Returns:
            Tuple of (sheet, range)
        """
        if "!" in range_str:
            # Format: "SheetName!A1:E100"
            sheet_name, range_part = range_str.split("!", 1)
            # Remove quotes if present
            sheet_name = sheet_name.strip("'\"")
            
            # Check if sheet exists
            sheet_names = [s.name for s in wb.sheets]
            if sheet_name not in sheet_names:
                raise ValueError(f"Sheet '{sheet_name}' not found")
            
            sheet = wb.sheets[sheet_name]
            return sheet, sheet.range(range_part)
        else:
            # Use default sheet
            sheet = wb.sheets[default_sheet_name]
            return sheet, sheet.range(range_str)
    
    @staticmethod
    def generate_unique_name(base_name: str, existing_names: list) -> str:
        """
        Generate a unique name by appending numbers if needed.
        
        Args:
            base_name: Base name to use
            existing_names: List of existing names to avoid
            
        Returns:
            Unique name
        """
        if base_name not in existing_names:
            return base_name
        
        counter = 1
        while f"{base_name}{counter}" in existing_names:
            counter += 1
        
        return f"{base_name}{counter}"
    
    @staticmethod
    def safe_com_operation(operation, *fallback_operations, default=None, error_msg=None):
        """
        Safely execute COM operations with fallback options.
        
        Args:
            operation: Primary operation to try
            fallback_operations: Alternative operations to try if primary fails
            default: Default value if all operations fail
            error_msg: Custom error message for logging
            
        Returns:
            Result of successful operation or default value
        """
        operations = [operation] + list(fallback_operations)
        
        for i, op in enumerate(operations):
            try:
                return op()
            except Exception as e:
                if i == len(operations) - 1:
                    # Last operation failed
                    if error_msg:
                        logger.warning(f"{error_msg}: {e}")
                    return default
                # Try next operation
                continue
        
        return default
    
    @staticmethod
    def add_pivot_field(pivot_table, field_name: str, field_type: str, field_names: list) -> Tuple[bool, Optional[str]]:
        """
        Add a field to pivot table with better error handling.
        
        Args:
            pivot_table: COM pivot table object
            field_name: Name of the field to add
            field_type: Type of field ("row", "column", "value")
            field_names: List of all available field names
            
        Returns:
            Tuple of (success, error_message)
        """
        if field_name not in field_names:
            return False, f"{field_type.capitalize()} field '{field_name}' not found in data headers"
        
        orientation_map = {
            "row": 1,      # xlRowField
            "column": 2,   # xlColumnField
            "value": 4,    # xlDataField
            "page": 3      # xlPageField
        }
        
        orientation = orientation_map.get(field_type, 1)
        
        # Try multiple methods to add field
        def method1():
            field = pivot_table.PivotFields(field_name)
            field.Orientation = orientation
            return True
        
        def method2():
            field_index = field_names.index(field_name) + 1
            field = pivot_table.PivotFields(field_index)
            field.Orientation = orientation
            return True
        
        result = ExcelHelper.safe_com_operation(
            method1,
            method2,
            default=False,
            error_msg=f"Failed to add {field_type} field '{field_name}'"
        )
        
        if result:
            return True, None
        else:
            return False, f"Failed to add {field_type} field '{field_name}'"
    
    @staticmethod
    def set_aggregation_function(pivot_table, field_index: int, agg_func: str) -> bool:
        """
        Set aggregation function for a data field.
        
        Args:
            pivot_table: COM pivot table object
            field_index: Index of the data field (1-based)
            agg_func: Aggregation function name
            
        Returns:
            Success status
        """
        agg_map = {
            'sum': -4157,      # xlSum
            'count': -4112,    # xlCount
            'average': -4106,  # xlAverage
            'avg': -4106,      # xlAverage (alias)
            'mean': -4106,     # xlAverage (alias)
            'max': -4136,      # xlMax
            'min': -4139,      # xlMin
            'product': -4149,  # xlProduct
            'stdev': -4155,    # xlStDev
            'var': -4164,      # xlVar
        }
        
        agg_constant = agg_map.get(agg_func.lower(), -4157)  # Default to sum
        
        try:
            if pivot_table.DataFields.Count >= field_index:
                data_field = pivot_table.DataFields(field_index)
                data_field.Function = agg_constant
                return True
        except Exception as e:
            logger.warning(f"Could not set aggregation function: {e}")
        
        return False


class PivotTableBuilder:
    """Builder class for creating pivot tables with intelligent defaults."""
    
    def __init__(self, wb: xw.Book):
        self.wb = wb
        self.helper = ExcelHelper()
    
    def find_best_location(self, sheet: xw.Sheet) -> str:
        """
        Find the best location for a new pivot table.
        
        Args:
            sheet: Target sheet
            
        Returns:
            Cell address for pivot table
        """
        return self.helper.find_empty_cell(sheet)
    
    def generate_unique_pivot_name(self) -> str:
        """
        Generate a unique pivot table name across the workbook.
        
        Returns:
            Unique pivot table name
        """
        existing_names = []
        
        try:
            for sheet in self.wb.sheets:
                try:
                    sheet_pivots = sheet.api.PivotTables()
                    for i in range(1, sheet_pivots.Count + 1):
                        existing_names.append(sheet_pivots.Item(i).Name)
                except:
                    continue
        except:
            pass
        
        return self.helper.generate_unique_name("PivotTable", existing_names)
    
    def get_or_create_pivot_sheet(self, preferred_name: Optional[str] = None) -> xw.Sheet:
        """
        Get existing sheet or create new one for pivot table.
        
        Args:
            preferred_name: Preferred sheet name
            
        Returns:
            Sheet object
        """
        sheet_names = [s.name for s in self.wb.sheets]
        
        if preferred_name:
            if preferred_name in sheet_names:
                return self.wb.sheets[preferred_name]
            else:
                return self.wb.sheets.add(preferred_name)
        
        # Generate unique sheet name
        base_name = "PivotTable"
        sheet_name = self.helper.generate_unique_name(base_name, sheet_names)
        
        return self.wb.sheets.add(sheet_name)