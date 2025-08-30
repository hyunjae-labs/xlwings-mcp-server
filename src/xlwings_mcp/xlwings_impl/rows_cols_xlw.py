"""
xlwings implementation for row and column operations.
Handles insertion and deletion of rows and columns in Excel worksheets.
"""

import xlwings as xw
from typing import Dict, Any
import logging
import os

logger = logging.getLogger(__name__)


def insert_rows_xlw(
    filepath: str,
    sheet_name: str,
    start_row: int,
    count: int = 1
) -> Dict[str, Any]:
    """
    Insert one or more rows in Excel using xlwings.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        start_row: Row number to start insertion (1-based)
        count: Number of rows to insert
        
    Returns:
        Dict with success message or error
    """
    app = None
    wb = None
    
    try:
        logger.info(f"📊 Inserting {count} rows at row {start_row} in {sheet_name}")
        
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
        
        # Insert rows using xlwings
        # Get the range for the row to insert at
        target_row = sheet.range(f"{start_row}:{start_row}")
        
        # Insert rows - use COM API for better control
        for _ in range(count):
            target_row.api.Insert()
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully inserted {count} rows at row {start_row}")
        return {
            "message": f"Successfully inserted {count} rows at row {start_row}",
            "sheet": sheet_name,
            "start_row": start_row,
            "count": count
        }
        
    except Exception as e:
        logger.error(f"Error inserting rows: {e}")
        return {"error": str(e)}
        
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()


def insert_columns_xlw(
    filepath: str,
    sheet_name: str,
    start_col: int,
    count: int = 1
) -> Dict[str, Any]:
    """
    Insert one or more columns in Excel using xlwings.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        start_col: Column number to start insertion (1-based)
        count: Number of columns to insert
        
    Returns:
        Dict with success message or error
    """
    app = None
    wb = None
    
    try:
        logger.info(f"📊 Inserting {count} columns at column {start_col} in {sheet_name}")
        
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
        
        # Convert column number to letter
        def col_num_to_letter(n):
            string = ""
            while n > 0:
                n, remainder = divmod(n - 1, 26)
                string = chr(65 + remainder) + string
            return string
        
        col_letter = col_num_to_letter(start_col)
        
        # Insert columns using xlwings
        target_col = sheet.range(f"{col_letter}:{col_letter}")
        
        # Insert columns - use COM API for better control
        for _ in range(count):
            target_col.api.Insert()
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully inserted {count} columns at column {col_letter}")
        return {
            "message": f"Successfully inserted {count} columns at column {col_letter}",
            "sheet": sheet_name,
            "start_col": start_col,
            "count": count
        }
        
    except Exception as e:
        logger.error(f"Error inserting columns: {e}")
        return {"error": str(e)}
        
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()


def delete_sheet_rows_xlw(
    filepath: str,
    sheet_name: str,
    start_row: int,
    count: int = 1
) -> Dict[str, Any]:
    """
    Delete one or more rows in Excel using xlwings.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        start_row: Row number to start deletion (1-based)
        count: Number of rows to delete
        
    Returns:
        Dict with success message or error
    """
    app = None
    wb = None
    
    try:
        logger.info(f"🗑️ Deleting {count} rows starting from row {start_row} in {sheet_name}")
        
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
        
        # Delete rows using xlwings
        # Delete from bottom to top to maintain row indices
        for i in range(count):
            row_to_delete = sheet.range(f"{start_row}:{start_row}")
            row_to_delete.api.Delete()
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully deleted {count} rows starting from row {start_row}")
        return {
            "message": f"Successfully deleted {count} rows starting from row {start_row}",
            "sheet": sheet_name,
            "start_row": start_row,
            "count": count
        }
        
    except Exception as e:
        logger.error(f"Error deleting rows: {e}")
        return {"error": str(e)}
        
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()


def delete_sheet_columns_xlw(
    filepath: str,
    sheet_name: str,
    start_col: int,
    count: int = 1
) -> Dict[str, Any]:
    """
    Delete one or more columns in Excel using xlwings.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        start_col: Column number to start deletion (1-based)
        count: Number of columns to delete
        
    Returns:
        Dict with success message or error
    """
    app = None
    wb = None
    
    try:
        logger.info(f"🗑️ Deleting {count} columns starting from column {start_col} in {sheet_name}")
        
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
        
        # Convert column number to letter
        def col_num_to_letter(n):
            string = ""
            while n > 0:
                n, remainder = divmod(n - 1, 26)
                string = chr(65 + remainder) + string
            return string
        
        col_letter = col_num_to_letter(start_col)
        
        # Delete columns using xlwings
        # Delete multiple times since we delete one at a time
        for i in range(count):
            col_to_delete = sheet.range(f"{col_letter}:{col_letter}")
            col_to_delete.api.Delete()
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully deleted {count} columns starting from column {col_letter}")
        return {
            "message": f"Successfully deleted {count} columns starting from column {col_letter}",
            "sheet": sheet_name,
            "start_col": start_col,
            "count": count
        }
        
    except Exception as e:
        logger.error(f"Error deleting columns: {e}")
        return {"error": str(e)}
        
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()

def insert_rows_xlw_with_wb(
    wb,
    sheet_name: str,
    start_row: int,
    count: int = 1
) -> Dict[str, Any]:
    """Session-based version using existing workbook object.
    
    Args:
        wb: Workbook object from session
        sheet_name: Name of worksheet
        start_row: Row number to start insertion (1-based)
        count: Number of rows to insert
        
    Returns:
        Dict with success message or error
    """
    try:
        logger.info(f"📊 Inserting {count} rows at row {start_row} in {sheet_name}")
        
        # Check if sheet exists
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        sheet = wb.sheets[sheet_name]
        
        # Insert rows using xlwings
        # Get the range for the row to insert at
        target_row = sheet.range(f"{start_row}:{start_row}")
        
        # Insert rows - use COM API for better control
        for _ in range(count):
            target_row.api.Insert()
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully inserted {count} rows at row {start_row}")
        return {
            "message": f"Successfully inserted {count} rows at row {start_row}",
            "sheet": sheet_name,
            "start_row": start_row,
            "count": count
        }
        
    except Exception as e:
        logger.error(f"Error inserting rows: {e}")
        return {"error": str(e)}

def insert_columns_xlw_with_wb(
    wb,
    sheet_name: str,
    start_col: int,
    count: int = 1
) -> Dict[str, Any]:
    """Session-based version using existing workbook object.
    
    Args:
        wb: Workbook object from session
        sheet_name: Name of worksheet
        start_col: Column number to start insertion (1-based)
        count: Number of columns to insert
        
    Returns:
        Dict with success message or error
    """
    try:
        logger.info(f"📊 Inserting {count} columns at column {start_col} in {sheet_name}")
        
        # Check if sheet exists
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        sheet = wb.sheets[sheet_name]
        
        # Convert column number to letter
        def col_num_to_letter(n):
            string = ""
            while n > 0:
                n, remainder = divmod(n - 1, 26)
                string = chr(65 + remainder) + string
            return string
        
        col_letter = col_num_to_letter(start_col)
        
        # Insert columns using xlwings
        target_col = sheet.range(f"{col_letter}:{col_letter}")
        
        # Insert columns - use COM API for better control
        for _ in range(count):
            target_col.api.Insert()
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully inserted {count} columns at column {col_letter}")
        return {
            "message": f"Successfully inserted {count} columns at column {col_letter}",
            "sheet": sheet_name,
            "start_col": start_col,
            "count": count
        }
        
    except Exception as e:
        logger.error(f"Error inserting columns: {e}")
        return {"error": str(e)}

def delete_sheet_rows_xlw_with_wb(
    wb,
    sheet_name: str,
    start_row: int,
    count: int = 1
) -> Dict[str, Any]:
    """Session-based version using existing workbook object.
    
    Args:
        wb: Workbook object from session
        sheet_name: Name of worksheet
        start_row: Row number to start deletion (1-based)
        count: Number of rows to delete
        
    Returns:
        Dict with success message or error
    """
    try:
        logger.info(f"🗑️ Deleting {count} rows starting from row {start_row} in {sheet_name}")
        
        # Check if sheet exists
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        sheet = wb.sheets[sheet_name]
        
        # Delete rows using xlwings
        # Delete from bottom to top to maintain row indices
        for i in range(count):
            row_to_delete = sheet.range(f"{start_row}:{start_row}")
            row_to_delete.api.Delete()
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully deleted {count} rows starting from row {start_row}")
        return {
            "message": f"Successfully deleted {count} rows starting from row {start_row}",
            "sheet": sheet_name,
            "start_row": start_row,
            "count": count
        }
        
    except Exception as e:
        logger.error(f"Error deleting rows: {e}")
        return {"error": str(e)}

def delete_sheet_columns_xlw_with_wb(
    wb,
    sheet_name: str,
    start_col: int,
    count: int = 1
) -> Dict[str, Any]:
    """Session-based version using existing workbook object.
    
    Args:
        wb: Workbook object from session
        sheet_name: Name of worksheet
        start_col: Column number to start deletion (1-based)
        count: Number of columns to delete
        
    Returns:
        Dict with success message or error
    """
    try:
        logger.info(f"🗑️ Deleting {count} columns starting from column {start_col} in {sheet_name}")
        
        # Check if sheet exists
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        sheet = wb.sheets[sheet_name]
        
        # Convert column number to letter
        def col_num_to_letter(n):
            string = ""
            while n > 0:
                n, remainder = divmod(n - 1, 26)
                string = chr(65 + remainder) + string
            return string
        
        col_letter = col_num_to_letter(start_col)
        
        # Delete columns using xlwings
        # Delete multiple times since we delete one at a time
        for i in range(count):
            col_to_delete = sheet.range(f"{col_letter}:{col_letter}")
            col_to_delete.api.Delete()
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully deleted {count} columns starting from column {col_letter}")
        return {
            "message": f"Successfully deleted {count} columns starting from column {col_letter}",
            "sheet": sheet_name,
            "start_col": start_col,
            "count": count
        }
        
    except Exception as e:
        logger.error(f"Error deleting columns: {e}")
        return {"error": str(e)}