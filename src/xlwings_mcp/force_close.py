"""
Force close utility for Excel workbooks.
Uses pywin32 to close specific workbooks by file path without saving.
"""

import os
import logging
import sys
from typing import Optional, Dict, Any

logger = logging.getLogger(__name__)

# Platform-specific implementation
if sys.platform == "win32":
    try:
        import win32com.client
        import pythoncom
        PYWIN32_AVAILABLE = True
    except ImportError:
        PYWIN32_AVAILABLE = False
        logger.warning("pywin32 not available - force close functionality disabled")
else:
    PYWIN32_AVAILABLE = False


def force_close_workbook_by_path(filepath: str) -> Dict[str, Any]:
    """
    Force close a specific workbook from any running Excel process.
    
    Args:
        filepath: Absolute path to the workbook to close
        
    Returns:
        Dictionary with 'closed' (bool) and 'message' (str)
    """
    
    if not PYWIN32_AVAILABLE:
        return {
            "closed": False,
            "message": "Force close not available (pywin32 not installed or not on Windows)"
        }
    
    # Normalize the file path
    target_path = os.path.abspath(filepath).lower()
    
    try:
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        found = False
        closed = False
        
        try:
            # Try to connect to running Excel instance
            xl = win32com.client.GetObject(Class="Excel.Application")
            
            # Check all open workbooks
            for wb in xl.Workbooks:
                try:
                    # Compare full paths (case-insensitive on Windows)
                    wb_path = os.path.abspath(wb.FullName).lower()
                    
                    if wb_path == target_path:
                        logger.info(f"Found workbook to force close: {wb.FullName}")
                        found = True
                        
                        # Force close without saving
                        wb.Close(SaveChanges=False)
                        closed = True
                        logger.info(f"Successfully force closed: {filepath}")
                        break
                        
                except Exception as e:
                    logger.warning(f"Error checking/closing workbook: {e}")
                    continue
            
            # If no workbooks remain, optionally quit Excel
            if closed and xl.Workbooks.Count == 0:
                try:
                    xl.Quit()
                    logger.info("Excel application quit (no remaining workbooks)")
                except:
                    pass
                    
        except Exception as e:
            # No Excel instance running or other COM error
            logger.debug(f"Could not connect to Excel: {e}")
            return {
                "closed": False,
                "message": f"No Excel instance found or cannot connect: {str(e)}"
            }
        
        finally:
            # Uninitialize COM
            pythoncom.CoUninitialize()
        
        if not found:
            return {
                "closed": False,
                "message": f"Workbook not found in any Excel instance: {filepath}"
            }
        
        return {
            "closed": closed,
            "message": f"Successfully force closed workbook: {filepath}"
        }
        
    except Exception as e:
        logger.error(f"Force close failed for {filepath}: {e}")
        return {
            "closed": False,
            "message": f"Force close failed: {str(e)}"
        }


def force_close_all_excel_instances() -> Dict[str, Any]:
    """
    Force close all Excel instances (emergency recovery).
    
    Returns:
        Dictionary with 'count' (int) and 'message' (str)
    """
    
    if not PYWIN32_AVAILABLE:
        return {
            "count": 0,
            "message": "Force close not available (pywin32 not installed or not on Windows)"
        }
    
    count = 0
    
    try:
        pythoncom.CoInitialize()
        
        try:
            # Try to connect to Excel
            xl = win32com.client.GetObject(Class="Excel.Application")
            
            # Close all workbooks without saving
            while xl.Workbooks.Count > 0:
                try:
                    xl.Workbooks(1).Close(SaveChanges=False)
                    count += 1
                except:
                    break
            
            # Quit Excel
            xl.Quit()
            logger.info(f"Force closed {count} workbooks and quit Excel")
            
        except Exception as e:
            logger.debug(f"No Excel instance to close: {e}")
            
        finally:
            pythoncom.CoUninitialize()
            
        return {
            "count": count,
            "message": f"Closed {count} workbook(s) and quit Excel"
        }
        
    except Exception as e:
        logger.error(f"Force close all failed: {e}")
        return {
            "count": 0,
            "message": f"Force close all failed: {str(e)}"
        }