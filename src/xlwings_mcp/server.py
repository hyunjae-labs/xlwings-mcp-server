import logging
import os
from typing import Any, List, Dict, Optional

from mcp.server.fastmcp import FastMCP

# Excel 엔진: xlwings 구현으로 완전 전환 완료

# Import exceptions
from xlwings_mcp.exceptions import (
    ValidationError,
    WorkbookError,
    SheetError,
    DataError,
    FormattingError,
    CalculationError,
    PivotError,
    ChartError
)

# Import from xlwings_mcp package (only functions still used by server)
from xlwings_mcp.workbook import get_workbook_info
from xlwings_mcp.sheet import (
    copy_sheet,
    delete_sheet,
    rename_sheet,
)

# Import session management
from xlwings_mcp.session import SESSION_MANAGER
from xlwings_mcp.force_close import force_close_workbook_by_path

# Get project root directory path for log file path.
# When using the stdio transmission method,
# relative paths may cause log files to fail to create
# due to the client's running location and permission issues,
# resulting in the program not being able to run.
# Thus using os.path.join(ROOT_DIR, "excel-mcp.log") instead.

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
LOG_FILE = os.path.join(ROOT_DIR, "excel-mcp.log")

# Initialize EXCEL_FILES_PATH variable without assigning a value
EXCEL_FILES_PATH = None

# xlwings 구현 사용 (openpyxl 마이그레이션 완료)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        # Referring to https://github.com/modelcontextprotocol/python-sdk/issues/409#issuecomment-2816831318
        # The stdio mode server MUST NOT write anything to its stdout that is not a valid MCP message.
        logging.FileHandler(LOG_FILE)
    ],
)
logger = logging.getLogger("excel-mcp")
logger.info("🚀 Excel MCP Server starting - xlwings 모드 활성화")

# Error message templates for consistent error reporting
ERROR_TEMPLATES = {
    'SESSION_NOT_FOUND': "SESSION_NOT_FOUND: Session '{session_id}' not found. It may have expired after {ttl} minutes of inactivity. Use open_workbook() to create a new session.",
    'SESSION_TIMEOUT': "SESSION_TIMEOUT: Session '{session_id}' expired at {time}. Create new session with open_workbook()",
    'FILE_LOCKED': "FILE_ACCESS_ERROR: '{filepath}' is locked by another process. Use force_close_workbook_by_path() to force close it first.",
    'FILE_NOT_FOUND': "FILE_NOT_FOUND: '{filepath}' does not exist. Check the path or create a new workbook with create_workbook().",
    'SHEET_NOT_FOUND': "SHEET_NOT_FOUND: Sheet '{sheet_name}' not found in workbook. Available sheets: {sheets}",
    'INVALID_RANGE': "INVALID_RANGE: Range '{range}' is not valid. Use format like 'A1' or 'A1:B10'.",
    'PARAMETER_MISSING': "PARAMETER_MISSING: Either {param1} or {param2} must be provided.",
}

# Initialize FastMCP server
mcp = FastMCP(
    "excel-mcp",
    instructions="Excel MCP Server for manipulating Excel files"
)

def get_excel_path(filename: str) -> str:
    """Get full path to Excel file.
    
    Args:
        filename: Name of Excel file
        
    Returns:
        Full path to Excel file
    """
    # If filename is already an absolute path, return it
    if os.path.isabs(filename):
        return filename

    # Check if in SSE mode (EXCEL_FILES_PATH is not None)
    if EXCEL_FILES_PATH is None:
        # Must use absolute path
        raise ValueError(f"Invalid filename: {filename}, must be an absolute path when not in SSE mode")

    # In SSE mode, if it's a relative path, resolve it based on EXCEL_FILES_PATH
    return os.path.join(EXCEL_FILES_PATH, filename)

# ============================================================================
# SESSION MANAGEMENT TOOLS (NEW)
# ============================================================================

@mcp.tool()
def open_workbook(
    filepath: str,
    visible: bool = False,
    read_only: bool = False
) -> Dict[str, Any]:
    """
    Open an Excel workbook and create a session.
    
    Args:
        filepath: Path to Excel file
        visible: Whether to show Excel window (default: False)
        read_only: Whether to open in read-only mode (default: False)
        
    Returns:
        Dictionary with session_id, filepath, visible, read_only, and sheets
    """
    try:
        full_path = get_excel_path(filepath)
        session_id = SESSION_MANAGER.open_workbook(full_path, visible, read_only)
        
        # Get session info
        session = SESSION_MANAGER.get_session(session_id)
        if not session:
            raise WorkbookError(f"Failed to create session for {filepath}")
        
        return {
            "session_id": session_id,
            "filepath": session.filepath,
            "visible": session.visible,
            "read_only": session.read_only,
            "sheets": [sheet.name for sheet in session.workbook.sheets]
        }
        
    except Exception as e:
        logger.error(f"Error opening workbook: {e}")
        raise WorkbookError(f"Failed to open workbook: {str(e)}")

@mcp.tool()
def close_workbook(
    session_id: str,
    save: bool = True
) -> str:
    """
    Close a workbook session.
    
    Args:
        session_id: Session ID from open_workbook
        save: Whether to save changes (default: True)
        
    Returns:
        Success message
    """
    try:
        success = SESSION_MANAGER.close_workbook(session_id, save)
        if not success:
            raise WorkbookError(f"Session {session_id} not found")
        
        return f"Workbook session {session_id} closed successfully"
        
    except Exception as e:
        logger.error(f"Error closing workbook: {e}")
        raise WorkbookError(f"Failed to close workbook: {str(e)}")

@mcp.tool()
def list_workbooks() -> List[Dict[str, Any]]:
    """
    List all open workbook sessions.
    
    Returns:
        List of session information dictionaries
    """
    try:
        return SESSION_MANAGER.list_sessions()
    except Exception as e:
        logger.error(f"Error listing workbooks: {e}")
        raise WorkbookError(f"Failed to list workbooks: {str(e)}")

@mcp.tool()
def force_close_workbook_by_path_tool(
    filepath: str
) -> Dict[str, Any]:
    """
    Force close a specific workbook by file path (without saving).
    
    Args:
        filepath: Path to the workbook to force close
        
    Returns:
        Dictionary with 'closed' (bool) and 'message' (str)
    """
    try:
        full_path = get_excel_path(filepath)
        return force_close_workbook_by_path(full_path)
    except Exception as e:
        logger.error(f"Error force closing workbook: {e}")
        return {
            "closed": False,
            "message": f"Failed to force close workbook: {str(e)}"
        }

# ============================================================================
# EXISTING TOOLS (TO BE UPDATED WITH SESSION SUPPORT)
# ============================================================================

@mcp.tool()
def apply_formula(
    session_id: str,
    sheet_name: str,
    cell: str,
    formula: str
) -> str:
    """
    Apply Excel formula to cell.
    
    Args:
        session_id: Session ID from open_workbook (required)
        sheet_name: Name of worksheet
        cell: Cell address (e.g., "A1")
        formula: Excel formula to apply
    """
    try:
        # Get session (required)
        session = SESSION_MANAGER.get_session(session_id)
        if not session:
            return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                session_id=session_id, 
                ttl=10  # Default TTL is 10 minutes (600 seconds)
            )
        
        with session.lock:
            from xlwings_mcp.xlwings_impl.calculations_xlw import apply_formula_xlw_with_wb
            result = apply_formula_xlw_with_wb(session.workbook, sheet_name, cell, formula)
        
        return result.get("message", "Formula applied successfully") if "error" not in result else f"Error: {result['error']}"
            
    except (ValidationError, CalculationError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error applying formula: {e}")
        raise

@mcp.tool()
def validate_formula_syntax(
    sheet_name: str,
    cell: str,
    formula: str,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None
) -> str:
    """
    Validate Excel formula syntax without applying it.
    
    Args:
        sheet_name: Name of worksheet
        cell: Cell address (e.g., "A1")
        formula: Excel formula to validate
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.calculations_xlw import validate_formula_syntax_xlw_with_wb
                result = validate_formula_syntax_xlw_with_wb(session.workbook, sheet_name, cell, formula)
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.xlwings_impl.calculations_xlw import validate_formula_syntax_xlw
            result = validate_formula_syntax_xlw(full_path, sheet_name, cell, formula)
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        return result.get("message", "Formula validation completed") if "error" not in result else f"Error: {result['error']}"
    except (ValidationError, CalculationError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error validating formula: {e}")
        raise

@mcp.tool()
def format_range(
    sheet_name: str,
    start_cell: str,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None,
    end_cell: Optional[str] = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_size: Optional[int] = None,
    font_color: Optional[str] = None,
    bg_color: Optional[str] = None,
    border_style: Optional[str] = None,
    border_color: Optional[str] = None,
    number_format: Optional[str] = None,
    alignment: Optional[str] = None,
    wrap_text: bool = False,
    merge_cells: bool = False,
    protection: Optional[Dict[str, Any]] = None,
    conditional_format: Optional[Dict[str, Any]] = None
) -> str:
    """
    Apply formatting to a range of cells.
    
    Args:
        sheet_name: Name of worksheet
        start_cell: Starting cell
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        end_cell: Ending cell (optional)
        bold: Apply bold formatting
        italic: Apply italic formatting
        underline: Apply underline formatting
        font_size: Font size
        font_color: Font color
        bg_color: Background color
        border_style: Border style
        border_color: Border color
        number_format: Number format
        alignment: Text alignment
        wrap_text: Enable text wrapping
        merge_cells: Merge cells in range
        protection: Cell protection settings (optional)
        conditional_format: Conditional formatting settings (optional)
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.formatting_xlw import format_range_xlw_with_wb
                result = format_range_xlw_with_wb(
                    session.workbook,
                    sheet_name=sheet_name,
                    start_cell=start_cell,
                    end_cell=end_cell,
                    bold=bold,
                    italic=italic,
                    underline=underline,
                    font_size=font_size,
                    font_color=font_color,
                    bg_color=bg_color,
                    border_style=border_style,
                    border_color=border_color,
                    number_format=number_format,
                    alignment=alignment,
                    wrap_text=wrap_text,
                    merge_cells=merge_cells
                )
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.xlwings_impl.formatting_xlw import format_range_xlw
            result = format_range_xlw(
                filepath=full_path,
                sheet_name=sheet_name,
                start_cell=start_cell,
                end_cell=end_cell,
                bold=bold,
                italic=italic,
                underline=underline,
                font_size=font_size,
                font_color=font_color,
                bg_color=bg_color,
                border_style=border_style,
                border_color=border_color,
                number_format=number_format,
                alignment=alignment,
                wrap_text=wrap_text,
                merge_cells=merge_cells
            )
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        return result.get("message", "Range formatted successfully") if "error" not in result else f"Error: {result['error']}"
    except (ValidationError, FormattingError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error formatting range: {e}")
        raise

@mcp.tool()
def read_data_from_excel(
    session_id: str,
    sheet_name: str,
    start_cell: Optional[str] = None,
    end_cell: Optional[str] = None,
    preview_only: bool = False
) -> str:
    """
    Read data from Excel worksheet with cell metadata including validation rules.
    
    Args:
        session_id: Session ID from open_workbook (required)
        sheet_name: Name of worksheet
        start_cell: Starting cell (default A1)
        end_cell: Ending cell (optional, auto-expands if not provided)
        preview_only: Whether to return preview only
    """
    try:
        # Get session (required)
        session = SESSION_MANAGER.get_session(session_id)
        if not session:
            return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                session_id=session_id, 
                ttl=10  # Default TTL is 10 minutes (600 seconds)
            )
        
        with session.lock:
            from xlwings_mcp.xlwings_impl.data_xlw import read_data_from_excel_xlw_with_wb
            return read_data_from_excel_xlw_with_wb(session.workbook, sheet_name, start_cell, end_cell, preview_only)
        
    except (ValidationError, DataError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error reading data: {e}")
        raise

@mcp.tool()
def write_data_to_excel(
    session_id: str,
    sheet_name: str,
    data: List[List],
    start_cell: Optional[str] = None
) -> str:
    """
    Write data to Excel worksheet.
    Excel formula will write to cell without any verification.

    Args:
        session_id: Session ID from open_workbook (required)
        sheet_name: Name of worksheet to write to
        data: List of lists containing data to write to the worksheet, sublists are assumed to be rows
        start_cell: Cell to start writing to (optional, auto-finds appropriate location)
    """
    try:
        # Get session (required)
        session = SESSION_MANAGER.get_session(session_id)
        if not session:
            return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                session_id=session_id, 
                ttl=10  # Default TTL is 10 minutes (600 seconds)
            )
        
        with session.lock:
            from xlwings_mcp.xlwings_impl.data_xlw import write_data_to_excel_xlw_with_wb
            result = write_data_to_excel_xlw_with_wb(session.workbook, sheet_name, data, start_cell)
        
        return result.get("message", "Data written successfully") if "error" not in result else f"Error: {result['error']}"
            
    except (ValidationError, DataError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error writing data: {e}")
        raise

@mcp.tool()
def create_workbook(
    session_id: Optional[str] = None,
    filepath: Optional[str] = None
) -> str:
    """
    Create new Excel workbook.
    
    Args:
        session_id: Session ID for creating workbook in existing session (optional)
        filepath: Path to create new Excel file (legacy, deprecated)
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session (though this is less common for creating new workbooks)
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return f"Error: Session {session_id} not found. Please open the workbook first using open_workbook()."
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.workbook_xlw import create_workbook_xlw_with_wb
                result = create_workbook_xlw_with_wb(session.workbook)
                return result.get("message", "Workbook created successfully") if "error" not in result else f"Error: {result['error']}"
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.workbook import create_workbook as create_workbook_impl
            create_workbook_impl(full_path)
            return f"Created workbook at {full_path}"
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
    except WorkbookError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating workbook: {e}")
        raise

@mcp.tool()
def create_worksheet(
    sheet_name: str,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None
) -> str:
    """
    Create new worksheet in workbook.
    
    Args:
        sheet_name: Name of the new worksheet
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.workbook_xlw import create_worksheet_xlw_with_wb
                result = create_worksheet_xlw_with_wb(session.workbook, sheet_name)
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.workbook import create_sheet as create_worksheet_impl
            result = create_worksheet_impl(full_path, sheet_name)
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        return result.get("message", "Worksheet created successfully") if "error" not in result else f"Error: {result['error']}"
        
    except (ValidationError, WorkbookError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating worksheet: {e}")
        raise

@mcp.tool()
def create_chart(
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None,
    title: str = "",
    x_axis: str = "",
    y_axis: str = ""
) -> str:
    """
    Create chart in worksheet.
    
    Args:
        sheet_name: Name of worksheet
        data_range: Data range for chart
        chart_type: Type of chart
        target_cell: Cell where chart will be placed
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        title: Chart title (optional)
        x_axis: X-axis label (optional)
        y_axis: Y-axis label (optional)
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.advanced_xlw import create_chart_xlw_with_wb
                result = create_chart_xlw_with_wb(
                    session.workbook,
                    sheet_name=sheet_name,
                    data_range=data_range,
                    chart_type=chart_type,
                    target_cell=target_cell,
                    title=title,
                    x_axis=x_axis,
                    y_axis=y_axis
                )
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.xlwings_impl.advanced_xlw import create_chart_xlw
            result = create_chart_xlw(
                filepath=full_path,
                sheet_name=sheet_name,
                data_range=data_range,
                chart_type=chart_type,
                target_cell=target_cell,
                title=title,
                x_axis=x_axis,
                y_axis=y_axis
            )
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        return result.get("message", "Chart created successfully") if "error" not in result else f"Error: {result['error']}"
        
    except (ValidationError, ChartError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating chart: {e}")
        raise

@mcp.tool()
def create_pivot_table(
    sheet_name: str,
    data_range: str,
    rows: List[str],
    values: List[str],
    session_id: Optional[str] = None,
    filepath: Optional[str] = None,
    columns: Optional[List[str]] = None,
    agg_func: str = "mean",
    target_sheet: Optional[str] = None,
    target_cell: Optional[str] = None,
    pivot_name: Optional[str] = None
) -> str:
    """
    Create pivot table in worksheet.
    
    Args:
        sheet_name: Name of worksheet containing source data
        data_range: Source data range (e.g., "A1:E100" or "Sheet2!A1:E100")
        rows: Field names for row labels
        values: Field names for values
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        columns: Field names for column labels (optional)
        agg_func: Aggregation function (sum, count, average, max, min)
        target_sheet: Target sheet for pivot table (optional, auto-created if not exists)
        target_cell: Target cell for pivot table (optional, finds empty area if not provided)
        pivot_name: Custom name for pivot table (optional, auto-generated if not provided)
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.advanced_xlw import create_pivot_table_xlw_with_wb
                result = create_pivot_table_xlw_with_wb(
                    session.workbook,
                    sheet_name=sheet_name,
                    data_range=data_range,
                    rows=rows,
                    values=values,
                    columns=columns,
                    agg_func=agg_func,
                    target_sheet=target_sheet,
                    target_cell=target_cell,
                    pivot_name=pivot_name
                )
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.xlwings_impl.advanced_xlw import create_pivot_table_xlw
            result = create_pivot_table_xlw(
                filepath=full_path,
                sheet_name=sheet_name,
                data_range=data_range,
                rows=rows,
                values=values,
                columns=columns,
                agg_func=agg_func,
                target_sheet=target_sheet,
                target_cell=target_cell,
                pivot_name=pivot_name
            )
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        # Handle warnings in response
        if "warnings" in result and result["warnings"]:
            warning_msg = "; ".join(result["warnings"])
            return f"{result.get('message', 'Pivot table created')} (Warnings: {warning_msg})"
        
        return result.get("message", "Pivot table created successfully") if "error" not in result else f"Error: {result['error']}"
        
    except (ValidationError, PivotError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating pivot table: {e}")
        raise

@mcp.tool()
def create_table(
    sheet_name: str,
    data_range: str,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None,
    table_name: Optional[str] = None,
    table_style: str = "TableStyleMedium9"
) -> str:
    """
    Creates a native Excel table from a specified range of data.
    
    Args:
        sheet_name: Name of worksheet
        data_range: Range of data to create table from
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        table_name: Name for the table (optional)
        table_style: Style for the table (optional)
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.advanced_xlw import create_table_xlw_with_wb
                result = create_table_xlw_with_wb(
                    session.workbook,
                    sheet_name=sheet_name,
                    data_range=data_range,
                    table_name=table_name,
                    table_style=table_style
                )
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.xlwings_impl.advanced_xlw import create_table_xlw
            result = create_table_xlw(
                filepath=full_path,
                sheet_name=sheet_name,
                data_range=data_range,
                table_name=table_name,
                table_style=table_style
            )
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        return result.get("message", "Table created successfully") if "error" not in result else f"Error: {result['error']}"
        
    except DataError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating table: {e}")
        raise

@mcp.tool()
def copy_worksheet(
    source_sheet: str,
    target_sheet: str,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None
) -> str:
    """
    Copy worksheet within workbook.
    
    Args:
        source_sheet: Name of the source worksheet
        target_sheet: Name of the target worksheet
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.workbook_xlw import copy_worksheet_xlw_with_wb
                result = copy_worksheet_xlw_with_wb(session.workbook, source_sheet, target_sheet)
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            result = copy_sheet(full_path, source_sheet, target_sheet)
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        return result.get("message", "Worksheet copied successfully") if "error" not in result else f"Error: {result['error']}"
        
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error copying worksheet: {e}")
        raise

@mcp.tool()
def delete_worksheet(
    sheet_name: str,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None
) -> str:
    """
    Delete worksheet from workbook.
    
    Args:
        sheet_name: Name of the worksheet to delete
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.workbook_xlw import delete_worksheet_xlw_with_wb
                result = delete_worksheet_xlw_with_wb(session.workbook, sheet_name)
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            result = delete_sheet(full_path, sheet_name)
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        return result.get("message", "Worksheet deleted successfully") if "error" not in result else f"Error: {result['error']}"
        
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error deleting worksheet: {e}")
        raise

@mcp.tool()
def rename_worksheet(
    old_name: str,
    new_name: str,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None
) -> str:
    """
    Rename worksheet in workbook.
    
    Args:
        old_name: Current name of the worksheet
        new_name: New name for the worksheet
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.workbook_xlw import rename_worksheet_xlw_with_wb
                result = rename_worksheet_xlw_with_wb(session.workbook, old_name, new_name)
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            result = rename_sheet(full_path, old_name, new_name)
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        return result.get("message", "Worksheet renamed successfully") if "error" not in result else f"Error: {result['error']}"
        
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error renaming worksheet: {e}")
        raise

@mcp.tool()
def get_workbook_metadata(
    session_id: Optional[str] = None,
    filepath: Optional[str] = None,
    include_ranges: bool = False
) -> str:
    """
    Get metadata about workbook including sheets, ranges, etc.
    
    Args:
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        include_ranges: Whether to include range information
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.workbook_xlw import get_workbook_metadata_xlw_with_wb
                result = get_workbook_metadata_xlw_with_wb(session.workbook, include_ranges=include_ranges)
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.xlwings_impl.workbook_xlw import get_workbook_metadata_xlw
            result = get_workbook_metadata_xlw(full_path, include_ranges=include_ranges)
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        if "error" in result:
            return f"Error: {result['error']}"
        import json
        return json.dumps(result, indent=2, default=str, ensure_ascii=False)
            
    except (ValidationError, WorkbookError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error getting workbook metadata: {e}")
        raise

@mcp.tool()
def merge_cells(
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None
) -> str:
    """
    Merge a range of cells.
    
    Args:
        sheet_name: Name of worksheet
        start_cell: Starting cell
        end_cell: Ending cell
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.range_xlw import merge_cells_xlw_with_wb
                result = merge_cells_xlw_with_wb(session.workbook, sheet_name, start_cell, end_cell)
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.xlwings_impl.range_xlw import merge_cells_xlw
            result = merge_cells_xlw(full_path, sheet_name, start_cell, end_cell)
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        return result.get("message", "Cells merged successfully") if "error" not in result else f"Error: {result['error']}"
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error merging cells: {e}")
        raise

@mcp.tool()
def unmerge_cells(
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None
) -> str:
    """
    Unmerge a range of cells.
    
    Args:
        sheet_name: Name of worksheet
        start_cell: Starting cell
        end_cell: Ending cell
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.range_xlw import unmerge_cells_xlw_with_wb
                result = unmerge_cells_xlw_with_wb(session.workbook, sheet_name, start_cell, end_cell)
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.xlwings_impl.range_xlw import unmerge_cells_xlw
            result = unmerge_cells_xlw(full_path, sheet_name, start_cell, end_cell)
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        return result.get("message", "Cells unmerged successfully") if "error" not in result else f"Error: {result['error']}"
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error unmerging cells: {e}")
        raise

@mcp.tool()
def get_merged_cells(
    sheet_name: str,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None
) -> str:
    """
    Get merged cells in a worksheet.
    
    Args:
        sheet_name: Name of worksheet
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.range_xlw import get_merged_cells_xlw_with_wb
                result = get_merged_cells_xlw_with_wb(session.workbook, sheet_name)
                if "error" in result:
                    return f"Error: {result['error']}"
                import json
                return json.dumps(result, indent=2, default=str)
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.xlwings_impl.range_xlw import get_merged_cells_xlw
            result = get_merged_cells_xlw(full_path, sheet_name)
            if "error" in result:
                return f"Error: {result['error']}"
            import json
            return json.dumps(result, indent=2, default=str)
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error getting merged cells: {e}")
        raise

@mcp.tool()
def copy_range(
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None,
    target_sheet: Optional[str] = None
) -> str:
    """
    Copy a range of cells to another location.
    
    Args:
        sheet_name: Name of source worksheet
        source_start: Starting cell of source range
        source_end: Ending cell of source range
        target_start: Starting cell of target range
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        target_sheet: Target worksheet (optional, uses source sheet if not provided)
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.range_xlw import copy_range_xlw_with_wb
                result = copy_range_xlw_with_wb(
                    session.workbook,
                    sheet_name,
                    source_start,
                    source_end,
                    target_start,
                    target_sheet or sheet_name  # Use source sheet if target_sheet is None
                )
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.xlwings_impl.range_xlw import copy_range_xlw
            result = copy_range_xlw(
                full_path,
                sheet_name,
                source_start,
                source_end,
                target_start,
                target_sheet or sheet_name  # Use source sheet if target_sheet is None
            )
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        return result.get("message", "Range copied successfully") if "error" not in result else f"Error: {result['error']}"
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error copying range: {e}")
        raise

@mcp.tool()
def delete_range(
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None,
    shift_direction: str = "up"
) -> str:
    """
    Delete a range of cells and shift remaining cells.
    
    Args:
        sheet_name: Name of worksheet
        start_cell: Starting cell
        end_cell: Ending cell
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        shift_direction: Direction to shift cells ("up" or "left")
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.range_xlw import delete_range_xlw_with_wb
                result = delete_range_xlw_with_wb(
                    session.workbook,
                    sheet_name,
                    start_cell,
                    end_cell,
                    shift_direction
                )
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.xlwings_impl.range_xlw import delete_range_xlw
            result = delete_range_xlw(
                full_path,
                sheet_name,
                start_cell,
                end_cell,
                shift_direction
            )
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        return result.get("message", "Range deleted successfully") if "error" not in result else f"Error: {result['error']}"
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error deleting range: {e}")
        raise

@mcp.tool()
def validate_excel_range(
    sheet_name: str,
    start_cell: str,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None,
    end_cell: Optional[str] = None
) -> str:
    """
    Validate if a range exists and is properly formatted.
    
    Args:
        sheet_name: Name of worksheet
        start_cell: Starting cell
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        end_cell: Ending cell (optional)
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.validation_xlw import validate_excel_range_xlw_with_wb
                result = validate_excel_range_xlw_with_wb(session.workbook, sheet_name, start_cell, end_cell)
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.xlwings_impl.validation_xlw import validate_excel_range_xlw
            result = validate_excel_range_xlw(full_path, sheet_name, start_cell, end_cell)
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        return result.get("message", "Range validation completed") if "error" not in result else f"Error: {result['error']}"
            
    except (ValidationError, DataError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error validating range: {e}")
        raise

@mcp.tool()
def get_data_validation_info(
    sheet_name: str,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None
) -> str:
    """
    Get all data validation rules in a worksheet.
    
    This tool helps identify which cell ranges have validation rules
    and what types of validation are applied.
    
    Args:
        sheet_name: Name of worksheet
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
        
    Returns:
        JSON string containing all validation rules in the worksheet
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.validation_xlw import get_data_validation_info_xlw_with_wb
                result = get_data_validation_info_xlw_with_wb(session.workbook, sheet_name)
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.xlwings_impl.validation_xlw import get_data_validation_info_xlw
            result = get_data_validation_info_xlw(full_path, sheet_name)
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        if "error" in result:
            return f"Error: {result['error']}"
        import json
        return json.dumps(result, indent=2, default=str)
        
    except (ValidationError, DataError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error getting validation info: {e}")
        raise

@mcp.tool()
def insert_rows(
    sheet_name: str,
    start_row: int,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None,
    count: int = 1
) -> str:
    """
    Insert one or more rows starting at the specified row.
    
    Args:
        sheet_name: Name of worksheet
        start_row: Row number to start inserting at
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        count: Number of rows to insert
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.rows_cols_xlw import insert_rows_xlw_with_wb
                result = insert_rows_xlw_with_wb(session.workbook, sheet_name, start_row, count)
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.xlwings_impl.rows_cols_xlw import insert_rows_xlw
            result = insert_rows_xlw(full_path, sheet_name, start_row, count)
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        if "error" in result:
            return f"Error: {result['error']}"
        return result["message"]
        
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error inserting rows: {e}")
        raise

@mcp.tool()
def insert_columns(
    sheet_name: str,
    start_col: int,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None,
    count: int = 1
) -> str:
    """
    Insert one or more columns starting at the specified column.
    
    Args:
        sheet_name: Name of worksheet
        start_col: Column number to start inserting at
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        count: Number of columns to insert
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.rows_cols_xlw import insert_columns_xlw_with_wb
                result = insert_columns_xlw_with_wb(session.workbook, sheet_name, start_col, count)
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.xlwings_impl.rows_cols_xlw import insert_columns_xlw
            result = insert_columns_xlw(full_path, sheet_name, start_col, count)
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        if "error" in result:
            return f"Error: {result['error']}"
        return result["message"]
        
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error inserting columns: {e}")
        raise

@mcp.tool()
def delete_sheet_rows(
    sheet_name: str,
    start_row: int,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None,
    count: int = 1
) -> str:
    """
    Delete one or more rows starting at the specified row.
    
    Args:
        sheet_name: Name of worksheet
        start_row: Row number to start deleting from
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        count: Number of rows to delete
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.rows_cols_xlw import delete_sheet_rows_xlw_with_wb
                result = delete_sheet_rows_xlw_with_wb(session.workbook, sheet_name, start_row, count)
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.xlwings_impl.rows_cols_xlw import delete_sheet_rows_xlw
            result = delete_sheet_rows_xlw(full_path, sheet_name, start_row, count)
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        if "error" in result:
            return f"Error: {result['error']}"
        return result["message"]
        
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error deleting rows: {e}")
        raise

@mcp.tool()
def delete_sheet_columns(
    sheet_name: str,
    start_col: int,
    session_id: Optional[str] = None,
    filepath: Optional[str] = None,
    count: int = 1
) -> str:
    """
    Delete one or more columns starting at the specified column.
    
    Args:
        sheet_name: Name of worksheet
        start_col: Column number to start deleting from
        session_id: Session ID from open_workbook (preferred)
        filepath: Path to Excel file (legacy, deprecated)
        count: Number of columns to delete
        
    Note: Use session_id for better performance. filepath parameter is deprecated.
    """
    try:
        # Support both new (session_id) and old (filepath) API
        if session_id:
            # New API: use session
            session = SESSION_MANAGER.get_session(session_id)
            if not session:
                return ERROR_TEMPLATES['SESSION_NOT_FOUND'].format(
                    session_id=session_id, 
                    ttl=10  # Default TTL is 10 minutes (600 seconds)
                )
            
            with session.lock:
                from xlwings_mcp.xlwings_impl.rows_cols_xlw import delete_sheet_columns_xlw_with_wb
                result = delete_sheet_columns_xlw_with_wb(session.workbook, sheet_name, start_col, count)
        elif filepath:
            # Legacy API: backwards compatibility
            logger.warning("Using deprecated filepath parameter. Please use session_id instead.")
            full_path = get_excel_path(filepath)
            from xlwings_mcp.xlwings_impl.rows_cols_xlw import delete_sheet_columns_xlw
            result = delete_sheet_columns_xlw(full_path, sheet_name, start_col, count)
        else:
            return ERROR_TEMPLATES['PARAMETER_MISSING'].format(
                param1='session_id',
                param2='filepath'
            )
        
        if "error" in result:
            return f"Error: {result['error']}"
        return result["message"]
        
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error deleting columns: {e}")
        raise

async def run_sse():
    """Run Excel MCP server in SSE mode."""
    # Assign value to EXCEL_FILES_PATH in SSE mode
    global EXCEL_FILES_PATH
    EXCEL_FILES_PATH = os.environ.get("EXCEL_FILES_PATH", "./excel_files")
    # Create directory if it doesn't exist
    os.makedirs(EXCEL_FILES_PATH, exist_ok=True)
    
    try:
        logger.info(f"Starting Excel MCP server with SSE transport (files directory: {EXCEL_FILES_PATH})")
        await mcp.run_sse_async()
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        # Clean up all sessions on shutdown
        try:
            SESSION_MANAGER.close_all_sessions()
            logger.info("All Excel sessions closed")
        except Exception as e:
            logger.error(f"Error closing sessions during shutdown: {e}")
        logger.info("Server shutdown complete")

async def run_streamable_http():
    """Run Excel MCP server in streamable HTTP mode."""
    # Assign value to EXCEL_FILES_PATH in streamable HTTP mode
    global EXCEL_FILES_PATH
    EXCEL_FILES_PATH = os.environ.get("EXCEL_FILES_PATH", "./excel_files")
    # Create directory if it doesn't exist
    os.makedirs(EXCEL_FILES_PATH, exist_ok=True)
    
    try:
        logger.info(f"Starting Excel MCP server with streamable HTTP transport (files directory: {EXCEL_FILES_PATH})")
        await mcp.run_streamable_http_async()
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        # Clean up all sessions on shutdown
        try:
            SESSION_MANAGER.close_all_sessions()
            logger.info("All Excel sessions closed")
        except Exception as e:
            logger.error(f"Error closing sessions during shutdown: {e}")
        logger.info("Server shutdown complete")

def run_stdio():
    """Run Excel MCP server in stdio mode."""
    # No need to assign EXCEL_FILES_PATH in stdio mode
    
    try:
        logger.info("Starting Excel MCP server with stdio transport")
        mcp.run(transport="stdio")
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        # Clean up all sessions on shutdown
        try:
            SESSION_MANAGER.close_all_sessions()
            logger.info("All Excel sessions closed")
        except Exception as e:
            logger.error(f"Error closing sessions during shutdown: {e}")
        logger.info("Server shutdown complete")