"""
xlwings implementation for formatting and validation features.
Includes cell formatting and formula validation functionality.
"""

import xlwings as xw
from typing import Dict, Any, Optional
import logging
import os

logger = logging.getLogger(__name__)

def format_range_xlw(
    filepath: str,
    sheet_name: str,
    start_cell: str,
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
    merge_cells: bool = False
) -> Dict[str, Any]:
    """
    Apply formatting to a range of cells using xlwings.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        start_cell: Starting cell for formatting
        end_cell: Ending cell for formatting (optional, defaults to start_cell)
        bold: Apply bold formatting
        italic: Apply italic formatting
        underline: Apply underline formatting
        font_size: Font size in points
        font_color: Font color (hex code or color name)
        bg_color: Background color (hex code or color name)
        border_style: Border style (thin, medium, thick, double)
        border_color: Border color (hex code or color name)
        number_format: Number format string (e.g., "0.00", "#,##0", "mm/dd/yyyy")
        alignment: Text alignment (left, center, right, justify)
        wrap_text: Enable text wrapping
        merge_cells: Merge the cell range
        
    Returns:
        Dict with success message or error
    """
    app = None
    wb = None
    
    try:
        logger.info(f"🎨 Applying formatting to range {start_cell}:{end_cell or start_cell} in {sheet_name}")
        
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
        
        # Get the range to format
        if end_cell:
            range_obj = sheet.range(f"{start_cell}:{end_cell}")
        else:
            range_obj = sheet.range(start_cell)
        
        # Apply font formatting
        if bold:
            range_obj.font.bold = True
        if italic:
            range_obj.font.italic = True
        if underline:
            range_obj.font.underline = True
        if font_size:
            range_obj.font.size = font_size
        
        # Apply font color
        if font_color:
            try:
                # Convert hex color to RGB if needed
                if font_color.startswith('#'):
                    # Remove # and convert hex to RGB
                    hex_color = font_color.lstrip('#')
                    rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
                    range_obj.font.color = rgb
                else:
                    # Use color name or index
                    range_obj.font.color = font_color
            except:
                logger.warning(f"Could not apply font color: {font_color}")
        
        # Apply background color
        if bg_color:
            try:
                if bg_color.startswith('#'):
                    hex_color = bg_color.lstrip('#')
                    rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
                    range_obj.color = rgb
                else:
                    range_obj.color = bg_color
            except:
                logger.warning(f"Could not apply background color: {bg_color}")
        
        # Apply borders using COM API
        if border_style:
            range_com = range_obj.api
            
            # Map border styles to Excel constants
            border_map = {
                'thin': 1,      # xlThin
                'medium': -4138, # xlMedium
                'thick': 4,     # xlThick
                'double': -4119, # xlDouble
                'dotted': -4118, # xlDot
                'dashed': -4115  # xlDash
            }
            
            style_constant = border_map.get(border_style.lower(), 1)
            
            # Apply to all borders
            for border_index in [7, 8, 9, 10]:  # xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight
                border = range_com.Borders(border_index)
                border.LineStyle = style_constant
                
                if border_color:
                    try:
                        if border_color.startswith('#'):
                            hex_color = border_color.lstrip('#')
                            rgb_val = int(hex_color[:2], 16) + (int(hex_color[2:4], 16) << 8) + (int(hex_color[4:6], 16) << 16)
                            border.Color = rgb_val
                    except:
                        pass
        
        # Apply number format
        if number_format:
            range_obj.number_format = number_format
        
        # Apply alignment
        if alignment:
            alignment_map = {
                'left': -4131,    # xlLeft
                'center': -4108,  # xlCenter
                'right': -4152,   # xlRight
                'justify': -4130  # xlJustify
            }
            
            if alignment.lower() in alignment_map:
                range_obj.api.HorizontalAlignment = alignment_map[alignment.lower()]
        
        # Apply text wrapping
        if wrap_text:
            range_obj.api.WrapText = True
        
        # Merge cells if requested
        if merge_cells:
            range_obj.merge()
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully applied formatting to range")
        return {
            "message": f"Successfully applied formatting to range {start_cell}:{end_cell or start_cell}",
            "range": f"{start_cell}:{end_cell or start_cell}",
            "sheet": sheet_name,
            "formatting_applied": {
                "bold": bold,
                "italic": italic,
                "underline": underline,
                "font_size": font_size,
                "font_color": font_color,
                "bg_color": bg_color,
                "border_style": border_style,
                "number_format": number_format,
                "alignment": alignment,
                "wrap_text": wrap_text,
                "merged": merge_cells
            }
        }
        
    except Exception as e:
        logger.error(f"❌ Error applying formatting: {str(e)}")
        return {"error": str(e)}
        
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()


def validate_formula_syntax_xlw(
    filepath: str,
    sheet_name: str,
    cell: str,
    formula: str
) -> Dict[str, Any]:
    """
    Validate Excel formula syntax using xlwings without applying it.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        cell: Target cell for formula
        formula: Formula to validate
        
    Returns:
        Dict with validation result or error
    """
    app = None
    wb = None
    
    try:
        logger.info(f"🔍 Validating formula syntax: {formula}")
        
        # Check if file exists
        if not os.path.exists(filepath):
            return {"error": f"File not found: {filepath}"}
        
        # Ensure formula starts with =
        if not formula.startswith('='):
            formula = '=' + formula
        
        # Open Excel app and workbook
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(filepath)
        
        # Check if sheet exists
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        sheet = wb.sheets[sheet_name]
        
        # Try to apply the formula to a temporary cell to validate
        try:
            # Store original value
            target_cell = sheet.range(cell)
            original_value = target_cell.value
            original_formula = target_cell.formula
            
            # Try to set the formula
            target_cell.formula = formula
            
            # Check if Excel accepted the formula
            # If there's an error, Excel will show #NAME?, #VALUE!, etc.
            cell_value = target_cell.value
            
            # Check for common Excel errors
            excel_errors = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A']
            
            formula_valid = True
            error_type = None
            
            if isinstance(cell_value, str) and cell_value in excel_errors:
                formula_valid = False
                error_type = cell_value
            
            # Restore original value/formula
            if original_formula:
                target_cell.formula = original_formula
            else:
                target_cell.value = original_value
            
            # Don't save - we were just validating
            
            if formula_valid:
                logger.info(f"✅ Formula syntax is valid: {formula}")
                return {
                    "message": f"Formula syntax is valid",
                    "formula": formula,
                    "cell": cell,
                    "valid": True
                }
            else:
                logger.warning(f"⚠️ Formula has error: {error_type}")
                return {
                    "message": f"Formula contains error: {error_type}",
                    "formula": formula,
                    "cell": cell,
                    "valid": False,
                    "error_type": error_type
                }
                
        except Exception as e:
            # If we can't set the formula, it's invalid
            logger.error(f"❌ Invalid formula syntax: {str(e)}")
            return {
                "message": f"Invalid formula syntax: {str(e)}",
                "formula": formula,
                "cell": cell,
                "valid": False,
                "error": str(e)
            }
        
    except Exception as e:
        logger.error(f"❌ Error validating formula: {str(e)}")
        return {"error": str(e)}
        
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()