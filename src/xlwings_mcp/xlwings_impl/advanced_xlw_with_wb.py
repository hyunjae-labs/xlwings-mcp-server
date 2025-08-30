"""
Session-based versions of advanced xlwings functions.
These functions use existing workbook objects instead of opening/closing Excel apps.
"""

import xlwings as xw
from typing import Dict, Any, List, Optional
import logging

logger = logging.getLogger(__name__)

def create_chart_xlw_with_wb(
    wb,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis: str = "",
    y_axis: str = ""
) -> Dict[str, Any]:
    """Session-based version using existing workbook object.
    
    Args:
        wb: Workbook object from session
        sheet_name: Name of worksheet
        data_range: Range of data for chart (e.g., "A1:C10")
        chart_type: Type of chart (line, bar, pie, scatter, area, column)
        target_cell: Cell where chart will be positioned
        title: Chart title
        x_axis: X-axis label
        y_axis: Y-axis label
        
    Returns:
        Dict with success message or error
    """
    try:
        logger.info(f"📈 Creating {chart_type} chart in {sheet_name}")
        
        # Check if sheet exists
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        sheet = wb.sheets[sheet_name]
        
        # Map chart types to Excel constants
        chart_type_map = {
            'line': 4,          # xlLine
            'bar': 57,          # xlBarClustered
            'column': 51,       # xlColumnClustered
            'pie': 5,           # xlPie
            'scatter': 74,      # xlXYScatter
            'area': 1,          # xlArea
        }
        
        if chart_type.lower() not in chart_type_map:
            return {"error": f"Unsupported chart type: {chart_type}"}
        
        excel_chart_type = chart_type_map[chart_type.lower()]
        
        # Get data range first
        data_range_obj = sheet.range(data_range)
        
        # Create chart using xlwings method
        chart = sheet.charts.add()
        
        # Set data source
        chart.set_source_data(data_range_obj)
        
        # Set chart type - handle COM API properly
        try:
            if hasattr(chart, 'chart_type'):
                # Use xlwings built-in chart type property
                chart.chart_type = chart_type.lower()
            else:
                # Use COM API more carefully
                chart_api = chart.api
                if hasattr(chart_api, 'ChartType'):
                    chart_api.ChartType = excel_chart_type
                else:
                    logger.warning("Cannot set chart type - using default")
        except Exception as e:
            logger.warning(f"Chart type setting failed: {e}, using default")
        
        # Set chart position
        target = sheet.range(target_cell)
        chart.top = target.top
        chart.left = target.left
        
        # Calculate chart size based on data range
        data_rows = data_range_obj.rows.count
        data_cols = data_range_obj.columns.count
        
        # Dynamic sizing based on data
        chart.width = min(600, max(400, data_cols * 80))  # Adjust width based on columns
        chart.height = min(450, max(300, data_rows * 15))  # Adjust height based on rows
        
        # Set chart properties safely
        try:
            chart_com = chart.api
            
            # Set title
            if title and hasattr(chart_com, 'HasTitle'):
                chart_com.HasTitle = True
                if hasattr(chart_com, 'ChartTitle'):
                    chart_com.ChartTitle.Text = title
            
            # Set axis labels
            if hasattr(chart_com, 'Axes'):
                try:
                    if x_axis:
                        x_axis_obj = chart_com.Axes(1)  # xlCategory
                        x_axis_obj.HasTitle = True
                        x_axis_obj.AxisTitle.Text = x_axis
                        
                    if y_axis:
                        y_axis_obj = chart_com.Axes(2)  # xlValue
                        y_axis_obj.HasTitle = True
                        y_axis_obj.AxisTitle.Text = y_axis
                except Exception as e:
                    logger.warning(f"Axis label setting failed: {e}")
        except:
            # Some chart types don't have axes
            pass
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully created {chart_type} chart")
        return {
            "message": f"Successfully created {chart_type} chart",
            "chart_type": chart_type,
            "data_range": data_range,
            "position": target_cell,
            "sheet": sheet_name
        }
        
    except Exception as e:
        logger.error(f"❌ Error creating chart: {str(e)}")
        return {"error": str(e)}

def create_pivot_table_xlw_with_wb(
    wb,
    sheet_name: str,
    data_range: str,
    rows: List[str],
    values: List[str],
    columns: Optional[List[str]] = None,
    agg_func: str = "sum",
    target_sheet: Optional[str] = None,
    target_cell: str = None,
    pivot_name: Optional[str] = None
) -> Dict[str, Any]:
    """Session-based version using existing workbook object.
    
    Args:
        wb: Workbook object from session
        sheet_name: Name of worksheet containing source data
        data_range: Source data range (e.g., "A1:E100" or "Sheet2!A1:E100")
        rows: Field names for row labels
        values: Field names for values
        columns: Field names for column labels (optional)
        agg_func: Aggregation function (sum, count, average, max, min)
        target_sheet: Target sheet for pivot table (optional)
        target_cell: Target cell for pivot table (optional, default finds empty area)
        pivot_name: Custom name for pivot table (optional)
        
    Returns:
        Dict with success message or error
    """
    try:
        logger.info(f"📊 Creating pivot table in {sheet_name}")
        
        # Check if sheet exists
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        # Parse data range to support cross-sheet references
        if "!" in data_range:
            # Format: "SheetName!A1:E100"
            source_sheet_name, range_part = data_range.split("!", 1)
            # Remove quotes if present
            source_sheet_name = source_sheet_name.strip('\'"')
            if source_sheet_name not in sheet_names:
                return {"error": f"Source sheet '{source_sheet_name}' not found"}
            source_sheet = wb.sheets[source_sheet_name]
            source_range = source_sheet.range(range_part)
        else:
            # Use the provided sheet_name
            source_sheet = wb.sheets[sheet_name]
            source_range = source_sheet.range(data_range)
        
        # Determine target sheet for pivot table
        if target_sheet:
            # Use specified target sheet
            if target_sheet not in sheet_names:
                # Create if doesn't exist
                pivot_sheet = wb.sheets.add(target_sheet)
            else:
                pivot_sheet = wb.sheets[target_sheet]
        else:
            # Auto-generate unique pivot sheet name
            pivot_sheet_name = "PivotTable"
            counter = 1
            while pivot_sheet_name in sheet_names:
                pivot_sheet_name = f"PivotTable{counter}"
                counter += 1
            pivot_sheet = wb.sheets.add(pivot_sheet_name)
        
        # Determine target cell position
        if not target_cell:
            # Find empty area automatically
            used_range = pivot_sheet.used_range
            if used_range:
                # Place below existing content with some spacing
                target_cell = f"A{used_range.last_cell.row + 3}"
            else:
                target_cell = "A3"  # Default position if sheet is empty
        
        # Use COM API to create pivot table
        pivot_cache = wb.api.PivotCaches().Create(
            SourceType=1,  # xlDatabase
            SourceData=source_range.api
        )
        
        # Generate unique pivot table name if not provided
        if not pivot_name:
            existing_pivots = []
            try:
                # Try to get existing pivot table names
                for sheet in wb.sheets:
                    try:
                        sheet_pivots = sheet.api.PivotTables()
                        for i in range(1, sheet_pivots.Count + 1):
                            existing_pivots.append(sheet_pivots.Item(i).Name)
                    except:
                        pass
            except:
                pass
            
            # Generate unique name
            pivot_name = "PivotTable1"
            counter = 1
            while pivot_name in existing_pivots:
                counter += 1
                pivot_name = f"PivotTable{counter}"
        
        pivot_table = pivot_cache.CreatePivotTable(
            TableDestination=pivot_sheet.range(target_cell).api,
            TableName=pivot_name
        )
        
        # Get field names from first row of data (use source_range which is already parsed)
        header_range = source_range.rows[0]
        field_names = [cell.value for cell in header_range]
        
        # Track warnings for partial failures
        warnings = []
        
        # Add row fields - try different COM API access methods
        for row_field in rows:
            if row_field in field_names:
                success = False
                try:
                    # Method 1: Direct string access
                    field = pivot_table.PivotFields(row_field)
                    field.Orientation = 1  # xlRowField
                    success = True
                except:
                    try:
                        # Method 2: Index access
                        field_index = field_names.index(row_field) + 1
                        field = pivot_table.PivotFields(field_index)
                        field.Orientation = 1  # xlRowField
                        success = True
                    except Exception as e:
                        error_msg = f"Failed to add row field '{row_field}': {str(e)}"
                        logger.warning(error_msg)
                        warnings.append(error_msg)
            else:
                warnings.append(f"Row field '{row_field}' not found in data headers")
        
        # Add column fields
        if columns:
            for col_field in columns:
                if col_field in field_names:
                    success = False
                    try:
                        # Method 1: Direct string access
                        field = pivot_table.PivotFields(col_field)
                        field.Orientation = 2  # xlColumnField
                        success = True
                    except:
                        try:
                            # Method 2: Index access
                            field_index = field_names.index(col_field) + 1
                            field = pivot_table.PivotFields(field_index)
                            field.Orientation = 2  # xlColumnField
                            success = True
                        except Exception as e:
                            error_msg = f"Failed to add column field '{col_field}': {str(e)}"
                            logger.warning(error_msg)
                            warnings.append(error_msg)
                else:
                    warnings.append(f"Column field '{col_field}' not found in data headers")
        
        # Add value fields with aggregation
        for value_field in values:
            if value_field in field_names:
                success = False
                try:
                    # Method 1: Direct string access
                    field = pivot_table.PivotFields(value_field)
                    field.Orientation = 4  # xlDataField
                    success = True
                    logger.info(f"Added value field '{value_field}' successfully")
                except:
                    try:
                        # Method 2: Index access
                        field_index = field_names.index(value_field) + 1
                        field = pivot_table.PivotFields(field_index)
                        field.Orientation = 4  # xlDataField
                        success = True
                        logger.info(f"Added value field '{value_field}' using index")
                    except Exception as e:
                        error_msg = f"Failed to add value field '{value_field}': {str(e)}"
                        logger.warning(error_msg)
                        warnings.append(error_msg)
                
                # Try to set aggregation function if field was added successfully
                if success and agg_func.lower() != 'sum':
                    try:
                        # Safer approach: iterate through DataFields to find our field
                        agg_map = {
                            'count': -4112,    # xlCount
                            'average': -4106,  # xlAverage
                            'mean': -4106,     # xlAverage (alias)
                            'max': -4136,      # xlMax
                            'min': -4139,      # xlMin
                        }
                        
                        if agg_func.lower() in agg_map:
                            # Wait a moment for COM to update
                            import time
                            time.sleep(0.1)
                            
                            # Try to find and update the data field
                            for i in range(1, pivot_table.DataFields.Count + 1):
                                try:
                                    data_field = pivot_table.DataFields(i)
                                    # Check if this is our field (name contains the original field name)
                                    if value_field in str(data_field.SourceName):
                                        data_field.Function = agg_map[agg_func.lower()]
                                        logger.info(f"Set aggregation to {agg_func} for {value_field}")
                                        break
                                except:
                                    continue
                    except Exception as e:
                        # Non-critical: aggregation function setting failed
                        logger.debug(f"Could not set aggregation function for {value_field}: {e}")
            else:
                warnings.append(f"Value field '{value_field}' not found in data headers")
        
        # Apply default pivot table style
        pivot_table.TableStyle2 = "PivotStyleMedium9"
        
        # Save the workbook
        wb.save()
        
        # Prepare result
        result = {
            "message": f"Successfully created pivot table '{pivot_name}'",
            "pivot_name": pivot_name,
            "pivot_sheet": pivot_sheet.name,
            "pivot_cell": target_cell,
            "source_range": data_range,
            "source_sheet": source_sheet.name,
            "rows": rows,
            "columns": columns or [],
            "values": values,
            "aggregation": agg_func
        }
        
        # Add warnings if any
        if warnings:
            result["warnings"] = warnings
            logger.info(f"⚠️ Pivot table created with warnings: {warnings}")
        else:
            logger.info(f"✅ Successfully created pivot table '{pivot_name}' at {pivot_sheet.name}!{target_cell}")
        
        return result
        
    except Exception as e:
        logger.error(f"❌ Error creating pivot table: {str(e)}")
        return {"error": str(e)}

def create_table_xlw_with_wb(
    wb,
    sheet_name: str,
    data_range: str,
    table_name: Optional[str] = None,
    table_style: str = "TableStyleMedium9"
) -> Dict[str, Any]:
    """Session-based version using existing workbook object.
    
    Args:
        wb: Workbook object from session
        sheet_name: Name of worksheet
        data_range: Range of data to convert to table (e.g., "A1:D10")
        table_name: Name for the table (optional)
        table_style: Excel table style name
        
    Returns:
        Dict with success message or error
    """
    try:
        logger.info(f"📋 Creating Excel table in {sheet_name}")
        
        # Check if sheet exists
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        sheet = wb.sheets[sheet_name]
        
        # Get data range
        range_obj = sheet.range(data_range)
        
        # Generate table name if not provided
        if not table_name:
            existing_tables = sheet.api.ListObjects
            table_name = f"Table{existing_tables.Count + 1}"
        
        # Create table using COM API
        sheet_com = sheet.api
        table = sheet_com.ListObjects.Add(
            SourceType=1,  # xlSrcRange
            Source=range_obj.api,
            XlListObjectHasHeaders=1  # xlYes
        )
        
        # Set table name
        table.Name = table_name
        
        # Apply table style
        table.TableStyle = table_style
        
        # Enable filtering
        table.ShowAutoFilter = True
        
        # Enable total row (optional, disabled by default)
        table.ShowTotals = False
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully created table '{table_name}'")
        return {
            "message": f"Successfully created Excel table",
            "table_name": table_name,
            "data_range": data_range,
            "style": table_style,
            "sheet": sheet_name,
            "has_headers": True,
            "has_filter": True
        }
        
    except Exception as e:
        logger.error(f"❌ Error creating table: {str(e)}")
        return {"error": str(e)}