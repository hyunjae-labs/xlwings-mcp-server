"""
xlwings implementation for advanced Excel features.
Includes charts, pivot tables, and Excel tables functionality.
"""

import xlwings as xw
from typing import Dict, Any, List, Optional
import logging
import os

logger = logging.getLogger(__name__)

def create_chart_xlw(
    filepath: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis: str = "",
    y_axis: str = ""
) -> Dict[str, Any]:
    """
    Create a chart in Excel using xlwings.
    
    Args:
        filepath: Path to Excel file
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
    app = None
    wb = None
    
    try:
        logger.info(f"📈 Creating {chart_type} chart in {sheet_name}")
        
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
        chart.width = 400  # Default width
        chart.height = 300  # Default height
        
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
        
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()


def create_pivot_table_xlw(
    filepath: str,
    sheet_name: str,
    data_range: str,
    rows: List[str],
    values: List[str],
    columns: Optional[List[str]] = None,
    agg_func: str = "sum"
) -> Dict[str, Any]:
    """
    Create a pivot table in Excel using xlwings.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        data_range: Source data range (e.g., "A1:E100")
        rows: Field names for row labels
        values: Field names for values
        columns: Field names for column labels (optional)
        agg_func: Aggregation function (sum, count, average, max, min)
        
    Returns:
        Dict with success message or error
    """
    app = None
    wb = None
    
    try:
        logger.info(f"📊 Creating pivot table in {sheet_name}")
        
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
        
        # Create a new sheet for pivot table
        pivot_sheet_name = "PivotTable"
        counter = 1
        while pivot_sheet_name in sheet_names:
            pivot_sheet_name = f"PivotTable{counter}"
            counter += 1
        
        pivot_sheet = wb.sheets.add(pivot_sheet_name)
        
        # Get source data range
        source_range = sheet.range(data_range)
        
        # Use COM API to create pivot table
        pivot_cache = wb.api.PivotCaches().Create(
            SourceType=1,  # xlDatabase
            SourceData=source_range.api
        )
        
        pivot_table = pivot_cache.CreatePivotTable(
            TableDestination=pivot_sheet.range("A3").api,
            TableName="PivotTable1"
        )
        
        # Get field names from first row of data
        header_range = sheet.range(data_range).rows[0]
        field_names = [cell.value for cell in header_range]
        
        # Add row fields - try different COM API access methods
        for row_field in rows:
            if row_field in field_names:
                try:
                    # Method 1: Direct string access
                    field = pivot_table.PivotFields(row_field)
                    field.Orientation = 1  # xlRowField
                except:
                    try:
                        # Method 2: Index access
                        field_index = field_names.index(row_field) + 1
                        field = pivot_table.PivotFields(field_index)
                        field.Orientation = 1  # xlRowField
                    except Exception as e:
                        logger.warning(f"Failed to add row field {row_field}: {e}")
        
        # Add column fields
        if columns:
            for col_field in columns:
                if col_field in field_names:
                    try:
                        # Method 1: Direct string access
                        field = pivot_table.PivotFields(col_field)
                        field.Orientation = 2  # xlColumnField
                    except:
                        try:
                            # Method 2: Index access
                            field_index = field_names.index(col_field) + 1
                            field = pivot_table.PivotFields(field_index)
                            field.Orientation = 2  # xlColumnField
                        except Exception as e:
                            logger.warning(f"Failed to add column field {col_field}: {e}")
        
        # Add value fields with aggregation
        agg_map = {
            'sum': -4157,      # xlSum
            'count': -4112,    # xlCount
            'average': -4106,  # xlAverage
            'max': -4136,      # xlMax
            'min': -4139,      # xlMin
        }
        
        agg_constant = agg_map.get(agg_func.lower(), -4157)  # Default to sum
        
        for value_field in values:
            if value_field in field_names:
                try:
                    # Method 1: Direct string access
                    field = pivot_table.PivotFields(value_field)
                    field.Orientation = 4  # xlDataField
                    # Set aggregation function
                    if pivot_table.DataFields.Count > 0:
                        data_field = pivot_table.DataFields(1)  # First data field
                        data_field.Function = agg_constant
                except:
                    try:
                        # Method 2: Index access
                        field_index = field_names.index(value_field) + 1
                        field = pivot_table.PivotFields(field_index)
                        field.Orientation = 4  # xlDataField
                        if pivot_table.DataFields.Count > 0:
                            data_field = pivot_table.DataFields(1)
                            data_field.Function = agg_constant
                    except Exception as e:
                        logger.warning(f"Failed to add value field {value_field}: {e}")
        
        # Apply default pivot table style
        pivot_table.TableStyle2 = "PivotStyleMedium9"
        
        # Save the workbook
        wb.save()
        
        logger.info(f"✅ Successfully created pivot table in {pivot_sheet_name}")
        return {
            "message": f"Successfully created pivot table",
            "pivot_sheet": pivot_sheet_name,
            "source_range": data_range,
            "rows": rows,
            "columns": columns or [],
            "values": values,
            "aggregation": agg_func
        }
        
    except Exception as e:
        logger.error(f"❌ Error creating pivot table: {str(e)}")
        return {"error": str(e)}
        
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()


def create_table_xlw(
    filepath: str,
    sheet_name: str,
    data_range: str,
    table_name: Optional[str] = None,
    table_style: str = "TableStyleMedium9"
) -> Dict[str, Any]:
    """
    Create an Excel table (ListObject) using xlwings.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        data_range: Range of data to convert to table (e.g., "A1:D10")
        table_name: Name for the table (optional)
        table_style: Excel table style name
        
    Returns:
        Dict with success message or error
    """
    app = None
    wb = None
    
    try:
        logger.info(f"📋 Creating Excel table in {sheet_name}")
        
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
        
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()