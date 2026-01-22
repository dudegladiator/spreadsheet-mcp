"""Google Sheets MCP Server - FastMCP implementation with all tools."""

import json
from mcp.server.fastmcp import FastMCP

from .sheets_client import get_client

# Initialize the MCP server
mcp = FastMCP(
    "Google Sheets",
    instructions="""
Google Sheets MCP Server - Full spreadsheet manipulation capabilities.

## Getting Started
1. Get spreadsheet_id from URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
2. Use get_spreadsheet_info first to see available sheets and their IDs
3. Use A1 notation for ranges (examples below)

## A1 Notation Reference
- "Sheet1!A1:D10" - Cells A1 to D10 in Sheet1
- "Sheet1!A:A" - Entire column A
- "Sheet1!1:5" - Rows 1 through 5
- "A1:B5" - Range in first visible sheet
- "'My Sheet'!A1" - Sheet names with spaces need quotes

## Common Workflows

### Reading Data
1. Use read_cells to get values from a range
2. Use batch_read for multiple ranges at once
3. Use get_last_row to find where data ends

### Writing Data
1. Use write_cells for single range (supports formulas like "=SUM(A1:A10)")
2. Use batch_write for multiple ranges
3. Use append_rows to add data at the end

### Formulas
Write formulas as string values starting with "=":
- "=SUM(A1:A10)" - Sum values
- "=AVERAGE(B:B)" - Average entire column
- "=VLOOKUP(A1,Sheet2!A:B,2,FALSE)" - Lookup values
- "=IF(A1>100,\"High\",\"Low\")" - Conditional logic

### Formatting
Use format_cells with 0-based row/column indices:
- Row 1 = index 0, Column A = index 0
- end_row/end_col are exclusive (row 0-1 means just row 0)

### Charts
1. First write your data to the sheet
2. Use create_chart with chart_type: BAR, LINE, PIE, COLUMN, AREA, SCATTER
3. Charts are positioned by row/column offset

### Sharing
Use share_spreadsheet to:
- Share with specific email (role: reader/writer/commenter)
- Make public with make_public=True

## Tips
- Always check get_spreadsheet_info first to get sheet_id values
- sheet_id is numeric (e.g., 0, 123456), sheet name is text (e.g., "Sheet1")
- Row/column indices in formatting are 0-based
- A1 notation in read/write is 1-based (A1 is first cell)
""",
)


# =============================================================================
# Spreadsheet Management Tools
# =============================================================================


@mcp.tool()
def create_spreadsheet(title: str, sheet_names: str = "") -> str:
    """Create a new Google Spreadsheet.

    Args:
        title: The title for the new spreadsheet.
        sheet_names: Comma-separated list of sheet names (optional).

    Returns:
        JSON with spreadsheet_id, title, url, and sheets.
    """
    client = get_client()
    sheets = (
        [s.strip() for s in sheet_names.split(",") if s.strip()]
        if sheet_names
        else None
    )
    result = client.create_spreadsheet(title, sheets)
    return json.dumps(result, indent=2)


@mcp.tool()
def get_spreadsheet_info(spreadsheet_id: str) -> str:
    """Get information about a spreadsheet including all sheets.

    Args:
        spreadsheet_id: The ID of the spreadsheet (from URL).

    Returns:
        JSON with spreadsheet metadata, sheets list, and URL.
    """
    client = get_client()
    result = client.get_spreadsheet_info(spreadsheet_id)
    return json.dumps(result, indent=2)


# =============================================================================
# Sheet Management Tools
# =============================================================================


@mcp.tool()
def list_sheets(spreadsheet_id: str) -> str:
    """List all sheets/tabs in a spreadsheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.

    Returns:
        JSON list of sheets with their IDs, titles, and dimensions.
    """
    client = get_client()
    result = client.list_sheets(spreadsheet_id)
    return json.dumps(result, indent=2)


@mcp.tool()
def create_sheet(spreadsheet_id: str, title: str, index: int = -1) -> str:
    """Create a new sheet/tab in a spreadsheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        title: Name for the new sheet.
        index: Position for the new sheet (-1 for end).

    Returns:
        JSON with new sheet's ID, title, and index.
    """
    client = get_client()
    idx = None if index < 0 else index
    result = client.create_sheet(spreadsheet_id, title, idx)
    return json.dumps(result, indent=2)


@mcp.tool()
def delete_sheet(spreadsheet_id: str, sheet_id: int) -> str:
    """Delete a sheet from a spreadsheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: The numeric ID of the sheet (from list_sheets).

    Returns:
        Success message.
    """
    client = get_client()
    client.delete_sheet(spreadsheet_id, sheet_id)
    return json.dumps({"success": True, "message": f"Sheet {sheet_id} deleted"})


@mcp.tool()
def rename_sheet(spreadsheet_id: str, sheet_id: int, new_title: str) -> str:
    """Rename a sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: The numeric ID of the sheet.
        new_title: New name for the sheet.

    Returns:
        Success message.
    """
    client = get_client()
    client.rename_sheet(spreadsheet_id, sheet_id, new_title)
    return json.dumps({"success": True, "message": f"Sheet renamed to '{new_title}'"})


@mcp.tool()
def duplicate_sheet(
    spreadsheet_id: str, sheet_id: int, new_title: str, insert_index: int = -1
) -> str:
    """Duplicate/copy a sheet within the same spreadsheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: The numeric ID of the sheet to duplicate.
        new_title: Name for the duplicated sheet.
        insert_index: Position for the copy (-1 for end).

    Returns:
        JSON with new sheet's ID, title, and index.
    """
    client = get_client()
    idx = None if insert_index < 0 else insert_index
    result = client.duplicate_sheet(spreadsheet_id, sheet_id, new_title, idx)
    return json.dumps(result, indent=2)


# =============================================================================
# Cell Operations Tools
# =============================================================================


@mcp.tool()
def read_cells(spreadsheet_id: str, range_notation: str) -> str:
    """Read cell values from a spreadsheet range.

    Examples:
        - read_cells("abc123", "Sheet1!A1:D10") → Read 10 rows, 4 columns
        - read_cells("abc123", "Sheet1!A:A") → Read entire column A
        - read_cells("abc123", "A1:B5") → Read from first visible sheet

    Args:
        spreadsheet_id: The ID from the spreadsheet URL.
        range_notation: A1 notation range. Use 'SheetName!A1:B5' format.

    Returns:
        JSON 2D array of cell values. Empty cells may be omitted.
    """
    client = get_client()
    result = client.read_cells(spreadsheet_id, range_notation)
    return json.dumps(result, indent=2)


@mcp.tool()
def write_cells(spreadsheet_id: str, range_notation: str, values: str) -> str:
    """Write values to cells. Supports formulas (start with =).

    Examples:
        - write_cells("abc123", "Sheet1!A1", '[["Name", "Age"], ["Alice", 30]]')
        - write_cells("abc123", "Sheet1!E1", '[["=SUM(A1:D1)"], ["=SUM(A2:D2)"]]')
        - write_cells("abc123", "A1", '[["Header1", "Header2", "Header3"]]')

    Args:
        spreadsheet_id: The ID from the spreadsheet URL.
        range_notation: Starting cell in A1 notation (e.g., "Sheet1!A1").
        values: JSON 2D array. Each inner array is a row.
                Formulas: use "=SUM(A1:A10)", "=AVERAGE(B:B)", etc.

    Returns:
        JSON with updated_cells count and updated_range.
    """
    client = get_client()
    parsed_values = json.loads(values)
    result = client.write_cells(spreadsheet_id, range_notation, parsed_values)
    return json.dumps(result, indent=2)


@mcp.tool()
def batch_read(spreadsheet_id: str, ranges: str) -> str:
    """Read values from multiple ranges at once.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        ranges: Comma-separated A1 notation ranges (e.g., "Sheet1!A1:B5,Sheet1!D1:E5").

    Returns:
        JSON object mapping each range to its values.
    """
    client = get_client()
    range_list = [r.strip() for r in ranges.split(",")]
    result = client.batch_read(spreadsheet_id, range_list)
    return json.dumps(result, indent=2)


@mcp.tool()
def batch_write(spreadsheet_id: str, data: str) -> str:
    """Write values to multiple ranges at once.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        data: JSON array of objects with 'range' and 'values' keys.
              Example: '[{"range": "A1:B2", "values": [[1,2],[3,4]]}, {"range": "D1", "values": [["Hello"]]}]'

    Returns:
        JSON with total cells, rows, columns updated.
    """
    client = get_client()
    parsed_data = json.loads(data)
    result = client.batch_write(spreadsheet_id, parsed_data)
    return json.dumps(result, indent=2)


@mcp.tool()
def clear_cells(spreadsheet_id: str, range_notation: str) -> str:
    """Clear all values in a range (keeps formatting).

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        range_notation: A1 notation range to clear.

    Returns:
        Success message.
    """
    client = get_client()
    client.clear_cells(spreadsheet_id, range_notation)
    return json.dumps({"success": True, "message": f"Cleared range: {range_notation}"})


@mcp.tool()
def append_rows(spreadsheet_id: str, range_notation: str, values: str) -> str:
    """Append rows to the end of data in a sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        range_notation: A1 notation indicating which columns (e.g., "Sheet1!A:D").
        values: JSON 2D array of rows to append.

    Returns:
        JSON with updated range and cells appended.
    """
    client = get_client()
    parsed_values = json.loads(values)
    result = client.append_rows(spreadsheet_id, range_notation, parsed_values)
    return json.dumps(result, indent=2)


# =============================================================================
# Row/Column Operations Tools
# =============================================================================


@mcp.tool()
def insert_rows(
    spreadsheet_id: str, sheet_id: int, start_index: int, num_rows: int
) -> str:
    """Insert empty rows at a specific position.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: The numeric ID of the sheet.
        start_index: Row index to insert at (0-based, so row 1 = index 0).
        num_rows: Number of rows to insert.

    Returns:
        Success message.
    """
    client = get_client()
    client.insert_rows(spreadsheet_id, sheet_id, start_index, num_rows)
    return json.dumps(
        {"success": True, "message": f"Inserted {num_rows} rows at index {start_index}"}
    )


@mcp.tool()
def insert_columns(
    spreadsheet_id: str, sheet_id: int, start_index: int, num_columns: int
) -> str:
    """Insert empty columns at a specific position.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: The numeric ID of the sheet.
        start_index: Column index to insert at (0-based, so column A = index 0).
        num_columns: Number of columns to insert.

    Returns:
        Success message.
    """
    client = get_client()
    client.insert_columns(spreadsheet_id, sheet_id, start_index, num_columns)
    return json.dumps(
        {
            "success": True,
            "message": f"Inserted {num_columns} columns at index {start_index}",
        }
    )


@mcp.tool()
def delete_rows(
    spreadsheet_id: str, sheet_id: int, start_index: int, num_rows: int
) -> str:
    """Delete rows from a sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: The numeric ID of the sheet.
        start_index: Starting row index (0-based).
        num_rows: Number of rows to delete.

    Returns:
        Success message.
    """
    client = get_client()
    client.delete_rows(spreadsheet_id, sheet_id, start_index, num_rows)
    return json.dumps(
        {
            "success": True,
            "message": f"Deleted {num_rows} rows starting at index {start_index}",
        }
    )


@mcp.tool()
def delete_columns(
    spreadsheet_id: str, sheet_id: int, start_index: int, num_columns: int
) -> str:
    """Delete columns from a sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: The numeric ID of the sheet.
        start_index: Starting column index (0-based).
        num_columns: Number of columns to delete.

    Returns:
        Success message.
    """
    client = get_client()
    client.delete_columns(spreadsheet_id, sheet_id, start_index, num_columns)
    return json.dumps(
        {
            "success": True,
            "message": f"Deleted {num_columns} columns starting at index {start_index}",
        }
    )


# =============================================================================
# Formatting Tools
# =============================================================================


@mcp.tool()
def format_cells(
    spreadsheet_id: str,
    sheet_id: int,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
    bold: bool = False,
    italic: bool = False,
    font_size: int = 0,
    font_color: str = "",
    background_color: str = "",
    alignment: str = "",
) -> str:
    """Apply formatting to a range of cells.

    Examples:
        - Format header row bold with blue background:
          format_cells("abc", 0, 0, 1, 0, 5, bold=True, background_color="#4285F4")
        - Center align cells A1:C10:
          format_cells("abc", 0, 0, 10, 0, 3, alignment="CENTER")

    Note: Indices are 0-based. Row 1 in sheets = index 0. Column A = index 0.
    end_row and end_col are EXCLUSIVE (like Python ranges).

    Args:
        spreadsheet_id: The ID from the spreadsheet URL.
        sheet_id: Numeric sheet ID (get from list_sheets or get_spreadsheet_info).
        start_row: Starting row index (0-based, so row 1 = 0).
        end_row: Ending row index (exclusive, so rows 1-5 = start=0, end=5).
        start_col: Starting column index (0-based, A=0, B=1, C=2...).
        end_col: Ending column index (exclusive).
        bold: Make text bold (True/False).
        italic: Make text italic (True/False).
        font_size: Font size in points (0 to skip, typical: 10-14).
        font_color: Hex color for text (e.g., "#FF0000"=red, "#FFFFFF"=white).
        background_color: Hex color for cell background (e.g., "#4285F4"=blue).
        alignment: "LEFT", "CENTER", or "RIGHT".

    Returns:
        Success message.
    """
    client = get_client()

    # Convert hex colors to RGB dicts
    def hex_to_rgb(hex_color: str) -> dict | None:
        if not hex_color:
            return None
        hex_color = hex_color.lstrip("#")
        if len(hex_color) != 6:
            return None
        r = int(hex_color[0:2], 16) / 255
        g = int(hex_color[2:4], 16) / 255
        b = int(hex_color[4:6], 16) / 255
        return {"red": r, "green": g, "blue": b}

    client.format_cells(
        spreadsheet_id=spreadsheet_id,
        sheet_id=sheet_id,
        start_row=start_row,
        end_row=end_row,
        start_col=start_col,
        end_col=end_col,
        bold=bold if bold else None,
        italic=italic if italic else None,
        font_size=font_size if font_size > 0 else None,
        font_color=hex_to_rgb(font_color),
        background_color=hex_to_rgb(background_color),
        horizontal_alignment=alignment if alignment else None,
    )
    return json.dumps({"success": True, "message": "Formatting applied"})


@mcp.tool()
def set_column_width(
    spreadsheet_id: str, sheet_id: int, start_col: int, end_col: int, width: int
) -> str:
    """Set the width of columns.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: The numeric ID of the sheet.
        start_col: Starting column index (0-based, A=0).
        end_col: Ending column index (exclusive).
        width: Width in pixels.

    Returns:
        Success message.
    """
    client = get_client()
    client.set_column_width(spreadsheet_id, sheet_id, start_col, end_col, width)
    return json.dumps({"success": True, "message": f"Column width set to {width}px"})


@mcp.tool()
def merge_cells(
    spreadsheet_id: str,
    sheet_id: int,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
    merge_type: str = "MERGE_ALL",
) -> str:
    """Merge cells in a range.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: The numeric ID of the sheet.
        start_row: Starting row index (0-based).
        end_row: Ending row index (exclusive).
        start_col: Starting column index (0-based).
        end_col: Ending column index (exclusive).
        merge_type: "MERGE_ALL", "MERGE_COLUMNS", or "MERGE_ROWS".

    Returns:
        Success message.
    """
    client = get_client()
    client.merge_cells(
        spreadsheet_id, sheet_id, start_row, end_row, start_col, end_col, merge_type
    )
    return json.dumps({"success": True, "message": "Cells merged"})


# =============================================================================
# Chart Tools
# =============================================================================


@mcp.tool()
def create_chart(
    spreadsheet_id: str,
    sheet_id: int,
    chart_type: str,
    data_range: str = "",
    title: str = "",
    position_row: int = 0,
    position_col: int = 5,
    domain_column: int = 0,
    series_columns: str = ""
) -> str:
    """Create an embedded chart with MULTIPLE SERIES support.

    Examples:
        - Multi-series line chart (auto-detect all columns B-F as series):
          create_chart("abc", 0, "LINE", "Sheet1!A1:F10", "Stock Prices")
          
        - Explicit series columns (B=1, C=2, D=3, E=4 as separate lines):
          create_chart("abc", 0, "LINE", "A1:E10", "Comparison", 0, 7, 0, "1,2,3,4")
          
        - Pie chart (single series from column B):
          create_chart("abc", 0, "PIE", "A1:B5", "Distribution")

    Chart Types:
        - BAR: Horizontal bars (multiple series = grouped bars)
        - COLUMN: Vertical bars (multiple series = grouped columns)
        - LINE: Line graph (multiple series = multiple lines, great for trends)
        - PIE: Circular chart (single series only)
        - AREA: Stacked area chart
        - SCATTER: X-Y scatter plot

    Multi-Series Example:
        Data: | Date | ONGC | BPCL | IOCL | NTPC | CIL |
        Range: "Sheet1!A1:F30"
        domain_column: 0 (Date column as X-axis)
        series_columns: "1,2,3,4,5" (or leave empty to auto-detect)
        Result: 5 lines on one chart, each representing a company

    Args:
        spreadsheet_id: The ID from the spreadsheet URL.
        sheet_id: Numeric sheet ID where chart will be placed.
        chart_type: BAR, LINE, PIE, AREA, COLUMN, or SCATTER.
        data_range: A1 notation of data range (e.g., "Sheet1!A1:F30").
        title: Chart title displayed at top.
        position_row: Row offset for chart placement (0-based).
        position_col: Column offset for chart placement (0=A, 5=F, 7=H).
        domain_column: Column index for X-axis labels (0=first column of range).
        series_columns: Comma-separated indices for series data (e.g., "1,2,3,4,5").
                       Empty = auto-detect all columns after domain_column.

    Returns:
        JSON with chart_id for later reference.
    """
    client = get_client()
    
    # Parse series_columns string to list
    series_list = None
    if series_columns.strip():
        series_list = [int(x.strip()) for x in series_columns.split(",") if x.strip()]
    
    result = client.create_chart(
        spreadsheet_id=spreadsheet_id,
        sheet_id=sheet_id,
        chart_type=chart_type,
        data_range=data_range,
        title=title,
        position_row=position_row,
        position_col=position_col,
        domain_column=domain_column,
        series_columns=series_list
    )
    return json.dumps(result, indent=2)


@mcp.tool()
def list_charts(spreadsheet_id: str) -> str:
    """List all charts in a spreadsheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.

    Returns:
        JSON list of charts with IDs, titles, and positions.
    """
    client = get_client()
    result = client.list_charts(spreadsheet_id)
    return json.dumps(result, indent=2)


@mcp.tool()
def delete_chart(spreadsheet_id: str, chart_id: int) -> str:
    """Delete a chart from the spreadsheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        chart_id: The ID of the chart (from list_charts).

    Returns:
        Success message.
    """
    client = get_client()
    client.delete_chart(spreadsheet_id, chart_id)
    return json.dumps({"success": True, "message": f"Chart {chart_id} deleted"})


# =============================================================================
# Data Operation Tools
# =============================================================================


@mcp.tool()
def sort_range(
    spreadsheet_id: str,
    sheet_id: int,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
    sort_column: int,
    ascending: bool = True,
) -> str:
    """Sort data in a range by a column.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_id: The numeric ID of the sheet.
        start_row: Starting row index (0-based).
        end_row: Ending row index (exclusive).
        start_col: Starting column index (0-based).
        end_col: Ending column index (exclusive).
        sort_column: Column index to sort by (0-based, relative to range).
        ascending: True for A-Z/0-9, False for Z-A/9-0.

    Returns:
        Success message.
    """
    client = get_client()
    client.sort_range(
        spreadsheet_id,
        sheet_id,
        start_row,
        end_row,
        start_col,
        end_col,
        sort_column,
        ascending,
    )
    order = "ascending" if ascending else "descending"
    return json.dumps(
        {"success": True, "message": f"Range sorted by column {sort_column} ({order})"}
    )


@mcp.tool()
def find_replace(
    spreadsheet_id: str,
    find: str,
    replace: str,
    sheet_id: int = -1,
    match_case: bool = False,
    match_entire_cell: bool = False,
) -> str:
    """Find and replace text across the spreadsheet or a specific sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        find: Text to search for.
        replace: Text to replace with.
        sheet_id: Limit to specific sheet (-1 for all sheets).
        match_case: Case-sensitive matching.
        match_entire_cell: Only match if cell equals find text exactly.

    Returns:
        JSON with number of occurrences changed.
    """
    client = get_client()
    sid = None if sheet_id < 0 else sheet_id
    result = client.find_replace(
        spreadsheet_id,
        find,
        replace,
        sheet_id=sid,
        match_case=match_case,
        match_entire_cell=match_entire_cell,
    )
    return json.dumps(result, indent=2)


@mcp.tool()
def get_last_row(spreadsheet_id: str, sheet_name: str, column: str = "A") -> str:
    """Find the last row with data in a column.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        sheet_name: Name of the sheet.
        column: Column letter to check (default "A").

    Returns:
        JSON with the last row number (1-based).
    """
    client = get_client()
    last_row = client.get_last_row(spreadsheet_id, sheet_name, column)
    return json.dumps({"last_row": last_row, "sheet": sheet_name, "column": column})


# =============================================================================
# Sharing Tools
# =============================================================================


@mcp.tool()
def share_spreadsheet(
    spreadsheet_id: str,
    email: str = "",
    role: str = "reader",
    make_public: bool = False,
) -> str:
    """Share a spreadsheet with someone or make it publicly accessible.

    Args:
        spreadsheet_id: The ID of the spreadsheet.
        email: Email address to share with (optional if make_public=True).
        role: Permission level - "reader", "writer", or "commenter".
        make_public: If True, makes spreadsheet accessible to anyone with link.

    Returns:
        JSON with sharing result.
    """
    client = get_client()
    result = client.share_spreadsheet(
        spreadsheet_id=spreadsheet_id,
        email=email if email else None,
        role=role,
        make_public=make_public,
    )
    return json.dumps(result, indent=2)


# =============================================================================
# Main Entry Point
# =============================================================================


def main():
    """Run the MCP server."""
    mcp.run()


if __name__ == "__main__":
    main()
