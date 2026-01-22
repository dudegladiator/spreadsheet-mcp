# Google Sheets MCP Server

A Model Context Protocol (MCP) server providing full Google Sheets functionality. Read, write, format cells, create charts, use formulas, and manage spreadsheets directly from Claude, Cursor, or any MCP-compatible client.

## Why I Built This

I created this MCP server for my college assignments, where I frequently work with data in Google Sheets — mostly financial sheets. Tasks like creating charts, writing formulas, formatting tables, and organizing data were repetitive and time-consuming. With this tool, I can automate all of that and focus on what actually matters — logical thinking and analysis — instead of doing the same manual work over and over again. :)

## Features

| Category | Capabilities |
|----------|-------------|
| **Cell Operations** | Read, write, batch operations, append rows, clear cells |
| **Formulas** | Full support (`=SUM()`, `=VLOOKUP()`, `=IF()`, etc.) |
| **Sheet Management** | Create, delete, rename, duplicate sheets |
| **Formatting** | Bold, italic, colors, alignment, column widths, merge cells |
| **Charts** | Bar, Line, Pie, Column, Area, Scatter charts |
| **Data Operations** | Sort, find/replace, get last row |
| **Sharing** | Share with users or make public |

## Quick Start

### 1. Install

```bash
cd /path/to/spreadsheet-mcp
uv sync
```

### 2. Google Cloud Setup (Free, One-Time)

1. **Create Project**: Go to [Google Cloud Console](https://console.cloud.google.com) → New Project
2. **Enable APIs**:
   - [Enable Sheets API](https://console.cloud.google.com/apis/library/sheets.googleapis.com)
   - [Enable Drive API](https://console.cloud.google.com/apis/library/drive.googleapis.com)
3. **Create Service Account**: 
   - Go to APIs & Services → Credentials → Create Credentials → Service Account
   - Download JSON key → Save as `credentials/service-account.json`
4. **Share Spreadsheets**: Share your sheets with the service account email (found in JSON)

### 3. Run

```bash
uv run spreadsheet-mcp
```

## MCP Client Configuration

### Claude Desktop

Add to your Claude Desktop config file:
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "google-sheets": {
      "command": "uv",
      "args": [
        "--directory",
        "/path/to/spreadsheet-mcp",
        "run",
        "spreadsheet-mcp"
      ]
    }
  }
}
```

### Claude Code (CLI)

Add to your project's `.mcp.json` file:

```json
{
  "mcpServers": {
    "google-sheets": {
      "command": "uv",
      "args": [
        "--directory",
        "/path/to/spreadsheet-mcp",
        "run",
        "spreadsheet-mcp"
      ]
    }
  }
}
```

### Cursor

Add to Cursor's MCP settings (Settings → MCP → Add Server):

```json
{
  "mcpServers": {
    "google-sheets": {
      "command": "uv",
      "args": [
        "--directory",
        "/path/to/spreadsheet-mcp",
        "run",
        "spreadsheet-mcp"
      ]
    }
  }
}
```

Or add to `~/.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "google-sheets": {
      "command": "uv",
      "args": [
        "--directory",
        "/path/to/spreadsheet-mcp",
        "run",
        "spreadsheet-mcp"
      ]
    }
  }
}
```

### Gemini CLI

For Gemini CLI with MCP support, add to your MCP configuration:

```json
{
  "mcpServers": {
    "google-sheets": {
      "command": "uv",
      "args": [
        "--directory",
        "/path/to/spreadsheet-mcp",
        "run",
        "spreadsheet-mcp"
      ]
    }
  }
}
```

### Using with npx (Alternative)

If you prefer npx over uv:

```json
{
  "mcpServers": {
    "google-sheets": {
      "command": "npx",
      "args": [
        "-y",
        "uv",
        "--directory",
        "/path/to/spreadsheet-mcp",
        "run",
        "spreadsheet-mcp"
      ]
    }
  }
}
```

**Note:** Replace `/path/to/spreadsheet-mcp` with the actual path where you cloned this repository.

## Available Tools (27 Total)

### Spreadsheet & Sheet Management

| Tool | Description | Example |
|------|-------------|---------|
| `get_spreadsheet_info` | Get metadata and sheet list | "Show me info about spreadsheet abc123" |
| `create_spreadsheet` | Create new spreadsheet | "Create a spreadsheet called 'Budget 2024'" |
| `list_sheets` | List all sheets/tabs | "What sheets are in this spreadsheet?" |
| `create_sheet` | Add a new sheet | "Add a sheet called 'Summary'" |
| `delete_sheet` | Remove a sheet | "Delete the sheet with ID 12345" |
| `rename_sheet` | Rename a sheet | "Rename sheet 0 to 'Data'" |
| `duplicate_sheet` | Copy a sheet | "Duplicate the main sheet" |

### Cell Operations

| Tool | Description | Example |
|------|-------------|---------|
| `read_cells` | Read values from range | "Read cells A1 to D10 from Sheet1" |
| `write_cells` | Write values (supports formulas) | "Write headers Name, Age, Score to A1" |
| `batch_read` | Read multiple ranges at once | "Read A1:A10 and C1:C10" |
| `batch_write` | Write to multiple ranges | "Write data to multiple locations" |
| `append_rows` | Add rows at end of data | "Append these new records" |
| `clear_cells` | Clear a range | "Clear cells B2:D10" |
| `get_last_row` | Find last row with data | "What's the last row in column A?" |

### Row/Column Operations

| Tool | Description | Example |
|------|-------------|---------|
| `insert_rows` | Insert empty rows | "Insert 5 rows at row 10" |
| `insert_columns` | Insert empty columns | "Insert 2 columns at column C" |
| `delete_rows` | Delete rows | "Delete rows 5-10" |
| `delete_columns` | Delete columns | "Delete column B" |

### Formatting

| Tool | Description | Example |
|------|-------------|---------|
| `format_cells` | Apply formatting | "Make header row bold with blue background" |
| `set_column_width` | Adjust column width | "Set column A width to 200 pixels" |
| `merge_cells` | Merge cell range | "Merge cells A1:C1 for title" |

### Charts

| Tool | Description | Example |
|------|-------------|---------|
| `create_chart` | Create embedded chart | "Create a bar chart from sales data" |
| `list_charts` | List all charts | "What charts exist in this spreadsheet?" |
| `delete_chart` | Remove a chart | "Delete chart 123456" |

### Data Operations

| Tool | Description | Example |
|------|-------------|---------|
| `sort_range` | Sort data by column | "Sort data by column B ascending" |
| `find_replace` | Find and replace text | "Replace 'N/A' with '0' everywhere" |

### Sharing

| Tool | Description | Example |
|------|-------------|---------|
| `share_spreadsheet` | Share or make public | "Make this spreadsheet public" |

## Usage Examples

### Example 1: Create a Sales Report

```
User: Create a sales report in spreadsheet abc123

Claude will:
1. get_spreadsheet_info("abc123") - Check existing sheets
2. create_sheet("abc123", "Sales Report") - Create new sheet
3. write_cells("abc123", "Sales Report!A1", '[["Product", "Q1", "Q2", "Q3", "Q4", "Total"]]')
4. write_cells("abc123", "Sales Report!A2", '[["Widget", 100, 150, 200, 180, "=SUM(B2:E2)"]]')
5. format_cells(..., bold=True, background_color="#4285F4") - Format header
6. create_chart(..., "COLUMN", "Quarterly Sales") - Add chart
```

### Example 2: Analyze Existing Data

```
User: Summarize the data in Sheet1

Claude will:
1. read_cells("abc123", "Sheet1!A1:Z1") - Read headers
2. get_last_row("abc123", "Sheet1", "A") - Find data extent
3. read_cells("abc123", "Sheet1!A1:D100") - Read all data
4. Provide summary and insights
```

### Example 3: Format a Table

```
User: Make the spreadsheet look professional

Claude will:
1. format_cells(..., bold=True, background_color="#1a73e8", font_color="#FFFFFF") - Header
2. set_column_width(..., 150) - Adjust widths
3. format_cells(..., alignment="CENTER") - Center data
```

## A1 Notation Reference

| Notation | Description |
|----------|-------------|
| `Sheet1!A1:D10` | Cells A1 to D10 in Sheet1 |
| `Sheet1!A:A` | Entire column A in Sheet1 |
| `Sheet1!1:5` | Rows 1-5 in Sheet1 |
| `A1:B10` | Range in first visible sheet |
| `'My Sheet'!A1` | Sheet name with spaces (use quotes) |

## Index Reference

- **Rows**: 0-based in formatting (Row 1 = index 0)
- **Columns**: 0-based (A=0, B=1, C=2, D=3...)
- **sheet_id**: Numeric ID from `list_sheets` or `get_spreadsheet_info`

## Common Formulas

```
=SUM(A1:A10)           Sum values
=AVERAGE(B:B)          Average entire column
=COUNT(A:A)            Count numbers
=COUNTA(A:A)           Count non-empty cells
=MAX(A1:A100)          Maximum value
=MIN(A1:A100)          Minimum value
=VLOOKUP(A1,B:C,2,0)   Lookup value
=IF(A1>100,"High","Low") Conditional
=CONCATENATE(A1," ",B1)  Join text
=TODAY()               Current date
=NOW()                 Current date/time
```

## Troubleshooting

| Error | Solution |
|-------|----------|
| "Service account credentials not found" | Check `credentials/service-account.json` exists |
| "403 Forbidden" | Share spreadsheet with service account email |
| "API not enabled" | Enable Sheets & Drive APIs in Google Cloud Console |
| "Quota exceeded" | Wait a minute, you're hitting rate limits |

## File Structure

```
spreadsheet-mcp/
├── pyproject.toml
├── README.md
├── credentials/
│   └── service-account.json  (your key, gitignored)
└── src/spreadsheet_mcp/
    ├── __init__.py
    ├── auth.py              # Google authentication
    ├── sheets_client.py     # API wrapper (1000+ lines)
    └── server.py            # MCP server (27 tools)
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Author

- [@dudegladiator](https://github.com/dudegladiator)

## License

MIT
