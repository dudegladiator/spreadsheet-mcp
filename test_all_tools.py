#!/usr/bin/env python3
"""Test all Google Sheets MCP tools."""

import json
import time
from src.spreadsheet_mcp.sheets_client import get_client

# Colors for output
GREEN = "\033[92m"
RED = "\033[91m"
YELLOW = "\033[93m"
RESET = "\033[0m"
BOLD = "\033[1m"

def log_test(name: str, status: str, result: str = ""):
    if status == "PASS":
        print(f"{GREEN}✓ {name}{RESET}")
    elif status == "FAIL":
        print(f"{RED}✗ {name}: {result}{RESET}")
    else:
        print(f"{YELLOW}○ {name}: {result}{RESET}")

def main():
    print(f"\n{BOLD}=== Google Sheets MCP Server - Full Test Suite ==={RESET}\n")
    
    client = get_client()
    spreadsheet_id = None
    sheet_id = None
    second_sheet_id = None
    chart_id = None
    
    try:
        # =====================================================================
        # 1. SPREADSHEET MANAGEMENT
        # =====================================================================
        print(f"\n{BOLD}1. Spreadsheet Management{RESET}")
        
        # Test: create_spreadsheet
        try:
            result = client.create_spreadsheet(
                "MCP Test Spreadsheet", 
                ["TestSheet1", "TestSheet2"]
            )
            spreadsheet_id = result["spreadsheet_id"]
            log_test("create_spreadsheet", "PASS")
            print(f"   Created: {result['url']}")
        except Exception as e:
            log_test("create_spreadsheet", "FAIL", str(e))
            return
        
        time.sleep(1)  # Rate limiting
        
        # Test: get_spreadsheet_info
        try:
            result = client.get_spreadsheet_info(spreadsheet_id)
            assert result["title"] == "MCP Test Spreadsheet"
            assert len(result["sheets"]) == 2
            sheet_id = result["sheets"][0]["sheet_id"]
            second_sheet_id = result["sheets"][1]["sheet_id"]
            log_test("get_spreadsheet_info", "PASS")
        except Exception as e:
            log_test("get_spreadsheet_info", "FAIL", str(e))
        
        # =====================================================================
        # 2. SHEET MANAGEMENT
        # =====================================================================
        print(f"\n{BOLD}2. Sheet Management{RESET}")
        
        # Test: list_sheets
        try:
            result = client.list_sheets(spreadsheet_id)
            assert len(result) == 2
            log_test("list_sheets", "PASS")
        except Exception as e:
            log_test("list_sheets", "FAIL", str(e))
        
        # Test: create_sheet
        try:
            result = client.create_sheet(spreadsheet_id, "NewSheet")
            new_sheet_id = result["sheet_id"]
            assert result["title"] == "NewSheet"
            log_test("create_sheet", "PASS")
        except Exception as e:
            log_test("create_sheet", "FAIL", str(e))
        
        time.sleep(0.5)
        
        # Test: rename_sheet
        try:
            client.rename_sheet(spreadsheet_id, new_sheet_id, "RenamedSheet")
            log_test("rename_sheet", "PASS")
        except Exception as e:
            log_test("rename_sheet", "FAIL", str(e))
        
        # Test: duplicate_sheet
        try:
            result = client.duplicate_sheet(spreadsheet_id, sheet_id, "DuplicatedSheet")
            dup_sheet_id = result["sheet_id"]
            log_test("duplicate_sheet", "PASS")
        except Exception as e:
            log_test("duplicate_sheet", "FAIL", str(e))
        
        time.sleep(0.5)
        
        # Test: delete_sheet
        try:
            client.delete_sheet(spreadsheet_id, dup_sheet_id)
            log_test("delete_sheet", "PASS")
        except Exception as e:
            log_test("delete_sheet", "FAIL", str(e))
        
        # =====================================================================
        # 3. CELL OPERATIONS
        # =====================================================================
        print(f"\n{BOLD}3. Cell Operations{RESET}")
        
        # Test: write_cells
        try:
            result = client.write_cells(
                spreadsheet_id, 
                "TestSheet1!A1",
                [
                    ["Name", "Age", "Score"],
                    ["Alice", 25, 95],
                    ["Bob", 30, 87],
                    ["Charlie", 22, 92],
                    ["Diana", 28, 88]
                ]
            )
            assert result["updated_cells"] == 15
            log_test("write_cells", "PASS")
        except Exception as e:
            log_test("write_cells", "FAIL", str(e))
        
        time.sleep(0.5)
        
        # Test: read_cells
        try:
            result = client.read_cells(spreadsheet_id, "TestSheet1!A1:C5")
            assert len(result) == 5
            assert result[0][0] == "Name"
            assert result[1][0] == "Alice"
            log_test("read_cells", "PASS")
        except Exception as e:
            log_test("read_cells", "FAIL", str(e))
        
        # Test: write formula
        try:
            result = client.write_cells(
                spreadsheet_id,
                "TestSheet1!D1",
                [
                    ["Total"],
                    ["=B2+C2"],
                    ["=B3+C3"],
                    ["=B4+C4"],
                    ["=B5+C5"]
                ]
            )
            log_test("write_cells (formulas)", "PASS")
        except Exception as e:
            log_test("write_cells (formulas)", "FAIL", str(e))
        
        time.sleep(0.5)
        
        # Test: batch_read
        try:
            result = client.batch_read(
                spreadsheet_id,
                ["TestSheet1!A1:A5", "TestSheet1!C1:C5"]
            )
            assert len(result) == 2
            log_test("batch_read", "PASS")
        except Exception as e:
            log_test("batch_read", "FAIL", str(e))
        
        # Test: batch_write
        try:
            result = client.batch_write(
                spreadsheet_id,
                [
                    {"range": "TestSheet1!E1", "values": [["Category"]]},
                    {"range": "TestSheet1!E2:E5", "values": [["A"], ["B"], ["A"], ["B"]]}
                ]
            )
            assert result["total_updated_cells"] == 5
            log_test("batch_write", "PASS")
        except Exception as e:
            log_test("batch_write", "FAIL", str(e))
        
        # Test: append_rows
        try:
            result = client.append_rows(
                spreadsheet_id,
                "TestSheet1!A:E",
                [["Eve", 26, 90, "=B6+C6", "A"]]
            )
            log_test("append_rows", "PASS")
        except Exception as e:
            log_test("append_rows", "FAIL", str(e))
        
        time.sleep(0.5)
        
        # Test: get_last_row
        try:
            result = client.get_last_row(spreadsheet_id, "TestSheet1", "A")
            assert result == 6  # Header + 5 data rows
            log_test("get_last_row", "PASS")
        except Exception as e:
            log_test("get_last_row", "FAIL", str(e))
        
        # Test: clear_cells
        try:
            # Write some data to clear
            client.write_cells(spreadsheet_id, "TestSheet1!G1:G3", [["X"], ["Y"], ["Z"]])
            client.clear_cells(spreadsheet_id, "TestSheet1!G1:G3")
            result = client.read_cells(spreadsheet_id, "TestSheet1!G1:G3")
            assert result == []
            log_test("clear_cells", "PASS")
        except Exception as e:
            log_test("clear_cells", "FAIL", str(e))
        
        # =====================================================================
        # 4. ROW/COLUMN OPERATIONS
        # =====================================================================
        print(f"\n{BOLD}4. Row/Column Operations{RESET}")
        
        # Test: insert_rows
        try:
            client.insert_rows(spreadsheet_id, sheet_id, 2, 1)
            log_test("insert_rows", "PASS")
        except Exception as e:
            log_test("insert_rows", "FAIL", str(e))
        
        time.sleep(0.5)
        
        # Test: delete_rows
        try:
            client.delete_rows(spreadsheet_id, sheet_id, 2, 1)
            log_test("delete_rows", "PASS")
        except Exception as e:
            log_test("delete_rows", "FAIL", str(e))
        
        # Test: insert_columns
        try:
            client.insert_columns(spreadsheet_id, sheet_id, 5, 1)
            log_test("insert_columns", "PASS")
        except Exception as e:
            log_test("insert_columns", "FAIL", str(e))
        
        time.sleep(0.5)
        
        # Test: delete_columns
        try:
            client.delete_columns(spreadsheet_id, sheet_id, 5, 1)
            log_test("delete_columns", "PASS")
        except Exception as e:
            log_test("delete_columns", "FAIL", str(e))
        
        # =====================================================================
        # 5. FORMATTING
        # =====================================================================
        print(f"\n{BOLD}5. Formatting{RESET}")
        
        # Test: format_cells (header bold with background)
        try:
            client.format_cells(
                spreadsheet_id=spreadsheet_id,
                sheet_id=sheet_id,
                start_row=0,
                end_row=1,
                start_col=0,
                end_col=5,
                bold=True,
                background_color={"red": 0.26, "green": 0.52, "blue": 0.96},  # Google Blue
                font_color={"red": 1, "green": 1, "blue": 1}  # White text
            )
            log_test("format_cells (bold + colors)", "PASS")
        except Exception as e:
            log_test("format_cells", "FAIL", str(e))
        
        time.sleep(0.5)
        
        # Test: set_column_width
        try:
            client.set_column_width(spreadsheet_id, sheet_id, 0, 1, 150)
            log_test("set_column_width", "PASS")
        except Exception as e:
            log_test("set_column_width", "FAIL", str(e))
        
        # Test: merge_cells
        try:
            # Write a title to merge
            client.write_cells(spreadsheet_id, "TestSheet1!A10", [["Merged Title"]])
            client.merge_cells(spreadsheet_id, sheet_id, 9, 10, 0, 3)
            log_test("merge_cells", "PASS")
        except Exception as e:
            log_test("merge_cells", "FAIL", str(e))
        
        # =====================================================================
        # 6. DATA OPERATIONS
        # =====================================================================
        print(f"\n{BOLD}6. Data Operations{RESET}")
        
        # Test: sort_range
        try:
            # Sort data by Age (column B, index 1)
            client.sort_range(
                spreadsheet_id=spreadsheet_id,
                sheet_id=sheet_id,
                start_row=1,  # Skip header
                end_row=7,
                start_col=0,
                end_col=5,
                sort_column=1,  # Age column
                ascending=True
            )
            log_test("sort_range", "PASS")
        except Exception as e:
            log_test("sort_range", "FAIL", str(e))
        
        time.sleep(0.5)
        
        # Test: find_replace
        try:
            result = client.find_replace(
                spreadsheet_id=spreadsheet_id,
                find="A",
                replace="Category A",
                sheet_id=sheet_id
            )
            log_test("find_replace", "PASS")
            print(f"   Replaced {result['occurrences_changed']} occurrences")
        except Exception as e:
            log_test("find_replace", "FAIL", str(e))
        
        # =====================================================================
        # 7. CHARTS
        # =====================================================================
        print(f"\n{BOLD}7. Charts{RESET}")
        
        # First, add clean data for chart
        try:
            client.write_cells(
                spreadsheet_id,
                "TestSheet2!A1",
                [
                    ["Month", "Sales"],
                    ["Jan", 100],
                    ["Feb", 150],
                    ["Mar", 120],
                    ["Apr", 200],
                    ["May", 180]
                ]
            )
        except Exception as e:
            print(f"   Warning: Could not prepare chart data: {e}")
        
        time.sleep(0.5)
        
        # Test: create_chart
        try:
            result = client.create_chart(
                spreadsheet_id=spreadsheet_id,
                sheet_id=second_sheet_id,
                chart_type="COLUMN",
                title="Monthly Sales",
                position_row=0,
                position_col=4
            )
            chart_id = result["chart_id"]
            log_test("create_chart", "PASS")
        except Exception as e:
            log_test("create_chart", "FAIL", str(e))
        
        time.sleep(0.5)
        
        # Test: list_charts
        try:
            result = client.list_charts(spreadsheet_id)
            assert len(result) >= 1
            log_test("list_charts", "PASS")
            print(f"   Found {len(result)} chart(s)")
        except Exception as e:
            log_test("list_charts", "FAIL", str(e))
        
        # Test: delete_chart
        if chart_id:
            try:
                client.delete_chart(spreadsheet_id, chart_id)
                log_test("delete_chart", "PASS")
            except Exception as e:
                log_test("delete_chart", "FAIL", str(e))
        
        # =====================================================================
        # SUMMARY
        # =====================================================================
        print(f"\n{BOLD}=== Test Complete ==={RESET}")
        print(f"\nSpreadsheet URL: https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit")
        print("\n⚠️  The test spreadsheet was created and is still available.")
        print("   You can view it to verify the operations worked correctly.")
        
    except Exception as e:
        print(f"\n{RED}Unexpected error: {e}{RESET}")
        raise
    finally:
        # Optionally delete the test spreadsheet
        # (leaving it so user can inspect)
        pass


if __name__ == "__main__":
    main()
