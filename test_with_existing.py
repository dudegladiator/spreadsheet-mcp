#!/usr/bin/env python3
"""Test all Google Sheets MCP tools using an existing spreadsheet."""

import json
import sys
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
    if len(sys.argv) < 2:
        print("Usage: python test_with_existing.py <spreadsheet_id>")
        print("\nTo get spreadsheet ID from URL:")
        print("  https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit")
        print("                                        ^^^^^^^^^^^^^^")
        print("\nMake sure you've shared the spreadsheet with your service account email!")
        sys.exit(1)
    
    spreadsheet_id = sys.argv[1]
    
    print(f"\n{BOLD}=== Google Sheets MCP Server - Full Test Suite ==={RESET}")
    print(f"Testing with spreadsheet: {spreadsheet_id}\n")
    
    client = get_client()
    sheet_id = None
    test_sheet_id = None
    chart_id = None
    passed = 0
    failed = 0
    
    try:
        # =====================================================================
        # 1. SPREADSHEET INFO
        # =====================================================================
        print(f"\n{BOLD}1. Spreadsheet Management{RESET}")
        
        # Test: get_spreadsheet_info
        try:
            result = client.get_spreadsheet_info(spreadsheet_id)
            sheet_id = result["sheets"][0]["sheet_id"]
            print(f"   Title: {result['title']}")
            print(f"   Sheets: {[s['title'] for s in result['sheets']]}")
            log_test("get_spreadsheet_info", "PASS")
            passed += 1
        except Exception as e:
            log_test("get_spreadsheet_info", "FAIL", str(e))
            failed += 1
            print(f"\n{RED}Cannot access spreadsheet. Make sure you shared it with your service account email.{RESET}")
            print("Find your service account email in credentials/service-account.json (client_email field)")
            return
        
        # =====================================================================
        # 2. SHEET MANAGEMENT
        # =====================================================================
        print(f"\n{BOLD}2. Sheet Management{RESET}")
        
        # Test: list_sheets
        try:
            result = client.list_sheets(spreadsheet_id)
            log_test("list_sheets", "PASS")
            passed += 1
        except Exception as e:
            log_test("list_sheets", "FAIL", str(e))
            failed += 1
        
        # Test: create_sheet
        try:
            result = client.create_sheet(spreadsheet_id, "MCPTestSheet")
            test_sheet_id = result["sheet_id"]
            log_test("create_sheet", "PASS")
            passed += 1
        except Exception as e:
            log_test("create_sheet", "FAIL", str(e))
            failed += 1
            # Try to find existing test sheet
            try:
                sheets = client.list_sheets(spreadsheet_id)
                for s in sheets:
                    if "MCPTest" in s["title"]:
                        test_sheet_id = s["sheet_id"]
                        break
            except:
                pass
        
        time.sleep(0.5)
        
        # Test: rename_sheet
        if test_sheet_id:
            try:
                client.rename_sheet(spreadsheet_id, test_sheet_id, "MCPTestRenamed")
                log_test("rename_sheet", "PASS")
                passed += 1
            except Exception as e:
                log_test("rename_sheet", "FAIL", str(e))
                failed += 1
        
        # Test: duplicate_sheet
        dup_sheet_id = None
        try:
            result = client.duplicate_sheet(spreadsheet_id, sheet_id, "DuplicatedSheet")
            dup_sheet_id = result["sheet_id"]
            log_test("duplicate_sheet", "PASS")
            passed += 1
        except Exception as e:
            log_test("duplicate_sheet", "FAIL", str(e))
            failed += 1
        
        time.sleep(0.5)
        
        # Test: delete_sheet (delete the duplicated one)
        if dup_sheet_id:
            try:
                client.delete_sheet(spreadsheet_id, dup_sheet_id)
                log_test("delete_sheet", "PASS")
                passed += 1
            except Exception as e:
                log_test("delete_sheet", "FAIL", str(e))
                failed += 1
        
        # =====================================================================
        # 3. CELL OPERATIONS
        # =====================================================================
        print(f"\n{BOLD}3. Cell Operations{RESET}")
        
        work_sheet = test_sheet_id if test_sheet_id else sheet_id
        sheet_name = "MCPTestRenamed" if test_sheet_id else "Sheet1"
        
        # Test: write_cells
        try:
            result = client.write_cells(
                spreadsheet_id, 
                f"{sheet_name}!A1",
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
            passed += 1
        except Exception as e:
            log_test("write_cells", "FAIL", str(e))
            failed += 1
        
        time.sleep(0.5)
        
        # Test: read_cells
        try:
            result = client.read_cells(spreadsheet_id, f"{sheet_name}!A1:C5")
            assert len(result) == 5
            assert result[0][0] == "Name"
            log_test("read_cells", "PASS")
            passed += 1
        except Exception as e:
            log_test("read_cells", "FAIL", str(e))
            failed += 1
        
        # Test: write formula
        try:
            result = client.write_cells(
                spreadsheet_id,
                f"{sheet_name}!D1",
                [
                    ["Total"],
                    ["=B2+C2"],
                    ["=B3+C3"],
                    ["=B4+C4"],
                    ["=B5+C5"]
                ]
            )
            log_test("write_cells (formulas)", "PASS")
            passed += 1
        except Exception as e:
            log_test("write_cells (formulas)", "FAIL", str(e))
            failed += 1
        
        time.sleep(0.5)
        
        # Verify formula results
        try:
            result = client.read_cells(spreadsheet_id, f"{sheet_name}!D2:D5")
            # Should have computed values (120, 117, 114, 116)
            print(f"   Formula results: {[row[0] for row in result]}")
        except:
            pass
        
        # Test: batch_read
        try:
            result = client.batch_read(
                spreadsheet_id,
                [f"{sheet_name}!A1:A5", f"{sheet_name}!C1:C5"]
            )
            assert len(result) == 2
            log_test("batch_read", "PASS")
            passed += 1
        except Exception as e:
            log_test("batch_read", "FAIL", str(e))
            failed += 1
        
        # Test: batch_write
        try:
            result = client.batch_write(
                spreadsheet_id,
                [
                    {"range": f"{sheet_name}!E1", "values": [["Category"]]},
                    {"range": f"{sheet_name}!E2:E5", "values": [["A"], ["B"], ["A"], ["B"]]}
                ]
            )
            assert result["total_updated_cells"] == 5
            log_test("batch_write", "PASS")
            passed += 1
        except Exception as e:
            log_test("batch_write", "FAIL", str(e))
            failed += 1
        
        # Test: append_rows
        try:
            result = client.append_rows(
                spreadsheet_id,
                f"{sheet_name}!A:E",
                [["Eve", 26, 90, "=B6+C6", "A"]]
            )
            log_test("append_rows", "PASS")
            passed += 1
        except Exception as e:
            log_test("append_rows", "FAIL", str(e))
            failed += 1
        
        time.sleep(0.5)
        
        # Test: get_last_row
        try:
            result = client.get_last_row(spreadsheet_id, sheet_name, "A")
            assert result >= 5
            log_test("get_last_row", "PASS")
            passed += 1
        except Exception as e:
            log_test("get_last_row", "FAIL", str(e))
            failed += 1
        
        # Test: clear_cells
        try:
            client.write_cells(spreadsheet_id, f"{sheet_name}!G1:G3", [["X"], ["Y"], ["Z"]])
            client.clear_cells(spreadsheet_id, f"{sheet_name}!G1:G3")
            result = client.read_cells(spreadsheet_id, f"{sheet_name}!G1:G3")
            assert result == []
            log_test("clear_cells", "PASS")
            passed += 1
        except Exception as e:
            log_test("clear_cells", "FAIL", str(e))
            failed += 1
        
        # =====================================================================
        # 4. ROW/COLUMN OPERATIONS
        # =====================================================================
        print(f"\n{BOLD}4. Row/Column Operations{RESET}")
        
        # Test: insert_rows
        try:
            client.insert_rows(spreadsheet_id, work_sheet, 2, 1)
            log_test("insert_rows", "PASS")
            passed += 1
        except Exception as e:
            log_test("insert_rows", "FAIL", str(e))
            failed += 1
        
        time.sleep(0.5)
        
        # Test: delete_rows
        try:
            client.delete_rows(spreadsheet_id, work_sheet, 2, 1)
            log_test("delete_rows", "PASS")
            passed += 1
        except Exception as e:
            log_test("delete_rows", "FAIL", str(e))
            failed += 1
        
        # Test: insert_columns
        try:
            client.insert_columns(spreadsheet_id, work_sheet, 5, 1)
            log_test("insert_columns", "PASS")
            passed += 1
        except Exception as e:
            log_test("insert_columns", "FAIL", str(e))
            failed += 1
        
        time.sleep(0.5)
        
        # Test: delete_columns
        try:
            client.delete_columns(spreadsheet_id, work_sheet, 5, 1)
            log_test("delete_columns", "PASS")
            passed += 1
        except Exception as e:
            log_test("delete_columns", "FAIL", str(e))
            failed += 1
        
        # =====================================================================
        # 5. FORMATTING
        # =====================================================================
        print(f"\n{BOLD}5. Formatting{RESET}")
        
        # Test: format_cells (header bold with background)
        try:
            client.format_cells(
                spreadsheet_id=spreadsheet_id,
                sheet_id=work_sheet,
                start_row=0,
                end_row=1,
                start_col=0,
                end_col=5,
                bold=True,
                background_color={"red": 0.26, "green": 0.52, "blue": 0.96},
                font_color={"red": 1, "green": 1, "blue": 1}
            )
            log_test("format_cells (bold + colors)", "PASS")
            passed += 1
        except Exception as e:
            log_test("format_cells", "FAIL", str(e))
            failed += 1
        
        time.sleep(0.5)
        
        # Test: set_column_width
        try:
            client.set_column_width(spreadsheet_id, work_sheet, 0, 1, 150)
            log_test("set_column_width", "PASS")
            passed += 1
        except Exception as e:
            log_test("set_column_width", "FAIL", str(e))
            failed += 1
        
        # Test: merge_cells
        try:
            client.write_cells(spreadsheet_id, f"{sheet_name}!A10", [["Merged Title"]])
            client.merge_cells(spreadsheet_id, work_sheet, 9, 10, 0, 3)
            log_test("merge_cells", "PASS")
            passed += 1
        except Exception as e:
            log_test("merge_cells", "FAIL", str(e))
            failed += 1
        
        # =====================================================================
        # 6. DATA OPERATIONS
        # =====================================================================
        print(f"\n{BOLD}6. Data Operations{RESET}")
        
        # Test: sort_range
        try:
            client.sort_range(
                spreadsheet_id=spreadsheet_id,
                sheet_id=work_sheet,
                start_row=1,
                end_row=6,
                start_col=0,
                end_col=5,
                sort_column=1,
                ascending=True
            )
            log_test("sort_range", "PASS")
            passed += 1
        except Exception as e:
            log_test("sort_range", "FAIL", str(e))
            failed += 1
        
        time.sleep(0.5)
        
        # Test: find_replace
        try:
            result = client.find_replace(
                spreadsheet_id=spreadsheet_id,
                find="A",
                replace="Category A",
                sheet_id=work_sheet
            )
            log_test("find_replace", "PASS")
            passed += 1
            if result['occurrences_changed'] > 0:
                print(f"   Replaced {result['occurrences_changed']} occurrences")
        except Exception as e:
            log_test("find_replace", "FAIL", str(e))
            failed += 1
        
        # =====================================================================
        # 7. CHARTS
        # =====================================================================
        print(f"\n{BOLD}7. Charts{RESET}")
        
        # Test: create_chart
        try:
            result = client.create_chart(
                spreadsheet_id=spreadsheet_id,
                sheet_id=work_sheet,
                chart_type="COLUMN",
                data_range=f"{sheet_name}!A1:C6",
                title="Age vs Score",
                position_row=0,
                position_col=7
            )
            chart_id = result["chart_id"]
            log_test("create_chart", "PASS")
            passed += 1
        except Exception as e:
            log_test("create_chart", "FAIL", str(e))
            failed += 1
        
        time.sleep(0.5)
        
        # Test: list_charts
        try:
            result = client.list_charts(spreadsheet_id)
            log_test("list_charts", "PASS")
            passed += 1
            print(f"   Found {len(result)} chart(s)")
        except Exception as e:
            log_test("list_charts", "FAIL", str(e))
            failed += 1
        
        # Test: delete_chart
        if chart_id:
            try:
                client.delete_chart(spreadsheet_id, chart_id)
                log_test("delete_chart", "PASS")
                passed += 1
            except Exception as e:
                log_test("delete_chart", "FAIL", str(e))
                failed += 1
        
        # =====================================================================
        # SUMMARY
        # =====================================================================
        print(f"\n{BOLD}=== Test Summary ==={RESET}")
        print(f"{GREEN}Passed: {passed}{RESET}")
        print(f"{RED}Failed: {failed}{RESET}")
        print(f"\nSpreadsheet URL: https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit")
        print("\nOpen the spreadsheet to see the results!")
        
        # Cleanup option
        if test_sheet_id:
            print(f"\n{YELLOW}Note: Test sheet 'MCPTestRenamed' was created. Delete it manually if needed.{RESET}")
        
    except Exception as e:
        print(f"\n{RED}Unexpected error: {e}{RESET}")
        raise


if __name__ == "__main__":
    main()
