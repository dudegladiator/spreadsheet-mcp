"""Google Sheets API Client Wrapper."""

from typing import Any
from googleapiclient.discovery import Resource

from .auth import get_sheets_service, get_drive_service


class SheetsClient:
    """Wrapper class for Google Sheets API operations."""

    def __init__(self):
        self._service: Resource | None = None

    @property
    def service(self) -> Resource:
        """Lazy-load the Sheets API service."""
        if self._service is None:
            self._service = get_sheets_service()
        return self._service

    @property
    def spreadsheets(self):
        """Access the spreadsheets resource."""
        return self.service.spreadsheets()

    @property
    def values(self):
        """Access the spreadsheets.values resource."""
        return self.service.spreadsheets().values()

    # =========================================================================
    # Spreadsheet Management
    # =========================================================================

    def create_spreadsheet(
        self, title: str, sheet_names: list[str] | None = None
    ) -> dict:
        """Create a new spreadsheet.

        Args:
            title: The title of the new spreadsheet.
            sheet_names: Optional list of sheet names to create.

        Returns:
            Spreadsheet metadata including spreadsheetId and URL.
        """
        body: dict[str, Any] = {"properties": {"title": title}}

        if sheet_names:
            body["sheets"] = [{"properties": {"title": name}} for name in sheet_names]

        result = self.spreadsheets.create(body=body).execute()
        return {
            "spreadsheet_id": result["spreadsheetId"],
            "title": result["properties"]["title"],
            "url": result["spreadsheetUrl"],
            "sheets": [s["properties"]["title"] for s in result.get("sheets", [])],
        }

    def get_spreadsheet_info(self, spreadsheet_id: str) -> dict:
        """Get metadata about a spreadsheet.

        Args:
            spreadsheet_id: The ID of the spreadsheet.

        Returns:
            Spreadsheet metadata including title, sheets, and URL.
        """
        result = self.spreadsheets.get(spreadsheetId=spreadsheet_id).execute()
        sheets_info = []
        for sheet in result.get("sheets", []):
            props = sheet["properties"]
            sheets_info.append(
                {
                    "sheet_id": props["sheetId"],
                    "title": props["title"],
                    "index": props["index"],
                    "row_count": props.get("gridProperties", {}).get("rowCount", 0),
                    "column_count": props.get("gridProperties", {}).get(
                        "columnCount", 0
                    ),
                }
            )

        return {
            "spreadsheet_id": result["spreadsheetId"],
            "title": result["properties"]["title"],
            "url": result["spreadsheetUrl"],
            "locale": result["properties"].get("locale", ""),
            "sheets": sheets_info,
        }

    # =========================================================================
    # Sheet Management
    # =========================================================================

    def list_sheets(self, spreadsheet_id: str) -> list[dict]:
        """List all sheets in a spreadsheet.

        Args:
            spreadsheet_id: The ID of the spreadsheet.

        Returns:
            List of sheet information dictionaries.
        """
        info = self.get_spreadsheet_info(spreadsheet_id)
        return info["sheets"]

    def create_sheet(
        self, spreadsheet_id: str, title: str, index: int | None = None
    ) -> dict:
        """Create a new sheet in a spreadsheet.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            title: The title for the new sheet.
            index: Optional position index for the new sheet.

        Returns:
            New sheet properties.
        """
        properties: dict[str, Any] = {"title": title}
        if index is not None:
            properties["index"] = index

        request = {"requests": [{"addSheet": {"properties": properties}}]}

        result = self.spreadsheets.batchUpdate(
            spreadsheetId=spreadsheet_id, body=request
        ).execute()

        new_sheet = result["replies"][0]["addSheet"]["properties"]
        return {
            "sheet_id": new_sheet["sheetId"],
            "title": new_sheet["title"],
            "index": new_sheet["index"],
        }

    def delete_sheet(self, spreadsheet_id: str, sheet_id: int) -> bool:
        """Delete a sheet from a spreadsheet.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            sheet_id: The numeric ID of the sheet to delete.

        Returns:
            True if successful.
        """
        request = {"requests": [{"deleteSheet": {"sheetId": sheet_id}}]}

        self.spreadsheets.batchUpdate(
            spreadsheetId=spreadsheet_id, body=request
        ).execute()
        return True

    def rename_sheet(self, spreadsheet_id: str, sheet_id: int, new_title: str) -> bool:
        """Rename a sheet.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            sheet_id: The numeric ID of the sheet.
            new_title: The new title for the sheet.

        Returns:
            True if successful.
        """
        request = {
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {"sheetId": sheet_id, "title": new_title},
                        "fields": "title",
                    }
                }
            ]
        }

        self.spreadsheets.batchUpdate(
            spreadsheetId=spreadsheet_id, body=request
        ).execute()
        return True

    def duplicate_sheet(
        self,
        spreadsheet_id: str,
        sheet_id: int,
        new_title: str,
        insert_index: int | None = None,
    ) -> dict:
        """Duplicate a sheet within the same spreadsheet.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            sheet_id: The numeric ID of the sheet to duplicate.
            new_title: Title for the duplicated sheet.
            insert_index: Optional position for the new sheet.

        Returns:
            New sheet properties.
        """
        request_body: dict[str, Any] = {
            "sourceSheetId": sheet_id,
            "newSheetName": new_title,
        }
        if insert_index is not None:
            request_body["insertSheetIndex"] = insert_index

        request = {"requests": [{"duplicateSheet": request_body}]}

        result = self.spreadsheets.batchUpdate(
            spreadsheetId=spreadsheet_id, body=request
        ).execute()

        new_sheet = result["replies"][0]["duplicateSheet"]["properties"]
        return {
            "sheet_id": new_sheet["sheetId"],
            "title": new_sheet["title"],
            "index": new_sheet["index"],
        }

    def _get_sheet_id_by_name(self, spreadsheet_id: str, sheet_name: str) -> int:
        """Get sheet ID from sheet name.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            sheet_name: The name/title of the sheet.

        Returns:
            The numeric sheet ID.
        """
        sheets = self.list_sheets(spreadsheet_id)
        for sheet in sheets:
            if sheet["title"] == sheet_name:
                return sheet["sheet_id"]
        raise ValueError(f"Sheet '{sheet_name}' not found in spreadsheet")

    # =========================================================================
    # Cell Operations
    # =========================================================================

    def read_cells(self, spreadsheet_id: str, range_notation: str) -> list[list]:
        """Read values from a range.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            range_notation: A1 notation range (e.g., "Sheet1!A1:D10").

        Returns:
            2D list of cell values.
        """
        result = self.values.get(
            spreadsheetId=spreadsheet_id,
            range=range_notation,
            valueRenderOption="FORMATTED_VALUE",
        ).execute()
        return result.get("values", [])

    def write_cells(
        self,
        spreadsheet_id: str,
        range_notation: str,
        values: list[list],
        value_input_option: str = "USER_ENTERED",
    ) -> dict:
        """Write values to a range.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            range_notation: A1 notation range (e.g., "Sheet1!A1").
            values: 2D list of values to write. Use formulas like "=SUM(A1:A10)".
            value_input_option: How to interpret input (USER_ENTERED or RAW).

        Returns:
            Update result with number of updated cells.
        """
        result = self.values.update(
            spreadsheetId=spreadsheet_id,
            range=range_notation,
            valueInputOption=value_input_option,
            body={"values": values},
        ).execute()

        return {
            "updated_range": result.get("updatedRange", ""),
            "updated_rows": result.get("updatedRows", 0),
            "updated_columns": result.get("updatedColumns", 0),
            "updated_cells": result.get("updatedCells", 0),
        }

    def batch_read(
        self, spreadsheet_id: str, ranges: list[str]
    ) -> dict[str, list[list]]:
        """Read values from multiple ranges.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            ranges: List of A1 notation ranges.

        Returns:
            Dictionary mapping range to values.
        """
        result = self.values.batchGet(
            spreadsheetId=spreadsheet_id,
            ranges=ranges,
            valueRenderOption="FORMATTED_VALUE",
        ).execute()

        output = {}
        for value_range in result.get("valueRanges", []):
            output[value_range["range"]] = value_range.get("values", [])
        return output

    def batch_write(
        self,
        spreadsheet_id: str,
        data: list[dict],
        value_input_option: str = "USER_ENTERED",
    ) -> dict:
        """Write values to multiple ranges.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            data: List of {"range": "A1:B2", "values": [[1, 2], [3, 4]]} dicts.
            value_input_option: How to interpret input (USER_ENTERED or RAW).

        Returns:
            Batch update result.
        """
        result = self.values.batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"valueInputOption": value_input_option, "data": data},
        ).execute()

        return {
            "total_updated_cells": result.get("totalUpdatedCells", 0),
            "total_updated_rows": result.get("totalUpdatedRows", 0),
            "total_updated_columns": result.get("totalUpdatedColumns", 0),
        }

    def clear_cells(self, spreadsheet_id: str, range_notation: str) -> bool:
        """Clear values in a range.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            range_notation: A1 notation range to clear.

        Returns:
            True if successful.
        """
        self.values.clear(spreadsheetId=spreadsheet_id, range=range_notation).execute()
        return True

    def append_rows(
        self,
        spreadsheet_id: str,
        range_notation: str,
        values: list[list],
        value_input_option: str = "USER_ENTERED",
    ) -> dict:
        """Append rows to a sheet.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            range_notation: A1 notation (e.g., "Sheet1!A:A" to append to column A).
            values: 2D list of values to append.
            value_input_option: How to interpret input.

        Returns:
            Append result with updated range.
        """
        result = self.values.append(
            spreadsheetId=spreadsheet_id,
            range=range_notation,
            valueInputOption=value_input_option,
            insertDataOption="INSERT_ROWS",
            body={"values": values},
        ).execute()

        updates = result.get("updates", {})
        return {
            "updated_range": updates.get("updatedRange", ""),
            "updated_rows": updates.get("updatedRows", 0),
            "updated_cells": updates.get("updatedCells", 0),
        }

    # =========================================================================
    # Row/Column Operations
    # =========================================================================

    def insert_rows(
        self, spreadsheet_id: str, sheet_id: int, start_index: int, num_rows: int
    ) -> bool:
        """Insert empty rows.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            sheet_id: The numeric ID of the sheet.
            start_index: Row index to insert at (0-based).
            num_rows: Number of rows to insert.

        Returns:
            True if successful.
        """
        request = {
            "requests": [
                {
                    "insertDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": start_index,
                            "endIndex": start_index + num_rows,
                        },
                        "inheritFromBefore": start_index > 0,
                    }
                }
            ]
        }

        self.spreadsheets.batchUpdate(
            spreadsheetId=spreadsheet_id, body=request
        ).execute()
        return True

    def insert_columns(
        self, spreadsheet_id: str, sheet_id: int, start_index: int, num_columns: int
    ) -> bool:
        """Insert empty columns.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            sheet_id: The numeric ID of the sheet.
            start_index: Column index to insert at (0-based).
            num_columns: Number of columns to insert.

        Returns:
            True if successful.
        """
        request = {
            "requests": [
                {
                    "insertDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": start_index,
                            "endIndex": start_index + num_columns,
                        },
                        "inheritFromBefore": start_index > 0,
                    }
                }
            ]
        }

        self.spreadsheets.batchUpdate(
            spreadsheetId=spreadsheet_id, body=request
        ).execute()
        return True

    def delete_rows(
        self, spreadsheet_id: str, sheet_id: int, start_index: int, num_rows: int
    ) -> bool:
        """Delete rows.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            sheet_id: The numeric ID of the sheet.
            start_index: Starting row index (0-based).
            num_rows: Number of rows to delete.

        Returns:
            True if successful.
        """
        request = {
            "requests": [
                {
                    "deleteDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": start_index,
                            "endIndex": start_index + num_rows,
                        }
                    }
                }
            ]
        }

        self.spreadsheets.batchUpdate(
            spreadsheetId=spreadsheet_id, body=request
        ).execute()
        return True

    def delete_columns(
        self, spreadsheet_id: str, sheet_id: int, start_index: int, num_columns: int
    ) -> bool:
        """Delete columns.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            sheet_id: The numeric ID of the sheet.
            start_index: Starting column index (0-based).
            num_columns: Number of columns to delete.

        Returns:
            True if successful.
        """
        request = {
            "requests": [
                {
                    "deleteDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": start_index,
                            "endIndex": start_index + num_columns,
                        }
                    }
                }
            ]
        }

        self.spreadsheets.batchUpdate(
            spreadsheetId=spreadsheet_id, body=request
        ).execute()
        return True

    # =========================================================================
    # Formatting
    # =========================================================================

    def format_cells(
        self,
        spreadsheet_id: str,
        sheet_id: int,
        start_row: int,
        end_row: int,
        start_col: int,
        end_col: int,
        bold: bool | None = None,
        italic: bool | None = None,
        font_size: int | None = None,
        font_color: dict | None = None,
        background_color: dict | None = None,
        horizontal_alignment: str | None = None,
    ) -> bool:
        """Apply formatting to a range of cells.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            sheet_id: The numeric ID of the sheet.
            start_row: Starting row index (0-based).
            end_row: Ending row index (exclusive).
            start_col: Starting column index (0-based).
            end_col: Ending column index (exclusive).
            bold: Make text bold.
            italic: Make text italic.
            font_size: Font size in points.
            font_color: RGB color dict {"red": 0-1, "green": 0-1, "blue": 0-1}.
            background_color: RGB color dict for cell background.
            horizontal_alignment: "LEFT", "CENTER", or "RIGHT".

        Returns:
            True if successful.
        """
        cell_format: dict[str, Any] = {}
        fields = []

        # Text format
        text_format: dict[str, Any] = {}
        if bold is not None:
            text_format["bold"] = bold
            fields.append("userEnteredFormat.textFormat.bold")
        if italic is not None:
            text_format["italic"] = italic
            fields.append("userEnteredFormat.textFormat.italic")
        if font_size is not None:
            text_format["fontSize"] = font_size
            fields.append("userEnteredFormat.textFormat.fontSize")
        if font_color is not None:
            text_format["foregroundColor"] = font_color
            fields.append("userEnteredFormat.textFormat.foregroundColor")

        if text_format:
            cell_format["textFormat"] = text_format

        # Background color
        if background_color is not None:
            cell_format["backgroundColor"] = background_color
            fields.append("userEnteredFormat.backgroundColor")

        # Alignment
        if horizontal_alignment is not None:
            cell_format["horizontalAlignment"] = horizontal_alignment
            fields.append("userEnteredFormat.horizontalAlignment")

        if not fields:
            return True  # Nothing to format

        request = {
            "requests": [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": start_row,
                            "endRowIndex": end_row,
                            "startColumnIndex": start_col,
                            "endColumnIndex": end_col,
                        },
                        "cell": {"userEnteredFormat": cell_format},
                        "fields": ",".join(fields),
                    }
                }
            ]
        }

        self.spreadsheets.batchUpdate(
            spreadsheetId=spreadsheet_id, body=request
        ).execute()
        return True

    def set_column_width(
        self,
        spreadsheet_id: str,
        sheet_id: int,
        start_col: int,
        end_col: int,
        width: int,
    ) -> bool:
        """Set column width.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            sheet_id: The numeric ID of the sheet.
            start_col: Starting column index (0-based).
            end_col: Ending column index (exclusive).
            width: Width in pixels.

        Returns:
            True if successful.
        """
        request = {
            "requests": [
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": start_col,
                            "endIndex": end_col,
                        },
                        "properties": {"pixelSize": width},
                        "fields": "pixelSize",
                    }
                }
            ]
        }

        self.spreadsheets.batchUpdate(
            spreadsheetId=spreadsheet_id, body=request
        ).execute()
        return True

    def merge_cells(
        self,
        spreadsheet_id: str,
        sheet_id: int,
        start_row: int,
        end_row: int,
        start_col: int,
        end_col: int,
        merge_type: str = "MERGE_ALL",
    ) -> bool:
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
            True if successful.
        """
        request = {
            "requests": [
                {
                    "mergeCells": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": start_row,
                            "endRowIndex": end_row,
                            "startColumnIndex": start_col,
                            "endColumnIndex": end_col,
                        },
                        "mergeType": merge_type,
                    }
                }
            ]
        }

        self.spreadsheets.batchUpdate(
            spreadsheetId=spreadsheet_id, body=request
        ).execute()
        return True

    # =========================================================================
    # Charts
    # =========================================================================

    def create_chart(
        self,
        spreadsheet_id: str,
        sheet_id: int,
        chart_type: str,
        data_range: str,
        title: str = "",
        position_row: int = 0,
        position_col: int = 0,
    ) -> dict:
        """Create an embedded chart.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            sheet_id: The numeric ID of the sheet where chart will be placed.
            chart_type: Type of chart: "BAR", "LINE", "PIE", "AREA", "COLUMN", "SCATTER".
            data_range: A1 notation of data range (e.g., "Sheet1!A1:B10").
            title: Chart title.
            position_row: Row offset for chart position.
            position_col: Column offset for chart position.

        Returns:
            Chart info including chart ID.
        """
        # Parse the data range to get source sheet ID
        # For simplicity, we'll use the same sheet_id for data source

        chart_spec: dict[str, Any] = {
            "title": title,
            "basicChart": {
                "chartType": chart_type.upper(),
                "legendPosition": "BOTTOM_LEGEND",
                "domains": [
                    {
                        "domain": {
                            "sourceRange": {
                                "sources": [
                                    {
                                        "sheetId": sheet_id,
                                        "startRowIndex": 0,
                                        "endRowIndex": 100,
                                        "startColumnIndex": 0,
                                        "endColumnIndex": 1,
                                    }
                                ]
                            }
                        }
                    }
                ],
                "series": [
                    {
                        "series": {
                            "sourceRange": {
                                "sources": [
                                    {
                                        "sheetId": sheet_id,
                                        "startRowIndex": 0,
                                        "endRowIndex": 100,
                                        "startColumnIndex": 1,
                                        "endColumnIndex": 2,
                                    }
                                ]
                            }
                        },
                        "targetAxis": "LEFT_AXIS",
                    }
                ],
                "headerCount": 1,
            },
        }

        # Special handling for PIE charts
        if chart_type.upper() == "PIE":
            chart_spec = {
                "title": title,
                "pieChart": {
                    "legendPosition": "RIGHT_LEGEND",
                    "domain": {
                        "sourceRange": {
                            "sources": [
                                {
                                    "sheetId": sheet_id,
                                    "startRowIndex": 0,
                                    "endRowIndex": 100,
                                    "startColumnIndex": 0,
                                    "endColumnIndex": 1,
                                }
                            ]
                        }
                    },
                    "series": {
                        "sourceRange": {
                            "sources": [
                                {
                                    "sheetId": sheet_id,
                                    "startRowIndex": 0,
                                    "endRowIndex": 100,
                                    "startColumnIndex": 1,
                                    "endColumnIndex": 2,
                                }
                            ]
                        }
                    },
                },
            }

        request = {
            "requests": [
                {
                    "addChart": {
                        "chart": {
                            "spec": chart_spec,
                            "position": {
                                "overlayPosition": {
                                    "anchorCell": {
                                        "sheetId": sheet_id,
                                        "rowIndex": position_row,
                                        "columnIndex": position_col,
                                    },
                                    "offsetXPixels": 0,
                                    "offsetYPixels": 0,
                                    "widthPixels": 600,
                                    "heightPixels": 400,
                                }
                            },
                        }
                    }
                }
            ]
        }

        result = self.spreadsheets.batchUpdate(
            spreadsheetId=spreadsheet_id, body=request
        ).execute()

        chart = result["replies"][0]["addChart"]["chart"]
        return {"chart_id": chart["chartId"], "position": chart.get("position", {})}

    def list_charts(self, spreadsheet_id: str) -> list[dict]:
        """List all charts in a spreadsheet.

        Args:
            spreadsheet_id: The ID of the spreadsheet.

        Returns:
            List of chart information.
        """
        result = self.spreadsheets.get(spreadsheetId=spreadsheet_id).execute()
        charts = []

        for sheet in result.get("sheets", []):
            sheet_title = sheet["properties"]["title"]
            for chart in sheet.get("charts", []):
                charts.append(
                    {
                        "chart_id": chart["chartId"],
                        "sheet": sheet_title,
                        "title": chart.get("spec", {}).get("title", ""),
                        "position": chart.get("position", {}),
                    }
                )

        return charts

    def delete_chart(self, spreadsheet_id: str, chart_id: int) -> bool:
        """Delete a chart.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            chart_id: The ID of the chart to delete.

        Returns:
            True if successful.
        """
        request = {"requests": [{"deleteEmbeddedObject": {"objectId": chart_id}}]}

        self.spreadsheets.batchUpdate(
            spreadsheetId=spreadsheet_id, body=request
        ).execute()
        return True

    # =========================================================================
    # Data Operations
    # =========================================================================

    def sort_range(
        self,
        spreadsheet_id: str,
        sheet_id: int,
        start_row: int,
        end_row: int,
        start_col: int,
        end_col: int,
        sort_column: int,
        ascending: bool = True,
    ) -> bool:
        """Sort data in a range.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            sheet_id: The numeric ID of the sheet.
            start_row: Starting row index (0-based).
            end_row: Ending row index (exclusive).
            start_col: Starting column index (0-based).
            end_col: Ending column index (exclusive).
            sort_column: Column index to sort by (0-based).
            ascending: Sort order (True for ascending).

        Returns:
            True if successful.
        """
        request = {
            "requests": [
                {
                    "sortRange": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": start_row,
                            "endRowIndex": end_row,
                            "startColumnIndex": start_col,
                            "endColumnIndex": end_col,
                        },
                        "sortSpecs": [
                            {
                                "dimensionIndex": sort_column,
                                "sortOrder": "ASCENDING" if ascending else "DESCENDING",
                            }
                        ],
                    }
                }
            ]
        }

        self.spreadsheets.batchUpdate(
            spreadsheetId=spreadsheet_id, body=request
        ).execute()
        return True

    def find_replace(
        self,
        spreadsheet_id: str,
        find: str,
        replace: str,
        sheet_id: int | None = None,
        match_case: bool = False,
        match_entire_cell: bool = False,
    ) -> dict:
        """Find and replace values.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            find: Text to find.
            replace: Text to replace with.
            sheet_id: Optional sheet ID to limit search to.
            match_case: Case-sensitive matching.
            match_entire_cell: Match entire cell contents only.

        Returns:
            Result with number of replacements.
        """
        find_replace_request: dict[str, Any] = {
            "find": find,
            "replacement": replace,
            "matchCase": match_case,
            "matchEntireCell": match_entire_cell,
        }

        # allSheets and sheetId are mutually exclusive
        if sheet_id is not None:
            find_replace_request["sheetId"] = sheet_id
        else:
            find_replace_request["allSheets"] = True

        request = {"requests": [{"findReplace": find_replace_request}]}

        result = self.spreadsheets.batchUpdate(
            spreadsheetId=spreadsheet_id, body=request
        ).execute()

        fr_result = result["replies"][0]["findReplace"]
        return {
            "occurrences_changed": fr_result.get("occurrencesChanged", 0),
            "rows_changed": fr_result.get("rowsChanged", 0),
            "sheets_changed": fr_result.get("sheetsChanged", 0),
            "values_changed": fr_result.get("valuesChanged", 0),
        }

    def get_last_row(
        self, spreadsheet_id: str, sheet_name: str, column: str = "A"
    ) -> int:
        """Find the last row with data in a column.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            sheet_name: Name of the sheet.
            column: Column letter to check (default "A").

        Returns:
            Last row number with data (1-based), or 0 if empty.
        """
        range_notation = f"{sheet_name}!{column}:{column}"
        values = self.read_cells(spreadsheet_id, range_notation)
        return len(values)

    # =========================================================================
    # Sharing
    # =========================================================================

    def share_spreadsheet(
        self,
        spreadsheet_id: str,
        email: str | None = None,
        role: str = "reader",
        make_public: bool = False,
    ) -> dict:
        """Share a spreadsheet with someone or make it public.

        Args:
            spreadsheet_id: The ID of the spreadsheet.
            email: Email address to share with (optional if make_public=True).
            role: Permission role - "reader", "writer", or "commenter".
            make_public: If True, makes the spreadsheet accessible to anyone with link.

        Returns:
            Result with permission details.
        """
        drive = get_drive_service()

        if make_public:
            permission = {"type": "anyone", "role": role}
        elif email:
            permission = {"type": "user", "role": role, "emailAddress": email}
        else:
            raise ValueError("Either email or make_public=True must be provided")

        result = (
            drive.permissions()
            .create(
                fileId=spreadsheet_id,
                body=permission,
                sendNotificationEmail=bool(email),
            )
            .execute()
        )

        return {
            "permission_id": result.get("id", ""),
            "role": role,
            "type": "public" if make_public else "user",
            "email": email or "anyone",
        }


# Global client instance
_client: SheetsClient | None = None


def get_client() -> SheetsClient:
    """Get the global SheetsClient instance."""
    global _client
    if _client is None:
        _client = SheetsClient()
    return _client
