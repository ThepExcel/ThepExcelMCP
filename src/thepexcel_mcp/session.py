"""ExcelSession — COM attach/launch, workbook routing, shared error helpers.

One session object is created at server startup and reused across all tool
calls. Each public method calls pythoncom.CoInitialize() because FastMCP may
dispatch tool calls from different threads.
"""

from __future__ import annotations

import os

import pythoncom
import win32com.client
from fastmcp.exceptions import ToolError


class ExcelSession:
    """Thin wrapper around Excel.Application COM object.

    Attach strategy:
    1. GetActiveObject("Excel.Application") — zero-cost, uses already-running Excel.
    2. If that fails AND THEPEXCEL_MCP_AUTOLAUNCH=1 → Dispatch a new visible instance.
    3. Otherwise raise ToolError with an actionable message.
    """

    def get_app(self) -> win32com.client.CDispatch:
        """Return the running Excel.Application COM object."""
        pythoncom.CoInitialize()
        try:
            return win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            if os.environ.get("THEPEXCEL_MCP_AUTOLAUNCH") == "1":
                app = win32com.client.Dispatch("Excel.Application")
                app.Visible = True
                return app
            raise ToolError(
                "Excel is not running — open Excel first, then retry. "
                "Or set THEPEXCEL_MCP_AUTOLAUNCH=1 to auto-launch."
            )

    def get_workbook(self, workbook: str | None = None) -> win32com.client.CDispatch:
        """Return a Workbook COM object.

        Args:
            workbook: Workbook name (e.g. "Sales.xlsx"). Pass None for active workbook.
        """
        app = self.get_app()
        if workbook:
            try:
                return app.Workbooks(workbook)
            except Exception:
                available = [app.Workbooks(i + 1).Name for i in range(app.Workbooks.Count)]
                raise ToolError(
                    f"Workbook '{workbook}' not found. "
                    f"Open workbooks: {available or ['(none)']}"
                )
        wb = app.ActiveWorkbook
        if wb is None:
            raise ToolError("No active workbook — open an Excel file first.")
        return wb

    def get_sheet(
        self, name: str | None, workbook: str | None = None
    ) -> win32com.client.CDispatch:
        """Return a Worksheet COM object.

        Args:
            name: Sheet name. Pass None for active sheet.
            workbook: Workbook name. Pass None for active workbook.
        """
        wb = self.get_workbook(workbook)
        if name is None:
            return wb.ActiveSheet
        try:
            return wb.Sheets(name)
        except Exception:
            available = [wb.Sheets(i + 1).Name for i in range(wb.Sheets.Count)]
            raise ToolError(
                f"Sheet '{name}' not found. Available: {available}"
            )

    @staticmethod
    def wrap(exc: Exception, context: str = "") -> ToolError:
        """Convert a raw COM exception to an actionable ToolError."""
        prefix = f"{context}: " if context else ""
        return ToolError(f"{prefix}{exc}")
