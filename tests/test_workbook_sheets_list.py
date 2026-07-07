"""Unit tests for workbook._list/_info and sheets._list — no Excel required.

Covers the Phase 1 "bind each collection item once per iteration" optimization:
these each used to call Workbooks(i+1)/Sheets(i+1) TWICE per loop iteration
(once for .Name, once for the equality check) — verify it's now once, and
that _info's two separate sheet loops are merged into a single pass.
"""

from __future__ import annotations

from unittest.mock import MagicMock, patch

from conftest import make_mock_session


# ── workbook._list ───────────────────────────────────────────────────────────────

class TestWorkbookList:
    def _make_app(self, names, active_name):
        app = MagicMock()
        wbs = []
        for n in names:
            wb = MagicMock()
            wb.Name = n
            wbs.append(wb)
        app.Workbooks.Count = len(wbs)
        app.Workbooks.side_effect = lambda i, _wbs=wbs: _wbs[i - 1]
        active = MagicMock()
        active.Name = active_name
        app.ActiveWorkbook = active
        return app, wbs

    def test_names_and_active_flag(self):
        from thepexcel_mcp.domains.workbook import _list

        app, wbs = self._make_app(["Book1.xlsx", "Book2.xlsx"], "Book2.xlsx")
        with patch("thepexcel_mcp.domains.workbook._session") as sess:
            sess.get_app.return_value = app
            result = _list()

        assert result["count"] == 2
        assert result["workbooks"] == [
            {"name": "Book1.xlsx", "active": False},
            {"name": "Book2.xlsx", "active": True},
        ]

    def test_workbooks_indexed_once_per_item(self):
        """Workbooks(i+1) must be called exactly once per item, not twice
        (bind-once optimization) — total calls == item count."""
        from thepexcel_mcp.domains.workbook import _list

        app, wbs = self._make_app(["A.xlsx", "B.xlsx", "C.xlsx"], "A.xlsx")
        with patch("thepexcel_mcp.domains.workbook._session") as sess:
            sess.get_app.return_value = app
            _list()

        assert app.Workbooks.call_count == 3

    def test_no_active_workbook(self):
        from thepexcel_mcp.domains.workbook import _list

        app, wbs = self._make_app(["A.xlsx"], "A.xlsx")
        app.ActiveWorkbook = None
        with patch("thepexcel_mcp.domains.workbook._session") as sess:
            sess.get_app.return_value = app
            result = _list()

        assert result["workbooks"] == [{"name": "A.xlsx", "active": False}]


# ── workbook._info ───────────────────────────────────────────────────────────────

class TestWorkbookInfo:
    def _make_wb(self, sheet_specs):
        """sheet_specs: list of (name, list_object_count)."""
        wb = MagicMock()
        sheets = []
        for name, lo_count in sheet_specs:
            ws = MagicMock()
            ws.Name = name
            ws.ListObjects.Count = lo_count
            sheets.append(ws)
        wb.Sheets.Count = len(sheets)
        wb.Sheets.side_effect = lambda i, _s=sheets: _s[i - 1]
        wb.Queries.Count = 2
        wb.Names.Count = 5
        wb.Name = "Test.xlsx"
        wb.FullName = "C:/Test.xlsx"
        wb.Saved = True
        return wb, sheets

    def test_merged_single_pass_result(self):
        from thepexcel_mcp.domains.workbook import _info

        wb, sheets = self._make_wb([("Sheet1", 1), ("Sheet2", 2), ("Sheet3", 0)])
        ms = make_mock_session()
        ms.get_workbook.return_value = wb
        with patch("thepexcel_mcp.domains.workbook._session", ms):
            result = _info(None)

        assert result["sheets"] == ["Sheet1", "Sheet2", "Sheet3"]
        assert result["sheet_count"] == 3
        assert result["table_count"] == 3  # 1 + 2 + 0
        assert result["query_count"] == 2
        assert result["defined_name_count"] == 5
        assert result["saved"] is True

    def test_sheets_indexed_once_per_item_single_pass(self):
        """The old code looped wb.Sheets(i+1) twice (once for names, once for
        table counts) — must now be a single pass: exactly N Sheets() calls."""
        from thepexcel_mcp.domains.workbook import _info

        wb, sheets = self._make_wb([("A", 0), ("B", 1), ("C", 0), ("D", 3)])
        ms = make_mock_session()
        ms.get_workbook.return_value = wb
        with patch("thepexcel_mcp.domains.workbook._session", ms):
            _info(None)

        assert wb.Sheets.call_count == 4


# ── sheets._list ─────────────────────────────────────────────────────────────────

class TestSheetsList:
    def _make_wb(self, names, active_name):
        wb = MagicMock()
        sheets = []
        for n in names:
            ws = MagicMock()
            ws.Name = n
            sheets.append(ws)
        wb.Sheets.Count = len(sheets)
        wb.Sheets.side_effect = lambda i, _s=sheets: _s[i - 1]
        active = MagicMock()
        active.Name = active_name
        wb.ActiveSheet = active
        return wb, sheets

    def test_names_and_active_flag(self):
        from thepexcel_mcp.domains.sheets import _list

        wb, sheets = self._make_wb(["Sheet1", "Sheet2"], "Sheet1")
        ms = make_mock_session()
        ms.get_workbook.return_value = wb
        with patch("thepexcel_mcp.domains.sheets._session", ms):
            result = _list(None)

        assert result["count"] == 2
        assert result["sheets"] == [
            {"name": "Sheet1", "active": True},
            {"name": "Sheet2", "active": False},
        ]

    def test_sheets_indexed_once_per_item(self):
        from thepexcel_mcp.domains.sheets import _list

        wb, sheets = self._make_wb(["A", "B", "C"], "B")
        ms = make_mock_session()
        ms.get_workbook.return_value = wb
        with patch("thepexcel_mcp.domains.sheets._session", ms):
            _list(None)

        assert wb.Sheets.call_count == 3
