"""Unit tests for table and pivot modules — no Excel required.

Tests cover:
- Argument validation (action dispatch, _require guards)
- COM constant correctness (aggregation functions, areas, totals functions)
- Pure Python helpers (_func_name, _resolve_area, _TOTALS_FUNC)
"""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


# ── Tables: argument validation ────────────────────────────────────────────────

class TestTableActionDispatch:
    """Verify _require guards fire before any COM call.

    run_com is a transparent passthrough so _require guards execute synchronously
    in the test process (no STA worker, no Excel needed).
    """

    def _call(self, **kwargs):
        # Import here so that COM imports at module level don't blow up test
        # collection on non-Windows. We mock _session so no live Excel needed.
        from unittest.mock import patch
        mock_session = make_mock_session()
        mock_session.get_workbook.side_effect = ToolError("no excel")
        with patch("thepexcel_mcp.domains.tables._session", mock_session):
            from thepexcel_mcp.domains.tables import table_action
            return table_action(**kwargs)

    def test_unknown_action(self):
        with pytest.raises(ToolError, match="Unknown action"):
            self._call(action="frobnicate")

    def test_create_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="create", range="A1:D10")

    def test_create_missing_range(self):
        with pytest.raises(ToolError, match="requires 'range'"):
            self._call(action="create", name="T1")

    def test_read_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="read")

    def test_append_rows_missing_values(self):
        # name provided → gets past _require(name) → hits values check
        # (ToolError is raised before run_com, so no passthrough needed)
        from unittest.mock import patch
        mock_session = make_mock_session()
        with patch("thepexcel_mcp.domains.tables._session", mock_session):
            from thepexcel_mcp.domains.tables import table_action
            with pytest.raises(ToolError, match="requires 'values'"):
                table_action(action="append_rows", name="T1")

    def test_append_rows_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="append_rows", values=[[1, 2]])

    def test_add_column_missing_column_name(self):
        with pytest.raises(ToolError, match="requires 'column_name'"):
            self._call(action="add_column", name="T1")

    def test_sort_missing_sort_column(self):
        with pytest.raises(ToolError, match="requires 'sort_column'"):
            self._call(action="sort", name="T1")

    def test_filter_missing_filter_column(self):
        with pytest.raises(ToolError, match="requires 'filter_column'"):
            self._call(action="filter", name="T1")

    def test_toggle_totals_missing_show_totals(self):
        with pytest.raises(ToolError, match="requires 'show_totals'"):
            self._call(action="toggle_totals", name="T1")

    def test_rename_missing_new_name(self):
        with pytest.raises(ToolError, match="requires 'new_name'"):
            self._call(action="rename", name="T1")

    def test_set_style_missing_style(self):
        with pytest.raises(ToolError, match="requires 'style'"):
            self._call(action="set_style", name="T1")


class TestTotalsConstantMapping:
    """Verify _TOTALS_FUNC constant values match Excel VBA spec."""

    def setup_method(self):
        from thepexcel_mcp.domains.tables import _TOTALS_FUNC
        self.funcs = _TOTALS_FUNC

    def test_sum_is_1(self):
        assert self.funcs["sum"] == 1

    def test_average_is_2(self):
        assert self.funcs["average"] == 2
        assert self.funcs["avg"] == 2

    def test_count_is_3(self):
        assert self.funcs["count"] == 3

    def test_countnums_is_4(self):
        assert self.funcs["countnums"] == 4

    def test_max_is_5(self):
        assert self.funcs["max"] == 5

    def test_min_is_6(self):
        assert self.funcs["min"] == 6

    def test_stddev_is_7(self):
        assert self.funcs["stddev"] == 7

    def test_var_is_9(self):
        assert self.funcs["var"] == 9

    def test_none_is_0(self):
        assert self.funcs["none"] == 0


class TestTableFilterOp:
    """Verify filter op validation."""

    def test_invalid_filter_op_raises(self):
        from unittest.mock import MagicMock, patch
        mock_wb = MagicMock()
        mock_session = make_mock_session()
        mock_session.get_workbook.return_value = mock_wb
        # Make the fast-path ActiveSheet lookup miss too, so _find_table falls
        # through to full enumeration and raises "not found" as before.
        mock_wb.ActiveSheet.ListObjects.side_effect = Exception("not found")
        # Return empty table list so _find_table raises ToolError
        mock_wb.Sheets.Count = 0
        with patch("thepexcel_mcp.domains.tables._session", mock_session):
            from thepexcel_mcp.domains.tables import table_action
            with pytest.raises(ToolError):  # either "not found" or "invalid op"
                table_action(
                    action="filter",
                    name="T1",
                    filter_column="Status",
                    filter_op="startswith",
                    filter_value="A",
                )


# ── Pivots: argument validation ────────────────────────────────────────────────

class TestPivotActionDispatch:
    """Verify _require guards fire before any COM call."""

    def _call(self, **kwargs):
        from unittest.mock import patch
        mock_session = make_mock_session()
        mock_session.get_workbook.side_effect = ToolError("no excel")
        with patch("thepexcel_mcp.domains.pivots._session", mock_session):
            from thepexcel_mcp.domains.pivots import pivot_action
            return pivot_action(**kwargs)

    def test_unknown_action(self):
        with pytest.raises(ToolError, match="Unknown action"):
            self._call(action="frobnicate")

    def test_create_missing_source(self):
        with pytest.raises(ToolError, match="requires 'source'"):
            self._call(action="create", name="P1")

    def test_create_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="create", source="SalesTable")

    def test_add_field_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="add_field", field="Region", area="rows")

    def test_add_field_missing_field(self):
        with pytest.raises(ToolError, match="requires 'field'"):
            self._call(action="add_field", name="P1", area="rows")

    def test_add_field_missing_area(self):
        with pytest.raises(ToolError, match="requires 'area'"):
            self._call(action="add_field", name="P1", field="Region")

    def test_remove_field_missing_field(self):
        with pytest.raises(ToolError, match="requires 'field'"):
            self._call(action="remove_field", name="P1")

    def test_refresh_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="refresh")

    def test_delete_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="delete")

    def test_read_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="read")


class TestResolveDestSheet:
    """_resolve_dest_sheet: omitted → Pivot_<name>; named-existing → reuse;
    named-but-absent → CREATE (not a 'sheet not found' error)."""

    def _fn(self):
        from thepexcel_mcp.domains.pivots import _resolve_dest_sheet
        return _resolve_dest_sheet

    def test_omitted_creates_pivot_named_sheet(self):
        from unittest.mock import MagicMock, patch
        wb, app = MagicMock(), MagicMock()
        new_ws = MagicMock()
        wb.Sheets.Add.return_value = new_ws
        with patch("thepexcel_mcp.domains.pivots._session") as sess:
            ws = self._fn()(wb, app, None, "Sales", None)
        wb.Sheets.Add.assert_called_once()
        assert new_ws.Name == "Pivot_Sales"
        assert ws is new_ws
        sess.get_sheet.assert_not_called()

    def test_named_existing_sheet_is_reused(self):
        from unittest.mock import MagicMock, patch
        wb, app = MagicMock(), MagicMock()
        existing = MagicMock()
        with patch("thepexcel_mcp.domains.pivots._session") as sess:
            sess.get_sheet.return_value = existing
            ws = self._fn()(wb, app, "Report", "P1", None)
        assert ws is existing
        wb.Sheets.Add.assert_not_called()

    def test_named_absent_sheet_is_created(self):
        """REGRESSION: a named dest_sheet that doesn't exist must be CREATED,
        not raise 'sheet not found' (the awkward behavior fixed for v0.2)."""
        from unittest.mock import MagicMock, patch
        wb, app = MagicMock(), MagicMock()
        new_ws = MagicMock()
        wb.Sheets.Add.return_value = new_ws
        with patch("thepexcel_mcp.domains.pivots._session") as sess:
            sess.get_sheet.side_effect = Exception("Sheet 'Report' not found")
            ws = self._fn()(wb, app, "Report", "P1", None)
        wb.Sheets.Add.assert_called_once()
        assert new_ws.Name == "Report"
        assert ws is new_ws

    def test_named_absent_truncated_to_31_chars(self):
        from unittest.mock import MagicMock, patch
        wb, app = MagicMock(), MagicMock()
        new_ws = MagicMock()
        wb.Sheets.Add.return_value = new_ws
        with patch("thepexcel_mcp.domains.pivots._session") as sess:
            sess.get_sheet.side_effect = Exception("not found")
            self._fn()(wb, app, "S" * 40, "P1", None)
        assert new_ws.Name == "S" * 31


class TestPivotConstants:
    """Verify COM constant values match VBA spec (from sbroenne PivotTableTypes.cs)."""

    def setup_method(self):
        import thepexcel_mcp.domains.pivots as p
        self.mod = p

    def test_orientation_constants(self):
        assert self.mod._XL_HIDDEN == 0
        assert self.mod._XL_ROW_FIELD == 1
        assert self.mod._XL_COLUMN_FIELD == 2
        assert self.mod._XL_PAGE_FIELD == 3
        assert self.mod._XL_DATA_FIELD == 4

    def test_aggregation_sum(self):
        assert self.mod._AGG_FUNC["sum"] == -4157

    def test_aggregation_count(self):
        assert self.mod._AGG_FUNC["count"] == -4112

    def test_aggregation_average(self):
        assert self.mod._AGG_FUNC["average"] == -4106
        assert self.mod._AGG_FUNC["avg"] == -4106

    def test_aggregation_max(self):
        assert self.mod._AGG_FUNC["max"] == -4136

    def test_aggregation_min(self):
        assert self.mod._AGG_FUNC["min"] == -4139

    def test_aggregation_product(self):
        assert self.mod._AGG_FUNC["product"] == -4149

    def test_aggregation_stddev(self):
        assert self.mod._AGG_FUNC["stddev"] == -4155

    def test_aggregation_var(self):
        assert self.mod._AGG_FUNC["var"] == -4164

    def test_source_type_database(self):
        assert self.mod._XL_DATABASE == 1

    def test_source_type_external(self):
        assert self.mod._XL_EXTERNAL == 2

    def test_layout_constants(self):
        assert self.mod._LAYOUT["compact"] == 0
        assert self.mod._LAYOUT["tabular"] == 1
        assert self.mod._LAYOUT["outline"] == 2


class TestResolveArea:
    """Verify _resolve_area maps strings to orientation constants correctly."""

    def setup_method(self):
        from thepexcel_mcp.domains.pivots import _resolve_area, _XL_ROW_FIELD, \
            _XL_COLUMN_FIELD, _XL_PAGE_FIELD, _XL_DATA_FIELD
        self._resolve = _resolve_area
        self.ROW = _XL_ROW_FIELD
        self.COL = _XL_COLUMN_FIELD
        self.FILTER = _XL_PAGE_FIELD
        self.DATA = _XL_DATA_FIELD

    def test_rows(self):
        assert self._resolve("rows") == self.ROW
        assert self._resolve("row") == self.ROW

    def test_columns(self):
        assert self._resolve("columns") == self.COL
        assert self._resolve("column") == self.COL
        assert self._resolve("col") == self.COL

    def test_filters(self):
        assert self._resolve("filters") == self.FILTER
        assert self._resolve("filter") == self.FILTER

    def test_values(self):
        assert self._resolve("values") == self.DATA
        assert self._resolve("value") == self.DATA
        assert self._resolve("data") == self.DATA

    def test_invalid_area(self):
        with pytest.raises(ToolError, match="area="):
            self._resolve("pages")


class TestFuncName:
    """Verify _func_name round-trips COM constants to readable names."""

    def setup_method(self):
        from thepexcel_mcp.domains.pivots import _func_name
        self._fn = _func_name

    def test_sum(self):
        assert self._fn(-4157) == "sum"

    def test_count(self):
        assert self._fn(-4112) == "count"

    def test_average(self):
        assert self._fn(-4106) == "average"

    def test_unknown_returns_string(self):
        result = self._fn(999)
        assert result == "999"


# ── tables._create: single workbook resolution (no double get_sheet) ──────────

class TestCreateSingleWorkbookResolution:
    def _make_wb(self, sheet_count=0):
        from unittest.mock import MagicMock
        wb = MagicMock()
        wb.Sheets.Count = sheet_count
        return wb

    def test_create_never_calls_get_sheet(self):
        """_create must resolve the sheet from the already-fetched wb, not via
        _session.get_sheet() (which would re-resolve the workbook)."""
        from unittest.mock import MagicMock
        from thepexcel_mcp.domains.tables import _create

        wb = self._make_wb(sheet_count=0)
        named_ws = MagicMock()
        wb.Sheets.return_value = named_ws
        lo = MagicMock()
        named_ws.ListObjects.Add.return_value = lo

        mock_session = make_mock_session()
        mock_session.get_workbook.return_value = wb

        from unittest.mock import patch
        with patch("thepexcel_mcp.domains.tables._session", mock_session):
            with patch("thepexcel_mcp.domains.tables._table_info", return_value={"name": "T1"}):
                _create("T1", "A1:B2", "Data", None, "TableStyleMedium2", True)

        mock_session.get_sheet.assert_not_called()
        wb.Sheets.assert_any_call("Data")

    def test_create_defaults_to_active_sheet_when_sheet_omitted(self):
        from unittest.mock import MagicMock
        from thepexcel_mcp.domains.tables import _create

        wb = self._make_wb(sheet_count=0)
        active_ws = MagicMock()
        wb.ActiveSheet = active_ws
        lo = MagicMock()
        active_ws.ListObjects.Add.return_value = lo

        mock_session = make_mock_session()
        mock_session.get_workbook.return_value = wb

        from unittest.mock import patch
        with patch("thepexcel_mcp.domains.tables._session", mock_session):
            with patch("thepexcel_mcp.domains.tables._table_info", return_value={"name": "T1"}):
                _create("T1", "A1:B2", None, None, "TableStyleMedium2", True)

        mock_session.get_sheet.assert_not_called()
        active_ws.Range.assert_called_once_with("A1:B2")

    def test_create_sheet_not_found_preserves_error_format(self):
        from unittest.mock import MagicMock

        wb = self._make_wb(sheet_count=2)
        s1, s2 = MagicMock(), MagicMock()
        s1.Name, s2.Name = "Sheet1", "Sheet2"

        def sheets_lookup(key):
            if key == "Ghost":
                raise Exception("subscript out of range")
            return {1: s1, 2: s2}[key]

        wb.Sheets.side_effect = sheets_lookup

        mock_session = make_mock_session()
        mock_session.get_workbook.return_value = wb

        from unittest.mock import patch
        with patch("thepexcel_mcp.domains.tables._session", mock_session):
            from thepexcel_mcp.domains.tables import _create
            with pytest.raises(ToolError, match=r"Sheet 'Ghost' not found. Available: \['Sheet1', 'Sheet2'\]"):
                _create("T1", "A1:B2", "Ghost", None, "TableStyleMedium2", True)


# ── tables._read: page-scoped reads ─────────────────────────────────────────────

class TestTableReadPageScoped:
    def _make_lo(self, col_names, total_rows, row=2, col=1):
        """lo.DataBodyRange geometry: total_rows x len(col_names), header at row 1."""
        from unittest.mock import MagicMock

        lo = MagicMock()
        lo.ListColumns.Count = len(col_names)
        lo.ListColumns.side_effect = lambda i: MagicMock(Name=col_names[i - 1])

        dbr = MagicMock()
        dbr.Rows.Count = total_rows
        dbr.Columns.Count = len(col_names)
        dbr.Row = row
        dbr.Column = col
        ws = MagicMock()
        dbr.Parent = ws
        lo.DataBodyRange = dbr
        return lo, dbr, ws

    def test_page_scoped_block_built_from_databodyrange_geometry(self):
        from thepexcel_mcp.domains.tables import _read

        lo, dbr, ws = self._make_lo(["A", "B"], total_rows=150, row=2, col=1)
        page_rng = MagicMock()
        page_rng.Value = tuple((i, i * 2) for i in range(100))
        ws.Range.return_value = page_rng

        mock_session = make_mock_session()
        mock_session.get_workbook.return_value = MagicMock()

        from unittest.mock import patch
        with patch("thepexcel_mcp.domains.tables._session", mock_session):
            with patch("thepexcel_mcp.domains.tables._find_table", return_value=lo):
                result = _read("T1", None, None, offset=0, limit=100)

        ws.Cells.assert_any_call(2, 1)     # first data row
        ws.Cells.assert_any_call(101, 2)   # last row of this page, last col
        assert result["total_rows"] == 150
        assert result["has_more"] is True
        assert result["next_offset"] == 100
        assert len(result["values"]) == 100

    def test_offset_beyond_total_rows_skips_com_read(self):
        from thepexcel_mcp.domains.tables import _read

        lo, dbr, ws = self._make_lo(["A"], total_rows=5)

        mock_session = make_mock_session()
        from unittest.mock import patch
        with patch("thepexcel_mcp.domains.tables._session", mock_session):
            with patch("thepexcel_mcp.domains.tables._find_table", return_value=lo):
                result = _read("T1", None, None, offset=50, limit=100)

        assert result["values"] == []
        assert result["has_more"] is False
        ws.Range.assert_not_called()

    def test_column_filter_applied_after_paging(self):
        from thepexcel_mcp.domains.tables import _read

        lo, dbr, ws = self._make_lo(["A", "B", "C"], total_rows=2)
        page_rng = MagicMock()
        page_rng.Value = ((1, 2, 3), (4, 5, 6))
        ws.Range.return_value = page_rng

        mock_session = make_mock_session()
        from unittest.mock import patch
        with patch("thepexcel_mcp.domains.tables._session", mock_session):
            with patch("thepexcel_mcp.domains.tables._find_table", return_value=lo):
                result = _read("T1", None, columns=["C", "A"], offset=0, limit=100)

        assert result["columns"] == ["C", "A"]
        assert result["values"] == [[3, 1], [6, 4]]

    def test_no_data_body_range_returns_empty(self):
        from thepexcel_mcp.domains.tables import _read

        lo, dbr, ws = self._make_lo(["A"], total_rows=0)
        lo.DataBodyRange = None

        mock_session = make_mock_session()
        from unittest.mock import patch
        with patch("thepexcel_mcp.domains.tables._session", mock_session):
            with patch("thepexcel_mcp.domains.tables._find_table", return_value=lo):
                result = _read("T1", None, None, offset=0, limit=100)

        assert result == {
            "columns": ["A"], "values": [], "total_rows": 0,
            "has_more": False, "next_offset": None,
        }


# ── tables._append_rows: single block write ─────────────────────────────────────

class TestAppendRowsBlockWrite:
    def _make_lo(self, n_cols=3, existing_rows=2):
        """lo.DataBodyRange geometry BEFORE the append (existing_rows rows)."""
        lo = MagicMock()
        lo.ListColumns.Count = n_cols
        wb_app = MagicMock()

        dbr_before = MagicMock()
        dbr_before.Row = 2  # header at row 1, data starts row 2
        dbr_before.Column = 1
        dbr_before.Rows.Count = existing_rows

        ws = MagicMock()
        dbr_before.Parent = ws

        # DataBodyRange is read TWICE: once inside the guard (post-Add, for the
        # block geometry) and once after (for the final row-count verify).
        # Simulate ListRows.Add() growing the row count between reads.
        state = {"count": existing_rows}

        def add_row():
            state["count"] += 1
            return MagicMock()

        lo.ListRows.Add.side_effect = add_row

        def dbr_prop():
            d = MagicMock()
            d.Row = 2
            d.Column = 1
            d.Rows.Count = state["count"]
            d.Parent = ws
            return d

        type(lo).DataBodyRange = property(lambda self: dbr_prop())
        lo.Parent.Application = wb_app
        return lo, ws, wb_app

    def _call(self, lo, values, wb_app):
        wb = MagicMock()
        wb.Application = wb_app
        mock_session = make_mock_session()
        mock_session.get_workbook.return_value = wb
        from unittest.mock import patch
        with patch("thepexcel_mcp.domains.tables._session", mock_session):
            with patch("thepexcel_mcp.domains.tables._find_table", return_value=lo):
                from thepexcel_mcp.domains.tables import _append_rows
                return _append_rows("T1", None, values)

    def test_adds_a_listrow_per_value_row(self):
        lo, ws, wb_app = self._make_lo(n_cols=2, existing_rows=0)
        result = self._call(lo, [[1, 2], [3, 4], [5, 6]], wb_app)

        assert lo.ListRows.Add.call_count == 3
        assert result["appended_rows"] == 3
        assert result["total_rows"] == 3

    def test_single_block_value_write_not_per_cell(self):
        """The new rows' data must be written via ONE block .Value assignment,
        not R×C per-cell writes."""
        lo, ws, wb_app = self._make_lo(n_cols=2, existing_rows=0)
        block = MagicMock()
        ws.Range.return_value = block

        self._call(lo, [[1, 2], [3, 4]], wb_app)

        # Exactly one Range(...) call built the write target, one Value assignment.
        ws.Range.assert_called_once()
        assert isinstance(block.Value, tuple)
        assert block.Value == ((1, 2), (3, 4))

    def test_ragged_rows_padded_with_none_up_to_widest_row(self):
        lo, ws, wb_app = self._make_lo(n_cols=5, existing_rows=0)
        block = MagicMock()
        ws.Range.return_value = block

        self._call(lo, [[1, 2, 3], [4]], wb_app)

        # n_cols capped at widest supplied row (3), not the full table width (5)
        assert block.Value == ((1, 2, 3), (4, None, None))

    def test_column_count_never_exceeds_table_width(self):
        lo, ws, wb_app = self._make_lo(n_cols=2, existing_rows=0)
        block = MagicMock()
        ws.Range.return_value = block

        # Rows supply MORE values than the table has columns — extras dropped.
        self._call(lo, [[1, 2, 3, 4]], wb_app)

        assert block.Value == ((1, 2),)

    def test_bulk_guard_wraps_the_write(self):
        lo, ws, wb_app = self._make_lo(n_cols=2, existing_rows=0)

        from unittest.mock import patch
        with patch("thepexcel_mcp.domains.tables.bulk_guard") as bg:
            bg.return_value.__enter__ = MagicMock(return_value=None)
            bg.return_value.__exit__ = MagicMock(return_value=False)
            self._call(lo, [[1, 2]], wb_app)

        bg.assert_called_once_with(wb_app)


# ── _find_table fast path ────────────────────────────────────────────────────────

class TestFindTableFastPath:
    def test_active_sheet_hit_skips_full_enumeration(self):
        from unittest.mock import MagicMock
        from thepexcel_mcp.domains.tables import _find_table

        wb = MagicMock()
        target_lo = MagicMock()
        wb.ActiveSheet.ListObjects.return_value = target_lo
        # Sheets(i) enumeration must NEVER be reached — fast path hits first
        wb.Sheets.side_effect = AssertionError("full enumeration must not run")

        result = _find_table(wb, "SalesTable")

        assert result is target_lo

    def test_given_sheet_hit_tried_before_active_sheet(self):
        from unittest.mock import MagicMock
        from thepexcel_mcp.domains.tables import _find_table

        wb = MagicMock()
        target_lo = MagicMock()
        named_ws = MagicMock()
        named_ws.ListObjects.return_value = target_lo
        wb.Sheets.return_value = named_ws
        wb.ActiveSheet.ListObjects.side_effect = AssertionError("should not reach ActiveSheet")

        result = _find_table(wb, "SalesTable", sheet="Data")

        assert result is target_lo
        wb.Sheets.assert_called_once_with("Data")

    def test_fast_path_miss_falls_back_to_full_enumeration(self):
        from unittest.mock import MagicMock
        from fastmcp.exceptions import ToolError
        from thepexcel_mcp.domains.tables import _find_table

        wb = MagicMock()
        wb.ActiveSheet.ListObjects.side_effect = Exception("not on active sheet")
        wb.Sheets.Count = 0  # empty enumeration → not found

        with pytest.raises(ToolError, match="not found"):
            _find_table(wb, "Ghost")

    def test_fast_path_miss_then_full_enumeration_finds_it(self):
        from unittest.mock import MagicMock
        from thepexcel_mcp.domains.tables import _find_table

        wb = MagicMock()
        wb.ActiveSheet.ListObjects.side_effect = Exception("not on active sheet")

        target_lo = MagicMock()
        target_lo.Name = "SalesTable"
        ws1 = MagicMock()
        ws1.ListObjects.Count = 1
        ws1.ListObjects.side_effect = lambda j: target_lo
        wb.Sheets.Count = 1
        wb.Sheets.side_effect = lambda i: ws1

        result = _find_table(wb, "SalesTable")

        assert result is target_lo


# ── pivots._read: page-scoped reads ─────────────────────────────────────────────

class TestPivotReadPageScoped:
    def _make_pt(self, total_rows, total_cols, row=1, col=1):
        from unittest.mock import MagicMock

        pt = MagicMock()
        rng = MagicMock()
        rng.Rows.Count = total_rows
        rng.Columns.Count = total_cols
        rng.Row = row
        rng.Column = col
        ws = MagicMock()
        rng.Parent = ws
        pt.TableRange1 = rng
        return pt, rng, ws

    def test_page_scoped_block_built_from_tablerange_geometry(self):
        from thepexcel_mcp.domains.pivots import _read

        pt, rng, ws = self._make_pt(total_rows=120, total_cols=3, row=3, col=1)
        page_rng = MagicMock()
        page_rng.Value = tuple((i, i, i) for i in range(100))
        ws.Range.return_value = page_rng

        mock_session = make_mock_session()
        from unittest.mock import patch
        with patch("thepexcel_mcp.domains.pivots._session", mock_session):
            with patch("thepexcel_mcp.domains.pivots._find_pivot", return_value=pt):
                result = _read("P1", None, offset=0, limit=100)

        ws.Cells.assert_any_call(3, 1)
        ws.Cells.assert_any_call(102, 3)
        assert result["total_rows"] == 120
        assert result["has_more"] is True
        assert result["next_offset"] == 100

    def test_offset_beyond_total_rows_skips_com_read(self):
        from thepexcel_mcp.domains.pivots import _read

        pt, rng, ws = self._make_pt(total_rows=5, total_cols=2)

        mock_session = make_mock_session()
        from unittest.mock import patch
        with patch("thepexcel_mcp.domains.pivots._session", mock_session):
            with patch("thepexcel_mcp.domains.pivots._find_pivot", return_value=pt):
                result = _read("P1", None, offset=50, limit=100)

        assert result["values"] == []
        assert result["has_more"] is False
        ws.Range.assert_not_called()

    def test_single_empty_cell_reports_zero_rows(self):
        """Legacy quirk preserved: whole-value None on a 1x1 range → total_rows=0."""
        from thepexcel_mcp.domains.pivots import _read

        pt, rng, ws = self._make_pt(total_rows=1, total_cols=1)
        rng.Value = None

        mock_session = make_mock_session()
        from unittest.mock import patch
        with patch("thepexcel_mcp.domains.pivots._session", mock_session):
            with patch("thepexcel_mcp.domains.pivots._find_pivot", return_value=pt):
                result = _read("P1", None, offset=0, limit=100)

        assert result == {"values": [], "total_rows": 0, "has_more": False, "next_offset": None}
        ws.Range.assert_not_called()


# ── _find_pivot fast path ─────────────────────────────────────────────────────

class TestFindPivotFastPath:
    def test_active_sheet_hit_skips_full_enumeration(self):
        from unittest.mock import MagicMock
        from thepexcel_mcp.domains.pivots import _find_pivot

        wb = MagicMock()
        target_pt = MagicMock()
        wb.ActiveSheet.PivotTables.return_value = target_pt
        wb.Sheets.side_effect = AssertionError("full enumeration must not run")

        result = _find_pivot(wb, "SalesPivot")

        assert result is target_pt

    def test_fast_path_miss_falls_back_to_full_enumeration(self):
        from unittest.mock import MagicMock
        from fastmcp.exceptions import ToolError
        from thepexcel_mcp.domains.pivots import _find_pivot

        wb = MagicMock()
        wb.ActiveSheet.PivotTables.side_effect = Exception("not on active sheet")
        wb.Sheets.Count = 0

        with pytest.raises(ToolError, match="not found"):
            _find_pivot(wb, "Ghost")


# ── PQ warnings integration ────────────────────────────────────────────────────

class TestPQWarnings:
    """Verify _get_warnings returns warning strings and never raises."""

    def test_warnings_for_unnecessary_buffer(self):
        # unnecessary-buffer fires at severity=warning (no join present)
        from unittest.mock import MagicMock
        from thepexcel_mcp.domains.powerquery import _get_warnings
        mock_q = MagicMock()
        mock_q.Name = "TestQ"
        mock_q.Formula = (
            'let Source = Excel.CurrentWorkbook(){[Name="T"]}[Content],\n'
            '    Buffered = Table.Buffer(Source)\n'
            'in Buffered'
        )
        warnings = _get_warnings(mock_q)
        assert isinstance(warnings, list)
        # unnecessary-buffer rule should fire at warning level
        assert any("unnecessary-buffer" in w for w in warnings)

    def test_clean_query_no_warnings(self):
        from unittest.mock import MagicMock
        from thepexcel_mcp.domains.powerquery import _get_warnings
        mock_q = MagicMock()
        mock_q.Name = "CleanQ"
        mock_q.Formula = (
            'let Source = Excel.CurrentWorkbook(){[Name="T1"]}[Content] in Source'
        )
        warnings = _get_warnings(mock_q)
        assert warnings == []

    def test_get_warnings_never_raises(self):
        """Even if analyzer throws, _get_warnings returns [] not exception."""
        from unittest.mock import MagicMock, patch
        from thepexcel_mcp.domains.powerquery import _get_warnings
        mock_q = MagicMock()
        mock_q.Name = "Q"
        mock_q.Formula = "not valid M code *** !!!"
        # Should not raise
        result = _get_warnings(mock_q)
        assert isinstance(result, list)
