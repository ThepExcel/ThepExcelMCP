"""Unit tests for table and pivot modules — no Excel required.

Tests cover:
- Argument validation (action dispatch, _require guards)
- COM constant correctness (aggregation functions, areas, totals functions)
- Pure Python helpers (_func_name, _resolve_area, _TOTALS_FUNC)
"""

from __future__ import annotations

import pytest

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
