"""Unit tests for excel_outline (outline.py).

No Excel required — all COM calls are intercepted via make_mock_session.

VERIFY-EFFECT discipline: every test asserts the SPECIFIC COM method that
was called, and for group_* actions, asserts the OutlineLevel read-back
value is echoed in the result dict.
"""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock, patch, call

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


# ── Helpers ────────────────────────────────────────────────────────────────────

def _make_sheet_mock(name: str = "Sheet1") -> MagicMock:
    """Minimal Worksheet mock sufficient for outline.py tests."""
    ws = MagicMock()
    ws.Name = name

    # ws.Rows(<spec>) returns a range mock whose .OutlineLevel = 2 (grouped)
    rows_range = MagicMock()
    rows_range.OutlineLevel = 2
    ws.Rows.return_value = rows_range

    # ws.Columns(<spec>) returns a range mock whose .OutlineLevel = 2 (grouped)
    cols_range = MagicMock()
    cols_range.OutlineLevel = 2
    ws.Columns.return_value = cols_range

    # ws.Outline — has ShowLevels method
    ws.Outline = MagicMock()

    # ws.Cells — has ClearOutline method
    ws.Cells = MagicMock()

    return ws


def _call_outline(action, mock_session, mock_ws, **kwargs):
    """Patch _session + excel_guard, invoke outline_action."""
    mock_ws.Application = MagicMock()
    mock_session.get_sheet.return_value = mock_ws

    with patch("thepexcel_mcp.domains.outline._session", mock_session):
        with patch("thepexcel_mcp.domains.outline.excel_guard") as eg:
            eg.return_value.__enter__ = MagicMock(return_value=None)
            eg.return_value.__exit__ = MagicMock(return_value=False)
            from thepexcel_mcp.domains.outline import outline_action
            return outline_action(action, sheet=None, workbook=None, **kwargs)


# ── Action validation ──────────────────────────────────────────────────────────

class TestActionValidation:
    def test_unknown_action_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="Unknown action"):
            _call_outline("bogus", ms, ws)

    def test_group_rows_missing_rows_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="rows"):
            _call_outline("group_rows", ms, ws)  # rows=None (default)

    def test_group_columns_missing_columns_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="columns"):
            _call_outline("group_columns", ms, ws)  # columns=None

    def test_ungroup_rows_missing_rows_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="rows"):
            _call_outline("ungroup_rows", ms, ws)

    def test_ungroup_columns_missing_columns_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="columns"):
            _call_outline("ungroup_columns", ms, ws)

    def test_show_levels_no_args_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="row_levels"):
            _call_outline("show_levels", ms, ws)  # both None

    def test_show_levels_row_levels_out_of_range_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="row_levels"):
            _call_outline("show_levels", ms, ws, row_levels=9)

    def test_show_levels_column_levels_out_of_range_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="column_levels"):
            _call_outline("show_levels", ms, ws, column_levels=0)

    def test_show_levels_row_below_min_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="row_levels"):
            _call_outline("show_levels", ms, ws, row_levels=0)

    def test_unknown_action_lists_valid(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="Valid"):
            _call_outline("noop", ms, ws)


# ── group_rows ────────────────────────────────────────────────────────────────

class TestGroupRows:
    def test_group_calls_group_on_rows_range(self):
        """VERIFY-EFFECT: ws.Rows('2:5').Group() must be called.

        ws.Rows is called twice: first with '2:5' (Group), then with 2 (read-back).
        Use assert_any_call to check the Group call regardless of order.
        """
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rows_rng = MagicMock()
        rows_rng.OutlineLevel = 2
        ws.Rows.return_value = rows_rng

        _call_outline("group_rows", ms, ws, rows="2:5")

        ws.Rows.assert_any_call("2:5")
        rows_rng.Group.assert_called_once()

    def test_group_rows_result_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()

        result = _call_outline("group_rows", ms, ws, rows="2:5")

        assert result["outline"] == "group_rows"
        assert result["sheet"] == "Sheet1"
        assert "applied" in result
        assert result["applied"]["rows"] == "2:5"

    def test_group_rows_outline_level_echoed_in_result(self):
        """VERIFY-EFFECT: OutlineLevel read-back must appear in applied dict."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rows_rng = MagicMock()
        rows_rng.OutlineLevel = 3  # simulate nested group
        ws.Rows.return_value = rows_rng

        result = _call_outline("group_rows", ms, ws, rows="3:6")

        assert result["applied"]["outline_level"] == 3

    def test_group_rows_reads_first_row_for_level(self):
        """OutlineLevel read-back uses first row number parsed from spec."""
        ms = make_mock_session()
        ws = _make_sheet_mock()

        row_spec_calls = []
        first_rng = MagicMock()
        first_rng.OutlineLevel = 2

        def rows_side_effect(spec):
            row_spec_calls.append(spec)
            return first_rng

        ws.Rows.side_effect = rows_side_effect

        _call_outline("group_rows", ms, ws, rows="4:8")

        # First call = "4:8" (the Group call)
        assert row_spec_calls[0] == "4:8"
        # Second call = 4 (the int first-row read-back)
        assert row_spec_calls[1] == 4

    def test_group_rows_com_error_wraps_as_tool_error(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Rows.return_value.Group.side_effect = Exception("COM error")
        ms.wrap.side_effect = lambda e, ctx: ToolError(f"{ctx}: {e}")

        with pytest.raises(ToolError):
            _call_outline("group_rows", ms, ws, rows="2:5")


# ── group_columns ─────────────────────────────────────────────────────────────

class TestGroupColumns:
    def test_group_calls_group_on_columns_range(self):
        """VERIFY-EFFECT: ws.Columns('B:D').Group() must be called.

        ws.Columns is called twice: first with 'B:D' (Group), then with 'B' (read-back).
        Use assert_any_call to check the Group call regardless of order.
        """
        ms = make_mock_session()
        ws = _make_sheet_mock()
        cols_rng = MagicMock()
        cols_rng.OutlineLevel = 2
        ws.Columns.return_value = cols_rng

        _call_outline("group_columns", ms, ws, columns="B:D")

        ws.Columns.assert_any_call("B:D")
        cols_rng.Group.assert_called_once()

    def test_group_columns_result_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()

        result = _call_outline("group_columns", ms, ws, columns="C:E")

        assert result["outline"] == "group_columns"
        assert result["sheet"] == "Sheet1"
        assert result["applied"]["columns"] == "C:E"

    def test_group_columns_outline_level_echoed(self):
        """VERIFY-EFFECT: OutlineLevel read-back in applied dict."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        cols_rng = MagicMock()
        cols_rng.OutlineLevel = 2
        ws.Columns.return_value = cols_rng

        result = _call_outline("group_columns", ms, ws, columns="B:D")

        assert result["applied"]["outline_level"] == 2

    def test_group_columns_reads_first_col_letter(self):
        """OutlineLevel read-back uses first col letter from spec."""
        ms = make_mock_session()
        ws = _make_sheet_mock()

        col_spec_calls = []
        first_rng = MagicMock()
        first_rng.OutlineLevel = 2

        def cols_side_effect(spec):
            col_spec_calls.append(spec)
            return first_rng

        ws.Columns.side_effect = cols_side_effect

        _call_outline("group_columns", ms, ws, columns="D:F")

        assert col_spec_calls[0] == "D:F"
        assert col_spec_calls[1] == "D"

    def test_group_columns_com_error_wraps_as_tool_error(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Columns.return_value.Group.side_effect = Exception("COM error")
        ms.wrap.side_effect = lambda e, ctx: ToolError(f"{ctx}: {e}")

        with pytest.raises(ToolError):
            _call_outline("group_columns", ms, ws, columns="B:D")


# ── ungroup_rows ──────────────────────────────────────────────────────────────

class TestUngroupRows:
    def test_ungroup_calls_ungroup_on_rows_range(self):
        """VERIFY-EFFECT: ws.Rows('2:5').Ungroup() must be called.

        ws.Rows is called twice: first with '2:5' (Ungroup), then with 2 (read-back).
        """
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rows_rng = MagicMock()
        rows_rng.OutlineLevel = 1  # after ungroup
        ws.Rows.return_value = rows_rng

        _call_outline("ungroup_rows", ms, ws, rows="2:5")

        ws.Rows.assert_any_call("2:5")
        rows_rng.Ungroup.assert_called_once()

    def test_ungroup_rows_result_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()

        result = _call_outline("ungroup_rows", ms, ws, rows="2:5")

        assert result["outline"] == "ungroup_rows"
        assert result["sheet"] == "Sheet1"
        assert result["applied"]["rows"] == "2:5"
        assert "outline_level" in result["applied"]

    def test_ungroup_rows_level_1_after_ungroup(self):
        """After full ungroup, OutlineLevel returns to 1."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rows_rng = MagicMock()
        rows_rng.OutlineLevel = 1
        ws.Rows.return_value = rows_rng

        result = _call_outline("ungroup_rows", ms, ws, rows="2:5")

        assert result["applied"]["outline_level"] == 1

    def test_ungroup_rows_reads_first_row_for_level(self):
        """OutlineLevel read-back uses first row number parsed from spec (same pattern as group_rows)."""
        ms = make_mock_session()
        ws = _make_sheet_mock()

        row_spec_calls = []
        first_rng = MagicMock()
        first_rng.OutlineLevel = 1

        def rows_side_effect(spec):
            row_spec_calls.append(spec)
            return first_rng

        ws.Rows.side_effect = rows_side_effect

        _call_outline("ungroup_rows", ms, ws, rows="3:7")

        assert row_spec_calls[0] == "3:7"
        assert row_spec_calls[1] == 3

    def test_ungroup_rows_com_error_wraps(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Rows.return_value.Ungroup.side_effect = Exception("RPC error")
        ms.wrap.side_effect = lambda e, ctx: ToolError(str(e))

        with pytest.raises(ToolError):
            _call_outline("ungroup_rows", ms, ws, rows="2:5")


# ── ungroup_columns ───────────────────────────────────────────────────────────

class TestUngroupColumns:
    def test_ungroup_calls_ungroup_on_columns_range(self):
        """VERIFY-EFFECT: ws.Columns('B:D').Ungroup() must be called.

        ws.Columns is called twice: first with 'B:D' (Ungroup), then with 'B' (read-back).
        """
        ms = make_mock_session()
        ws = _make_sheet_mock()
        cols_rng = MagicMock()
        cols_rng.OutlineLevel = 1
        ws.Columns.return_value = cols_rng

        _call_outline("ungroup_columns", ms, ws, columns="B:D")

        ws.Columns.assert_any_call("B:D")
        cols_rng.Ungroup.assert_called_once()

    def test_ungroup_columns_result_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()

        result = _call_outline("ungroup_columns", ms, ws, columns="B:D")

        assert result["outline"] == "ungroup_columns"
        assert result["applied"]["columns"] == "B:D"
        assert "outline_level" in result["applied"]

    def test_ungroup_columns_reads_first_col_letter(self):
        """OutlineLevel read-back uses first col letter from spec (same pattern as group_columns)."""
        ms = make_mock_session()
        ws = _make_sheet_mock()

        col_spec_calls = []
        first_rng = MagicMock()
        first_rng.OutlineLevel = 1

        def cols_side_effect(spec):
            col_spec_calls.append(spec)
            return first_rng

        ws.Columns.side_effect = cols_side_effect

        _call_outline("ungroup_columns", ms, ws, columns="C:E")

        assert col_spec_calls[0] == "C:E"
        assert col_spec_calls[1] == "C"

    def test_ungroup_columns_com_error_wraps(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Columns.return_value.Ungroup.side_effect = Exception("COM error")
        ms.wrap.side_effect = lambda e, ctx: ToolError(str(e))

        with pytest.raises(ToolError):
            _call_outline("ungroup_columns", ms, ws, columns="B:D")


# ── show_levels ───────────────────────────────────────────────────────────────

class TestShowLevels:
    def test_show_row_levels_only(self):
        """ShowLevels(RowLevels=2) called when only row_levels supplied."""
        ms = make_mock_session()
        ws = _make_sheet_mock()

        result = _call_outline("show_levels", ms, ws, row_levels=2)

        ws.Outline.ShowLevels.assert_called_once_with(RowLevels=2)
        assert result["outline"] == "show_levels"
        assert result["applied"]["row_levels"] == 2
        assert "column_levels" not in result["applied"]

    def test_show_column_levels_only(self):
        """ShowLevels(ColumnLevels=1) called when only column_levels supplied."""
        ms = make_mock_session()
        ws = _make_sheet_mock()

        result = _call_outline("show_levels", ms, ws, column_levels=1)

        ws.Outline.ShowLevels.assert_called_once_with(ColumnLevels=1)
        assert result["applied"]["column_levels"] == 1
        assert "row_levels" not in result["applied"]

    def test_show_both_levels(self):
        """ShowLevels(RowLevels=3, ColumnLevels=1) when both supplied."""
        ms = make_mock_session()
        ws = _make_sheet_mock()

        result = _call_outline("show_levels", ms, ws, row_levels=3, column_levels=1)

        ws.Outline.ShowLevels.assert_called_once_with(RowLevels=3, ColumnLevels=1)
        assert result["applied"]["row_levels"] == 3
        assert result["applied"]["column_levels"] == 1

    def test_show_levels_boundary_max(self):
        """Max level 8 is valid."""
        ms = make_mock_session()
        ws = _make_sheet_mock()

        result = _call_outline("show_levels", ms, ws, row_levels=8)

        ws.Outline.ShowLevels.assert_called_once_with(RowLevels=8)

    def test_show_levels_boundary_min(self):
        """Min level 1 is valid."""
        ms = make_mock_session()
        ws = _make_sheet_mock()

        result = _call_outline("show_levels", ms, ws, column_levels=1)

        ws.Outline.ShowLevels.assert_called_once_with(ColumnLevels=1)

    def test_show_levels_result_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("Data")

        result = _call_outline("show_levels", ms, ws, row_levels=2)

        assert result["outline"] == "show_levels"
        assert result["sheet"] == "Data"
        assert isinstance(result["applied"], dict)

    def test_show_levels_com_error_wraps(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Outline.ShowLevels.side_effect = Exception("ShowLevels failed")
        ms.wrap.side_effect = lambda e, ctx: ToolError(str(e))

        with pytest.raises(ToolError):
            _call_outline("show_levels", ms, ws, row_levels=1)


# ── clear ─────────────────────────────────────────────────────────────────────

class TestClear:
    def test_clear_calls_cells_clear_outline(self):
        """VERIFY-EFFECT: ws.Cells.ClearOutline() must be called."""
        ms = make_mock_session()
        ws = _make_sheet_mock()

        _call_outline("clear", ms, ws)

        ws.Cells.ClearOutline.assert_called_once()

    def test_clear_result_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("Summary")

        result = _call_outline("clear", ms, ws)

        assert result["outline"] == "clear"
        assert result["sheet"] == "Summary"
        assert result["applied"]["cleared"] is True

    def test_clear_com_error_wraps(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Cells.ClearOutline.side_effect = Exception("ClearOutline failed")
        ms.wrap.side_effect = lambda e, ctx: ToolError(str(e))

        with pytest.raises(ToolError):
            _call_outline("clear", ms, ws)


# ── Result shape consistency ──────────────────────────────────────────────────

class TestResultShape:
    """Every result dict must carry 'outline', 'sheet', 'applied' keys."""

    def test_group_rows_has_all_keys(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("MySheet")
        result = _call_outline("group_rows", ms, ws, rows="2:5")
        assert "outline" in result
        assert "sheet" in result
        assert "applied" in result
        assert result["sheet"] == "MySheet"

    def test_group_columns_has_all_keys(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_outline("group_columns", ms, ws, columns="B:C")
        assert "outline" in result and "sheet" in result and "applied" in result

    def test_ungroup_rows_has_all_keys(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_outline("ungroup_rows", ms, ws, rows="2:5")
        assert "outline" in result and "sheet" in result and "applied" in result

    def test_ungroup_columns_has_all_keys(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_outline("ungroup_columns", ms, ws, columns="B:C")
        assert "outline" in result and "sheet" in result and "applied" in result

    def test_show_levels_has_all_keys(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_outline("show_levels", ms, ws, row_levels=1)
        assert "outline" in result and "sheet" in result and "applied" in result

    def test_clear_has_all_keys(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_outline("clear", ms, ws)
        assert "outline" in result and "sheet" in result and "applied" in result
