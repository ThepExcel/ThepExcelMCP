"""Unit tests for excel_view (view.py).

No Excel required — all COM calls are intercepted via make_mock_session.

VERIFY-EFFECT discipline: every test asserts the SPECIFIC COM property that
was set to the SPECIFIC value, not just "no exception raised".
"""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock, patch, call

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


# ── Helpers ────────────────────────────────────────────────────────────────────

def _make_sheet_mock(name: str = "Sheet1"):
    """Minimal Worksheet mock sufficient for view.py tests."""
    ws = MagicMock()
    ws.Name = name
    # ws.Range(cell).Row / .Column for freeze_panes cell lookup
    ws.Range.return_value = MagicMock(Row=2, Column=2)  # default: "B2"
    return ws


def _make_window_mock():
    """Window COM mock — all relevant view properties settable."""
    win = MagicMock()
    # Default state: nothing frozen, gridlines/headings visible
    win.FreezePanes = False
    win.SplitRow = 0
    win.SplitColumn = 0
    win.DisplayGridlines = True
    win.DisplayHeadings = True
    win.Zoom = 100
    return win


def _call_view(action, mock_session, mock_ws, mock_win, **kwargs):
    """Patch _session + sheet resolution + window, invoke view_action.

    Window wiring mirrors the fixed _dispatch: wb.Windows(1) is the target
    window (workbook-scoped), NOT app.ActiveWindow (global). This makes the
    harness able to detect the wrong-window class of bug.
    """
    mock_wb = MagicMock()
    # wb.Windows(1) returns the target window — COM call-indexing style
    mock_wb.Windows.return_value = mock_win

    mock_app = MagicMock()
    # app.ActiveWindow deliberately left as a DIFFERENT mock object so that any
    # code still reading app.ActiveWindow will NOT accidentally use mock_win.
    mock_app.ActiveWindow = MagicMock(name="wrong_window")

    mock_ws.Parent = mock_wb
    mock_ws.Application = mock_app

    mock_session.get_sheet.return_value = mock_ws

    with patch("thepexcel_mcp.domains.view._session", mock_session):
        with patch("thepexcel_mcp.domains.view.excel_guard") as eg:
            eg.return_value.__enter__ = MagicMock(return_value=None)
            eg.return_value.__exit__ = MagicMock(return_value=False)
            from thepexcel_mcp.domains.view import view_action
            return view_action(action, sheet=None, workbook=None, **kwargs)


# ── Action validation ──────────────────────────────────────────────────────────

class TestViewActionValidation:
    def test_unknown_action_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()
        with pytest.raises(ToolError, match="Unknown action"):
            _call_view("bogus", ms, ws, win)

    def test_gridlines_requires_show(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()
        with pytest.raises(ToolError, match="show"):
            _call_view("gridlines", ms, ws, win)  # show=None (default)

    def test_zoom_requires_zoom_param(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()
        with pytest.raises(ToolError, match="zoom"):
            _call_view("zoom", ms, ws, win)  # zoom=None (default)

    def test_headings_requires_show(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()
        with pytest.raises(ToolError, match="show"):
            _call_view("headings", ms, ws, win)  # show=None (default)

    def test_zoom_below_min_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()
        with pytest.raises(ToolError, match="10"):
            _call_view("zoom", ms, ws, win, zoom=9)

    def test_zoom_above_max_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()
        with pytest.raises(ToolError, match="400"):
            _call_view("zoom", ms, ws, win, zoom=401)

    def test_freeze_panes_cell_a1_raises(self):
        """Freezing at A1 means nothing above/left — should raise."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.return_value = MagicMock(Row=1, Column=1)  # A1
        win = _make_window_mock()
        with pytest.raises(ToolError, match="A1"):
            _call_view("freeze_panes", ms, ws, win, cell="A1")

    def test_freeze_panes_no_spec_raises(self):
        """No cell, freeze_rows, or freeze_cols → rows=0, cols=0 → error."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()
        with pytest.raises(ToolError, match="freeze_panes"):
            # All default → freeze_rows=None, freeze_cols=None → 0,0
            _call_view("freeze_panes", ms, ws, win)


# ── freeze_panes ───────────────────────────────────────────────────────────────

class TestFreezePanes:
    def test_cell_b2_freezes_row1_col1(self):
        """cell='B2' → SplitRow=1, SplitColumn=1, FreezePanes=True."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.return_value = MagicMock(Row=2, Column=2)  # B2
        win = _make_window_mock()

        result = _call_view("freeze_panes", ms, ws, win, cell="B2")

        # VERIFY EFFECT: correct properties set to specific values
        assert win.SplitRow == 1
        assert win.SplitColumn == 1
        assert win.FreezePanes is True

        assert result["view"] == "freeze_panes"
        assert result["applied"]["freeze_rows"] == 1
        assert result["applied"]["freeze_cols"] == 1
        assert result["applied"]["cell"] == "B2"

    def test_cell_c1_freezes_cols_only(self):
        """cell='C1' → Row=1 → freeze_rows=0, Column=3 → freeze_cols=2."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.return_value = MagicMock(Row=1, Column=3)  # C1
        win = _make_window_mock()

        result = _call_view("freeze_panes", ms, ws, win, cell="C1")

        assert win.SplitRow == 0
        assert win.SplitColumn == 2
        assert win.FreezePanes is True

    def test_freeze_rows_direct(self):
        """freeze_rows=3 without cell → SplitRow=3, SplitColumn=0."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()

        result = _call_view("freeze_panes", ms, ws, win, freeze_rows=3)

        assert win.SplitRow == 3
        assert win.SplitColumn == 0
        assert win.FreezePanes is True
        assert result["applied"]["freeze_rows"] == 3

    def test_freeze_cols_direct(self):
        """freeze_cols=2 without cell → SplitRow=0, SplitColumn=2."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()

        result = _call_view("freeze_panes", ms, ws, win, freeze_cols=2)

        assert win.SplitRow == 0
        assert win.SplitColumn == 2
        assert win.FreezePanes is True

    def test_freeze_unfreezes_first(self):
        """Must set FreezePanes=False BEFORE re-applying to avoid no-op."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.return_value = MagicMock(Row=2, Column=2)
        win = _make_window_mock()

        # Track property assignments in order
        calls = []
        type(win).__setattr__ = lambda self, name, val, calls=calls: (
            calls.append((name, val)) or object.__setattr__(self, name, val)
        )

        _call_view("freeze_panes", ms, ws, win, cell="B2")

        # FreezePanes=False must appear before FreezePanes=True in call sequence
        fp_calls = [(n, v) for n, v in calls if n == "FreezePanes"]
        assert fp_calls[0] == ("FreezePanes", False), (
            "FreezePanes must be set to False before True (re-freeze pattern)"
        )
        assert fp_calls[-1] == ("FreezePanes", True)

    def test_invalid_cell_raises(self):
        """An invalid cell reference must raise ToolError, not COM crash."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.side_effect = Exception("Invalid range")
        win = _make_window_mock()

        with pytest.raises(ToolError, match="Invalid cell reference"):
            _call_view("freeze_panes", ms, ws, win, cell="NOTACELL")

    def test_sheet_activated_before_window(self):
        """ws.Activate() must be called to select the sheet; wb.Activate() is
        also called first to bring the workbook to the foreground."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.return_value = MagicMock(Row=2, Column=2)
        win = _make_window_mock()

        _call_view("freeze_panes", ms, ws, win, cell="B2")

        ws.Activate.assert_called_once()
        # wb.Activate() must also have been called
        ws.Parent.Activate.assert_called_once()


# ── unfreeze_panes ────────────────────────────────────────────────────────────

class TestUnfreezePanes:
    def test_unfreeze_sets_false_and_zeros_splits(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()

        result = _call_view("unfreeze_panes", ms, ws, win)

        # VERIFY EFFECT: FreezePanes=False and splits reset to 0
        assert win.FreezePanes is False
        assert win.SplitRow == 0
        assert win.SplitColumn == 0

        assert result["view"] == "unfreeze_panes"
        assert result["applied"]["frozen"] is False

    def test_sheet_activated(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()

        _call_view("unfreeze_panes", ms, ws, win)

        ws.Activate.assert_called_once()


# ── gridlines ─────────────────────────────────────────────────────────────────

class TestGridlines:
    def test_hide_gridlines_sets_false(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()

        result = _call_view("gridlines", ms, ws, win, show=False)

        assert win.DisplayGridlines is False
        assert result["view"] == "gridlines"
        assert result["applied"]["show"] is False

    def test_show_gridlines_sets_true(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()
        win.DisplayGridlines = False  # start hidden

        result = _call_view("gridlines", ms, ws, win, show=True)

        assert win.DisplayGridlines is True
        assert result["applied"]["show"] is True

    def test_sheet_activated(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()

        _call_view("gridlines", ms, ws, win, show=True)

        ws.Activate.assert_called_once()


# ── zoom ──────────────────────────────────────────────────────────────────────

class TestZoom:
    def test_zoom_150_sets_property(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()

        result = _call_view("zoom", ms, ws, win, zoom=150)

        assert win.Zoom == 150
        assert result["view"] == "zoom"
        assert result["applied"]["zoom"] == 150

    def test_zoom_boundary_min(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()

        _call_view("zoom", ms, ws, win, zoom=10)
        assert win.Zoom == 10

    def test_zoom_boundary_max(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()

        _call_view("zoom", ms, ws, win, zoom=400)
        assert win.Zoom == 400

    def test_zoom_stores_as_int(self):
        """Zoom must be passed as integer (not float) to COM."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()

        _call_view("zoom", ms, ws, win, zoom=75)
        assert isinstance(win.Zoom, int)

    def test_sheet_activated(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()

        _call_view("zoom", ms, ws, win, zoom=100)

        ws.Activate.assert_called_once()


# ── headings ──────────────────────────────────────────────────────────────────

class TestHeadings:
    def test_hide_headings_sets_false(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()

        result = _call_view("headings", ms, ws, win, show=False)

        assert win.DisplayHeadings is False
        assert result["view"] == "headings"
        assert result["applied"]["show"] is False

    def test_show_headings_sets_true(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()
        win.DisplayHeadings = False

        result = _call_view("headings", ms, ws, win, show=True)

        assert win.DisplayHeadings is True
        assert result["applied"]["show"] is True

    def test_sheet_activated(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()

        _call_view("headings", ms, ws, win, show=True)

        ws.Activate.assert_called_once()


# ── Result shape ───────────────────────────────────────────────────────────────

class TestResultShape:
    """Every result dict must carry 'view', 'sheet', 'applied' keys."""

    def test_freeze_panes_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("MySheet")
        ws.Range.return_value = MagicMock(Row=2, Column=2)
        win = _make_window_mock()
        result = _call_view("freeze_panes", ms, ws, win, cell="B2")
        assert "view" in result
        assert "sheet" in result
        assert "applied" in result
        assert result["sheet"] == "MySheet"

    def test_unfreeze_panes_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("Data")
        win = _make_window_mock()
        result = _call_view("unfreeze_panes", ms, ws, win)
        assert result["sheet"] == "Data"
        assert isinstance(result["applied"], dict)

    def test_gridlines_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()
        result = _call_view("gridlines", ms, ws, win, show=False)
        assert "view" in result and "applied" in result

    def test_zoom_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()
        result = _call_view("zoom", ms, ws, win, zoom=80)
        assert "view" in result and "applied" in result

    def test_headings_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        win = _make_window_mock()
        result = _call_view("headings", ms, ws, win, show=True)
        assert "view" in result and "applied" in result


# ── Wrong-window regression ───────────────────────────────────────────────────

class TestWrongWindowRegression:
    """Guard against the bug where app.ActiveWindow points to a different
    workbook's window than the target workbook, causing silent no-ops.

    Each test constructs the mismatch explicitly:
      - win_target  = wb.Windows(1)   (the window we MUST mutate)
      - win_other   = app.ActiveWindow (a bystander workbook's window)

    After the action, win_target must be mutated and win_other must be
    unchanged. This fails against the old impl (ws.Activate(); app.ActiveWindow)
    and passes only after the wb.Windows(1) fix.
    """

    def _make_mismatch_session(self, name: str = "Sheet1"):
        """Return (mock_session, ws, win_target, win_other) with ActiveWindow
        pointing at a different mock than wb.Windows(1)."""
        ms = make_mock_session()

        ws = MagicMock()
        ws.Name = name
        ws.Range.return_value = MagicMock(Row=2, Column=2)  # B2

        win_target = _make_window_mock()   # the correct workbook window
        win_other = MagicMock(name="win_other")  # bystander window

        mock_wb = MagicMock()
        mock_wb.Windows.return_value = win_target  # wb.Windows(1) → target

        mock_app = MagicMock()
        mock_app.ActiveWindow = win_other  # global ActiveWindow → wrong window

        ws.Parent = mock_wb
        ws.Application = mock_app

        ms.get_sheet.return_value = ws
        return ms, ws, win_target, win_other

    def _invoke(self, action, ms, ws, **kwargs):
        with patch("thepexcel_mcp.domains.view._session", ms):
            with patch("thepexcel_mcp.domains.view.excel_guard") as eg:
                eg.return_value.__enter__ = MagicMock(return_value=None)
                eg.return_value.__exit__ = MagicMock(return_value=False)
                from thepexcel_mcp.domains.view import view_action
                return view_action(action, sheet=None, workbook=None, **kwargs)

    def test_freeze_panes_mutates_target_window_not_active_window(self):
        ms, ws, win_target, win_other = self._make_mismatch_session()

        self._invoke("freeze_panes", ms, ws, cell="B2")

        # Target window must be frozen
        assert win_target.FreezePanes is True
        assert win_target.SplitRow == 1
        assert win_target.SplitColumn == 1
        # Bystander window must be untouched: FreezePanes on an unset MagicMock
        # auto-attribute is a MagicMock object, never True or False.
        assert win_other.FreezePanes is not True
        assert win_other.FreezePanes is not False

    def test_gridlines_mutates_target_window_not_active_window(self):
        ms, ws, win_target, win_other = self._make_mismatch_session()

        self._invoke("gridlines", ms, ws, show=False)

        assert win_target.DisplayGridlines is False
        # win_other.DisplayGridlines should never have been set to False
        # (MagicMock attribute access returns a new Mock, not False)
        assert win_other.DisplayGridlines is not False

    def test_zoom_mutates_target_window_not_active_window(self):
        ms, ws, win_target, win_other = self._make_mismatch_session()

        self._invoke("zoom", ms, ws, zoom=150)

        assert win_target.Zoom == 150
        # win_other.Zoom is an auto-created MagicMock attribute, never set to 150
        assert win_other.Zoom != 150

    def test_headings_mutates_target_window_not_active_window(self):
        ms, ws, win_target, win_other = self._make_mismatch_session()

        self._invoke("headings", ms, ws, show=False)

        assert win_target.DisplayHeadings is False
        assert win_other.DisplayHeadings is not False

    def test_unfreeze_mutates_target_window_not_active_window(self):
        ms, ws, win_target, win_other = self._make_mismatch_session()

        self._invoke("unfreeze_panes", ms, ws)

        assert win_target.FreezePanes is False
        assert win_target.SplitRow == 0
        assert win_target.SplitColumn == 0
        # win_other properties were never touched (remain auto-MagicMock)
        assert win_other.FreezePanes is not False
