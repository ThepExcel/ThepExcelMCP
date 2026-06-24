"""Unit tests for excel_protection (protection.py).

No Excel required — all COM calls are intercepted via make_mock_session.

VERIFY-EFFECT discipline: every mutation test asserts the SPECIFIC COM property
or method call that was triggered, not just "no exception raised".

Coverage
--------
- Action validation (unknown action → ToolError listing valid actions)
- protect_sheet: COM Protect() called with correct kwargs; ws.ProtectContents read-back
- unprotect_sheet: COM Unprotect() called; ProtectContents False after; 1004 → ToolError
- protect_workbook: wb.Protect() called; ProtectStructure read-back
- unprotect_workbook: wb.Unprotect() called; ProtectStructure False; 1004 → ToolError
- set_locked: rng.Locked / rng.FormulaHidden set + read-back; missing-range error
- status: reads all 5 protection flags, returns them in applied dict
- allow flags dict vs individual bool params
- Password omitted when None; Password passed when given
"""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock, patch, call

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


# ── Helpers ────────────────────────────────────────────────────────────────────

def _make_sheet_mock(name: str = "Sheet1"):
    ws = MagicMock()
    ws.Name = name
    # Default: sheet not protected
    ws.ProtectContents = False
    ws.ProtectDrawingObjects = False
    ws.ProtectScenarios = False
    return ws


def _make_workbook_mock(name: str = "Book1.xlsx"):
    wb = MagicMock()
    wb.Name = name
    wb.ProtectStructure = False
    wb.ProtectWindows = False
    return wb


def _call_protection(action, mock_session, mock_ws, mock_wb=None, **kwargs):
    """Patch _session + COM objects, invoke protection_action."""
    if mock_wb is None:
        mock_wb = _make_workbook_mock()

    mock_app = MagicMock()
    mock_ws.Parent = mock_wb
    mock_ws.Application = mock_app
    mock_session.get_sheet.return_value = mock_ws

    with patch("thepexcel_mcp.domains.protection._session", mock_session):
        with patch("thepexcel_mcp.domains.protection.excel_guard") as eg:
            eg.return_value.__enter__ = MagicMock(return_value=None)
            eg.return_value.__exit__ = MagicMock(return_value=False)
            from thepexcel_mcp.domains.protection import protection_action
            return protection_action(action, sheet=None, workbook=None, **kwargs)


# ── Action validation ──────────────────────────────────────────────────────────

class TestActionValidation:
    def test_unknown_action_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="Unknown action"):
            _call_protection("bogus_action", ms, ws)

    def test_unknown_action_lists_valid(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError) as exc_info:
            _call_protection("foo", ms, ws)
        msg = str(exc_info.value)
        assert "protect_sheet" in msg
        assert "status" in msg


# ── protect_sheet ──────────────────────────────────────────────────────────────

class TestProtectSheet:
    def test_protect_calls_ws_protect(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        # Simulate Protect() making ws.ProtectContents = True
        def _protect(**kw):
            ws.ProtectContents = True
        ws.Protect.side_effect = _protect

        _call_protection("protect_sheet", ms, ws)

        ws.Protect.assert_called_once()

    def test_protect_passes_contents_true_by_default(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        def _protect(**kw):
            ws.ProtectContents = True
        ws.Protect.side_effect = _protect

        _call_protection("protect_sheet", ms, ws)

        _, kwargs = ws.Protect.call_args
        assert kwargs.get("Contents") is True

    def test_protect_drawing_objects_default_true(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        def _protect(**kw):
            ws.ProtectContents = True
        ws.Protect.side_effect = _protect

        _call_protection("protect_sheet", ms, ws)

        _, kwargs = ws.Protect.call_args
        assert kwargs.get("DrawingObjects") is True

    def test_protect_scenarios_default_true(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        def _protect(**kw):
            ws.ProtectContents = True
        ws.Protect.side_effect = _protect

        _call_protection("protect_sheet", ms, ws)

        _, kwargs = ws.Protect.call_args
        assert kwargs.get("Scenarios") is True

    def test_protect_with_password(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        def _protect(**kw):
            ws.ProtectContents = True
        ws.Protect.side_effect = _protect

        _call_protection("protect_sheet", ms, ws, password="secret")

        _, kwargs = ws.Protect.call_args
        assert kwargs.get("Password") == "secret"

    def test_protect_without_password_omits_kwarg(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        def _protect(**kw):
            ws.ProtectContents = True
        ws.Protect.side_effect = _protect

        _call_protection("protect_sheet", ms, ws)  # password=None default

        _, kwargs = ws.Protect.call_args
        assert "Password" not in kwargs

    def test_protect_allow_flag_individual_param(self):
        """allow_filtering=True → AllowFiltering=True in COM call."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        def _protect(**kw):
            ws.ProtectContents = True
        ws.Protect.side_effect = _protect

        _call_protection("protect_sheet", ms, ws, allow_filtering=True)

        _, kwargs = ws.Protect.call_args
        assert kwargs.get("AllowFiltering") is True

    def test_protect_allow_flag_via_dict(self):
        """allow={'AllowSorting': True} → AllowSorting=True in COM call."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        def _protect(**kw):
            ws.ProtectContents = True
        ws.Protect.side_effect = _protect

        _call_protection("protect_sheet", ms, ws, allow={"AllowSorting": True})

        _, kwargs = ws.Protect.call_args
        assert kwargs.get("AllowSorting") is True

    def test_protect_allow_dict_falls_back_to_individual_for_unlisted_flags(self):
        """allow dict + individual param: dict key wins; unlisted flags fall back to individual bool."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        def _protect(**kw):
            ws.ProtectContents = True
        ws.Protect.side_effect = _protect

        # allow dict only specifies Sorting; allow_filtering via individual param
        _call_protection(
            "protect_sheet", ms, ws,
            allow={"AllowSorting": True},
            allow_filtering=True,
        )

        _, kwargs = ws.Protect.call_args
        assert kwargs.get("AllowSorting") is True      # from dict
        assert kwargs.get("AllowFiltering") is True    # fell back to individual param

    def test_protect_verify_effect_reads_protect_contents(self):
        """Result dict must include verify.ProtectContents == True."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        def _protect(**kw):
            ws.ProtectContents = True
        ws.Protect.side_effect = _protect

        result = _call_protection("protect_sheet", ms, ws)

        assert result["verify"]["ProtectContents"] is True
        assert result["protection"] == "protect_sheet"
        assert "applied" in result

    def test_protect_protect_contents_still_false_raises(self):
        """If ProtectContents remains False after Protect(), raise ToolError."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        # ws.ProtectContents stays False — simulates silent no-op
        ws.Protect.side_effect = lambda **kw: None

        with pytest.raises(ToolError, match="ProtectContents"):
            _call_protection("protect_sheet", ms, ws)

    def test_protect_result_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("MySheet")
        def _protect(**kw):
            ws.ProtectContents = True
        ws.Protect.side_effect = _protect

        result = _call_protection("protect_sheet", ms, ws)

        assert result["protection"] == "protect_sheet"
        assert result["sheet"] == "MySheet"
        assert isinstance(result["applied"], dict)
        assert isinstance(result["verify"], dict)


# ── unprotect_sheet ────────────────────────────────────────────────────────────

class TestUnprotectSheet:
    def test_unprotect_calls_ws_unprotect(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.ProtectContents = False  # already unprotected after call

        _call_protection("unprotect_sheet", ms, ws)

        ws.Unprotect.assert_called_once()

    def test_unprotect_with_password(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.ProtectContents = False

        _call_protection("unprotect_sheet", ms, ws, password="secret")

        ws.Unprotect.assert_called_once_with(Password="secret")

    def test_unprotect_without_password_no_kwarg(self):
        """password=None → call ws.Unprotect() with no args."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.ProtectContents = False

        _call_protection("unprotect_sheet", ms, ws)

        ws.Unprotect.assert_called_once_with()

    def test_unprotect_verify_effect_protect_contents_false(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.ProtectContents = False

        result = _call_protection("unprotect_sheet", ms, ws)

        assert result["verify"]["ProtectContents"] is False
        assert result["applied"]["protected"] is False

    def test_unprotect_wrong_password_1004_raises_tool_error(self):
        """COM error 1004 → ToolError mentioning 'wrong password'."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Unprotect.side_effect = Exception("(-2146827284, '1004', ...)")

        with pytest.raises(ToolError, match="wrong password"):
            _call_protection("unprotect_sheet", ms, ws, password="wrong")

    def test_unprotect_protect_contents_still_true_raises(self):
        """If ProtectContents stays True after Unprotect(), raise ToolError."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.ProtectContents = True  # still True after call — simulate no-op
        ws.Unprotect.side_effect = lambda **kw: None

        with pytest.raises(ToolError, match="ProtectContents"):
            _call_protection("unprotect_sheet", ms, ws)


# ── protect_workbook ───────────────────────────────────────────────────────────

class TestProtectWorkbook:
    def test_protect_wb_calls_wb_protect(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        wb = _make_workbook_mock()
        def _protect(**kw):
            wb.ProtectStructure = True
        wb.Protect.side_effect = _protect

        _call_protection("protect_workbook", ms, ws, mock_wb=wb)

        wb.Protect.assert_called_once()

    def test_protect_wb_structure_default_true(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        wb = _make_workbook_mock()
        def _protect(**kw):
            wb.ProtectStructure = True
        wb.Protect.side_effect = _protect

        _call_protection("protect_workbook", ms, ws, mock_wb=wb)

        _, kwargs = wb.Protect.call_args
        assert kwargs.get("Structure") is True

    def test_protect_wb_windows_default_false(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        wb = _make_workbook_mock()
        def _protect(**kw):
            wb.ProtectStructure = True
        wb.Protect.side_effect = _protect

        _call_protection("protect_workbook", ms, ws, mock_wb=wb)

        _, kwargs = wb.Protect.call_args
        assert kwargs.get("Windows") is False

    def test_protect_wb_with_password(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        wb = _make_workbook_mock()
        def _protect(**kw):
            wb.ProtectStructure = True
        wb.Protect.side_effect = _protect

        _call_protection("protect_workbook", ms, ws, mock_wb=wb, password="pw123")

        _, kwargs = wb.Protect.call_args
        assert kwargs.get("Password") == "pw123"

    def test_protect_wb_no_password_omits_kwarg(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        wb = _make_workbook_mock()
        def _protect(**kw):
            wb.ProtectStructure = True
        wb.Protect.side_effect = _protect

        _call_protection("protect_workbook", ms, ws, mock_wb=wb)

        _, kwargs = wb.Protect.call_args
        assert "Password" not in kwargs

    def test_protect_wb_verify_effect(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        wb = _make_workbook_mock()
        def _protect(**kw):
            wb.ProtectStructure = True
        wb.Protect.side_effect = _protect

        result = _call_protection("protect_workbook", ms, ws, mock_wb=wb)

        assert result["verify"]["ProtectStructure"] is True
        assert result["protection"] == "protect_workbook"

    def test_protect_wb_structure_still_false_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        wb = _make_workbook_mock()
        wb.Protect.side_effect = lambda **kw: None  # wb.ProtectStructure stays False

        with pytest.raises(ToolError, match="ProtectStructure"):
            _call_protection("protect_workbook", ms, ws, mock_wb=wb)

    def test_protect_wb_structure_false_does_not_raise(self):
        """structure=False is a valid call; guard must NOT fire when ProtectStructure is False."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        wb = _make_workbook_mock()
        # ProtectStructure stays False — that IS the correct outcome for structure=False
        wb.Protect.side_effect = lambda **kw: None

        # Must NOT raise — the verify-effect guard is only meaningful when structure=True
        result = _call_protection("protect_workbook", ms, ws, mock_wb=wb, structure=False)
        assert result["applied"]["structure"] is False


# ── unprotect_workbook ─────────────────────────────────────────────────────────

class TestUnprotectWorkbook:
    def test_unprotect_wb_calls_wb_unprotect(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        wb = _make_workbook_mock()
        wb.ProtectStructure = False

        _call_protection("unprotect_workbook", ms, ws, mock_wb=wb)

        wb.Unprotect.assert_called_once()

    def test_unprotect_wb_with_password(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        wb = _make_workbook_mock()
        wb.ProtectStructure = False

        _call_protection("unprotect_workbook", ms, ws, mock_wb=wb, password="pw")

        wb.Unprotect.assert_called_once_with(Password="pw")

    def test_unprotect_wb_no_password_no_kwarg(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        wb = _make_workbook_mock()
        wb.ProtectStructure = False

        _call_protection("unprotect_workbook", ms, ws, mock_wb=wb)

        wb.Unprotect.assert_called_once_with()

    def test_unprotect_wb_wrong_password_raises_tool_error(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        wb = _make_workbook_mock()
        wb.Unprotect.side_effect = Exception("1004 workbook is protected")

        with pytest.raises(ToolError, match="wrong password"):
            _call_protection("unprotect_workbook", ms, ws, mock_wb=wb, password="bad")

    def test_unprotect_wb_verify_effect(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        wb = _make_workbook_mock()
        wb.ProtectStructure = False

        result = _call_protection("unprotect_workbook", ms, ws, mock_wb=wb)

        assert result["verify"]["ProtectStructure"] is False
        assert result["applied"]["protected"] is False


# ── set_locked ────────────────────────────────────────────────────────────────

class TestSetLocked:
    def _make_range_mock(self, locked=True, formula_hidden=False):
        rng = MagicMock()
        rng.Address = "$A$1:$B$5"
        rng.Locked = locked
        rng.FormulaHidden = formula_hidden
        return rng

    def test_set_locked_true_sets_com_property(self):
        """Mock starts locked=False; code must set it True for verify to pass."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rng = self._make_range_mock(locked=False)  # start opposite so assignment is observable
        ws.Range.return_value = rng

        result = _call_protection("set_locked", ms, ws, range="A1:B5", locked=True)

        assert rng.Locked is True      # domain code must have done rng.Locked = True
        assert result["verify"]["Locked"] is True
        assert result["applied"]["locked"] is True

    def test_set_locked_false_sets_com_property(self):
        """Mock starts locked=True; code must set it False for verify to pass."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rng = self._make_range_mock(locked=True)  # start opposite so assignment is observable
        ws.Range.return_value = rng

        result = _call_protection("set_locked", ms, ws, range="A1:B5", locked=False)

        assert rng.Locked is False
        assert result["verify"]["Locked"] is False

    def test_set_locked_zero_is_valid_falsy_guard(self):
        """locked=False (falsy) must NOT be mistaken for 'not provided'."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rng = self._make_range_mock(locked=False)
        ws.Range.return_value = rng

        # Must not raise — False is a valid value for locked
        result = _call_protection("set_locked", ms, ws, range="A1", locked=False)
        assert "Locked" in result["verify"]

    def test_set_formula_hidden_true(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rng = self._make_range_mock(formula_hidden=True)
        ws.Range.return_value = rng

        result = _call_protection("set_locked", ms, ws, range="A1", hidden=True)

        assert rng.FormulaHidden is True
        assert result["verify"]["FormulaHidden"] is True

    def test_set_locked_and_hidden_together(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rng = self._make_range_mock(locked=True, formula_hidden=True)
        ws.Range.return_value = rng

        result = _call_protection("set_locked", ms, ws, range="A1:C3", locked=True, hidden=True)

        assert "Locked" in result["verify"]
        assert "FormulaHidden" in result["verify"]

    def test_set_locked_result_contains_note(self):
        """Result must include a note explaining that Locked needs sheet protection."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rng = self._make_range_mock()
        ws.Range.return_value = rng

        result = _call_protection("set_locked", ms, ws, range="A1", locked=True)

        note = result["applied"].get("note", "")
        assert "protect" in note.lower()

    def test_set_locked_missing_range_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="range"):
            _call_protection("set_locked", ms, ws, locked=True)  # range=None

    def test_set_locked_neither_locked_nor_hidden_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="locked"):
            _call_protection("set_locked", ms, ws, range="A1")  # both None

    def test_set_locked_invalid_range_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.side_effect = Exception("Invalid range reference")

        with pytest.raises(ToolError, match="Invalid range"):
            _call_protection("set_locked", ms, ws, range="NOTVALID", locked=True)

    def test_set_locked_verify_effect_raises_when_assignment_silently_ignored(self):
        """If COM silently rejects the Locked assignment (e.g. sheet protected),
        the read-back will differ from requested value → must raise ToolError."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        # Simulate a protected sheet: Locked write is ignored, stays True
        rng = self._make_range_mock(locked=True)

        # Override __setattr__ to make rng.Locked ignore assignment (stays True)
        original_locked = True

        class _StubbornRange:
            """Mimics a range where Locked is read-only (sheet protected)."""
            def __init__(self):
                self.Address = "$A$1"
                self.FormulaHidden = False

            @property
            def Locked(self):
                return original_locked  # always True regardless of writes

            @Locked.setter
            def Locked(self, value):
                pass  # silently ignored — sheet is protected

        stubborn = _StubbornRange()
        ws.Range.return_value = stubborn

        with pytest.raises(ToolError, match="Locked"):
            _call_protection("set_locked", ms, ws, range="A1", locked=False)

    def test_set_locked_formula_hidden_verify_raises_when_silently_ignored(self):
        """FormulaHidden write silently ignored → ToolError on mismatch."""
        ms = make_mock_session()
        ws = _make_sheet_mock()

        class _StubbornRange:
            def __init__(self):
                self.Address = "$A$1"
                self.Locked = True

            @property
            def FormulaHidden(self):
                return False  # always False

            @FormulaHidden.setter
            def FormulaHidden(self, value):
                pass

        ws.Range.return_value = _StubbornRange()

        with pytest.raises(ToolError, match="FormulaHidden"):
            _call_protection("set_locked", ms, ws, range="A1", hidden=True)


# ── status ────────────────────────────────────────────────────────────────────

class TestStatus:
    def test_status_reads_sheet_protect_contents(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        wb = _make_workbook_mock()
        ws.ProtectContents = True
        ws.ProtectDrawingObjects = True
        ws.ProtectScenarios = True
        wb.ProtectStructure = False
        wb.ProtectWindows = False

        result = _call_protection("status", ms, ws, mock_wb=wb)

        assert result["applied"]["sheet_protected"] is True
        assert result["applied"]["drawing_objects"] is True
        assert result["applied"]["scenarios"] is True

    def test_status_reads_workbook_protect_structure(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        wb = _make_workbook_mock()
        wb.ProtectStructure = True
        wb.ProtectWindows = True

        result = _call_protection("status", ms, ws, mock_wb=wb)

        assert result["applied"]["workbook_structure"] is True
        assert result["applied"]["workbook_windows"] is True

    def test_status_all_false_when_unprotected(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        wb = _make_workbook_mock()
        # Both mocks default to False for protection flags

        result = _call_protection("status", ms, ws, mock_wb=wb)

        assert result["applied"]["sheet_protected"] is False
        assert result["applied"]["workbook_structure"] is False

    def test_status_result_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("DataSheet")
        wb = _make_workbook_mock("Sales.xlsx")

        result = _call_protection("status", ms, ws, mock_wb=wb)

        assert result["protection"] == "status"
        assert result["sheet"] == "DataSheet"
        assert result["workbook"] == "Sales.xlsx"
        applied = result["applied"]
        for key in ("sheet_protected", "drawing_objects", "scenarios",
                    "workbook_structure", "workbook_windows"):
            assert key in applied, f"Missing key '{key}' in applied"

    def test_status_no_mutations(self):
        """status must not call Protect/Unprotect — read-only."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        wb = _make_workbook_mock()

        _call_protection("status", ms, ws, mock_wb=wb)

        ws.Protect.assert_not_called()
        ws.Unprotect.assert_not_called()
        wb.Protect.assert_not_called()
        wb.Unprotect.assert_not_called()
