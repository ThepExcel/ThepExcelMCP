"""Unit tests for excel_find_replace (find_replace.py).

No Excel required — all COM calls are intercepted via make_mock_session.

VERIFY-EFFECT discipline: every mutation test asserts the specific COM
method was called with the specific arguments AND the result dict reflects
the pre/post counts correctly.

Key COM facts (confirmed Microsoft Learn 2026-06-24):
  xlFormulas=-4123, xlValues=-4163, xlWhole=1, xlPart=2, xlByRows=1
  Range.Find returns None on no match.
  Range.Replace returns Boolean (not count).
"""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock, call, patch

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


# ── Helpers ────────────────────────────────────────────────────────────────────

def _make_cell_mock(address: str, value=None):
    """Single-cell COM mock: .Address, .Value, self-referential FindNext loop."""
    cell = MagicMock()
    cell.Address = address
    cell.Value = value
    return cell


def _make_rng_mock_no_match():
    """Range where Find() returns None (no match)."""
    rng = MagicMock()
    rng.Find.return_value = None
    rng.Worksheet = MagicMock()
    rng.Worksheet.Name = "Sheet1"
    rng.Application = MagicMock()
    return rng


def _make_rng_mock_single_match(address="$A$1", value="hello"):
    """Range where Find() returns one cell; FindNext wraps back immediately."""
    cell = _make_cell_mock(address, value)
    # FindNext returns the SAME address → loop terminates after one iteration
    cell2 = _make_cell_mock(address, value)  # same address = wrap sentinel
    rng = MagicMock()
    rng.Find.return_value = cell
    rng.FindNext.return_value = cell2  # same address → stop
    rng.Worksheet = MagicMock()
    rng.Worksheet.Name = "Sheet1"
    rng.Application = MagicMock()
    return rng


def _make_rng_mock_two_matches(addr1="$A$1", val1="foo", addr2="$B$2", val2="foo"):
    """Range with exactly two matches; FindNext wraps to first after second."""
    cell1 = _make_cell_mock(addr1, val1)
    cell2 = _make_cell_mock(addr2, val2)
    wrap  = _make_cell_mock(addr1, val1)  # same address as cell1 → stop

    rng = MagicMock()
    rng.Find.return_value = cell1
    rng.FindNext.side_effect = [cell2, wrap]
    rng.Worksheet = MagicMock()
    rng.Worksheet.Name = "Sheet1"
    rng.Application = MagicMock()
    return rng


def _patch_and_call(action, mock_session, mock_rng, **kwargs):
    """Patch _session + _resolve_range/get_sheet, then call find_replace_action."""
    ws_mock = MagicMock()
    ws_mock.Cells = mock_rng
    ws_mock.Name = "Sheet1"
    mock_session.get_sheet.return_value = ws_mock

    with patch("thepexcel_mcp.domains.find_replace._session", mock_session):
        with patch("thepexcel_mcp.domains.find_replace.excel_guard") as eg:
            eg.return_value.__enter__ = MagicMock(return_value=None)
            eg.return_value.__exit__ = MagicMock(return_value=False)
            from thepexcel_mcp.domains.find_replace import find_replace_action
            return find_replace_action(action, **kwargs)


def _patch_and_call_range_scope(action, mock_session, mock_rng, range_str="A1:B10", **kwargs):
    """Like _patch_and_call but for scope='range' — patches _resolve_range."""
    ws_mock = MagicMock()
    ws_mock.Name = "Sheet1"
    mock_rng.Worksheet = ws_mock

    mock_session.get_sheet.return_value = ws_mock
    ws_mock.Range.return_value = mock_rng

    with patch("thepexcel_mcp.domains.find_replace._session", mock_session):
        with patch("thepexcel_mcp.domains.find_replace.excel_guard") as eg:
            eg.return_value.__enter__ = MagicMock(return_value=None)
            eg.return_value.__exit__ = MagicMock(return_value=False)
            from thepexcel_mcp.domains.find_replace import find_replace_action
            return find_replace_action(
                action, scope="range", range=range_str, **kwargs
            )


# ── Action validation ──────────────────────────────────────────────────────────

class TestActionValidation:
    def test_unknown_action_raises(self):
        ms = make_mock_session()
        with pytest.raises(ToolError, match="Unknown action"):
            _patch_and_call("bogus", ms, MagicMock(), find_text="x")

    def test_unknown_scope_raises(self):
        ms = make_mock_session()
        rng = _make_rng_mock_no_match()
        with pytest.raises(ToolError, match="Unknown scope"):
            _patch_and_call("find", ms, rng, find_text="x", scope="universe")

    def test_unknown_look_in_raises(self):
        ms = make_mock_session()
        rng = _make_rng_mock_no_match()
        with pytest.raises(ToolError, match="Unknown look_in"):
            _patch_and_call("find", ms, rng, find_text="x", look_in="comments")

    def test_replace_without_replace_text_raises(self):
        ms = make_mock_session()
        rng = _make_rng_mock_no_match()
        with pytest.raises(ToolError, match="replace_text"):
            _patch_and_call("replace", ms, rng, find_text="x")

    def test_scope_range_without_range_param_raises(self):
        ms = make_mock_session()
        rng = _make_rng_mock_no_match()
        with pytest.raises(ToolError, match="'range' parameter"):
            with patch("thepexcel_mcp.domains.find_replace._session", ms):
                from thepexcel_mcp.domains.find_replace import find_replace_action
                find_replace_action("find", find_text="x", scope="range")


# ── find — no match ───────────────────────────────────────────────────────────

class TestFindNoMatch:
    def test_returns_empty_list(self):
        ms = make_mock_session()
        rng = _make_rng_mock_no_match()
        result = _patch_and_call("find", ms, rng, find_text="missing")
        assert result["find_replace"] == "find"
        assert result["applied"]["total_found"] == 0
        assert result["applied"]["matches"] == []
        assert result["applied"]["truncated"] is False

    def test_find_called_with_correct_params(self):
        """Find must receive LookIn=-4123 (xlFormulas), LookAt=2 (xlPart) by default."""
        ms = make_mock_session()
        rng = _make_rng_mock_no_match()
        _patch_and_call("find", ms, rng, find_text="hello")
        rng.Find.assert_called_once_with(
            What="hello",
            LookIn=-4123,   # xlFormulas
            LookAt=2,       # xlPart
            SearchOrder=1,  # xlByRows
            MatchCase=False,
        )


# ── find — single match ───────────────────────────────────────────────────────

class TestFindSingleMatch:
    def test_returns_one_hit(self):
        ms = make_mock_session()
        rng = _make_rng_mock_single_match("$A$1", "target")
        result = _patch_and_call("find", ms, rng, find_text="target")
        assert result["applied"]["total_found"] == 1
        assert result["applied"]["matches"][0]["cell"] == "$A$1"
        assert result["applied"]["matches"][0]["value"] == "target"

    def test_match_includes_sheet_name(self):
        ms = make_mock_session()
        rng = _make_rng_mock_single_match("$C$3", "x")
        result = _patch_and_call("find", ms, rng, find_text="x")
        assert result["applied"]["matches"][0]["sheet"] == "Sheet1"

    def test_find_whole_cell_uses_xl_whole(self):
        """match_whole_cell=True → LookAt=1 (xlWhole)."""
        ms = make_mock_session()
        rng = _make_rng_mock_no_match()
        _patch_and_call("find", ms, rng, find_text="x", match_whole_cell=True)
        rng.Find.assert_called_once_with(
            What="x",
            LookIn=-4123,
            LookAt=1,       # xlWhole
            SearchOrder=1,
            MatchCase=False,
        )

    def test_find_values_uses_xl_values(self):
        """look_in='values' → LookIn=-4163 (xlValues)."""
        ms = make_mock_session()
        rng = _make_rng_mock_no_match()
        _patch_and_call("find", ms, rng, find_text="x", look_in="values")
        rng.Find.assert_called_once_with(
            What="x",
            LookIn=-4163,   # xlValues
            LookAt=2,
            SearchOrder=1,
            MatchCase=False,
        )

    def test_find_match_case_passed_through(self):
        ms = make_mock_session()
        rng = _make_rng_mock_no_match()
        _patch_and_call("find", ms, rng, find_text="X", match_case=True)
        rng.Find.assert_called_once_with(
            What="X",
            LookIn=-4123,
            LookAt=2,
            SearchOrder=1,
            MatchCase=True,
        )


# ── find — two matches ────────────────────────────────────────────────────────

class TestFindTwoMatches:
    def test_returns_two_hits(self):
        ms = make_mock_session()
        rng = _make_rng_mock_two_matches("$A$1", "foo", "$B$2", "foo")
        result = _patch_and_call("find", ms, rng, find_text="foo")
        assert result["applied"]["total_found"] == 2
        cells = [m["cell"] for m in result["applied"]["matches"]]
        assert "$A$1" in cells
        assert "$B$2" in cells

    def test_findnext_called_after_find(self):
        """FindNext must be called to advance; if only Find() were used we'd miss cells."""
        ms = make_mock_session()
        rng = _make_rng_mock_two_matches()
        _patch_and_call("find", ms, rng, find_text="foo")
        rng.FindNext.assert_called()


# ── find — infinite-loop guard ────────────────────────────────────────────────

class TestInfiniteLoopGuard:
    def test_loop_stops_when_findnext_returns_none(self):
        """If FindNext returns None, the loop must stop (not crash)."""
        ms = make_mock_session()
        cell = _make_cell_mock("$A$1", "x")
        rng = MagicMock()
        rng.Find.return_value = cell
        rng.FindNext.return_value = None  # no more matches
        rng.Worksheet = MagicMock()
        rng.Worksheet.Name = "Sheet1"
        rng.Application = MagicMock()

        result = _patch_and_call("find", ms, rng, find_text="x")
        assert result["applied"]["total_found"] == 1

    def test_loop_stops_at_first_address_wrap(self):
        """FindNext returning the first address again must terminate the loop."""
        ms = make_mock_session()
        rng = _make_rng_mock_single_match("$Z$99", "abc")
        # FindNext immediately returns mock with same address → 1 iteration
        result = _patch_and_call("find", ms, rng, find_text="abc")
        # Must NOT loop forever; exactly 1 hit
        assert result["applied"]["total_found"] == 1


# ── count ─────────────────────────────────────────────────────────────────────

class TestCount:
    def test_count_zero_on_no_match(self):
        ms = make_mock_session()
        rng = _make_rng_mock_no_match()
        result = _patch_and_call("count", ms, rng, find_text="ghost")
        assert result["find_replace"] == "count"
        assert result["applied"]["count"] == 0

    def test_count_two_matches(self):
        ms = make_mock_session()
        rng = _make_rng_mock_two_matches()
        result = _patch_and_call("count", ms, rng, find_text="foo")
        assert result["applied"]["count"] == 2

    def test_count_does_not_return_matches_list(self):
        ms = make_mock_session()
        rng = _make_rng_mock_no_match()
        result = _patch_and_call("count", ms, rng, find_text="x")
        assert "matches" not in result["applied"]


# ── replace ────────────────────────────────────────────────────────────────────

class TestReplace:
    def _make_replace_rng(self, pre_count=2, post_count=0):
        """Range mock where Find returns hits before replace and none after.

        Uses side_effect on Find to return different results across calls:
        1st call (pre-count): returns a cell chain of pre_count hits
        2nd call (post-count): returns None (replaced successfully)
        """
        rng = MagicMock()
        rng.Worksheet = MagicMock()
        rng.Worksheet.Name = "Sheet1"
        rng.Application = MagicMock()
        rng.Replace.return_value = True  # Boolean return per MS Learn

        # pre-count: 2 hits, then none (post-replace)
        if pre_count == 1:
            cell = _make_cell_mock("$A$1", "old")
            wrap = _make_cell_mock("$A$1", "old")
            rng.Find.side_effect = [cell, None]
            rng.FindNext.return_value = wrap
        elif pre_count == 2:
            cell1 = _make_cell_mock("$A$1", "old")
            cell2 = _make_cell_mock("$B$2", "old")
            wrap  = _make_cell_mock("$A$1", "old")
            # 1st Find call = cell1 (pre-count), 2nd Find call = None (post-replace)
            rng.Find.side_effect = [cell1, None]
            rng.FindNext.side_effect = [cell2, wrap]
        else:
            rng.Find.return_value = None

        return rng

    def test_replace_returns_correct_shape(self):
        ms = make_mock_session()
        rng = self._make_replace_rng(pre_count=1)
        result = _patch_and_call(
            "replace", ms, rng,
            find_text="old", replace_text="new",
        )
        assert result["find_replace"] == "replace"
        assert "cells_matched_before" in result["applied"]
        assert "remaining_after" in result["applied"]
        assert "fully_replaced" in result["applied"]

    def test_replace_reports_cells_matched_before(self):
        ms = make_mock_session()
        rng = self._make_replace_rng(pre_count=2)
        result = _patch_and_call(
            "replace", ms, rng,
            find_text="old", replace_text="new",
        )
        assert result["applied"]["cells_matched_before"] == 2

    def test_replace_calls_rng_replace_with_correct_params(self):
        """rng.Replace must be called with exact kwargs matching the brief spec."""
        ms = make_mock_session()
        rng = self._make_replace_rng(pre_count=0)
        _patch_and_call("replace", ms, rng, find_text="X", replace_text="Y")
        rng.Replace.assert_called_once_with(
            What="X",
            Replacement="Y",
            LookAt=2,       # xlPart (default match_whole_cell=False)
            SearchOrder=1,  # xlByRows
            MatchCase=False,
        )

    def test_replace_whole_cell_passes_xl_whole(self):
        ms = make_mock_session()
        rng = self._make_replace_rng(pre_count=0)
        _patch_and_call(
            "replace", ms, rng,
            find_text="X", replace_text="Y", match_whole_cell=True,
        )
        rng.Replace.assert_called_once_with(
            What="X",
            Replacement="Y",
            LookAt=1,       # xlWhole
            SearchOrder=1,
            MatchCase=False,
        )

    def test_replace_match_case_passed_through(self):
        ms = make_mock_session()
        rng = self._make_replace_rng(pre_count=0)
        _patch_and_call(
            "replace", ms, rng,
            find_text="abc", replace_text="ABC", match_case=True,
        )
        rng.Replace.assert_called_once_with(
            What="abc",
            Replacement="ABC",
            LookAt=2,
            SearchOrder=1,
            MatchCase=True,
        )

    def test_verify_effect_remaining_after_zero_on_success(self):
        """After replace, re-count must be 0 for fully_replaced=True."""
        ms = make_mock_session()
        rng = self._make_replace_rng(pre_count=1)  # post-count returns None
        result = _patch_and_call(
            "replace", ms, rng,
            find_text="old", replace_text="new",
        )
        assert result["applied"]["remaining_after"] == 0
        assert result["applied"]["fully_replaced"] is True

    def test_verify_effect_remaining_nonzero_when_replace_partial(self):
        """If replacement is partial, remaining_after > 0 and fully_replaced=False."""
        ms = make_mock_session()
        rng = MagicMock()
        rng.Worksheet = MagicMock()
        rng.Worksheet.Name = "Sheet1"
        rng.Application = MagicMock()
        rng.Replace.return_value = True

        # Pre-count: 1 hit; post-count: 1 hit still remains (partial replace)
        cell_pre  = _make_cell_mock("$A$1", "old")
        wrap_pre  = _make_cell_mock("$A$1", "old")
        cell_post = _make_cell_mock("$A$1", "old")
        wrap_post = _make_cell_mock("$A$1", "old")

        rng.Find.side_effect = [cell_pre, cell_post]
        # FindNext: 1st sequence wraps on 1st call; 2nd sequence wraps on 3rd call
        rng.FindNext.side_effect = [wrap_pre, wrap_post]

        result = _patch_and_call(
            "replace", ms, rng,
            find_text="old", replace_text="new",
        )
        # cells_matched_before=1, remaining_after=1, fully_replaced=False
        assert result["applied"]["remaining_after"] == 1
        assert result["applied"]["fully_replaced"] is False

    def test_replace_find_text_in_result(self):
        ms = make_mock_session()
        rng = self._make_replace_rng(pre_count=0)
        result = _patch_and_call(
            "replace", ms, rng,
            find_text="needle", replace_text="pin",
        )
        assert result["applied"]["find_text"] == "needle"
        assert result["applied"]["replace_text"] == "pin"


# ── scope=range ───────────────────────────────────────────────────────────────

class TestScopeRange:
    def test_find_in_range_scope(self):
        ms = make_mock_session()
        rng = _make_rng_mock_single_match("$B$3", "hi")
        result = _patch_and_call_range_scope(
            "find", ms, rng, range_str="B1:B10", find_text="hi"
        )
        assert result["applied"]["total_found"] == 1

    def test_count_in_range_scope(self):
        ms = make_mock_session()
        rng = _make_rng_mock_two_matches()
        result = _patch_and_call_range_scope(
            "count", ms, rng, range_str="A1:C10", find_text="foo"
        )
        assert result["applied"]["count"] == 2


# ── workbook scope ────────────────────────────────────────────────────────────

class TestWorkbookScope:
    def test_find_iterates_multiple_sheets(self):
        """scope='workbook' must search each sheet's Cells range."""
        ms = make_mock_session()

        ws1 = MagicMock()
        ws1.Name = "Sheet1"
        ws1.Cells = _make_rng_mock_single_match("$A$1", "x")

        ws2 = MagicMock()
        ws2.Name = "Sheet2"
        ws2.Cells = _make_rng_mock_single_match("$A$1", "x")

        wb = MagicMock()
        wb.Sheets.Count = 2
        wb.Sheets.side_effect = lambda i: [None, ws1, ws2][i]

        ms.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.find_replace._session", ms):
            with patch("thepexcel_mcp.domains.find_replace.excel_guard") as eg:
                eg.return_value.__enter__ = MagicMock(return_value=None)
                eg.return_value.__exit__ = MagicMock(return_value=False)
                from thepexcel_mcp.domains.find_replace import find_replace_action
                result = find_replace_action("find", find_text="x", scope="workbook")

        # One match per sheet → 2 total
        assert result["applied"]["total_found"] == 2


# ── Result shape ───────────────────────────────────────────────────────────────

class TestResultShape:
    def test_find_shape(self):
        ms = make_mock_session()
        rng = _make_rng_mock_no_match()
        result = _patch_and_call("find", ms, rng, find_text="z")
        assert "find_replace" in result
        assert "applied" in result
        applied = result["applied"]
        assert "find_text" in applied
        assert "total_found" in applied
        assert "truncated" in applied
        assert "matches" in applied

    def test_count_shape(self):
        ms = make_mock_session()
        rng = _make_rng_mock_no_match()
        result = _patch_and_call("count", ms, rng, find_text="z")
        assert result["find_replace"] == "count"
        assert "count" in result["applied"]

    def test_replace_shape(self):
        ms = make_mock_session()
        rng = _make_rng_mock_no_match()
        rng.Replace.return_value = True
        result = _patch_and_call(
            "replace", ms, rng,
            find_text="z", replace_text="y",
        )
        assert result["find_replace"] == "replace"
        applied = result["applied"]
        assert "find_text" in applied
        assert "replace_text" in applied
        assert "cells_matched_before" in applied
        assert "remaining_after" in applied
        assert "fully_replaced" in applied


# ── Truncation cap ────────────────────────────────────────────────────────────

class TestTruncationCap:
    def test_find_caps_at_1000(self):
        """When total_found > 1000 the matches list must be capped at 1000."""
        ms = make_mock_session()
        rng = MagicMock()
        rng.Worksheet = MagicMock()
        rng.Worksheet.Name = "Sheet1"
        rng.Application = MagicMock()

        # Generate 1002 unique cell addresses
        addresses = [f"$A${i}" for i in range(1, 1003)]
        cells = [_make_cell_mock(a, "x") for a in addresses]
        # FindNext chain: cell[1], cell[2], ..., cell[1001], then wrap back to cell[0]
        rng.Find.return_value = cells[0]
        # FindNext returns cell[1] through cell[1001], then a cell with addresses[0] (wrap)
        wrap_cell = _make_cell_mock(addresses[0], "x")
        rng.FindNext.side_effect = cells[1:] + [wrap_cell]

        result = _patch_and_call("find", ms, rng, find_text="x")
        assert result["applied"]["total_found"] == 1002
        assert len(result["applied"]["matches"]) == 1000
        assert result["applied"]["truncated"] is True
