"""Unit tests for excel_diff (diff.py).

No Excel required — all COM calls are intercepted via make_mock_session.

VERIFY-EFFECT discipline: tests assert the specific COM property/method
was called, and that the returned diff correctly reflects the actual
cell-level comparison (not just "no exception raised").

Helpers used:
- _to_2d()      — normalises scalar/flat-tuple/tuple-of-tuples
- _col_to_letter() — column number → A1-style letter
- _offset_to_a1()  — base row/col + offsets → A1 address
- _cell_differs()  — compare mode logic
- _build_diffs()   — core diff engine
"""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock, patch, PropertyMock

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


# ── Helper: invoke diff_action through a patched session ──────────────────────

def _call_diff(action, mock_session, **kwargs):
    """Patch _session and invoke diff_action."""
    with patch("thepexcel_mcp.domains.diff._session", mock_session):
        from thepexcel_mcp.domains.diff import diff_action
        return diff_action(action, **kwargs)


def _make_range_mock(
    value=None,
    formula=None,
    rows_count: int = 1,
    cols_count: int = 1,
    row: int = 1,
    col: int = 1,
    sheet_name: str = "Sheet1",
):
    """Build a minimal COM Range mock for diff tests."""
    rng = MagicMock()
    rng.Value = value
    rng.Formula = formula
    rng.Row = row
    rng.Column = col
    rng.Rows.Count = rows_count
    rng.Columns.Count = cols_count
    rng.Worksheet.Name = sheet_name
    return rng


def _make_sheet_mock_with_used_range(
    sheet_name: str = "Sheet1",
    ur_row: int = 1,
    ur_col: int = 1,
    ur_rows_count: int = 3,
    ur_cols_count: int = 3,
    block_value=None,
    block_formula=None,
):
    """Build a Worksheet mock wired for _diff_sheets (UsedRange + Cells + Range)."""
    ws = MagicMock()
    ws.Name = sheet_name

    ur = MagicMock()
    ur.Row = ur_row
    ur.Column = ur_col
    ur.Rows.Count = ur_rows_count
    ur.Columns.Count = ur_cols_count
    ws.UsedRange = ur

    block_rng = MagicMock()
    block_rng.Value = block_value
    block_rng.Formula = block_formula
    # ws.Range(ws.Cells(...), ws.Cells(...)) → block_rng
    ws.Range.return_value = block_rng

    return ws, block_rng


# ── Action / param validation ─────────────────────────────────────────────────

class TestDiffActionValidation:
    def test_unknown_action_raises(self):
        ms = make_mock_session()
        with pytest.raises(ToolError, match="Unknown action"):
            _call_diff("bogus", ms, left_range="A1", right_range="A1")

    def test_invalid_compare_raises(self):
        ms = make_mock_session()
        with pytest.raises(ToolError, match="Unknown compare"):
            _call_diff(
                "ranges", ms,
                left_range="A1", right_range="A1",
                compare="bad",
            )

    def test_max_diffs_zero_raises(self):
        ms = make_mock_session()
        with pytest.raises(ToolError, match="max_diffs"):
            _call_diff(
                "ranges", ms,
                left_range="A1", right_range="A1",
                max_diffs=0,
            )

    def test_max_diffs_negative_raises(self):
        ms = make_mock_session()
        with pytest.raises(ToolError, match="max_diffs"):
            _call_diff(
                "ranges", ms,
                left_range="A1", right_range="A1",
                max_diffs=-5,
            )

    def test_ranges_missing_left_range_raises(self):
        ms = make_mock_session()
        with pytest.raises(ToolError, match="left_range"):
            _call_diff("ranges", ms, right_range="B1:B3")

    def test_ranges_missing_right_range_raises(self):
        ms = make_mock_session()
        with pytest.raises(ToolError, match="right_range"):
            # We need the session to not error before _diff_ranges
            l_rng = _make_range_mock(value="X", formula="X", rows_count=1, cols_count=1)
            ms.get_sheet.return_value = MagicMock()
            ms.get_sheet.return_value.Range.return_value = l_rng
            _call_diff("ranges", ms, left_range="A1")

    def test_sheets_missing_left_sheet_raises(self):
        ms = make_mock_session()
        with pytest.raises(ToolError, match="left_sheet"):
            _call_diff("sheets", ms, right_sheet="Sheet2")

    def test_sheets_missing_right_sheet_raises(self):
        ms = make_mock_session()
        with pytest.raises(ToolError, match="right_sheet"):
            _call_diff("sheets", ms, left_sheet="Sheet1")


# ── _to_2d helper ─────────────────────────────────────────────────────────────

class TestTo2D:
    def test_scalar_str(self):
        from thepexcel_mcp.domains.diff import _to_2d
        assert _to_2d("hello") == [["hello"]]

    def test_scalar_int(self):
        from thepexcel_mcp.domains.diff import _to_2d
        assert _to_2d(42) == [[42]]

    def test_scalar_none(self):
        from thepexcel_mcp.domains.diff import _to_2d
        assert _to_2d(None) == [[None]]

    def test_scalar_zero(self):
        """Zero is a valid value — must not be treated as None/empty."""
        from thepexcel_mcp.domains.diff import _to_2d
        assert _to_2d(0) == [[0]]

    def test_flat_tuple_single_row(self):
        from thepexcel_mcp.domains.diff import _to_2d
        assert _to_2d((1, 2, 3)) == [[1, 2, 3]]

    def test_tuple_of_tuples(self):
        from thepexcel_mcp.domains.diff import _to_2d
        raw = ((1, 2), (3, 4))
        assert _to_2d(raw) == [[1, 2], [3, 4]]

    def test_empty_inner_tuples_preserved(self):
        from thepexcel_mcp.domains.diff import _to_2d
        raw = (("a", "b"), ("c", "d"))
        assert _to_2d(raw) == [["a", "b"], ["c", "d"]]


# ── _col_to_letter ────────────────────────────────────────────────────────────

class TestColToLetter:
    def test_col_1_is_A(self):
        from thepexcel_mcp.domains.diff import _col_to_letter
        assert _col_to_letter(1) == "A"

    def test_col_26_is_Z(self):
        from thepexcel_mcp.domains.diff import _col_to_letter
        assert _col_to_letter(26) == "Z"

    def test_col_27_is_AA(self):
        from thepexcel_mcp.domains.diff import _col_to_letter
        assert _col_to_letter(27) == "AA"

    def test_col_52_is_AZ(self):
        from thepexcel_mcp.domains.diff import _col_to_letter
        assert _col_to_letter(52) == "AZ"

    def test_col_703_is_AAA(self):
        from thepexcel_mcp.domains.diff import _col_to_letter
        assert _col_to_letter(703) == "AAA"


# ── _offset_to_a1 ─────────────────────────────────────────────────────────────

class TestOffsetToA1:
    def test_base_a1_offset_00(self):
        from thepexcel_mcp.domains.diff import _offset_to_a1
        assert _offset_to_a1(1, 1, 0, 0) == "A1"

    def test_base_b2_offset_00(self):
        from thepexcel_mcp.domains.diff import _offset_to_a1
        assert _offset_to_a1(2, 2, 0, 0) == "B2"

    def test_base_a1_offset_12(self):
        from thepexcel_mcp.domains.diff import _offset_to_a1
        assert _offset_to_a1(1, 1, 1, 2) == "C2"

    def test_base_c3_offset_23(self):
        from thepexcel_mcp.domains.diff import _offset_to_a1
        # row = 3+2=5, col = 3+3=6 → F5
        assert _offset_to_a1(3, 3, 2, 3) == "F5"


# ── _cell_differs ─────────────────────────────────────────────────────────────

class TestCellDiffers:
    def test_value_same(self):
        from thepexcel_mcp.domains.diff import _cell_differs
        assert not _cell_differs(1, 1, "=A1", "=B1", "value")

    def test_value_different(self):
        from thepexcel_mcp.domains.diff import _cell_differs
        assert _cell_differs(1, 2, "=A1", "=A1", "value")

    def test_formula_same_value_different(self):
        """Formula compare: same formula → not different, even if values differ."""
        from thepexcel_mcp.domains.diff import _cell_differs
        assert not _cell_differs(1, 999, "=A1", "=A1", "formula")

    def test_formula_different_value_same(self):
        from thepexcel_mcp.domains.diff import _cell_differs
        assert _cell_differs(1, 1, "=A1", "=B1", "formula")

    def test_both_value_different_formula_same(self):
        from thepexcel_mcp.domains.diff import _cell_differs
        assert _cell_differs(1, 2, "=A1", "=A1", "both")

    def test_both_value_same_formula_different(self):
        from thepexcel_mcp.domains.diff import _cell_differs
        assert _cell_differs(1, 1, "=A1", "=B1", "both")

    def test_both_neither_different(self):
        from thepexcel_mcp.domains.diff import _cell_differs
        assert not _cell_differs(1, 1, "=A1", "=A1", "both")

    def test_none_vs_empty_string_are_different(self):
        """None (blank cell) and '' (empty string) must be treated as different."""
        from thepexcel_mcp.domains.diff import _cell_differs
        assert _cell_differs(None, "", None, None, "value")

    def test_none_vs_none_are_same(self):
        from thepexcel_mcp.domains.diff import _cell_differs
        assert not _cell_differs(None, None, None, None, "value")


# ── _build_diffs ──────────────────────────────────────────────────────────────

class TestBuildDiffs:
    def test_identical_grids_no_diffs(self):
        from thepexcel_mcp.domains.diff import _build_diffs
        grid = [[1, 2], [3, 4]]
        diffs, total, truncated = _build_diffs(
            grid, grid, None, None, "value", 500, 1, 1
        )
        assert total == 0
        assert not truncated
        assert diffs == []

    def test_one_diff_correct_cell_address(self):
        from thepexcel_mcp.domains.diff import _build_diffs
        left  = [[1, 2], [3, 4]]
        right = [[1, 2], [3, 99]]  # differs at (1, 1) → base(1,1)+1,1 = B2
        diffs, total, truncated = _build_diffs(
            left, right, None, None, "value", 500, 1, 1
        )
        assert total == 1
        assert diffs[0]["cell"] == "B2"
        assert diffs[0]["left_value"] == 4
        assert diffs[0]["right_value"] == 99
        assert diffs[0]["row"] == 1
        assert diffs[0]["col"] == 1

    def test_truncation_sets_flag_and_preserves_true_total(self):
        from thepexcel_mcp.domains.diff import _build_diffs
        # 3×3 grid, all different
        left  = [[i * 3 + j for j in range(3)] for i in range(3)]
        right = [[99 for _ in range(3)] for _ in range(3)]
        diffs, total, truncated = _build_diffs(
            left, right, None, None, "value", max_diffs=2, base_row=1, base_col=1
        )
        assert total == 9          # true total
        assert truncated is True
        assert len(diffs) == 2     # capped

    def test_overlap_used_for_different_sized_grids(self):
        from thepexcel_mcp.domains.diff import _build_diffs
        left  = [[1, 2, 3], [4, 5, 6]]          # 2×3
        right = [[1, 9], [4, 5], [7, 8]]         # 3×2 — overlap = 2×2
        diffs, total, _ = _build_diffs(
            left, right, None, None, "value", 500, 1, 1
        )
        # Only overlap (2×2) compared: (0,1)=2 vs 9 → diff; rest match
        assert total == 1
        assert diffs[0]["col"] == 1

    def test_formula_mode_includes_formula_keys(self):
        from thepexcel_mcp.domains.diff import _build_diffs
        left_v  = [[1]]
        right_v = [[1]]
        left_f  = [["=A1"]]
        right_f = [["=B1"]]  # formula differs
        diffs, total, _ = _build_diffs(
            left_v, right_v, left_f, right_f, "formula", 500, 1, 1
        )
        assert total == 1
        assert "left_formula" in diffs[0]
        assert "right_formula" in diffs[0]
        assert diffs[0]["left_formula"] == "=A1"
        assert diffs[0]["right_formula"] == "=B1"

    def test_value_mode_no_formula_keys_in_entry(self):
        from thepexcel_mcp.domains.diff import _build_diffs
        left  = [[1]]
        right = [[2]]
        diffs, _, _ = _build_diffs(left, right, None, None, "value", 500, 1, 1)
        assert "left_formula" not in diffs[0]
        assert "right_formula" not in diffs[0]

    def test_base_row_col_applied_to_cell_address(self):
        from thepexcel_mcp.domains.diff import _build_diffs
        left  = [[1, 2], [3, 4]]
        right = [[1, 2], [3, 99]]  # diff at (1,1)
        # Left range starts at C5 (row=5, col=3)
        diffs, _, _ = _build_diffs(left, right, None, None, "value", 500, 5, 3)
        # row_i=1, col_j=1 → row=5+1=6, col=3+1=4 → D6
        assert diffs[0]["cell"] == "D6"


# ── ranges action (integration-level mock) ────────────────────────────────────

class TestRangesAction:
    def _setup_session_ranges(self, ms, l_rng, r_rng, left_str="A1:B2", right_str="C1:D2"):
        """Wire ms.get_sheet to return sheets whose Range() returns our mocks."""
        l_ws = MagicMock()
        l_ws.Range.return_value = l_rng
        r_ws = MagicMock()
        r_ws.Range.return_value = r_rng

        def get_sheet(name, workbook):
            # The resolver just needs to call ws.Range(cell_part) or ws.Range(range_str)
            return l_ws if (name is None or name == "Sheet1") else r_ws

        ms.get_sheet.side_effect = get_sheet
        return l_ws, r_ws

    def test_identical_ranges_no_diffs(self):
        ms = make_mock_session()
        vals = ((1, 2), (3, 4))
        l_rng = _make_range_mock(value=vals, formula=vals, rows_count=2, cols_count=2)
        r_rng = _make_range_mock(value=vals, formula=vals, rows_count=2, cols_count=2)

        l_ws = MagicMock(); l_ws.Range.return_value = l_rng
        r_ws = MagicMock(); r_ws.Range.return_value = r_rng
        call_count = [0]

        def get_sheet(name, workbook):
            call_count[0] += 1
            return l_ws if call_count[0] <= 1 else r_ws

        ms.get_sheet.side_effect = get_sheet

        result = _call_diff(
            "ranges", ms,
            left_range="A1:B2", right_range="A1:B2",
        )
        assert result["diff"] == "ranges"
        assert result["applied"]["total_diffs"] == 0
        assert result["applied"]["truncated"] is False
        assert result["applied"]["dimensions_match"] is True

    def test_one_diff_detected(self):
        ms = make_mock_session()
        # Only cell (0,1) differs: 2 vs 99
        l_vals = ((1, 2), (3, 4))
        r_vals = ((1, 99), (3, 4))
        l_rng = _make_range_mock(value=l_vals, formula=l_vals, rows_count=2, cols_count=2, row=1, col=1)
        r_rng = _make_range_mock(value=r_vals, formula=r_vals, rows_count=2, cols_count=2, row=1, col=1)

        call_count = [0]
        l_ws = MagicMock(); l_ws.Range.return_value = l_rng
        r_ws = MagicMock(); r_ws.Range.return_value = r_rng

        def get_sheet(name, workbook):
            call_count[0] += 1
            return l_ws if call_count[0] <= 1 else r_ws

        ms.get_sheet.side_effect = get_sheet

        result = _call_diff(
            "ranges", ms,
            left_range="A1:B2", right_range="A1:B2",
        )
        assert result["applied"]["total_diffs"] == 1
        d = result["applied"]["diffs"][0]
        assert d["cell"] == "B1"     # row=0+1=1, col=0+1+1=2 → B1
        assert d["left_value"] == 2
        assert d["right_value"] == 99

    def test_shape_mismatch_produces_shape_note(self):
        ms = make_mock_session()
        l_rng = _make_range_mock(value=((1, 2, 3),), formula=((1, 2, 3),), rows_count=1, cols_count=3)
        r_rng = _make_range_mock(value=((1, 2),), formula=((1, 2),), rows_count=1, cols_count=2)

        call_count = [0]
        l_ws = MagicMock(); l_ws.Range.return_value = l_rng
        r_ws = MagicMock(); r_ws.Range.return_value = r_rng

        def get_sheet(name, workbook):
            call_count[0] += 1
            return l_ws if call_count[0] <= 1 else r_ws

        ms.get_sheet.side_effect = get_sheet

        result = _call_diff(
            "ranges", ms,
            left_range="A1:C1", right_range="A1:B1",
        )
        assert result["applied"]["dimensions_match"] is False
        assert "shape_note" in result["applied"]

    def test_single_cell_range_scalar_normalised(self):
        """Single cell .Value returns scalar — must be normalised to [[scalar]]."""
        ms = make_mock_session()
        # Left = "hello", Right = "world" → 1 diff
        l_rng = _make_range_mock(value="hello", formula="hello", rows_count=1, cols_count=1)
        r_rng = _make_range_mock(value="world", formula="world", rows_count=1, cols_count=1)

        call_count = [0]
        l_ws = MagicMock(); l_ws.Range.return_value = l_rng
        r_ws = MagicMock(); r_ws.Range.return_value = r_rng

        def get_sheet(name, workbook):
            call_count[0] += 1
            return l_ws if call_count[0] <= 1 else r_ws

        ms.get_sheet.side_effect = get_sheet

        result = _call_diff("ranges", ms, left_range="A1", right_range="B1")
        assert result["applied"]["total_diffs"] == 1
        assert result["applied"]["diffs"][0]["left_value"] == "hello"
        assert result["applied"]["diffs"][0]["right_value"] == "world"

    def test_formula_compare_reads_formula_property(self):
        """compare='formula' must read rng.Formula, not rng.Value.

        Verify-effect: rng.Value must NOT be accessed — we set it to a
        deliberately wrong tuple that would produce a false no-diff if used.
        """
        ms = make_mock_session()
        # Same .Value on both sides (would produce 0 diffs if accidentally used)
        # Different .Formula → should produce 1 diff
        l_rng = _make_range_mock(
            value=((10,),), formula=(("=A1",),), rows_count=1, cols_count=1
        )
        r_rng = _make_range_mock(
            value=((10,),), formula=(("=B1",),), rows_count=1, cols_count=1
        )

        call_count = [0]
        l_ws = MagicMock(); l_ws.Range.return_value = l_rng
        r_ws = MagicMock(); r_ws.Range.return_value = r_rng

        def get_sheet(name, workbook):
            call_count[0] += 1
            return l_ws if call_count[0] <= 1 else r_ws

        ms.get_sheet.side_effect = get_sheet

        result = _call_diff(
            "ranges", ms, left_range="A1", right_range="B1", compare="formula"
        )
        # Formulas differ → 1 diff (would be 0 if .Value was used for diff
        # detection and formula was ignored)
        assert result["applied"]["total_diffs"] == 1
        d = result["applied"]["diffs"][0]
        assert "left_formula" in d
        assert d["left_formula"] == "=A1"
        assert d["right_formula"] == "=B1"

    def test_formula_compare_ignores_differing_values(self):
        """compare='formula' must NOT report a diff when formulas are identical,
        even when .Value differs between the two ranges.

        Verify-effect: if the code accidentally used .Value for diff detection
        in formula mode, this test would report 1 diff instead of 0 — proving
        formula-mode is driven by .Formula, not .Value.
        """
        ms = make_mock_session()
        # Formulas identical, but Values deliberately different
        l_rng = _make_range_mock(
            value=((10,),),   # LEFT value
            formula=(("=A1",),),
            rows_count=1, cols_count=1,
        )
        r_rng = _make_range_mock(
            value=((99,),),   # RIGHT value — differs from left
            formula=(("=A1",),),  # same formula as left
            rows_count=1, cols_count=1,
        )

        call_count = [0]
        l_ws = MagicMock(); l_ws.Range.return_value = l_rng
        r_ws = MagicMock(); r_ws.Range.return_value = r_rng

        def get_sheet(name, workbook):
            call_count[0] += 1
            return l_ws if call_count[0] <= 1 else r_ws

        ms.get_sheet.side_effect = get_sheet

        result = _call_diff(
            "ranges", ms, left_range="A1", right_range="B1", compare="formula"
        )
        # Formulas same → 0 diffs.
        # If .Value were accidentally used for diff detection, this would be 1.
        assert result["applied"]["total_diffs"] == 0
        assert result["applied"]["diffs"] == []

    def test_truncation_and_total_diffs_preserved(self):
        """total_diffs must be the true count even when diffs are truncated."""
        ms = make_mock_session()
        # 3×3 = 9 cells, all different
        l_vals = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
        r_vals = ((0, 0, 0), (0, 0, 0), (0, 0, 0))
        l_rng = _make_range_mock(value=l_vals, formula=l_vals, rows_count=3, cols_count=3)
        r_rng = _make_range_mock(value=r_vals, formula=r_vals, rows_count=3, cols_count=3)

        call_count = [0]
        l_ws = MagicMock(); l_ws.Range.return_value = l_rng
        r_ws = MagicMock(); r_ws.Range.return_value = r_rng

        def get_sheet(name, workbook):
            call_count[0] += 1
            return l_ws if call_count[0] <= 1 else r_ws

        ms.get_sheet.side_effect = get_sheet

        result = _call_diff(
            "ranges", ms,
            left_range="A1:C3", right_range="A1:C3",
            max_diffs=3,
        )
        assert result["applied"]["total_diffs"] == 9
        assert result["applied"]["truncated"] is True
        assert len(result["applied"]["diffs"]) == 3


# ── sheets action ─────────────────────────────────────────────────────────────

class TestSheetsAction:
    def _make_sheet_session(self, l_ws, r_ws, ms):
        call_count = [0]

        def get_sheet(name, workbook):
            call_count[0] += 1
            return l_ws if call_count[0] <= 1 else r_ws

        ms.get_sheet.side_effect = get_sheet

    def test_identical_sheets_no_diffs(self):
        ms = make_mock_session()
        vals = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
        l_ws, l_block = _make_sheet_mock_with_used_range(
            "Sheet1", block_value=vals, block_formula=vals,
        )
        r_ws, r_block = _make_sheet_mock_with_used_range(
            "Sheet2", block_value=vals, block_formula=vals,
        )
        self._make_sheet_session(l_ws, r_ws, ms)

        result = _call_diff(
            "sheets", ms, left_sheet="Sheet1", right_sheet="Sheet2"
        )
        assert result["diff"] == "sheets"
        assert result["applied"]["total_diffs"] == 0
        assert result["applied"]["left_sheet"] == "Sheet1"
        assert result["applied"]["right_sheet"] == "Sheet2"

    def test_diff_detected_in_sheets(self):
        ms = make_mock_session()
        l_vals = ((1, 2), (3, 4))
        r_vals = ((1, 2), (3, 99))  # (1,1) differs
        l_ws, _ = _make_sheet_mock_with_used_range(
            "Sheet1", block_value=l_vals, block_formula=l_vals,
            ur_rows_count=2, ur_cols_count=2,
        )
        r_ws, _ = _make_sheet_mock_with_used_range(
            "Sheet2", block_value=r_vals, block_formula=r_vals,
            ur_rows_count=2, ur_cols_count=2,
        )
        self._make_sheet_session(l_ws, r_ws, ms)

        result = _call_diff(
            "sheets", ms, left_sheet="Sheet1", right_sheet="Sheet2"
        )
        assert result["applied"]["total_diffs"] == 1
        d = result["applied"]["diffs"][0]
        assert d["cell"] == "B2"    # base row=1, col=1; row_i=1, col_j=1 → B2
        assert d["left_value"] == 4
        assert d["right_value"] == 99

    def test_sheets_result_has_bounding_box(self):
        ms = make_mock_session()
        vals = ((1, 2), (3, 4))
        l_ws, _ = _make_sheet_mock_with_used_range(
            "Sheet1", block_value=vals, block_formula=vals,
            ur_rows_count=2, ur_cols_count=2,
        )
        r_ws, _ = _make_sheet_mock_with_used_range(
            "Sheet2", block_value=vals, block_formula=vals,
            ur_rows_count=3, ur_cols_count=2,  # right sheet taller
        )
        self._make_sheet_session(l_ws, r_ws, ms)

        result = _call_diff(
            "sheets", ms, left_sheet="Sheet1", right_sheet="Sheet2"
        )
        applied = result["applied"]
        assert "bounding_box" in applied
        assert "left_used_range" in applied
        assert "right_used_range" in applied
        # right sheet has 3 rows → bounding box max_row = 3
        assert applied["bounding_box"][0] == 3

    def test_sheets_calls_ws_range_cells_for_block(self):
        """ws.Range(ws.Cells(1,1), ws.Cells(maxrow, maxcol)) must be called.

        Verify-effect: check Range is called with exactly 2 positional args
        (the two Cells() results), confirming the bounding-box block pattern
        rather than a single-address Range("A1:...") shortcut.
        """
        ms = make_mock_session()
        vals = ((1,),)
        l_ws, _ = _make_sheet_mock_with_used_range(
            "Sheet1", block_value=vals, block_formula=vals,
            ur_rows_count=1, ur_cols_count=1,
        )
        r_ws, _ = _make_sheet_mock_with_used_range(
            "Sheet2", block_value=vals, block_formula=vals,
            ur_rows_count=1, ur_cols_count=1,
        )
        self._make_sheet_session(l_ws, r_ws, ms)

        _call_diff("sheets", ms, left_sheet="Sheet1", right_sheet="Sheet2")

        # ws.Range must be called with two positional args (Cells, Cells) per sheet
        l_args, _ = l_ws.Range.call_args
        r_args, _ = r_ws.Range.call_args
        assert len(l_args) == 2, "left ws.Range must be called with 2 Cells args"
        assert len(r_args) == 2, "right ws.Range must be called with 2 Cells args"
        # The two args are the results of ws.Cells(1,1) and ws.Cells(max_row, max_col)
        # Verify ws.Cells was called at all (confirming Cells-based addressing)
        l_ws.Cells.assert_called()
        r_ws.Cells.assert_called()

    def test_sheets_truncation_total_preserved(self):
        ms = make_mock_session()
        # 3×3 all different
        l_vals = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
        r_vals = ((0, 0, 0), (0, 0, 0), (0, 0, 0))
        l_ws, _ = _make_sheet_mock_with_used_range(
            "Sheet1", block_value=l_vals, block_formula=l_vals,
            ur_rows_count=3, ur_cols_count=3,
        )
        r_ws, _ = _make_sheet_mock_with_used_range(
            "Sheet2", block_value=r_vals, block_formula=r_vals,
            ur_rows_count=3, ur_cols_count=3,
        )
        self._make_sheet_session(l_ws, r_ws, ms)

        result = _call_diff(
            "sheets", ms,
            left_sheet="Sheet1", right_sheet="Sheet2",
            max_diffs=4,
        )
        assert result["applied"]["total_diffs"] == 9
        assert result["applied"]["truncated"] is True
        assert len(result["applied"]["diffs"]) == 4


# ── Result shape contract ──────────────────────────────────────────────────────

class TestResultShape:
    """Every result must have 'diff' and 'applied' keys with required sub-keys."""

    def _simple_range_result(self, compare="value"):
        ms = make_mock_session()
        vals = ((1, 2), (3, 4))
        l_rng = _make_range_mock(value=vals, formula=vals, rows_count=2, cols_count=2)
        r_rng = _make_range_mock(value=vals, formula=vals, rows_count=2, cols_count=2)

        call_count = [0]
        l_ws = MagicMock(); l_ws.Range.return_value = l_rng
        r_ws = MagicMock(); r_ws.Range.return_value = r_rng

        def get_sheet(name, workbook):
            call_count[0] += 1
            return l_ws if call_count[0] <= 1 else r_ws

        ms.get_sheet.side_effect = get_sheet
        return _call_diff("ranges", ms, left_range="A1:B2", right_range="A1:B2", compare=compare)

    def test_ranges_result_keys_present(self):
        result = self._simple_range_result()
        assert "diff" in result
        assert "applied" in result
        assert result["diff"] == "ranges"

    def test_ranges_applied_has_required_keys(self):
        result = self._simple_range_result()
        ap = result["applied"]
        for key in (
            "dimensions_left", "dimensions_right", "dimensions_match",
            "compare", "total_diffs", "truncated", "max_diffs", "diffs",
        ):
            assert key in ap, f"Missing key: {key}"

    def test_ranges_dimensions_are_lists(self):
        result = self._simple_range_result()
        assert isinstance(result["applied"]["dimensions_left"], list)
        assert isinstance(result["applied"]["dimensions_right"], list)

    def test_compare_field_echoed_in_result(self):
        result = self._simple_range_result(compare="both")
        assert result["applied"]["compare"] == "both"
