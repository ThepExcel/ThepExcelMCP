"""Unit tests for ranges._read / _read_spill page-scoped reads + the
exponential-probe spill scan — no Excel required.

Covers the Phase 2 "page-scoped reads" optimization (read only the page's
sub-range via bounded Cells(...) calls, not the whole range's .Value) and the
Phase 1 spill-scan optimization (exponential probe + binary search instead of
a linear HasSpill scan), verifying both preserve the original response shape.
"""

from __future__ import annotations

from unittest.mock import MagicMock, patch

from conftest import make_mock_session


# ── _normalize_rows ─────────────────────────────────────────────────────────────

class TestNormalizeRows:
    def test_none_returns_empty(self):
        from thepexcel_mcp.domains.ranges import _normalize_rows
        assert _normalize_rows(None) == []

    def test_scalar_wraps_in_1x1(self):
        from thepexcel_mcp.domains.ranges import _normalize_rows
        assert _normalize_rows(42) == [[42]]

    def test_flat_tuple_is_one_row(self):
        from thepexcel_mcp.domains.ranges import _normalize_rows
        assert _normalize_rows((1, 2, 3)) == [[1, 2, 3]]

    def test_tuple_of_tuples_is_multi_row(self):
        from thepexcel_mcp.domains.ranges import _normalize_rows
        assert _normalize_rows(((1, 2), (3, 4))) == [[1, 2], [3, 4]]

    def test_long_string_is_truncated(self):
        from thepexcel_mcp.domains.ranges import _normalize_rows, _MAX_CELL_LEN
        long_str = "x" * (_MAX_CELL_LEN + 10)
        result = _normalize_rows(((long_str,),))
        assert result[0][0] == "x" * _MAX_CELL_LEN + "…"


# ── _read: page-scoped reads ────────────────────────────────────────────────────

def _make_read_rng_mock(total_rows: int, total_cols: int, row: int = 1, col: int = 1):
    """Mock COM Range wired for the new page-scoped _read() path.

    rng.Parent -> ws; ws.Cells(...) builds the block bounds; ws.Range(a, b)
    returns the page sub-range whose .Value is set per-test.
    """
    page_rng = MagicMock()
    ws = MagicMock()
    ws.Range.return_value = page_rng
    ws.Cells.return_value = MagicMock()

    anchor = MagicMock()
    anchor.HasSpill = False  # no spill by default

    rng = MagicMock()
    rng.Rows.Count = total_rows
    rng.Columns.Count = total_cols
    rng.Row = row
    rng.Column = col
    rng.Parent = ws
    rng.Cells.return_value = anchor

    return rng, ws, page_rng, anchor


class TestReadPageScoped:
    def _call_read(self, rng, offset=0, limit=100, range_str="A1:C10"):
        with patch("thepexcel_mcp.domains.ranges._resolve_range", return_value=rng):
            from thepexcel_mcp.domains.ranges import _read
            return _read(range_str, None, None, offset, limit)

    def test_first_page_builds_bounded_block(self):
        rng, ws, page_rng, anchor = _make_read_rng_mock(total_rows=250, total_cols=3, row=1, col=1)
        page_rng.Value = tuple(tuple(f"r{r}c{c}" for c in range(3)) for r in range(100))

        result = self._call_read(rng, offset=0, limit=100)

        ws.Cells.assert_any_call(1, 1)     # top-left of page
        ws.Cells.assert_any_call(100, 3)   # bottom-right of page (row 100, col 3)
        assert result["total_rows"] == 250
        assert result["has_more"] is True
        assert result["next_offset"] == 100
        assert len(result["values"]) == 100

    def test_offset_page_builds_shifted_block(self):
        rng, ws, page_rng, anchor = _make_read_rng_mock(total_rows=250, total_cols=2, row=5, col=2)
        page_rng.Value = tuple((1, 2) for _ in range(100))

        result = self._call_read(rng, offset=100, limit=100, range_str="B5:C254")

        # r1=5,c1=2; page start row=5+100=105; last_row=5+min(200,250)-1=204
        ws.Cells.assert_any_call(105, 2)
        ws.Cells.assert_any_call(204, 3)
        assert result["total_rows"] == 250
        assert result["next_offset"] == 200

    def test_last_partial_page(self):
        rng, ws, page_rng, anchor = _make_read_rng_mock(total_rows=250, total_cols=1, row=1, col=1)
        page_rng.Value = tuple((i,) for i in range(50))

        result = self._call_read(rng, offset=200, limit=100)

        # last_row = 1 + min(300,250) - 1 = 250
        ws.Cells.assert_any_call(201, 1)
        ws.Cells.assert_any_call(250, 1)
        assert result["has_more"] is False
        assert result["next_offset"] is None

    def test_offset_beyond_total_rows_skips_com_read(self):
        rng, ws, page_rng, anchor = _make_read_rng_mock(total_rows=10, total_cols=2, row=1, col=1)

        result = self._call_read(rng, offset=50, limit=100)

        assert result["values"] == []
        assert result["has_more"] is False
        assert result["next_offset"] is None
        ws.Range.assert_not_called()

    def test_single_empty_cell_legacy_quirk(self):
        """A truly-empty 1x1 cell reports total_rows=0 (not 1), matching the
        old whole-value-read behavior — and skips any page COM read."""
        rng, ws, page_rng, anchor = _make_read_rng_mock(total_rows=1, total_cols=1, row=1, col=1)
        rng.Value = None

        result = self._call_read(rng, offset=0, limit=100, range_str="A1")

        assert result == {"values": [], "total_rows": 0, "has_more": False, "next_offset": 0}
        ws.Range.assert_not_called()

    def test_single_nonempty_cell_reads_scalar(self):
        rng, ws, page_rng, anchor = _make_read_rng_mock(total_rows=1, total_cols=1, row=1, col=1)
        rng.Value = 42
        page_rng.Value = 42

        result = self._call_read(rng, offset=0, limit=100, range_str="A1")

        assert result["values"] == [[42]]
        assert result["total_rows"] == 1
        assert result["has_more"] is False

    def test_raw_mode_reads_value2_not_value(self):
        rng, ws, page_rng, anchor = _make_read_rng_mock(
            total_rows=2, total_cols=1, row=1, col=1
        )
        page_rng.Value = (("typed",), ("typed",))
        page_rng.Value2 = ((45293.5,), (45294.5,))

        with patch("thepexcel_mcp.domains.ranges._resolve_range", return_value=rng):
            from thepexcel_mcp.domains.ranges import _read
            result = _read("A1:A2", None, None, 0, 100, value_mode="raw")

        assert result["values"] == [[45293.5], [45294.5]]

    def test_spill_metadata_still_populated(self):
        rng, ws, page_rng, anchor = _make_read_rng_mock(total_rows=5, total_cols=2, row=1, col=1)
        page_rng.Value = tuple((1, 2) for _ in range(5))
        anchor.HasSpill = True
        anchor.Formula2 = "=UNIQUE(A1:A10)"

        with patch("thepexcel_mcp.domains.ranges._spill_range_address", return_value="$A$1:$B$5"):
            result = self._call_read(rng, offset=0, limit=100, range_str="A1:B5")

        assert result["has_spill"] is True
        assert result["spill_range"] == "$A$1:$B$5"


# ── _read_spill: page-scoped reads ──────────────────────────────────────────────

class TestReadSpillPageScoped:
    def _make_anchor_and_ws(self):
        anchor_cell = MagicMock()
        anchor_cell.HasSpill = True
        anchor_cell.Address = "$A$1"
        ws = MagicMock()
        anchor_cell.Parent = ws
        resolved = MagicMock()
        resolved.Cells.return_value = anchor_cell
        return anchor_cell, ws, resolved

    def test_page_scoped_spill_read(self):
        anchor_cell, ws, resolved = self._make_anchor_and_ws()

        spill_rng = MagicMock()
        spill_rng.Rows.Count = 3
        spill_rng.Columns.Count = 2
        spill_rng.Row = 1
        spill_rng.Column = 1

        page_rng = MagicMock()
        page_rng.Value = ((1, 2), (3, 4), (5, 6))

        def range_side_effect(*args):
            return spill_rng if len(args) == 1 else page_rng

        ws.Range.side_effect = range_side_effect

        with patch("thepexcel_mcp.domains.ranges._resolve_range", return_value=resolved):
            with patch("thepexcel_mcp.domains.ranges._spill_range_address", return_value="$A$1:$B$3"):
                from thepexcel_mcp.domains.ranges import _read_spill
                result = _read_spill("A1", None, None, offset=0, limit=100)

        assert result["anchor"] == "$A$1"
        assert result["spill_range"] == "$A$1:$B$3"
        assert result["total_rows"] == 3
        assert result["values"] == [[1, 2], [3, 4], [5, 6]]
        assert result["has_more"] is False

    def test_no_spill_raises(self):
        from fastmcp.exceptions import ToolError
        import pytest

        anchor_cell = MagicMock()
        anchor_cell.HasSpill = False
        resolved = MagicMock()
        resolved.Cells.return_value = anchor_cell

        with patch("thepexcel_mcp.domains.ranges._resolve_range", return_value=resolved):
            from thepexcel_mcp.domains.ranges import _read_spill
            with pytest.raises(ToolError, match="no spill range"):
                _read_spill("A1", None, None, offset=0, limit=100)

    def test_single_cell_empty_spill_reports_zero(self):
        anchor_cell, ws, resolved = self._make_anchor_and_ws()
        spill_rng = MagicMock()
        spill_rng.Rows.Count = 1
        spill_rng.Columns.Count = 1
        spill_rng.Value = None
        ws.Range.return_value = spill_rng

        with patch("thepexcel_mcp.domains.ranges._resolve_range", return_value=resolved):
            with patch("thepexcel_mcp.domains.ranges._spill_range_address", return_value="$A$1"):
                from thepexcel_mcp.domains.ranges import _read_spill
                result = _read_spill("A1", None, None, offset=0, limit=100)

        assert result["total_rows"] == 0
        assert result["values"] == []

    def test_raw_mode_spill_reads_value2(self):
        anchor_cell, ws, resolved = self._make_anchor_and_ws()
        spill_rng = MagicMock()
        spill_rng.Rows.Count = 2
        spill_rng.Columns.Count = 1
        spill_rng.Row = 1
        spill_rng.Column = 1
        page_rng = MagicMock()
        page_rng.Value = (("typed",), ("typed",))
        page_rng.Value2 = ((1.25,), (2.5,))
        ws.Range.side_effect = lambda *args: spill_rng if len(args) == 1 else page_rng

        with patch("thepexcel_mcp.domains.ranges._resolve_range", return_value=resolved):
            with patch(
                "thepexcel_mcp.domains.ranges._spill_range_address",
                return_value="$A$1:$A$2",
            ):
                from thepexcel_mcp.domains.ranges import _read_spill
                result = _read_spill(
                    "A1", None, None, offset=0, limit=100, value_mode="raw"
                )

        assert result["values"] == [[1.25], [2.5]]


def test_range_action_rejects_invalid_value_mode():
    import pytest
    from fastmcp.exceptions import ToolError
    from thepexcel_mcp.domains.ranges import range_action

    with pytest.raises(ToolError, match="typed.*raw"):
        range_action("read", range="A1", value_mode="currency")


# ── _scan_spill_extent (exponential probe + binary search) ─────────────────────

class TestScanSpillExtent:
    def test_no_extension_beyond_anchor(self):
        from thepexcel_mcp.domains.ranges import _scan_spill_extent
        assert _scan_spill_extent(lambda k: False) == 0

    def test_small_extent(self):
        from thepexcel_mcp.domains.ranges import _scan_spill_extent
        assert _scan_spill_extent(lambda k: k <= 5) == 5

    def test_extent_of_one(self):
        from thepexcel_mcp.domains.ranges import _scan_spill_extent
        assert _scan_spill_extent(lambda k: k <= 1) == 1

    def test_fills_entire_cap(self):
        from thepexcel_mcp.domains.ranges import _scan_spill_extent
        assert _scan_spill_extent(lambda k: True, cap=999) == 999

    def test_respects_small_custom_cap(self):
        from thepexcel_mcp.domains.ranges import _scan_spill_extent
        assert _scan_spill_extent(lambda k: True, cap=3) == 3

    def test_matches_linear_scan_for_many_boundaries(self):
        """Cross-check against a brute-force linear reference scan."""
        from thepexcel_mcp.domains.ranges import _scan_spill_extent

        for boundary in (0, 1, 2, 3, 4, 5, 7, 10, 50, 100, 500, 998, 999):
            def is_true(k, b=boundary):
                return k <= b

            expected = 0
            for k in range(1, 1000):
                if is_true(k):
                    expected = k
                else:
                    break
            assert _scan_spill_extent(is_true) == expected, f"boundary={boundary}"


# ── _spill_range_address (fallback scan integration) ────────────────────────────

class TestSpillRangeAddressFallback:
    def test_fallback_scan_computes_correct_extent(self):
        from thepexcel_mcp.domains.ranges import _spill_range_address

        anchor = MagicMock()
        anchor.SpillingRange = None  # force fallback scan path
        anchor.Row = 5
        anchor.Column = 2
        ws = MagicMock()
        anchor.Parent = ws

        def cells_side_effect(r, c):
            cell = MagicMock()
            cell.HasSpill = (5 <= r <= 7) and (2 <= c <= 3)
            return cell

        ws.Cells.side_effect = cells_side_effect

        def range_side_effect(a, b):
            result = MagicMock()
            result.Address = "$B$5:$C$7"
            return result

        ws.Range.side_effect = range_side_effect

        addr = _spill_range_address(anchor)

        assert addr == "$B$5:$C$7"
        ws.Cells.assert_any_call(7, 3)  # computed end cell: row 7, col 3

    def test_uses_com_property_when_available(self):
        from thepexcel_mcp.domains.ranges import _spill_range_address

        anchor = MagicMock()
        sr = MagicMock()
        sr.Address = "$A$1:$A$3"
        anchor.SpillingRange = sr

        assert _spill_range_address(anchor) == "$A$1:$A$3"
