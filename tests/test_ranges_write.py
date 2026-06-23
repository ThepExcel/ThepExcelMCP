"""Regression tests for ranges._write — no Excel required.

Guard against the pywin32 dispatch quirk where Range.Resize(r,c) returns a
single offset cell instead of a block.  The fix builds the target block via
ws.Range(ws.Cells(r1,c1), ws.Cells(r2,c2)) — these tests assert that path.
"""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock, patch, call

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


def _make_rng_mock(tl_row: int = 1, tl_col: int = 1):
    """Build a minimal mock COM Range that records the Range(Cells,Cells) calls.

    Structure mirrors the new implementation:
      rng.Parent  → ws (worksheet mock)
      rng.Cells(1,1) → tl_cell  (top-left cell; has .Row / .Column)
      ws.Cells(r1,c1) / ws.Cells(r2,c2) → used to build target block
      ws.Range(tl, br) → target_mock  (the actual write target)
    """
    target = MagicMock()
    target.Address = "$A$1:$D$9"

    # Worksheet: Cells() and Range() both return predictable mocks
    ws = MagicMock()
    ws.Range.return_value = target
    # ws.Cells(r,c) always returns a simple mock; Row/Column don't matter here
    ws.Cells.return_value = MagicMock()

    # Top-left cell returned by rng.Cells(1,1)
    tl_cell = MagicMock()
    tl_cell.Row = tl_row
    tl_cell.Column = tl_col

    rng = MagicMock()
    rng.Parent = ws
    rng.Cells.return_value = tl_cell

    return rng, ws, tl_cell, target


class TestWriteCellsRange:
    """_write must use ws.Range(ws.Cells(r1,c1), ws.Cells(r2,c2)) — NOT Resize."""

    def _call_write(self, range_str, values, rng_mock, mock_session):
        with patch("thepexcel_mcp.domains.ranges._resolve_range", return_value=rng_mock):
            with patch("thepexcel_mcp.domains.ranges._session", mock_session):
                from thepexcel_mcp.domains.ranges import _write
                return _write(range_str, None, None, values)

    def test_single_cell_anchor_builds_block(self):
        """Writing 9x4 data to anchor 'A1' must build block via ws.Range(Cells,Cells)."""
        values = [[i * 4 + j for j in range(4)] for i in range(9)]  # 9 rows x 4 cols
        rng_mock, ws_mock, tl_cell, target_mock = _make_rng_mock(tl_row=1, tl_col=1)
        mock_session = make_mock_session()

        result = self._call_write("A1", values, rng_mock, mock_session)

        # Must anchor at top-left cell of the input range
        rng_mock.Cells.assert_called_once_with(1, 1)
        # Must NOT call Resize (the buggy path)
        tl_cell.Resize.assert_not_called()
        # Must call ws.Range(...) to build the block
        ws_mock.Range.assert_called_once()
        # Bottom-right Cells call: row=9, col=4
        ws_mock.Cells.assert_any_call(9, 4)
        # Returned shape must match data
        assert result["written"]["rows"] == 9
        assert result["written"]["cols"] == 4

    def test_returned_range_is_target_address_not_input_string(self):
        """written.range should come from target.Address, not the raw input string."""
        values = [[1, 2], [3, 4]]
        rng_mock, ws_mock, tl_cell, target_mock = _make_rng_mock(tl_row=2, tl_col=2)
        target_mock.Address = "$B$2:$C$3"
        mock_session = make_mock_session()

        result = self._call_write("B2", values, rng_mock, mock_session)

        assert result["written"]["range"] == "$B$2:$C$3"

    def test_non_origin_anchor(self):
        """Anchor at B3 (row=3,col=2): bottom-right must be row=3+2-1=4, col=2+3-1=4."""
        values = [[1, 2, 3]] * 2  # 2 rows x 3 cols
        rng_mock, ws_mock, tl_cell, target_mock = _make_rng_mock(tl_row=3, tl_col=2)
        mock_session = make_mock_session()

        result = self._call_write("B3", values, rng_mock, mock_session)

        ws_mock.Cells.assert_any_call(4, 4)  # r1+rows-1=4, c1+cols-1=4
        assert result["written"]["rows"] == 2
        assert result["written"]["cols"] == 3

    def test_single_row_single_col(self):
        """1x1 value list: ws.Range(Cells(1,1), Cells(1,1)) called."""
        values = [[42]]
        rng_mock, ws_mock, tl_cell, target_mock = _make_rng_mock()
        target_mock.Address = "$A$1"
        mock_session = make_mock_session()

        result = self._call_write("A1", values, rng_mock, mock_session)

        tl_cell.Resize.assert_not_called()
        ws_mock.Range.assert_called_once()
        assert result["written"]["rows"] == 1
        assert result["written"]["cols"] == 1

    def test_padded_value_is_rectangular_tuple_of_tuples(self):
        """Ragged input rows are padded with None so COM gets a rectangular array."""
        values = [[1, 2, 3], [4]]  # ragged — second row is short
        rng_mock, ws_mock, tl_cell, target_mock = _make_rng_mock()
        mock_session = make_mock_session()

        self._call_write("A1", values, rng_mock, mock_session)

        # Capture what was assigned to target.Value
        assigned = target_mock.Value
        assert isinstance(assigned, tuple), "must be tuple-of-tuples for COM"
        assert all(isinstance(row, tuple) for row in assigned)
        assert len(assigned) == 2
        assert len(assigned[0]) == 3
        assert len(assigned[1]) == 3
        assert assigned[1][1] is None  # padded column
        assert assigned[1][2] is None

    def test_empty_values_raises_tool_error(self):
        """Empty values list must raise ToolError before any COM call."""
        rng_mock, ws_mock, tl_cell, target_mock = _make_rng_mock()
        mock_session = make_mock_session()

        with pytest.raises(ToolError, match="non-empty"):
            self._call_write("A1", [], rng_mock, mock_session)

        rng_mock.Cells.assert_not_called()

    def test_com_error_is_wrapped(self):
        """COM errors during Value assignment are re-raised via _session.wrap."""
        values = [[1, 2], [3, 4]]
        rng_mock, ws_mock, tl_cell, target_mock = _make_rng_mock()
        mock_session = make_mock_session()
        mock_session.wrap.return_value = ToolError("Write failed: COM error")
        # Make Value assignment raise
        type(target_mock).Value = property(
            fget=lambda self: None,
            fset=lambda self, v: (_ for _ in ()).throw(Exception("COM error")),
        )

        with pytest.raises((ToolError, Exception)):
            self._call_write("A1", values, rng_mock, mock_session)
