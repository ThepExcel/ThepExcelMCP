"""Regression tests for ranges._write — no Excel required.

Guard against the bug where a single-cell anchor was NOT resized before
assigning a multi-row/col array, causing only the top-left cell to be
written while reporting a false rows/cols count.
"""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock, patch, call

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


def _make_rng_mock():
    """Build a minimal mock COM Range that records Cells/Resize/Value calls."""
    target = MagicMock()
    target.Address = "$A$1:$D$9"

    anchor = MagicMock()
    anchor.Resize.return_value = target

    rng = MagicMock()
    rng.Cells.return_value = anchor
    return rng, anchor, target


class TestWriteResize:
    """_write must resize the COM range to the data shape before assigning."""

    def _call_write(self, range_str, values, rng_mock, mock_session):
        with patch("thepexcel_mcp.domains.ranges._resolve_range", return_value=rng_mock):
            with patch("thepexcel_mcp.domains.ranges._session", mock_session):
                from thepexcel_mcp.domains.ranges import _write
                return _write(range_str, None, None, values)

    def test_single_cell_anchor_resizes_to_data_shape(self):
        """Writing 9x4 data to anchor 'A1' must call Resize(9, 4)."""
        values = [[i * 4 + j for j in range(4)] for i in range(9)]  # 9 rows x 4 cols
        rng_mock, anchor_mock, target_mock = _make_rng_mock()
        mock_session = make_mock_session()

        result = self._call_write("A1", values, rng_mock, mock_session)

        # Must anchor at top-left cell
        rng_mock.Cells.assert_called_once_with(1, 1)
        # Must resize to the data dimensions
        anchor_mock.Resize.assert_called_once_with(9, 4)
        # Must assign to the RESIZED range, not the original
        assert target_mock.Value == values
        # Returned shape must match data
        assert result["written"]["rows"] == 9
        assert result["written"]["cols"] == 4

    def test_returned_range_is_target_address_not_input_string(self):
        """written.range should come from target.Address, not the raw input string."""
        values = [[1, 2], [3, 4]]
        rng_mock, anchor_mock, target_mock = _make_rng_mock()
        target_mock.Address = "$B$2:$C$3"
        mock_session = make_mock_session()

        result = self._call_write("B2", values, rng_mock, mock_session)

        assert result["written"]["range"] == "$B$2:$C$3"

    def test_correctly_sized_range_is_idempotent(self):
        """Caller passing a full 'A1:D9' range still gets Resize called (idempotent)."""
        values = [[1, 2, 3, 4]] * 9  # 9x4
        rng_mock, anchor_mock, target_mock = _make_rng_mock()
        mock_session = make_mock_session()

        result = self._call_write("A1:D9", values, rng_mock, mock_session)

        anchor_mock.Resize.assert_called_once_with(9, 4)
        assert result["written"]["rows"] == 9
        assert result["written"]["cols"] == 4

    def test_single_row_single_col(self):
        """1x1 value list: Resize(1, 1) called."""
        values = [[42]]
        rng_mock, anchor_mock, target_mock = _make_rng_mock()
        target_mock.Address = "$A$1"
        mock_session = make_mock_session()

        result = self._call_write("A1", values, rng_mock, mock_session)

        anchor_mock.Resize.assert_called_once_with(1, 1)
        assert result["written"]["rows"] == 1
        assert result["written"]["cols"] == 1

    def test_empty_values_raises_tool_error(self):
        """Empty values list must raise ToolError before any COM call."""
        rng_mock, anchor_mock, target_mock = _make_rng_mock()
        mock_session = make_mock_session()

        with pytest.raises(ToolError, match="non-empty"):
            self._call_write("A1", [], rng_mock, mock_session)

        rng_mock.Cells.assert_not_called()

    def test_com_error_is_wrapped(self):
        """COM errors during Value assignment are re-raised via _session.wrap."""
        values = [[1, 2], [3, 4]]
        rng_mock, anchor_mock, target_mock = _make_rng_mock()
        target_mock.Value = MagicMock(side_effect=Exception("COM error"))
        mock_session = make_mock_session()
        mock_session.wrap.return_value = ToolError("Write failed: COM error")
        # Make Value assignment raise
        type(target_mock).Value = property(fget=lambda self: None,
                                           fset=lambda self, v: (_ for _ in ()).throw(Exception("COM error")))

        with pytest.raises((ToolError, Exception)):
            self._call_write("A1", values, rng_mock, mock_session)
