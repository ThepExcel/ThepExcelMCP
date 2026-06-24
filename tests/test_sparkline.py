"""Unit tests for excel_sparkline (sparkline.py).

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
    """Minimal Worksheet mock for sparkline.py tests."""
    ws = MagicMock()
    ws.Name = name
    return ws


def _make_rng_mock(address: str = "$F$2:$F$5", sheet_name: str = "Sheet1"):
    """Range mock with SparklineGroups wired up."""
    ws = _make_sheet_mock(sheet_name)
    rng = MagicMock()
    rng.Address = address
    rng.Worksheet = ws
    # Default: no existing sparkline groups
    rng.SparklineGroups.Count = 0
    return rng


def _call_sparkline(action, mock_session, mock_rng, **kwargs):
    """Patch _session + resolve_range + excel_guard, invoke sparkline_action."""
    mock_session.get_sheet.return_value = mock_rng.Worksheet
    # _resolve_range calls ws.Range(...) which should return mock_rng
    mock_rng.Worksheet.Range.return_value = mock_rng

    mock_app = MagicMock()
    # Default: ReferenceStyle already xlA1
    mock_app.ReferenceStyle = 1
    mock_rng.Application = mock_app

    with patch("thepexcel_mcp.domains.sparkline._session", mock_session):
        with patch("thepexcel_mcp.domains.sparkline.excel_guard") as eg:
            eg.return_value.__enter__ = MagicMock(return_value=None)
            eg.return_value.__exit__ = MagicMock(return_value=False)
            from thepexcel_mcp.domains.sparkline import sparkline_action
            return sparkline_action(action, location="F2:F5", sheet=None, workbook=None, **kwargs)


# ── Action validation ──────────────────────────────────────────────────────────

class TestSparklineActionValidation:
    def test_unknown_action_raises(self):
        ms = make_mock_session()
        rng = _make_rng_mock()
        with pytest.raises(ToolError, match="Unknown action"):
            _call_sparkline("bogus", ms, rng)

    def test_add_requires_data_range(self):
        ms = make_mock_session()
        rng = _make_rng_mock()
        with pytest.raises(ToolError, match="data_range"):
            _call_sparkline("add", ms, rng, spark_type="line")

    def test_add_rejects_unknown_spark_type(self):
        ms = make_mock_session()
        rng = _make_rng_mock()
        with pytest.raises(ToolError, match="spark_type"):
            _call_sparkline("add", ms, rng, data_range="B2:E5", spark_type="pie")

    def test_add_rejects_bad_color_format(self):
        ms = make_mock_session()
        rng = _make_rng_mock()
        # Simulate SparklineGroups.Add succeeding but color is invalid
        rng.SparklineGroups.Add.return_value = MagicMock()
        rng.SparklineGroups.Count = 1
        with pytest.raises(ToolError, match="color"):
            _call_sparkline("add", ms, rng, data_range="B2:E5", color="red")

    def test_add_rejects_short_hex_color(self):
        ms = make_mock_session()
        rng = _make_rng_mock()
        with pytest.raises(ToolError, match="RRGGBB"):
            _call_sparkline("add", ms, rng, data_range="B2:E5", color="#FFF")


# ── add action ────────────────────────────────────────────────────────────────

class TestSparklineAdd:
    def _setup_add(self, spark_type="line"):
        ms = make_mock_session()
        rng = _make_rng_mock()
        mock_group = MagicMock()
        rng.SparklineGroups.Add.return_value = mock_group
        # After Add, Count becomes 1
        rng.SparklineGroups.Count = 1
        return ms, rng, mock_group

    def test_add_line_calls_sparklinegroups_add(self):
        ms, rng, group = self._setup_add()
        _call_sparkline("add", ms, rng, data_range="B2:E5", spark_type="line")
        rng.SparklineGroups.Add.assert_called_once_with(Type=1, SourceData="B2:E5")

    def test_add_column_uses_type_int_2(self):
        ms, rng, group = self._setup_add()
        _call_sparkline("add", ms, rng, data_range="B2:E5", spark_type="column")
        rng.SparklineGroups.Add.assert_called_once_with(Type=2, SourceData="B2:E5")

    def test_add_win_loss_uses_type_int_3(self):
        ms, rng, group = self._setup_add()
        _call_sparkline("add", ms, rng, data_range="B2:E5", spark_type="win_loss")
        rng.SparklineGroups.Add.assert_called_once_with(Type=3, SourceData="B2:E5")

    def test_add_sets_marker_visible_true(self):
        ms, rng, group = self._setup_add()
        _call_sparkline("add", ms, rng, data_range="B2:E5", marker=True)
        # VERIFY-EFFECT: group.Points.Markers.Visible set to True
        assert group.Points.Markers.Visible is True

    def test_add_sets_marker_visible_false(self):
        ms, rng, group = self._setup_add()
        _call_sparkline("add", ms, rng, data_range="B2:E5", marker=False)
        assert group.Points.Markers.Visible is False

    def test_add_without_marker_does_not_set_visible_bool(self):
        """When marker=None, Markers.Visible must not be set to True or False."""
        ms, rng, group = self._setup_add()
        _call_sparkline("add", ms, rng, data_range="B2:E5")
        # Markers.Visible is an auto-attribute on MagicMock; if we never
        # assigned it, it is still a MagicMock object (not True / not False).
        visible = group.Points.Markers.Visible
        assert visible is not True and visible is not False

    def test_add_sets_color_as_bgr(self):
        """color '#FF0000' (red) → BGR = 255."""
        ms, rng, group = self._setup_add()
        _call_sparkline("add", ms, rng, data_range="B2:E5", color="#FF0000")
        # VERIFY-EFFECT: SeriesColor.Color set to BGR int
        assert group.SeriesColor.Color == 255  # 0x0000FF = 255

    def test_add_color_gold_bgr(self):
        """ThepExcel Gold #D4A84B → B=0x4B=75, G=0xA8=168, R=0xD4=212 → BGR = (75<<16)|(168<<8)|212 = 4957396."""
        ms, rng, group = self._setup_add()
        _call_sparkline("add", ms, rng, data_range="B2:E5", color="#D4A84B")
        expected_bgr = (0x4B << 16) | (0xA8 << 8) | 0xD4  # 4957396
        assert group.SeriesColor.Color == expected_bgr

    def test_add_without_color_does_not_touch_series_color(self):
        ms, rng, group = self._setup_add()
        _call_sparkline("add", ms, rng, data_range="B2:E5")
        # SeriesColor.Color must NOT have been assigned an integer.
        # On a MagicMock, assigning an int replaces the child mock with that int;
        # if we never assigned it the attribute is still a MagicMock (not an int).
        assert not isinstance(group.SeriesColor.Color, int)

    def test_add_forces_xl_a1_reference_style(self):
        """Application.ReferenceStyle must be set to xlA1 (=1) before Add.

        Calls _dispatch directly (bypassing _call_sparkline) so we can
        control the mock_app that the domain code receives via rng.Application.
        """
        from thepexcel_mcp.domains.sparkline import _dispatch

        ms, rng, _ = self._setup_add()

        mock_app = MagicMock()
        mock_app.ReferenceStyle = 2  # simulate xlR1C1
        rng.Application = mock_app

        ref_style_at_add_time = []

        def capture_add(Type, SourceData):
            ref_style_at_add_time.append(mock_app.ReferenceStyle)
            return MagicMock()

        rng.SparklineGroups.Add.side_effect = capture_add
        rng.SparklineGroups.Count = 1  # ensure count check passes

        ws = rng.Worksheet
        ms.get_sheet.return_value = ws
        ws.Range.return_value = rng

        with patch("thepexcel_mcp.domains.sparkline._session", ms):
            with patch("thepexcel_mcp.domains.sparkline.excel_guard") as eg:
                eg.return_value.__enter__ = MagicMock(return_value=None)
                eg.return_value.__exit__ = MagicMock(return_value=False)
                _dispatch("add", "F2:F5", None, None, "B2:E5", "line", None, None)

        assert ref_style_at_add_time == [1], (
            f"ReferenceStyle must be 1 (xlA1) when Add() is called; "
            f"got {ref_style_at_add_time}"
        )

    def test_add_restores_reference_style_after_add(self):
        """ReferenceStyle must be restored to original value after Add().

        Calls _dispatch directly so we control the mock_app instance.
        """
        from thepexcel_mcp.domains.sparkline import _dispatch

        ms, rng, _ = self._setup_add()

        mock_app = MagicMock()
        mock_app.ReferenceStyle = 2  # xlR1C1
        rng.Application = mock_app
        rng.SparklineGroups.Count = 1

        ws = rng.Worksheet
        ms.get_sheet.return_value = ws
        ws.Range.return_value = rng

        with patch("thepexcel_mcp.domains.sparkline._session", ms):
            with patch("thepexcel_mcp.domains.sparkline.excel_guard") as eg:
                eg.return_value.__enter__ = MagicMock(return_value=None)
                eg.return_value.__exit__ = MagicMock(return_value=False)
                _dispatch("add", "F2:F5", None, None, "B2:E5", "line", None, None)

        # After the action, ReferenceStyle should be restored to 2
        assert mock_app.ReferenceStyle == 2

    def test_add_raises_when_count_zero_after_add(self):
        """If SparklineGroups.Count stays 0 after Add, raise ToolError."""
        ms = make_mock_session()
        rng = _make_rng_mock()
        rng.SparklineGroups.Add.return_value = MagicMock()
        # Count stays 0 after Add (simulates silent no-op)
        rng.SparklineGroups.Count = 0
        with pytest.raises(ToolError, match="Count is still 0"):
            _call_sparkline("add", ms, rng, data_range="B2:E5")

    def test_add_result_shape(self):
        ms, rng, _ = self._setup_add()
        result = _call_sparkline("add", ms, rng, data_range="B2:E5", spark_type="column")
        assert result["sparkline"] == "add"
        assert "sheet" in result
        assert "applied" in result
        assert result["applied"]["spark_type"] == "column"
        assert result["applied"]["type_int"] == 2
        assert result["applied"]["data_range"] == "B2:E5"
        assert result["applied"]["groups_count"] == 1

    def test_add_result_includes_marker_when_set(self):
        ms, rng, _ = self._setup_add()
        result = _call_sparkline("add", ms, rng, data_range="B2:E5", marker=True)
        assert result["applied"]["marker"] is True

    def test_add_result_includes_color_when_set(self):
        ms, rng, _ = self._setup_add()
        result = _call_sparkline("add", ms, rng, data_range="B2:E5", color="#00FF00")
        assert result["applied"]["color"] == "#00FF00"

    def test_add_result_omits_marker_when_not_set(self):
        ms, rng, _ = self._setup_add()
        result = _call_sparkline("add", ms, rng, data_range="B2:E5")
        assert "marker" not in result["applied"]

    def test_add_result_omits_color_when_not_set(self):
        ms, rng, _ = self._setup_add()
        result = _call_sparkline("add", ms, rng, data_range="B2:E5")
        assert "color" not in result["applied"]


# ── clear action ──────────────────────────────────────────────────────────────

class TestSparklineClear:
    def test_clear_calls_sparklinegroups_clear(self):
        ms = make_mock_session()
        rng = _make_rng_mock()
        rng.SparklineGroups.Count = 0  # after clear, count = 0

        _call_sparkline("clear", ms, rng)

        # VERIFY-EFFECT: Clear() was called
        rng.SparklineGroups.Clear.assert_called_once()

    def test_clear_result_shape(self):
        ms = make_mock_session()
        rng = _make_rng_mock()
        rng.SparklineGroups.Count = 0

        result = _call_sparkline("clear", ms, rng)

        assert result["sparkline"] == "clear"
        assert "sheet" in result
        assert "applied" in result
        assert result["applied"]["cleared"] is True

    def test_clear_reports_count_zero(self):
        ms = make_mock_session()
        rng = _make_rng_mock()
        rng.SparklineGroups.Count = 0

        result = _call_sparkline("clear", ms, rng)

        assert result["applied"]["groups_count"] == 0

    def test_clear_reports_location(self):
        ms = make_mock_session()
        rng = _make_rng_mock(address="$F$2:$F$10")
        rng.SparklineGroups.Count = 0

        result = _call_sparkline("clear", ms, rng)

        assert result["applied"]["location"] == "$F$2:$F$10"

    def test_clear_raises_when_count_nonzero_after_clear(self):
        """If SparklineGroups.Count is still > 0 after Clear, raise ToolError."""
        ms = make_mock_session()
        rng = _make_rng_mock()
        # Simulate silent no-op: Clear() succeeds but count stays 2
        rng.SparklineGroups.Count = 2

        with pytest.raises(ToolError, match="Count is still"):
            _call_sparkline("clear", ms, rng)

    def test_clear_com_error_wrapped(self):
        ms = make_mock_session()
        ms.wrap.side_effect = ToolError("sparkline clear failed: COM error")
        rng = _make_rng_mock()
        rng.SparklineGroups.Clear.side_effect = Exception("COM error")

        with pytest.raises(ToolError, match="sparkline clear failed"):
            _call_sparkline("clear", ms, rng)


# ── list action ───────────────────────────────────────────────────────────────

class TestSparklineList:
    def _setup_list(self, count: int = 2):
        ms = make_mock_session()
        rng = _make_rng_mock()

        # Build mock groups for the location range's SparklineGroups.
        # (list scopes to the caller's location range, not UsedRange — sparkline
        # destination cells are often value-empty and excluded from UsedRange.)
        groups_col = MagicMock()
        groups_col.Count = count
        rng.SparklineGroups = groups_col

        # Wire Item(i) for each group
        def make_group(i, type_int, src, loc):
            g = MagicMock()
            g.Type = type_int
            g.SourceData = src
            g.Location.Address = loc
            return g

        if count >= 1:
            groups_col.Item.side_effect = lambda i: {
                1: make_group(1, 1, "B2:E5", "$F$2:$F$5"),
                2: make_group(2, 2, "B6:E10", "$F$6:$F$10"),
            }.get(i, MagicMock())
        else:
            groups_col.Item.side_effect = lambda i: (_ for _ in ()).throw(IndexError("no item"))

        return ms, rng

    def test_list_returns_count(self):
        ms, rng = self._setup_list(count=2)
        result = _call_sparkline("list", ms, rng)
        assert result["applied"]["groups_count"] == 2

    def test_list_result_shape(self):
        ms, rng = self._setup_list(count=1)
        result = _call_sparkline("list", ms, rng)
        assert result["sparkline"] == "list"
        assert "sheet" in result
        assert "applied" in result
        assert "groups" in result["applied"]

    def test_list_enumerates_group_type_and_source(self):
        ms, rng = self._setup_list(count=2)
        result = _call_sparkline("list", ms, rng)
        groups = result["applied"]["groups"]
        assert len(groups) == 2
        # First group: xlSparkLine=1 → type "line"
        assert groups[0]["type"] == "line"
        assert groups[0]["type_int"] == 1
        assert groups[0]["source_data"] == "B2:E5"
        assert groups[0]["location"] == "$F$2:$F$5"

    def test_list_maps_column_type_name(self):
        ms, rng = self._setup_list(count=2)
        result = _call_sparkline("list", ms, rng)
        groups = result["applied"]["groups"]
        # Second group: xlSparkColumn=2 → type "column"
        assert groups[1]["type"] == "column"

    def test_list_empty_sheet_returns_count_zero(self):
        ms, rng = self._setup_list(count=0)
        result = _call_sparkline("list", ms, rng)
        assert result["applied"]["groups_count"] == 0
        assert result["applied"]["groups"] == []

    def test_list_handles_enumeration_error_gracefully(self):
        """If a group's properties throw, the entry records the error string."""
        ms = make_mock_session()
        rng = _make_rng_mock()

        groups_col = MagicMock()
        groups_col.Count = 1
        rng.SparklineGroups = groups_col

        # Item(1) raises an error
        groups_col.Item.side_effect = Exception("RPC error")

        result = _call_sparkline("list", ms, rng)

        assert result["applied"]["groups_count"] == 1
        # The error should be recorded in the group entry
        assert "error" in result["applied"]["groups"][0]


# ── hex_to_bgr helper ─────────────────────────────────────────────────────────

class TestHexToBgr:
    """Isolated tests for the _hex_to_bgr color converter."""

    def _bgr(self, hex_color):
        from thepexcel_mcp.domains.sparkline import _hex_to_bgr
        return _hex_to_bgr(hex_color)

    def test_red_converts_to_255(self):
        assert self._bgr("#FF0000") == 255  # B=0 G=0 R=255 → 0x0000FF

    def test_blue_converts_correctly(self):
        assert self._bgr("#0000FF") == 0xFF0000  # B=255 G=0 R=0 → 0xFF0000

    def test_green_converts_correctly(self):
        assert self._bgr("#00FF00") == 0x00FF00  # B=0 G=255 R=0 → 0x00FF00

    def test_thepexcel_gold(self):
        """#D4A84B → R=0xD4=212, G=0xA8=168, B=0x4B=75 → BGR=(75<<16)|(168<<8)|212."""
        expected = (0x4B << 16) | (0xA8 << 8) | 0xD4
        assert self._bgr("#D4A84B") == expected

    def test_lowercase_hex_accepted(self):
        assert self._bgr("#ff0000") == 255

    def test_invalid_format_raises(self):
        with pytest.raises(ToolError, match="RRGGBB"):
            self._bgr("#FFF")

    def test_non_hex_chars_raise(self):
        with pytest.raises(ToolError, match="hex"):
            self._bgr("#GGGGGG")
