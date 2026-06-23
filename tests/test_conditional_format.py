"""Unit tests for conditional_format.py — excel_conditional_format domain.

No Excel required — all COM calls are intercepted via make_mock_session.
VERIFY-EFFECT mindset: every test asserts the SPECIFIC COM property/method was
called with the SPECIFIC value, not just "no exception raised".
"""

from __future__ import annotations

from unittest.mock import MagicMock, call, patch

import pytest

from conftest import make_mock_session
from fastmcp.exceptions import ToolError


# ── Test helper ───────────────────────────────────────────────────────────────

def _make_range_mock(address: str = "$A$1:$C$10") -> MagicMock:
    """Minimal COM Range mock for conditional_format.py tests."""
    rng = MagicMock()
    rng.Address = address
    rng.Application = MagicMock()

    # FormatConditions mock
    fc = MagicMock()
    rng.FormatConditions = fc

    # AddDatabar → databar object
    db = MagicMock()
    db.BarColor = MagicMock()
    fc.AddDatabar.return_value = db

    # AddColorScale → colorscale object
    cs = MagicMock()
    fc.AddColorScale.return_value = cs

    # AddIconSetCondition → iconset condition object
    ic = MagicMock()
    fc.AddIconSetCondition.return_value = ic

    # Add (cell_rule) → condition object with Interior + Font
    cond = MagicMock()
    cond.Interior = MagicMock()
    cond.Font = MagicMock()
    fc.Add.return_value = cond

    # AddTop10 → top10 object with Interior
    tb = MagicMock()
    tb.Interior = MagicMock()
    fc.AddTop10.return_value = tb

    # rng.Parent.Parent = Workbook (for icon_set wb.IconSets lookup)
    wb = MagicMock()
    ws = MagicMock()
    ws.Parent = wb
    rng.Parent = ws

    return rng


def _call_cf(action, rng_mock, mock_session, **kwargs):
    """Patch _resolve_range + _session + excel_guard; call conditional_format_action."""
    with patch("thepexcel_mcp.domains.conditional_format._resolve_range",
               return_value=rng_mock):
        with patch("thepexcel_mcp.domains.conditional_format._session", mock_session):
            with patch("thepexcel_mcp.domains.conditional_format.excel_guard") as eg:
                eg.return_value.__enter__ = MagicMock(return_value=None)
                eg.return_value.__exit__ = MagicMock(return_value=False)
                from thepexcel_mcp.domains.conditional_format import (
                    conditional_format_action,
                )
                return conditional_format_action(
                    action,
                    range="A1:C10",
                    sheet=None,
                    workbook=None,
                    **kwargs,
                )


# ── Dispatch validation ────────────────────────────────────────────────────────

class TestDispatchValidation:
    def test_unknown_action_raises(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="Unknown action"):
            _call_cf("bogus", rng, mock_session)

    def test_valid_action_names_in_error_message(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError) as exc_info:
            _call_cf("xyz", rng, mock_session)
        msg = str(exc_info.value)
        for name in ("cell_rule", "clear", "color_scale", "data_bar", "icon_set", "top_bottom"):
            assert name in msg


# ── data_bar ──────────────────────────────────────────────────────────────────

class TestDataBar:
    def test_addDatabar_called(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_cf("data_bar", rng, mock_session)
        rng.FormatConditions.AddDatabar.assert_called_once()
        assert result["conditional_format"] == "data_bar"
        assert result["range"] == "$A$1:$C$10"

    def test_color_set_on_BarColor(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        # #FF0000 red → BGR = 255
        result = _call_cf("data_bar", rng, mock_session, color="#FF0000")
        db = rng.FormatConditions.AddDatabar.return_value
        assert db.BarColor.Color == 255  # BGR of #FF0000
        assert result["applied"]["color"] == "#FF0000"

    def test_no_color_does_not_set_BarColor(self):
        """When color=None, BarColor.Color must NOT be touched.

        MagicMock attribute assignment is not reliably tracked via mock_calls,
        so we use a simple namespace object for BarColor and assert it has no
        Color attribute after the call (attribute assignment would create it).
        """
        import types
        mock_session = make_mock_session()
        rng = _make_range_mock()

        # Replace BarColor with a plain namespace — has no Color attr by default.
        # If the impl assigns db.BarColor.Color = ..., the namespace gains the attr.
        bar_color_ns = types.SimpleNamespace()
        db = rng.FormatConditions.AddDatabar.return_value
        db.BarColor = bar_color_ns

        _call_cf("data_bar", rng, mock_session)

        assert not hasattr(bar_color_ns, "Color"), (
            "BarColor.Color must NOT be set when color=None, "
            f"but it was set to {bar_color_ns.Color!r}"
        )

    def test_applied_empty_when_no_color(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_cf("data_bar", rng, mock_session)
        assert result["applied"] == {}


# ── color_scale ───────────────────────────────────────────────────────────────

class TestColorScale:
    def test_three_color_scale(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_cf("color_scale", rng, mock_session, scale_type=3)
        rng.FormatConditions.AddColorScale.assert_called_once_with(3)
        assert result["conditional_format"] == "color_scale"
        assert result["applied"]["scale_type"] == 3

    def test_two_color_scale(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_cf("color_scale", rng, mock_session, scale_type=2)
        rng.FormatConditions.AddColorScale.assert_called_once_with(2)
        assert result["applied"]["scale_type"] == 2

    def test_invalid_scale_type_raises(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="scale_type"):
            _call_cf("color_scale", rng, mock_session, scale_type=4)

    def test_scale_type_1_raises(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="scale_type"):
            _call_cf("color_scale", rng, mock_session, scale_type=1)


# ── icon_set ──────────────────────────────────────────────────────────────────

class TestIconSet:
    def test_default_traffic_lights(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_cf("icon_set", rng, mock_session, style="3traffic_lights")
        rng.FormatConditions.AddIconSetCondition.assert_called_once()
        # IconSet assigned from wb.IconSets(4) — value 4 = xl3TrafficLights1
        wb = rng.Parent.Parent
        wb.IconSets.assert_called_once_with(4)
        ic = rng.FormatConditions.AddIconSetCondition.return_value
        assert ic.IconSet == wb.IconSets.return_value
        assert result["conditional_format"] == "icon_set"
        assert result["applied"]["style"] == "3traffic_lights"
        assert result["applied"]["icon_set_const"] == 4

    def test_5arrows_style(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_cf("icon_set", rng, mock_session, style="5arrows")
        wb = rng.Parent.Parent
        wb.IconSets.assert_called_once_with(13)  # xl5Arrows = 13
        assert result["applied"]["icon_set_const"] == 13

    def test_invalid_style_raises(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="Unknown icon_set style"):
            _call_cf("icon_set", rng, mock_session, style="rainbow")

    def test_case_insensitive_style(self):
        """Style lookup is case-insensitive."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_cf("icon_set", rng, mock_session, style="3TRAFFIC_LIGHTS")
        assert result["applied"]["icon_set_const"] == 4


# ── cell_rule ─────────────────────────────────────────────────────────────────

class TestCellRule:
    def test_greater_than_rule(self):
        """operator=greater → Operator const 5, FormatConditions.Add called correctly."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_cf(
            "cell_rule", rng, mock_session,
            operator="greater", formula1="100",
        )
        rng.FormatConditions.Add.assert_called_once_with(
            Type=1,       # xlCellValue
            Operator=5,   # xlGreater
            Formula1="100",
        )
        assert result["conditional_format"] == "cell_rule"
        assert result["applied"]["operator"] == "greater"
        assert result["applied"]["formula1"] == "100"

    def test_between_requires_formula2(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="formula2"):
            _call_cf("cell_rule", rng, mock_session,
                     operator="between", formula1="10")

    def test_between_passes_both_formulas(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_cf(
            "cell_rule", rng, mock_session,
            operator="between", formula1="10", formula2="50",
        )
        rng.FormatConditions.Add.assert_called_once_with(
            Type=1,
            Operator=1,   # xlBetween
            Formula1="10",
            Formula2="50",
        )

    def test_fill_color_set_on_condition_interior(self):
        """fill_color must be set on the RETURNED condition object, not the range."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_cf(
            "cell_rule", rng, mock_session,
            operator="equal", formula1="1", fill_color="#FFFF00",
        )
        cond = rng.FormatConditions.Add.return_value
        # Yellow #FFFF00 → BGR = (0 << 16) | (255 << 8) | 255 = 65535
        assert cond.Interior.Color == 65535

    def test_font_color_set_on_condition_font(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_cf(
            "cell_rule", rng, mock_session,
            operator="less", formula1="0", font_color="#FF0000",
        )
        cond = rng.FormatConditions.Add.return_value
        # Red #FF0000 → BGR = 255
        assert cond.Font.Color == 255

    def test_missing_operator_raises(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="operator"):
            _call_cf("cell_rule", rng, mock_session, formula1="100")

    def test_missing_formula1_raises(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="formula1"):
            _call_cf("cell_rule", rng, mock_session, operator="greater")

    def test_unknown_operator_raises(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="Unknown operator"):
            _call_cf("cell_rule", rng, mock_session,
                     operator="contains", formula1="foo")

    def test_all_operator_constants(self):
        """Spot-check every operator maps to the documented VBA enum integer."""
        from thepexcel_mcp.domains.conditional_format import _OPERATORS
        expected = {
            "between": 1, "not_between": 2, "equal": 3, "not_equal": 4,
            "greater": 5, "less": 6, "greater_equal": 7, "less_equal": 8,
        }
        for name, value in expected.items():
            assert _OPERATORS[name] == value, (
                f"Operator '{name}' expected {value}, got {_OPERATORS[name]}"
            )

    def test_not_between_applied_keys(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_cf(
            "cell_rule", rng, mock_session,
            operator="not_between", formula1="5", formula2="95",
        )
        assert "formula2" in result["applied"]
        assert result["applied"]["formula2"] == "95"


# ── top_bottom ────────────────────────────────────────────────────────────────

class TestTopBottom:
    def test_top_10_count(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_cf("top_bottom", rng, mock_session,
                          kind="top", rank=10, percent=False)
        rng.FormatConditions.AddTop10.assert_called_once()
        tb = rng.FormatConditions.AddTop10.return_value
        assert tb.TopBottom == 1   # xlTop10Top
        assert tb.Rank == 10
        assert tb.Percent is False
        assert result["conditional_format"] == "top_bottom"
        assert result["applied"]["kind"] == "top"

    def test_bottom_20_percent(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_cf("top_bottom", rng, mock_session,
                          kind="bottom", rank=20, percent=True)
        tb = rng.FormatConditions.AddTop10.return_value
        assert tb.TopBottom == 0   # xlTop10Bottom
        assert tb.Rank == 20
        assert tb.Percent is True
        assert result["applied"]["percent"] is True

    def test_fill_color_on_top10_interior(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_cf("top_bottom", rng, mock_session,
                 kind="top", rank=5, percent=False, fill_color="#0000FF")
        tb = rng.FormatConditions.AddTop10.return_value
        # Blue #0000FF → BGR = 0xFF0000 = 16711680
        assert tb.Interior.Color == 0xFF0000

    def test_invalid_kind_raises(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="kind"):
            _call_cf("top_bottom", rng, mock_session, kind="middle")

    def test_top_bottom_constants_correct(self):
        """Verify xlTop10Bottom=0 and xlTop10Top=1 from MS Learn docs."""
        from thepexcel_mcp.domains.conditional_format import (
            _XL_TOP10_BOTTOM, _XL_TOP10_TOP,
        )
        assert _XL_TOP10_BOTTOM == 0
        assert _XL_TOP10_TOP == 1


# ── clear ─────────────────────────────────────────────────────────────────────

class TestClear:
    def test_delete_called(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_cf("clear", rng, mock_session)
        rng.FormatConditions.Delete.assert_called_once()
        assert result["conditional_format"] == "clear"
        assert result["applied"] == {}

    def test_range_preserved_in_result(self):
        mock_session = make_mock_session()
        rng = _make_range_mock("$D$5:$F$20")
        result = _call_cf("clear", rng, mock_session)
        assert result["range"] == "$D$5:$F$20"


# ── _hex_to_bgr reuse ─────────────────────────────────────────────────────────

class TestHexToBgrReuse:
    def test_uses_format_module_function(self):
        """conditional_format imports _hex_to_bgr from format — not duplicated."""
        import thepexcel_mcp.domains.conditional_format as cf_mod
        import thepexcel_mcp.domains.format as fmt_mod
        assert cf_mod._hex_to_bgr is fmt_mod._hex_to_bgr

    def test_resolve_range_reuse(self):
        """_resolve_range is imported from format — not reimplemented."""
        import thepexcel_mcp.domains.conditional_format as cf_mod
        import thepexcel_mcp.domains.format as fmt_mod
        assert cf_mod._resolve_range is fmt_mod._resolve_range
