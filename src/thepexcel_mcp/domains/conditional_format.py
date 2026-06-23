"""Range-based conditional formatting (FormatConditions).

Reuses format.py's _resolve_range and _hex_to_bgr — no duplication.

Supported actions
-----------------
data_bar     — AddDataBar; optionally set bar color.
color_scale  — AddColorScale(ColorScaleType=2|3); 2-color or 3-color gradient.
icon_set     — AddIconSetCondition; choose from named icon-set styles.
cell_rule    — FormatConditions.Add(Type=xlCellValue) with operator + colors.
top_bottom   — AddTop10; top/bottom N (count or percent).
clear        — FormatConditions.Delete() — wipes ALL CF rules on the range.

COM note: FormatConditions.Add (and the Add* variants) RETURN the new condition
object.  All format properties (Interior.Color, Font.Color, IconSet, …) must be
set on THAT returned object, not on the range itself.
"""

from __future__ import annotations

from fastmcp.exceptions import ToolError

from ..session import ExcelSession, excel_guard
from .format import _hex_to_bgr, _resolve_range  # reuse, no duplication

_session = ExcelSession()

# ── Verified enum constants (Microsoft Learn, 2024-04-01) ──────────────────────

# XlFormatConditionType (used in FormatConditions.Add Type= argument)
_XL_CELL_VALUE = 1   # xlCellValue

# XlFormatConditionOperator → used in FormatConditions.Add Operator= argument
_OPERATORS = {
    "between":        1,  # xlBetween
    "not_between":    2,  # xlNotBetween
    "equal":          3,  # xlEqual
    "not_equal":      4,  # xlNotEqual
    "greater":        5,  # xlGreater
    "less":           6,  # xlLess
    "greater_equal":  7,  # xlGreaterEqual
    "less_equal":     8,  # xlLessEqual
}

# XlTopBottom (Top10.TopBottom property)
_XL_TOP10_BOTTOM = 0   # xlTop10Bottom
_XL_TOP10_TOP    = 1   # xlTop10Top

# XlIconSet (IconSets index values) — verified from Microsoft Learn
_ICON_SET_STYLES = {
    "3arrows":          1,   # xl3Arrows
    "3arrows_gray":     2,   # xl3ArrowsGray
    "3flags":           3,   # xl3Flags
    "3traffic_lights":  4,   # xl3TrafficLights1  (default)
    "3traffic_lights2": 5,   # xl3TrafficLights2 (rimmed circles)
    "3signs":           6,   # xl3Signs
    "3symbols":         7,   # xl3Symbols
    "4arrows":          8,   # xl4Arrows
    "4arrows_gray":     9,   # xl4ArrowsGray
    "4red_to_black":    10,  # xl4RedToBlack
    "4crv":             11,  # xl4CRV
    "4traffic_lights":  12,  # xl4TrafficLights
    "5arrows":          13,  # xl5Arrows
    "5arrows_gray":     14,  # xl5ArrowsGray
    "5crv":             15,  # xl5CRV
    "5quarters":        16,  # xl5Quarters
}


# ── Public entry point ─────────────────────────────────────────────────────────

def conditional_format_action(
    action: str,
    range: str,
    sheet: str | None = None,
    workbook: str | None = None,
    # data_bar
    color: str | None = None,
    # color_scale
    scale_type: int = 3,
    # icon_set
    style: str = "3traffic_lights",
    # cell_rule
    operator: str | None = None,
    formula1: str | None = None,
    formula2: str | None = None,
    fill_color: str | None = None,
    font_color: str | None = None,
    # top_bottom
    kind: str = "top",
    rank: int = 10,
    percent: bool = False,
) -> dict:
    """Dispatch a conditional-formatting action.

    Actions
    -------
    data_bar
        Add a data-bar conditional format to the range.
    color_scale
        Add a 2- or 3-color scale conditional format.
    icon_set
        Add an icon-set conditional format.
    cell_rule
        Add a cell-value rule (greater-than, between, equal, …) with
        optional fill and font color.
    top_bottom
        Add a Top N / Bottom N conditional format.
    clear
        Delete all conditional formatting rules on the range.
    """
    valid = {"data_bar", "color_scale", "icon_set", "cell_rule", "top_bottom", "clear"}
    if action not in valid:
        raise ToolError(
            f"Unknown action '{action}'. Valid: {', '.join(sorted(valid))}."
        )
    return _session.run_com(
        _dispatch,
        action, range, sheet, workbook,
        color,
        scale_type,
        style,
        operator, formula1, formula2, fill_color, font_color,
        kind, rank, percent,
    )


# ── Worker-thread dispatcher ───────────────────────────────────────────────────

def _dispatch(
    action, range_str, sheet, workbook,
    color,
    scale_type,
    style,
    operator, formula1, formula2, fill_color, font_color,
    kind, rank, percent,
) -> dict:
    """Executed on the STA COM worker thread."""
    rng = _resolve_range(range_str, sheet, workbook)
    app = rng.Application

    with excel_guard(app):
        if action == "data_bar":
            return _data_bar(rng, color)
        if action == "color_scale":
            return _color_scale(rng, scale_type)
        if action == "icon_set":
            return _icon_set(rng, style)
        if action == "cell_rule":
            return _cell_rule(rng, operator, formula1, formula2, fill_color, font_color)
        if action == "top_bottom":
            return _top_bottom(rng, kind, rank, percent, fill_color)
        # action == "clear"
        return _clear(rng)


# ── Action implementations ─────────────────────────────────────────────────────

def _data_bar(rng, color: str | None) -> dict:
    try:
        db = rng.FormatConditions.AddDatabar()
        if color is not None:
            db.BarColor.Color = _hex_to_bgr(color)
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "data_bar failed")
    applied: dict = {}
    if color is not None:
        applied["color"] = color
    return {"conditional_format": "data_bar", "range": rng.Address, "applied": applied}


def _color_scale(rng, scale_type: int) -> dict:
    if scale_type not in (2, 3):
        raise ToolError(
            f"scale_type must be 2 (two-color) or 3 (three-color), got {scale_type!r}."
        )
    try:
        rng.FormatConditions.AddColorScale(scale_type)
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "color_scale failed")
    return {
        "conditional_format": "color_scale",
        "range": rng.Address,
        "applied": {"scale_type": scale_type},
    }


def _icon_set(rng, style: str) -> dict:
    icon_const = _ICON_SET_STYLES.get(style.lower())
    if icon_const is None:
        raise ToolError(
            f"Unknown icon_set style '{style}'. "
            f"Valid: {', '.join(sorted(_ICON_SET_STYLES))}."
        )
    try:
        ic = rng.FormatConditions.AddIconSetCondition()
        wb = rng.Parent.Parent  # Worksheet.Parent = Workbook
        ic.IconSet = wb.IconSets(icon_const)
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "icon_set failed")
    return {
        "conditional_format": "icon_set",
        "range": rng.Address,
        "applied": {"style": style, "icon_set_const": icon_const},
    }


def _cell_rule(
    rng,
    operator: str | None,
    formula1: str | None,
    formula2: str | None,
    fill_color: str | None,
    font_color: str | None,
) -> dict:
    if operator is None:
        raise ToolError(
            "cell_rule action requires 'operator'. "
            f"Valid: {', '.join(sorted(_OPERATORS))}."
        )
    op_const = _OPERATORS.get(operator.lower())
    if op_const is None:
        raise ToolError(
            f"Unknown operator '{operator}'. "
            f"Valid: {', '.join(sorted(_OPERATORS))}."
        )
    if formula1 is None:
        raise ToolError("cell_rule action requires 'formula1'.")
    # between / not_between require formula2
    if operator.lower() in ("between", "not_between") and formula2 is None:
        raise ToolError(
            f"operator='{operator}' requires 'formula2' as the second bound."
        )

    try:
        if formula2 is not None:
            cond = rng.FormatConditions.Add(
                Type=_XL_CELL_VALUE,
                Operator=op_const,
                Formula1=formula1,
                Formula2=formula2,
            )
        else:
            cond = rng.FormatConditions.Add(
                Type=_XL_CELL_VALUE,
                Operator=op_const,
                Formula1=formula1,
            )
        if fill_color is not None:
            cond.Interior.Color = _hex_to_bgr(fill_color)
        if font_color is not None:
            cond.Font.Color = _hex_to_bgr(font_color)
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "cell_rule failed")

    applied: dict = {"operator": operator, "formula1": formula1}
    if formula2 is not None:
        applied["formula2"] = formula2
    if fill_color is not None:
        applied["fill_color"] = fill_color
    if font_color is not None:
        applied["font_color"] = font_color
    return {"conditional_format": "cell_rule", "range": rng.Address, "applied": applied}


def _top_bottom(
    rng,
    kind: str,
    rank: int,
    percent: bool,
    fill_color: str | None,
) -> dict:
    kind_lower = kind.lower()
    if kind_lower not in ("top", "bottom"):
        raise ToolError(
            f"kind must be 'top' or 'bottom', got '{kind}'."
        )
    try:
        tb = rng.FormatConditions.AddTop10()
        tb.TopBottom = _XL_TOP10_TOP if kind_lower == "top" else _XL_TOP10_BOTTOM
        tb.Rank = rank
        tb.Percent = percent
        if fill_color is not None:
            tb.Interior.Color = _hex_to_bgr(fill_color)
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "top_bottom failed")

    applied: dict = {"kind": kind, "rank": rank, "percent": percent}
    if fill_color is not None:
        applied["fill_color"] = fill_color
    return {"conditional_format": "top_bottom", "range": rng.Address, "applied": applied}


def _clear(rng) -> dict:
    try:
        rng.FormatConditions.Delete()
    except Exception as e:
        raise _session.wrap(e, "clear conditional formats failed")
    return {"conditional_format": "clear", "range": rng.Address, "applied": {}}
