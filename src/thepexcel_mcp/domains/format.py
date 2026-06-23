"""Range formatting operations: font, fill, border, number_format, alignment,
column_width, row_height, autofit.

Color input convention
----------------------
All color parameters accept an ``"#RRGGBB"`` hex string (e.g. ``"#FF0000"`` for
red).  Internally Excel COM expects a **BGR integer** (0xBBGGRR), so we convert
on entry via _hex_to_bgr().  The conversion is transparent to callers.
"""

from __future__ import annotations

from contextlib import contextmanager

from fastmcp.exceptions import ToolError

from ..session import ExcelSession, excel_guard

_session = ExcelSession()

# XlBordersIndex values (from Excel VBA enum)
_XL_EDGE_LEFT          =  7
_XL_EDGE_TOP           =  8
_XL_EDGE_BOTTOM        =  9
_XL_EDGE_RIGHT         = 10
_XL_INSIDE_VERTICAL    = 11
_XL_INSIDE_HORIZONTAL  = 12

# XlBorderWeight
_BORDER_WEIGHT = {
    "thin":   2,   # xlThin
    "medium": -4138,  # xlMedium
    "thick":  4,   # xlThick
}

# XlLineStyle
_LINE_STYLE = {
    "continuous": 1,     # xlContinuous
    "dash":       -4115, # xlDash
    "double":     -4119, # xlDouble
    "none":       -4142, # xlLineStyleNone
    "thin":       1,     # alias → continuous
}

# XlHAlign
_HALIGN = {
    "general": 1,    # xlHAlignGeneral
    "left":    -4131, # xlHAlignLeft
    "center":  -4108, # xlHAlignCenter
    "right":   -4152, # xlHAlignRight
}

# XlVAlign
_VALIGN = {
    "top":    -4160, # xlVAlignTop
    "center": -4108, # xlVAlignCenter
    "bottom": -4107, # xlVAlignBottom
}


# ── Public entry point ─────────────────────────────────────────────────────────

def format_action(
    action: str,
    range: str,
    sheet: str | None = None,
    workbook: str | None = None,
    # font
    font_name: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    font_color: str | None = None,
    # fill
    fill_color: str | None = None,
    clear_fill: bool = False,
    # border
    border_sides: str = "outline",
    border_style: str = "continuous",
    border_weight: str = "thin",
    border_color: str | None = None,
    # number_format
    number_format: str | None = None,
    # alignment
    horizontal: str | None = None,
    vertical: str | None = None,
    wrap_text: bool | None = None,
    merge: bool | None = None,
    # column/row sizing
    width: float | None = None,
    height: float | None = None,
    # autofit
    autofit_columns: bool = True,
    autofit_rows: bool = False,
) -> dict:
    """Dispatch a formatting action.

    Actions
    -------
    font
        Set font name, size, bold, italic, underline, color.
    fill
        Set background (interior) color, or clear fill.
    border
        Apply borders on sides (outline/all/top/bottom/left/right/inside).
    number_format
        Set a NumberFormat code (e.g. ``"#,##0.00"``, ``"0.0%"``).
    alignment
        Set horizontal/vertical alignment, wrap_text, and optionally merge/unmerge.
    column_width
        Set explicit column width in character-width units.
    row_height
        Set explicit row height in points.
    autofit
        AutoFit columns and/or rows (autofit_columns, autofit_rows).
    """
    valid = {
        "font", "fill", "border", "number_format",
        "alignment", "column_width", "row_height", "autofit",
    }
    if action not in valid:
        raise ToolError(
            f"Unknown action '{action}'. Valid: {', '.join(sorted(valid))}."
        )
    return _session.run_com(
        _dispatch,
        action, range, sheet, workbook,
        font_name, font_size, bold, italic, underline, font_color,
        fill_color, clear_fill,
        border_sides, border_style, border_weight, border_color,
        number_format,
        horizontal, vertical, wrap_text, merge,
        width, height,
        autofit_columns, autofit_rows,
    )


# ── Worker-thread dispatcher ───────────────────────────────────────────────────

def _dispatch(
    action, range_str, sheet, workbook,
    font_name, font_size, bold, italic, underline, font_color,
    fill_color, clear_fill,
    border_sides, border_style, border_weight, border_color,
    number_format,
    horizontal, vertical, wrap_text, merge,
    width, height,
    autofit_columns, autofit_rows,
) -> dict:
    """Executed on the STA COM worker thread."""
    rng = _resolve_range(range_str, sheet, workbook)
    app = rng.Application

    with excel_guard(app):
        if action == "font":
            return _font(rng, font_name, font_size, bold, italic, underline, font_color)
        if action == "fill":
            return _fill(rng, fill_color, clear_fill)
        if action == "border":
            return _border(rng, border_sides, border_style, border_weight, border_color)
        if action == "number_format":
            return _number_format(rng, number_format)
        if action == "alignment":
            return _alignment(rng, horizontal, vertical, wrap_text, merge)
        if action == "column_width":
            return _column_width(rng, width)
        if action == "row_height":
            return _row_height(rng, height)
        # action == "autofit"
        return _autofit(rng, autofit_columns, autofit_rows)


# ── Range resolver (mirrors ranges.py pattern) ────────────────────────────────

def _resolve_range(range_str: str, sheet: str | None, workbook: str | None):
    if "!" in range_str:
        sheet_part, cell_part = range_str.split("!", 1)
        ws = _session.get_sheet(sheet_part.strip("'"), workbook)
        try:
            return ws.Range(cell_part)
        except Exception as e:
            raise _session.wrap(e, f"Invalid range '{cell_part}' on sheet '{sheet_part}'")
    ws = _session.get_sheet(sheet, workbook)
    try:
        return ws.Range(range_str)
    except Exception as e:
        raise _session.wrap(e, f"Invalid range '{range_str}'")


# ── Color helper ──────────────────────────────────────────────────────────────

def _hex_to_bgr(hex_color: str) -> int:
    """Convert ``"#RRGGBB"`` to Excel BGR integer (0xBBGGRR).

    Excel Interior.Color / Font.Color expect BGR byte order, not RGB.
    Example: "#FF0000" (red) → 0x0000FF (255).
    """
    h = hex_color.lstrip("#")
    if len(h) != 6:
        raise ToolError(
            f"Invalid color '{hex_color}'. Use '#RRGGBB' format (e.g. '#FF0000')."
        )
    try:
        r = int(h[0:2], 16)
        g = int(h[2:4], 16)
        b = int(h[4:6], 16)
    except ValueError:
        raise ToolError(
            f"Invalid color '{hex_color}'. Must be a valid hex string like '#D4A84B'."
        )
    return (b << 16) | (g << 8) | r


# ── Action implementations ────────────────────────────────────────────────────

def _font(
    rng,
    font_name: str | None,
    font_size: float | None,
    bold: bool | None,
    italic: bool | None,
    underline: bool | None,
    font_color: str | None,
) -> dict:
    applied = {}
    try:
        f = rng.Font
        if font_name is not None:
            f.Name = font_name
            applied["font_name"] = font_name
        if font_size is not None:
            f.Size = font_size
            applied["font_size"] = font_size
        if bold is not None:
            f.Bold = bold
            applied["bold"] = bold
        if italic is not None:
            f.Italic = italic
            applied["italic"] = italic
        if underline is not None:
            # xlUnderlineStyleSingle=2, xlUnderlineStyleNone=-4142
            f.Underline = 2 if underline else -4142
            applied["underline"] = underline
        if font_color is not None:
            f.Color = _hex_to_bgr(font_color)
            applied["font_color"] = font_color
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "Font formatting failed")
    return {"formatted": "font", "range": rng.Address, "applied": applied}


def _fill(rng, fill_color: str | None, clear_fill: bool) -> dict:
    try:
        interior = rng.Interior
        if clear_fill:
            interior.ColorIndex = -4142  # xlColorIndexNone
            return {"formatted": "fill", "range": rng.Address, "applied": {"cleared": True}}
        if fill_color is None:
            raise ToolError("fill action requires 'fill_color' (e.g. '#FFFF00') or 'clear_fill=True'.")
        interior.Color = _hex_to_bgr(fill_color)
        return {
            "formatted": "fill",
            "range": rng.Address,
            "applied": {"fill_color": fill_color},
        }
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "Fill formatting failed")


def _apply_border_side(rng, side_const: int, line_style: int, weight: int, color_bgr: int | None) -> None:
    b = rng.Borders(side_const)
    b.LineStyle = line_style
    if line_style != -4142:  # not xlLineStyleNone
        b.Weight = weight
        if color_bgr is not None:
            b.Color = color_bgr


def _border(
    rng,
    border_sides: str,
    border_style: str,
    border_weight: str,
    border_color: str | None,
) -> dict:
    line_style = _LINE_STYLE.get(border_style)
    if line_style is None:
        raise ToolError(
            f"Unknown border_style '{border_style}'. "
            "Valid: continuous, dash, double, none, thin."
        )
    weight = _BORDER_WEIGHT.get(border_weight, 2)  # default xlThin
    color_bgr = _hex_to_bgr(border_color) if border_color else None

    try:
        sides_key = border_sides.lower()
        if sides_key == "all":
            # All 4 edges + inside horizontals + verticals
            for side_const in (
                _XL_EDGE_LEFT, _XL_EDGE_TOP, _XL_EDGE_BOTTOM, _XL_EDGE_RIGHT,
                _XL_INSIDE_VERTICAL, _XL_INSIDE_HORIZONTAL,
            ):
                _apply_border_side(rng, side_const, line_style, weight, color_bgr)
        elif sides_key == "outline":
            for side_const in (
                _XL_EDGE_LEFT, _XL_EDGE_TOP, _XL_EDGE_BOTTOM, _XL_EDGE_RIGHT,
            ):
                _apply_border_side(rng, side_const, line_style, weight, color_bgr)
        elif sides_key == "inside":
            for side_const in (_XL_INSIDE_VERTICAL, _XL_INSIDE_HORIZONTAL):
                _apply_border_side(rng, side_const, line_style, weight, color_bgr)
        elif sides_key == "top":
            _apply_border_side(rng, _XL_EDGE_TOP, line_style, weight, color_bgr)
        elif sides_key == "bottom":
            _apply_border_side(rng, _XL_EDGE_BOTTOM, line_style, weight, color_bgr)
        elif sides_key == "left":
            _apply_border_side(rng, _XL_EDGE_LEFT, line_style, weight, color_bgr)
        elif sides_key == "right":
            _apply_border_side(rng, _XL_EDGE_RIGHT, line_style, weight, color_bgr)
        else:
            raise ToolError(
                f"Unknown border_sides '{border_sides}'. "
                "Valid: all, outline, inside, top, bottom, left, right."
            )
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "Border formatting failed")

    return {
        "formatted": "border",
        "range": rng.Address,
        "applied": {
            "sides": border_sides,
            "style": border_style,
            "weight": border_weight,
            "color": border_color,
        },
    }


def _number_format(rng, number_format: str | None) -> dict:
    if number_format is None:
        raise ToolError("number_format action requires the 'number_format' parameter.")
    try:
        rng.NumberFormat = number_format
    except Exception as e:
        raise _session.wrap(e, f"NumberFormat '{number_format}' failed")
    return {
        "formatted": "number_format",
        "range": rng.Address,
        "applied": {"number_format": number_format},
    }


def _alignment(
    rng,
    horizontal: str | None,
    vertical: str | None,
    wrap_text: bool | None,
    merge: bool | None,
) -> dict:
    applied = {}
    try:
        if horizontal is not None:
            h_const = _HALIGN.get(horizontal.lower())
            if h_const is None:
                raise ToolError(
                    f"Unknown horizontal alignment '{horizontal}'. "
                    "Valid: general, left, center, right."
                )
            rng.HorizontalAlignment = h_const
            applied["horizontal"] = horizontal
        if vertical is not None:
            v_const = _VALIGN.get(vertical.lower())
            if v_const is None:
                raise ToolError(
                    f"Unknown vertical alignment '{vertical}'. "
                    "Valid: top, center, bottom."
                )
            rng.VerticalAlignment = v_const
            applied["vertical"] = vertical
        if wrap_text is not None:
            rng.WrapText = wrap_text
            applied["wrap_text"] = wrap_text
        if merge is True:
            rng.Merge()
            applied["merged"] = True
        elif merge is False:
            rng.UnMerge()
            applied["merged"] = False
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "Alignment formatting failed")
    return {"formatted": "alignment", "range": rng.Address, "applied": applied}


def _column_width(rng, width: float | None) -> dict:
    if width is None:
        raise ToolError("column_width action requires the 'width' parameter (character-width units).")
    try:
        rng.ColumnWidth = width
    except Exception as e:
        raise _session.wrap(e, "ColumnWidth failed")
    return {
        "formatted": "column_width",
        "range": rng.Address,
        "applied": {"width": width},
    }


def _row_height(rng, height: float | None) -> dict:
    if height is None:
        raise ToolError("row_height action requires the 'height' parameter (points).")
    try:
        rng.RowHeight = height
    except Exception as e:
        raise _session.wrap(e, "RowHeight failed")
    return {
        "formatted": "row_height",
        "range": rng.Address,
        "applied": {"height": height},
    }


def _autofit(rng, autofit_columns: bool, autofit_rows: bool) -> dict:
    if not autofit_columns and not autofit_rows:
        raise ToolError(
            "autofit action requires at least one of autofit_columns=True or autofit_rows=True."
        )
    applied = {}
    try:
        if autofit_columns:
            rng.Columns.AutoFit()
            applied["columns"] = True
        if autofit_rows:
            rng.Rows.AutoFit()
            applied["rows"] = True
    except Exception as e:
        raise _session.wrap(e, "AutoFit failed")
    return {"formatted": "autofit", "range": rng.Address, "applied": applied}
