"""Mini in-cell sparkline charts via Range.SparklineGroups COM API.

Sparklines are tiny charts embedded in a cell (or range of cells). They
are managed through the SparklineGroups collection on the destination Range.

COM model overview
------------------
- ``ws.Range(location).SparklineGroups.Add(Type=<int>, SourceData=<str>)``
  creates a new SparklineGroup and returns it.
- The *location* range is WHERE the sparklines are rendered (destination).
- The *SourceData* string is the DATA range — must include sheet name if
  it differs from the destination sheet, e.g. "Sheet1!B2:E10".
- For a column of per-row sparklines, location "F2:F5" + SourceData
  "B2:E5" maps rows 1-to-1: F2 gets data from B2:E2, etc.

XlSparkType enum (confirmed from Microsoft Learn)
-------------------------------------------------
- xlSparkLine              = 1   (line sparkline)
- xlSparkColumn            = 2   (column/bar sparkline)
- xlSparkColumnStacked100  = 3   (win/loss sparkline)

CRITICAL: Application.ReferenceStyle MUST be xlA1 (= 1) before calling
SparklineGroups.Add, otherwise Excel may raise a COM error. We set it and
restore it around Add() calls.

Color convention
----------------
color parameter accepts ``"#RRGGBB"`` (e.g. ``"#FF0000"`` for red).
Excel's SeriesColor.Color expects a BGR integer (same as Interior.Color /
Font.Color). Conversion via _hex_to_bgr() replicated verbatim from format.py.

VERIFY-EFFECT contract
----------------------
- add:   after Add(), assert SparklineGroups.Count >= 1 via Count property.
- clear: after Clear(), assert SparklineGroups.Count == 0.
- list:  read-only; returns Count + per-group Type/SourceData/Location.

Note on SparklineGroups enumeration: the COM SparklineGroups collection
supports .Count and .Item(i) (1-based index). Iterating via Python for-in
over a win32com dispatch wrapper usually works, but .Item(i) is more
reliable. On error, fall back to Count-only.
"""

from __future__ import annotations

from fastmcp.exceptions import ToolError

from ..session import ExcelSession, excel_guard

_session = ExcelSession()

# XlSparkType — confirmed from Microsoft Learn:
# https://learn.microsoft.com/en-us/office/vba/api/excel.xlsparktype
_XL_SPARK_LINE                = 1   # xlSparkLine
_XL_SPARK_COLUMN              = 2   # xlSparkColumn
_XL_SPARK_COLUMN_STACKED100   = 3   # xlSparkColumnStacked100 (win/loss)

# XlReferenceStyle
_XL_A1 = 1  # xlA1 — required by SparklineGroups.Add

_SPARK_TYPE_MAP = {
    "line":      _XL_SPARK_LINE,
    "column":    _XL_SPARK_COLUMN,
    "win_loss":  _XL_SPARK_COLUMN_STACKED100,
}


# ── Public entry point ─────────────────────────────────────────────────────────

def sparkline_action(
    action: str,
    location: str,
    sheet: str | None = None,
    workbook: str | None = None,
    # add
    data_range: str | None = None,
    spark_type: str = "line",
    marker: bool | None = None,
    color: str | None = None,
) -> dict:
    """Dispatch a sparkline action.

    Actions
    -------
    add
        Add sparklines at *location* (destination cells) using data from
        *data_range*. spark_type is "line" (default), "column", or "win_loss".
        Optional: marker=True to show data-point markers (line type only),
        color="#RRGGBB" to set the series color.
    clear
        Remove all sparklines from *location*.
    list
        Report sparkline groups on the given sheet (scoped to *location* if
        provided, otherwise the entire sheet's used range).
    """
    valid = {"add", "clear", "list"}
    if action not in valid:
        raise ToolError(
            f"Unknown action '{action}'. Valid: {', '.join(sorted(valid))}."
        )
    return _session.run_com(
        _dispatch,
        action, location, sheet, workbook,
        data_range, spark_type, marker, color,
    )


# ── Worker-thread dispatcher ───────────────────────────────────────────────────

def _dispatch(
    action: str,
    location: str,
    sheet: str | None,
    workbook: str | None,
    data_range: str | None,
    spark_type: str,
    marker: bool | None,
    color: str | None,
) -> dict:
    """Executed on the STA COM worker thread."""
    rng = _resolve_range(location, sheet, workbook)
    app = rng.Application

    with excel_guard(app):
        if action == "add":
            return _add(rng, data_range, spark_type, marker, color, app)
        if action == "clear":
            return _clear(rng)
        # action == "list"
        return _list(rng)


# ── Range resolver (verbatim from validation.py) ──────────────────────────────

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


# ── Color helper (verbatim from format.py) ────────────────────────────────────

def _hex_to_bgr(hex_color: str) -> int:
    """Convert ``"#RRGGBB"`` to Excel BGR integer (0xBBGGRR).

    Excel SeriesColor.Color expects BGR byte order, not RGB.
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

def _add(rng, data_range: str | None, spark_type: str, marker: bool | None, color: str | None, app) -> dict:
    """Add sparklines at the location range from data_range source.

    CRITICAL: Application.ReferenceStyle must be xlA1 (=1) before Add().
    We save + restore the current style around the call.
    """
    if not data_range:
        raise ToolError(
            "add action requires 'data_range' (e.g. 'B2:E10' or 'Sheet1!B2:E10')."
        )
    type_int = _SPARK_TYPE_MAP.get(spark_type)
    if type_int is None:
        raise ToolError(
            f"Unknown spark_type '{spark_type}'. Valid: line, column, win_loss."
        )
    color_bgr: int | None = None
    if color is not None:
        color_bgr = _hex_to_bgr(color)

    try:
        # Save ReferenceStyle and force xlA1 — required by SparklineGroups.Add
        saved_ref_style = app.ReferenceStyle
        app.ReferenceStyle = _XL_A1
        try:
            group = rng.SparklineGroups.Add(Type=type_int, SourceData=data_range)
        finally:
            app.ReferenceStyle = saved_ref_style

        # Optional post-Add styling
        if marker is not None:
            group.Points.Markers.Visible = bool(marker)
        if color_bgr is not None:
            group.SeriesColor.Color = color_bgr

        # VERIFY-EFFECT: confirm at least one group now exists at this location
        count = rng.SparklineGroups.Count
        if count < 1:
            raise ToolError(
                "SparklineGroups.Add reported success but Count is still 0 — "
                "verify that ReferenceStyle is xlA1 and data_range is valid."
            )
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "sparkline add failed")

    applied: dict = {
        "location": rng.Address,
        "data_range": data_range,
        "spark_type": spark_type,
        "type_int": type_int,
        "groups_count": count,
    }
    if marker is not None:
        applied["marker"] = marker
    if color is not None:
        applied["color"] = color

    return {
        "sparkline": "add",
        "sheet": rng.Worksheet.Name,
        "applied": applied,
    }


def _clear(rng) -> dict:
    """Remove all sparklines from the location range.

    Uses SparklineGroups.Clear() which removes sparklines while preserving
    cell values. After Clear(), Count must be 0 (verify-effect).
    """
    try:
        rng.SparklineGroups.Clear()
        count_after = rng.SparklineGroups.Count
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "sparkline clear failed")

    # VERIFY-EFFECT: Clear() must actually remove all groups from this range.
    if count_after != 0:
        raise ToolError(
            f"SparklineGroups.Clear reported success but Count is still "
            f"{count_after} — sparklines may not have been removed."
        )

    return {
        "sparkline": "clear",
        "sheet": rng.Worksheet.Name,
        "applied": {
            "location": rng.Address,
            "cleared": True,
            "groups_count": count_after,
        },
    }


def _list(rng) -> dict:
    """Report sparkline groups on the range's sheet (best-effort enumeration).

    Enumerates via SparklineGroups.Item(i) (1-based). Each group exposes
    .Type (int), .SourceData (str), and .Location.Address (str).
    If enumeration fails for a group, that entry records the error string.
    """
    ws = rng.Worksheet
    try:
        # Scope to the caller's location range — its SparklineGroups covers
        # sparklines anchored in those cells. UsedRange is NOT a reliable scope:
        # sparkline destination cells are often value-empty and excluded from
        # UsedRange, so UsedRange.SparklineGroups misses them entirely.
        groups_col = rng.SparklineGroups
        count = groups_col.Count
    except Exception as e:
        raise _session.wrap(e, "sparkline list failed — could not read SparklineGroups")

    groups_info = []
    for i in range(1, count + 1):
        try:
            grp = groups_col.Item(i)
            type_int = grp.Type
            # Reverse-map type int to name for readability
            type_name = {v: k for k, v in _SPARK_TYPE_MAP.items()}.get(type_int, str(type_int))
            try:
                source = grp.SourceData
            except Exception:
                source = "<unavailable>"
            try:
                loc = grp.Location.Address
            except Exception:
                loc = "<unavailable>"
            groups_info.append({
                "index": i,
                "type": type_name,
                "type_int": type_int,
                "source_data": source,
                "location": loc,
            })
        except Exception as ex:
            groups_info.append({"index": i, "error": str(ex)})

    return {
        "sparkline": "list",
        "sheet": ws.Name,
        "applied": {
            "groups_count": count,
            "groups": groups_info,
        },
    }
