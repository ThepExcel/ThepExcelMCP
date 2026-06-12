"""Excel chart (ChartObject/Chart) operations — create, configure, export, delete."""

from __future__ import annotations

import os
import tempfile
import time

from fastmcp.exceptions import ToolError

from ..session import ExcelSession

_session = ExcelSession()

# XlChartType constants — verified against Microsoft Learn (xl enum)
# https://learn.microsoft.com/en-us/office/vba/api/excel.xlcharttype
_CHART_TYPE: dict[str, int] = {
    "column":              51,   # xlColumnClustered
    "column_stacked":      52,   # xlColumnStacked
    "column_stacked100":   53,   # xlColumnStacked100
    "bar":                 57,   # xlBarClustered
    "bar_stacked":         58,   # xlBarStacked
    "bar_stacked100":      59,   # xlBarStacked100
    "line":                 4,   # xlLine
    "line_stacked":        63,   # xlLineStacked
    "line_markers":        65,   # xlLineMarkers
    "area":                 1,   # xlArea
    "area_stacked":        76,   # xlAreaStacked
    "pie":                  5,   # xlPie
    "doughnut":         -4120,   # xlDoughnut
    "scatter":          -4169,   # xlXYScatter
    "scatter_lines":       74,   # xlXYScatterLines
    "bubble":              15,   # xlBubble
    "radar":            -4151,   # xlRadar
    "combo_column_line":  113,   # xlComboColumnClusteredLine
    "combo_column_line_2axis": 114,  # xlComboColumnClusteredLineSecondaryAxis
    "waterfall":          119,   # xlWaterfall
    "treemap":            117,   # xlTreemap
    "sunburst":           120,   # xlSunburst
    "histogram":          118,   # xlHistogram
    "funnel":             123,   # xlFunnel
}

# XlLegendPosition constants
_LEGEND_POSITION: dict[str, int] = {
    "bottom":  -4107,   # xlLegendPositionBottom
    "corner":      2,   # xlLegendPositionCorner
    "left":   -4131,   # xlLegendPositionLeft
    "right":  -4152,   # xlLegendPositionRight
    "top":    -4160,   # xlLegendPositionTop
}

# Default chart dimensions (points)
_DEFAULT_WIDTH  = 375
_DEFAULT_HEIGHT = 225


def chart_action(
    action: str,
    # identity
    name: str | None = None,
    workbook: str | None = None,
    sheet: str | None = None,
    # create
    source: str | None = None,
    chart_type: str | None = None,
    position: str | None = None,
    width: float | None = None,
    height: float | None = None,
    title: str | None = None,
    dest_sheet: str | None = None,
    # configure
    x_title: str | None = None,
    y_title: str | None = None,
    legend: bool | None = None,
    legend_position: str | None = None,
    data_labels: bool | None = None,
    series_names: list | None = None,
    secondary_series: str | None = None,
    # export_image
    output_path: str | None = None,
) -> dict:
    """Dispatch an Excel Chart action.

    Actions
    -------
    list
        All ChartObjects across all sheets: name, sheet, chart_type,
        source_ranges, position (left, top, width, height).
        Example: ``excel_chart(action="list")``
    create
        Create an embedded chart. Requires ``source`` (range address like
        ``"A1:C10"`` or a table/pivot table name) and ``chart_type``
        (friendly name from the list below).
        Optional: ``sheet`` (default active), ``position`` (anchor cell like
        ``"E2"``), ``width`` (points, default 375), ``height`` (points, default
        225), ``title``, ``dest_sheet`` (sheet to embed on; default same as
        source sheet).

        Supported chart_type values:
        column, column_stacked, column_stacked100,
        bar, bar_stacked, bar_stacked100,
        line, line_stacked, line_markers,
        area, area_stacked,
        pie, doughnut, scatter, scatter_lines, bubble, radar,
        combo_column_line, combo_column_line_2axis,
        waterfall, treemap, sunburst, histogram, funnel

        PivotCharts: pass a pivot table name as ``source``.
        The chart will be created using the pivot table's range as data source.
        Note: full pivot-chart binding (field sync on refresh) is not
        supported via this COM path — use it for static snapshot charts
        from pivot ranges only.
        Example: ``excel_chart(action="create", source="A1:C10",
        chart_type="column", title="Sales by Region")``
    configure
        Adjust chart appearance. Requires ``name``.
        Optional: ``title``, ``x_title`` (category axis), ``y_title`` (value
        axis), ``legend`` (bool show/hide), ``legend_position``
        (bottom/corner/left/right/top), ``data_labels`` (bool on/off for all
        series), ``series_names`` (list of new series name strings in order),
        ``secondary_series`` (series name to move to secondary Y axis).
        Example: ``excel_chart(action="configure", name="Chart 1",
        title="Q1 Sales", legend=True, legend_position="bottom")``
    set_source
        Change the chart's data range. Requires ``name`` and ``source``
        (range address).
        Example: ``excel_chart(action="set_source", name="Chart 1",
        source="A1:D20")``
    export_image
        Export chart to PNG. Requires ``name``.
        ``output_path`` defaults to ``%TEMP%/thepexcel_mcp/<name>.png``.
        Returns ``{"path": "<abs_path>", "chart": "<name>"}``.
        Example: ``excel_chart(action="export_image", name="Chart 1")``
    delete
        Delete a chart by name. Requires ``name``.
        Example: ``excel_chart(action="delete", name="Chart 1")``
    """
    if action == "list":
        return _session.run_com(_list, workbook)
    if action == "create":
        _require(source, "source", action)
        _require(chart_type, "chart_type", action)
        if chart_type not in _CHART_TYPE:
            raise ToolError(
                f"Unknown chart_type '{chart_type}'. "
                f"Valid: {sorted(_CHART_TYPE.keys())}"
            )
        return _session.run_com(
            _create, source, chart_type, sheet, workbook,
            position, width, height, title, dest_sheet,
        )
    if action == "configure":
        _require(name, "name", action)
        return _session.run_com(
            _configure, name, workbook,
            title, x_title, y_title,
            legend, legend_position,
            data_labels, series_names, secondary_series,
        )
    if action == "set_source":
        _require(name, "name", action)
        _require(source, "source", action)
        return _session.run_com(_set_source, name, workbook, source)
    if action == "export_image":
        _require(name, "name", action)
        return _session.run_com(_export_image, name, workbook, output_path)
    if action == "delete":
        _require(name, "name", action)
        return _session.run_com(_delete, name, workbook)
    raise ToolError(
        f"Unknown action '{action}'. "
        "Valid: list, create, configure, set_source, export_image, delete."
    )


# ── Helpers ────────────────────────────────────────────────────────────────────

def _require(value, param: str, action: str) -> None:
    if value is None:
        raise ToolError(f"action='{action}' requires '{param}'.")


def _find_chart(wb, name: str):
    """Return ChartObject COM object or raise ToolError with available names."""
    available = []
    for i in range(1, wb.Sheets.Count + 1):
        ws = wb.Sheets(i)
        try:
            cos = ws.ChartObjects()
        except Exception:
            continue
        for j in range(1, cos.Count + 1):
            co = cos(j)
            if co.Name == name:
                return co
            available.append(co.Name)
    raise ToolError(
        f"Chart '{name}' not found. "
        f"Available charts: {available or ['(none)']}. "
        "Use excel_chart(action='list') to see all charts."
    )


def _chart_info(co) -> dict:
    """Compact info dict for a ChartObject."""
    chart = co.Chart
    try:
        chart_type_val = chart.ChartType
        # Reverse-lookup friendly name
        type_name = next(
            (k for k, v in _CHART_TYPE.items() if v == chart_type_val),
            str(chart_type_val),
        )
    except Exception:
        type_name = "unknown"
    try:
        src = chart.SeriesCollection(1).Formula
    except Exception:
        src = ""
    return {
        "name": co.Name,
        "sheet": co.Parent.Name,
        "chart_type": type_name,
        "source": src,
        "left": co.Left,
        "top": co.Top,
        "width": co.Width,
        "height": co.Height,
    }


def _default_output_path(name: str) -> str:
    tmp = os.path.join(tempfile.gettempdir(), "thepexcel_mcp")
    os.makedirs(tmp, exist_ok=True)
    safe = "".join(c if c.isalnum() or c in "-_." else "_" for c in name)
    return os.path.join(tmp, f"{safe}.png")


# ── Action implementations ─────────────────────────────────────────────────────

def _list(workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    charts = []
    for i in range(1, wb.Sheets.Count + 1):
        ws = wb.Sheets(i)
        try:
            cos = ws.ChartObjects()
        except Exception:
            continue
        for j in range(1, cos.Count + 1):
            charts.append(_chart_info(cos(j)))
    return {"charts": charts, "count": len(charts)}


def _delete(name: str, workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    co = _find_chart(wb, name)
    sheet = co.Parent.Name
    try:
        co.Delete()
        return {"deleted": name, "sheet": sheet}
    except Exception as e:
        raise _session.wrap(e, f"Delete chart '{name}' failed")


def _create(
    source: str,
    chart_type: str,
    sheet: str | None,
    workbook: str | None,
    position: str | None,
    width: float | None,
    height: float | None,
    title: str | None,
    dest_sheet: str | None,
) -> dict:
    wb = _session.get_workbook(workbook)
    type_val = _CHART_TYPE[chart_type]
    w = width or _DEFAULT_WIDTH
    h = height or _DEFAULT_HEIGHT

    # Resolve source range or named object (table / pivot)
    source_ws = _session.get_sheet(sheet, workbook)
    try:
        src_rng = source_ws.Range(source)
    except Exception:
        # Not a range address — try as table name
        src_rng = None
        for i in range(1, wb.Sheets.Count + 1):
            s = wb.Sheets(i)
            for j in range(1, s.ListObjects.Count + 1):
                lo = s.ListObjects(j)
                if lo.Name == source:
                    src_rng = lo.Range
                    source_ws = s
                    break
            if src_rng is not None:
                break
        if src_rng is None:
            # Try as pivot table name
            for i in range(1, wb.Sheets.Count + 1):
                s = wb.Sheets(i)
                try:
                    pt_count = s.PivotTables().Count
                except Exception:
                    continue
                for j in range(1, pt_count + 1):
                    pt = s.PivotTables(j)
                    if pt.Name == source:
                        src_rng = pt.TableRange1
                        source_ws = s
                        break
                if src_rng is not None:
                    break
        if src_rng is None:
            raise ToolError(
                f"Source '{source}' is not a valid range address, table name, "
                "or pivot table name in this workbook."
            )

    dest_ws = source_ws
    if dest_sheet:
        dest_ws = _session.get_sheet(dest_sheet, workbook)

    # Determine anchor position
    left, top = 10.0, 10.0
    if position:
        try:
            anchor = dest_ws.Range(position)
            left = anchor.Left
            top = anchor.Top
        except Exception:
            pass  # Fall back to default position

    try:
        cos = dest_ws.ChartObjects()
        co = cos.Add(left, top, w, h)
        chart = co.Chart
        chart.SetSourceData(src_rng)
        chart.ChartType = type_val
        if title:
            chart.HasTitle = True
            chart.ChartTitle.Text = title
        return _chart_info(co)
    except Exception as e:
        raise _session.wrap(e, f"Create chart from '{source}' failed")


def _configure(
    name: str,
    workbook: str | None,
    title: str | None,
    x_title: str | None,
    y_title: str | None,
    legend: bool | None,
    legend_position: str | None,
    data_labels: bool | None,
    series_names: list | None,
    secondary_series: str | None,
) -> dict:
    wb = _session.get_workbook(workbook)
    co = _find_chart(wb, name)
    chart = co.Chart
    changes = []

    try:
        if title is not None:
            chart.HasTitle = bool(title)
            if title:
                chart.ChartTitle.Text = title
            changes.append("title")

        if x_title is not None:
            # xlCategory = 1
            ax = chart.Axes(1)
            ax.HasTitle = bool(x_title)
            if x_title:
                ax.AxisTitle.Text = x_title
            changes.append("x_title")

        if y_title is not None:
            # xlValue = 2
            ax = chart.Axes(2)
            ax.HasTitle = bool(y_title)
            if y_title:
                ax.AxisTitle.Text = y_title
            changes.append("y_title")

        if legend is not None:
            chart.HasLegend = legend
            changes.append("legend")

        if legend_position is not None and legend:
            if legend_position not in _LEGEND_POSITION:
                raise ToolError(
                    f"legend_position='{legend_position}' invalid. "
                    f"Valid: {list(_LEGEND_POSITION.keys())}"
                )
            chart.Legend.Position = _LEGEND_POSITION[legend_position]
            changes.append("legend_position")

        if data_labels is not None:
            for i in range(1, chart.SeriesCollection().Count + 1):
                chart.SeriesCollection(i).HasDataLabels = data_labels
            changes.append("data_labels")

        if series_names:
            for i, sname in enumerate(series_names, start=1):
                if i <= chart.SeriesCollection().Count:
                    chart.SeriesCollection(i).Name = sname
            changes.append("series_names")

        if secondary_series is not None:
            # xlSecondary = 2
            for i in range(1, chart.SeriesCollection().Count + 1):
                s = chart.SeriesCollection(i)
                if s.Name == secondary_series:
                    s.AxisGroup = 2
                    changes.append("secondary_axis")
                    break
            else:
                raise ToolError(
                    f"Series '{secondary_series}' not found in chart '{name}'."
                )

        return {"chart": name, "changes": changes}
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"Configure chart '{name}' failed")


def _set_source(name: str, workbook: str | None, source: str) -> dict:
    wb = _session.get_workbook(workbook)
    co = _find_chart(wb, name)
    chart = co.Chart
    # Resolve range — must be sheet-qualified or a range within any sheet
    ws = co.Parent  # chart's own sheet
    try:
        rng = ws.Range(source)
    except Exception:
        # Try as workbook-level range
        try:
            app = _session.get_app()
            rng = app.Range(source)
        except Exception as e:
            raise _session.wrap(e, f"Cannot resolve source range '{source}'")
    try:
        chart.SetSourceData(rng)
        return {"chart": name, "new_source": source}
    except Exception as e:
        raise _session.wrap(e, f"Set source on '{name}' failed")


def _export_image(
    name: str,
    workbook: str | None,
    output_path: str | None,
) -> dict:
    wb = _session.get_workbook(workbook)
    co = _find_chart(wb, name)
    path = output_path or _default_output_path(name)
    path = os.path.abspath(path)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    try:
        co.Chart.Export(path, "PNG")
    except Exception as e:
        raise _session.wrap(e, f"Export chart '{name}' to PNG failed")
    if not os.path.exists(path):
        raise ToolError(
            f"Chart.Export completed but file not found at '{path}'. "
            "Check that the path is writable and the chart is not empty."
        )
    return {"path": path, "chart": name}
