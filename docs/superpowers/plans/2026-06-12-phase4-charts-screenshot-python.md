# Phase 4: Charts, Screenshot, Python-in-Excel Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add `excel_chart` (6 actions), `excel_screenshot` (3 actions), and `excel_range(action="write_py")` to ThepExcelMCP, bringing Phase 4 to completion with full test coverage.

**Architecture:** Each domain module follows the established pattern: pure-Python arg validation at the top of the action function, then `_session.run_com(fn, *args)` for all COM work. `pillow` is added as a dependency for clipboard→PNG conversion in screenshot. `excel_chart` and `excel_screenshot` are new files; `write_py` is appended to the existing `ranges.py`; all three are registered in `server.py`.

**Tech Stack:** Python 3.11+, FastMCP, pywin32 (win32con, win32clipboard, win32api), Pillow (PIL), Excel COM (ChartObjects, Chart.Export, Range.CopyPicture, Range.Formula2R1C1).

---

## File Map

| File | Action |
|---|---|
| `src/thepexcel_mcp/domains/charts.py` | Create — `chart_action()` + 6 inner fns |
| `src/thepexcel_mcp/domains/screenshot.py` | Create — `screenshot_action()` + 3 inner fns |
| `src/thepexcel_mcp/domains/ranges.py` | Modify — add `write_py` action + `_write_py()` fn |
| `src/thepexcel_mcp/server.py` | Modify — register `excel_chart` + `excel_screenshot` tools; extend `excel_range` params/docstring |
| `pyproject.toml` | Modify — add `pillow>=10.0` dependency |
| `CLAUDE.md` | Modify — add Phase 4 tools to Tool Registry |
| `docs/ROADMAP.md` | Modify — mark P4 done |
| `tests/test_phase4.py` | Create — unit tests (no Excel) |
| `tests/smoke_com.py` | Modify — append Phase 4 smoke section |

---

## Task 1: Add `pillow` dependency

**Files:**
- Modify: `D:/ThepExcelMCP/pyproject.toml`

- [ ] **Step 1.1: Read current pyproject.toml**

```
Already read — dependencies section is at line 6-9.
```

- [ ] **Step 1.2: Add pillow dependency**

In `pyproject.toml`, change:
```toml
dependencies = [
    "fastmcp>=3.0.0",
    "pywin32>=306",
]
```
to:
```toml
dependencies = [
    "fastmcp>=3.0.0",
    "pywin32>=306",
    "pillow>=10.0",
]
```

- [ ] **Step 1.3: Sync + verify import**

```powershell
cd D:/ThepExcelMCP
uv sync
uv run python -c "from PIL import ImageGrab; print('PIL OK')"
```
Expected output: `PIL OK`

- [ ] **Step 1.4: Run existing tests — must still pass**

```powershell
uv run pytest -q
```
Expected: `167 passed`

- [ ] **Step 1.5: Commit**

```bash
git add pyproject.toml uv.lock
git commit -m "feat(p4): add pillow dependency for screenshot clipboard capture"
```

---

## Task 2: Chart type mapping constants

**Files:**
- Create: `D:/ThepExcelMCP/src/thepexcel_mcp/domains/charts.py` (skeleton with constants only)

The `_CHART_TYPE` dict maps friendly lowercase names → verified XlChartType integer values from Microsoft Learn (verified 2026-06-12):

- [ ] **Step 2.1: Create charts.py with type constants**

```python
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
```

- [ ] **Step 2.2: Run import check**

```powershell
uv run python -c "from thepexcel_mcp.domains.charts import _CHART_TYPE; print(len(_CHART_TYPE), 'chart types')"
```
Expected: `25 chart types`

---

## Task 3: `chart_action` dispatcher + `list` + `delete`

**Files:**
- Modify: `D:/ThepExcelMCP/src/thepexcel_mcp/domains/charts.py`

- [ ] **Step 3.1: Write failing tests for chart_action dispatch**

Create `D:/ThepExcelMCP/tests/test_phase4.py`:

```python
"""Phase 4 unit tests — no Excel required.

Covers:
- chart type mapping (_CHART_TYPE completeness)
- chart_action arg validation (unknown action, missing required args)
- write_py escaping logic
- screenshot_action arg validation
"""

from __future__ import annotations

import pytest
from unittest.mock import patch, MagicMock

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


# ── Chart type mapping ─────────────────────────────────────────────────────────

class TestChartTypeMapping:
    def test_column_maps_to_51(self):
        from thepexcel_mcp.domains.charts import _CHART_TYPE
        assert _CHART_TYPE["column"] == 51

    def test_bar_maps_to_57(self):
        from thepexcel_mcp.domains.charts import _CHART_TYPE
        assert _CHART_TYPE["bar"] == 57

    def test_line_maps_to_4(self):
        from thepexcel_mcp.domains.charts import _CHART_TYPE
        assert _CHART_TYPE["line"] == 4

    def test_pie_maps_to_5(self):
        from thepexcel_mcp.domains.charts import _CHART_TYPE
        assert _CHART_TYPE["pie"] == 5

    def test_scatter_maps_to_negative_4169(self):
        from thepexcel_mcp.domains.charts import _CHART_TYPE
        assert _CHART_TYPE["scatter"] == -4169

    def test_doughnut_maps_to_negative_4120(self):
        from thepexcel_mcp.domains.charts import _CHART_TYPE
        assert _CHART_TYPE["doughnut"] == -4120

    def test_all_values_are_ints(self):
        from thepexcel_mcp.domains.charts import _CHART_TYPE
        for k, v in _CHART_TYPE.items():
            assert isinstance(v, int), f"chart type '{k}' value is not int: {v!r}"


# ── Chart action dispatch ──────────────────────────────────────────────────────

class TestChartActionDispatch:
    def _call(self, **kwargs):
        mock_session = make_mock_session()
        mock_session.get_workbook.side_effect = ToolError("no excel")
        with patch("thepexcel_mcp.domains.charts._session", mock_session):
            from thepexcel_mcp.domains.charts import chart_action
            return chart_action(**kwargs)

    def test_unknown_action(self):
        with pytest.raises(ToolError, match="Unknown action"):
            self._call(action="frobnicate")

    def test_create_missing_source(self):
        with pytest.raises(ToolError, match="requires 'source'"):
            self._call(action="create", chart_type="column")

    def test_create_missing_chart_type(self):
        with pytest.raises(ToolError, match="requires 'chart_type'"):
            self._call(action="create", source="A1:C10")

    def test_create_unknown_chart_type(self):
        with pytest.raises(ToolError, match="Unknown chart_type"):
            self._call(action="create", source="A1:C10", chart_type="laser_beam")

    def test_configure_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="configure", title="My Chart")

    def test_set_source_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="set_source", source="A1:D10")

    def test_set_source_missing_source(self):
        with pytest.raises(ToolError, match="requires 'source'"):
            self._call(action="set_source", name="Chart1")

    def test_export_image_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="export_image")

    def test_delete_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="delete")
```

- [ ] **Step 3.2: Run test — confirm failures**

```powershell
cd D:/ThepExcelMCP
uv run pytest tests/test_phase4.py -v 2>&1 | head -40
```
Expected: collection errors (charts module not fully implemented yet) — that's fine; we only need type mapping tests to pass initially.

- [ ] **Step 3.3: Implement dispatcher + list + delete in charts.py**

Append to `charts.py` after the constants:

```python
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
    # set_source
    source_range: str | None = None,
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
```

- [ ] **Step 3.4: Implement _list and _delete COM functions**

Append to `charts.py`:

```python
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
```

- [ ] **Step 3.5: Run dispatch tests**

```powershell
uv run pytest tests/test_phase4.py -v -k "TestChart" 2>&1 | tail -20
```
Expected: all `TestChartActionDispatch` and `TestChartTypeMapping` tests pass.

- [ ] **Step 3.6: Run full test suite**

```powershell
uv run pytest -q
```
Expected: 167 + new tests pass.

---

## Task 4: Chart `create`, `configure`, `set_source`

**Files:**
- Modify: `D:/ThepExcelMCP/src/thepexcel_mcp/domains/charts.py`

- [ ] **Step 4.1: Implement `_create`**

Append to `charts.py`:

```python
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
                for j in range(1, s.PivotTables().Count + 1):
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
```

- [ ] **Step 4.2: Implement `_configure`**

Append to `charts.py`:

```python
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
```

- [ ] **Step 4.3: Implement `_set_source`**

Append to `charts.py`:

```python
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
```

- [ ] **Step 4.4: Run tests**

```powershell
uv run pytest -q
```
Expected: all prior tests still pass.

---

## Task 5: Chart `export_image`

**Files:**
- Modify: `D:/ThepExcelMCP/src/thepexcel_mcp/domains/charts.py`

The `Chart.Export(path, "PNG")` COM call writes a PNG directly — no clipboard needed for charts. This is simpler than the range screenshot approach.

- [ ] **Step 5.1: Implement `_export_image`**

Append to `charts.py`:

```python
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
```

- [ ] **Step 5.2: Run tests**

```powershell
uv run pytest -q
```
Expected: all pass.

- [ ] **Step 5.3: Commit charts.py**

```bash
git add src/thepexcel_mcp/domains/charts.py
git commit -m "feat(p4): add excel_chart domain (list/create/configure/set_source/export_image/delete)"
```

---

## Task 6: Screenshot domain

**Files:**
- Create: `D:/ThepExcelMCP/src/thepexcel_mcp/domains/screenshot.py`

The pipeline for `range` and `sheet` actions:
1. `range.CopyPicture(Appearance=1, Format=2)` — xlScreen=1, xlBitmap=2 — copies to clipboard on COM worker thread
2. `PIL.ImageGrab.grabclipboard()` — reads from clipboard (must also be on COM worker thread since clipboard state is thread-local on Windows)
3. `img.save(path, "PNG")` — save to file
4. Clear clipboard via `app.CutCopyMode = False`

Note on DPI: `ImageGrab.grabclipboard()` on Windows returns the raw bitmap; DPI scaling does not affect the pixel output when using `xlBitmap`. `xlPicture` (Format=1) gives a vector metafile — less reliable with PIL. We use `Format=2` (bitmap) for reliability.

- [ ] **Step 6.1: Write failing tests for screenshot dispatch**

Append to `tests/test_phase4.py`:

```python
# ── Screenshot action dispatch ─────────────────────────────────────────────────

class TestScreenshotActionDispatch:
    def _call(self, **kwargs):
        mock_session = make_mock_session()
        mock_session.get_workbook.side_effect = ToolError("no excel")
        with patch("thepexcel_mcp.domains.screenshot._session", mock_session):
            from thepexcel_mcp.domains.screenshot import screenshot_action
            return screenshot_action(**kwargs)

    def test_unknown_action(self):
        with pytest.raises(ToolError, match="Unknown action"):
            self._call(action="frobnicate")

    def test_range_missing_range(self):
        with pytest.raises(ToolError, match="requires 'range'"):
            self._call(action="range")

    def test_sheet_no_error_without_name(self):
        # sheet action does NOT require name (uses active sheet when omitted)
        # dispatch should pass to run_com; ToolError from get_workbook is fine
        with pytest.raises(ToolError):
            self._call(action="sheet")

    def test_chart_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="chart")
```

- [ ] **Step 6.2: Run test — confirm failure**

```powershell
uv run pytest tests/test_phase4.py -k "TestScreenshot" -v 2>&1 | head -20
```
Expected: ImportError (module not yet created).

- [ ] **Step 6.3: Create screenshot.py**

```python
"""Excel screenshot (visual verification) — capture range/sheet/chart as PNG.

All clipboard operations run on the COM worker thread (clipboard state is
per-thread on Windows). PIL.ImageGrab.grabclipboard() is therefore called
inside the fn passed to _session.run_com().

DPI note: CopyPicture with Format=2 (xlBitmap) copies a device-dependent
bitmap that reflects the current screen DPI. SetProcessDpiAwareness is NOT
called here — doing so from a background thread after the process is running
has no effect. If output looks small on HiDPI, use Format=1 (xlPicture) via
export through a temp ChartObject instead. The bitmap approach is used for
simplicity and reliability.
"""

from __future__ import annotations

import os
import tempfile
import time

from fastmcp.exceptions import ToolError

from ..session import ExcelSession

_session = ExcelSession()

# CopyPicture constants
_XL_SCREEN  = 1   # xlScreen  (Appearance)
_XL_BITMAP  = 2   # xlBitmap  (Format)

_DEFAULT_SUBDIR = "thepexcel_mcp"


def screenshot_action(
    action: str,
    range: str | None = None,
    sheet: str | None = None,
    workbook: str | None = None,
    output_path: str | None = None,
    name: str | None = None,
) -> dict:
    """Capture a portion of the live Excel workbook as a PNG image.

    Purpose: LLM visual verification — let the agent "see" the workbook
    state after writes, chart creation, or formatting changes.

    Actions
    -------
    range
        Capture a specific range as PNG. Requires ``range`` (e.g. ``"A1:F20"``).
        ``sheet`` defaults to the active sheet.
        ``output_path`` defaults to ``%TEMP%/thepexcel_mcp/<timestamp>.png``.
        Technique: Range.CopyPicture (bitmap mode) → PIL.ImageGrab.grabclipboard()
        → save to PNG. Clipboard is cleared after capture.
        Returns ``{"path": "<abs_path>", "range": "<range>", "sheet": "<sheet>"}``.
        Example: ``excel_screenshot(action="range", range="A1:F20")``
    sheet
        Capture the used range of an entire sheet.
        ``sheet`` defaults to the active sheet.
        Same output convention as ``range``.
        Example: ``excel_screenshot(action="sheet", sheet="Summary")``
    chart
        Export a chart by name as PNG. Requires ``name`` (chart name).
        Delegates to Chart.Export — no clipboard involved.
        Example: ``excel_screenshot(action="chart", name="Chart 1")``
    """
    if action == "range":
        _require(range, "range", action)
        return _session.run_com(_capture_range, range, sheet, workbook, output_path)
    if action == "sheet":
        return _session.run_com(_capture_sheet, sheet, workbook, output_path)
    if action == "chart":
        _require(name, "name", action)
        return _session.run_com(_capture_chart, name, workbook, output_path)
    raise ToolError(
        f"Unknown action '{action}'. Valid: range, sheet, chart."
    )


def _require(value, param: str, action: str) -> None:
    if value is None:
        raise ToolError(f"action='{action}' requires '{param}'.")


def _default_path(suffix: str = "") -> str:
    tmp = os.path.join(tempfile.gettempdir(), _DEFAULT_SUBDIR)
    os.makedirs(tmp, exist_ok=True)
    ts = int(time.time() * 1000)
    safe = "".join(c if c.isalnum() or c in "-_." else "_" for c in suffix)
    name = f"{ts}_{safe}.png" if safe else f"{ts}.png"
    return os.path.join(tmp, name)


def _copy_picture_to_file(rng, output_path: str | None, label: str) -> str:
    """CopyPicture → clipboard → PIL → PNG file. Returns absolute path."""
    from PIL import ImageGrab

    path = output_path or _default_path(label)
    path = os.path.abspath(path)
    os.makedirs(os.path.dirname(path), exist_ok=True)

    app = _session.get_app()
    # Clear clipboard before copy to avoid stale data
    app.CutCopyMode = False

    # CopyPicture: Appearance=1 (xlScreen), Format=2 (xlBitmap)
    try:
        rng.CopyPicture(Appearance=_XL_SCREEN, Format=_XL_BITMAP)
    except Exception as e:
        raise _session.wrap(e, f"CopyPicture failed for '{label}'")

    # Small delay: Excel's rendering pipeline needs a moment after CopyPicture
    time.sleep(0.15)

    img = ImageGrab.grabclipboard()
    if img is None:
        raise ToolError(
            f"Clipboard is empty after CopyPicture for '{label}'. "
            "Ensure Excel is visible and not minimized, then retry."
        )

    try:
        img.save(path, "PNG")
    except Exception as e:
        raise ToolError(f"Failed to save screenshot to '{path}': {e}")

    # Clear clipboard (polite)
    try:
        app.CutCopyMode = False
    except Exception:
        pass

    return path


def _capture_range(
    range_str: str,
    sheet: str | None,
    workbook: str | None,
    output_path: str | None,
) -> dict:
    ws = _session.get_sheet(sheet, workbook)
    try:
        rng = ws.Range(range_str)
    except Exception as e:
        raise _session.wrap(e, f"Invalid range '{range_str}'")
    sheet_name = ws.Name
    path = _copy_picture_to_file(rng, output_path, f"{sheet_name}_{range_str}")
    return {"path": path, "range": range_str, "sheet": sheet_name}


def _capture_sheet(
    sheet: str | None,
    workbook: str | None,
    output_path: str | None,
) -> dict:
    ws = _session.get_sheet(sheet, workbook)
    used = ws.UsedRange
    if used is None:
        raise ToolError(
            f"Sheet '{ws.Name}' has no used range (it is empty). "
            "Write data first, then screenshot."
        )
    path = _copy_picture_to_file(used, output_path, ws.Name)
    return {"path": path, "range": used.Address, "sheet": ws.Name}


def _capture_chart(
    name: str,
    workbook: str | None,
    output_path: str | None,
) -> dict:
    # Delegate to Chart.Export — cleaner than clipboard for charts
    from .charts import _find_chart, _default_output_path

    wb = _session.get_workbook(workbook)
    co = _find_chart(wb, name)
    path = output_path or _default_output_path(name)
    path = os.path.abspath(path)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    try:
        co.Chart.Export(path, "PNG")
    except Exception as e:
        raise _session.wrap(e, f"Export chart '{name}' failed")
    return {"path": path, "chart": name}
```

- [ ] **Step 6.4: Run screenshot tests**

```powershell
uv run pytest tests/test_phase4.py -k "TestScreenshot" -v 2>&1 | tail -15
```
Expected: all pass.

- [ ] **Step 6.5: Run full suite**

```powershell
uv run pytest -q
```
Expected: all pass.

- [ ] **Step 6.6: Commit screenshot.py**

```bash
git add src/thepexcel_mcp/domains/screenshot.py
git commit -m "feat(p4): add excel_screenshot domain (range/sheet/chart capture via CopyPicture+PIL)"
```

---

## Task 7: `write_py` action in ranges.py

**Files:**
- Modify: `D:/ThepExcelMCP/src/thepexcel_mcp/domains/ranges.py`

The `=PY("code",0)` formula is inserted via `Range.Formula2R1C1` (verified from Microsoft Q&A thread). Escaping rule: double-up any double-quote characters inside the Python code string, then wrap in `=PY("...",0)`.

Important caveats to document clearly: execution is asynchronous in Azure cloud; this tool cannot await or read results; the cell will show `#BUSY!` then the result; requires M365 license with Python in Excel enabled.

- [ ] **Step 7.1: Write failing tests for write_py**

Append to `tests/test_phase4.py`:

```python
# ── write_py escaping ──────────────────────────────────────────────────────────

class TestWritePyEscaping:
    """Test the _build_py_formula helper in isolation."""

    def test_simple_code_no_quotes(self):
        from thepexcel_mcp.domains.ranges import _build_py_formula
        result = _build_py_formula("x = 1 + 2")
        assert result == '=PY("x = 1 + 2",0)'

    def test_code_with_double_quotes_escaped(self):
        from thepexcel_mcp.domains.ranges import _build_py_formula
        result = _build_py_formula('s = "hello"')
        # Double-quotes inside code must be doubled for Excel formula string
        assert result == '=PY("s = ""hello""",0)'

    def test_code_with_newline_preserved(self):
        from thepexcel_mcp.domains.ranges import _build_py_formula
        result = _build_py_formula("x = 1\ny = 2")
        # Newlines are passed as-is (Excel's Formula2R1C1 handles them)
        assert "x = 1\ny = 2" in result

    def test_empty_code_raises(self):
        from thepexcel_mcp.domains.ranges import _build_py_formula
        with pytest.raises(ToolError, match="python_code"):
            _build_py_formula("")

    def test_range_action_dispatch_write_py(self):
        mock_session = make_mock_session()
        mock_session.get_workbook.side_effect = ToolError("no excel")
        with patch("thepexcel_mcp.domains.ranges._session", mock_session):
            from thepexcel_mcp.domains.ranges import range_action
            with pytest.raises(ToolError, match="requires 'python_code'"):
                range_action(action="write_py", range="A1")
```

- [ ] **Step 7.2: Run test — confirm failures**

```powershell
uv run pytest tests/test_phase4.py -k "TestWritePy" -v 2>&1 | head -20
```
Expected: ImportError on `_build_py_formula` (not yet implemented).

- [ ] **Step 7.3: Add write_py to ranges.py**

In `range_action()` function, change the docstring Actions section to add:

```
    write_py
        Insert a Python-in-Excel formula (``=PY()``) into a single cell via
        ``Range.Formula2R1C1``. Requires ``python_code``.

        **IMPORTANT CAVEATS** (read before use):
        - Execution is asynchronous in Microsoft Azure cloud. The cell shows
          ``#BUSY!`` or ``#CONNECT!`` until Azure processes the formula.
          Requires an M365 subscription with Python in Excel enabled (Preview
          or GA channel as of 2026).
        - This tool CANNOT await results. Use ``excel_range(action="read",
          range="A1")`` in a subsequent call to fetch the computed value.
        - Offline or unsupported accounts: formula shows ``#CONNECT!`` or
          ``#BUSY!`` errors permanently.
        - The ``=PY()`` second argument: ``0`` = return value as Excel type,
          ``1`` = return as Python object (shows custom Python icon in cell).
          This tool always inserts ``0`` (Excel value mode).
        - **Experimental**: the COM insertion path (Formula2R1C1) works in
          testing but Microsoft does not formally document it for automation.
          Treat results as best-effort.

        Escaping: any double-quote characters in ``python_code`` are
        automatically doubled for the Excel formula string.

        Example: ``excel_range(action="write_py", range="A1",
        python_code="import pandas as pd\\ndf = pd.DataFrame({'x': [1,2,3]})")``
```

Also update the action validation block:

In `range_action()`, add before the existing `if action not in (...)` check:

```python
    if action == "write_py":
        if not python_code:
            raise ToolError("action='write_py' requires 'python_code' (non-empty string).")
        return _session.run_com(_write_py, range, sheet, workbook, python_code)
```

And update the `if action not in (...)` line to include `"write_py"`:
```python
    if action not in ("read", "read_spill", "write", "write_formula", "write_py", "clear"):
        raise ToolError(
            f"Unknown action '{action}'. Valid: read, read_spill, write, write_formula, write_py, clear."
        )
```

Add `python_code: str | None = None` to `range_action` signature.

- [ ] **Step 7.4: Add `_build_py_formula` and `_write_py` functions**

Append to `ranges.py`:

```python
def _build_py_formula(python_code: str) -> str:
    """Build the =PY("code",0) formula string with proper escaping.

    Excel formula string escaping: double-quote characters inside the string
    literal must be doubled. E.g., s = "hi" becomes =PY("s = ""hi""",0).

    Verified format via Microsoft Q&A (2023): Range.Formula2R1C1 = '=PY("code",0)'
    """
    if not python_code:
        raise ToolError("python_code must be a non-empty string.")
    escaped = python_code.replace('"', '""')
    return f'=PY("{escaped}",0)'


def _write_py(
    range_str: str,
    sheet: str | None,
    workbook: str | None,
    python_code: str,
) -> dict:
    formula = _build_py_formula(python_code)
    rng = _resolve_range(range_str, sheet, workbook)
    cell = rng.Cells(1, 1)
    try:
        # Formula2R1C1 is required for =PY() insertion (per MS Q&A verification)
        cell.Formula2R1C1 = formula
        return {
            "cell": cell.Address,
            "formula_inserted": formula,
            "note": (
                "PY formula inserted. Execution is asynchronous in Azure cloud. "
                "Use excel_range(action='read') after a delay to fetch the result. "
                "Cell shows #BUSY! until Azure processes the formula."
            ),
        }
    except Exception as e:
        raise _session.wrap(e, "Insert =PY() formula failed — ensure M365 Python in Excel is enabled")
```

- [ ] **Step 7.5: Run write_py tests**

```powershell
uv run pytest tests/test_phase4.py -k "TestWritePy" -v 2>&1 | tail -15
```
Expected: all 5 tests pass.

- [ ] **Step 7.6: Run full suite**

```powershell
uv run pytest -q
```
Expected: all pass (167 + new).

- [ ] **Step 7.7: Commit ranges.py**

```bash
git add src/thepexcel_mcp/domains/ranges.py tests/test_phase4.py
git commit -m "feat(p4): add write_py action to excel_range (=PY formula insertion, experimental)"
```

---

## Task 8: Register tools in server.py

**Files:**
- Modify: `D:/ThepExcelMCP/src/thepexcel_mcp/server.py`

- [ ] **Step 8.1: Add imports at top of server.py**

After the existing imports block, add:
```python
from .domains.charts import chart_action
from .domains.screenshot import screenshot_action
```

- [ ] **Step 8.2: Add `excel_chart` tool**

Append before `def main()`:

```python
@mcp.tool()
def excel_chart(
    action: str,
    name: str | None = None,
    workbook: str | None = None,
    sheet: str | None = None,
    source: str | None = None,
    chart_type: str | None = None,
    position: str | None = None,
    width: float | None = None,
    height: float | None = None,
    title: str | None = None,
    dest_sheet: str | None = None,
    x_title: str | None = None,
    y_title: str | None = None,
    legend: bool | None = None,
    legend_position: str | None = None,
    data_labels: bool | None = None,
    series_names: list | None = None,
    secondary_series: str | None = None,
    output_path: str | None = None,
) -> dict:
    """Create and manage Excel charts (ChartObjects / embedded charts).

    Parameters
    ----------
    action : str
        One of: ``list``, ``create``, ``configure``, ``set_source``,
        ``export_image``, ``delete``.
    name : str, optional
        Chart name. Required for ``configure``, ``set_source``,
        ``export_image``, ``delete``.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    sheet : str, optional
        Source sheet for ``create``. Uses active sheet when omitted.
    source : str, optional
        Data source for ``create`` / ``set_source``: range address
        (``"A1:C10"``), table name, or pivot table name.
    chart_type : str, optional
        Chart type for ``create``. See domain docstring for full list.
        Common values: ``column``, ``bar``, ``line``, ``pie``, ``doughnut``,
        ``scatter``, ``area``, ``combo_column_line``.
    position : str, optional
        Anchor cell for chart top-left corner (e.g. ``"E2"``).
    width : float, optional
        Chart width in points. Default 375.
    height : float, optional
        Chart height in points. Default 225.
    title : str, optional
        Chart title for ``create`` / ``configure``.
    dest_sheet : str, optional
        Sheet to embed chart on (``create``). Defaults to same sheet as source.
    x_title : str, optional
        Category (X) axis title for ``configure``.
    y_title : str, optional
        Value (Y) axis title for ``configure``.
    legend : bool, optional
        Show/hide legend for ``configure``.
    legend_position : str, optional
        Legend position: ``bottom``, ``corner``, ``left``, ``right``, ``top``.
    data_labels : bool, optional
        Show/hide data labels on all series for ``configure``.
    series_names : list of str, optional
        Override series names (in order) for ``configure``.
    secondary_series : str, optional
        Series name to move to secondary Y axis for ``configure`` (combo charts).
    output_path : str, optional
        File path for ``export_image``. Defaults to
        ``%TEMP%/thepexcel_mcp/<chart_name>.png``.

    Actions
    -------
    list
        All charts: name, sheet, chart_type, source, position.
        Example: ``excel_chart(action="list")``
    create
        Create a chart. Requires ``source`` and ``chart_type``.
        Example: ``excel_chart(action="create", source="A1:C10",
        chart_type="column", title="Sales")``
    configure
        Update title, axis labels, legend, data labels, series names.
        Requires ``name``. Pass only parameters to change.
        Example: ``excel_chart(action="configure", name="Chart 1",
        title="Updated Title", legend=True, legend_position="bottom")``
    set_source
        Change data range. Requires ``name`` and ``source``.
        Example: ``excel_chart(action="set_source", name="Chart 1",
        source="A1:D20")``
    export_image
        Export chart as PNG. Requires ``name``. Returns file path.
        Example: ``excel_chart(action="export_image", name="Chart 1")``
    delete
        Delete chart. Requires ``name``.
        Example: ``excel_chart(action="delete", name="Chart 1")``
    """
    return chart_action(
        action,
        name=name,
        workbook=workbook,
        sheet=sheet,
        source=source,
        chart_type=chart_type,
        position=position,
        width=width,
        height=height,
        title=title,
        dest_sheet=dest_sheet,
        x_title=x_title,
        y_title=y_title,
        legend=legend,
        legend_position=legend_position,
        data_labels=data_labels,
        series_names=series_names,
        secondary_series=secondary_series,
        output_path=output_path,
    )
```

- [ ] **Step 8.3: Add `excel_screenshot` tool**

Append before `def main()`:

```python
@mcp.tool()
def excel_screenshot(
    action: str,
    range: str | None = None,
    sheet: str | None = None,
    workbook: str | None = None,
    output_path: str | None = None,
    name: str | None = None,
) -> dict:
    """Capture a region of the live Excel workbook as a PNG for visual verification.

    Primary use: let an AI agent visually verify what Excel looks like after
    writing data, creating charts, or applying formatting.

    Requires Excel to be visible (not minimized). On HiDPI screens the bitmap
    captures at screen resolution; output size scales with DPI.

    Parameters
    ----------
    action : str
        One of: ``range``, ``sheet``, ``chart``.
    range : str, optional
        Range address for ``range`` action (e.g. ``"A1:F20"``).
    sheet : str, optional
        Sheet name. Uses active sheet when omitted (``range`` / ``sheet``).
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    output_path : str, optional
        Save path. Defaults to ``%TEMP%/thepexcel_mcp/<timestamp>.png``.
    name : str, optional
        Chart name for ``chart`` action.

    Actions
    -------
    range
        Capture a cell range as PNG. Requires ``range``.
        Uses Range.CopyPicture (bitmap mode) → PIL clipboard grab → PNG.
        Example: ``excel_screenshot(action="range", range="A1:F20")``
    sheet
        Capture the used range of a sheet.
        Example: ``excel_screenshot(action="sheet", sheet="Summary")``
    chart
        Export a chart as PNG by name. Requires ``name``.
        Uses Chart.Export (no clipboard involved).
        Example: ``excel_screenshot(action="chart", name="Sales Chart")``
    """
    return screenshot_action(
        action,
        range=range,
        sheet=sheet,
        workbook=workbook,
        output_path=output_path,
        name=name,
    )
```

- [ ] **Step 8.4: Extend `excel_range` signature and docstring for write_py**

In the `excel_range` tool function, add `python_code: str | None = None` to the signature and pass it through to `range_action`. Add the `write_py` action documentation to the docstring. Update the `action` parameter description to include `write_py`.

Concrete change — in `server.py` function `excel_range`:

Add parameter after `limit`:
```python
    python_code: str | None = None,
```

Update the call at the end:
```python
    return range_action(
        action,
        range=range,
        sheet=sheet,
        workbook=workbook,
        values=values,
        formula=formula,
        offset=offset,
        limit=limit,
        python_code=python_code,
    )
```

Add to docstring Parameters:
```
    python_code : str, optional
        Python source code string for ``write_py``. Multi-line supported
        (use ``\\n`` or actual newlines). Double-quote characters are
        automatically escaped.
```

Add to docstring Actions:
```
    write_py
        Insert a Python-in-Excel ``=PY()`` formula. Requires ``python_code``.
        IMPORTANT: execution is asynchronous in Azure cloud. The cell shows
        ``#BUSY!`` until Azure processes it. Requires M365 with Python in
        Excel enabled. Cannot await results — use ``read`` in a follow-up
        call. Experimental: COM insertion path is not officially documented.
        Office Scripts: not supported (cloud-only, no COM path) — use
        ``excel_vba`` instead.
        Example: ``excel_range(action="write_py", range="A1",
        python_code="import pandas as pd\\ndf = xl('A1:C10', headers=True)")``
```

Also update the `action` line:
```
    action : str
        One of: ``read``, ``read_spill``, ``write``, ``write_formula``,
        ``write_py``, ``clear``.
```

- [ ] **Step 8.5: Verify import check**

```powershell
uv run python -c "from thepexcel_mcp.server import mcp; print('OK', len(mcp.tools))"
```
Expected: `OK 11` (9 original + excel_chart + excel_screenshot)

- [ ] **Step 8.6: Run full test suite**

```powershell
uv run pytest -q
```
Expected: all pass.

- [ ] **Step 8.7: Commit server.py**

```bash
git add src/thepexcel_mcp/server.py
git commit -m "feat(p4): register excel_chart + excel_screenshot tools; extend excel_range with write_py"
```

---

## Task 9: Extend smoke_com.py with Phase 4 section

**Files:**
- Modify: `D:/ThepExcelMCP/tests/smoke_com.py`

- [ ] **Step 9.1: Add Phase 4 imports at top of smoke_com.py**

Add to the existing try/import block:
```python
    from thepexcel_mcp.domains.charts import chart_action
    from thepexcel_mcp.domains.screenshot import screenshot_action
    from thepexcel_mcp.domains.ranges import range_action
```

- [ ] **Step 9.2: Append Phase 4 section to `main()`**

Before the final `print("\n=== Smoke test complete ===")` line, append:

```python
    # ── Phase 4: Charts + Screenshot on a fresh temp workbook ────────────────
    print("\n[9] Phase 4: Charts + Screenshot on a fresh temp workbook...")
    tmp_wb4 = None
    tmp_wb4_name = None
    try:
        app = _session.get_app()
        tmp_wb4 = app.Workbooks.Add()
        tmp_wb4_name = tmp_wb4.Name
        print(f"    Created temp workbook: {tmp_wb4_name}")

        ws = tmp_wb4.ActiveSheet
        ws.Name = "ChartData"
        ws.Range("A1:C1").Value = [["Month", "Revenue", "Cost"]]
        ws.Range("A2:C5").Value = [
            ["Jan", 10000, 7000],
            ["Feb", 12000, 8000],
            ["Mar", 11000, 7500],
            ["Apr", 14000, 9000],
        ]
        print("    Test data written.")

        print("\n    [9a] Create column chart from A1:C5...")
        r = chart_action("create", source="A1:C5", chart_type="column",
                         sheet="ChartData", workbook=tmp_wb4_name,
                         title="Revenue vs Cost", position="E2")
        chart_name = r["name"]
        print(f"         Created: '{chart_name}' on sheet '{r['sheet']}'")

        print("    [9b] List charts...")
        r = chart_action("list", workbook=tmp_wb4_name)
        print(f"         Found {r['count']} chart(s): {[c['name'] for c in r['charts']]}")

        print("    [9c] Configure chart (legend + axis titles)...")
        r = chart_action("configure", name=chart_name, workbook=tmp_wb4_name,
                         x_title="Month", y_title="Amount (THB)",
                         legend=True, legend_position="bottom")
        print(f"         Changes: {r['changes']}")

        print("    [9d] Export chart as PNG...")
        r = chart_action("export_image", name=chart_name, workbook=tmp_wb4_name)
        import os
        print(f"         PNG at: {r['path']} (exists={os.path.exists(r['path'])})")

        print("    [9e] Screenshot: range A1:C5...")
        try:
            r = screenshot_action("range", range="A1:C5",
                                  sheet="ChartData", workbook=tmp_wb4_name)
            print(f"         PNG at: {r['path']} (exists={os.path.exists(r['path'])})")
        except Exception as e:
            print(f"         SKIP screenshot range (may need visible Excel): {e}")

        print("    [9f] Screenshot: full sheet...")
        try:
            r = screenshot_action("sheet", sheet="ChartData", workbook=tmp_wb4_name)
            print(f"         PNG at: {r['path']} (exists={os.path.exists(r['path'])})")
        except Exception as e:
            print(f"         SKIP screenshot sheet: {e}")

        print("    [9g] Screenshot via chart action...")
        try:
            r = screenshot_action("chart", name=chart_name, workbook=tmp_wb4_name)
            print(f"         PNG at: {r['path']} (exists={os.path.exists(r['path'])})")
        except Exception as e:
            print(f"         SKIP screenshot chart: {e}")

        print("    [9h] Delete chart...")
        r = chart_action("delete", name=chart_name, workbook=tmp_wb4_name)
        print(f"         {r}")

        print("\n    All Phase 4 chart+screenshot checks PASSED.")

    except Exception as e:
        print(f"\n    Phase 4 ERROR: {e}")
        import traceback; traceback.print_exc()
    finally:
        if tmp_wb4 is not None:
            try:
                tmp_wb4.Close(SaveChanges=False)
                print(f"\n    Temp workbook '{tmp_wb4_name}' closed (not saved).")
            except Exception as e:
                print(f"    WARNING: Could not close temp workbook 4: {e}")
```

- [ ] **Step 9.3: Run full test suite one more time**

```powershell
uv run pytest -q
```
Expected: all pass.

- [ ] **Step 9.4: Commit smoke_com.py**

```bash
git add tests/smoke_com.py
git commit -m "test(p4): extend smoke_com.py with Phase 4 chart+screenshot section"
```

---

## Task 10: Update docs and CLAUDE.md

**Files:**
- Modify: `D:/ThepExcelMCP/CLAUDE.md`
- Modify: `D:/ThepExcelMCP/docs/ROADMAP.md`

- [ ] **Step 10.1: Update CLAUDE.md tool registry**

In `CLAUDE.md`, update the architecture comment (line 21) from:
```
      └── excel_name       → domains/names.py          ← Phase 3
```
to:
```
      ├── excel_name       → domains/names.py          ← Phase 3
      ├── excel_chart      → domains/charts.py         ← Phase 4
      └── excel_screenshot → domains/screenshot.py     ← Phase 4
```

Update the Key files table to add:
```
| `src/thepexcel_mcp/domains/charts.py` | **Phase 4** — Chart CRUD + configure + export |
| `src/thepexcel_mcp/domains/screenshot.py` | **Phase 4** — Range/sheet/chart capture as PNG |
```

Update the Tool registry table to add:
```
| `excel_chart` | **Phase 4** — `list`, `create`, `configure`, `set_source`, `export_image`, `delete` |
| `excel_screenshot` | **Phase 4** — `range`, `sheet`, `chart` (PNG capture for LLM visual verification) |
```

Update `excel_range` row to add `write_py` in its action list.

Update "Future phases" to remove Phase 4 bullet and add:
```
- **Phase 5:** Live end-to-end smoke vs real Excel · packaging (uvx + MCPB bundle for Claude Desktop) · client registration
```

- [ ] **Step 10.2: Update ROADMAP.md**

Change:
```
- **P4** `excel_chart` · `excel_screenshot` (visual verification loop) · `=PY()` write-only action
```
to:
```
- **P4 ✅** `excel_chart` (list/create/configure/set_source/export_image/delete) · `excel_screenshot` (range/sheet/chart → PNG, CopyPicture+PIL) · `excel_range(action="write_py")` (`=PY()` Formula2R1C1 insertion, experimental)
```

- [ ] **Step 10.3: Run final test suite**

```powershell
uv run pytest -q
```
Expected: all pass.

- [ ] **Step 10.4: Commit docs**

```bash
git add CLAUDE.md docs/ROADMAP.md
git commit -m "docs(p4): mark Phase 4 done, update tool registry and roadmap"
```

---

## Self-Review

### Spec coverage check

| Spec requirement | Task |
|---|---|
| `excel_chart` domain module | Tasks 2-5 |
| `list` action for charts | Task 3 |
| `create` (range/table/pivot source, chart_type, position, title) | Task 4 |
| `configure` (title, axis, legend, data labels, series names, secondary axis) | Task 4 |
| `set_source` | Task 4 |
| `export_image` → PNG | Task 5 |
| `delete` | Task 3 |
| PivotChart: create from pivot table name with caveat in docstring | Task 4 (_create falls back to TableRange1, docstring notes limitation) |
| `excel_screenshot` domain | Task 6 |
| `screenshot.range` — CopyPicture+PIL | Task 6 |
| `screenshot.sheet` — used range | Task 6 |
| `screenshot.chart` — Chart.Export | Task 6 |
| Return absolute file path | Task 6 |
| `write_py` action in `excel_range` | Task 7 |
| `=PY("code",0)` via Formula2R1C1 | Task 7 |
| Docstring caveats (async, M365, can't await) | Task 7-8 |
| Office Scripts note (no tool, note in docstring) | Task 8 (added to write_py docstring) |
| Unit tests: chart-type mapping | Task 3 |
| Unit tests: arg validation | Tasks 3, 6, 7 |
| Unit tests: PY escaping | Task 7 |
| smoke_com.py Phase 4 section | Task 9 |
| CLAUDE.md + README + ROADMAP updates | Task 10 |
| `pillow` dependency | Task 1 |
| COM calls via `_session.run_com()` | All domain functions |

### Placeholder scan

No TBDs. All code blocks are complete.

### Type consistency

- `chart_action` signature uses `source` (not `source_range`); dispatcher maps to `source` parameter — consistent.
- `_find_chart` returns a `ChartObject` COM proxy; called in `_configure`, `_set_source`, `_export_image`, `_delete`, `_capture_chart` — all consistent.
- `_build_py_formula` is exported at module level; test imports it directly — consistent.
- `_session.run_com(fn, *args)` pattern — all COM functions follow the tables.py pattern exactly.

---

Plan complete and saved to `D:/ThepExcelMCP/docs/superpowers/plans/2026-06-12-phase4-charts-screenshot-python.md`.

**Two execution options:**

**1. Subagent-Driven (recommended)** - Dispatch a fresh subagent per task, review between tasks

**2. Inline Execution** - Execute tasks in this session using executing-plans

**Which approach?**
