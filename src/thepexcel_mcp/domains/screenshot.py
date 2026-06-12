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
    fname = f"{ts}_{safe}.png" if safe else f"{ts}.png"
    return os.path.join(tmp, fname)


def _copy_picture_to_file(rng, output_path: str | None, label: str) -> str:
    """CopyPicture -> clipboard -> PIL -> PNG file. Returns absolute path."""
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
