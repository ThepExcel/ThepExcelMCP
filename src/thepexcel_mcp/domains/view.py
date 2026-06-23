"""Worksheet / window VIEW settings.

Controls sheet-level display properties via the Excel Window object:
freeze/unfreeze panes, gridlines, zoom level, and row/column headings.

Window binding strategy: we use wb.Windows(1) (workbook-scoped) rather than
app.ActiveWindow (global). This is critical for multi-workbook / ROT-fallback
scenarios: ws.Activate() only selects the sheet inside its own workbook; if
the workbook is not the foreground workbook, app.ActiveWindow still points to
a different workbook's window, so all mutations would silently land on the
wrong target. wb.Activate() + wb.Windows(1) is always the target window.
"""

from __future__ import annotations

from fastmcp.exceptions import ToolError

from ..session import ExcelSession, excel_guard

_session = ExcelSession()


# ── Public entry point ─────────────────────────────────────────────────────────

def view_action(
    action: str,
    sheet: str | None = None,
    workbook: str | None = None,
    # freeze_panes
    cell: str | None = None,
    freeze_rows: int | None = None,
    freeze_cols: int | None = None,
    # gridlines
    show: bool | None = None,
    # zoom
    zoom: int | None = None,
) -> dict:
    """Dispatch a view/display action.

    Actions
    -------
    freeze_panes
        Freeze panes at a cell (e.g. cell="B2" → freeze row 1 + column A).
        Alternatively supply freeze_rows and/or freeze_cols directly.
    unfreeze_panes
        Remove freeze from the active window pane.
    gridlines
        Show or hide gridlines (show=True/False).
    zoom
        Set worksheet zoom level 10–400 (zoom=100 for 100%).
    headings
        Show or hide row/column headings (show=True/False).
    """
    valid = {"freeze_panes", "unfreeze_panes", "gridlines", "zoom", "headings"}
    if action not in valid:
        raise ToolError(
            f"Unknown action '{action}'. Valid: {', '.join(sorted(valid))}."
        )
    return _session.run_com(
        _dispatch,
        action, sheet, workbook,
        cell, freeze_rows, freeze_cols,
        show, zoom,
    )


# ── Worker-thread dispatcher ───────────────────────────────────────────────────

def _dispatch(
    action: str,
    sheet: str | None,
    workbook: str | None,
    cell: str | None,
    freeze_rows: int | None,
    freeze_cols: int | None,
    show: bool | None,
    zoom: int | None,
) -> dict:
    """Executed on the STA COM worker thread.

    Window binding: wb.Activate() brings the workbook to the foreground, then
    ws.Activate() selects the sheet inside it. We then bind to wb.Windows(1)
    (workbook-scoped) rather than app.ActiveWindow (global) so that in a
    multi-workbook / ROT-fallback scenario we always mutate the target window,
    never a bystander workbook's window.
    """
    ws = _session.get_sheet(sheet, workbook)
    wb = ws.Parent        # the Workbook owning this sheet
    wb.Activate()         # bring the workbook to the foreground
    ws.Activate()         # select the sheet within it
    app = ws.Application
    win = wb.Windows(1)   # workbook-scoped window — always the right target

    with excel_guard(app):
        if action == "freeze_panes":
            return _freeze_panes(ws, win, cell, freeze_rows, freeze_cols)
        if action == "unfreeze_panes":
            return _unfreeze_panes(win, ws)
        if action == "gridlines":
            return _gridlines(win, show, ws)
        if action == "zoom":
            return _zoom(win, zoom, ws)
        # action == "headings"
        return _headings(win, show, ws)


# ── Action implementations ────────────────────────────────────────────────────

def _freeze_panes(ws, win, cell: str | None, freeze_rows: int | None, freeze_cols: int | None) -> dict:
    """Freeze via split offsets — more reliable than Select-based approach.

    Resolves freeze_rows/freeze_cols from cell when not supplied directly.
    cell="B2" → rows above = 1, cols left = 1.
    """
    try:
        # Resolve row/col count from cell address if provided
        if cell is not None:
            try:
                ref = ws.Range(cell)
            except Exception as e:
                raise ToolError(f"Invalid cell reference '{cell}': {e}")
            rows_to_freeze = ref.Row - 1   # rows ABOVE the split
            cols_to_freeze = ref.Column - 1  # cols LEFT of the split
        else:
            rows_to_freeze = freeze_rows if freeze_rows is not None else 0
            cols_to_freeze = freeze_cols if freeze_cols is not None else 0

        if rows_to_freeze < 0 or cols_to_freeze < 0:
            raise ToolError(
                "freeze_panes: cell must not be in row 1 or column A "
                "(there must be at least one row/column above/left to freeze)."
            )
        if rows_to_freeze == 0 and cols_to_freeze == 0:
            raise ToolError(
                "freeze_panes: specify cell (e.g. 'B2'), freeze_rows, or freeze_cols. "
                "Cell 'A1' results in no freeze — nothing to freeze above/left of A1."
            )

        # Unfreeze first (required before re-applying — Excel ignores FreezePanes=True
        # if panes are already frozen from a different position)
        win.FreezePanes = False
        win.SplitRow = rows_to_freeze
        win.SplitColumn = cols_to_freeze
        win.FreezePanes = True

        return {
            "view": "freeze_panes",
            "sheet": ws.Name,
            "applied": {
                "freeze_rows": rows_to_freeze,
                "freeze_cols": cols_to_freeze,
                "cell": cell,
            },
        }
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "freeze_panes failed")


def _unfreeze_panes(win, ws) -> dict:
    try:
        win.FreezePanes = False
        win.SplitRow = 0
        win.SplitColumn = 0
    except Exception as e:
        raise _session.wrap(e, "unfreeze_panes failed")
    return {
        "view": "unfreeze_panes",
        "sheet": ws.Name,
        "applied": {"frozen": False},
    }


def _gridlines(win, show: bool | None, ws) -> dict:
    if show is None:
        raise ToolError("gridlines action requires 'show' (True or False).")
    try:
        win.DisplayGridlines = bool(show)
    except Exception as e:
        raise _session.wrap(e, "gridlines failed")
    return {
        "view": "gridlines",
        "sheet": ws.Name,
        "applied": {"show": show},
    }


def _zoom(win, zoom: int | None, ws) -> dict:
    if zoom is None:
        raise ToolError("zoom action requires the 'zoom' parameter (integer 10–400).")
    if not (10 <= zoom <= 400):
        raise ToolError(
            f"zoom must be between 10 and 400 (got {zoom})."
        )
    try:
        win.Zoom = int(zoom)
    except Exception as e:
        raise _session.wrap(e, "zoom failed")
    return {
        "view": "zoom",
        "sheet": ws.Name,
        "applied": {"zoom": zoom},
    }


def _headings(win, show: bool | None, ws) -> dict:
    if show is None:
        raise ToolError("headings action requires 'show' (True or False).")
    try:
        win.DisplayHeadings = bool(show)
    except Exception as e:
        raise _session.wrap(e, "headings failed")
    return {
        "view": "headings",
        "sheet": ws.Name,
        "applied": {"show": show},
    }
