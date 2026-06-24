"""Row/column grouping & outline levels for Excel worksheets.

Wraps the following COM surface:
  ws.Rows(spec).Group()           — group rows (e.g. "2:5")
  ws.Columns(spec).Group()        — group columns (e.g. "B:D")
  ws.Rows(spec).Ungroup()         — ungroup rows
  ws.Columns(spec).Ungroup()      — ungroup columns
  ws.Outline.ShowLevels(...)      — show/collapse to specific outline level
  ws.Cells.ClearOutline()         — remove ALL outline groupings from sheet

Verify-effect strategy
----------------------
- group_rows  : read back ws.Rows(<first_row>).OutlineLevel — should be >= 2
- group_cols  : read back ws.Columns(<first_col>).OutlineLevel — should be >= 2
- ungroup_rows: read back ws.Rows(<first_row>).OutlineLevel — value returned as-is
- ungroup_cols: read back ws.Columns(<first_col>).OutlineLevel — value returned as-is
- show_levels : no single read-back property (collapse is visual); return the
                levels requested as the applied record
- clear       : no post-read (ClearOutline is destructive-complete); applied = True

COM gotchas
-----------
1. Group() must be called on a full-row or full-column range:
   ws.Rows("2:5") or ws.Columns("B:D"), NOT a cell range.
2. OutlineLevel is 1 for ungrouped rows/columns; >= 2 after grouping.
   Level 1 is the outermost summary level in Excel's outline model.
3. Max 8 nested outline levels. Exceeding raises a COM error — caught and
   wrapped as a ToolError.
4. ShowLevels(RowLevels=..., ColumnLevels=...) — both optional BUT at least
   one must be supplied. If 0 or omitted, that axis is unaffected.
5. Rows spec like "2:5" and column spec like "B:D" pass directly to
   ws.Rows() / ws.Columns(); Excel's own parser handles the rest.

Source: Microsoft Learn
  Range.OutlineLevel  — https://learn.microsoft.com/en-us/office/vba/api/excel.range.outlinelevel
  Range.Ungroup       — https://learn.microsoft.com/en-us/office/vba/api/excel.range.ungroup
  Outline.ShowLevels  — https://learn.microsoft.com/en-us/office/vba/api/excel.outline.showlevels
  Range.ClearOutline  — https://learn.microsoft.com/en-us/office/vba/api/excel.range.clearoutline
"""

from __future__ import annotations

from fastmcp.exceptions import ToolError

from ..session import ExcelSession, excel_guard

_session = ExcelSession()


# ── Public entry point ─────────────────────────────────────────────────────────

def outline_action(
    action: str,
    sheet: str | None = None,
    workbook: str | None = None,
    # row/column spec
    rows: str | None = None,
    columns: str | None = None,
    # show_levels
    row_levels: int | None = None,
    column_levels: int | None = None,
) -> dict:
    """Dispatch an outline/grouping action.

    Actions
    -------
    group_rows
        Group rows by spec (e.g. rows="2:5"). OutlineLevel of first row rises
        to >= 2 after success.
    group_columns
        Group columns by spec (e.g. columns="B:D"). OutlineLevel of first column
        rises to >= 2 after success.
    ungroup_rows
        Ungroup rows by spec (decreases outline level by 1).
    ungroup_columns
        Ungroup columns by spec (decreases outline level by 1).
    show_levels
        Show/collapse outline to specified level. Supply row_levels and/or
        column_levels (integers 1–8). At least one must be provided.
    clear
        Remove ALL row/column outline groupings from the sheet.
    """
    valid = {
        "group_rows", "group_columns",
        "ungroup_rows", "ungroup_columns",
        "show_levels", "clear",
    }
    if action not in valid:
        raise ToolError(
            f"Unknown action '{action}'. Valid: {', '.join(sorted(valid))}."
        )
    return _session.run_com(
        _dispatch,
        action, sheet, workbook,
        rows, columns,
        row_levels, column_levels,
    )


# ── Worker-thread dispatcher ───────────────────────────────────────────────────

def _dispatch(
    action: str,
    sheet: str | None,
    workbook: str | None,
    rows: str | None,
    columns: str | None,
    row_levels: int | None,
    column_levels: int | None,
) -> dict:
    """Executed on the STA COM worker thread."""
    ws = _session.get_sheet(sheet, workbook)
    app = ws.Application

    with excel_guard(app):
        if action == "group_rows":
            return _group_rows(ws, rows)
        if action == "group_columns":
            return _group_columns(ws, columns)
        if action == "ungroup_rows":
            return _ungroup_rows(ws, rows)
        if action == "ungroup_columns":
            return _ungroup_columns(ws, columns)
        if action == "show_levels":
            return _show_levels(ws, row_levels, column_levels)
        # action == "clear"
        return _clear(ws)


# ── Helpers ────────────────────────────────────────────────────────────────────

def _first_row_number(rows_spec: str) -> int:
    """Parse the first row number from a spec like '2:5' or '3'."""
    try:
        return int(rows_spec.split(":")[0].strip())
    except (ValueError, AttributeError):
        return 1


def _first_col_letter(cols_spec: str) -> str:
    """Parse the first column letter(s) from a spec like 'B:D' or 'C'."""
    return cols_spec.split(":")[0].strip()


# ── Action implementations ────────────────────────────────────────────────────

def _group_rows(ws, rows: str | None) -> dict:
    """Group a row range. VERIFY-EFFECT: OutlineLevel of first row >= 2."""
    if rows is None:
        raise ToolError(
            "group_rows action requires the 'rows' parameter (e.g. rows='2:5')."
        )
    try:
        ws.Rows(rows).Group()
        first_row = _first_row_number(rows)
        level = ws.Rows(first_row).OutlineLevel
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"group_rows '{rows}' failed")
    return {
        "outline": "group_rows",
        "sheet": ws.Name,
        "applied": {
            "rows": rows,
            "outline_level": level,
        },
    }


def _group_columns(ws, columns: str | None) -> dict:
    """Group a column range. VERIFY-EFFECT: OutlineLevel of first column >= 2."""
    if columns is None:
        raise ToolError(
            "group_columns action requires the 'columns' parameter (e.g. columns='B:D')."
        )
    try:
        ws.Columns(columns).Group()
        first_col = _first_col_letter(columns)
        level = ws.Columns(first_col).OutlineLevel
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"group_columns '{columns}' failed")
    return {
        "outline": "group_columns",
        "sheet": ws.Name,
        "applied": {
            "columns": columns,
            "outline_level": level,
        },
    }


def _ungroup_rows(ws, rows: str | None) -> dict:
    """Ungroup a row range (decrease outline level)."""
    if rows is None:
        raise ToolError(
            "ungroup_rows action requires the 'rows' parameter (e.g. rows='2:5')."
        )
    try:
        ws.Rows(rows).Ungroup()
        first_row = _first_row_number(rows)
        level = ws.Rows(first_row).OutlineLevel
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"ungroup_rows '{rows}' failed")
    return {
        "outline": "ungroup_rows",
        "sheet": ws.Name,
        "applied": {
            "rows": rows,
            "outline_level": level,
        },
    }


def _ungroup_columns(ws, columns: str | None) -> dict:
    """Ungroup a column range (decrease outline level)."""
    if columns is None:
        raise ToolError(
            "ungroup_columns action requires the 'columns' parameter (e.g. columns='B:D')."
        )
    try:
        ws.Columns(columns).Ungroup()
        first_col = _first_col_letter(columns)
        level = ws.Columns(first_col).OutlineLevel
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"ungroup_columns '{columns}' failed")
    return {
        "outline": "ungroup_columns",
        "sheet": ws.Name,
        "applied": {
            "columns": columns,
            "outline_level": level,
        },
    }


def _show_levels(ws, row_levels: int | None, column_levels: int | None) -> dict:
    """Show/collapse outline to the specified level(s).

    ShowLevels requires at least one of RowLevels or ColumnLevels.
    Valid range is 1–8 per axis.
    """
    if row_levels is None and column_levels is None:
        raise ToolError(
            "show_levels action requires at least one of 'row_levels' or "
            "'column_levels' (integers 1–8)."
        )
    # Validate ranges
    for name, val in (("row_levels", row_levels), ("column_levels", column_levels)):
        if val is not None and not (1 <= val <= 8):
            raise ToolError(
                f"'{name}' must be between 1 and 8 (got {val}). "
                "Excel supports a maximum of 8 outline levels."
            )
    try:
        kwargs: dict = {}
        if row_levels is not None:
            kwargs["RowLevels"] = row_levels
        if column_levels is not None:
            kwargs["ColumnLevels"] = column_levels
        ws.Outline.ShowLevels(**kwargs)
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "show_levels failed")
    applied: dict = {}
    if row_levels is not None:
        applied["row_levels"] = row_levels
    if column_levels is not None:
        applied["column_levels"] = column_levels
    return {
        "outline": "show_levels",
        "sheet": ws.Name,
        "applied": applied,
    }


def _clear(ws) -> dict:
    """Remove all outline groupings from the sheet via ws.Cells.ClearOutline()."""
    try:
        ws.Cells.ClearOutline()
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "clear outline failed")
    return {
        "outline": "clear",
        "sheet": ws.Name,
        "applied": {"cleared": True},
    }
