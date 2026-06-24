"""Cell-level diff between two ranges or sheets.

Compares values and/or formulas across two named ranges or entire sheets.
PURE READ — no mutations, no excel_guard needed, no deadlock surface.

COM properties used
-------------------
- ``rng.Value``    — returns scalar for 1 cell, flat tuple for 1 row/col,
                     tuple-of-tuples for a block. Normalised via _to_2d().
- ``rng.Formula``  — same shape contract as .Value.
- ``rng.Row``      — 1-based row index of range top-left (for A1 mapping).
- ``rng.Column``   — 1-based column index of range top-left.
- ``rng.Rows.Count`` / ``rng.Columns.Count`` — range dimensions.
- ``ws.UsedRange`` — bounding box of non-empty cells on the sheet.
- ``ws.UsedRange.Row`` / ``.Rows.Count`` / ``.Column`` / ``.Columns.Count``
                   — to compute last-used row/col.
- ``ws.Cells(r, c)`` — used to form the sheet-level block range.

Shape contract for rng.Value / rng.Formula (pywin32 late-binding)
-----------------------------------------------------------------
- 1 cell  → scalar (str / float / bool / None / pywintypes.datetime)
- 1 row   → flat tuple  e.g. (1, 2, 3)
- 1 col   → flat tuple  e.g. (1, 2, 3)  (same as single row)
- N×M     → tuple-of-tuples  ((r1c1, r1c2), (r2c1, r2c2), ...)

_to_2d() normalises all forms to list[list[Any]] so diff logic is uniform.
Confirmed in ranges.py::_extract_values (same project, same COM path).

None vs empty-string
--------------------
Raw values are compared directly: None (blank cell in COM) and "" (empty
string typed by user) are treated as DIFFERENT. This is the least-surprising
behaviour — callers who need blank-equivalence can post-filter.

A1 address mapping
------------------
Left range top-left is at (rng.Row, rng.Column) in 1-based Excel coords.
Diff item at (row_i, col_j) (0-based) → Excel row = rng.Row + row_i,
Excel col = rng.Column + col_j. Column integer → letter(s) via _col_to_letter().
"""

from __future__ import annotations

from typing import Any

from fastmcp.exceptions import ToolError

from ..session import ExcelSession

_session = ExcelSession()


# ── Public entry point ─────────────────────────────────────────────────────────

def diff_action(
    action: str,
    left_range: str | None = None,
    left_sheet: str | None = None,
    left_workbook: str | None = None,
    right_range: str | None = None,
    right_sheet: str | None = None,
    right_workbook: str | None = None,
    compare: str = "value",
    max_diffs: int = 500,
) -> dict:
    """Compare two ranges or sheets and report cell-level differences.

    Actions
    -------
    ranges
        Compare left_range vs right_range cell-by-cell.
        Required: left_range, right_range.
        Optional: left_sheet, left_workbook, right_sheet, right_workbook
        (also parsed from range string when it contains '!').
    sheets
        Compare two whole sheets over their combined UsedRange bounding box.
        Required: left_sheet, right_sheet.
        Optional: left_workbook, right_workbook.

    compare
        "value"   — compare rng.Value only (default)
        "formula" — compare rng.Formula only
        "both"    — diff if EITHER value or formula differs

    max_diffs
        Cap on diff entries returned (default 500).
        total_diffs always reflects the true count even when truncated=True.
    """
    valid = {"ranges", "sheets"}
    if action not in valid:
        raise ToolError(
            f"Unknown action '{action}'. Valid: {', '.join(sorted(valid))}."
        )
    _compare_valid = {"value", "formula", "both"}
    if compare not in _compare_valid:
        raise ToolError(
            f"Unknown compare mode '{compare}'. Valid: value, formula, both."
        )
    if not isinstance(max_diffs, int) or max_diffs < 1:
        raise ToolError("max_diffs must be a positive integer.")

    return _session.run_com(
        _dispatch,
        action,
        left_range, left_sheet, left_workbook,
        right_range, right_sheet, right_workbook,
        compare, max_diffs,
    )


# ── Worker-thread dispatcher ───────────────────────────────────────────────────

def _dispatch(
    action: str,
    left_range: str | None,
    left_sheet: str | None,
    left_workbook: str | None,
    right_range: str | None,
    right_sheet: str | None,
    right_workbook: str | None,
    compare: str,
    max_diffs: int,
) -> dict:
    """Executed on the STA COM worker thread. Pure read — no guard needed."""
    if action == "ranges":
        return _diff_ranges(
            left_range, left_sheet, left_workbook,
            right_range, right_sheet, right_workbook,
            compare, max_diffs,
        )
    # action == "sheets"
    return _diff_sheets(
        left_sheet, left_workbook,
        right_sheet, right_workbook,
        compare, max_diffs,
    )


# ── Range resolver (verbatim from validation.py / sparkline.py) ───────────────

def _resolve_range(range_str: str, sheet: str | None, workbook: str | None):
    """Return a COM Range, honouring 'Sheet1!A1:B10' notation."""
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


# ── Normalisation helpers ─────────────────────────────────────────────────────

def _to_2d(raw: Any) -> list[list[Any]]:
    """Normalise COM .Value / .Formula result to list[list[Any]].

    pywin32 shape contract (confirmed from ranges.py::_extract_values):
      - scalar       → [[scalar]]          (single cell)
      - flat tuple   → [list(raw)]         (single row or single column)
      - tuple-of-tuples → [[...], ...]     (multi-row block)
      - None         → [[None]]            (blank single cell)
    """
    if raw is None:
        return [[None]]
    if not isinstance(raw, tuple):
        # Single cell scalar
        return [[raw]]
    if raw and not isinstance(raw[0], tuple):
        # Single row or single column — flat tuple
        return [list(raw)]
    # Multi-row block
    return [list(row) for row in raw]


def _col_to_letter(col: int) -> str:
    """Convert a 1-based column number to A1-style column letters.

    Examples: 1→'A', 26→'Z', 27→'AA', 703→'AAA'.
    """
    letters = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _offset_to_a1(base_row: int, base_col: int, row_i: int, col_j: int) -> str:
    """Return the A1 address for the left range at 0-based (row_i, col_j).

    base_row / base_col are 1-based (from rng.Row / rng.Column).
    """
    r = base_row + row_i
    c = base_col + col_j
    return f"{_col_to_letter(c)}{r}"


# ── Core diff engine ──────────────────────────────────────────────────────────

def _cell_differs(
    lv: Any, rv: Any, lf: Any, rf: Any, compare: str,
) -> bool:
    """Return True if the cell pair is considered different per compare mode."""
    if compare == "value":
        return lv != rv
    if compare == "formula":
        return lf != rf
    # "both": different if EITHER value or formula differs
    return (lv != rv) or (lf != rf)


def _build_diffs(
    left_values: list[list[Any]],
    right_values: list[list[Any]],
    left_formulas: list[list[Any]] | None,
    right_formulas: list[list[Any]] | None,
    compare: str,
    max_diffs: int,
    base_row: int,
    base_col: int,
) -> tuple[list[dict], int, bool]:
    """Compare 2-D grids and return (diffs_capped, total_diffs, truncated).

    Compares over the overlap (min rows × min cols of both grids).
    base_row / base_col: left grid's 1-based Excel row/col for A1 mapping.
    """
    n_rows_l = len(left_values)
    n_cols_l = len(left_values[0]) if left_values else 0
    n_rows_r = len(right_values)
    n_cols_r = len(right_values[0]) if right_values else 0

    rows = min(n_rows_l, n_rows_r)
    cols = min(n_cols_l, n_cols_r)

    all_diffs: list[dict] = []

    for i in range(rows):
        row_l = left_values[i]
        row_r = right_values[i]
        lf_row = left_formulas[i] if left_formulas is not None else None
        rf_row = right_formulas[i] if right_formulas is not None else None

        for j in range(cols):
            lv = row_l[j] if j < len(row_l) else None
            rv = row_r[j] if j < len(row_r) else None
            lf = lf_row[j] if (lf_row is not None and j < len(lf_row)) else None
            rf = rf_row[j] if (rf_row is not None and j < len(rf_row)) else None

            if _cell_differs(lv, rv, lf, rf, compare):
                entry: dict = {
                    "cell": _offset_to_a1(base_row, base_col, i, j),
                    "row": i,
                    "col": j,
                    "left_value": lv,
                    "right_value": rv,
                }
                if compare in ("formula", "both"):
                    entry["left_formula"] = lf
                    entry["right_formula"] = rf
                all_diffs.append(entry)

    total_diffs = len(all_diffs)
    truncated = total_diffs > max_diffs
    return all_diffs[:max_diffs], total_diffs, truncated


# ── Action implementations ────────────────────────────────────────────────────

def _read_range_data(
    rng, compare: str,
) -> tuple[list[list[Any]], list[list[Any]] | None]:
    """Read value and/or formula 2-D grids from a COM Range.

    Returns (values_2d, formulas_2d_or_None).
    Always reads .Value for value/both, .Formula for formula/both.
    """
    if compare in ("value", "both"):
        values = _to_2d(rng.Value)
    else:
        # formula-only: populate values from Formula (avoids extra COM call for .Value)
        values = _to_2d(rng.Formula)

    if compare in ("formula", "both"):
        # Re-use already-fetched data when formula-only (values == formulas grid)
        formulas = values if compare == "formula" else _to_2d(rng.Formula)
    else:
        formulas = None

    return values, formulas


def _diff_ranges(
    left_range_str: str | None,
    left_sheet: str | None,
    left_workbook: str | None,
    right_range_str: str | None,
    right_sheet: str | None,
    right_workbook: str | None,
    compare: str,
    max_diffs: int,
) -> dict:
    """Compare two named ranges cell-by-cell."""
    if not left_range_str:
        raise ToolError("ranges action requires 'left_range'.")
    if not right_range_str:
        raise ToolError("ranges action requires 'right_range'.")

    try:
        l_rng = _resolve_range(left_range_str, left_sheet, left_workbook)
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"Could not resolve left range '{left_range_str}'")

    try:
        r_rng = _resolve_range(right_range_str, right_sheet, right_workbook)
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"Could not resolve right range '{right_range_str}'")

    try:
        n_rows_l = l_rng.Rows.Count
        n_cols_l = l_rng.Columns.Count
        n_rows_r = r_rng.Rows.Count
        n_cols_r = r_rng.Columns.Count
        base_row = l_rng.Row
        base_col = l_rng.Column

        left_vals, left_forms = _read_range_data(l_rng, compare)
        right_vals, right_forms = _read_range_data(r_rng, compare)

    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "diff ranges: COM read failed")

    diffs, total_diffs, truncated = _build_diffs(
        left_vals, right_vals, left_forms, right_forms,
        compare, max_diffs, base_row, base_col,
    )

    dims_match = (n_rows_l == n_rows_r) and (n_cols_l == n_cols_r)
    applied: dict = {
        "dimensions_left": [n_rows_l, n_cols_l],
        "dimensions_right": [n_rows_r, n_cols_r],
        "dimensions_match": dims_match,
        "compare": compare,
        "total_diffs": total_diffs,
        "truncated": truncated,
        "max_diffs": max_diffs,
        "diffs": diffs,
    }
    if not dims_match:
        applied["shape_note"] = (
            f"Shapes differ: left={n_rows_l}×{n_cols_l}, "
            f"right={n_rows_r}×{n_cols_r}. "
            f"Diff covers overlap: "
            f"{min(n_rows_l, n_rows_r)}×{min(n_cols_l, n_cols_r)}."
        )

    return {"diff": "ranges", "applied": applied}


def _diff_sheets(
    left_sheet: str | None,
    left_workbook: str | None,
    right_sheet: str | None,
    right_workbook: str | None,
    compare: str,
    max_diffs: int,
) -> dict:
    """Compare two whole sheets over their combined UsedRange bounding box.

    Bounding box = max(last used row of both) × max(last used col of both).
    This ensures no data from either sheet is excluded from the diff.
    """
    if not left_sheet:
        raise ToolError("sheets action requires 'left_sheet'.")
    if not right_sheet:
        raise ToolError("sheets action requires 'right_sheet'.")

    try:
        l_ws = _session.get_sheet(left_sheet, left_workbook)
        r_ws = _session.get_sheet(right_sheet, right_workbook)
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "diff sheets: could not resolve sheet")

    try:
        l_ur = l_ws.UsedRange
        r_ur = r_ws.UsedRange

        # Last used row/col = first + count - 1
        l_last_row = l_ur.Row + l_ur.Rows.Count - 1
        l_last_col = l_ur.Column + l_ur.Columns.Count - 1
        r_last_row = r_ur.Row + r_ur.Rows.Count - 1
        r_last_col = r_ur.Column + r_ur.Columns.Count - 1

        # Union bounding box starting at (1, 1)
        max_row = max(l_last_row, r_last_row)
        max_col = max(l_last_col, r_last_col)

        # Read the full block from each sheet (both start at Cell(1,1))
        l_block = l_ws.Range(l_ws.Cells(1, 1), l_ws.Cells(max_row, max_col))
        r_block = r_ws.Range(r_ws.Cells(1, 1), r_ws.Cells(max_row, max_col))

        left_vals, left_forms = _read_range_data(l_block, compare)
        right_vals, right_forms = _read_range_data(r_block, compare)

    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "diff sheets: COM read failed")

    # Sheet-level diff always starts at A1 (base row=1, col=1)
    diffs, total_diffs, truncated = _build_diffs(
        left_vals, right_vals, left_forms, right_forms,
        compare, max_diffs, base_row=1, base_col=1,
    )

    return {
        "diff": "sheets",
        "applied": {
            "left_sheet": left_sheet,
            "right_sheet": right_sheet,
            "left_used_range": [l_last_row, l_last_col],
            "right_used_range": [r_last_row, r_last_col],
            "bounding_box": [max_row, max_col],
            "compare": compare,
            "total_diffs": total_diffs,
            "truncated": truncated,
            "max_diffs": max_diffs,
            "diffs": diffs,
        },
    }
