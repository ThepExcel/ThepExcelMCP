"""Range read/write operations with pagination and dynamic array support."""

from __future__ import annotations

from fastmcp.exceptions import ToolError

from ..session import ExcelSession

_session = ExcelSession()

_DEFAULT_LIMIT = 100
_MAX_CELL_LEN = 500  # truncate long cell strings to keep responses compact


def range_action(
    action: str,
    range: str,
    sheet: str | None = None,
    workbook: str | None = None,
    values: list | None = None,
    formula: str | None = None,
    offset: int = 0,
    limit: int = _DEFAULT_LIMIT,
) -> dict:
    """Dispatch a range action.

    Range addressing examples
    -------------------------
    - ``"A1:C10"``            — standard A1 notation (active sheet)
    - ``"Sheet1!A1:C10"``     — sheet-qualified (overrides ``sheet`` param)
    - ``"SalesTable[Amount]"`` — structured table reference

    Actions
    -------
    read
        Read cell values as a 2-D list. Paginated: default 100 rows.
        Returns: ``{values, total_rows, has_more, next_offset}``.
        Cells with strings >500 chars are truncated with ``…`` appended.
    write
        Write a 2-D list of values via ``Range.Value``.
        ``values`` must be a list-of-lists (rows × columns).
    write_formula
        Write a formula string via ``Range.Formula2`` to the top-left cell.
        Excel spills dynamic array results automatically.
        ``formula`` must start with ``=`` (e.g. ``"=UNIQUE(A1:A100)"``).
    clear
        Clear cell contents (not formatting).
    """
    if action == "read":
        return _read(range, sheet, workbook, offset, limit)
    if action == "write":
        if values is None:
            raise ToolError("action='write' requires 'values' (2-D list).")
        return _write(range, sheet, workbook, values)
    if action == "write_formula":
        if not formula:
            raise ToolError("action='write_formula' requires 'formula' starting with '='.")
        return _write_formula(range, sheet, workbook, formula)
    if action == "clear":
        return _clear(range, sheet, workbook)
    raise ToolError(
        f"Unknown action '{action}'. Valid: read, write, write_formula, clear."
    )


def _resolve_range(
    range_str: str, sheet: str | None, workbook: str | None
) -> "win32com.client.CDispatch":  # noqa: F821
    """Return a COM Range object, honouring sheet-qualified notation."""
    # "Sheet1!A1:C10" — sheet in the range string takes priority
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


def _read(
    range_str: str,
    sheet: str | None,
    workbook: str | None,
    offset: int,
    limit: int,
) -> dict:
    rng = _resolve_range(range_str, sheet, workbook)
    # Get full value array (COM returns tuple-of-tuples or scalar)
    raw = rng.Value
    if raw is None:
        return {"values": [], "total_rows": 0, "has_more": False, "next_offset": 0}
    # Normalise: scalar → 1×1, 1-D tuple → 1-row
    if not isinstance(raw, tuple):
        raw = ((raw,),)
    elif raw and not isinstance(raw[0], tuple):
        raw = (raw,)
    total_rows = len(raw)
    page = raw[offset: offset + limit]
    rows = []
    for row in page:
        cells = []
        for cell in row:
            if isinstance(cell, str) and len(cell) > _MAX_CELL_LEN:
                cell = cell[:_MAX_CELL_LEN] + "…"
            cells.append(cell)
        rows.append(cells)
    has_more = (offset + limit) < total_rows
    return {
        "values": rows,
        "total_rows": total_rows,
        "has_more": has_more,
        "next_offset": offset + limit if has_more else None,
    }


def _write(
    range_str: str,
    sheet: str | None,
    workbook: str | None,
    values: list,
) -> dict:
    rng = _resolve_range(range_str, sheet, workbook)
    try:
        rng.Value = values
        rows = len(values)
        cols = len(values[0]) if values else 0
        return {"written": {"rows": rows, "cols": cols, "range": range_str}}
    except Exception as e:
        raise _session.wrap(e, "Write failed")


def _write_formula(
    range_str: str,
    sheet: str | None,
    workbook: str | None,
    formula: str,
) -> dict:
    if not formula.startswith("="):
        raise ToolError("Formula must start with '=' (e.g. '=UNIQUE(A1:A100)').")
    rng = _resolve_range(range_str, sheet, workbook)
    # Use Formula2: supports dynamic-array spill (XLOOKUP, UNIQUE, FILTER, etc.)
    # Write only to the top-left cell; Excel handles the spill range.
    cell = rng.Cells(1, 1)
    try:
        cell.Formula2 = formula
        return {"formula_written": formula, "cell": cell.Address}
    except Exception as e:
        raise _session.wrap(e, "Write formula failed")


def _clear(range_str: str, sheet: str | None, workbook: str | None) -> dict:
    rng = _resolve_range(range_str, sheet, workbook)
    try:
        rng.ClearContents()
        return {"cleared": range_str}
    except Exception as e:
        raise _session.wrap(e, "Clear failed")
