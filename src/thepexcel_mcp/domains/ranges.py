"""Range read/write operations with pagination, dynamic array and spill support."""

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
    python_code: str | None = None,
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
        When a cell is the anchor of a spill, the response includes
        ``has_spill: true`` and ``spill_range``.
        When a cell is part of a spill (not the anchor), the response includes
        ``spill_parent`` with the anchor cell address.
    read_spill
        Given an anchor cell address, returns the full spill range.
        If the cell has no spill (HasSpill is False), returns a clear error.
        Response: ``{anchor, spill_range, values, total_rows, has_more,
        next_offset}``. Paginated via ``offset`` / ``limit``.
    write
        Write a 2-D list of values via ``Range.Value``.
        ``values`` must be a list-of-lists (rows × columns).
    write_formula
        Write a formula string via ``Range.Formula2`` to the top-left cell.
        Excel spills dynamic array results automatically.
        ``formula`` must start with ``=`` (e.g. ``"=UNIQUE(A1:A100)"``).
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
        - Office Scripts: not supported (cloud-only, no COM path) — use
          ``excel_vba`` instead.

        Escaping: any double-quote characters in ``python_code`` are
        automatically doubled for the Excel formula string.

        Example: ``excel_range(action="write_py", range="A1",
        python_code="import pandas as pd\\ndf = pd.DataFrame({'x': [1,2,3]})")``
    clear
        Clear cell contents (not formatting).
    """
    # Validate args (pure Python) before entering the COM worker
    if action == "write" and values is None:
        raise ToolError("action='write' requires 'values' (2-D list).")
    if action == "write_formula" and not formula:
        raise ToolError("action='write_formula' requires 'formula' starting with '='.")
    if action == "write_py":
        if not python_code:
            raise ToolError("action='write_py' requires 'python_code' (non-empty string).")
        return _session.run_com(_write_py, range, sheet, workbook, python_code)
    if action not in ("read", "read_spill", "write", "write_formula", "write_py", "clear"):
        raise ToolError(
            f"Unknown action '{action}'. Valid: read, read_spill, write, write_formula, write_py, clear."
        )
    return _session.run_com(_dispatch, action, range, sheet, workbook, values, formula, offset, limit)


def _dispatch(
    action: str,
    range_str: str,
    sheet: str | None,
    workbook: str | None,
    values,
    formula: str | None,
    offset: int,
    limit: int,
) -> dict:
    """Executed on the COM worker thread."""
    if action == "read":
        return _read(range_str, sheet, workbook, offset, limit)
    if action == "read_spill":
        return _read_spill(range_str, sheet, workbook, offset, limit)
    if action == "write":
        return _write(range_str, sheet, workbook, values)
    if action == "write_formula":
        return _write_formula(range_str, sheet, workbook, formula)
    return _clear(range_str, sheet, workbook)  # action == "clear"


def _resolve_range(range_str: str, sheet: str | None, workbook: str | None):
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


def _extract_values(raw, offset: int, limit: int) -> tuple[list, int]:
    """Normalise COM value result to list-of-lists, return (page, total_rows)."""
    if raw is None:
        return [], 0
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
    return rows, total_rows


def _read(
    range_str: str,
    sheet: str | None,
    workbook: str | None,
    offset: int,
    limit: int,
) -> dict:
    rng = _resolve_range(range_str, sheet, workbook)
    raw = rng.Value
    rows, total_rows = _extract_values(raw, offset, limit)
    if total_rows == 0:
        return {"values": [], "total_rows": 0, "has_more": False, "next_offset": 0}
    has_more = (offset + limit) < total_rows
    result = {
        "values": rows,
        "total_rows": total_rows,
        "has_more": has_more,
        "next_offset": offset + limit if has_more else None,
    }
    # Spill metadata for the top-left cell of the range
    try:
        anchor = rng.Cells(1, 1)
        if anchor.HasSpill:
            result["has_spill"] = True
            result["spill_range"] = anchor.SpillingRange.Address
        elif anchor.SpillParent is not None:
            result["spill_parent"] = anchor.SpillParent.Address
    except Exception:
        pass  # SpillParent/HasSpill not available in old Excel builds — ignore
    return result


def _read_spill(
    range_str: str,
    sheet: str | None,
    workbook: str | None,
    offset: int,
    limit: int,
) -> dict:
    """Return the full spill range for a dynamic-array anchor cell.

    Raises ToolError if the cell has no spill.
    """
    anchor_cell = _resolve_range(range_str, sheet, workbook).Cells(1, 1)
    try:
        has_spill = anchor_cell.HasSpill
    except Exception as e:
        raise _session.wrap(e, f"Cannot check HasSpill on '{range_str}'")
    if not has_spill:
        raise ToolError(
            f"Cell '{range_str}' has no spill range (HasSpill=False). "
            "Use action='read' to read the cell value directly."
        )
    spill_rng = anchor_cell.SpillingRange
    spill_addr = spill_rng.Address
    raw = spill_rng.Value
    rows, total_rows = _extract_values(raw, offset, limit)
    has_more = (offset + limit) < total_rows
    return {
        "anchor": anchor_cell.Address,
        "spill_range": spill_addr,
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


def _build_py_formula(python_code: str) -> str:
    """Build the =PY("code",0) formula string with proper escaping.

    Excel formula string escaping: double-quote characters inside the code
    string must be doubled. Example: s = "hi" becomes the formula
    =PY("s = \\"hi\\"",0) where each embedded quote is doubled.

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
