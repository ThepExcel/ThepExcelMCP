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
    value_mode: str = "typed",
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
        Spill metadata is only computed on the offset-0 page (pagination
        re-pays the same anchor probe for an identical answer on every later
        page, so it's skipped for ``offset > 0``):
        When a cell is the anchor of a spill, the response includes
        ``has_spill: true`` and ``spill_range``.
        When the requested range is a single cell (offset 0) that is part of
        a spill (not the anchor), the response includes ``spill_parent`` with
        the anchor cell address — this check is skipped for multi-cell reads,
        where "part of someone else's spill" isn't a meaningful answer.
        ``value_mode="typed"`` (default) uses ``Range.Value`` so Excel dates
        and currency keep their Python types. ``value_mode="raw"`` uses
        ``Range.Value2`` for faster bulk reads and returns those values as
        Excel serial numbers.
    read_spill
        Given an anchor cell address, returns the full spill range.
        If the cell has no spill (HasSpill is False), returns a clear error.
        Response: ``{anchor, spill_range, values, total_rows, has_more,
        next_offset}``. Paginated via ``offset`` / ``limit``.
    write
        Write a 2-D list of values via ``Range.Value``.
        ``values`` must be a list-of-lists (rows × columns).
        The target range is auto-resized to the data shape, so callers may
        pass just the top-left anchor cell (e.g. ``"A1"``) and the full block
        will be written correctly.
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
    if value_mode not in ("typed", "raw"):
        raise ToolError("value_mode must be 'typed' or 'raw'.")
    if action not in ("read", "read_spill") and value_mode != "typed":
        raise ToolError("value_mode='raw' is only valid for read/read_spill actions.")
    if action == "write_py":
        if not python_code:
            raise ToolError("action='write_py' requires 'python_code' (non-empty string).")
        return _session.run_com(_write_py, range, sheet, workbook, python_code)
    if action not in ("read", "read_spill", "write", "write_formula", "write_py", "clear"):
        raise ToolError(
            f"Unknown action '{action}'. Valid: read, read_spill, write, write_formula, write_py, clear."
        )
    return _session.run_com(
        _dispatch,
        action,
        range,
        sheet,
        workbook,
        values,
        formula,
        offset,
        limit,
        value_mode,
    )


def _dispatch(
    action: str,
    range_str: str,
    sheet: str | None,
    workbook: str | None,
    values,
    formula: str | None,
    offset: int,
    limit: int,
    value_mode: str,
) -> dict:
    """Executed on the COM worker thread."""
    if action == "read":
        return _read(range_str, sheet, workbook, offset, limit, value_mode)
    if action == "read_spill":
        return _read_spill(range_str, sheet, workbook, offset, limit, value_mode)
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


def _normalize_rows(raw) -> list:
    """Normalise a COM .Value/.Formula read (already page-sized) to list-of-lists.

    Handles the pywin32 shape contract: None (empty), scalar (1 cell), flat
    tuple (1 row or 1 col), tuple-of-tuples (N×M block). Truncates long
    strings same as before. Unlike the old ``_extract_values``, this does NOT
    slice by offset/limit — the caller is expected to have already read only
    the page-sized sub-range via a bounded ``Cells(...)`` block.
    """
    if raw is None:
        return []
    if not isinstance(raw, tuple):
        raw = ((raw,),)
    elif raw and not isinstance(raw[0], tuple):
        raw = (raw,)
    rows = []
    for row in raw:
        cells = []
        for cell in row:
            if isinstance(cell, str) and len(cell) > _MAX_CELL_LEN:
                cell = cell[:_MAX_CELL_LEN] + "…"
            cells.append(cell)
        rows.append(cells)
    return rows


def _read_values(rng, value_mode: str):
    """Read a COM Range using typed ``Value`` or faster raw ``Value2``."""
    return rng.Value2 if value_mode == "raw" else rng.Value


def _read(
    range_str: str,
    sheet: str | None,
    workbook: str | None,
    offset: int,
    limit: int,
    value_mode: str = "typed",
) -> dict:
    rng = _resolve_range(range_str, sheet, workbook)
    total_rows = rng.Rows.Count
    total_cols = rng.Columns.Count

    # Legacy quirk (kept for byte-identical response shape): a truly-empty
    # single cell has Rows.Count==Columns.Count==1 but .Value is None — the
    # old whole-value read reported that as total_rows=0, not "1 empty row".
    if total_rows == 1 and total_cols == 1 and _read_values(rng, value_mode) is None:
        return {"values": [], "total_rows": 0, "has_more": False, "next_offset": 0}

    if offset >= total_rows:
        rows: list = []
    else:
        r1, c1 = rng.Row, rng.Column
        last_row = r1 + min(offset + limit, total_rows) - 1
        ws = rng.Parent
        # Page-scoped read: only marshal the rows this page needs, via an
        # explicit Cells(...)-bounded block — never .Resize/.Offset (pywin32
        # indexed-property quirk, see ranges._write) and never the whole
        # range's .Value (that was O(total_rows) per page — quadratic across
        # pages for a large range read a page at a time).
        page_rng = ws.Range(
            ws.Cells(r1 + offset, c1),
            ws.Cells(last_row, c1 + total_cols - 1),
        )
        rows = _normalize_rows(_read_values(page_rng, value_mode))

    has_more = (offset + limit) < total_rows
    result = {
        "values": rows,
        "total_rows": total_rows,
        "has_more": has_more,
        "next_offset": offset + limit if has_more else None,
    }
    # Spill metadata for the top-left cell of the range — offset-0 only.
    # SpillingRange is a newer COM property not exposed via pywin32 late-binding
    # dispatch on all Excel builds. We use a HasSpill row/column scan fallback
    # that is robust across Excel versions (verified 2026-06).
    # Pagination re-pays this same anchor probe (2-4 COM round-trips) on every
    # page for an identical answer, so it only runs once, on the first page.
    if offset == 0:
        try:
            anchor = rng.Cells(1, 1)
            if anchor.HasSpill and _is_spill_anchor(anchor):
                result["has_spill"] = True
                result["spill_range"] = _spill_range_address(anchor)
            elif not anchor.HasSpill and total_rows == 1 and total_cols == 1:
                # "Part of someone else's spill" is only a meaningful question
                # for a single-cell read — skip the extra SpillParent
                # round-trip on multi-cell reads.
                parent_addr = _spill_parent_address(anchor)
                if parent_addr:
                    result["spill_parent"] = parent_addr
        except Exception:
            pass  # HasSpill not available in old Excel builds — ignore
    return result


def _is_spill_anchor(cell) -> bool:
    """True if the cell is the anchor (top-left) of a spill range.

    An anchor has HasSpill=True AND has a formula in Formula2.
    Spill-overflow cells have HasSpill=True but empty Formula2.
    """
    try:
        return bool(cell.HasSpill) and bool(cell.Formula2)
    except Exception:
        return False


_SPILL_SCAN_CAP = 999  # max offset from anchor; matches the old linear scan's
                        # 1000-cell cap (offsets 0..999, offset 0 == the anchor).


def _scan_spill_extent(is_true, cap: int = _SPILL_SCAN_CAP) -> int:
    """Return the largest offset (>=0) for which ``is_true(offset)`` holds.

    Precondition: ``is_true(0)`` is already known True (the anchor itself).
    Exponential probe (1, 2, 4, ...) brackets the True/False boundary, then a
    binary search narrows it down — same result as checking every offset
    0..cap linearly, but O(log cap) calls to ``is_true`` instead of O(cap).
    """
    lo, probe = 0, 1
    while probe <= cap and is_true(probe):
        lo = probe
        probe *= 2
    hi = min(probe, cap + 1)
    while hi - lo > 1:
        mid = (lo + hi) // 2
        if mid <= cap and is_true(mid):
            lo = mid
        else:
            hi = mid
    return lo


def _spill_range_address(anchor) -> str:
    """Return the address of the full spill range from an anchor cell.

    SpillingRange is not reliably accessible via pywin32 late-binding
    (returns None on some Excel builds despite HasSpill=True).
    Fallback: exponential-probe + binary-search outward from the anchor via
    HasSpill flags (same 1000-cell cap and result as the old linear scan,
    with O(log n) COM round-trips instead of O(n)).
    """
    # First try the COM property directly
    try:
        sr = anchor.SpillingRange
        if sr is not None:
            return sr.Address
    except Exception:
        pass

    ws = anchor.Parent
    anchor_row = anchor.Row
    anchor_col = anchor.Column

    def _row_has_spill(offset: int) -> bool:
        try:
            return bool(ws.Cells(anchor_row + offset, anchor_col).HasSpill)
        except Exception:
            return False

    def _col_has_spill(offset: int) -> bool:
        try:
            return bool(ws.Cells(anchor_row, anchor_col + offset).HasSpill)
        except Exception:
            return False

    try:
        row_extent = _scan_spill_extent(_row_has_spill)
    except Exception:
        row_extent = 0
    max_row = anchor_row + row_extent

    try:
        col_extent = _scan_spill_extent(_col_has_spill)
    except Exception:
        col_extent = 0
    max_col = anchor_col + col_extent

    end_cell = ws.Cells(max_row, max_col)
    return ws.Range(anchor, end_cell).Address


def _spill_parent_address(cell) -> str | None:
    """Return the anchor cell address if this cell is part of a foreign spill.

    SpillParent is also unreliable in late-binding; we try it but don't crash.
    """
    try:
        sp = cell.SpillParent
        if sp is not None:
            return sp.Address
    except Exception:
        pass
    return None


def _read_spill(
    range_str: str,
    sheet: str | None,
    workbook: str | None,
    offset: int,
    limit: int,
    value_mode: str = "typed",
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
    # Unlike _read's incidental spill metadata, this probe IS the pagination
    # boundary itself (total_rows/total_cols below come from it) — every
    # page call needs it, offset 0 included, since each MCP call is stateless
    # (no cross-call cache of a prior page's spill_addr). Nothing to trim here
    # without adding a cache, which is out of this pass's scope.
    spill_addr = _spill_range_address(anchor_cell)
    ws = anchor_cell.Parent
    spill_rng = ws.Range(spill_addr)

    total_rows = spill_rng.Rows.Count
    total_cols = spill_rng.Columns.Count
    if total_rows == 1 and total_cols == 1 and _read_values(spill_rng, value_mode) is None:
        return {
            "anchor": anchor_cell.Address,
            "spill_range": spill_addr,
            "values": [],
            "total_rows": 0,
            "has_more": False,
            "next_offset": 0,
        }
    if offset >= total_rows:
        rows: list = []
    else:
        r1, c1 = spill_rng.Row, spill_rng.Column
        last_row = r1 + min(offset + limit, total_rows) - 1
        page_rng = ws.Range(
            ws.Cells(r1 + offset, c1),
            ws.Cells(last_row, c1 + total_cols - 1),
        )
        rows = _normalize_rows(_read_values(page_rng, value_mode))
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
    if not values:
        raise ToolError("write requires a non-empty 2-D list of values.")
    rows = len(values)
    cols = max((len(r) if isinstance(r, (list, tuple)) else 1) for r in values)
    # Normalize to a rectangular tuple-of-tuples so COM accepts ragged lists.
    padded = tuple(
        tuple(r[c] if isinstance(r, (list, tuple)) and c < len(r) else None for c in range(cols))
        for r in values
    )
    rng = _resolve_range(range_str, sheet, workbook)
    try:
        # pywin32 dispatch quirk: Range.Resize(r,c) and Range.Offset(r,c) act as
        # indexed-property access → return a SINGLE offset cell (Count==1), not a
        # block.  Build the target block explicitly via Range(Cells, Cells) instead
        # — this is what xlwings does internally and is verified correct.
        ws = rng.Parent
        tl = rng.Cells(1, 1)
        r1, c1 = tl.Row, tl.Column
        target = ws.Range(ws.Cells(r1, c1), ws.Cells(r1 + rows - 1, c1 + cols - 1))
        target.Value = padded
        return {"written": {"rows": rows, "cols": cols, "range": target.Address}}
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
