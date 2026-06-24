"""Find / Replace / Count text across a range, sheet, or entire workbook.

Actions
-------
find    Locate all cells matching find_text.  Returns list of {cell, value}
        (capped at FIND_CAP=1000) plus total_found + truncated flag.
count   Same search loop, returns only the count.
replace Pre-count matching cells, call rng.Replace(), re-count remaining.
        Verify-effect: remaining_after must be 0 (or report if non-zero).

COM API (confirmed from Microsoft Learn 2026-06-24)
----------------------------------------------------
Range.Find(What, After, LookIn, LookAt, SearchOrder, MatchCase)
  - LookIn:      xlFormulas=-4123  xlValues=-4163
  - LookAt:      xlWhole=1  xlPart=2
  - SearchOrder: xlByRows=1  xlByColumns=2
  - Returns None when no match (Python win32com maps VBA Nothing → None).

Range.FindNext(After) → next matching Range or None.
Range.Replace(What, Replacement, LookAt, SearchOrder, MatchCase) → Boolean.
  Range.Replace does NOT return a count; count is obtained via the find loop.

Infinite-loop guard (CRITICAL)
------------------------------
rng.FindNext loops back to the beginning when it wraps. We save the .Address
of the FIRST found cell and stop as soon as FindNext returns that address
again (or returns None). Failing to do this causes an infinite loop.

scope resolution
----------------
- "range"    → resolved via _resolve_range(range, sheet, workbook)
- "sheet"    → ws.Cells  (entire sheet)
- "workbook" → iterate all worksheets, merge results

excel_guard: wrapped around Replace (mutation); find/count are read-only.
"""

from __future__ import annotations

from fastmcp.exceptions import ToolError

from ..session import ExcelSession, excel_guard

_session = ExcelSession()

# XlFindLookIn — confirmed from Microsoft Learn (excel.xlfindlookin)
_XL_FORMULAS = -4123   # xlFormulas
_XL_VALUES   = -4163   # xlValues

# XlLookAt — confirmed from Microsoft Learn (excel.xllookat)
_XL_WHOLE = 1   # xlWhole
_XL_PART  = 2   # xlPart

# XlSearchOrder — confirmed from Microsoft Learn (excel.xlsearchorder)
_XL_BY_ROWS = 1   # xlByRows

# Maximum cells returned in find result list (prevent oversized response)
_FIND_CAP = 1000

_LOOK_IN_MAP = {
    "formulas": _XL_FORMULAS,
    "values":   _XL_VALUES,
}


# ── Public entry point ─────────────────────────────────────────────────────────

def find_replace_action(
    action: str,
    find_text: str,
    replace_text: str | None = None,
    scope: str = "sheet",
    range: str | None = None,
    sheet: str | None = None,
    workbook: str | None = None,
    match_case: bool = False,
    match_whole_cell: bool = False,
    look_in: str = "formulas",
) -> dict:
    """Dispatch a find/replace/count action.

    Actions
    -------
    find
        Return all matching cell addresses + their current values.
        Results capped at 1000 entries (truncated flag set if more exist).
    count
        Return only the count of matching cells.
    replace
        Replace find_text with replace_text.  Requires replace_text.
        Reports cells_matched_before and remaining_after.

    Parameters
    ----------
    find_text        : text to search for.
    replace_text     : replacement text (required for replace action only).
    scope            : "range" | "sheet" (default) | "workbook".
    range            : cell/range address, required when scope="range".
    sheet            : sheet name; omit to use active sheet.
    workbook         : workbook name; omit to use active workbook.
    match_case       : True for case-sensitive search (default False).
    match_whole_cell : True to match whole cell only (default False, xlPart).
    look_in          : "formulas" (default) | "values".
    """
    valid = {"find", "count", "replace"}
    if action not in valid:
        raise ToolError(
            f"Unknown action '{action}'. Valid: {', '.join(sorted(valid))}."
        )
    if scope not in ("range", "sheet", "workbook"):
        raise ToolError(
            f"Unknown scope '{scope}'. Valid: range, sheet, workbook."
        )
    if look_in not in _LOOK_IN_MAP:
        raise ToolError(
            f"Unknown look_in '{look_in}'. Valid: formulas, values."
        )
    if action == "replace" and replace_text is None:
        raise ToolError("replace action requires 'replace_text'.")
    if scope == "range" and not range:
        raise ToolError("scope='range' requires the 'range' parameter.")

    return _session.run_com(
        _dispatch,
        action, find_text, replace_text,
        scope, range, sheet, workbook,
        match_case, match_whole_cell, look_in,
    )


# ── Worker-thread dispatcher ───────────────────────────────────────────────────

def _dispatch(
    action: str,
    find_text: str,
    replace_text: str | None,
    scope: str,
    range_str: str | None,
    sheet: str | None,
    workbook: str | None,
    match_case: bool,
    match_whole_cell: bool,
    look_in: str,
) -> dict:
    """Executed on the STA COM worker thread."""
    look_in_int  = _LOOK_IN_MAP[look_in]
    look_at_int  = _XL_WHOLE if match_whole_cell else _XL_PART

    # Resolve search scope(s): list of (range_obj, ws_name)
    scopes = _resolve_scopes(scope, range_str, sheet, workbook)

    app = scopes[0][0].Application

    if action == "count":
        return _count(scopes, find_text, look_in_int, look_at_int, match_case)

    if action == "find":
        return _find(scopes, find_text, look_in_int, look_at_int, match_case)

    # action == "replace"
    with excel_guard(app):
        return _replace(
            scopes, find_text, replace_text,
            look_in_int, look_at_int, match_case, app,
        )


# ── Scope resolution ──────────────────────────────────────────────────────────

def _resolve_scopes(
    scope: str,
    range_str: str | None,
    sheet: str | None,
    workbook: str | None,
) -> list[tuple]:
    """Return list of (rng, sheet_name) pairs for the requested scope.

    - scope="range"    → single resolved range
    - scope="sheet"    → ws.Cells (full sheet)
    - scope="workbook" → ws.Cells for every sheet in the workbook
    """
    if scope == "range":
        rng = _resolve_range(range_str, sheet, workbook)
        return [(rng, rng.Worksheet.Name)]

    if scope == "sheet":
        ws = _session.get_sheet(sheet, workbook)
        return [(ws.Cells, ws.Name)]

    # scope == "workbook"
    wb = _session.get_workbook(workbook)
    result = []
    try:
        count = wb.Sheets.Count
        for i in range(1, count + 1):
            ws = wb.Sheets(i)
            result.append((ws.Cells, ws.Name))
    except Exception as e:
        raise _session.wrap(e, "Failed to enumerate workbook sheets")
    return result


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


# ── Find loop helper ──────────────────────────────────────────────────────────

def _find_all(rng, find_text: str, look_in_int: int, look_at_int: int, match_case: bool):
    """Iterate all matching cells in rng using Find + FindNext loop.

    Returns list of (address_str, value_str).

    CRITICAL infinite-loop guard: save the first hit's .Address; stop when
    FindNext returns that same address again (wrap detection).
    Range.Find returns None when nothing found — do NOT crash on None.
    """
    hits: list[tuple[str, str]] = []
    try:
        cell = rng.Find(
            What=find_text,
            LookIn=look_in_int,
            LookAt=look_at_int,
            SearchOrder=_XL_BY_ROWS,
            MatchCase=match_case,
        )
        if cell is None:
            return hits

        first_address = cell.Address
        while True:
            try:
                val = str(cell.Value) if cell.Value is not None else ""
            except Exception:
                val = ""
            hits.append((cell.Address, val))

            cell = rng.FindNext(cell)
            if cell is None or cell.Address == first_address:
                break

    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "Find loop failed")

    return hits


# ── Action implementations ────────────────────────────────────────────────────

def _find(
    scopes: list[tuple],
    find_text: str,
    look_in_int: int,
    look_at_int: int,
    match_case: bool,
) -> dict:
    """Locate all matching cells across all scopes."""
    all_hits: list[dict] = []
    for rng, ws_name in scopes:
        hits = _find_all(rng, find_text, look_in_int, look_at_int, match_case)
        for addr, val in hits:
            all_hits.append({"sheet": ws_name, "cell": addr, "value": val})

    total = len(all_hits)
    truncated = total > _FIND_CAP
    return {
        "find_replace": "find",
        "applied": {
            "find_text": find_text,
            "total_found": total,
            "truncated": truncated,
            "matches": all_hits[:_FIND_CAP],
        },
    }


def _count(
    scopes: list[tuple],
    find_text: str,
    look_in_int: int,
    look_at_int: int,
    match_case: bool,
) -> dict:
    """Return only the count of matching cells across all scopes."""
    total = 0
    for rng, _ws_name in scopes:
        hits = _find_all(rng, find_text, look_in_int, look_at_int, match_case)
        total += len(hits)

    return {
        "find_replace": "count",
        "applied": {
            "find_text": find_text,
            "count": total,
        },
    }


def _replace(
    scopes: list[tuple],
    find_text: str,
    replace_text: str,
    look_in_int: int,
    look_at_int: int,
    match_case: bool,
    app,
) -> dict:
    """Replace all occurrences of find_text with replace_text.

    Steps:
    1. Pre-count matching cells (find loop).
    2. Call rng.Replace() on each scope range.
    3. VERIFY-EFFECT: re-count remaining matches — must be 0.
       Report cells_matched_before + remaining_after.

    Range.Replace returns Boolean (not a count) — we use the find loop
    both before and after to measure the actual effect.
    """
    # Step 1: pre-count
    cells_matched_before = 0
    for rng, _ws_name in scopes:
        hits = _find_all(rng, find_text, look_in_int, look_at_int, match_case)
        cells_matched_before += len(hits)

    # Step 2: replace
    try:
        for rng, _ws_name in scopes:
            rng.Replace(
                What=find_text,
                Replacement=replace_text,
                LookAt=look_at_int,
                SearchOrder=_XL_BY_ROWS,
                MatchCase=match_case,
            )
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "Replace failed")

    # Step 3: VERIFY-EFFECT — re-count; should be 0
    remaining_after = 0
    for rng, _ws_name in scopes:
        hits = _find_all(rng, find_text, look_in_int, look_at_int, match_case)
        remaining_after += len(hits)

    return {
        "find_replace": "replace",
        "applied": {
            "find_text": find_text,
            "replace_text": replace_text,
            "cells_matched_before": cells_matched_before,
            "remaining_after": remaining_after,
            "fully_replaced": remaining_after == 0,
        },
    }
