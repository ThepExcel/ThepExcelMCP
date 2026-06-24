"""Worksheet and workbook PROTECTION + cell lock state.

Operations
----------
protect_sheet       ws.Protect(Password, DrawingObjects, Contents, Scenarios, + allow flags)
unprotect_sheet     ws.Unprotect(Password)          — catches wrong-password 1004
protect_workbook    wb.Protect(Password, Structure, Windows)
unprotect_workbook  wb.Unprotect(Password)          — catches wrong-password 1004
set_locked          ws.Range(range).Locked / FormulaHidden
status              read-only snapshot of all protection flags

COM API (confirmed from Microsoft Learn)
----------------------------------------
Worksheet.Protect(*Password*, *DrawingObjects*, *Contents*, *Scenarios*,
                  *UserInterfaceOnly*, *AllowFormattingCells*, ...)
    All params Optional Variant.  Password omitted ≡ no-password.
    DrawingObjects / Contents / Scenarios default True.
    All Allow* flags default False.
  Read-back: ws.ProtectContents (bool), ws.ProtectDrawingObjects (bool),
             ws.ProtectScenarios (bool), ws.Protection.Allow* (bool)

Worksheet.Unprotect(*Password*) — raises COM error 1004 on wrong password.
  Read-back: ws.ProtectContents == False

Workbook.Protect(*Password*, *Structure*, *Windows*)
    Structure: True → sheet order/add/delete locked. Default False.
    Windows:   True → workbook window layout locked. Default False.
  Read-back: wb.ProtectStructure (bool), wb.ProtectWindows (bool)

Workbook.Unprotect(*Password*) — raises COM error 1004 on wrong password.
  Read-back: wb.ProtectStructure == False

Range.Locked   (bool) — cell-level locked flag; only effective under sheet protection.
Range.FormulaHidden (bool) — hide formula bar content when sheet is protected.
  Read-back: rng.Locked, rng.FormulaHidden

GOTCHAS
-------
* Password=None or "" → omit the keyword arg entirely (passing Password="" to
  Workbook.Protect raises COM 1004 in some Excel versions). We use conditional
  kwargs to avoid passing it when None.
* Wrong password to Unprotect raises COM error 1004 (same code as other errors).
  We wrap the Unprotect call in excel_guard to suppress any dialog, then convert
  to an actionable ToolError mentioning "wrong password".
* Locked / FormulaHidden are meaningless until the sheet is protected — we
  document this in the action return value rather than blocking the call, because
  cells are typically pre-configured BEFORE calling protect_sheet.
* Wrap ALL mutations in excel_guard(app) to suppress Excel's confirmation dialogs.
* Falsy-value guard: use `if x is None` not `if not x` — False and 0 are valid
  protect flag values.
"""

from __future__ import annotations

from fastmcp.exceptions import ToolError

from ..session import ExcelSession, excel_guard

_session = ExcelSession()

# Allow-flag parameter names for protect_sheet (keep in deterministic order)
_ALLOW_FLAGS = [
    "AllowFormattingCells",
    "AllowFormattingColumns",
    "AllowFormattingRows",
    "AllowInsertingColumns",
    "AllowInsertingRows",
    "AllowInsertingHyperlinks",
    "AllowDeletingColumns",
    "AllowDeletingRows",
    "AllowSorting",
    "AllowFiltering",
    "AllowUsingPivotTables",
]


# ── Public entry point ─────────────────────────────────────────────────────────

def protection_action(
    action: str,
    # sheet / workbook target
    sheet: str | None = None,
    workbook: str | None = None,
    # password (shared by protect/unprotect actions)
    password: str | None = None,
    # protect_sheet structural flags
    drawing_objects: bool = True,
    contents: bool = True,
    scenarios: bool = True,
    # protect_sheet allow flags (pass-through as dict or individual bools)
    allow: dict | None = None,
    allow_formatting_cells: bool = False,
    allow_formatting_columns: bool = False,
    allow_formatting_rows: bool = False,
    allow_inserting_columns: bool = False,
    allow_inserting_rows: bool = False,
    allow_inserting_hyperlinks: bool = False,
    allow_deleting_columns: bool = False,
    allow_deleting_rows: bool = False,
    allow_sorting: bool = False,
    allow_filtering: bool = False,
    allow_using_pivot_tables: bool = False,
    # protect_workbook flags
    structure: bool = True,
    windows: bool = False,
    # set_locked
    range: str | None = None,
    locked: bool | None = None,
    hidden: bool | None = None,
) -> dict:
    """Dispatch a protection action.

    Actions
    -------
    protect_sheet       Protect the worksheet (optionally with password + allow flags).
    unprotect_sheet     Remove sheet protection (supply password if one was set).
    protect_workbook    Protect workbook structure/windows.
    unprotect_workbook  Remove workbook protection.
    set_locked          Set Locked and/or FormulaHidden on a cell range.
                        Note: these flags only take effect while the sheet is protected.
    status              Return current protection status (sheet + workbook).
    """
    valid = {
        "protect_sheet",
        "unprotect_sheet",
        "protect_workbook",
        "unprotect_workbook",
        "set_locked",
        "status",
    }
    if action not in valid:
        raise ToolError(
            f"Unknown action '{action}'. Valid: {', '.join(sorted(valid))}."
        )

    # Build the allow-flags dict: 'allow' dict overrides individual bool params
    resolved_allow: dict[str, bool] = {}
    if allow is not None:
        # Caller supplied a dict — dict keys win; flags absent from the dict
        # fall back to the individual bool params (so mixed use works).
        _individual = {
            "AllowFormattingCells":      allow_formatting_cells,
            "AllowFormattingColumns":    allow_formatting_columns,
            "AllowFormattingRows":       allow_formatting_rows,
            "AllowInsertingColumns":     allow_inserting_columns,
            "AllowInsertingRows":        allow_inserting_rows,
            "AllowInsertingHyperlinks":  allow_inserting_hyperlinks,
            "AllowDeletingColumns":      allow_deleting_columns,
            "AllowDeletingRows":         allow_deleting_rows,
            "AllowSorting":              allow_sorting,
            "AllowFiltering":            allow_filtering,
            "AllowUsingPivotTables":     allow_using_pivot_tables,
        }
        for flag in _ALLOW_FLAGS:
            snake = _flag_to_snake(flag)
            if flag in allow:
                resolved_allow[flag] = bool(allow[flag])
            elif snake in allow:
                resolved_allow[flag] = bool(allow[snake])
            else:
                # Genuinely absent from the dict — fall back to individual param
                resolved_allow[flag] = _individual[flag]
    else:
        # Use individual bool params
        resolved_allow = {
            "AllowFormattingCells":      allow_formatting_cells,
            "AllowFormattingColumns":    allow_formatting_columns,
            "AllowFormattingRows":       allow_formatting_rows,
            "AllowInsertingColumns":     allow_inserting_columns,
            "AllowInsertingRows":        allow_inserting_rows,
            "AllowInsertingHyperlinks":  allow_inserting_hyperlinks,
            "AllowDeletingColumns":      allow_deleting_columns,
            "AllowDeletingRows":         allow_deleting_rows,
            "AllowSorting":              allow_sorting,
            "AllowFiltering":            allow_filtering,
            "AllowUsingPivotTables":     allow_using_pivot_tables,
        }

    return _session.run_com(
        _dispatch,
        action, sheet, workbook, password,
        drawing_objects, contents, scenarios,
        resolved_allow,
        structure, windows,
        range, locked, hidden,
    )


def _flag_to_snake(flag: str) -> str:
    """AllowFormattingCells → allow_formatting_cells (for dict lookup)."""
    # strip leading 'Allow', then insert _ before each uppercase
    import re
    name = flag[len("Allow"):]
    snake = re.sub(r"(?<!^)(?=[A-Z])", "_", name).lower()
    return "allow_" + snake


# ── Worker-thread dispatcher ───────────────────────────────────────────────────

def _dispatch(
    action: str,
    sheet: str | None,
    workbook: str | None,
    password: str | None,
    drawing_objects: bool,
    contents: bool,
    scenarios: bool,
    resolved_allow: dict[str, bool],
    structure: bool,
    windows: bool,
    range_str: str | None,
    locked: bool | None,
    hidden: bool | None,
) -> dict:
    """Executed on the STA COM worker thread."""
    ws = _session.get_sheet(sheet, workbook)
    wb = ws.Parent
    app = ws.Application

    with excel_guard(app):
        if action == "protect_sheet":
            return _protect_sheet(ws, password, drawing_objects, contents, scenarios, resolved_allow)
        if action == "unprotect_sheet":
            return _unprotect_sheet(ws, password)
        if action == "protect_workbook":
            return _protect_workbook(wb, password, structure, windows)
        if action == "unprotect_workbook":
            return _unprotect_workbook(wb, password)
        if action == "set_locked":
            return _set_locked(ws, range_str, locked, hidden)
        # action == "status"
        return _status(ws, wb)


# ── Action implementations ────────────────────────────────────────────────────

def _protect_sheet(ws, password: str | None, drawing_objects: bool, contents: bool,
                   scenarios: bool, allow_flags: dict[str, bool]) -> dict:
    """Protect a worksheet.

    Password is omitted from the COM call when None (no password), because
    passing Password="" can behave differently from omitting it in some Excel
    versions.
    """
    try:
        kwargs: dict = {
            "DrawingObjects": drawing_objects,
            "Contents":       contents,
            "Scenarios":      scenarios,
        }
        # Merge allow flags
        kwargs.update(allow_flags)
        if password is not None:
            kwargs["Password"] = password
        ws.Protect(**kwargs)
        # VERIFY EFFECT
        if not ws.ProtectContents:
            raise ToolError(
                "protect_sheet: ws.ProtectContents is still False after Protect() — "
                "check that Contents=True was passed."
            )
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "protect_sheet failed")

    applied: dict = {
        "drawing_objects": drawing_objects,
        "contents":        contents,
        "scenarios":       scenarios,
    }
    applied.update({k: v for k, v in allow_flags.items() if v})  # include only True flags
    if password is not None:
        applied["password_set"] = True
    return {
        "protection": "protect_sheet",
        "sheet": ws.Name,
        "applied": applied,
        "verify": {"ProtectContents": ws.ProtectContents},
    }


def _unprotect_sheet(ws, password: str | None) -> dict:
    """Remove sheet protection.

    COM raises error 1004 on wrong password; we catch it and convert to
    an actionable ToolError. excel_guard is already active (from _dispatch)
    so Excel's password-mismatch dialog is suppressed.
    """
    try:
        if password is not None:
            ws.Unprotect(Password=password)
        else:
            ws.Unprotect()
        # VERIFY EFFECT
        if ws.ProtectContents:
            raise ToolError(
                "unprotect_sheet: ws.ProtectContents is still True after Unprotect(). "
                "This sheet may require a password."
            )
    except ToolError:
        raise
    except Exception as e:
        msg = str(e)
        if "1004" in msg or "password" in msg.lower() or "protected" in msg.lower():
            raise ToolError(
                "unprotect_sheet: wrong password or the sheet is not protected. "
                f"Original error: {e}"
            )
        raise _session.wrap(e, "unprotect_sheet failed")
    return {
        "protection": "unprotect_sheet",
        "sheet": ws.Name,
        "applied": {"protected": False},
        "verify": {"ProtectContents": ws.ProtectContents},
    }


def _protect_workbook(wb, password: str | None, structure: bool, windows: bool) -> dict:
    """Protect workbook structure and/or windows."""
    try:
        kwargs: dict = {
            "Structure": structure,
            "Windows":   windows,
        }
        if password is not None:
            kwargs["Password"] = password
        wb.Protect(**kwargs)
        # VERIFY EFFECT — only meaningful when structure=True was requested
        if structure and not wb.ProtectStructure:
            raise ToolError(
                "protect_workbook: wb.ProtectStructure is still False after Protect(). "
                "Check that Structure=True was passed."
            )
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "protect_workbook failed")

    applied: dict = {"structure": structure, "windows": windows}
    if password is not None:
        applied["password_set"] = True
    return {
        "protection": "protect_workbook",
        "workbook": wb.Name,
        "applied": applied,
        "verify": {
            "ProtectStructure": wb.ProtectStructure,
            "ProtectWindows":   wb.ProtectWindows,
        },
    }


def _unprotect_workbook(wb, password: str | None) -> dict:
    """Remove workbook protection."""
    try:
        if password is not None:
            wb.Unprotect(Password=password)
        else:
            wb.Unprotect()
        # VERIFY EFFECT
        if wb.ProtectStructure:
            raise ToolError(
                "unprotect_workbook: wb.ProtectStructure is still True after Unprotect(). "
                "This workbook may require a password."
            )
    except ToolError:
        raise
    except Exception as e:
        msg = str(e)
        if "1004" in msg or "password" in msg.lower() or "protected" in msg.lower():
            raise ToolError(
                "unprotect_workbook: wrong password or the workbook is not protected. "
                f"Original error: {e}"
            )
        raise _session.wrap(e, "unprotect_workbook failed")
    return {
        "protection": "unprotect_workbook",
        "workbook": wb.Name,
        "applied": {"protected": False},
        "verify": {"ProtectStructure": wb.ProtectStructure},
    }


def _set_locked(ws, range_str: str | None, locked: bool | None, hidden: bool | None) -> dict:
    """Set Locked and/or FormulaHidden on a range.

    NOTE: Locked and FormulaHidden only take effect while the sheet is
    protected. Configure these flags BEFORE calling protect_sheet.
    """
    if range_str is None:
        raise ToolError("set_locked action requires 'range' (e.g. 'A1:B10').")
    if locked is None and hidden is None:
        raise ToolError(
            "set_locked action requires at least one of 'locked' (bool) "
            "or 'hidden' (bool)."
        )
    try:
        rng = ws.Range(range_str)
    except Exception as e:
        raise ToolError(f"Invalid range '{range_str}': {e}")

    try:
        if locked is not None:
            rng.Locked = bool(locked)
        if hidden is not None:
            rng.FormulaHidden = bool(hidden)
    except Exception as e:
        raise _session.wrap(e, "set_locked failed")

    # VERIFY EFFECT — read back from COM and raise if assignment was silently ignored
    applied: dict = {"range": rng.Address}
    verify: dict = {}
    if locked is not None:
        actual_locked = rng.Locked
        applied["locked"] = locked
        verify["Locked"] = actual_locked
        if actual_locked != locked:
            raise ToolError(
                f"set_locked: rng.Locked read-back is {actual_locked!r} "
                f"but {locked!r} was requested. The sheet may be protected — "
                "call unprotect_sheet first, then set_locked."
            )
    if hidden is not None:
        actual_hidden = rng.FormulaHidden
        applied["formula_hidden"] = hidden
        verify["FormulaHidden"] = actual_hidden
        if actual_hidden != hidden:
            raise ToolError(
                f"set_locked: rng.FormulaHidden read-back is {actual_hidden!r} "
                f"but {hidden!r} was requested. The sheet may be protected — "
                "call unprotect_sheet first, then set_locked."
            )

    applied["note"] = (
        "Locked/FormulaHidden only take effect while the sheet is protected. "
        "Call protect_sheet after configuring cell lock state."
    )
    return {
        "protection": "set_locked",
        "sheet": ws.Name,
        "applied": applied,
        "verify": verify,
    }


def _status(ws, wb) -> dict:
    """Return current protection flags for the sheet and workbook."""
    try:
        sheet_protected   = ws.ProtectContents
        drawing_objects   = ws.ProtectDrawingObjects
        scenarios         = ws.ProtectScenarios
        wb_structure      = wb.ProtectStructure
        wb_windows        = wb.ProtectWindows
    except Exception as e:
        raise _session.wrap(e, "status read failed")
    return {
        "protection": "status",
        "sheet": ws.Name,
        "workbook": wb.Name,
        "applied": {
            "sheet_protected":     sheet_protected,
            "drawing_objects":     drawing_objects,
            "scenarios":           scenarios,
            "workbook_structure":  wb_structure,
            "workbook_windows":    wb_windows,
        },
    }
