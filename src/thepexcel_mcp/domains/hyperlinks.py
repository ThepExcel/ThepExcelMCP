"""Worksheet hyperlink management.

Supports adding URL, internal (sheet/cell), email, and file hyperlinks to cells;
listing all hyperlinks on a sheet; and deleting hyperlinks from a range.

COM API
-------
ws.Hyperlinks.Add(Anchor, Address, SubAddress, ScreenTip, TextToDisplay)
  - Anchor  : Required — a Range object (ws.Range(cell)), NOT a string.
  - Address : Required — the URL/path/mailto string. For internal links MUST be
              an empty string "" (some Excel builds reject None; omitting it
              causes a COM TypeError because the parameter is marked Required).
  - SubAddress, ScreenTip, TextToDisplay : Optional Variant.

Read-back properties (Hyperlink object)
  .Address        — the URL/path/mailto string
  .SubAddress     — the in-document anchor (e.g. "Sheet2!A1")
  .TextToDisplay  — the visible cell text
  .ScreenTip      — the hover tooltip
  .Range.Address  — the cell address of the anchor (absolute $A$1 form)

Delete
  ws.Range(range).Hyperlinks.Delete() removes all hyperlinks in the range
  but keeps the cell text and value intact.

Gotchas
-------
- Anchor MUST be ws.Range(cell), not a bare string.
- For internal links, Address must be "" (empty string), not None or omitted.
- Passing SubAddress="" for URL hyperlinks is harmless; Excel stores it as "".
- pywin32 supports keyword args matching VBA named params exactly.
- Wrap the Add call in excel_guard to suppress modal dialogs on the STA thread.
- The count of hyperlinks before Delete() is the only way to report removals
  (ws.Range(range).Hyperlinks.Count before, then .Delete()).
- TextToDisplay, when not supplied by the caller, defaults to the cell address
  in Excel's internal representation — we pass the caller's value or omit the
  keyword entirely (omitting passes no argument, Excel uses its own default).
"""

from __future__ import annotations

from fastmcp.exceptions import ToolError

from ..session import ExcelSession, excel_guard

_session = ExcelSession()

# Valid link types
_LINK_TYPES = {"url", "internal", "email", "file"}


# ── Public entry point ─────────────────────────────────────────────────────────

def hyperlink_action(
    action: str,
    sheet: str | None = None,
    workbook: str | None = None,
    # add
    cell: str | None = None,
    link_type: str | None = None,
    target: str | None = None,
    sub_address: str | None = None,
    text_to_display: str | None = None,
    screen_tip: str | None = None,
    # delete
    range: str | None = None,
) -> dict:
    """Dispatch a hyperlink action.

    Actions
    -------
    add
        Add a hyperlink to *cell* on the target sheet.
        link_type: "url" | "internal" | "email" | "file"
        target   : the URL, sheet!cell ref, email address, or file path.
        sub_address     : (optional) in-document anchor; used as-is for
                          'internal' when target already contains the sheet ref.
        text_to_display : (optional) visible cell text.
        screen_tip      : (optional) hover tooltip.
    list
        Return all hyperlinks on the sheet as a list of dicts with keys:
        anchor, address, sub_address, text, screen_tip.
    delete
        Clear all hyperlinks from *range* (keeps cell text/value).
        Reports the count of hyperlinks removed.
    """
    valid = {"add", "list", "delete"}
    if action not in valid:
        raise ToolError(
            f"Unknown action '{action}'. Valid: {', '.join(sorted(valid))}."
        )
    return _session.run_com(
        _dispatch,
        action, sheet, workbook,
        cell, link_type, target, sub_address,
        text_to_display, screen_tip, range,
    )


# ── Worker-thread dispatcher ───────────────────────────────────────────────────

def _dispatch(
    action: str,
    sheet: str | None,
    workbook: str | None,
    cell: str | None,
    link_type: str | None,
    target: str | None,
    sub_address: str | None,
    text_to_display: str | None,
    screen_tip: str | None,
    range_str: str | None,
) -> dict:
    """Executed on the STA COM worker thread."""
    ws = _session.get_sheet(sheet, workbook)
    app = ws.Application

    with excel_guard(app):
        if action == "add":
            return _add(ws, cell, link_type, target, sub_address,
                        text_to_display, screen_tip)
        if action == "list":
            return _list(ws)
        # action == "delete"
        return _delete(ws, range_str)


# ── Action implementations ────────────────────────────────────────────────────

def _add(
    ws,
    cell: str | None,
    link_type: str | None,
    target: str | None,
    sub_address: str | None,
    text_to_display: str | None,
    screen_tip: str | None,
) -> dict:
    """Add a hyperlink to *cell*."""
    if not cell:
        raise ToolError("add action requires 'cell' (anchor address, e.g. 'A1').")
    if link_type is None:
        raise ToolError(
            f"add action requires 'link_type'. Valid: {', '.join(sorted(_LINK_TYPES))}."
        )
    if link_type not in _LINK_TYPES:
        raise ToolError(
            f"Unknown link_type '{link_type}'. Valid: {', '.join(sorted(_LINK_TYPES))}."
        )
    if not target:
        raise ToolError("add action requires 'target' (the URL, cell ref, email, or path).")

    # Resolve the anchor Range object (must be a COM Range, not a string)
    try:
        anchor = ws.Range(cell)
    except Exception as e:
        raise ToolError(f"Invalid cell reference '{cell}': {e}")

    # Build Address and SubAddress based on link_type
    if link_type == "url":
        address = target
        sub_addr = sub_address if sub_address is not None else ""
    elif link_type == "internal":
        # Address MUST be "" for internal links; SubAddress carries the sheet!cell
        address = ""
        sub_addr = target if sub_address is None else sub_address
    elif link_type == "email":
        # Normalize: add mailto: prefix if missing
        address = target if target.startswith("mailto:") else f"mailto:{target}"
        sub_addr = sub_address if sub_address is not None else ""
    else:
        # file
        address = target
        sub_addr = sub_address if sub_address is not None else ""

    try:
        # Build kwargs dict so we can conditionally include optional args.
        # Always pass SubAddress — Excel accepts "" fine and it avoids positional
        # ambiguity in pywin32 when ScreenTip/TextToDisplay are also supplied.
        kwargs: dict = dict(
            Anchor=anchor,
            Address=address,
            SubAddress=sub_addr,
        )
        if screen_tip is not None:
            kwargs["ScreenTip"] = screen_tip
        if text_to_display is not None:
            kwargs["TextToDisplay"] = text_to_display

        hl = ws.Hyperlinks.Add(**kwargs)

        # VERIFY EFFECT: read back the created hyperlink's properties from COM,
        # including the canonical anchor address (hl.Range.Address, not the
        # caller-supplied string which may be relative or shorthand).
        readback_cell = hl.Range.Address
        readback_address = hl.Address
        readback_sub = hl.SubAddress
        readback_text = hl.TextToDisplay
        readback_tip = hl.ScreenTip
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"Hyperlinks.Add failed on '{cell}'")

    return {
        "hyperlink": "add",
        "sheet": ws.Name,
        "applied": {
            "cell": readback_cell,
            "link_type": link_type,
            "address": readback_address,
            "sub_address": readback_sub,
            "text_to_display": readback_text,
            "screen_tip": readback_tip,
        },
    }


def _list(ws) -> dict:
    """Return all hyperlinks on the sheet."""
    try:
        hyperlinks = ws.Hyperlinks
        count = hyperlinks.Count
        items = []
        for i in range(1, count + 1):
            hl = hyperlinks(i)
            items.append({
                "anchor": hl.Range.Address,
                "address": hl.Address,
                "sub_address": hl.SubAddress,
                "text": hl.TextToDisplay,
                "screen_tip": hl.ScreenTip,
            })
    except Exception as e:
        raise _session.wrap(e, "Hyperlinks.list failed")
    return {
        "hyperlink": "list",
        "sheet": ws.Name,
        "count": count,
        "hyperlinks": items,
    }


def _delete(ws, range_str: str | None) -> dict:
    """Delete all hyperlinks in *range_str* (keeps cell text)."""
    if not range_str:
        raise ToolError(
            "delete action requires 'range' (e.g. 'A1', 'A1:C5', 'Sheet2!A1:B3')."
        )
    # Resolve range — support sheet-qualified refs (Sheet!A1:B3)
    if "!" in range_str:
        sheet_part, cell_part = range_str.split("!", 1)
        target_ws = _session.get_sheet(sheet_part.strip("'"), ws.Parent.Name)
        try:
            rng = target_ws.Range(cell_part)
        except Exception as e:
            raise _session.wrap(e, f"Invalid range '{cell_part}' on sheet '{sheet_part}'")
        actual_ws = target_ws
    else:
        try:
            rng = ws.Range(range_str)
        except Exception as e:
            raise _session.wrap(e, f"Invalid range '{range_str}'")
        actual_ws = ws

    try:
        count_before = rng.Hyperlinks.Count
        rng.Hyperlinks.Delete()
    except Exception as e:
        raise _session.wrap(e, f"Hyperlinks.Delete failed on range '{range_str}'")

    return {
        "hyperlink": "delete",
        "sheet": actual_ws.Name,
        "applied": {
            "range": rng.Address,
            "removed": count_before,
        },
    }
