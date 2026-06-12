"""Sheet-level operations: list, add, rename, delete."""

from __future__ import annotations

from fastmcp.exceptions import ToolError

from ..session import ExcelSession

_session = ExcelSession()


def sheet_action(
    action: str,
    name: str | None = None,
    workbook: str | None = None,
    new_name: str | None = None,
) -> dict:
    """Dispatch a sheet action.

    Actions
    -------
    list
        Return all sheet names in the workbook.
    add
        Add a new sheet. ``name`` sets the sheet name (optional; Excel auto-names if omitted).
    rename
        Rename an existing sheet. Requires ``name`` (current) and ``new_name``.
    delete
        Delete a sheet by name. Irreversible.
    """
    if action == "list":
        return _list(workbook)
    if action == "add":
        return _add(name, workbook)
    if action == "rename":
        if not name or not new_name:
            raise ToolError("action='rename' requires 'name' and 'new_name'.")
        return _rename(name, new_name, workbook)
    if action == "delete":
        if not name:
            raise ToolError("action='delete' requires 'name'.")
        return _delete(name, workbook)
    raise ToolError(
        f"Unknown action '{action}'. Valid: list, add, rename, delete."
    )


def _list(workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    active = wb.ActiveSheet.Name if wb.ActiveSheet else None
    sheets = [
        {"name": wb.Sheets(i + 1).Name, "active": wb.Sheets(i + 1).Name == active}
        for i in range(wb.Sheets.Count)
    ]
    return {"sheets": sheets, "count": len(sheets)}


def _add(name: str | None, workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    try:
        # Add after last sheet
        ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        if name:
            ws.Name = name
        return {"added": ws.Name}
    except Exception as e:
        raise _session.wrap(e, "Add sheet failed")


def _rename(name: str, new_name: str, workbook: str | None) -> dict:
    ws = _session.get_sheet(name, workbook)
    try:
        ws.Name = new_name
        return {"renamed": {"from": name, "to": new_name}}
    except Exception as e:
        raise _session.wrap(e, "Rename sheet failed")


def _delete(name: str, workbook: str | None) -> dict:
    ws = _session.get_sheet(name, workbook)
    wb = ws.Parent
    app = wb.Application
    app.DisplayAlerts = False
    try:
        ws.Delete()
    except Exception as e:
        app.DisplayAlerts = True
        raise _session.wrap(e, f"Delete sheet '{name}' failed")
    app.DisplayAlerts = True
    return {"deleted": name}
