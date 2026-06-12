"""Workbook-level operations: list, info, open, save, close."""

from __future__ import annotations

from fastmcp.exceptions import ToolError

from ..session import ExcelSession

_session = ExcelSession()


def workbook_action(
    action: str,
    workbook: str | None = None,
    path: str | None = None,
) -> dict:
    """Dispatch a workbook action.

    Actions
    -------
    list
        Return names + active-flag for every open workbook.
    info
        Return sheet / table / query / defined-name counts for a workbook.
    open
        Open a workbook from a file path (requires ``path``).
    save
        Save the workbook (Ctrl+S equivalent).
    close
        Close without saving. Does NOT prompt — caller's responsibility.
    """
    if action == "list":
        return _list()
    if action == "info":
        return _info(workbook)
    if action == "open":
        if not path:
            raise ToolError("action='open' requires the 'path' parameter.")
        return _open(path)
    if action == "save":
        return _save(workbook)
    if action == "close":
        return _close(workbook)
    raise ToolError(
        f"Unknown action '{action}'. Valid: list, info, open, save, close."
    )


def _list() -> dict:
    app = _session.get_app()
    active_name = app.ActiveWorkbook.Name if app.ActiveWorkbook else None
    workbooks = [
        {
            "name": app.Workbooks(i + 1).Name,
            "active": app.Workbooks(i + 1).Name == active_name,
        }
        for i in range(app.Workbooks.Count)
    ]
    return {"workbooks": workbooks, "count": len(workbooks)}


def _info(workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    # Count sheets
    sheet_names = [wb.Sheets(i + 1).Name for i in range(wb.Sheets.Count)]
    # Count tables (ListObjects) across all sheets
    table_count = sum(
        wb.Sheets(i + 1).ListObjects.Count for i in range(wb.Sheets.Count)
    )
    # Count queries
    query_count = wb.Queries.Count
    # Count defined names
    name_count = wb.Names.Count
    return {
        "name": wb.Name,
        "path": wb.FullName,
        "sheets": sheet_names,
        "sheet_count": len(sheet_names),
        "table_count": table_count,
        "query_count": query_count,
        "defined_name_count": name_count,
        "saved": wb.Saved,
    }


def _open(path: str) -> dict:
    app = _session.get_app()
    try:
        wb = app.Workbooks.Open(path)
        return {"opened": wb.Name, "path": wb.FullName}
    except Exception as e:
        raise _session.wrap(e, f"Cannot open '{path}'")


def _save(workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    try:
        wb.Save()
        return {"saved": wb.Name}
    except Exception as e:
        raise _session.wrap(e, "Save failed")


def _close(workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    name = wb.Name
    try:
        wb.Close(SaveChanges=False)
        return {"closed": name}
    except Exception as e:
        raise _session.wrap(e, "Close failed")
