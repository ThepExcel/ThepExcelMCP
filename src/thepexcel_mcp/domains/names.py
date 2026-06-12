"""Named ranges and LAMBDA formula names.

Excel's Names collection covers both:
  - Named ranges/constants: Name.RefersTo = "=Sheet1!$A$1:$B$10"
  - Named LAMBDA formulas: Name.RefersTo starts with "=LAMBDA("

LAMBDA authoring example
------------------------
To define a named LAMBDA that doubles its input:

    excel_name(action="set",
               name="DOUBLE",
               refers_to="=LAMBDA(x, x*2)")

Once defined, use it in cells: =DOUBLE(A1)

LAMBDA formulas with multiple params:

    excel_name(action="set",
               name="TAX_CALC",
               refers_to='=LAMBDA(amount, rate, amount * rate / 100)')

Scope
-----
Workbook-scoped names (scope="workbook") are visible everywhere.
Sheet-scoped names (scope="SheetName") are local to that sheet and take
precedence over workbook-scoped names with the same name on that sheet.
Use sheet scope to shadow a global name on a specific sheet.
"""

from __future__ import annotations

from fastmcp.exceptions import ToolError

from ..session import ExcelSession

_session = ExcelSession()


def name_action(
    action: str,
    workbook: str | None = None,
    name: str | None = None,
    refers_to: str | None = None,
    scope: str | None = None,
) -> dict:
    """Dispatch a named-range / LAMBDA action.

    Parameters
    ----------
    action : str
        One of: ``list``, ``get``, ``set``, ``delete``.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    name : str, optional
        Name to get/set/delete. Case-insensitive.
    refers_to : str, optional
        ``set`` — the definition string. Must start with ``=``.
        For a named range: ``"=Sheet1!$A$1:$B$10"``
        For a LAMBDA: ``"=LAMBDA(x, x*2)"``
        For a constant: ``"=42"`` or ``"=\"hello\""``
    scope : str, optional
        ``"workbook"`` (default) for workbook-level scope, or a sheet name
        for sheet-local scope. Sheet-local names take precedence on their sheet.

    Actions
    -------
    list
        All defined names with: name, refers_to, scope, is_lambda.
        ``is_lambda`` is True when RefersTo starts with ``=LAMBDA(``.
        Example: ``excel_name(action="list")``
    get
        Return details for a single name. Requires ``name``.
        Example: ``excel_name(action="get", name="SalesTotal")``
    set
        Create or update a name. Requires ``name`` and ``refers_to``.
        ``refers_to`` must start with ``=``.
        Example (named range):
            ``excel_name(action="set", name="MyRange",
            refers_to="=Sheet1!$A$1:$C$10")``
        Example (LAMBDA):
            ``excel_name(action="set", name="DOUBLE",
            refers_to="=LAMBDA(x, x*2)")``
        Example (sheet-scoped):
            ``excel_name(action="set", name="Total",
            refers_to="=Sheet1!$D$10", scope="Sheet1")``
    delete
        Delete a name. Requires ``name``.
        Example: ``excel_name(action="delete", name="OldName")``
    """
    if action == "list":
        return _session.run_com(_list, workbook)
    if action == "get":
        if not name:
            raise ToolError("action='get' requires 'name'.")
        return _session.run_com(_get, workbook, name)
    if action == "set":
        if not name:
            raise ToolError("action='set' requires 'name'.")
        if not refers_to:
            raise ToolError("action='set' requires 'refers_to' (must start with '=').")
        if not refers_to.startswith("="):
            raise ToolError(
                f"'refers_to' must start with '=' (got: {refers_to!r}). "
                "Example: '=Sheet1!$A$1' or '=LAMBDA(x, x*2)'."
            )
        return _session.run_com(_set, workbook, name, refers_to, scope)
    if action == "delete":
        if not name:
            raise ToolError("action='delete' requires 'name'.")
        return _session.run_com(_delete, workbook, name)
    raise ToolError(
        f"Unknown action '{action}'. Valid: list, get, set, delete."
    )


# ── COM helpers (run inside the worker thread) ────────────────────────────────

def _is_lambda(refers_to: str) -> bool:
    """Heuristic: name is a LAMBDA formula if RefersTo starts with '=LAMBDA('."""
    return refers_to.upper().startswith("=LAMBDA(")


def _name_info(n) -> dict:
    """Compact info dict for a Name COM object."""
    try:
        ref = n.RefersTo
    except Exception:
        ref = ""
    scope_name = "workbook"
    try:
        parent = n.Parent
        # If parent is a Worksheet, it's sheet-scoped
        if hasattr(parent, "Type"):
            scope_name = parent.Name
    except Exception:
        pass
    return {
        "name": n.Name.split("!")[-1],  # strip sheet prefix from sheet-scoped names
        "refers_to": ref,
        "scope": scope_name,
        "is_lambda": _is_lambda(ref),
    }


def _list(workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    names = []
    for i in range(1, wb.Names.Count + 1):
        try:
            names.append(_name_info(wb.Names.Item(i)))
        except Exception:
            pass
    return {"names": names, "count": len(names)}


def _get(workbook: str | None, name: str) -> dict:
    wb = _session.get_workbook(workbook)
    n = _find_name(wb, name)
    return _name_info(n)


def _set(
    workbook: str | None,
    name: str,
    refers_to: str,
    scope: str | None,
) -> dict:
    wb = _session.get_workbook(workbook)
    # Determine scope object: None/workbook → workbook, else sheet
    if scope and scope.lower() != "workbook":
        ws = _session.get_sheet(scope, workbook)
        # Sheet-scoped name: use Worksheet.Names.Add
        try:
            ws.Names.Add(Name=name, RefersTo=refers_to)
            return {"set": name, "refers_to": refers_to, "scope": scope}
        except Exception as e:
            raise _session.wrap(e, f"Set sheet-scoped name '{name}' failed")
    # Workbook-scoped: wb.Names.Add
    try:
        wb.Names.Add(Name=name, RefersTo=refers_to)
        return {"set": name, "refers_to": refers_to, "scope": "workbook"}
    except Exception as e:
        raise _session.wrap(e, f"Set name '{name}' failed")


def _delete(workbook: str | None, name: str) -> dict:
    wb = _session.get_workbook(workbook)
    n = _find_name(wb, name)
    full_name = n.Name
    try:
        n.Delete()
        return {"deleted": full_name}
    except Exception as e:
        raise _session.wrap(e, f"Delete name '{name}' failed")


def _find_name(wb, name: str):
    """Return a Name COM object or raise ToolError with available names."""
    # Try exact match first (case-insensitive)
    for i in range(1, wb.Names.Count + 1):
        n = wb.Names.Item(i)
        bare = n.Name.split("!")[-1]
        if bare.lower() == name.lower() or n.Name.lower() == name.lower():
            return n
    available = []
    for i in range(1, wb.Names.Count + 1):
        available.append(wb.Names.Item(i).Name.split("!")[-1])
    raise ToolError(
        f"Name '{name}' not found. Defined names: {available or ['(none)']}"
    )
