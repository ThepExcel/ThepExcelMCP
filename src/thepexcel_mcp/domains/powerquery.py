"""Power Query operations ported from ThepExcel-PQ-MCP PoC.

All 10 actions from the PoC plus analyze/analyze_raw wiring.
The tricky Connections.Add2 + ListObjects.Add load-to-table pattern is
preserved verbatim — it took real debugging to get right (see PoC excel_com.py).
"""

from __future__ import annotations

from fastmcp.exceptions import ToolError

from ..analysis.pq_analyzer import analyze_mcode
from ..session import ExcelSession

_session = ExcelSession()


def powerquery_action(
    action: str,
    # identify
    name: str | None = None,
    workbook: str | None = None,
    # create / update
    formula: str | None = None,
    description: str | None = None,
    new_name: str | None = None,
    # load_to_table
    sheet_name: str | None = None,
    # analyze_raw
    raw_formula: str | None = None,
    # load_to_datamodel (Phase 2)
    load_to_datamodel: bool = False,
) -> dict:
    """Dispatch a Power Query action.

    Actions
    -------
    list
        List all queries in the workbook (name, description, line count,
        and ``model_connection`` flag when detectable).
    get
        Return M code + metadata for a single query. Requires ``name``.
    create
        Create a new query. Requires ``name`` and ``formula`` (let…in M code).
    update
        Update formula, description, or rename. Requires ``name``.
        Pass only the fields you want to change; others are preserved.
    delete
        Delete a query by name. Requires ``name``.
    refresh
        Refresh a single query via its "Query - <name>" connection. Requires ``name``.
    refresh_all
        Refresh all queries and data connections in the workbook.
    load_to_table
        Load a connection-only query to a new worksheet Table. Requires ``name``.
        Creates a sheet named ``sheet_name`` (defaults to query name, max 31 chars).
        Uses the proven Connections.Add2 + Mashup-OLEDB pattern.
    load_to_datamodel
        Load a query directly to the Data Model (no worksheet table). Requires ``name``.
        Uses Connections.Add2 with CreateModelConnection=True.
        After loading, use excel_datamodel(action="list_tables") to confirm.
    analyze
        Analyze M code from an existing query for anti-patterns. Requires ``name``.
    analyze_raw
        Analyze M code provided directly (no Excel needed). Requires ``raw_formula``.
    """
    # Validate args (pure Python) before entering the COM worker
    # analyze_raw is pure Python (no COM) — skip run_com
    if action == "analyze_raw":
        _require(raw_formula, "raw_formula", action)
        return _analyze_raw(raw_formula, name or "unnamed")
    if action == "list":
        return _session.run_com(_list, workbook)
    if action == "get":
        _require(name, "name", action)
        return _session.run_com(_get, name, workbook)
    if action == "create":
        _require(name, "name", action)
        _require(formula, "formula", action)
        return _session.run_com(_create, name, formula, description or "", workbook)
    if action == "update":
        _require(name, "name", action)
        return _session.run_com(_update, name, formula, description, new_name, workbook)
    if action == "delete":
        _require(name, "name", action)
        return _session.run_com(_delete, name, workbook)
    if action == "refresh":
        _require(name, "name", action)
        return _session.run_com(_refresh, name, workbook)
    if action == "refresh_all":
        return _session.run_com(_refresh_all, workbook)
    if action == "load_to_table":
        _require(name, "name", action)
        return _session.run_com(_load_to_table, name, sheet_name, workbook)
    if action == "load_to_datamodel":
        _require(name, "name", action)
        return _session.run_com(_load_to_datamodel, name, workbook)
    if action == "analyze":
        _require(name, "name", action)
        return _session.run_com(_analyze, name, workbook)
    raise ToolError(
        f"Unknown action '{action}'. Valid: list, get, create, update, delete, "
        "refresh, refresh_all, load_to_table, load_to_datamodel, analyze, analyze_raw."
    )


# ── Helpers ────────────────────────────────────────────────────────────────────

def _require(value, param: str, action: str) -> None:
    if not value:
        raise ToolError(f"action='{action}' requires '{param}'.")


def _query_obj(wb, name: str):
    """Return COM query object or raise ToolError with available names."""
    try:
        return wb.Queries(name)
    except Exception:
        available = [wb.Queries.Item(i + 1).Name for i in range(wb.Queries.Count)]
        raise ToolError(
            f"Query '{name}' not found. Available: {available or ['(none)']}"
        )


def _format_result(q) -> dict:
    return {
        "name": q.Name,
        "formula": q.Formula,
        "description": q.Description or "",
    }


# ── Action implementations ─────────────────────────────────────────────────────

def _list(workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    # Build set of query names that have a model connection
    model_conn_names: set[str] = set()
    try:
        for i in range(1, wb.Connections.Count + 1):
            c = wb.Connections.Item(i)
            try:
                # Type 7 = xlConnectionTypeOLEDB with CreateModelConnection
                # Check for ModelConnection type (6) or OLEDB with model flag
                if c.Type == 6:  # xlConnectionTypeModel
                    cname = c.Name  # e.g. "Query - SalesData" or bare query name
                    if cname.startswith("Query - "):
                        model_conn_names.add(cname[len("Query - "):])
            except Exception:
                pass
    except Exception:
        pass
    queries = []
    for i in range(1, wb.Queries.Count + 1):
        q = wb.Queries.Item(i)
        entry = {
            "name": q.Name,
            "description": q.Description or "",
            "line_count": q.Formula.count("\n") + 1,
        }
        if q.Name in model_conn_names:
            entry["model_connection"] = True
        queries.append(entry)
    return {"queries": queries, "count": len(queries)}


def _get(name: str, workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    q = _query_obj(wb, name)
    return _format_result(q)


def _create(name: str, formula: str, description: str, workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    # Guard duplicate
    for i in range(1, wb.Queries.Count + 1):
        if wb.Queries.Item(i).Name == name:
            raise ToolError(f"Query '{name}' already exists. Use action='update' instead.")
    try:
        q = wb.Queries.Add(name, formula)
        if description:
            q.Description = description
        result = _format_result(q)
        result["warnings"] = _get_warnings(q)
        return result
    except Exception as e:
        raise _session.wrap(e, f"Create query '{name}' failed")


def _update(
    name: str,
    formula: str | None,
    description: str | None,
    new_name: str | None,
    workbook: str | None,
) -> dict:
    wb = _session.get_workbook(workbook)
    q = _query_obj(wb, name)
    try:
        if formula is not None:
            q.Formula = formula
        if description is not None:
            q.Description = description
        if new_name is not None:
            q.Name = new_name
        result = _format_result(q)
        # Include M code warnings only when formula was changed
        if formula is not None:
            result["warnings"] = _get_warnings(q)
        return result
    except Exception as e:
        raise _session.wrap(e, f"Update query '{name}' failed")


def _delete(name: str, workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    _query_obj(wb, name).Delete()
    return {"deleted": name}


def _refresh(name: str, workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    conn_name = f"Query - {name}"
    try:
        wb.Connections(conn_name).Refresh()
        return {"refreshed": name}
    except Exception:
        # Fallback: RefreshAll (covers connection-only queries with no load target)
        try:
            _query_obj(wb, name)  # verify query exists first
            wb.RefreshAll()
            return {"refreshed": name, "method": "refresh_all_fallback"}
        except ToolError:
            raise
        except Exception as e:
            raise _session.wrap(e, f"Refresh '{name}' failed")


def _refresh_all(workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    wb.RefreshAll()
    return {"refreshed_all": True, "workbook": wb.Name}


def _load_to_table(
    name: str, sheet_name: str | None, workbook: str | None
) -> dict:
    """Load a query result to a worksheet Table.

    Proven Connections.Add2 + Mashup-OLEDB pattern from PoC excel_com.py.
    This exact call sequence is required — deviating breaks the load target.
    """
    wb = _session.get_workbook(workbook)
    app = wb.Application

    # Verify query exists
    _query_obj(wb, name)

    target_name = (sheet_name or name)[:31]

    # Delete existing sheet with same name (clean slate)
    app.DisplayAlerts = False
    try:
        wb.Sheets(target_name).Delete()
    except Exception:
        pass
    app.DisplayAlerts = True

    ws = wb.Sheets.Add()
    ws.Name = target_name

    # Remove stale connections for this query
    to_delete = []
    for i in range(1, wb.Connections.Count + 1):
        c = wb.Connections.Item(i)
        try:
            if name in c.OLEDBConnection.Connection:
                to_delete.append(c.Name)
        except Exception:
            continue
    for cname in to_delete:
        try:
            wb.Connections(cname).Delete()
        except Exception:
            pass

    # Step 1: Create OLEDB Mashup connection via Connections.Add2
    # lCmdtype=2 (xlCmdSql), Location WITHOUT quotes, CommandText as SQL SELECT.
    # Exact form required — see PoC debugging notes.
    conn_name = f"Query - {name}"
    conn_str = (
        "OLEDB;"
        "Provider=Microsoft.Mashup.OleDb.1;"
        "Data Source=$Workbook$;"
        f"Location={name}"
    )
    command_text = f"SELECT * FROM [{name}]"

    try:
        conn = wb.Connections.Add2(
            conn_name,       # Name
            "",              # Description
            conn_str,        # ConnectionString
            command_text,    # CommandText
            2,               # lCmdtype = xlCmdSql
            False,           # CreateModelConnection
            False,           # ImportRelationships
        )
    except Exception as e:
        raise _session.wrap(e, "Connections.Add2 failed")

    # Step 2: Add ListObject using the connection object
    # xlSrcExternal=0, xlYes=1
    dest = ws.Range("A1")
    try:
        lo = ws.ListObjects.Add(0, conn, True, 1, dest)
    except Exception as e:
        raise _session.wrap(e, "ListObjects.Add failed")

    # Step 3: Configure QueryTable and refresh synchronously
    qt = lo.QueryTable
    qt.CommandType = 2   # xlCmdSql
    qt.CommandText = command_text
    qt.AdjustColumnWidth = True
    qt.BackgroundQuery = False
    try:
        qt.Refresh(False)  # synchronous
    except Exception as e:
        raise _session.wrap(e, "QueryTable.Refresh failed")

    # Rename the ListObject to the query name for easy reference
    try:
        lo.Name = name
    except Exception:
        pass  # not fatal — name collision on existing table is harmless

    rows = lo.Range.Rows.Count - 1  # subtract header row
    return {
        "loaded": name,
        "sheet": ws.Name,
        "table": lo.Name,
        "rows": rows,
    }


def _load_to_datamodel(name: str, workbook: str | None) -> dict:
    """Load a query result directly to the Data Model (no worksheet table).

    Uses Connections.Add2 with CreateModelConnection=True — same Mashup-OLEDB
    pattern as _load_to_table but skips the ListObjects.Add step.
    After Add2 + Refresh, wb.Model.Initialize() flushes metadata.
    """
    wb = _session.get_workbook(workbook)

    # Verify query exists
    _query_obj(wb, name)

    conn_name = f"Query - {name}"
    conn_str = (
        "OLEDB;"
        "Provider=Microsoft.Mashup.OleDb.1;"
        "Data Source=$Workbook$;"
        f"Location={name}"
    )
    command_text = f"SELECT * FROM [{name}]"

    # Remove stale model connection for this query
    to_delete = []
    for i in range(1, wb.Connections.Count + 1):
        c = wb.Connections.Item(i)
        try:
            if name in (c.OLEDBConnection.Connection or ""):
                to_delete.append(c.Name)
        except Exception:
            pass
    for cname in to_delete:
        try:
            wb.Connections(cname).Delete()
        except Exception:
            pass

    try:
        conn = wb.Connections.Add2(
            conn_name,
            "",
            conn_str,
            command_text,
            2,      # lCmdtype = xlCmdSql
            True,   # CreateModelConnection = True
            False,  # ImportRelationships
        )
        conn.Refresh()
    except Exception as e:
        raise _session.wrap(e, f"load_to_datamodel: Connections.Add2 failed for '{name}'")

    # Flush model metadata
    model_table_count = None
    try:
        wb.Model.Initialize()
        model_table_count = wb.Model.ModelTables.Count
    except Exception:
        pass

    return {
        "loaded_to_datamodel": name,
        "connection": conn_name,
        "model_table_count": model_table_count,
    }


def _get_warnings(q) -> list[str]:
    """Run M code analyzer and return warning/error messages (non-blocking)."""
    try:
        result = analyze_mcode(q.Name, q.Formula)
        return [
            f"[{i.severity}] {i.rule}: {i.message}"
            for i in result.issues
            if i.severity in ("warning", "error")
        ]
    except Exception:
        return []


def _analyze(name: str, workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    q = _query_obj(wb, name)
    result = analyze_mcode(q.Name, q.Formula)
    return _format_analysis(result)


def _analyze_raw(formula: str, query_name: str) -> dict:
    result = analyze_mcode(query_name, formula)
    return _format_analysis(result)


def _format_analysis(result) -> dict:
    return {
        "query_name": result.query_name,
        "step_count": result.step_count,
        "estimated_complexity": result.estimated_complexity,
        "summary": result.summary,
        "issues": [
            {
                "severity": issue.severity,
                "rule": issue.rule,
                "message": issue.message,
                "line": issue.line,
                "suggestion": issue.suggestion,
            }
            for issue in result.issues
        ],
    }
