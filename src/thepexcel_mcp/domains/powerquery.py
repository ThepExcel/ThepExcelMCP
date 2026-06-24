"""Power Query operations ported from ThepExcel-PQ-MCP PoC.

All 10 actions from the PoC plus analyze/analyze_raw wiring.
The tricky Connections.Add2 + ListObjects.Add load-to-table pattern is
preserved verbatim — it took real debugging to get right (see PoC excel_com.py).
"""

from __future__ import annotations

import re

from fastmcp.exceptions import ToolError

from ..analysis.pq_analyzer import analyze_mcode
from ..session import ExcelSession

_session = ExcelSession()


# ── Parameter query pure-Python helpers ───────────────────────────────────────
# Logic verified in scratch/tier3/exp_pq_params.py (9/9 pure-Python + live COM).
# Copied VERBATIM from that file; only renamed to module-private _names.

def _m_escape_text(value: str) -> str:
    """Escape a Python string into an M text literal body (without outer quotes).

    M text literals escape a double-quote by doubling it: "" inside the string.
    """
    return value.replace('"', '""')


def _split_literal_and_meta(formula: str) -> tuple[str, str]:
    """Split a parameter formula into (leading_literal, meta_part).

    Splits by structure, NOT by searching for ' meta ' as a substring, so text
    values that happen to contain the word ' meta ' are handled correctly.

    - Quoted literal: scan forward past the closing quote (doubled "" = escaped
      quote inside the string; keep scanning). The literal is everything up to
      and including the closing quote.
    - Bare number literal: token up to the first whitespace.

    Returns (literal_token, remainder_including_leading_space_and_meta_keyword).
    Raises ValueError if no meta record follows (caller treats as non-parameter).
    """
    f = formula.lstrip()
    if f.startswith('"'):
        # Walk the quoted text literal character by character
        i = 1  # start after opening quote
        while i < len(f):
            if f[i] == '"':
                # Peek ahead: doubled quote = escaped quote inside string
                if i + 1 < len(f) and f[i + 1] == '"':
                    i += 2  # consume both and continue inside string
                    continue
                else:
                    # Closing quote found
                    literal = f[:i + 1]
                    meta_part = f[i + 1:]  # everything after closing quote
                    if not re.search(r"\bmeta\b", meta_part):
                        raise ValueError("not a parameter formula (no meta record after quoted literal)")
                    return literal, meta_part
            i += 1
        raise ValueError("not a parameter formula (unterminated quoted literal)")
    else:
        # Bare number literal — token up to first whitespace
        m = re.search(r"\s", f)
        if not m:
            raise ValueError("not a parameter formula (no whitespace after literal)")
        literal = f[:m.start()]
        meta_part = f[m.start():]
        if not re.search(r"\bmeta\b", meta_part):
            raise ValueError("not a parameter formula (no meta record after number literal)")
        return literal, meta_part


def _build_parameter_m(value, ptype: str | None = None, required: bool = True) -> str:
    """Assemble the M for a parameter query.

    ptype: "Text" | "Number" | None (None → inferred from Python type).
    Text values are emitted quoted+escaped; Number values bare.
    """
    if ptype is None:
        ptype = "Number" if isinstance(value, (int, float)) and not isinstance(value, bool) else "Text"
    if ptype == "Text":
        literal = '"' + _m_escape_text(str(value)) + '"'
    else:  # Number (or any non-text scalar) → bare literal
        literal = repr(value) if isinstance(value, float) else str(value)
    return (
        f"{literal} meta "
        f"[IsParameterQuery=true, Type=\"{ptype}\", IsParameterQueryRequired={str(required).lower()}]"
    )


# Detection: a query is a parameter if its formula contains IsParameterQuery=true
_RE_IS_PARAM = re.compile(r"IsParameterQuery\s*=\s*true", re.IGNORECASE)
# A query is a function if its body starts with a parameter list "(...) =>"
_RE_FUNC_SIG = re.compile(r"^\s*\([^)]*\)\s*(as\s+[^=]+?)?=>", re.DOTALL)
_RE_TYPE_FIELD = re.compile(r"Type\s*=\s*\"([^\"]*)\"")


def _is_parameter(formula: str) -> bool:
    return bool(_RE_IS_PARAM.search(formula or ""))


def _is_function(formula: str) -> bool:
    f = (formula or "").lstrip()
    return bool(_RE_FUNC_SIG.match(f))


def _parse_parameter(formula: str) -> dict:
    """Return {value, type, raw_literal} parsed from a parameter formula."""
    try:
        raw, _meta_part = _split_literal_and_meta(formula.lstrip())
    except ValueError:
        # Fallback: treat entire formula (stripped) as the raw literal
        raw = formula.strip()
    tmatch = _RE_TYPE_FIELD.search(formula)
    ptype = tmatch.group(1) if tmatch else None
    # decode the literal back to a Python value (best-effort, for display)
    if raw.startswith('"') and raw.endswith('"'):
        value = raw[1:-1].replace('""', '"')
    else:
        try:
            value = float(raw) if ("." in raw or "e" in raw.lower()) else int(raw)
        except ValueError:
            value = raw
    return {"value": value, "type": ptype, "raw_literal": raw}


def _set_parameter_formula(old_formula: str, new_value, ptype: str | None = None) -> str:
    """Produce a new formula with the literal replaced but meta record preserved."""
    try:
        _old_literal, meta_part = _split_literal_and_meta(old_formula.lstrip())
    except ValueError as exc:
        raise ValueError(f"not a parameter formula (no meta record): {exc}") from exc

    if ptype is None:
        tmatch = _RE_TYPE_FIELD.search(old_formula)
        ptype = tmatch.group(1) if tmatch else None
    if ptype == "Text":
        literal = '"' + _m_escape_text(str(new_value)) + '"'
    else:
        literal = repr(new_value) if isinstance(new_value, float) else str(new_value)
    # Normalize to exactly one leading space before "meta"
    meta_part = " " + meta_part.lstrip()
    return f"{literal}{meta_part}"


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
    # parameter management (Tier-3)
    value=None,
    param_type: str | None = None,
    required: bool = True,
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
    create_parameter
        Create a new Power Query parameter (a query whose formula is a scalar meta
        record). Requires ``name`` and ``value``. Optional ``param_type`` ("Text" |
        "Number"; inferred from Python type when omitted). ``required`` defaults True.
    get_parameter
        Return parsed {value, type} for a parameter query. Requires ``name``.
        Raises ToolError if the query exists but is not a parameter.
    set_parameter
        Update the scalar value of an existing parameter query. Requires ``name``
        and ``value``. Optionally pass ``param_type`` to coerce the type; otherwise
        the existing type in the meta record is preserved.
    list_parameters
        List all parameter queries in the workbook (those with IsParameterQuery=true
        in their formula). Returns [{name, value, type}].
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
    if action == "create_parameter":
        _require(name, "name", action)
        if value is None:
            raise ToolError(f"action='create_parameter' requires 'value'.")
        return _session.run_com(_create_parameter, name, value, param_type, required, workbook)
    if action == "get_parameter":
        _require(name, "name", action)
        return _session.run_com(_get_parameter, name, workbook)
    if action == "set_parameter":
        _require(name, "name", action)
        if value is None:
            raise ToolError(f"action='set_parameter' requires 'value'.")
        return _session.run_com(_set_parameter, name, value, param_type, workbook)
    if action == "list_parameters":
        return _session.run_com(_list_parameters, workbook)
    raise ToolError(
        f"Unknown action '{action}'. Valid: list, get, create, update, delete, "
        "refresh, refresh_all, load_to_table, load_to_datamodel, analyze, analyze_raw, "
        "create_parameter, get_parameter, set_parameter, list_parameters."
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
        formula = q.Formula
        entry = {
            "name": q.Name,
            "description": q.Description or "",
            "line_count": formula.count("\n") + 1,
            "is_parameter": _is_parameter(formula),
            "is_function": _is_function(formula),
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
        # Set BackgroundQuery=False to request synchronous refresh.
        # Note: the Mashup/PQ engine requires Excel's UI message pump to
        # complete model population. This works when Excel is visible (real
        # Claude usage) but may deadlock in headless COM automation tests.
        try:
            conn.OLEDBConnection.BackgroundQuery = False
        except Exception:
            pass
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


def _create_parameter(
    name: str,
    value,
    param_type: str | None,
    required: bool,
    workbook: str | None,
) -> dict:
    wb = _session.get_workbook(workbook)
    # Guard duplicate (mirror _create)
    for i in range(1, wb.Queries.Count + 1):
        if wb.Queries.Item(i).Name == name:
            raise ToolError(
                f"Query '{name}' already exists. Use action='set_parameter' to update its value."
            )
    formula = _build_parameter_m(value, param_type, required)
    try:
        q = wb.Queries.Add(name, formula)
    except Exception as e:
        raise _session.wrap(e, f"Create parameter '{name}' failed")
    # Verify-EFFECT: read back and confirm the meta record survived the round-trip
    actual = q.Formula
    if not _is_parameter(actual):
        raise ToolError(
            f"create_parameter verify-effect failed: formula written but IsParameterQuery "
            f"not detected in read-back. Got: {actual!r}"
        )
    parsed = _parse_parameter(actual)
    return {
        "created_parameter": name,
        "formula": actual,
        "value": parsed["value"],
        "type": parsed["type"],
    }


def _get_parameter(name: str, workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    q = _query_obj(wb, name)
    if not _is_parameter(q.Formula):
        raise ToolError(
            f"'{name}' is not a parameter query (IsParameterQuery=true not found in formula). "
            f"Use action='get' to retrieve raw M code."
        )
    parsed = _parse_parameter(q.Formula)
    return {
        "name": name,
        "value": parsed["value"],
        "type": parsed["type"],
        "raw_literal": parsed["raw_literal"],
        "formula": q.Formula,
    }


def _set_parameter(
    name: str,
    value,
    param_type: str | None,
    workbook: str | None,
) -> dict:
    wb = _session.get_workbook(workbook)
    q = _query_obj(wb, name)
    if not _is_parameter(q.Formula):
        raise ToolError(
            f"'{name}' is not a parameter query. Use action='update' to rewrite arbitrary M code."
        )
    new_formula = _set_parameter_formula(q.Formula, value, param_type)
    try:
        q.Formula = new_formula
    except Exception as e:
        raise _session.wrap(e, f"set_parameter '{name}': Formula write failed")
    # Verify-EFFECT: read back and compare against the USER-REQUESTED value,
    # not a re-parse of new_formula (a consistent round-trip bug would pass that check).
    actual = q.Formula
    parsed = _parse_parameter(actual)
    readback_val = parsed["value"]
    # Numeric coercion: compare as numbers to avoid float-repr mismatches
    try:
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            if abs(float(readback_val) - float(value)) > 1e-12:
                raise ToolError(
                    f"set_parameter verify-effect mismatch for '{name}': "
                    f"requested {value!r}, read back {readback_val!r} (formula: {actual!r})"
                )
        else:
            if str(readback_val) != str(value):
                raise ToolError(
                    f"set_parameter verify-effect mismatch for '{name}': "
                    f"requested {value!r}, read back {readback_val!r} (formula: {actual!r})"
                )
    except ToolError:
        raise
    except Exception:
        # Fallback: string comparison (e.g. type coercion edge case)
        if str(readback_val) != str(value):
            raise ToolError(
                f"set_parameter verify-effect mismatch for '{name}': "
                f"requested {value!r}, read back {readback_val!r} (formula: {actual!r})"
            )
    return {
        "set_parameter": name,
        "formula": actual,
        "value": parsed["value"],
    }


def _list_parameters(workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    params = []
    for i in range(1, wb.Queries.Count + 1):
        q = wb.Queries.Item(i)
        if _is_parameter(q.Formula):
            parsed = _parse_parameter(q.Formula)
            params.append({
                "name": q.Name,
                "value": parsed["value"],
                "type": parsed["type"],
            })
    return {"parameters": params, "count": len(params)}


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
