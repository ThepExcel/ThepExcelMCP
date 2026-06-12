"""Excel Table (ListObject) operations — create, read, modify, filter, sort."""

from __future__ import annotations

from fastmcp.exceptions import ToolError

from ..session import ExcelSession

_session = ExcelSession()

# xlTotalsCalculation constants
_TOTALS_FUNC = {
    "none": 0,
    "sum": 1,
    "average": 2,
    "avg": 2,
    "count": 3,
    "countnums": 4,
    "max": 5,
    "min": 6,
    "stddev": 7,
    "var": 9,
}

# AutoFilter operator constants
_XL_AND = 1
_XL_OR = 2
_XL_FILTER_VALUES = 7

# Range.Sort constants
_XL_ASCENDING = 1
_XL_DESCENDING = 2
_XL_YES = 1  # Header row present

_DEFAULT_STYLE = "TableStyleMedium2"
_DEFAULT_LIMIT = 100


def table_action(
    action: str,
    # identity
    name: str | None = None,
    workbook: str | None = None,
    sheet: str | None = None,
    # create
    range: str | None = None,
    style: str | None = None,
    has_headers: bool = True,
    # read
    columns: list[str] | None = None,
    offset: int = 0,
    limit: int = _DEFAULT_LIMIT,
    # append_rows
    values: list | None = None,
    # add_column
    column_name: str | None = None,
    formula: str | None = None,
    # sort
    sort_column: str | None = None,
    ascending: bool = True,
    # filter
    filter_column: str | None = None,
    filter_op: str | None = None,
    filter_value: str | None = None,
    # toggle_totals
    show_totals: bool | None = None,
    total_func: str | None = None,
    # rename
    new_name: str | None = None,
    # delete
    keep_data: bool = True,
) -> dict:
    """Dispatch an Excel Table action.

    Actions
    -------
    list
        All ListObjects across all sheets: name, sheet, range, columns, row count,
        has_totals, style.
        Example: ``excel_table(action="list")``
    create
        Create a table from a range. Requires ``range`` and ``name``.
        ``sheet`` defaults to active sheet. ``style`` defaults to
        ``"TableStyleMedium2"``. ``has_headers`` defaults to True.
        Example: ``excel_table(action="create", range="A1:D10", name="SalesTable")``
    read
        Read table data paginated. Requires ``name``.
        ``columns`` filters to specific column names. Pagination via
        ``offset`` and ``limit`` (default 100). Returns ``{values, columns,
        total_rows, has_more, next_offset}``.
        Note: structured references like ``"SalesTable[Amount]"`` also work
        directly via ``excel_range(action="read", range="SalesTable[Amount]")``.
        Example: ``excel_table(action="read", name="SalesTable", limit=50)``
    append_rows
        Append rows to a table. Requires ``name`` and ``values`` (2-D list).
        Example: ``excel_table(action="append_rows", name="SalesTable",
        values=[[101, "Jan", 5000], [102, "Feb", 6000]])``
    add_column
        Add a column to the table. Requires ``name`` and ``column_name``.
        ``formula`` adds a calculated column (structured reference, e.g.
        ``"=[@Amount]*0.1"``). Without formula, column is blank.
        Example: ``excel_table(action="add_column", name="Sales",
        column_name="Tax", formula="=[@Amount]*0.1")``
    sort
        Sort table by a column. Requires ``name`` and ``sort_column``.
        ``ascending`` defaults to True.
        Example: ``excel_table(action="sort", name="Sales",
        sort_column="Amount", ascending=False)``
    filter
        Set autofilter on a column. Requires ``name``, ``filter_column``,
        and one of:
          - ``filter_op="equals"`` + ``filter_value``
          - ``filter_op="contains"`` + ``filter_value`` (uses wildcard ``*value*``)
          - ``filter_op="greater"`` / ``"less"`` + ``filter_value``
          - ``filter_op="clear_filters"`` (``filter_column`` still required)
        Example: ``excel_table(action="filter", name="Sales",
        filter_column="Region", filter_op="equals", filter_value="North")``
    set_style
        Change table style. Requires ``name`` and ``style``.
        Example: ``excel_table(action="set_style", name="Sales",
        style="TableStyleLight1")``
    toggle_totals
        Show/hide the totals row. Requires ``name`` and ``show_totals`` (bool).
        Optionally set per-column aggregation: ``total_func`` (sum/average/
        count/max/min/stddev/var/none) with ``column_name``.
        Example: ``excel_table(action="toggle_totals", name="Sales",
        show_totals=True, column_name="Amount", total_func="sum")``
    rename
        Rename the table. Requires ``name`` and ``new_name``.
        Example: ``excel_table(action="rename", name="Sales", new_name="SalesData")``
    delete
        Delete the table. Requires ``name``.
        ``keep_data=True`` (default) unlinks (Unlist) while keeping cell data.
        ``keep_data=False`` deletes the entire data range.
        Example: ``excel_table(action="delete", name="OldTable")``
    """
    if action == "list":
        return _list(workbook)
    if action == "create":
        _require(name, "name", action)
        _require(range, "range", action)
        return _create(name, range, sheet, workbook, style or _DEFAULT_STYLE, has_headers)
    if action == "read":
        _require(name, "name", action)
        return _read(name, workbook, columns, offset, limit)
    if action == "append_rows":
        _require(name, "name", action)
        if not values:
            raise ToolError("action='append_rows' requires 'values' (2-D list).")
        return _append_rows(name, workbook, values)
    if action == "add_column":
        _require(name, "name", action)
        _require(column_name, "column_name", action)
        return _add_column(name, workbook, column_name, formula)
    if action == "sort":
        _require(name, "name", action)
        _require(sort_column, "sort_column", action)
        return _sort(name, workbook, sort_column, ascending)
    if action == "filter":
        _require(name, "name", action)
        _require(filter_column, "filter_column", action)
        if filter_op == "clear_filters":
            return _clear_filters(name, workbook)
        _require(filter_op, "filter_op", action)
        return _filter(name, workbook, filter_column, filter_op, filter_value)
    if action == "set_style":
        _require(name, "name", action)
        _require(style, "style", action)
        return _set_style(name, workbook, style)
    if action == "toggle_totals":
        _require(name, "name", action)
        if show_totals is None:
            raise ToolError("action='toggle_totals' requires 'show_totals' (bool).")
        return _toggle_totals(name, workbook, show_totals, column_name, total_func)
    if action == "rename":
        _require(name, "name", action)
        _require(new_name, "new_name", action)
        return _rename(name, workbook, new_name)
    if action == "delete":
        _require(name, "name", action)
        return _delete(name, workbook, keep_data)
    raise ToolError(
        f"Unknown action '{action}'. Valid: list, create, read, append_rows, "
        "add_column, sort, filter, set_style, toggle_totals, rename, delete."
    )


# ── Helpers ────────────────────────────────────────────────────────────────────

def _require(value, param: str, action: str) -> None:
    if value is None:
        raise ToolError(f"action='{action}' requires '{param}'.")


def _find_table(wb, name: str):
    """Return a ListObject COM object or raise ToolError with available names."""
    available = []
    for i in range(1, wb.Sheets.Count + 1):
        ws = wb.Sheets(i)
        for j in range(1, ws.ListObjects.Count + 1):
            lo = ws.ListObjects(j)
            if lo.Name == name:
                return lo
            available.append(lo.Name)
    raise ToolError(
        f"Table '{name}' not found. Available tables: {available or ['(none)']}"
    )


def _col_index(lo, col_name: str) -> int:
    """Return 1-based column index by name, or raise ToolError."""
    for i in range(1, lo.ListColumns.Count + 1):
        if lo.ListColumns(i).Name == col_name:
            return i
    cols = [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)]
    raise ToolError(f"Column '{col_name}' not found. Columns: {cols}")


def _table_info(lo) -> dict:
    """Compact info dict for a ListObject."""
    ws_name = lo.Parent.Name
    cols = [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)]
    # DataBodyRange is None when table has headers only (no data rows)
    dbr = lo.DataBodyRange
    row_count = dbr.Rows.Count if dbr is not None else 0
    style_obj = lo.TableStyle
    style_name = style_obj.Name if style_obj is not None else ""
    return {
        "name": lo.Name,
        "sheet": ws_name,
        "range": lo.Range.Address,
        "columns": cols,
        "row_count": row_count,
        "has_totals": bool(lo.ShowTotals),
        "style": style_name,
    }


# ── Action implementations ─────────────────────────────────────────────────────

def _list(workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    tables = []
    for i in range(1, wb.Sheets.Count + 1):
        ws = wb.Sheets(i)
        for j in range(1, ws.ListObjects.Count + 1):
            tables.append(_table_info(ws.ListObjects(j)))
    return {"tables": tables, "count": len(tables)}


def _create(
    name: str,
    range_str: str,
    sheet: str | None,
    workbook: str | None,
    style: str,
    has_headers: bool,
) -> dict:
    wb = _session.get_workbook(workbook)
    ws = _session.get_sheet(sheet, workbook)
    # Guard duplicate name across workbook
    for i in range(1, wb.Sheets.Count + 1):
        s = wb.Sheets(i)
        for j in range(1, s.ListObjects.Count + 1):
            if s.ListObjects(j).Name == name:
                raise ToolError(
                    f"Table '{name}' already exists on sheet '{s.Name}'. "
                    "Use a different name or delete the existing table first."
                )
    try:
        rng = ws.Range(range_str)
    except Exception as e:
        raise _session.wrap(e, f"Invalid range '{range_str}'")
    try:
        # xlSrcRange = 1, xlYes = 1 (has headers), xlNo = 2
        header_const = 1 if has_headers else 2
        lo = ws.ListObjects.Add(1, rng, None, header_const)
        lo.Name = name
        lo.TableStyle = style
        return _table_info(lo)
    except Exception as e:
        raise _session.wrap(e, f"Create table '{name}' failed")


def _read(
    name: str,
    workbook: str | None,
    columns: list[str] | None,
    offset: int,
    limit: int,
) -> dict:
    wb = _session.get_workbook(workbook)
    lo = _find_table(wb, name)
    all_cols = [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)]

    # Determine which column indices to include
    if columns:
        invalid = [c for c in columns if c not in all_cols]
        if invalid:
            raise ToolError(f"Columns not found in table: {invalid}. Available: {all_cols}")
        col_indices = [all_cols.index(c) for c in columns]
        out_cols = columns
    else:
        col_indices = list(range(len(all_cols)))
        out_cols = all_cols

    dbr = lo.DataBodyRange
    if dbr is None:
        return {
            "columns": out_cols,
            "values": [],
            "total_rows": 0,
            "has_more": False,
            "next_offset": None,
        }

    raw = dbr.Value
    if raw is None:
        rows_data = []
    elif not isinstance(raw, tuple):
        rows_data = [[raw]]
    elif raw and not isinstance(raw[0], tuple):
        rows_data = [list(raw)]
    else:
        rows_data = [list(r) for r in raw]

    total_rows = len(rows_data)
    page = rows_data[offset: offset + limit]
    filtered = [[row[i] for i in col_indices] for row in page]
    has_more = (offset + limit) < total_rows
    return {
        "columns": out_cols,
        "values": filtered,
        "total_rows": total_rows,
        "has_more": has_more,
        "next_offset": offset + limit if has_more else None,
    }


def _append_rows(name: str, workbook: str | None, values: list) -> dict:
    wb = _session.get_workbook(workbook)
    lo = _find_table(wb, name)
    try:
        for row_values in values:
            new_row = lo.ListRows.Add()
            for col_idx, val in enumerate(row_values, start=1):
                if col_idx <= lo.ListColumns.Count:
                    new_row.Range.Cells(1, col_idx).Value = val
        dbr = lo.DataBodyRange
        new_row_count = dbr.Rows.Count if dbr is not None else 0
        return {
            "table": name,
            "appended_rows": len(values),
            "total_rows": new_row_count,
        }
    except Exception as e:
        raise _session.wrap(e, f"Append rows to '{name}' failed")


def _add_column(
    name: str, workbook: str | None, col_name: str, formula: str | None
) -> dict:
    wb = _session.get_workbook(workbook)
    lo = _find_table(wb, name)
    # Guard duplicate column name
    for i in range(1, lo.ListColumns.Count + 1):
        if lo.ListColumns(i).Name == col_name:
            raise ToolError(f"Column '{col_name}' already exists in table '{name}'.")
    try:
        new_col = lo.ListColumns.Add()
        new_col.Name = col_name
        if formula:
            # Write formula to first data cell — Excel auto-fills structured ref
            dbr = lo.DataBodyRange
            if dbr is not None:
                col_idx = lo.ListColumns(col_name).Index
                dbr.Columns(col_idx).Cells(1, 1).Formula2 = formula
        return {
            "table": name,
            "column_added": col_name,
            "column_count": lo.ListColumns.Count,
        }
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"Add column '{col_name}' to '{name}' failed")


def _sort(name: str, workbook: str | None, col_name: str, ascending: bool) -> dict:
    wb = _session.get_workbook(workbook)
    lo = _find_table(wb, name)
    col_idx = _col_index(lo, col_name)
    try:
        sort_range = lo.Range
        col_range = lo.ListColumns(col_idx).Range
        sort_range.Sort(
            Key1=col_range,
            Order1=_XL_ASCENDING if ascending else _XL_DESCENDING,
            Header=_XL_YES,
        )
        return {"table": name, "sorted_by": col_name, "ascending": ascending}
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"Sort table '{name}' failed")


def _filter(
    name: str,
    workbook: str | None,
    col_name: str,
    op: str,
    value: str | None,
) -> dict:
    wb = _session.get_workbook(workbook)
    lo = _find_table(wb, name)
    col_idx = _col_index(lo, col_name)
    valid_ops = ("equals", "contains", "greater", "less", "clear_filters")
    if op not in valid_ops:
        raise ToolError(f"filter_op='{op}' invalid. Valid: {valid_ops}")
    if op != "clear_filters" and value is None:
        raise ToolError(f"filter_op='{op}' requires 'filter_value'.")
    try:
        rng = lo.Range
        if op == "equals":
            rng.AutoFilter(Field=col_idx, Criteria1=value)
        elif op == "contains":
            rng.AutoFilter(Field=col_idx, Criteria1=f"*{value}*")
        elif op == "greater":
            rng.AutoFilter(Field=col_idx, Criteria1=f">{value}")
        elif op == "less":
            rng.AutoFilter(Field=col_idx, Criteria1=f"<{value}")
        return {
            "table": name,
            "filtered_column": col_name,
            "op": op,
            "value": value,
        }
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"Filter table '{name}' failed")


def _clear_filters(name: str, workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    lo = _find_table(wb, name)
    try:
        # ShowAllData clears all active autofilters while keeping the dropdown arrows
        ws = lo.Parent
        if ws.AutoFilterMode:
            ws.AutoFilter.ShowAllData()
        return {"table": name, "filters_cleared": True}
    except Exception as e:
        raise _session.wrap(e, f"Clear filters on '{name}' failed")


def _set_style(name: str, workbook: str | None, style: str) -> dict:
    wb = _session.get_workbook(workbook)
    lo = _find_table(wb, name)
    try:
        lo.TableStyle = style
        return {"table": name, "style": style}
    except Exception as e:
        raise _session.wrap(e, f"Set style on '{name}' failed — style name may not exist")


def _toggle_totals(
    name: str,
    workbook: str | None,
    show_totals: bool,
    col_name: str | None,
    total_func: str | None,
) -> dict:
    wb = _session.get_workbook(workbook)
    lo = _find_table(wb, name)
    try:
        lo.ShowTotals = show_totals
        if show_totals and col_name and total_func:
            func_key = total_func.lower()
            if func_key not in _TOTALS_FUNC:
                raise ToolError(
                    f"total_func='{total_func}' invalid. "
                    f"Valid: {list(_TOTALS_FUNC.keys())}"
                )
            col_idx = _col_index(lo, col_name)
            lo.ListColumns(col_idx).TotalsCalculation = _TOTALS_FUNC[func_key]
        return {
            "table": name,
            "show_totals": show_totals,
            "column": col_name,
            "total_func": total_func,
        }
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"Toggle totals on '{name}' failed")


def _rename(name: str, workbook: str | None, new_name: str) -> dict:
    wb = _session.get_workbook(workbook)
    lo = _find_table(wb, name)
    # Guard duplicate
    for i in range(1, wb.Sheets.Count + 1):
        ws = wb.Sheets(i)
        for j in range(1, ws.ListObjects.Count + 1):
            if ws.ListObjects(j).Name == new_name:
                raise ToolError(
                    f"Table name '{new_name}' already exists. Choose a different name."
                )
    try:
        lo.Name = new_name
        return {"renamed": {"from": name, "to": new_name}}
    except Exception as e:
        raise _session.wrap(e, f"Rename table '{name}' failed")


def _delete(name: str, workbook: str | None, keep_data: bool) -> dict:
    wb = _session.get_workbook(workbook)
    lo = _find_table(wb, name)
    addr = lo.Range.Address
    ws_name = lo.Parent.Name
    try:
        if keep_data:
            # Unlist: converts table to normal range, preserving data and formatting
            lo.Unlist()
        else:
            lo.Delete()
        return {
            "deleted": name,
            "sheet": ws_name,
            "range": addr,
            "data_kept": keep_data,
        }
    except Exception as e:
        raise _session.wrap(e, f"Delete table '{name}' failed")
