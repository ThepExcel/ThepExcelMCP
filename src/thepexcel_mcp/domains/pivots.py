"""PivotTable operations — create, configure fields, layout, refresh, read."""

from __future__ import annotations

from fastmcp.exceptions import ToolError

from ..session import ExcelSession

_session = ExcelSession()

# ── COM constants (verified against sbroenne/mcp-server-excel PivotTableTypes.cs) ──
# PivotCache.Create SourceType
_XL_DATABASE = 1    # range or table source
_XL_EXTERNAL = 2    # external connection / data model

# PivotTable version
_XL_PIVOT_TABLE_VERSION_14 = 4  # Excel 2010+ (fully compatible with XLSX)

# PivotField.Orientation
_XL_HIDDEN = 0
_XL_ROW_FIELD = 1
_XL_COLUMN_FIELD = 2
_XL_PAGE_FIELD = 3   # "Filter" area
_XL_DATA_FIELD = 4   # "Values" area

# PivotField.Function (XlConsolidationFunction)
_AGG_FUNC = {
    "sum":      -4157,
    "count":    -4112,
    "average":  -4106,
    "avg":      -4106,
    "max":      -4136,
    "min":      -4139,
    "product":  -4149,
    "stddev":   -4155,
    "stddevp":  -4156,
    "var":      -4164,
    "varp":     -4165,
    # distinct_count is only supported in Data Model pivots via DAX measure;
    # not a VBA Function constant — handled as a special case below.
}

# PivotTable.RowAxisLayout (RowLayout)
_LAYOUT = {
    "compact": 0,
    "tabular": 1,
    "outline": 2,
}

_AREA_NAME = {
    _XL_HIDDEN: "hidden",
    _XL_ROW_FIELD: "rows",
    _XL_COLUMN_FIELD: "columns",
    _XL_PAGE_FIELD: "filters",
    _XL_DATA_FIELD: "values",
}

_AREA_STR_TO_CONST = {
    "rows": _XL_ROW_FIELD,
    "row": _XL_ROW_FIELD,
    "columns": _XL_COLUMN_FIELD,
    "column": _XL_COLUMN_FIELD,
    "col": _XL_COLUMN_FIELD,
    "filters": _XL_PAGE_FIELD,
    "filter": _XL_PAGE_FIELD,
    "values": _XL_DATA_FIELD,
    "value": _XL_DATA_FIELD,
    "data": _XL_DATA_FIELD,
}

_DEFAULT_LIMIT = 100


def pivot_action(
    action: str,
    # identity
    name: str | None = None,
    workbook: str | None = None,
    # create
    source: str | None = None,
    dest_sheet: str | None = None,
    dest_cell: str | None = None,
    # add_field / remove_field / move_field
    field: str | None = None,
    area: str | None = None,
    aggregation: str | None = None,
    number_format: str | None = None,
    # set_layout
    layout: str | None = None,
    subtotals: bool | None = None,
    grand_totals: bool | None = None,
    # read (pagination)
    offset: int = 0,
    limit: int = _DEFAULT_LIMIT,
) -> dict:
    """Dispatch an Excel PivotTable action.

    Actions
    -------
    list
        All PivotTables in the workbook: name, sheet, location, source,
        fields by area.
        Example: ``excel_pivot(action="list")``
    create
        Create a new PivotTable. Requires ``source`` and ``name``.
        ``source`` can be:
          - A range address: ``"Sheet1!A1:D100"``
          - A table name: ``"SalesTable"``
          - ``"datamodel"`` to use the workbook Data Model
        ``dest_sheet`` and ``dest_cell`` (default ``"A3"``) control placement.
        When omitted, a new sheet named ``"Pivot_<name>"`` is created.
        Example: ``excel_pivot(action="create", source="SalesTable",
        name="SalesPivot")``
    add_field
        Add a field to an area. Requires ``name``, ``field``, ``area``
        (rows/columns/filters/values). For values area, ``aggregation``
        sets the function (sum/count/average/max/min — default sum).
        ``number_format`` optionally sets a format code (e.g. ``"#,##0"``).
        Example: ``excel_pivot(action="add_field", name="SalesPivot",
        field="Region", area="rows")``
        Example: ``excel_pivot(action="add_field", name="SalesPivot",
        field="Amount", area="values", aggregation="sum",
        number_format="#,##0.00")``
    remove_field
        Remove a field from its current area. Requires ``name`` and ``field``.
        Example: ``excel_pivot(action="remove_field", name="SalesPivot",
        field="Region")``
    move_field
        Move a field to a different area. Requires ``name``, ``field``,
        and ``area``. Equivalent to remove + add.
        Example: ``excel_pivot(action="move_field", name="SalesPivot",
        field="Quarter", area="columns")``
    set_layout
        Configure layout. Requires ``name``.
        ``layout``: compact/outline/tabular.
        ``subtotals``: True/False to show/hide subtotals for all row fields.
        ``grand_totals``: True/False for both row and column grand totals.
        Example: ``excel_pivot(action="set_layout", name="SalesPivot",
        layout="tabular", grand_totals=False)``
    refresh
        Refresh the PivotTable data. Requires ``name``.
        Example: ``excel_pivot(action="refresh", name="SalesPivot")``
    delete
        Delete the PivotTable (removes the pivot object, keeps the sheet).
        Requires ``name``.
        Example: ``excel_pivot(action="delete", name="SalesPivot")``
    read
        Read the PivotTable result values paginated via TableRange2.
        Returns ``{values, total_rows, has_more, next_offset}``.
        Example: ``excel_pivot(action="read", name="SalesPivot", limit=50)``
    """
    if action == "list":
        return _list(workbook)
    if action == "create":
        _require(source, "source", action)
        _require(name, "name", action)
        return _create(name, source, dest_sheet, dest_cell, workbook)
    if action == "add_field":
        _require(name, "name", action)
        _require(field, "field", action)
        _require(area, "area", action)
        return _add_field(name, workbook, field, area, aggregation, number_format)
    if action == "remove_field":
        _require(name, "name", action)
        _require(field, "field", action)
        return _remove_field(name, workbook, field)
    if action == "move_field":
        _require(name, "name", action)
        _require(field, "field", action)
        _require(area, "area", action)
        return _move_field(name, workbook, field, area, aggregation, number_format)
    if action == "set_layout":
        _require(name, "name", action)
        return _set_layout(name, workbook, layout, subtotals, grand_totals)
    if action == "refresh":
        _require(name, "name", action)
        return _refresh(name, workbook)
    if action == "delete":
        _require(name, "name", action)
        return _delete(name, workbook)
    if action == "read":
        _require(name, "name", action)
        return _read(name, workbook, offset, limit)
    raise ToolError(
        f"Unknown action '{action}'. Valid: list, create, add_field, remove_field, "
        "move_field, set_layout, refresh, delete, read."
    )


# ── Helpers ────────────────────────────────────────────────────────────────────

def _require(value, param: str, action: str) -> None:
    if value is None:
        raise ToolError(f"action='{action}' requires '{param}'.")


def _find_pivot(wb, name: str):
    """Return a PivotTable COM object or raise ToolError with available names."""
    available = []
    for i in range(1, wb.Sheets.Count + 1):
        ws = wb.Sheets(i)
        try:
            pts = ws.PivotTables()
            for j in range(1, pts.Count + 1):
                pt = pts.Item(j)
                if pt.Name == name:
                    return pt
                available.append(pt.Name)
        except Exception:
            pass
    raise ToolError(
        f"PivotTable '{name}' not found. Available: {available or ['(none)']}"
    )


def _find_table_range(wb, table_name: str) -> str:
    """Return 'SheetName!RangeAddress' for a named ListObject."""
    for i in range(1, wb.Sheets.Count + 1):
        ws = wb.Sheets(i)
        for j in range(1, ws.ListObjects.Count + 1):
            lo = ws.ListObjects(j)
            if lo.Name == table_name:
                # Quote sheet name in case it contains spaces
                sheet_quoted = f"'{ws.Name}'"
                return f"{sheet_quoted}!{lo.Range.Address}"
    raise ToolError(
        f"Table '{table_name}' not found. Use a range address like "
        f"'Sheet1!A1:D100' or a valid table name."
    )


def _pivot_fields_by_area(pt) -> dict:
    """Return fields grouped by area as a compact dict."""
    areas: dict[str, list] = {
        "rows": [], "columns": [], "filters": [], "values": [], "hidden": []
    }
    try:
        pfs = pt.PivotFields()
        for i in range(1, pfs.Count + 1):
            try:
                pf = pfs.Item(i)
                orientation = int(pf.Orientation)
                area_key = _AREA_NAME.get(orientation, "hidden")
                entry: dict = {"name": pf.SourceName or pf.Name}
                if orientation == _XL_DATA_FIELD:
                    try:
                        entry["function"] = _func_name(int(pf.Function))
                    except Exception:
                        pass
                areas[area_key].append(entry)
            except Exception:
                pass
    except Exception:
        pass
    return {k: v for k, v in areas.items() if v}  # omit empty areas


def _func_name(com_func: int) -> str:
    for k, v in _AGG_FUNC.items():
        if v == com_func and k not in ("avg",):
            return k
    return str(com_func)


def _pivot_info(pt, ws_name: str) -> dict:
    info = {
        "name": pt.Name,
        "sheet": ws_name,
        "location": pt.TableRange1.Address,
        "fields": _pivot_fields_by_area(pt),
    }
    try:
        info["source"] = str(pt.SourceData)
    except Exception:
        info["source"] = "unknown"
    return info


# ── Action implementations ─────────────────────────────────────────────────────

def _list(workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    pivots = []
    for i in range(1, wb.Sheets.Count + 1):
        ws = wb.Sheets(i)
        try:
            pts = ws.PivotTables()
            for j in range(1, pts.Count + 1):
                try:
                    pivots.append(_pivot_info(pts.Item(j), ws.Name))
                except Exception:
                    pass
        except Exception:
            pass
    return {"pivot_tables": pivots, "count": len(pivots)}


def _create(
    name: str,
    source: str,
    dest_sheet: str | None,
    dest_cell: str | None,
    workbook: str | None,
) -> dict:
    wb = _session.get_workbook(workbook)
    app = wb.Application

    # Resolve source → source_data string and source_type constant
    is_datamodel = source.strip().lower() == "datamodel"

    if is_datamodel:
        source_type = _XL_EXTERNAL
    else:
        # If source looks like a table name (no ! and no space or colon), look it up
        if "!" not in source and ":" not in source:
            source_data = _find_table_range(wb, source)
        else:
            source_data = source
        source_type = _XL_DATABASE

    # Determine destination sheet and cell
    if dest_sheet is None:
        # Create a new sheet named "Pivot_<name>" (max 31 chars)
        new_sheet_name = f"Pivot_{name}"[:31]
        # Remove existing sheet with same name if any
        app.DisplayAlerts = False
        try:
            wb.Sheets(new_sheet_name).Delete()
        except Exception:
            pass
        app.DisplayAlerts = True
        dest_ws = wb.Sheets.Add()
        dest_ws.Name = new_sheet_name
    else:
        dest_ws = _session.get_sheet(dest_sheet, workbook)

    cell_addr = dest_cell or "A3"
    dest_range = dest_ws.Range(cell_addr)

    try:
        if is_datamodel:
            # Data Model pivot: use workbook connection "ThisWorkbookDataModel"
            # PivotCaches.Create(xlExternal, connection_obj)
            try:
                conn = wb.Connections("ThisWorkbookDataModel")
            except Exception:
                raise ToolError(
                    "Data Model connection not found. "
                    "Load at least one table into the Data Model first via Power Query "
                    "or by enabling 'Add this data to the Data Model' on import."
                )
            pc = wb.PivotCaches().Create(
                SourceType=_XL_EXTERNAL,
                SourceData=conn,
                Version=_XL_PIVOT_TABLE_VERSION_14,
            )
        else:
            # Range or table source
            pc = wb.PivotCaches().Create(
                SourceType=_XL_DATABASE,
                SourceData=source_data,
                Version=_XL_PIVOT_TABLE_VERSION_14,
            )

        pt = pc.CreatePivotTable(
            TableDestination=dest_range,
            TableName=name,
        )
        pt.RefreshTable()

        return {
            "created": name,
            "sheet": dest_ws.Name,
            "location": pt.TableRange1.Address,
            "source": source,
            "available_fields": _get_available_fields(pt),
        }
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"Create PivotTable '{name}' failed")


def _get_available_fields(pt) -> list[str]:
    """Return list of field names available in the pivot (hidden + placed)."""
    fields = []
    try:
        pfs = pt.PivotFields()
        for i in range(1, pfs.Count + 1):
            try:
                pf = pfs.Item(i)
                fields.append(pf.SourceName or pf.Name)
            except Exception:
                pass
    except Exception:
        pass
    return fields


def _resolve_area(area: str) -> int:
    key = area.lower().strip()
    if key not in _AREA_STR_TO_CONST:
        raise ToolError(
            f"area='{area}' invalid. Valid: rows, columns, filters, values."
        )
    return _AREA_STR_TO_CONST[key]


def _is_datamodel_pivot(pt) -> bool:
    """True if this pivot was created from the workbook Data Model."""
    try:
        # Data model pivots have SourceData == None / not applicable
        # and their PivotCaches.SourceType == xlExternal (2)
        return int(pt.PivotCache().SourceType) == _XL_EXTERNAL
    except Exception:
        return False


def _add_field(
    name: str,
    workbook: str | None,
    field_name: str,
    area: str,
    aggregation: str | None,
    number_format: str | None,
) -> dict:
    wb = _session.get_workbook(workbook)
    pt = _find_pivot(wb, name)
    orientation = _resolve_area(area)

    # Data Model pivots expose measures as CubeFields with captions like
    # "[Measures].[MeasureName]" or "[TableName].[ColumnName]".
    # We try PivotFields first (works for regular pivots); on failure
    # try a CubeField match by caption or by "[Measures].[field_name]" pattern.
    pf = None
    try:
        pf = pt.PivotFields(field_name)
    except Exception:
        pass

    if pf is None and _is_datamodel_pivot(pt):
        # Try CubeFields — enumerate and match by Name or caption
        cube_candidates = [
            f"[Measures].[{field_name}]",
            field_name,
        ]
        try:
            cfs = pt.CubeFields()
            for ci in range(1, cfs.Count + 1):
                cf = cfs.Item(ci)
                try:
                    cf_name = cf.Name  # e.g. "[Measures].[Total Sales]"
                    if cf_name in cube_candidates or cf_name == field_name:
                        cf.Orientation = orientation
                        if number_format:
                            try:
                                cf.NumberFormat = number_format
                            except Exception:
                                pass
                        return {
                            "pivot": name,
                            "field": cf_name,
                            "area": _AREA_NAME.get(orientation, area),
                            "note": "data_model_cubefield",
                        }
                except Exception:
                    continue
        except Exception:
            pass

    if pf is None:
        available = _get_available_fields(pt)
        raise ToolError(
            f"Field '{field_name}' not found in PivotTable '{name}'. "
            f"Available fields: {available}"
        )

    try:
        pf.Orientation = orientation

        if orientation == _XL_DATA_FIELD and aggregation:
            agg_key = aggregation.lower()
            if agg_key not in _AGG_FUNC:
                raise ToolError(
                    f"aggregation='{aggregation}' invalid. "
                    f"Valid: {[k for k in _AGG_FUNC if k != 'avg']}"
                )
            pf.Function = _AGG_FUNC[agg_key]

        if orientation == _XL_DATA_FIELD and number_format:
            pf.NumberFormat = number_format

        result = {
            "pivot": name,
            "field": field_name,
            "area": _AREA_NAME.get(orientation, area),
        }
        if orientation == _XL_DATA_FIELD:
            try:
                result["aggregation"] = _func_name(int(pf.Function))
            except Exception:
                pass
        return result
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"Add field '{field_name}' to '{name}' failed")


def _remove_field(name: str, workbook: str | None, field_name: str) -> dict:
    wb = _session.get_workbook(workbook)
    pt = _find_pivot(wb, name)
    try:
        pf = pt.PivotFields(field_name)
    except Exception:
        raise ToolError(f"Field '{field_name}' not found in PivotTable '{name}'.")
    try:
        old_orientation = int(pf.Orientation)
        if old_orientation == _XL_HIDDEN:
            raise ToolError(
                f"Field '{field_name}' is not currently placed in any area."
            )
        pf.Orientation = _XL_HIDDEN
        return {
            "pivot": name,
            "field": field_name,
            "removed_from": _AREA_NAME.get(old_orientation, "unknown"),
        }
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"Remove field '{field_name}' from '{name}' failed")


def _move_field(
    name: str,
    workbook: str | None,
    field_name: str,
    area: str,
    aggregation: str | None,
    number_format: str | None,
) -> dict:
    # Move = remove from current + add to target area
    wb = _session.get_workbook(workbook)
    pt = _find_pivot(wb, name)
    orientation = _resolve_area(area)
    try:
        pf = pt.PivotFields(field_name)
    except Exception:
        available = _get_available_fields(pt)
        raise ToolError(
            f"Field '{field_name}' not found in '{name}'. Available: {available}"
        )
    try:
        old_orientation = int(pf.Orientation)
        pf.Orientation = orientation
        if orientation == _XL_DATA_FIELD and aggregation:
            agg_key = aggregation.lower()
            if agg_key not in _AGG_FUNC:
                raise ToolError(
                    f"aggregation='{aggregation}' invalid. "
                    f"Valid: {[k for k in _AGG_FUNC if k != 'avg']}"
                )
            pf.Function = _AGG_FUNC[agg_key]
        if orientation == _XL_DATA_FIELD and number_format:
            pf.NumberFormat = number_format
        return {
            "pivot": name,
            "field": field_name,
            "moved_from": _AREA_NAME.get(old_orientation, "hidden"),
            "moved_to": _AREA_NAME.get(orientation, area),
        }
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"Move field '{field_name}' in '{name}' failed")


def _set_layout(
    name: str,
    workbook: str | None,
    layout: str | None,
    subtotals: bool | None,
    grand_totals: bool | None,
) -> dict:
    wb = _session.get_workbook(workbook)
    pt = _find_pivot(wb, name)
    result: dict = {"pivot": name}
    try:
        if layout is not None:
            key = layout.lower()
            if key not in _LAYOUT:
                raise ToolError(
                    f"layout='{layout}' invalid. Valid: {list(_LAYOUT.keys())}"
                )
            # RowAxisLayout sets form for all row fields
            pt.RowAxisLayout(_LAYOUT[key])
            result["layout"] = layout

        if subtotals is not None:
            # Apply to all row fields: Subtotals[1] = automatic
            try:
                pfs = pt.PivotFields()
                for i in range(1, pfs.Count + 1):
                    try:
                        pf = pfs.Item(i)
                        if int(pf.Orientation) == _XL_ROW_FIELD:
                            pf.Subtotals[1] = subtotals
                    except Exception:
                        pass
            except Exception:
                pass
            result["subtotals"] = subtotals

        if grand_totals is not None:
            pt.RowGrand = grand_totals
            pt.ColumnGrand = grand_totals
            result["grand_totals"] = grand_totals

        return result
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"Set layout on '{name}' failed")


def _refresh(name: str, workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    pt = _find_pivot(wb, name)
    try:
        pt.RefreshTable()
        return {"refreshed": name}
    except Exception as e:
        raise _session.wrap(e, f"Refresh PivotTable '{name}' failed")


def _delete(name: str, workbook: str | None) -> dict:
    wb = _session.get_workbook(workbook)
    pt = _find_pivot(wb, name)
    sheet_name = pt.Parent.Name
    try:
        pt.TableRange2.ClearContents()
        pt.TableRange2.Clear()
        # TableRange2 includes filter fields; this removes the pivot structure
        # For a full delete, use PivotTableWizard.Delete() or clear the range
        # The standard VBA approach is to select the range and delete
        pt.TableRange2.Delete()
        return {"deleted": name, "sheet": sheet_name}
    except Exception as e:
        raise _session.wrap(e, f"Delete PivotTable '{name}' failed")


def _read(name: str, workbook: str | None, offset: int, limit: int) -> dict:
    wb = _session.get_workbook(workbook)
    pt = _find_pivot(wb, name)
    try:
        # TableRange2 includes page/filter fields at top; TableRange1 is just the body
        rng = pt.TableRange1
        raw = rng.Value
        if raw is None:
            return {"values": [], "total_rows": 0, "has_more": False, "next_offset": None}
        if not isinstance(raw, tuple):
            raw = ((raw,),)
        elif raw and not isinstance(raw[0], tuple):
            raw = (raw,)
        total_rows = len(raw)
        page = raw[offset: offset + limit]
        rows = [list(r) for r in page]
        has_more = (offset + limit) < total_rows
        return {
            "values": rows,
            "total_rows": total_rows,
            "has_more": has_more,
            "next_offset": offset + limit if has_more else None,
        }
    except Exception as e:
        raise _session.wrap(e, f"Read PivotTable '{name}' failed")
