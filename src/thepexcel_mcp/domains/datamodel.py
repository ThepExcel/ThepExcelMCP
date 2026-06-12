"""Data Model (Power Pivot) + DAX measure operations.

COM object chain:
    wb.Model                           → Model object
    wb.Model.ModelTables               → ModelTables collection
    wb.Model.ModelRelationships        → ModelRelationships collection
    wb.Model.ModelMeasures             → ModelMeasures collection
    wb.Model.ModelFormatGeneral        → format factory (property, not ctor)
    wb.Model.ModelFormatDecimalNumber  → same pattern
    wb.Model.ModelFormatCurrency       → same pattern
    wb.Model.ModelFormatPercentageNumber → same pattern

Format objects are properties on the Model object — no separate constructor.
(COM pattern confirmed from sbroenne/mcp-server-excel DataModelCommands.Helpers.cs)

add_table technique: Connections.Add2(…, CreateModelConnection=True)
followed by wb.Model.Initialize() to flush the model metadata.
This mirrors the load_to_table Mashup-OLEDB pattern but sets
CreateModelConnection=True so Excel adds the query to the model
without creating a worksheet output table.
"""

from __future__ import annotations

import pythoncom
from fastmcp.exceptions import ToolError

from ..session import ExcelSession

_session = ExcelSession()


# ── Format type → Model property name ─────────────────────────────────────────

_FORMAT_MAP = {
    "general":    "ModelFormatGeneral",
    "decimal":    "ModelFormatDecimalNumber",
    "number":     "ModelFormatDecimalNumber",
    "currency":   "ModelFormatCurrency",
    "percent":    "ModelFormatPercentageNumber",
    "percentage": "ModelFormatPercentageNumber",
    "whole":      "ModelFormatWholeNumber",
    "integer":    "ModelFormatWholeNumber",
    "boolean":    "ModelFormatBoolean",
    "date":       "ModelFormatDate",
    "scientific": "ModelFormatScientificNumber",
}

VALID_FORMATS = sorted(set(_FORMAT_MAP.keys()))


def _get_format_obj(model, format_type: str | None):
    """Return a ModelFormatXxx COM object from wb.Model.<PropertyName>.

    Falls back to ModelFormatGeneral for unknown/None format_type.
    Pattern: properties on Model, not constructors — per sbroenne Helpers.cs.
    """
    key = (format_type or "general").lower().strip()
    prop = _FORMAT_MAP.get(key, "ModelFormatGeneral")
    try:
        return getattr(model, prop)
    except Exception:
        # Fallback: some Excel versions need initialization before format props work
        try:
            model.Initialize()
            return getattr(model, prop)
        except Exception:
            return model.ModelFormatGeneral


def datamodel_action(
    action: str,
    workbook: str | None = None,
    # add_table
    source_type: str | None = None,
    source_name: str | None = None,
    # add_relationship / delete_relationship
    from_table: str | None = None,
    from_column: str | None = None,
    to_table: str | None = None,
    to_column: str | None = None,
    relationship_index: int | None = None,
    # add_measure / update_measure / delete_measure
    measure_name: str | None = None,
    table: str | None = None,
    formula: str | None = None,
    format_type: str | None = None,
    decimal_places: int | None = None,
    use_thousand_sep: bool | None = None,
    description: str | None = None,
    new_formula: str | None = None,
    new_format_type: str | None = None,
    new_description: str | None = None,
) -> dict:
    """Dispatch a Data Model / DAX action.

    Actions
    -------
    info
        Model summary: table count, relationship count, measure count.
    list_tables
        All ModelTables: name, source_name, record_count, columns.
    add_table
        Add a workbook Table or Power Query to the Data Model.
        Requires ``source_type`` (table|query) and ``source_name``.
    list_relationships
        All ModelRelationships: FK table.column → PK table.column, active flag.
    add_relationship
        Create a relationship. Requires ``from_table``, ``from_column``,
        ``to_table``, ``to_column``.
        from = FK side (many), to = PK side (one).
    delete_relationship
        Delete by index (1-based). Requires ``relationship_index``.
    list_measures
        All ModelMeasures: name, table, formula, format, description.
    add_measure
        Add a DAX measure. Requires ``measure_name``, ``table``, ``formula``.
        ``format_type``: general|decimal|number|currency|percent|whole|integer|boolean|date|scientific.
        ``decimal_places``, ``use_thousand_sep`` apply to decimal/whole formats.
    update_measure
        Update formula, format, or description. Requires ``measure_name``.
        Pass only fields to change (new_formula, new_format_type, new_description).
    delete_measure
        Delete a measure by name. Requires ``measure_name``.
    refresh
        Model.Refresh() — refreshes all data sources and reprocesses.
    """
    if action == "info":
        return _info(workbook)
    if action == "list_tables":
        return _list_tables(workbook)
    if action == "add_table":
        _require(source_type, "source_type", action)
        _require(source_name, "source_name", action)
        return _add_table(source_type, source_name, workbook)
    if action == "list_relationships":
        return _list_relationships(workbook)
    if action == "add_relationship":
        _require(from_table, "from_table", action)
        _require(from_column, "from_column", action)
        _require(to_table, "to_table", action)
        _require(to_column, "to_column", action)
        return _add_relationship(from_table, from_column, to_table, to_column, workbook)
    if action == "delete_relationship":
        _require(relationship_index, "relationship_index", action)
        return _delete_relationship(relationship_index, workbook)
    if action == "list_measures":
        return _list_measures(workbook)
    if action == "add_measure":
        _require(measure_name, "measure_name", action)
        _require(table, "table", action)
        _require(formula, "formula", action)
        return _add_measure(
            measure_name, table, formula, format_type, decimal_places,
            use_thousand_sep, description, workbook
        )
    if action == "update_measure":
        _require(measure_name, "measure_name", action)
        return _update_measure(
            measure_name, new_formula, new_format_type, new_description,
            decimal_places, use_thousand_sep, workbook
        )
    if action == "delete_measure":
        _require(measure_name, "measure_name", action)
        return _delete_measure(measure_name, workbook)
    if action == "refresh":
        return _refresh(workbook)
    raise ToolError(
        f"Unknown action '{action}'. Valid: info, list_tables, add_table, "
        "list_relationships, add_relationship, delete_relationship, "
        "list_measures, add_measure, update_measure, delete_measure, refresh."
    )


# ── Helpers ────────────────────────────────────────────────────────────────────

def _require(value, param: str, action: str) -> None:
    if not value and value != 0:
        raise ToolError(f"action='{action}' requires '{param}'.")


def _get_model(workbook: str | None):
    """Return (wb, model) or raise ToolError if no model exists."""
    wb = _session.get_workbook(workbook)
    try:
        model = wb.Model
        if model is None:
            raise ToolError(
                "This workbook has no Data Model. "
                "Load at least one table to the Data Model first: "
                "use excel_datamodel(action='add_table', source_type='table', source_name='MyTable') "
                "or import data via Power Query with 'Only Create Connection' + 'Add this data to the Data Model'."
            )
        return wb, model
    except ToolError:
        raise
    except Exception:
        raise ToolError(
            "Data Model not available. Requires Excel with PowerPivot add-in enabled "
            "(included in Microsoft 365 / Office 2019+)."
        )


def _find_model_table(model, table_name: str):
    """Return ModelTable COM object by name or raise ToolError."""
    try:
        mt = model.ModelTables(table_name)
        return mt
    except Exception:
        available = []
        try:
            for i in range(1, model.ModelTables.Count + 1):
                available.append(model.ModelTables.Item(i).Name)
        except Exception:
            pass
        raise ToolError(
            f"Model table '{table_name}' not found. "
            f"Available model tables: {available or ['(none)']}"
        )


def _find_model_column(model, table_name: str, col_name: str):
    """Return ModelTableColumn COM object."""
    mt = _find_model_table(model, table_name)
    try:
        return mt.ModelTableColumns(col_name)
    except Exception:
        available = []
        try:
            for i in range(1, mt.ModelTableColumns.Count + 1):
                available.append(mt.ModelTableColumns.Item(i).Name)
        except Exception:
            pass
        raise ToolError(
            f"Column '{col_name}' not found in model table '{table_name}'. "
            f"Available columns: {available or ['(none)']}"
        )


def _find_measure(model, measure_name: str):
    """Return ModelMeasure COM object by name or raise ToolError."""
    try:
        for i in range(1, model.ModelMeasures.Count + 1):
            m = model.ModelMeasures.Item(i)
            if m.Name == measure_name:
                return m
    except Exception:
        pass
    available = _list_measure_names(model)
    raise ToolError(
        f"Measure '{measure_name}' not found. "
        f"Available measures: {available or ['(none)']}"
    )


def _list_measure_names(model) -> list[str]:
    names = []
    try:
        for i in range(1, model.ModelMeasures.Count + 1):
            names.append(model.ModelMeasures.Item(i).Name)
    except Exception:
        pass
    return names


def _format_info(m) -> str:
    """Return a short format string for a measure's FormatInformation."""
    try:
        fmt = m.FormatInformation
        cls = type(fmt).__name__.lower()
        if "general" in cls:
            return "general"
        if "decimal" in cls:
            return "decimal"
        if "currency" in cls:
            return "currency"
        if "percent" in cls:
            return "percent"
        if "whole" in cls or "wholenumber" in cls:
            return "whole"
        if "boolean" in cls:
            return "boolean"
        if "date" in cls:
            return "date"
        if "scientific" in cls:
            return "scientific"
        # pywin32 late-binding: use ProgID or COM type name
        return str(cls)
    except Exception:
        return "general"


def _apply_format_options(fmt_obj, decimal_places: int | None, use_thousand_sep: bool | None):
    """Apply DecimalPlaces / UseThousandSeparator when the format object supports them."""
    if decimal_places is not None:
        try:
            fmt_obj.DecimalPlaces = decimal_places
        except Exception:
            pass
    if use_thousand_sep is not None:
        try:
            fmt_obj.UseThousandSeparator = use_thousand_sep
        except Exception:
            pass


# ── Action implementations ─────────────────────────────────────────────────────

def _info(workbook: str | None) -> dict:
    wb, model = _get_model(workbook)
    result: dict = {"workbook": wb.Name}
    try:
        result["table_count"] = model.ModelTables.Count
    except Exception:
        result["table_count"] = 0
    try:
        result["relationship_count"] = model.ModelRelationships.Count
    except Exception:
        result["relationship_count"] = 0
    try:
        result["measure_count"] = model.ModelMeasures.Count
    except Exception:
        result["measure_count"] = 0
    return result


def _list_tables(workbook: str | None) -> dict:
    wb, model = _get_model(workbook)
    tables = []
    count = 0
    try:
        count = model.ModelTables.Count
        for i in range(1, count + 1):
            mt = model.ModelTables.Item(i)
            entry: dict = {"name": mt.Name}
            try:
                entry["source_name"] = mt.SourceName
            except Exception:
                entry["source_name"] = None
            try:
                entry["record_count"] = mt.RecordCount
            except Exception:
                entry["record_count"] = None
            cols = []
            try:
                for j in range(1, mt.ModelTableColumns.Count + 1):
                    c = mt.ModelTableColumns.Item(j)
                    cols.append(c.Name)
            except Exception:
                pass
            entry["columns"] = cols
            tables.append(entry)
    except Exception:
        pass
    return {"tables": tables, "count": count}


def _add_table(
    source_type: str,
    source_name: str,
    workbook: str | None,
) -> dict:
    """Add a workbook Table (ListObject) or PQ query to the Data Model.

    Technique: Connections.Add2(…, CreateModelConnection=True).
    After Add2, wb.Model.Initialize() flushes metadata so the table
    appears in ModelTables immediately.

    For source_type='table': the ListObject must already exist.
    For source_type='query': the Power Query must already exist as a connection.
    """
    wb = _session.get_workbook(workbook)
    stype = source_type.lower().strip()

    if stype not in ("table", "query"):
        raise ToolError(
            f"source_type='{source_type}' invalid. Valid: table, query."
        )

    # Verify source exists
    if stype == "table":
        found = False
        for i in range(1, wb.Sheets.Count + 1):
            ws = wb.Sheets(i)
            try:
                for j in range(1, ws.ListObjects.Count + 1):
                    if ws.ListObjects(j).Name == source_name:
                        found = True
                        break
            except Exception:
                pass
            if found:
                break
        if not found:
            raise ToolError(
                f"Table '{source_name}' not found. "
                "Create it first with excel_table(action='create', …)."
            )
    else:
        # query: verify it exists in wb.Queries
        found = False
        try:
            for i in range(1, wb.Queries.Count + 1):
                if wb.Queries.Item(i).Name == source_name:
                    found = True
                    break
        except Exception:
            pass
        if not found:
            raise ToolError(
                f"Query '{source_name}' not found. "
                "Create it first with excel_powerquery(action='create', …)."
            )

    # Build Mashup-OLEDB connection string (same as load_to_table but
    # CreateModelConnection=True, no worksheet target)
    conn_name = f"Query - {source_name}"
    conn_str = (
        "OLEDB;"
        "Provider=Microsoft.Mashup.OleDb.1;"
        "Data Source=$Workbook$;"
        f"Location={source_name}"
    )
    command_text = f"SELECT * FROM [{source_name}]"

    try:
        # Remove stale model connection for the same source if present
        to_delete = []
        for i in range(1, wb.Connections.Count + 1):
            c = wb.Connections.Item(i)
            try:
                if source_name in (c.OLEDBConnection.Connection or ""):
                    to_delete.append(c.Name)
            except Exception:
                pass
        for cname in to_delete:
            try:
                wb.Connections(cname).Delete()
            except Exception:
                pass

        # CreateModelConnection=True — adds to Data Model, no worksheet output
        conn = wb.Connections.Add2(
            conn_name,        # Name
            "",               # Description
            conn_str,         # ConnectionString
            command_text,     # CommandText
            2,                # lCmdtype = xlCmdSql
            True,             # CreateModelConnection ← key difference
            False,            # ImportRelationships
        )
        conn.Refresh()
    except Exception as e:
        raise _session.wrap(e, f"Add '{source_name}' to Data Model failed")

    # Flush model metadata
    try:
        wb.Model.Initialize()
    except Exception:
        pass

    # Confirm table now appears in model
    model_table_count = 0
    try:
        model_table_count = wb.Model.ModelTables.Count
    except Exception:
        pass

    return {
        "added_to_model": source_name,
        "source_type": stype,
        "model_table_count": model_table_count,
    }


def _list_relationships(workbook: str | None) -> dict:
    wb, model = _get_model(workbook)
    rels = []
    count = 0
    try:
        count = model.ModelRelationships.Count
        for i in range(1, count + 1):
            r = model.ModelRelationships.Item(i)
            entry: dict = {"index": i}
            try:
                entry["from_table"] = r.ForeignKeyTable.Name
                entry["from_column"] = r.ForeignKeyColumn.Name
                entry["to_table"] = r.PrimaryKeyTable.Name
                entry["to_column"] = r.PrimaryKeyColumn.Name
                entry["active"] = bool(r.Active)
            except Exception as inner:
                entry["error"] = str(inner)
            rels.append(entry)
    except Exception:
        pass
    return {"relationships": rels, "count": count}


def _add_relationship(
    from_table: str,
    from_column: str,
    to_table: str,
    to_column: str,
    workbook: str | None,
) -> dict:
    """Create a one-to-many relationship: from (FK/many) → to (PK/one)."""
    wb, model = _get_model(workbook)
    fk_col = _find_model_column(model, from_table, from_column)
    pk_col = _find_model_column(model, to_table, to_column)
    try:
        rel = model.ModelRelationships.Add(
            ForeignKeyColumn=fk_col,
            PrimaryKeyColumn=pk_col,
        )
        return {
            "created_relationship": {
                "from_table": from_table,
                "from_column": from_column,
                "to_table": to_table,
                "to_column": to_column,
                "active": bool(rel.Active),
            }
        }
    except Exception as e:
        raise _session.wrap(e, f"Add relationship {from_table}.{from_column} → {to_table}.{to_column} failed")


def _delete_relationship(index: int, workbook: str | None) -> dict:
    wb, model = _get_model(workbook)
    count = 0
    try:
        count = model.ModelRelationships.Count
    except Exception:
        pass
    if index < 1 or index > count:
        raise ToolError(
            f"relationship_index={index} out of range. "
            f"Valid: 1–{count}. Use list_relationships to see current indexes."
        )
    try:
        rel = model.ModelRelationships.Item(index)
        info = {
            "deleted_index": index,
            "from_table": rel.ForeignKeyTable.Name,
            "from_column": rel.ForeignKeyColumn.Name,
            "to_table": rel.PrimaryKeyTable.Name,
            "to_column": rel.PrimaryKeyColumn.Name,
        }
        rel.Delete()
        return info
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"Delete relationship[{index}] failed")


def _list_measures(workbook: str | None) -> dict:
    wb, model = _get_model(workbook)
    measures = []
    count = 0
    try:
        count = model.ModelMeasures.Count
        for i in range(1, count + 1):
            m = model.ModelMeasures.Item(i)
            entry: dict = {
                "name": m.Name,
                "formula": m.Formula,
                "format": _format_info(m),
            }
            try:
                entry["table"] = m.AssociatedTable.Name
            except Exception:
                entry["table"] = None
            try:
                entry["description"] = m.Description or ""
            except Exception:
                entry["description"] = ""
            measures.append(entry)
    except Exception:
        pass
    return {"measures": measures, "count": count}


def _add_measure(
    measure_name: str,
    table_name: str,
    formula: str,
    format_type: str | None,
    decimal_places: int | None,
    use_thousand_sep: bool | None,
    description: str | None,
    workbook: str | None,
) -> dict:
    wb, model = _get_model(workbook)
    mt = _find_model_table(model, table_name)

    # Guard duplicate
    existing = _list_measure_names(model)
    if measure_name in existing:
        raise ToolError(
            f"Measure '{measure_name}' already exists. Use action='update_measure' instead."
        )

    fmt_obj = _get_format_obj(model, format_type)
    _apply_format_options(fmt_obj, decimal_places, use_thousand_sep)

    # FormatInformation is required (never None/Type.Missing) per sbroenne note
    import pythoncom as _pc  # noqa — already imported at top but being explicit
    desc_arg = description if description else pythoncom.Empty

    try:
        m = model.ModelMeasures.Add(
            MeasureName=measure_name,
            AssociatedTable=mt,
            Formula=formula,
            FormatInformation=fmt_obj,
            Description=desc_arg,
        )
        return {
            "created_measure": {
                "name": m.Name,
                "table": mt.Name,
                "formula": m.Formula,
                "format": _format_info(m),
                "description": description or "",
            }
        }
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"Add measure '{measure_name}' failed")


def _update_measure(
    measure_name: str,
    new_formula: str | None,
    new_format_type: str | None,
    new_description: str | None,
    decimal_places: int | None,
    use_thousand_sep: bool | None,
    workbook: str | None,
) -> dict:
    """Update a measure's formula, format, or description in place.

    ModelMeasure has no direct .Formula setter — the only way to update is
    delete + re-add preserving the rest. We capture existing values first.
    """
    wb, model = _get_model(workbook)
    m = _find_measure(model, measure_name)

    # Capture existing values before delete
    try:
        old_formula = m.Formula
    except Exception:
        old_formula = new_formula or ""
    try:
        old_table = m.AssociatedTable
        old_table_name = old_table.Name
    except Exception:
        old_table = None
        old_table_name = None
    try:
        old_desc = m.Description or ""
    except Exception:
        old_desc = ""
    old_format = _format_info(m)

    formula_to_use = new_formula if new_formula is not None else old_formula
    format_to_use = new_format_type if new_format_type is not None else old_format
    desc_to_use = new_description if new_description is not None else old_desc

    if old_table is None:
        raise ToolError(
            f"Cannot determine AssociatedTable for measure '{measure_name}'. "
            "Delete and re-add the measure manually."
        )

    try:
        m.Delete()
    except Exception as e:
        raise _session.wrap(e, f"Delete (for update) of measure '{measure_name}' failed")

    # Re-add with new values
    fmt_obj = _get_format_obj(model, format_to_use)
    _apply_format_options(fmt_obj, decimal_places, use_thousand_sep)
    desc_arg = desc_to_use if desc_to_use else pythoncom.Empty

    try:
        mt = _find_model_table(model, old_table_name)
        m_new = model.ModelMeasures.Add(
            MeasureName=measure_name,
            AssociatedTable=mt,
            Formula=formula_to_use,
            FormatInformation=fmt_obj,
            Description=desc_arg,
        )
        return {
            "updated_measure": {
                "name": m_new.Name,
                "table": old_table_name,
                "formula": m_new.Formula,
                "format": _format_info(m_new),
                "description": desc_to_use,
            }
        }
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"Re-add (update) measure '{measure_name}' failed")


def _delete_measure(measure_name: str, workbook: str | None) -> dict:
    wb, model = _get_model(workbook)
    m = _find_measure(model, measure_name)
    try:
        m.Delete()
        return {"deleted_measure": measure_name}
    except Exception as e:
        raise _session.wrap(e, f"Delete measure '{measure_name}' failed")


def _refresh(workbook: str | None) -> dict:
    wb, model = _get_model(workbook)
    try:
        model.Refresh()
        return {"refreshed": True, "workbook": wb.Name}
    except Exception as e:
        raise _session.wrap(e, "Model.Refresh() failed")
