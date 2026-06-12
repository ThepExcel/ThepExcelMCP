"""ThepExcelMCP — FastMCP server entry point.

4 coarse action-dispatch tools; each tool routes to a domain module.
Tool docstrings are the LLM-facing API — kept precise and example-rich.
"""

from __future__ import annotations

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError  # noqa: F401 — re-exported for domain modules

from .domains.charts import chart_action
from .domains.datamodel import datamodel_action
from .domains.names import name_action
from .domains.pivots import pivot_action
from .domains.powerquery import powerquery_action
from .domains.ranges import range_action
from .domains.screenshot import screenshot_action
from .domains.sheets import sheet_action
from .domains.tables import table_action
from .domains.vba import vba_action
from .domains.workbook import workbook_action

mcp = FastMCP(
    "ThepExcelMCP",
    instructions=(
        "Controls a live running Excel Desktop instance via COM. "
        "Windows only. Excel must be open before calling any tool. "
        "Use excel_workbook(action='list') to discover open workbooks, "
        "then pass the workbook name to other tools."
    ),
)


@mcp.tool()
def excel_workbook(
    action: str,
    workbook: str | None = None,
    path: str | None = None,
) -> dict:
    """Manage Excel workbooks.

    Parameters
    ----------
    action : str
        One of: ``list``, ``info``, ``open``, ``save``, ``close``.
    workbook : str, optional
        Workbook name (e.g. ``"Sales.xlsx"``). Uses active workbook when omitted
        (not applicable to ``list`` and ``open``).
    path : str, optional
        Full file path. Required for ``open`` (e.g. ``"C:/data/Sales.xlsx"``).

    Actions
    -------
    list
        Returns all open workbooks with an active-flag.
        Example: ``excel_workbook(action="list")``
    info
        Returns sheet names, table/query/name counts, and saved status.
        Example: ``excel_workbook(action="info", workbook="Sales.xlsx")``
    open
        Opens a workbook from disk.
        Example: ``excel_workbook(action="open", path="C:/data/Sales.xlsx")``
    save
        Saves the workbook (equivalent to Ctrl+S).
    close
        Closes without saving. Use with caution — data loss is possible.
    """
    return workbook_action(action, workbook=workbook, path=path)


@mcp.tool()
def excel_sheet(
    action: str,
    name: str | None = None,
    workbook: str | None = None,
    new_name: str | None = None,
) -> dict:
    """Manage worksheets within a workbook.

    Parameters
    ----------
    action : str
        One of: ``list``, ``add``, ``rename``, ``delete``.
    name : str, optional
        Sheet name. Required for ``rename`` (current name) and ``delete``.
        Used as the new sheet's name for ``add`` (Excel auto-names if omitted).
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    new_name : str, optional
        New name for the sheet. Required for ``rename``.

    Actions
    -------
    list
        Returns all sheet names with an active-flag.
        Example: ``excel_sheet(action="list")``
    add
        Adds a new sheet after the last sheet.
        Example: ``excel_sheet(action="add", name="Summary")``
    rename
        Renames a sheet. Requires ``name`` and ``new_name``.
        Example: ``excel_sheet(action="rename", name="Sheet1", new_name="Data")``
    delete
        Deletes a sheet. Irreversible. Requires ``name``.
        Example: ``excel_sheet(action="delete", name="OldData")``
    """
    return sheet_action(action, name=name, workbook=workbook, new_name=new_name)


@mcp.tool()
def excel_range(
    action: str,
    range: str,
    sheet: str | None = None,
    workbook: str | None = None,
    values: list | None = None,
    formula: str | None = None,
    offset: int = 0,
    limit: int = 100,
    python_code: str | None = None,
) -> dict:
    """Read and write cell ranges.

    Parameters
    ----------
    action : str
        One of: ``read``, ``read_spill``, ``write``, ``write_formula``,
        ``write_py``, ``clear``.
    range : str
        Range address. Examples:
        - ``"A1:C10"`` — standard A1 notation (active sheet)
        - ``"Sheet1!A1:C10"`` — sheet-qualified (overrides ``sheet`` param)
        - ``"SalesTable[Amount]"`` — structured table column reference
        - ``"A1"`` — single cell
    sheet : str, optional
        Sheet name. Uses active sheet when omitted.
        Ignored when ``range`` already contains a sheet qualifier (``Sheet1!...``).
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    values : list of lists, optional
        2-D list for ``write``. Outer list = rows, inner list = columns.
        Example: ``[[1, "Jan"], [2, "Feb"]]``
    formula : str, optional
        Formula string for ``write_formula``. Must start with ``=``.
        Example: ``"=UNIQUE(A1:A100)"``
    offset : int, optional
        Row offset for ``read`` pagination (default 0).
    limit : int, optional
        Max rows to return per ``read`` call (default 100). Reduce for large ranges.
    python_code : str, optional
        Python source code string for ``write_py``. Multi-line supported
        (use ``\\n`` or actual newlines). Double-quote characters are
        automatically escaped for the Excel formula string.

    Actions
    -------
    read
        Returns cell values as a 2-D list with pagination metadata.
        Strings longer than 500 chars are truncated. Response includes
        ``has_more`` and ``next_offset`` for continuation.
        When the top-left cell is a spill anchor, the response also includes
        ``has_spill: true`` and ``spill_range`` (the spill address).
        When the top-left cell is part of someone else's spill, the response
        includes ``spill_parent`` (the anchor cell address).
        Example: ``excel_range(action="read", range="A1:D100")``
    read_spill
        Returns the complete spill range for a dynamic-array anchor cell
        (XLOOKUP, UNIQUE, FILTER, SORT, SEQUENCE, etc.).
        Raises an error if the cell has no spill (HasSpill=False).
        Paginated via ``offset`` / ``limit``.
        Example: ``excel_range(action="read_spill", range="E1")``
    write
        Writes a 2-D list of values starting at the top-left of the range.
        Example: ``excel_range(action="write", range="A1", values=[[1,"Jan"],[2,"Feb"]])``
    write_formula
        Writes a dynamic-array formula via ``Formula2`` to the top-left cell.
        Excel spills results automatically (XLOOKUP, UNIQUE, FILTER, SORT, etc.).
        Example: ``excel_range(action="write_formula", range="E1", formula="=SORT(A1:A20)")``
    write_py
        Insert a Python-in-Excel ``=PY()`` formula. Requires ``python_code``.
        IMPORTANT: execution is asynchronous in Azure cloud. The cell shows
        ``#BUSY!`` until Azure processes it. Requires M365 with Python in
        Excel enabled. Cannot await results — use ``read`` in a follow-up call.
        Experimental: COM insertion path is not officially documented by Microsoft.
        Office Scripts: not supported (cloud-only, no COM path) — use
        ``excel_vba`` instead.
        Example: ``excel_range(action="write_py", range="A1",
        python_code="import pandas as pd\\ndf = xl('A1:C10', headers=True)")``
    clear
        Clears cell contents (preserves formatting).
        Example: ``excel_range(action="clear", range="A1:Z100")``
    """
    return range_action(
        action,
        range=range,
        sheet=sheet,
        workbook=workbook,
        values=values,
        formula=formula,
        offset=offset,
        limit=limit,
        python_code=python_code,
    )


@mcp.tool()
def excel_powerquery(
    action: str,
    name: str | None = None,
    workbook: str | None = None,
    formula: str | None = None,
    description: str | None = None,
    new_name: str | None = None,
    sheet_name: str | None = None,
    raw_formula: str | None = None,
) -> dict:
    """Manage Power Query (M code) queries in Excel.

    Parameters
    ----------
    action : str
        One of: ``list``, ``get``, ``create``, ``update``, ``delete``,
        ``refresh``, ``refresh_all``, ``load_to_table``, ``load_to_datamodel``,
        ``analyze``, ``analyze_raw``.
    name : str, optional
        Query name. Required for most actions.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    formula : str, optional
        Complete M code (``let … in …`` expression) for ``create`` / ``update``.
    description : str, optional
        Human-readable description for ``create`` / ``update``.
    new_name : str, optional
        New query name for ``update`` (rename).
    sheet_name : str, optional
        Target sheet name for ``load_to_table``. Defaults to query name (max 31 chars).
    raw_formula : str, optional
        M code string for ``analyze_raw`` (no Excel connection required).

    Actions
    -------
    list
        Lists all queries with name, description, line count, and ``model_connection``
        flag when the query is loaded to the Data Model.
        Example: ``excel_powerquery(action="list")``
    get
        Returns full M code and metadata for a query.
        Example: ``excel_powerquery(action="get", name="SalesData")``
    create
        Creates a new query. ``name`` and ``formula`` are required.
        Example formula (let…in):
            ``let Source = Excel.CurrentWorkbook(){[Name="T1"]}[Content] in Source``
        Example: ``excel_powerquery(action="create", name="Q1", formula="let ... in ...")``
    update
        Updates formula, description, or name. Pass only fields to change.
        Example: ``excel_powerquery(action="update", name="Q1", new_name="SalesQuery")``
    delete
        Deletes a query permanently.
        Example: ``excel_powerquery(action="delete", name="OldQuery")``
    refresh
        Refreshes a specific query (synchronous via its connection).
        Example: ``excel_powerquery(action="refresh", name="SalesData")``
    refresh_all
        Refreshes all queries and connections in the workbook.
        Example: ``excel_powerquery(action="refresh_all")``
    load_to_table
        Loads a connection-only query to a new worksheet Table.
        Creates or replaces the target sheet. Uses Connections.Add2 + Mashup-OLEDB pattern.
        Example: ``excel_powerquery(action="load_to_table", name="SalesData", sheet_name="Output")``
    load_to_datamodel
        Loads a query directly to the Data Model (no worksheet table output).
        Uses Connections.Add2 with CreateModelConnection=True.
        After loading, the query appears in excel_datamodel(action="list_tables").
        Example: ``excel_powerquery(action="load_to_datamodel", name="SalesData")``
    analyze
        Analyzes M code from an existing query for anti-patterns
        (query folding, unnecessary buffers, hardcoded paths, etc.).
        Example: ``excel_powerquery(action="analyze", name="SalesData")``
    analyze_raw
        Analyzes M code provided directly — no Excel connection needed.
        Example: ``excel_powerquery(action="analyze_raw", raw_formula="let ... in ...")``
    """
    return powerquery_action(
        action,
        name=name,
        workbook=workbook,
        formula=formula,
        description=description,
        new_name=new_name,
        sheet_name=sheet_name,
        raw_formula=raw_formula,
    )


@mcp.tool()
def excel_table(
    action: str,
    name: str | None = None,
    workbook: str | None = None,
    sheet: str | None = None,
    range: str | None = None,
    style: str | None = None,
    has_headers: bool = True,
    columns: list | None = None,
    offset: int = 0,
    limit: int = 100,
    values: list | None = None,
    column_name: str | None = None,
    formula: str | None = None,
    sort_column: str | None = None,
    ascending: bool = True,
    filter_column: str | None = None,
    filter_op: str | None = None,
    filter_value: str | None = None,
    show_totals: bool | None = None,
    total_func: str | None = None,
    new_name: str | None = None,
    keep_data: bool = True,
) -> dict:
    """Manage Excel Tables (ListObjects) — create, read, sort, filter, and more.

    Parameters
    ----------
    action : str
        One of: ``list``, ``create``, ``read``, ``append_rows``, ``add_column``,
        ``sort``, ``filter``, ``set_style``, ``toggle_totals``, ``rename``, ``delete``.
    name : str, optional
        Table name (e.g. ``"SalesTable"``). Required for most actions.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    sheet : str, optional
        Sheet name for ``create``. Uses active sheet when omitted.
    range : str, optional
        Range address for ``create`` (e.g. ``"A1:D10"``).
    style : str, optional
        Table style name for ``create`` / ``set_style``
        (e.g. ``"TableStyleMedium2"``). Defaults to ``"TableStyleMedium2"``
        for ``create``.
    has_headers : bool, optional
        Whether the first row is a header row for ``create``. Default True.
    columns : list of str, optional
        Column names to include in ``read`` output. All columns when omitted.
    offset : int, optional
        Row offset for ``read`` pagination (default 0).
    limit : int, optional
        Max rows per ``read`` call (default 100).
    values : list of lists, optional
        2-D list of rows for ``append_rows``.
    column_name : str, optional
        New column name for ``add_column``. Also used as the target column
        for ``toggle_totals`` per-column aggregation.
    formula : str, optional
        Calculated column formula for ``add_column``
        (e.g. ``"=[@Amount]*0.1"``). Structured references auto-fill.
    sort_column : str, optional
        Column name to sort by (``sort`` action).
    ascending : bool, optional
        Sort order for ``sort`` action. Default True.
    filter_column : str, optional
        Column to filter on (``filter`` action).
    filter_op : str, optional
        Filter operation: ``equals``, ``contains``, ``greater``, ``less``,
        or ``clear_filters``.
    filter_value : str, optional
        Value to filter by (not required for ``clear_filters``).
    show_totals : bool, optional
        Show/hide totals row for ``toggle_totals``.
    total_func : str, optional
        Per-column aggregation for ``toggle_totals``:
        sum/average/count/max/min/stddev/var/none.
    new_name : str, optional
        New table name for ``rename``.
    keep_data : bool, optional
        For ``delete``: True (default) = unlink (keeps data), False = delete range.

    Actions
    -------
    list
        Returns all tables in the workbook.
        Example: ``excel_table(action="list")``
    create
        Creates a table from a range. Requires ``range`` and ``name``.
        Example: ``excel_table(action="create", range="A1:D10", name="SalesTable")``
    read
        Reads table data paginated. Requires ``name``.
        Tip: structured references like ``"SalesTable[Amount]"`` also work
        via ``excel_range(action="read", range="SalesTable[Amount]")``.
        Example: ``excel_table(action="read", name="SalesTable", limit=50)``
    append_rows
        Appends rows below the table. Requires ``name`` and ``values``.
        Example: ``excel_table(action="append_rows", name="SalesTable",
        values=[[101, "Jan", 5000]])``
    add_column
        Adds a column (with optional calculated formula). Requires ``name``
        and ``column_name``.
        Example: ``excel_table(action="add_column", name="SalesTable",
        column_name="Tax", formula="=[@Amount]*0.1")``
    sort
        Sorts table by a column. Requires ``name`` and ``sort_column``.
        Example: ``excel_table(action="sort", name="SalesTable",
        sort_column="Amount", ascending=False)``
    filter
        Applies an autofilter. Requires ``name``, ``filter_column``,
        ``filter_op`` (equals/contains/greater/less/clear_filters).
        Example: ``excel_table(action="filter", name="SalesTable",
        filter_column="Region", filter_op="equals", filter_value="North")``
    set_style
        Changes table style. Requires ``name`` and ``style``.
        Example: ``excel_table(action="set_style", name="SalesTable",
        style="TableStyleLight1")``
    toggle_totals
        Shows/hides totals row. Requires ``name`` and ``show_totals``.
        Example: ``excel_table(action="toggle_totals", name="SalesTable",
        show_totals=True, column_name="Amount", total_func="sum")``
    rename
        Renames the table. Requires ``name`` and ``new_name``.
        Example: ``excel_table(action="rename", name="Sales", new_name="SalesData")``
    delete
        Removes the table. Requires ``name``.
        ``keep_data=True`` (default) converts to normal range; ``keep_data=False``
        deletes the data range entirely.
        Example: ``excel_table(action="delete", name="OldTable")``
    """
    return table_action(
        action,
        name=name,
        workbook=workbook,
        sheet=sheet,
        range=range,
        style=style,
        has_headers=has_headers,
        columns=columns,
        offset=offset,
        limit=limit,
        values=values,
        column_name=column_name,
        formula=formula,
        sort_column=sort_column,
        ascending=ascending,
        filter_column=filter_column,
        filter_op=filter_op,
        filter_value=filter_value,
        show_totals=show_totals,
        total_func=total_func,
        new_name=new_name,
        keep_data=keep_data,
    )


@mcp.tool()
def excel_pivot(
    action: str,
    name: str | None = None,
    workbook: str | None = None,
    source: str | None = None,
    dest_sheet: str | None = None,
    dest_cell: str | None = None,
    field: str | None = None,
    area: str | None = None,
    aggregation: str | None = None,
    number_format: str | None = None,
    layout: str | None = None,
    subtotals: bool | None = None,
    grand_totals: bool | None = None,
    offset: int = 0,
    limit: int = 100,
) -> dict:
    """Manage Excel PivotTables — create, configure fields, layout, and read.

    Parameters
    ----------
    action : str
        One of: ``list``, ``create``, ``add_field``, ``remove_field``,
        ``move_field``, ``set_layout``, ``refresh``, ``delete``, ``read``.
    name : str, optional
        PivotTable name. Required for most actions.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    source : str, optional
        Data source for ``create``:
        - Range address: ``"Sheet1!A1:D100"``
        - Table name: ``"SalesTable"`` (resolved automatically)
        - ``"datamodel"`` for the workbook Data Model
    dest_sheet : str, optional
        Destination sheet for ``create``. Creates a new sheet when omitted.
    dest_cell : str, optional
        Top-left cell for the pivot (default ``"A3"``).
    field : str, optional
        Field name for ``add_field`` / ``remove_field`` / ``move_field``.
    area : str, optional
        Target area: ``rows``, ``columns``, ``filters``, or ``values``.
    aggregation : str, optional
        Aggregation for ``values`` area:
        sum/count/average/max/min (default sum).
    number_format : str, optional
        Number format code for value fields (e.g. ``"#,##0.00"``).
    layout : str, optional
        Layout for ``set_layout``: ``compact``, ``outline``, or ``tabular``.
    subtotals : bool, optional
        Show/hide subtotals on all row fields (``set_layout``).
    grand_totals : bool, optional
        Show/hide grand totals rows and columns (``set_layout``).
    offset : int, optional
        Row offset for ``read`` pagination (default 0).
    limit : int, optional
        Max rows per ``read`` call (default 100).

    Actions
    -------
    list
        Returns all PivotTables: name, sheet, location, source, fields by area.
        Example: ``excel_pivot(action="list")``
    create
        Creates a PivotTable. Requires ``source`` and ``name``.
        For a table source: ``excel_pivot(action="create",
        source="SalesTable", name="SalesPivot")``
        For a range: ``excel_pivot(action="create",
        source="Sheet1!A1:D100", name="SalesPivot")``
    add_field
        Adds a field to an area. Requires ``name``, ``field``, ``area``.
        Example: ``excel_pivot(action="add_field", name="SalesPivot",
        field="Region", area="rows")``
        Example: ``excel_pivot(action="add_field", name="SalesPivot",
        field="Amount", area="values", aggregation="sum",
        number_format="#,##0.00")``
    remove_field
        Removes a field from its current area. Requires ``name`` and ``field``.
        Example: ``excel_pivot(action="remove_field", name="SalesPivot",
        field="Region")``
    move_field
        Moves a field to a different area. Requires ``name``, ``field``, ``area``.
        Example: ``excel_pivot(action="move_field", name="SalesPivot",
        field="Quarter", area="columns")``
    set_layout
        Configures layout. Requires ``name``. Pass any combination of
        ``layout``, ``subtotals``, ``grand_totals``.
        Example: ``excel_pivot(action="set_layout", name="SalesPivot",
        layout="tabular", grand_totals=False)``
    refresh
        Refreshes the PivotTable from its source data. Requires ``name``.
        Example: ``excel_pivot(action="refresh", name="SalesPivot")``
    delete
        Deletes the PivotTable. Requires ``name``.
        Example: ``excel_pivot(action="delete", name="SalesPivot")``
    read
        Reads pivot result values paginated. Requires ``name``.
        Example: ``excel_pivot(action="read", name="SalesPivot", limit=50)``
    """
    return pivot_action(
        action,
        name=name,
        workbook=workbook,
        source=source,
        dest_sheet=dest_sheet,
        dest_cell=dest_cell,
        field=field,
        area=area,
        aggregation=aggregation,
        number_format=number_format,
        layout=layout,
        subtotals=subtotals,
        grand_totals=grand_totals,
        offset=offset,
        limit=limit,
    )


@mcp.tool()
def excel_datamodel(
    action: str,
    workbook: str | None = None,
    # add_table
    source_type: str | None = None,
    source_name: str | None = None,
    # relationships
    from_table: str | None = None,
    from_column: str | None = None,
    to_table: str | None = None,
    to_column: str | None = None,
    relationship_index: int | None = None,
    # measures
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
    """Manage the Excel Data Model (Power Pivot) — tables, relationships, and DAX measures.

    Parameters
    ----------
    action : str
        One of: ``info``, ``list_tables``, ``add_table``, ``list_relationships``,
        ``add_relationship``, ``delete_relationship``, ``list_measures``,
        ``add_measure``, ``update_measure``, ``delete_measure``, ``refresh``.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    source_type : str, optional
        For ``add_table``: ``"table"`` (ListObject) or ``"query"`` (Power Query).
    source_name : str, optional
        Name of the table or query to add to the model.
    from_table : str, optional
        FK (many-side) table name for relationship operations.
    from_column : str, optional
        FK column name in ``from_table``.
    to_table : str, optional
        PK (one-side) table name for relationship operations.
    to_column : str, optional
        PK column name in ``to_table``.
    relationship_index : int, optional
        1-based index for ``delete_relationship``.
    measure_name : str, optional
        Measure name. Required for measure operations.
    table : str, optional
        Model table to associate the measure with (``add_measure``).
    formula : str, optional
        DAX formula string (e.g. ``"=SUM(Sales[Amount])"``) for ``add_measure``.
    format_type : str, optional
        Measure format: ``general``, ``decimal``, ``number``, ``currency``,
        ``percent``, ``whole``, ``integer``, ``boolean``, ``date``, ``scientific``.
        Default: ``general``.
    decimal_places : int, optional
        Decimal places for decimal/whole/currency formats.
    use_thousand_sep : bool, optional
        Thousand separator for decimal/whole/currency formats.
    description : str, optional
        Measure description for ``add_measure``.
    new_formula : str, optional
        New DAX formula for ``update_measure``.
    new_format_type : str, optional
        New format type for ``update_measure``.
    new_description : str, optional
        New description for ``update_measure``.

    Actions
    -------
    info
        Model summary: table count, relationship count, measure count.
        Example: ``excel_datamodel(action="info")``
    list_tables
        All model tables with names, source, record count, columns.
        Example: ``excel_datamodel(action="list_tables")``
    add_table
        Add an existing workbook Table or Power Query to the Data Model.
        Uses Connections.Add2 with CreateModelConnection=True.
        Requires ``source_type`` and ``source_name``.
        Example: ``excel_datamodel(action="add_table",
        source_type="table", source_name="SalesTable")``
        Example: ``excel_datamodel(action="add_table",
        source_type="query", source_name="SalesQuery")``
    list_relationships
        All relationships: FK table.column → PK table.column, active flag.
        Example: ``excel_datamodel(action="list_relationships")``
    add_relationship
        Create a one-to-many relationship. Requires ``from_table``,
        ``from_column`` (FK/many side), ``to_table``, ``to_column`` (PK/one side).
        Example: ``excel_datamodel(action="add_relationship",
        from_table="Sales", from_column="ProductID",
        to_table="Products", to_column="ProductID")``
    delete_relationship
        Delete a relationship by 1-based index from ``list_relationships``.
        Example: ``excel_datamodel(action="delete_relationship",
        relationship_index=1)``
    list_measures
        All DAX measures: name, table, formula, format, description.
        Example: ``excel_datamodel(action="list_measures")``
    add_measure
        Add a DAX measure. Requires ``measure_name``, ``table``, ``formula``.
        Example: ``excel_datamodel(action="add_measure",
        measure_name="Total Sales", table="Sales",
        formula="=SUM(Sales[Amount])", format_type="currency",
        decimal_places=2, use_thousand_sep=True)``
    update_measure
        Update formula, format, or description. Requires ``measure_name``.
        Pass only the fields to change.
        Example: ``excel_datamodel(action="update_measure",
        measure_name="Total Sales", new_formula="=SUMX(Sales, Sales[Amount])")``
    delete_measure
        Delete a measure by name. Requires ``measure_name``.
        Example: ``excel_datamodel(action="delete_measure",
        measure_name="OldMeasure")``
    refresh
        Refresh all Data Model data sources (Model.Refresh()).
        Example: ``excel_datamodel(action="refresh")``
    """
    return datamodel_action(
        action,
        workbook=workbook,
        source_type=source_type,
        source_name=source_name,
        from_table=from_table,
        from_column=from_column,
        to_table=to_table,
        to_column=to_column,
        relationship_index=relationship_index,
        measure_name=measure_name,
        table=table,
        formula=formula,
        format_type=format_type,
        decimal_places=decimal_places,
        use_thousand_sep=use_thousand_sep,
        description=description,
        new_formula=new_formula,
        new_format_type=new_format_type,
        new_description=new_description,
    )


@mcp.tool()
def excel_vba(
    action: str,
    workbook: str | None = None,
    module_name: str | None = None,
    code: str | None = None,
    proc_name: str | None = None,
    args: list | None = None,
) -> dict:
    """Execute and manage VBA modules in Excel.

    Security: Entire tool disabled unless THEPEXCEL_MCP_ENABLE_VBA=1 is set.
    A pre-flight check also verifies that Excel's "Trust access to the VBA
    project object model" setting is enabled (AccessVBOM=1 in registry).

    Parameters
    ----------
    action : str
        One of: ``list_modules``, ``get_module``, ``write_module``,
        ``delete_module``, ``run``.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    module_name : str, optional
        VBA module name for get/write/delete operations.
    code : str, optional
        Full VBA source code for ``write_module``.
    proc_name : str, optional
        Macro name for ``run``. Format: ``"Module1.ProcName"`` or ``"ProcName"``.
    args : list, optional
        Positional arguments for ``run`` (scalars only: str/int/float/bool/None).

    Actions
    -------
    list_modules
        List all VBA modules: name, type, line count.
        Types: standard, class, form, document.
        Example: ``excel_vba(action="list_modules")``
    get_module
        Return full source code of a module. Requires ``module_name``.
        Example: ``excel_vba(action="get_module", module_name="Module1")``
    write_module
        Create or replace a standard module. Requires ``module_name`` and ``code``.
        If the module exists, all lines are replaced atomically.
        Note: to save a workbook with VBA, save as .xlsm or .xlsb — COM cannot
        coerce a .xlsx file to a macro-enabled format automatically.
        Example: ``excel_vba(action="write_module", module_name="Helpers",
        code="Sub Hello()\\n  MsgBox \\"Hi\\"\\nEnd Sub")``
    delete_module
        Delete a standard module. Requires ``module_name``.
        Document modules (Sheet*, ThisWorkbook) cannot be deleted.
        Example: ``excel_vba(action="delete_module", module_name="OldModule")``
    run
        Execute a macro via Application.Run. Requires ``proc_name``.
        Functions return their value; Subs return null.
        Example: ``excel_vba(action="run", proc_name="Module1.CalcTax",
        args=[1000, 0.07])``
    """
    return vba_action(
        action,
        workbook=workbook,
        module_name=module_name,
        code=code,
        proc_name=proc_name,
        args=args,
    )


@mcp.tool()
def excel_name(
    action: str,
    workbook: str | None = None,
    name: str | None = None,
    refers_to: str | None = None,
    scope: str | None = None,
) -> dict:
    """Manage defined names: named ranges, constants, and LAMBDA formulas.

    Parameters
    ----------
    action : str
        One of: ``list``, ``get``, ``set``, ``delete``.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    name : str, optional
        Name to get/set/delete. Case-insensitive.
    refers_to : str, optional
        Definition for ``set``. Must start with ``=``.
        Named range: ``"=Sheet1!$A$1:$B$10"``
        LAMBDA: ``"=LAMBDA(x, x*2)"``
        Constant: ``"=42"``
    scope : str, optional
        ``"workbook"`` (default) for workbook-scope, or a sheet name for
        sheet-local scope. Sheet-local names shadow workbook names on that sheet.

    Actions
    -------
    list
        All defined names with: name, refers_to, scope, is_lambda flag.
        ``is_lambda=true`` when RefersTo starts with ``=LAMBDA(``.
        Example: ``excel_name(action="list")``
    get
        Details for a single name. Requires ``name``.
        Example: ``excel_name(action="get", name="SalesTotal")``
    set
        Create or update a name. Requires ``name`` and ``refers_to``.
        Example (named range):
            ``excel_name(action="set", name="MyRange",
            refers_to="=Sheet1!$A$1:$C$10")``
        Example (LAMBDA — available as =DOUBLE(A1) in cells):
            ``excel_name(action="set", name="DOUBLE",
            refers_to="=LAMBDA(x, x*2)")``
        Example (LAMBDA with multiple params):
            ``excel_name(action="set", name="TAX",
            refers_to='=LAMBDA(amount, rate, amount*rate/100)')``
    delete
        Delete a name. Requires ``name``.
        Example: ``excel_name(action="delete", name="OldName")``
    """
    return name_action(
        action,
        workbook=workbook,
        name=name,
        refers_to=refers_to,
        scope=scope,
    )


@mcp.tool()
def excel_chart(
    action: str,
    name: str | None = None,
    workbook: str | None = None,
    sheet: str | None = None,
    source: str | None = None,
    chart_type: str | None = None,
    position: str | None = None,
    width: float | None = None,
    height: float | None = None,
    title: str | None = None,
    dest_sheet: str | None = None,
    x_title: str | None = None,
    y_title: str | None = None,
    legend: bool | None = None,
    legend_position: str | None = None,
    data_labels: bool | None = None,
    series_names: list | None = None,
    secondary_series: str | None = None,
    output_path: str | None = None,
) -> dict:
    """Create and manage Excel charts (ChartObjects / embedded charts).

    Parameters
    ----------
    action : str
        One of: ``list``, ``create``, ``configure``, ``set_source``,
        ``export_image``, ``delete``.
    name : str, optional
        Chart name. Required for ``configure``, ``set_source``,
        ``export_image``, ``delete``.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    sheet : str, optional
        Source sheet for ``create``. Uses active sheet when omitted.
    source : str, optional
        Data source for ``create`` / ``set_source``: range address
        (``"A1:C10"``), table name, or pivot table name.
    chart_type : str, optional
        Chart type for ``create``. Common values: ``column``, ``bar``,
        ``line``, ``pie``, ``doughnut``, ``scatter``, ``area``,
        ``combo_column_line``. Full list in domain docstring.
    position : str, optional
        Anchor cell for chart top-left corner (e.g. ``"E2"``).
    width : float, optional
        Chart width in points. Default 375.
    height : float, optional
        Chart height in points. Default 225.
    title : str, optional
        Chart title for ``create`` / ``configure``.
    dest_sheet : str, optional
        Sheet to embed chart on (``create``). Defaults to same sheet as source.
    x_title : str, optional
        Category (X) axis title for ``configure``.
    y_title : str, optional
        Value (Y) axis title for ``configure``.
    legend : bool, optional
        Show/hide legend for ``configure``.
    legend_position : str, optional
        Legend position: ``bottom``, ``corner``, ``left``, ``right``, ``top``.
    data_labels : bool, optional
        Show/hide data labels on all series for ``configure``.
    series_names : list of str, optional
        Override series names (in order) for ``configure``.
    secondary_series : str, optional
        Series name to move to secondary Y axis for ``configure`` (combo charts).
    output_path : str, optional
        File path for ``export_image``. Defaults to
        ``%TEMP%/thepexcel_mcp/<chart_name>.png``.

    Actions
    -------
    list
        All charts: name, sheet, chart_type, source, position.
        Example: ``excel_chart(action="list")``
    create
        Create a chart. Requires ``source`` and ``chart_type``.
        Example: ``excel_chart(action="create", source="A1:C10",
        chart_type="column", title="Sales")``
    configure
        Update title, axis labels, legend, data labels, series names.
        Requires ``name``. Pass only parameters to change.
        Example: ``excel_chart(action="configure", name="Chart 1",
        title="Updated Title", legend=True, legend_position="bottom")``
    set_source
        Change data range. Requires ``name`` and ``source``.
        Example: ``excel_chart(action="set_source", name="Chart 1",
        source="A1:D20")``
    export_image
        Export chart as PNG. Requires ``name``. Returns file path.
        Example: ``excel_chart(action="export_image", name="Chart 1")``
    delete
        Delete chart. Requires ``name``.
        Example: ``excel_chart(action="delete", name="Chart 1")``
    """
    return chart_action(
        action,
        name=name,
        workbook=workbook,
        sheet=sheet,
        source=source,
        chart_type=chart_type,
        position=position,
        width=width,
        height=height,
        title=title,
        dest_sheet=dest_sheet,
        x_title=x_title,
        y_title=y_title,
        legend=legend,
        legend_position=legend_position,
        data_labels=data_labels,
        series_names=series_names,
        secondary_series=secondary_series,
        output_path=output_path,
    )


@mcp.tool()
def excel_screenshot(
    action: str,
    range: str | None = None,
    sheet: str | None = None,
    workbook: str | None = None,
    output_path: str | None = None,
    name: str | None = None,
) -> dict:
    """Capture a region of the live Excel workbook as a PNG for visual verification.

    Primary use: let an AI agent visually verify what Excel looks like after
    writing data, creating charts, or applying formatting.

    Requires Excel to be visible (not minimized). On HiDPI screens the bitmap
    captures at screen resolution; output size scales with DPI.

    Parameters
    ----------
    action : str
        One of: ``range``, ``sheet``, ``chart``.
    range : str, optional
        Range address for ``range`` action (e.g. ``"A1:F20"``).
    sheet : str, optional
        Sheet name. Uses active sheet when omitted (``range`` / ``sheet``).
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    output_path : str, optional
        Save path. Defaults to ``%TEMP%/thepexcel_mcp/<timestamp>.png``.
    name : str, optional
        Chart name for ``chart`` action.

    Actions
    -------
    range
        Capture a cell range as PNG. Requires ``range``.
        Uses Range.CopyPicture (bitmap mode) -> PIL clipboard grab -> PNG.
        Example: ``excel_screenshot(action="range", range="A1:F20")``
    sheet
        Capture the used range of a sheet.
        Example: ``excel_screenshot(action="sheet", sheet="Summary")``
    chart
        Export a chart as PNG by name. Requires ``name``.
        Uses Chart.Export (no clipboard involved).
        Example: ``excel_screenshot(action="chart", name="Sales Chart")``
    """
    return screenshot_action(
        action,
        range=range,
        sheet=sheet,
        workbook=workbook,
        output_path=output_path,
        name=name,
    )


def main() -> None:
    """Run the MCP server over stdio."""
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
