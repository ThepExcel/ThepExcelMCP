"""ThepExcelMCP — FastMCP server entry point.

4 coarse action-dispatch tools; each tool routes to a domain module.
Tool docstrings are the LLM-facing API — kept precise and example-rich.
"""

from __future__ import annotations

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError  # noqa: F401 — re-exported for domain modules

from .domains.pivots import pivot_action
from .domains.powerquery import powerquery_action
from .domains.ranges import range_action
from .domains.sheets import sheet_action
from .domains.tables import table_action
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
) -> dict:
    """Read and write cell ranges.

    Parameters
    ----------
    action : str
        One of: ``read``, ``write``, ``write_formula``, ``clear``.
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

    Actions
    -------
    read
        Returns cell values as a 2-D list with pagination metadata.
        Strings longer than 500 chars are truncated. Response includes
        ``has_more`` and ``next_offset`` for continuation.
        Example: ``excel_range(action="read", range="A1:D100")``
    write
        Writes a 2-D list of values starting at the top-left of the range.
        Example: ``excel_range(action="write", range="A1", values=[[1,"Jan"],[2,"Feb"]])``
    write_formula
        Writes a dynamic-array formula via ``Formula2`` to the top-left cell.
        Excel spills results automatically (XLOOKUP, UNIQUE, FILTER, SORT, etc.).
        Example: ``excel_range(action="write_formula", range="E1", formula="=SORT(A1:A20)")``
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
        ``refresh``, ``refresh_all``, ``load_to_table``, ``analyze``, ``analyze_raw``.
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
        Lists all queries with name, description, line count.
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


def main() -> None:
    """Run the MCP server over stdio."""
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
