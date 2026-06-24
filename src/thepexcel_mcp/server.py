"""ThepExcelMCP — FastMCP server entry point.

4 coarse action-dispatch tools; each tool routes to a domain module.
Tool docstrings are the LLM-facing API — kept precise and example-rich.
"""

from __future__ import annotations

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError  # noqa: F401 — re-exported for domain modules

from .domains.charts import chart_action
from .domains.comments import comment_action
from .domains.conditional_format import conditional_format_action
from .domains.datamodel import datamodel_action
from .domains.diff import diff_action
from .domains.find_replace import find_replace_action
from .domains.format import format_action
from .domains.hyperlinks import hyperlink_action
from .domains.names import name_action
from .domains.outline import outline_action
from .domains.page_setup import page_setup_action
from .domains.pivots import pivot_action
from .domains.powerquery import powerquery_action
from .domains.protection import protection_action
from .domains.ranges import range_action
from .domains.screenshot import screenshot_action
from .domains.shapes import shape_action
from .domains.sheets import sheet_action
from .domains.slicer import slicer_action
from .domains.snapshot import snapshot_action
from .domains.sparkline import sparkline_action
from .domains.tables import table_action
from .domains.validation import validation_action
from .domains.vba import vba_action
from .domains.view import view_action
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
        One of: ``list``, ``info``, ``open``, ``save``, ``close``,
        ``create``, ``save_as``.
    workbook : str, optional
        Workbook name (e.g. ``"Sales.xlsx"``). Uses active workbook when omitted
        (not applicable to ``list``, ``open``, and ``create``).
    path : str, optional
        Full file path. Required for ``open`` and ``save_as``.
        Optional for ``create`` (if given, SaveAs is called immediately).
        Example: ``"C:/data/Sales.xlsx"``

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
    create
        Creates a new blank workbook (Workbooks.Add()). If ``path`` is given,
        immediately SaveAs to that path. File format inferred from extension
        (.xlsx=51, .xlsm=52, .xlsb=50, .xls=56, .csv=6).
        Example: ``excel_workbook(action="create")``
        Example: ``excel_workbook(action="create", path="C:/data/New.xlsx")``
    save_as
        Saves the active/named workbook to a new file path. File format is
        inferred from the extension. Suppresses overwrite prompts automatically.
        Example: ``excel_workbook(action="save_as", path="C:/data/Copy.xlsx")``
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
    value=None,
    param_type: str | None = None,
    required: bool = True,
) -> dict:
    """Manage Power Query (M code) queries in Excel.

    Parameters
    ----------
    action : str
        One of: ``list``, ``get``, ``create``, ``update``, ``delete``,
        ``refresh``, ``refresh_all``, ``load_to_table``, ``load_to_datamodel``,
        ``analyze``, ``analyze_raw``, ``create_parameter``, ``get_parameter``,
        ``set_parameter``, ``list_parameters``.
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
    value : int | float | str, optional
        Scalar value for the PQ parameter (Number or Text). Required for
        ``create_parameter`` and ``set_parameter``.
    param_type : str, optional
        Force the M type string: ``"Text"`` or ``"Number"``. When omitted, inferred
        from the Python type of ``value`` (int/float → Number, str → Text).
    required : bool, optional
        Sets ``IsParameterQueryRequired`` in the meta record. Defaults ``True``.

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
    create_parameter
        Creates a PQ parameter query (scalar meta record). Requires ``name`` and
        ``value``. Optional ``param_type`` (``"Text"`` | ``"Number"``); inferred
        when omitted.
        Example: ``excel_powerquery(action="create_parameter", name="MaxRows", value=1000)``
    get_parameter
        Returns parsed ``{value, type}`` for a parameter query. Requires ``name``.
        Raises ToolError if the query is not a parameter.
        Example: ``excel_powerquery(action="get_parameter", name="MaxRows")``
    set_parameter
        Updates the scalar value of an existing parameter query. Requires ``name``
        and ``value``. Pass ``param_type`` to coerce; otherwise existing type is kept.
        Example: ``excel_powerquery(action="set_parameter", name="MaxRows", value=500)``
    list_parameters
        Lists all parameter queries in the workbook (``IsParameterQuery=true``).
        Returns ``[{name, value, type}]``.
        Example: ``excel_powerquery(action="list_parameters")``
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
        value=value,
        param_type=param_type,
        required=required,
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
    # cube_formula / cube_value / cube_member
    target_cell: str | None = None,
    measure: str | None = None,
    members: list | None = None,
    kind: str = "cubevalue",
    member_expression: str | None = None,
    caption: str | None = None,
    connection: str = "ThisWorkbookDataModel",
    sheet: str | None = None,
) -> dict:
    """Manage the Excel Data Model (Power Pivot) — tables, relationships, and DAX measures.

    Parameters
    ----------
    action : str
        One of: ``info``, ``list_tables``, ``add_table``, ``list_relationships``,
        ``add_relationship``, ``delete_relationship``, ``list_measures``,
        ``add_measure``, ``update_measure``, ``delete_measure``, ``refresh``,
        ``cube_formula``, ``cube_value``, ``cube_member``,
        ``add_calculated_column``, ``add_calculated_table``.
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
    target_cell : str, optional
        Cell address (e.g. ``"B2"``) to write the formula into for
        ``cube_value`` / ``cube_member``.
    measure : str, optional
        Friendly or MDX measure name for ``cube_formula`` (kind=cubevalue)
        and ``cube_value``.
    members : list, optional
        Optional list of MDX member-expression strings (slicers/tuples) for
        CUBEVALUE (``cube_formula`` / ``cube_value``).
    kind : str, optional
        Formula kind for ``cube_formula``: ``"cubevalue"`` (default) or
        ``"cubemember"``.
    member_expression : str, optional
        Full MDX member expression for ``cube_formula`` (kind=cubemember) and
        ``cube_member``.
    caption : str, optional
        Optional display caption for CUBEMEMBER (3rd argument).
    connection : str, optional
        Data-model connection literal. Defaults to ``"ThisWorkbookDataModel"``.
    sheet : str, optional
        Sheet name for ``cube_value`` / ``cube_member``; ``None`` = active sheet.

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
    cube_formula
        Build and return a CUBEVALUE or CUBEMEMBER formula string (pure Python,
        no COM). Use ``kind`` to select: ``"cubevalue"`` (default) or
        ``"cubemember"``. For cubevalue, requires ``measure``; optional ``members``
        list of MDX tuples. For cubemember, requires ``member_expression``; optional
        ``caption``.
        Example: ``excel_datamodel(action="cube_formula", kind="cubevalue",
        measure="[Measures].[Total Sales]",
        members=["[Date].[Year].&[2024]"])``
    cube_value
        Write a CUBEVALUE formula to a cell and return the resolved value.
        Requires ``target_cell`` and ``measure``; optional ``members``, ``sheet``,
        ``connection``.
        Example: ``excel_datamodel(action="cube_value", target_cell="B2",
        measure="[Measures].[Total Sales]")``
    cube_member
        Write a CUBEMEMBER formula to a cell and return the resolved value.
        Requires ``target_cell`` and ``member_expression``; optional ``caption``,
        ``sheet``, ``connection``.
        Example: ``excel_datamodel(action="cube_member", target_cell="A2",
        member_expression="[Date].[Year].&[2024]")``
    add_calculated_column
        Not supported — calculated columns require Power BI / Analysis Services.
        Use ``add_measure`` for DAX aggregations instead. Always raises ToolError.
        Example: ``excel_datamodel(action="add_calculated_column")``
    add_calculated_table
        Not supported — calculated tables require Power BI / Analysis Services.
        Always raises ToolError.
        Example: ``excel_datamodel(action="add_calculated_table")``
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
        target_cell=target_cell,
        measure=measure,
        members=members,
        kind=kind,
        member_expression=member_expression,
        caption=caption,
        connection=connection,
        sheet=sheet,
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


@mcp.tool()
def excel_format(
    action: str,
    range: str,
    sheet: str | None = None,
    workbook: str | None = None,
    # font
    font_name: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    font_color: str | None = None,
    # fill
    fill_color: str | None = None,
    clear_fill: bool = False,
    # border
    border_sides: str = "outline",
    border_style: str = "continuous",
    border_weight: str = "thin",
    border_color: str | None = None,
    # number_format
    number_format: str | None = None,
    # alignment
    horizontal: str | None = None,
    vertical: str | None = None,
    wrap_text: bool | None = None,
    merge: bool | None = None,
    # column/row sizing
    width: float | None = None,
    height: float | None = None,
    # autofit
    autofit_columns: bool = True,
    autofit_rows: bool = False,
) -> dict:
    """Apply formatting to any Excel range.

    All color parameters accept ``"#RRGGBB"`` hex strings (converted to Excel
    BGR internally). No existing data or formulas are affected — formatting only.

    Parameters
    ----------
    action : str
        One of: ``font``, ``fill``, ``border``, ``number_format``,
        ``alignment``, ``column_width``, ``row_height``, ``autofit``.
    range : str
        Range address. Same syntax as ``excel_range``:
        ``"A1:C10"``, ``"Sheet1!A1:C10"``, ``"SalesTable[Amount]"``.
    sheet : str, optional
        Sheet name. Uses active sheet when omitted.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    font_name : str, optional
        Font family name for ``font`` action (e.g. ``"Calibri"``).
    font_size : float, optional
        Font size in points for ``font`` action (e.g. ``12.0``).
    bold : bool, optional
        Bold on/off for ``font`` action.
    italic : bool, optional
        Italic on/off for ``font`` action.
    underline : bool, optional
        Underline on/off for ``font`` action (single underline).
    font_color : str, optional
        Font color ``"#RRGGBB"`` for ``font`` action (e.g. ``"#FF0000"``).
    fill_color : str, optional
        Background color ``"#RRGGBB"`` for ``fill`` action.
    clear_fill : bool, optional
        When True, removes the background fill (``fill`` action).
    border_sides : str, optional
        Which sides to apply border (``border`` action):
        ``outline`` (default), ``all``, ``inside``, ``top``, ``bottom``,
        ``left``, ``right``.
    border_style : str, optional
        Line style for ``border`` action:
        ``continuous`` (default), ``dash``, ``double``, ``none``.
    border_weight : str, optional
        Line weight for ``border`` action:
        ``thin`` (default), ``medium``, ``thick``.
    border_color : str, optional
        Border color ``"#RRGGBB"`` for ``border`` action.
    number_format : str, optional
        Excel NumberFormat code for ``number_format`` action.
        Examples: ``"#,##0.00"``, ``"0.0%"``, ``"yyyy-mm-dd"``,
        ``"$#,##0.00"``.
    horizontal : str, optional
        Horizontal alignment for ``alignment`` action:
        ``general``, ``left``, ``center``, ``right``.
    vertical : str, optional
        Vertical alignment for ``alignment`` action:
        ``top``, ``center``, ``bottom``.
    wrap_text : bool, optional
        Wrap text on/off for ``alignment`` action.
    merge : bool, optional
        True = merge cells, False = unmerge, None = no change.
        For ``alignment`` action.
    width : float, optional
        Column width in character-width units for ``column_width`` action.
    height : float, optional
        Row height in points for ``row_height`` action.
    autofit_columns : bool, optional
        AutoFit column widths (default True for ``autofit`` action).
    autofit_rows : bool, optional
        AutoFit row heights (default False for ``autofit`` action).

    Actions
    -------
    font
        Set font properties. Pass any combination of font_name, font_size,
        bold, italic, underline, font_color.
        Example: ``excel_format(action="font", range="A1:C1",
        bold=True, font_size=14, font_color="#FFFFFF")``
    fill
        Set background color or clear fill.
        Example: ``excel_format(action="fill", range="A1:C1",
        fill_color="#D4A84B")``
        Example: ``excel_format(action="fill", range="A1:C1", clear_fill=True)``
    border
        Apply borders. Defaults: sides=outline, style=continuous, weight=thin.
        Example: ``excel_format(action="border", range="A1:D10",
        border_sides="all", border_weight="thin")``
        Example: ``excel_format(action="border", range="A1:D10",
        border_sides="outline", border_weight="medium", border_color="#000000")``
    number_format
        Set a NumberFormat code string.
        Example: ``excel_format(action="number_format", range="D2:D100",
        number_format="#,##0.00")``
        Example: ``excel_format(action="number_format", range="E2:E100",
        number_format="0.0%")``
    alignment
        Set horizontal/vertical alignment, wrap_text, or merge cells.
        Example: ``excel_format(action="alignment", range="A1:D1",
        horizontal="center", vertical="center", wrap_text=True)``
        Example: ``excel_format(action="alignment", range="A1:D1", merge=True)``
    column_width
        Set explicit column width in character-width units.
        Example: ``excel_format(action="column_width", range="A:A", width=20)``
    row_height
        Set explicit row height in points.
        Example: ``excel_format(action="row_height", range="1:1", height=30)``
    autofit
        AutoFit columns and/or rows to content.
        Example: ``excel_format(action="autofit", range="A:D")``
        Example: ``excel_format(action="autofit", range="A1:D20",
        autofit_columns=True, autofit_rows=True)``
    """
    return format_action(
        action,
        range=range,
        sheet=sheet,
        workbook=workbook,
        font_name=font_name,
        font_size=font_size,
        bold=bold,
        italic=italic,
        underline=underline,
        font_color=font_color,
        fill_color=fill_color,
        clear_fill=clear_fill,
        border_sides=border_sides,
        border_style=border_style,
        border_weight=border_weight,
        border_color=border_color,
        number_format=number_format,
        horizontal=horizontal,
        vertical=vertical,
        wrap_text=wrap_text,
        merge=merge,
        width=width,
        height=height,
        autofit_columns=autofit_columns,
        autofit_rows=autofit_rows,
    )


@mcp.tool()
def excel_view(
    action: str,
    sheet: str | None = None,
    workbook: str | None = None,
    # freeze_panes
    cell: str | None = None,
    freeze_rows: int | None = None,
    freeze_cols: int | None = None,
    # gridlines / headings
    show: bool | None = None,
    # zoom
    zoom: int | None = None,
) -> dict:
    """Control worksheet display settings: freeze panes, gridlines, zoom, headings.

    All mutations target the workbook-scoped window (``wb.Windows(1)``), not
    the global ``Application.ActiveWindow``, so results are correct even when
    the target workbook is not the foreground workbook.

    Parameters
    ----------
    action : str
        One of: ``freeze_panes``, ``unfreeze_panes``, ``gridlines``,
        ``zoom``, ``headings``.
    sheet : str, optional
        Sheet name. Uses active sheet when omitted.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    cell : str, optional
        Anchor cell for ``freeze_panes`` (e.g. ``"B2"`` freezes row 1 and
        column A). Overrides freeze_rows/freeze_cols when provided.
    freeze_rows : int, optional
        Number of rows to freeze from the top (``freeze_panes``).
    freeze_cols : int, optional
        Number of columns to freeze from the left (``freeze_panes``).
    show : bool, optional
        True = show, False = hide for ``gridlines`` and ``headings`` actions.
    zoom : int, optional
        Zoom percentage 10–400 for ``zoom`` action (e.g. 150 for 150%).

    Actions
    -------
    freeze_panes
        Freeze rows/columns. Supply ``cell="B2"`` (row 1 + col A frozen),
        or ``freeze_rows=2``, ``freeze_cols=1`` directly.
        Examples::

            excel_view(action="freeze_panes", sheet="Sales", cell="B2")
            excel_view(action="freeze_panes", workbook="Report.xlsx", freeze_rows=1)

    unfreeze_panes
        Remove any freeze from the target sheet's window.
        Example::

            excel_view(action="unfreeze_panes", sheet="Sheet1")

    gridlines
        Show or hide gridlines on the target sheet.
        Examples::

            excel_view(action="gridlines", show=False)   # hide
            excel_view(action="gridlines", show=True)    # show

    zoom
        Set zoom level 10–400.
        Examples::

            excel_view(action="zoom", sheet="Dashboard", zoom=80)
            excel_view(action="zoom", zoom=150)

    headings
        Show or hide row/column headings (the 1/2/3… A/B/C… labels).
        Examples::

            excel_view(action="headings", show=False)
            excel_view(action="headings", sheet="Print", show=True)
    """
    return view_action(
        action,
        sheet=sheet,
        workbook=workbook,
        cell=cell,
        freeze_rows=freeze_rows,
        freeze_cols=freeze_cols,
        show=show,
        zoom=zoom,
    )


@mcp.tool()
def excel_conditional_format(
    action: str,
    range: str,
    sheet: str | None = None,
    workbook: str | None = None,
    # data_bar
    color: str | None = None,
    # color_scale
    scale_type: int = 3,
    # icon_set
    style: str = "3traffic_lights",
    # cell_rule
    operator: str | None = None,
    formula1: str | None = None,
    formula2: str | None = None,
    fill_color: str | None = None,
    font_color: str | None = None,
    # top_bottom
    kind: str = "top",
    rank: int = 10,
    percent: bool = False,
) -> dict:
    """Add or remove conditional formatting rules on an Excel range via COM.

    Parameters
    ----------
    action : str
        One of: cell_rule | data_bar | color_scale | icon_set | top_bottom | clear
    range : str
        Excel range address, e.g. "A1:D20" or "TableName[Column]".
    sheet : str, optional
        Sheet name. Defaults to the active sheet.
    workbook : str, optional
        Workbook name. Defaults to the active workbook.
    color : str, optional
        "#RRGGBB" bar color for data_bar.
    scale_type : int, optional
        2 (two-color) or 3 (three-color) for color_scale. Default 3.
    style : str, optional
        Icon-set style name for icon_set. Default "3traffic_lights".
        Valid: 3arrows, 3arrows_gray, 3flags, 3traffic_lights, 3traffic_lights2,
        3signs, 3symbols, 4arrows, 4arrows_gray, 4red_to_black, 4crv,
        4traffic_lights, 5arrows, 5arrows_gray, 5crv, 5quarters.
    operator : str, optional
        Required for cell_rule. One of: between | not_between | equal | not_equal
        | greater | less | greater_equal | less_equal.
    formula1 : str, optional
        Required for cell_rule — first threshold value (e.g. "100").
    formula2 : str, optional
        Required for cell_rule with operator=between or not_between.
    fill_color : str, optional
        "#RRGGBB" fill color for cell_rule or top_bottom.
    font_color : str, optional
        "#RRGGBB" font color for cell_rule.
    kind : str, optional
        "top" or "bottom" for top_bottom. Default "top".
    rank : int, optional
        N value for top_bottom (top/bottom N items). Default 10.
    percent : bool, optional
        If True, rank is a percentage (top/bottom N%). Default False.

    Actions
    -------
    data_bar
        Add a data-bar rule. Optionally set bar color via color="#RRGGBB".
        Example: excel_conditional_format(action="data_bar", range="B2:B20",
                                          color="#0070C0")
    color_scale
        Add a 2- or 3-color gradient scale.
        Example: excel_conditional_format(action="color_scale", range="C2:C20",
                                          scale_type=3)
    icon_set
        Add an icon-set rule (traffic lights, arrows, flags, …).
        Example: excel_conditional_format(action="icon_set", range="D2:D20",
                                          style="3arrows")
    cell_rule
        Add a cell-value conditional format with operator, threshold, and colors.
        Example: excel_conditional_format(action="cell_rule", range="A2:A20",
                                          operator="greater", formula1="1000",
                                          fill_color="#FFFF00", font_color="#FF0000")
    top_bottom
        Highlight the top or bottom N cells (count or percent).
        Example: excel_conditional_format(action="top_bottom", range="B2:B20",
                                          kind="top", rank=5, percent=False,
                                          fill_color="#00B050")
    clear
        Delete ALL conditional formatting rules on the range.
        Example: excel_conditional_format(action="clear", range="A1:Z100")
    """
    return conditional_format_action(
        action,
        range=range,
        sheet=sheet,
        workbook=workbook,
        color=color,
        scale_type=scale_type,
        style=style,
        operator=operator,
        formula1=formula1,
        formula2=formula2,
        fill_color=fill_color,
        font_color=font_color,
        kind=kind,
        rank=rank,
        percent=percent,
    )


@mcp.tool()
def excel_validation(
    action: str,
    range: str,
    sheet: str | None = None,
    workbook: str | None = None,
    formula1: str | None = None,
    formula2: str | None = None,
    operator: str | None = None,
    in_cell_dropdown: bool = True,
    ignore_blank: bool = True,
    show_error: bool = True,
) -> dict:
    """Add or remove data validation on an Excel range.

    Always deletes any existing validation before adding new rules (required by
    Excel COM — adding over existing validation raises error 1004).

    Parameters
    ----------
    action : str
        One of: ``list``, ``whole_number``, ``decimal``, ``date``,
        ``text_length``, ``custom``, ``clear``.
    range : str
        Range address: ``"B2:B100"``, ``"Sheet1!C2:C50"``,
        ``"SalesTable[Region]"``.
    sheet : str, optional
        Sheet name. Uses active sheet when omitted.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    formula1 : str, optional
        First (or only) constraint value or formula.
        For ``list``: comma-separated values ``"Yes,No,Maybe"`` or a range
        reference ``"=$A$1:$A$5"``.
        For ``custom``: an Excel formula returning TRUE/FALSE, e.g. ``"=ISNUMBER(B2)"``.
        For constraint types: the bound value, e.g. ``"0"`` or ``"2026-01-01"``.
    formula2 : str, optional
        Upper bound — required when ``operator`` is ``between`` or ``not_between``.
    operator : str, optional
        Comparison operator for ``whole_number``, ``decimal``, ``date``,
        ``text_length``. One of:
        ``between``, ``not_between``, ``equal``, ``not_equal``,
        ``greater``, ``less``, ``greater_equal``, ``less_equal``.
        Must be OMITTED for ``list`` and ``custom``.
    in_cell_dropdown : bool, optional
        Show dropdown arrow in the cell (``list`` action only, default True).
    ignore_blank : bool, optional
        Allow blank cells to bypass validation (default True).
    show_error : bool, optional
        Show an error alert when invalid data is entered (default True).

    Actions
    -------
    list
        Restrict entry to a dropdown list of values.
        Example: ``excel_validation(action="list", range="B2:B100",
        formula1="Yes,No,Maybe")``
        Example: ``excel_validation(action="list", range="B2:B100",
        formula1="=$A$1:$A$5")``
    whole_number
        Restrict entry to whole integers.
        Example: ``excel_validation(action="whole_number", range="C2:C50",
        formula1="1", formula2="100", operator="between")``
        Example: ``excel_validation(action="whole_number", range="C2:C50",
        formula1="0", operator="greater")``
    decimal
        Restrict entry to decimal numbers.
        Example: ``excel_validation(action="decimal", range="D2:D50",
        formula1="0.0", formula2="1.0", operator="between")``
    date
        Restrict entry to dates.
        Example: ``excel_validation(action="date", range="E2:E50",
        formula1="2026-01-01", operator="greater_equal")``
    text_length
        Restrict text by character count.
        Example: ``excel_validation(action="text_length", range="F2:F200",
        formula1="50", operator="less_equal")``
    custom
        Validate using an arbitrary Excel formula (evaluates to TRUE = valid).
        Example: ``excel_validation(action="custom", range="G2:G50",
        formula1="=ISNUMBER(G2)")``
        Example: ``excel_validation(action="custom", range="H2:H50",
        formula1="=LEN(H2)<=100")``
    clear
        Remove all validation from the range.
        Example: ``excel_validation(action="clear", range="B2:B100")``
    """
    return validation_action(
        action,
        range=range,
        sheet=sheet,
        workbook=workbook,
        formula1=formula1,
        formula2=formula2,
        operator=operator,
        in_cell_dropdown=in_cell_dropdown,
        ignore_blank=ignore_blank,
        show_error=show_error,
    )


@mcp.tool()
def excel_slicer(
    action: str,
    workbook: str | None = None,
    source: str | None = None,
    field: str | None = None,
    sheet: str | None = None,
    caption: str | None = None,
    name: str | None = None,
    top: float = 50.0,
    left: float = 50.0,
    width: float = 144.0,
    height: float = 144.0,
    slicer: str | None = None,
    pivots: list[str] | None = None,
) -> dict:
    """Manage slicers and timelines on PivotTables and Tables.

    Parameters
    ----------
    action : str
        One of: ``add``, ``add_timeline``, ``list``, ``delete``, ``connect``
    workbook : str, optional
        Workbook name (active workbook if omitted).
    source : str
        (add / add_timeline) Name of the PivotTable or Table (ListObject) to filter.
    field : str
        (add / add_timeline) Column / field name to filter on.
    sheet : str, optional
        Sheet where the slicer control is placed (active sheet if omitted).
    caption : str, optional
        Visible label on the slicer header. Defaults to *field*.
    name : str, optional
        Unique slicer-cache name. Excel auto-generates if omitted.
    top : float
        Top position in points (default 50.0).
    left : float
        Left position in points (default 50.0).
    width : float
        Width in points (default 144.0 ≈ 2 inches).
    height : float
        Height in points (default 144.0 ≈ 2 inches).
    slicer : str
        (delete / connect) Name of the slicer cache to target.
    pivots : list[str]
        (connect) List of PivotTable names to connect the slicer to.

    Actions
    -------
    add
        Create a slicer for a Table or PivotTable field.

        Example::

            excel_slicer(action="add", source="SalesTable", field="Region",
                         caption="Region Filter", top=50, left=200, width=150, height=200)

    add_timeline
        Create a date-field timeline slicer (requires a date-typed field in a PivotTable).

        Example::

            excel_slicer(action="add_timeline", source="SalesPivot", field="OrderDate")

    list
        List all slicer caches in the workbook.

        Example::

            excel_slicer(action="list")

    delete
        Remove a slicer cache (and its visual controls) by cache name.

        Example::

            excel_slicer(action="delete", slicer="Slicer_Region")

    connect
        Connect an existing slicer cache to one or more additional PivotTables.

        Example::

            excel_slicer(action="connect", slicer="Slicer_Region",
                         pivots=["PivotTable1", "PivotTable2"])
    """
    return slicer_action(
        action,
        workbook=workbook,
        source=source,
        field=field,
        sheet=sheet,
        caption=caption,
        name=name,
        top=top,
        left=left,
        width=width,
        height=height,
        slicer=slicer,
        pivots=pivots,
    )


@mcp.tool()
def excel_page_setup(
    action: str,
    sheet: str | None = None,
    workbook: str | None = None,
    # set
    orientation: str | None = None,
    paper_size: str | None = None,
    fit_to_wide: int | None = None,
    fit_to_tall: int | None = None,
    scale: int | None = None,
    top: float | None = None,
    bottom: float | None = None,
    left: float | None = None,
    right: float | None = None,
    header: float | None = None,
    footer: float | None = None,
    center_horizontally: bool | None = None,
    center_vertically: bool | None = None,
    print_gridlines: bool | None = None,
    black_and_white: bool | None = None,
    # print_area
    address: str | None = None,
    # print_titles
    rows: str | None = None,
    cols: str | None = None,
    # header_footer
    left_header: str | None = None,
    center_header: str | None = None,
    right_header: str | None = None,
    left_footer: str | None = None,
    center_footer: str | None = None,
    right_footer: str | None = None,
    # export_pdf
    path: str | None = None,
    scope: str = "sheet",
    open_after: bool = False,
) -> dict:
    """Configure worksheet page setup and export to PDF.

    Controls orientation, paper size, margins, fit-to-page scaling, print area,
    print titles (repeating rows/columns), headers/footers, and PDF export.
    Margins are given in INCHES (converted to points internally).

    Parameters
    ----------
    action : str
        One of: ``set``, ``print_area``, ``print_titles``, ``header_footer``,
        ``export_pdf``, ``get``.
    sheet : str, optional
        Sheet name. Uses active sheet when omitted.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    orientation : str, optional
        ``"portrait"`` or ``"landscape"`` (``set`` action).
    paper_size : str, optional
        ``"a4"``, ``"letter"``, ``"a3"``, ``"legal"`` (``set`` action).
    fit_to_wide : int, optional
        FitToPagesWide — scale to N pages wide. Also disables Zoom (``set``).
    fit_to_tall : int, optional
        FitToPagesTall — scale to N pages tall. Also disables Zoom (``set``).
    scale : int, optional
        Print zoom 10–400 (``set``). Mutually exclusive with fit_to_*.
    top, bottom, left, right, header, footer : float, optional
        Margins in INCHES (``set``).
    center_horizontally, center_vertically : bool, optional
        Center the printout on the page (``set``).
    print_gridlines : bool, optional
        Print cell gridlines (``set``).
    black_and_white : bool, optional
        Print in black & white (``set``).
    address : str, optional
        Print-area range for ``print_area`` (e.g. ``"A1:F50"``; ``""`` clears it).
    rows : str, optional
        Repeating title rows for ``print_titles`` (e.g. ``"$1:$1"``).
    cols : str, optional
        Repeating title columns for ``print_titles`` (e.g. ``"$A:$A"``).
    left_header, center_header, right_header : str, optional
        Header slots (``header_footer``). Excel codes: ``&P`` page, ``&N`` pages,
        ``&D`` date, ``&T`` time, ``&F`` file name, ``&A`` sheet name.
    left_footer, center_footer, right_footer : str, optional
        Footer slots (``header_footer``).
    path : str, optional
        Full path ending in ``.pdf`` for ``export_pdf``.
    scope : str, optional
        ``"sheet"`` (default) or ``"workbook"`` for ``export_pdf``.
    open_after : bool, optional
        Open the PDF after export (``export_pdf``, default False).

    Actions
    -------
    set
        Set any subset of page-setup properties; only supplied values change.
        Example::

            excel_page_setup(action="set", sheet="Report",
                             orientation="landscape", paper_size="a4",
                             fit_to_wide=1, top=0.5)

    print_area
        Set or clear the print area.
        Example::

            excel_page_setup(action="print_area", address="A1:F50")
            excel_page_setup(action="print_area", address="")   # clear

    print_titles
        Set repeating header rows/columns for printing.
        Example::

            excel_page_setup(action="print_titles", rows="$1:$1", cols="$A:$A")

    header_footer
        Set header/footer slots (supports Excel codes).
        Example::

            excel_page_setup(action="header_footer",
                             center_header="ThepExcel Report",
                             right_footer="Page &P of &N")

    export_pdf
        Export the sheet or workbook to PDF.
        Example::

            excel_page_setup(action="export_pdf", path="C:/temp/report.pdf",
                             scope="sheet")

    get
        Read back the current page-setup properties as a dict.
        Example::

            excel_page_setup(action="get", sheet="Report")
    """
    return page_setup_action(
        action,
        sheet=sheet,
        workbook=workbook,
        orientation=orientation,
        paper_size=paper_size,
        fit_to_wide=fit_to_wide,
        fit_to_tall=fit_to_tall,
        scale=scale,
        top=top,
        bottom=bottom,
        left=left,
        right=right,
        header=header,
        footer=footer,
        center_horizontally=center_horizontally,
        center_vertically=center_vertically,
        print_gridlines=print_gridlines,
        black_and_white=black_and_white,
        address=address,
        rows=rows,
        cols=cols,
        left_header=left_header,
        center_header=center_header,
        right_header=right_header,
        left_footer=left_footer,
        center_footer=center_footer,
        right_footer=right_footer,
        path=path,
        scope=scope,
        open_after=open_after,
    )


@mcp.tool()
def excel_comment(
    action: str,
    cell: str | None = None,
    sheet: str | None = None,
    workbook: str | None = None,
    text: str | None = None,
    kind: str = "note",
) -> dict:
    """Add, edit, reply to, delete, or read cell comments (notes and threaded).

    Excel has two comment systems: legacy yellow sticky **notes** (``kind="note"``)
    and modern **threaded** comments (``kind="threaded"``). Notes support edit;
    threaded comments support replies. Notes are added with delete-before-add to
    avoid COM error 1004.

    Parameters
    ----------
    action : str
        One of: ``add``, ``edit``, ``reply``, ``delete``, ``list``, ``get``.
    cell : str, optional
        Single-cell address (e.g. ``"A1"``). Required for all actions except ``list``.
    sheet : str, optional
        Sheet name. Uses active sheet when omitted.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    text : str, optional
        Comment text. Required for ``add``, ``edit``, ``reply``.
    kind : str, optional
        ``"note"`` (default) for legacy notes, ``"threaded"`` for threaded comments,
        or ``"all"`` for the ``list``/``get`` actions to return both kinds.

    Actions
    -------
    add
        Add a note or threaded comment.
        Example::

            excel_comment(action="add", cell="A1", text="Check this figure")
            excel_comment(action="add", cell="B2", text="Needs review",
                          kind="threaded")

    edit
        Replace the text of an existing note (``kind="note"`` only; for threaded,
        use ``add`` to delete+re-add).
        Example::

            excel_comment(action="edit", cell="A1", text="Updated note")

    reply
        Add a reply to a threaded comment (``kind="threaded"`` only).
        Example::

            excel_comment(action="reply", cell="B2", text="Confirmed",
                          kind="threaded")

    delete
        Delete the note or threaded comment on a cell.
        Example::

            excel_comment(action="delete", cell="A1")
            excel_comment(action="delete", cell="B2", kind="threaded")

    list
        List all notes and/or threaded comments on the sheet.
        Example::

            excel_comment(action="list", kind="all")

    get
        Read the note and/or threaded comment on a single cell.
        Example::

            excel_comment(action="get", cell="A1", kind="all")
    """
    return comment_action(
        action,
        cell=cell,
        sheet=sheet,
        workbook=workbook,
        text=text,
        kind=kind,
    )


@mcp.tool()
def excel_hyperlink(
    action: str,
    sheet: str | None = None,
    workbook: str | None = None,
    cell: str | None = None,
    link_type: str | None = None,
    target: str | None = None,
    sub_address: str | None = None,
    text_to_display: str | None = None,
    screen_tip: str | None = None,
    range: str | None = None,
) -> dict:
    """Add, list, or delete worksheet hyperlinks.

    Supports URL, internal (sheet/cell), email, and file hyperlinks. Deleting
    hyperlinks keeps the cell text and value intact.

    Parameters
    ----------
    action : str
        One of: ``add``, ``list``, ``delete``.
    sheet : str, optional
        Sheet name. Uses active sheet when omitted.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    cell : str, optional
        Anchor cell for ``add`` (e.g. ``"A1"``).
    link_type : str, optional
        Required for ``add``. One of: ``url``, ``internal``, ``email``, ``file``.
    target : str, optional
        The URL, file path, or email address (``add``). For ``internal`` links,
        use ``sub_address`` instead.
    sub_address : str, optional
        In-document anchor for ``internal`` links (e.g. ``"Sheet1!C3"``).
    text_to_display : str, optional
        Visible cell text for ``add``.
    screen_tip : str, optional
        Hover tooltip for ``add``.
    range : str, optional
        Range to clear hyperlinks from for ``delete`` (e.g. ``"A1"`` or ``"A1:A10"``).

    Actions
    -------
    add
        Add a hyperlink to a cell.
        Example::

            excel_hyperlink(action="add", cell="A1", link_type="url",
                            target="https://www.thepexcel.com",
                            text_to_display="ThepExcel")
            excel_hyperlink(action="add", cell="A2", link_type="internal",
                            sub_address="Sheet1!C3", text_to_display="Jump")

    list
        List all hyperlinks on the sheet.
        Example::

            excel_hyperlink(action="list")

    delete
        Remove all hyperlinks from a range (keeps cell text/value).
        Example::

            excel_hyperlink(action="delete", range="A1:A10")
    """
    return hyperlink_action(
        action,
        sheet=sheet,
        workbook=workbook,
        cell=cell,
        link_type=link_type,
        target=target,
        sub_address=sub_address,
        text_to_display=text_to_display,
        screen_tip=screen_tip,
        range=range,
    )


@mcp.tool()
def excel_outline(
    action: str,
    sheet: str | None = None,
    workbook: str | None = None,
    rows: str | None = None,
    columns: str | None = None,
    row_levels: int | None = None,
    column_levels: int | None = None,
) -> dict:
    """Group/ungroup rows and columns and control outline levels.

    Wraps Excel's row/column grouping (outline). Grouping raises the OutlineLevel
    of the affected rows/columns to >= 2; ungrouping lowers it.

    Parameters
    ----------
    action : str
        One of: ``group_rows``, ``group_columns``, ``ungroup_rows``,
        ``ungroup_columns``, ``show_levels``, ``clear``.
    sheet : str, optional
        Sheet name. Uses active sheet when omitted.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    rows : str, optional
        Row spec for ``group_rows`` / ``ungroup_rows`` (e.g. ``"2:5"``).
    columns : str, optional
        Column spec for ``group_columns`` / ``ungroup_columns`` (e.g. ``"B:D"``).
    row_levels : int, optional
        Outline level (1–8) to show/collapse rows to (``show_levels``).
    column_levels : int, optional
        Outline level (1–8) to show/collapse columns to (``show_levels``).

    Actions
    -------
    group_rows
        Group a row range.
        Example::

            excel_outline(action="group_rows", rows="2:5")

    group_columns
        Group a column range.
        Example::

            excel_outline(action="group_columns", columns="B:D")

    ungroup_rows
        Ungroup a row range (lowers outline level by 1).
        Example::

            excel_outline(action="ungroup_rows", rows="2:5")

    ungroup_columns
        Ungroup a column range.
        Example::

            excel_outline(action="ungroup_columns", columns="B:D")

    show_levels
        Show/collapse the outline to a specific level. At least one of
        ``row_levels`` / ``column_levels`` is required.
        Example::

            excel_outline(action="show_levels", row_levels=1)

    clear
        Remove ALL row/column outline groupings from the sheet.
        Example::

            excel_outline(action="clear")
    """
    return outline_action(
        action,
        sheet=sheet,
        workbook=workbook,
        rows=rows,
        columns=columns,
        row_levels=row_levels,
        column_levels=column_levels,
    )


@mcp.tool()
def excel_protection(
    action: str,
    sheet: str | None = None,
    workbook: str | None = None,
    password: str | None = None,
    drawing_objects: bool = True,
    contents: bool = True,
    scenarios: bool = True,
    allow: dict | None = None,
    allow_formatting_cells: bool = False,
    allow_formatting_columns: bool = False,
    allow_formatting_rows: bool = False,
    allow_inserting_columns: bool = False,
    allow_inserting_rows: bool = False,
    allow_inserting_hyperlinks: bool = False,
    allow_deleting_columns: bool = False,
    allow_deleting_rows: bool = False,
    allow_sorting: bool = False,
    allow_filtering: bool = False,
    allow_using_pivot_tables: bool = False,
    structure: bool = True,
    windows: bool = False,
    range: str | None = None,
    locked: bool | None = None,
    hidden: bool | None = None,
) -> dict:
    """Protect/unprotect worksheets and workbooks and set cell lock state.

    Worksheet protection locks cell editing (respecting per-cell ``Locked`` flags);
    workbook protection locks the sheet structure and/or window layout. Cell
    ``Locked`` / ``FormulaHidden`` flags only take effect while the sheet is
    protected — configure them with ``set_locked`` BEFORE ``protect_sheet``.

    IMPORTANT: keep track of any password you set — Excel cannot recover a lost
    sheet/workbook password.

    Parameters
    ----------
    action : str
        One of: ``protect_sheet``, ``unprotect_sheet``, ``protect_workbook``,
        ``unprotect_workbook``, ``set_locked``, ``status``.
    sheet : str, optional
        Sheet name. Uses active sheet when omitted.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    password : str, optional
        Password for protect/unprotect. Must match on unprotect.
    drawing_objects, contents, scenarios : bool, optional
        ``protect_sheet`` structural flags (all default True).
    allow : dict, optional
        Optional dict of allow-flag overrides for ``protect_sheet``.
    allow_formatting_cells, allow_formatting_columns, allow_formatting_rows,
    allow_inserting_columns, allow_inserting_rows, allow_inserting_hyperlinks,
    allow_deleting_columns, allow_deleting_rows, allow_sorting, allow_filtering,
    allow_using_pivot_tables : bool, optional
        Per-action permissions for ``protect_sheet`` (all default False).
    structure : bool, optional
        Lock sheet structure for ``protect_workbook`` (default True).
    windows : bool, optional
        Lock window layout for ``protect_workbook`` (default False).
    range : str, optional
        Cell range for ``set_locked`` (e.g. ``"A1:B10"``).
    locked : bool, optional
        Set the ``Locked`` flag on the range (``set_locked``).
    hidden : bool, optional
        Set the ``FormulaHidden`` flag on the range (``set_locked``).

    Actions
    -------
    protect_sheet
        Protect a worksheet, optionally with a password and allow-flags.
        Example::

            excel_protection(action="protect_sheet", sheet="Data",
                             password="secret", allow_sorting=True)

    unprotect_sheet
        Remove sheet protection (supply the same password).
        Example::

            excel_protection(action="unprotect_sheet", sheet="Data",
                             password="secret")

    protect_workbook
        Protect workbook structure and/or windows.
        Example::

            excel_protection(action="protect_workbook", password="secret",
                             structure=True)

    unprotect_workbook
        Remove workbook protection.
        Example::

            excel_protection(action="unprotect_workbook", password="secret")

    set_locked
        Set ``Locked`` and/or ``FormulaHidden`` on a range (takes effect under
        sheet protection).
        Example::

            excel_protection(action="set_locked", range="A1:A10", locked=False)

    status
        Read back current sheet + workbook protection flags.
        Example::

            excel_protection(action="status", sheet="Data")
    """
    return protection_action(
        action,
        sheet=sheet,
        workbook=workbook,
        password=password,
        drawing_objects=drawing_objects,
        contents=contents,
        scenarios=scenarios,
        allow=allow,
        allow_formatting_cells=allow_formatting_cells,
        allow_formatting_columns=allow_formatting_columns,
        allow_formatting_rows=allow_formatting_rows,
        allow_inserting_columns=allow_inserting_columns,
        allow_inserting_rows=allow_inserting_rows,
        allow_inserting_hyperlinks=allow_inserting_hyperlinks,
        allow_deleting_columns=allow_deleting_columns,
        allow_deleting_rows=allow_deleting_rows,
        allow_sorting=allow_sorting,
        allow_filtering=allow_filtering,
        allow_using_pivot_tables=allow_using_pivot_tables,
        structure=structure,
        windows=windows,
        range=range,
        locked=locked,
        hidden=hidden,
    )


@mcp.tool()
def excel_sparkline(
    action: str,
    location: str | None = None,
    sheet: str | None = None,
    workbook: str | None = None,
    data_range: str | None = None,
    spark_type: str = "line",
    marker: bool | None = None,
    color: str | None = None,
) -> dict:
    """Add, clear, or list in-cell sparkline mini-charts.

    Sparklines are tiny charts rendered inside cells. ``location`` is the
    destination range (where the sparklines render) and ``data_range`` is the
    source data. For a column of per-row sparklines, location ``"F2:F4"`` +
    data_range ``"B2:E4"`` maps each destination cell to the matching data row.

    Parameters
    ----------
    action : str
        One of: ``add``, ``clear``, ``list``.
    location : str
        REQUIRED for all actions — the destination range (e.g. ``"F2:F4"``).
    sheet : str, optional
        Sheet name. Uses active sheet when omitted.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    data_range : str, optional
        Source data range for ``add`` (e.g. ``"B2:E4"`` or ``"Sheet1!B2:E4"``).
    spark_type : str, optional
        ``"line"`` (default), ``"column"``, or ``"win_loss"`` for ``add``.
    marker : bool, optional
        Show data-point markers (line type only) for ``add``.
    color : str, optional
        ``"#RRGGBB"`` series color for ``add``.

    Actions
    -------
    add
        Add sparklines at ``location`` using data from ``data_range``.
        Example::

            excel_sparkline(action="add", location="F2:F4",
                            data_range="B2:E4", spark_type="line")
            excel_sparkline(action="add", location="G2:G4",
                            data_range="B2:E4", spark_type="column")

    clear
        Remove all sparklines from ``location``.
        Example::

            excel_sparkline(action="clear", location="F2:F4")

    list
        Report sparkline groups on the sheet (scoped to ``location``).
        Example::

            excel_sparkline(action="list", location="F2")
    """
    return sparkline_action(
        action,
        location=location,
        sheet=sheet,
        workbook=workbook,
        data_range=data_range,
        spark_type=spark_type,
        marker=marker,
        color=color,
    )


@mcp.tool()
def excel_shape(
    action: str,
    sheet: str | None = None,
    workbook: str | None = None,
    cell: str | None = None,
    left: float | None = None,
    top: float | None = None,
    width: float | None = None,
    height: float | None = None,
    filename: str | None = None,
    text: str | None = None,
    shape_type: str | None = None,
    name: str | None = None,
) -> dict:
    """Add and manage drawing objects (images, text boxes, AutoShapes) on a sheet.

    Positioning is in POINTS (1 inch = 72 pt). Supply either a *cell* anchor
    (the shape's upper-left snaps to that cell's corner) OR explicit *left* and
    *top*. Adding a shape returns its generated ``name`` under ``applied`` — use
    that name to target ``move`` and ``delete``.

    Parameters
    ----------
    action : str
        One of: ``add_image``, ``add_textbox``, ``add_shape``, ``list``,
        ``delete``, ``move``.
    sheet : str, optional
        Sheet name. Uses active sheet when omitted.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    cell : str, optional
        Anchor cell (e.g. ``"B3"``). Takes priority over *left*/*top*.
    left : float, optional
        Explicit X position in points (required with *top* when *cell* omitted).
    top : float, optional
        Explicit Y position in points.
    width : float, optional
        Width in points. Required for ``add_textbox``/``add_shape``. For
        ``add_image`` pass ``-1`` (or omit) to keep native image width.
    height : float, optional
        Height in points (same rules as *width*).
    filename : str, optional
        Absolute path to the image file (``add_image`` only).
    text : str, optional
        Initial text for ``add_textbox``, or label text for ``add_shape``.
    shape_type : str, optional
        AutoShape type for ``add_shape``. Named: ``rectangle``,
        ``rounded_rectangle``, ``oval``, ``right_arrow``, ``diamond``,
        ``triangle``, ``hexagon``, ``star``, ``cloud``, ``heart`` — or a raw
        MsoAutoShapeType integer as a string (e.g. ``"1"``).
    name : str, optional
        Target shape name for ``move`` and ``delete``.

    Actions
    -------
    add_image
        Embed an image into the sheet. Requires an absolute *filename*.
        Example::

            excel_shape(action="add_image", filename="D:/logo.png",
                        cell="F2", width=-1, height=-1)

    add_textbox
        Add a horizontal text box. *width*/*height* required.
        Example::

            excel_shape(action="add_textbox", cell="B2", text="Hello box",
                        width=120, height=40)

    add_shape
        Add an AutoShape. *shape_type*, *width*, *height* required.
        Example::

            excel_shape(action="add_shape", shape_type="oval", cell="D2",
                        width=80, height=60)

    list
        Enumerate shapes (name, shape_type, left, top, width, height).
        Example: ``excel_shape(action="list")``

    move
        Reposition (and optionally resize) the named shape.
        Example::

            excel_shape(action="move", name="TextBox 1", left=10, top=10)

    delete
        Delete the named shape.
        Example: ``excel_shape(action="delete", name="Oval 2")``
    """
    return shape_action(
        action,
        sheet=sheet,
        workbook=workbook,
        cell=cell,
        left=left,
        top=top,
        width=width,
        height=height,
        filename=filename,
        text=text,
        shape_type=shape_type,
        name=name,
    )


@mcp.tool()
def excel_find_replace(
    action: str,
    find_text: str,
    replace_text: str | None = None,
    scope: str = "sheet",
    range: str | None = None,
    sheet: str | None = None,
    workbook: str | None = None,
    match_case: bool = False,
    match_whole_cell: bool = False,
    look_in: str = "formulas",
) -> dict:
    """Find, count, or replace text across a range, sheet, or whole workbook.

    ``look_in`` controls what is searched: ``"formulas"`` (default — matches the
    underlying formula/literal text, the same as Excel's Find dialog default) or
    ``"values"`` (matches the displayed/computed value). For replacing literal
    cell content, ``"values"`` is usually the intuitive choice.

    Parameters
    ----------
    action : str
        One of: ``find``, ``count``, ``replace``.
    find_text : str
        Text to search for (REQUIRED).
    replace_text : str, optional
        Replacement text. REQUIRED for ``replace``.
    scope : str, optional
        ``"sheet"`` (default — entire active/named sheet), ``"range"`` (limited
        to *range*, which is then required), or ``"workbook"`` (all sheets).
    range : str, optional
        Range address when ``scope="range"`` (e.g. ``"A1:A100"``).
    sheet : str, optional
        Sheet name. Uses active sheet when omitted.
    workbook : str, optional
        Workbook name. Uses active workbook when omitted.
    match_case : bool, optional
        Case-sensitive search (default False).
    match_whole_cell : bool, optional
        Match the ENTIRE cell content only (default False = substring match).
    look_in : str, optional
        ``"formulas"`` (default) or ``"values"``.

    Actions
    -------
    find
        Return matching cell addresses + their current values (capped at 1000,
        ``truncated`` flag set if more exist).
        Example::

            excel_find_replace(action="find", find_text="apple",
                               scope="sheet", look_in="values")

    count
        Return only the number of matching cells.
        Example::

            excel_find_replace(action="count", find_text="apple",
                               match_whole_cell=True, look_in="values")

    replace
        Replace every match. Reports ``cells_matched_before`` and
        ``remaining_after`` (0 = fully replaced).
        Example::

            excel_find_replace(action="replace", find_text="apple",
                               replace_text="orange", scope="sheet",
                               look_in="values")
    """
    return find_replace_action(
        action,
        find_text,
        replace_text=replace_text,
        scope=scope,
        range=range,
        sheet=sheet,
        workbook=workbook,
        match_case=match_case,
        match_whole_cell=match_whole_cell,
        look_in=look_in,
    )


@mcp.tool()
def excel_diff(
    action: str,
    left_range: str | None = None,
    left_sheet: str | None = None,
    left_workbook: str | None = None,
    right_range: str | None = None,
    right_sheet: str | None = None,
    right_workbook: str | None = None,
    compare: str = "value",
    max_diffs: int = 500,
) -> dict:
    """Diff two ranges or two whole sheets, reporting cell-level differences.

    PURE READ — no mutation. Each difference is reported with the LEFT-side A1
    cell address, ``left_value``, and ``right_value`` (plus formulas when
    ``compare`` includes formulas). ``total_diffs`` is the true count even when
    the returned ``diffs`` list is truncated to *max_diffs*.

    Parameters
    ----------
    action : str
        One of: ``ranges``, ``sheets``.
    left_range : str, optional
        Left range address (REQUIRED for ``ranges``), e.g. ``"A1:B3"`` or
        ``"Sheet1!A1:B3"``.
    left_sheet : str, optional
        Sheet for the left side. REQUIRED for ``sheets``.
    left_workbook : str, optional
        Workbook for the left side. Uses active workbook when omitted.
    right_range : str, optional
        Right range address (REQUIRED for ``ranges``).
    right_sheet : str, optional
        Sheet for the right side. REQUIRED for ``sheets``.
    right_workbook : str, optional
        Workbook for the right side. Uses active workbook when omitted.
    compare : str, optional
        ``"value"`` (default), ``"formula"``, or ``"both"`` (differ if EITHER
        value or formula differs).
    max_diffs : int, optional
        Cap on returned diff entries (default 500).

    Actions
    -------
    ranges
        Compare *left_range* vs *right_range* cell-by-cell. Reports
        ``dimensions_match`` and the per-cell ``diffs``.
        Example::

            excel_diff(action="ranges", left_range="A1:B3",
                       right_range="D1:E3", compare="value")

    sheets
        Compare two whole sheets over their combined used-range bounding box.
        Example::

            excel_diff(action="sheets", left_sheet="Sheet1",
                       right_sheet="Sheet2", compare="value")
    """
    return diff_action(
        action,
        left_range=left_range,
        left_sheet=left_sheet,
        left_workbook=left_workbook,
        right_range=right_range,
        right_sheet=right_sheet,
        right_workbook=right_workbook,
        compare=compare,
        max_diffs=max_diffs,
    )


@mcp.tool()
def excel_snapshot(
    action: str,
    workbook: str | None = None,
    snapshot_id: str | None = None,
) -> dict:
    """Create and restore non-destructive, on-disk safety copies of a workbook.

    SAFETY (the defining property): this tool NEVER closes, overwrites, or
    reverts the workbook you are editing. ``snapshot`` writes a COPY to disk via
    ``SaveCopyAs`` (which leaves the live workbook's name, path, and dirty flag
    untouched); ``restore`` only OPENS that copy as a separate new workbook
    alongside everything else. There is no in-place / close-and-reopen path.

    Parameters
    ----------
    action : str
        One of: ``snapshot``, ``list``, ``restore``, ``delete``.
    workbook : str, optional
        Workbook name to snapshot. Uses the active workbook when omitted.
        (Only used by ``snapshot``.)
    snapshot_id : str, optional
        The snapshot id returned by ``snapshot``. REQUIRED for ``restore`` and
        ``delete``.

    Actions
    -------
    snapshot
        Save a point-in-time copy of *workbook* to a temp directory via
        ``SaveCopyAs``. The source extension is preserved, so an .xlsm keeps its
        macros. The live workbook is NOT modified, saved, or rebound.
        Returns ``{id, path, workbook, size_bytes, created}``.
        Example::

            excel_snapshot(action="snapshot")               # active workbook
            excel_snapshot(action="snapshot", workbook="Sales.xlsx")

    list
        List every snapshot taken this session with ``{id, workbook, path,
        created, exists, size_bytes}`` (``exists`` reflects whether the file is
        still on disk). Read-only.

    restore
        Open the snapshot ``snapshot_id`` as a NEW, separate workbook — your
        original workbook is left exactly as it is. To revert, copy what you
        need out of the opened restored copy by hand.
        Example::

            excel_snapshot(action="restore", snapshot_id="snap_1_Sales")

    delete
        Delete the snapshot file from disk and remove it from the registry.
        Example::

            excel_snapshot(action="delete", snapshot_id="snap_1_Sales")
    """
    return snapshot_action(
        action,
        workbook=workbook,
        snapshot_id=snapshot_id,
    )


def main() -> None:
    """Run the MCP server over stdio."""
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
