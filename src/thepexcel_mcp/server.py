"""ThepExcelMCP — FastMCP server entry point.

4 coarse action-dispatch tools; each tool routes to a domain module.
Tool docstrings are the LLM-facing API — kept precise and example-rich.
"""

from __future__ import annotations

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError  # noqa: F401 — re-exported for domain modules

from .domains.powerquery import powerquery_action
from .domains.ranges import range_action
from .domains.sheets import sheet_action
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


def main() -> None:
    """Run the MCP server over stdio."""
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
