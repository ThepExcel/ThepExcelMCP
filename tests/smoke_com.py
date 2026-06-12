"""Manual COM smoke test — requires a live Excel instance.

NOT collected by pytest (see pyproject.toml collect_ignore).
Run manually: uv run python tests/smoke_com.py

Phase 0 section: READ-ONLY checks on existing workbooks.
Phase 1 section: Creates a TEMPORARY workbook, exercises table+pivot ops,
then closes WITHOUT saving. Pre-existing workbooks are NEVER touched.
"""

from __future__ import annotations

import sys

try:
    from thepexcel_mcp.domains.workbook import workbook_action
    from thepexcel_mcp.domains.powerquery import powerquery_action
    from thepexcel_mcp.domains.sheets import sheet_action
    from thepexcel_mcp.domains.tables import table_action
    from thepexcel_mcp.domains.pivots import pivot_action
    from thepexcel_mcp.session import ExcelSession
except ImportError as e:
    print(f"Import failed: {e}")
    sys.exit(1)


def main():
    print("=== ThepExcelMCP COM Smoke Test ===\n")

    # ── Phase 0: read-only checks on existing workbooks ──────────────────────
    print("[1] Listing open workbooks...")
    try:
        result = workbook_action("list")
        print(f"    Workbooks: {result}")
    except Exception as e:
        print(f"    ERROR: {e}")
        print("    (Is Excel running?)")
        return

    workbooks = result.get("workbooks", [])
    if not workbooks:
        print("    No workbooks open — skipping further checks.")
        return

    active = next((w["name"] for w in workbooks if w.get("active")), workbooks[0]["name"])
    print(f"\n[2] Workbook info for active: '{active}'")
    try:
        info = workbook_action("info", workbook=active)
        print(f"    Sheets: {info['sheets']}")
        print(f"    Queries: {info['query_count']}, Tables: {info['table_count']}")
    except Exception as e:
        print(f"    ERROR: {e}")

    print(f"\n[3] Listing sheets in '{active}'...")
    try:
        sheets = sheet_action("list", workbook=active)
        print(f"    {sheets}")
    except Exception as e:
        print(f"    ERROR: {e}")

    print(f"\n[4] Listing Power Query queries in '{active}'...")
    try:
        queries = powerquery_action("list", workbook=active)
        print(f"    Found {queries['count']} queries")
        for q in queries["queries"]:
            print(f"    - {q['name']} ({q['line_count']} lines)")
    except Exception as e:
        print(f"    ERROR: {e}")

    print(f"\n[5] Listing tables in '{active}'...")
    try:
        tables = table_action("list", workbook=active)
        print(f"    Found {tables['count']} tables")
        for t in tables["tables"]:
            print(f"    - {t['name']} on {t['sheet']} ({t['row_count']} rows, cols: {t['columns']})")
    except Exception as e:
        print(f"    ERROR: {e}")

    print(f"\n[6] Listing PivotTables in '{active}'...")
    try:
        pivots = pivot_action("list", workbook=active)
        print(f"    Found {pivots['count']} pivot tables")
        for pt in pivots["pivot_tables"]:
            print(f"    - {pt['name']} on {pt['sheet']} | fields: {pt.get('fields', {})}")
    except Exception as e:
        print(f"    ERROR: {e}")

    # ── Phase 1: temporary workbook table+pivot exercise ─────────────────────
    print("\n[7] Phase 1: creating temporary workbook for table+pivot tests...")
    _session = ExcelSession()
    app = None
    tmp_wb = None
    tmp_wb_name = None
    try:
        app = _session.get_app()
        tmp_wb = app.Workbooks.Add()
        tmp_wb_name = tmp_wb.Name
        print(f"    Created temp workbook: {tmp_wb_name}")

        ws = tmp_wb.ActiveSheet
        ws.Name = "Data"

        # Write test data: headers + 5 rows
        headers = [["Region", "Product", "Quarter", "Amount"]]
        data = [
            ["North", "Widget", "Q1", 1000],
            ["South", "Gadget", "Q1", 1500],
            ["North", "Widget", "Q2", 1200],
            ["South", "Gadget", "Q2", 1800],
            ["North", "Gadget", "Q1", 900],
        ]
        ws.Range("A1:D1").Value = headers
        ws.Range("A2:D6").Value = data
        print("    Test data written.")

        # --- Table tests ---
        print("\n    [7a] Create table from A1:D6...")
        r = table_action("create", name="TestTable", range="A1:D6",
                         sheet="Data", workbook=tmp_wb_name)
        print(f"         {r}")

        print("    [7b] List tables...")
        r = table_action("list", workbook=tmp_wb_name)
        print(f"         {r}")

        print("    [7c] Read table (limit=3)...")
        r = table_action("read", name="TestTable", workbook=tmp_wb_name, limit=3)
        print(f"         columns={r['columns']}, rows={len(r['values'])}, "
              f"has_more={r['has_more']}")

        print("    [7d] Append 2 rows...")
        r = table_action("append_rows", name="TestTable", workbook=tmp_wb_name,
                         values=[["East", "Widget", "Q2", 2000],
                                  ["West", "Gadget", "Q2", 2200]])
        print(f"         {r}")

        print("    [7e] Add calculated column (Tax = Amount*0.1)...")
        r = table_action("add_column", name="TestTable", workbook=tmp_wb_name,
                         column_name="Tax", formula="=[@Amount]*0.1")
        print(f"         {r}")

        print("    [7f] Sort by Amount descending...")
        r = table_action("sort", name="TestTable", workbook=tmp_wb_name,
                         sort_column="Amount", ascending=False)
        print(f"         {r}")

        print("    [7g] Filter Region=North...")
        r = table_action("filter", name="TestTable", workbook=tmp_wb_name,
                         filter_column="Region", filter_op="equals",
                         filter_value="North")
        print(f"         {r}")

        print("    [7h] Clear filters...")
        r = table_action("filter", name="TestTable", workbook=tmp_wb_name,
                         filter_column="Region", filter_op="clear_filters")
        print(f"         {r}")

        print("    [7i] Toggle totals (show, sum on Amount)...")
        r = table_action("toggle_totals", name="TestTable", workbook=tmp_wb_name,
                         show_totals=True, column_name="Amount", total_func="sum")
        print(f"         {r}")

        print("    [7j] Set style TableStyleLight2...")
        r = table_action("set_style", name="TestTable", workbook=tmp_wb_name,
                         style="TableStyleLight2")
        print(f"         {r}")

        print("    [7k] Rename table...")
        r = table_action("rename", name="TestTable", workbook=tmp_wb_name,
                         new_name="SalesData")
        print(f"         {r}")

        # --- Pivot tests ---
        print("\n    [7l] Create PivotTable from table SalesData...")
        r = pivot_action("create", name="SalesPivot", source="SalesData",
                         workbook=tmp_wb_name)
        print(f"         {r}")

        print("    [7m] Add Region to rows...")
        r = pivot_action("add_field", name="SalesPivot", workbook=tmp_wb_name,
                         field="Region", area="rows")
        print(f"         {r}")

        print("    [7n] Add Amount to values (sum)...")
        r = pivot_action("add_field", name="SalesPivot", workbook=tmp_wb_name,
                         field="Amount", area="values", aggregation="sum",
                         number_format="#,##0")
        print(f"         {r}")

        print("    [7o] List pivot tables...")
        r = pivot_action("list", workbook=tmp_wb_name)
        print(f"         {r}")

        print("    [7p] Read pivot (limit=10)...")
        r = pivot_action("read", name="SalesPivot", workbook=tmp_wb_name)
        print(f"         total_rows={r['total_rows']}, rows={r['values'][:3]}")

        print("    [7q] Set layout tabular...")
        r = pivot_action("set_layout", name="SalesPivot", workbook=tmp_wb_name,
                         layout="tabular", grand_totals=True)
        print(f"         {r}")

        print("    [7r] Refresh pivot...")
        r = pivot_action("refresh", name="SalesPivot", workbook=tmp_wb_name)
        print(f"         {r}")

        print("    [7s] Remove field Region...")
        r = pivot_action("remove_field", name="SalesPivot", workbook=tmp_wb_name,
                         field="Region")
        print(f"         {r}")

        print("    [7t] Delete table (keep_data=True)...")
        r = table_action("delete", name="SalesData", workbook=tmp_wb_name,
                         keep_data=True)
        print(f"         {r}")

        print("\n    All Phase 1 checks PASSED.")

    except Exception as e:
        print(f"\n    Phase 1 ERROR: {e}")
    finally:
        if tmp_wb is not None:
            try:
                tmp_wb.Close(SaveChanges=False)
                print(f"\n    Temp workbook '{tmp_wb_name}' closed (not saved).")
            except Exception as e:
                print(f"    WARNING: Could not close temp workbook: {e}")

    print("\n=== Smoke test complete ===")


if __name__ == "__main__":
    main()
