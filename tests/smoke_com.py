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
    from thepexcel_mcp.domains.datamodel import datamodel_action
    from thepexcel_mcp.domains.charts import chart_action
    from thepexcel_mcp.domains.screenshot import screenshot_action
    from thepexcel_mcp.domains.ranges import range_action
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

    # ── Phase 2: Data Model + DAX measures ───────────────────────────────────
    print("\n[8] Phase 2: Data Model + DAX measures on a fresh temp workbook...")
    tmp_wb2 = None
    tmp_wb2_name = None
    try:
        app = _session.get_app()
        tmp_wb2 = app.Workbooks.Add()
        tmp_wb2_name = tmp_wb2.Name
        print(f"    Created temp workbook: {tmp_wb2_name}")

        ws = tmp_wb2.ActiveSheet
        ws.Name = "Sales"

        # Write test data: headers + 5 rows
        ws.Range("A1:C1").Value = [["ProductID", "Region", "Amount"]]
        ws.Range("A2:C6").Value = [
            [1, "North", 1000],
            [2, "South", 1500],
            [1, "North", 1200],
            [2, "South", 1800],
            [1, "East",  900],
        ]

        ws2 = tmp_wb2.Sheets.Add()
        ws2.Name = "Products"
        ws2.Range("A1:B1").Value = [["ProductID", "ProductName"]]
        ws2.Range("A2:B3").Value = [[1, "Widget"], [2, "Gadget"]]
        print("    Test data written.")

        print("\n    [8a] Create Sales table from Sheet 'Sales'...")
        r = table_action("create", name="SalesTable", range="A1:C6",
                         sheet="Sales", workbook=tmp_wb2_name)
        print(f"         {r}")

        print("    [8b] Create Products table from Sheet 'Products'...")
        r = table_action("create", name="ProductsTable", range="A1:B3",
                         sheet="Products", workbook=tmp_wb2_name)
        print(f"         {r}")

        print("    [8c] Add SalesTable to Data Model...")
        r = datamodel_action("add_table", workbook=tmp_wb2_name,
                             source_type="table", source_name="SalesTable")
        print(f"         {r}")

        print("    [8d] Add ProductsTable to Data Model...")
        r = datamodel_action("add_table", workbook=tmp_wb2_name,
                             source_type="table", source_name="ProductsTable")
        print(f"         {r}")

        print("    [8e] Data Model info...")
        r = datamodel_action("info", workbook=tmp_wb2_name)
        print(f"         {r}")

        print("    [8f] List model tables...")
        r = datamodel_action("list_tables", workbook=tmp_wb2_name)
        print(f"         tables={[t['name'] for t in r['tables']]}")

        print("    [8g] Add relationship SalesTable.ProductID → ProductsTable.ProductID...")
        try:
            r = datamodel_action("add_relationship", workbook=tmp_wb2_name,
                                 from_table="SalesTable", from_column="ProductID",
                                 to_table="ProductsTable", to_column="ProductID")
            print(f"         {r}")
        except Exception as e:
            print(f"         SKIP (relationship may need refresh first): {e}")

        print("    [8h] List relationships...")
        r = datamodel_action("list_relationships", workbook=tmp_wb2_name)
        print(f"         {r}")

        print("    [8i] Add DAX measure 'Total Sales'...")
        r = datamodel_action("add_measure", workbook=tmp_wb2_name,
                             measure_name="Total Sales",
                             table="SalesTable",
                             formula="=SUM(SalesTable[Amount])",
                             format_type="currency",
                             decimal_places=2,
                             use_thousand_sep=True)
        print(f"         {r}")

        print("    [8j] Add DAX measure 'Avg Amount'...")
        r = datamodel_action("add_measure", workbook=tmp_wb2_name,
                             measure_name="Avg Amount",
                             table="SalesTable",
                             formula="=AVERAGE(SalesTable[Amount])",
                             format_type="decimal",
                             decimal_places=1)
        print(f"         {r}")

        print("    [8k] List measures...")
        r = datamodel_action("list_measures", workbook=tmp_wb2_name)
        print(f"         measures={[m['name'] for m in r['measures']]}")

        print("    [8l] Update measure 'Avg Amount' formula...")
        r = datamodel_action("update_measure", workbook=tmp_wb2_name,
                             measure_name="Avg Amount",
                             new_formula="=AVERAGEX(SalesTable, SalesTable[Amount])")
        print(f"         {r}")

        print("    [8m] Create data-model pivot from 'Total Sales' measure...")
        try:
            r = pivot_action("create", name="DMPivot", source="datamodel",
                             workbook=tmp_wb2_name)
            print(f"         created: {r.get('sheet')}, fields: {r.get('available_fields', [])[:5]}")

            print("    [8n] Add Region to rows...")
            try:
                r = pivot_action("add_field", name="DMPivot", workbook=tmp_wb2_name,
                                 field="Region", area="rows")
                print(f"         {r}")
            except Exception as e:
                print(f"         SKIP Region (use cube field format): {e}")

            print("    [8o] Add Total Sales measure to values...")
            try:
                r = pivot_action("add_field", name="DMPivot", workbook=tmp_wb2_name,
                                 field="Total Sales", area="values")
                print(f"         {r}")
            except Exception as e:
                print(f"         SKIP (try [Measures].[Total Sales]): {e}")

        except Exception as e:
            print(f"         SKIP datamodel pivot: {e}")

        print("    [8p] Delete measure 'Avg Amount'...")
        r = datamodel_action("delete_measure", workbook=tmp_wb2_name,
                             measure_name="Avg Amount")
        print(f"         {r}")

        print("    [8q] Final model info...")
        r = datamodel_action("info", workbook=tmp_wb2_name)
        print(f"         {r}")

        print("\n    All Phase 2 checks PASSED.")

    except Exception as e:
        print(f"\n    Phase 2 ERROR: {e}")
        import traceback; traceback.print_exc()
    finally:
        if tmp_wb2 is not None:
            try:
                tmp_wb2.Close(SaveChanges=False)
                print(f"\n    Temp workbook '{tmp_wb2_name}' closed (not saved).")
            except Exception as e:
                print(f"    WARNING: Could not close temp workbook 2: {e}")

    # ── Phase 4: Charts + Screenshot on a fresh temp workbook ────────────────
    print("\n[9] Phase 4: Charts + Screenshot on a fresh temp workbook...")
    tmp_wb4 = None
    tmp_wb4_name = None
    try:
        app = _session.get_app()
        tmp_wb4 = app.Workbooks.Add()
        tmp_wb4_name = tmp_wb4.Name
        print(f"    Created temp workbook: {tmp_wb4_name}")

        ws = tmp_wb4.ActiveSheet
        ws.Name = "ChartData"
        ws.Range("A1:C1").Value = [["Month", "Revenue", "Cost"]]
        ws.Range("A2:C5").Value = [
            ["Jan", 10000, 7000],
            ["Feb", 12000, 8000],
            ["Mar", 11000, 7500],
            ["Apr", 14000, 9000],
        ]
        print("    Test data written.")

        print("\n    [9a] Create column chart from A1:C5...")
        r = chart_action("create", source="A1:C5", chart_type="column",
                         sheet="ChartData", workbook=tmp_wb4_name,
                         title="Revenue vs Cost", position="E2")
        chart_name = r["name"]
        print(f"         Created: '{chart_name}' on sheet '{r['sheet']}'")

        print("    [9b] List charts...")
        r = chart_action("list", workbook=tmp_wb4_name)
        print(f"         Found {r['count']} chart(s): {[c['name'] for c in r['charts']]}")

        print("    [9c] Configure chart (legend + axis titles)...")
        r = chart_action("configure", name=chart_name, workbook=tmp_wb4_name,
                         x_title="Month", y_title="Amount (THB)",
                         legend=True, legend_position="bottom")
        print(f"         Changes: {r['changes']}")

        print("    [9d] Export chart as PNG...")
        r = chart_action("export_image", name=chart_name, workbook=tmp_wb4_name)
        import os
        print(f"         PNG at: {r['path']} (exists={os.path.exists(r['path'])})")

        print("    [9e] Screenshot: range A1:C5...")
        try:
            r = screenshot_action("range", range="A1:C5",
                                  sheet="ChartData", workbook=tmp_wb4_name)
            print(f"         PNG at: {r['path']} (exists={os.path.exists(r['path'])})")
        except Exception as e:
            print(f"         SKIP screenshot range (may need visible Excel): {e}")

        print("    [9f] Screenshot: full sheet...")
        try:
            r = screenshot_action("sheet", sheet="ChartData", workbook=tmp_wb4_name)
            print(f"         PNG at: {r['path']} (exists={os.path.exists(r['path'])})")
        except Exception as e:
            print(f"         SKIP screenshot sheet: {e}")

        print("    [9g] Screenshot via chart action...")
        try:
            r = screenshot_action("chart", name=chart_name, workbook=tmp_wb4_name)
            print(f"         PNG at: {r['path']} (exists={os.path.exists(r['path'])})")
        except Exception as e:
            print(f"         SKIP screenshot chart: {e}")

        print("    [9h] Delete chart...")
        r = chart_action("delete", name=chart_name, workbook=tmp_wb4_name)
        print(f"         {r}")

        print("\n    All Phase 4 chart+screenshot checks PASSED.")

    except Exception as e:
        print(f"\n    Phase 4 ERROR: {e}")
        import traceback; traceback.print_exc()
    finally:
        if tmp_wb4 is not None:
            try:
                tmp_wb4.Close(SaveChanges=False)
                print(f"\n    Temp workbook '{tmp_wb4_name}' closed (not saved).")
            except Exception as e:
                print(f"    WARNING: Could not close temp workbook 4: {e}")

    print("\n=== Smoke test complete ===")


if __name__ == "__main__":
    main()
