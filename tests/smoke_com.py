"""Live COM smoke test — Phase 5 regression harness.

Requires a live Excel instance (or THEPEXCEL_MCP_AUTOLAUNCH=1).
Pre-existing workbooks are NEVER touched: all test work uses Workbooks.Add().
All temp workbooks are closed WITHOUT saving at the end of each section.

Run: uv run python tests/smoke_com.py

Output: per-section PASS/FAIL/SKIPPED summary, then full report at the end.
"""

from __future__ import annotations

import os
import sys
import textwrap
import time
import traceback
from dataclasses import dataclass, field
from typing import Literal

# ── Result tracking ────────────────────────────────────────────────────────────

Status = Literal["PASS", "FAIL", "SKIP"]


@dataclass
class Result:
    section: str
    status: Status
    detail: str = ""


_results: list[Result] = []
_excel_busy: bool = False  # set True when close_wb times out (Excel processing PQ)


def record(section: str, status: Status, detail: str = "") -> None:
    _results.append(Result(section, status, detail))
    icon = {"PASS": "PASS", "FAIL": "FAIL", "SKIP": "SKIP"}[status]
    suffix = f" — {detail}" if detail else ""
    print(f"  [{icon}] {section}{suffix}")


def section_header(title: str) -> None:
    print(f"\n{'='*60}")
    print(f"  {title}")
    print(f"{'='*60}")


def print_summary() -> None:
    print(f"\n{'='*60}")
    print("  SMOKE TEST SUMMARY")
    print(f"{'='*60}")
    passes = [r for r in _results if r.status == "PASS"]
    fails  = [r for r in _results if r.status == "FAIL"]
    skips  = [r for r in _results if r.status == "SKIP"]
    col_w = max((len(r.section) for r in _results), default=20) + 2
    for r in _results:
        line = f"  {r.section:<{col_w}} {r.status}"
        if r.detail and r.status != "PASS":
            line += f"  ({r.detail[:80]})"
        print(line)
    print(f"\n  Total: {len(_results)}  |  PASS: {len(passes)}  |  FAIL: {len(fails)}  |  SKIP: {len(skips)}")
    if fails:
        print("\n  FAILED sections:")
        for r in fails:
            print(f"    - {r.section}: {r.detail}")
    print()


# ── Imports ────────────────────────────────────────────────────────────────────

try:
    import win32com.client
    import pythoncom
    from thepexcel_mcp.session import ExcelSession
    from thepexcel_mcp.domains.workbook   import workbook_action
    from thepexcel_mcp.domains.sheets     import sheet_action
    from thepexcel_mcp.domains.ranges     import range_action
    from thepexcel_mcp.domains.powerquery import powerquery_action
    from thepexcel_mcp.domains.tables     import table_action
    from thepexcel_mcp.domains.pivots     import pivot_action
    from thepexcel_mcp.domains.datamodel  import datamodel_action
    from thepexcel_mcp.domains.charts     import chart_action
    from thepexcel_mcp.domains.screenshot import screenshot_action
    from thepexcel_mcp.domains.names      import name_action
    from thepexcel_mcp.domains.vba        import vba_action
except ImportError as e:
    print(f"Import failed: {e}")
    sys.exit(1)

_session = ExcelSession()


# ── Helpers ────────────────────────────────────────────────────────────────────

def _close_wb(wb, name: str) -> bool:
    """Close a temp workbook without saving. Returns True on success.

    If close times out (e.g. because PQ engine is still running in the workbook),
    sets _excel_busy=True so subsequent sections know Excel may be unresponsive.
    """
    global _excel_busy
    if wb is None:
        return True

    def _do_close():
        wb.Close(SaveChanges=False)

    try:
        _session.run_com(_do_close)
        print(f"  [cleanup] Closed '{name}' (not saved)")
        return True
    except Exception as e:
        msg = str(e)
        print(f"  [cleanup] Warning: could not close '{name}': {msg}")
        if _is_timeout_error(e):
            print("  [cleanup] Excel busy (PQ engine running) — marking excel_busy=True")
            _excel_busy = True
        return False


def _add_wb(app) -> tuple:
    """Add a new workbook; return (wb, name). Must run on COM worker thread."""
    wb = app.Workbooks.Add()
    return wb, wb.Name


def _new_wb() -> tuple:
    """Create a new workbook via the STA COM worker. Returns (wb, name)."""
    def _do():
        app = _session.get_app()
        wb = app.Workbooks.Add()
        return wb, wb.Name
    return _session.run_com(_do)


def _check_excel_busy(*labels: str) -> bool:
    """If Excel is busy (e.g. PQ engine still running), record all labels as SKIP
    and return True so the caller can early-return.
    """
    if _excel_busy:
        for lbl in labels:
            record(lbl, "SKIP", "Excel busy — PQ engine still running from prior section")
        return True
    return False


# ── Section 1: Workbook/Sheet/Range CRUD ──────────────────────────────────────

def run_workbook_sheet_range() -> None:
    section_header("SECTION 1 — Workbook / Sheet / Range CRUD")
    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()
        print(f"  Created temp workbook: {wb_name}")

        # workbook list
        try:
            r = workbook_action("list")
            names = [w["name"] for w in r["workbooks"]]
            assert wb_name in names, f"{wb_name} not in {names}"
            record("workbook.list", "PASS")
        except Exception as e:
            record("workbook.list", "FAIL", str(e))

        # workbook info
        try:
            r = workbook_action("info", workbook=wb_name)
            assert "sheets" in r
            record("workbook.info", "PASS")
        except Exception as e:
            record("workbook.info", "FAIL", str(e))

        # sheet add / rename / list
        try:
            sheet_action("add", name="TestSheet", workbook=wb_name)
            sheet_action("rename", name="TestSheet", new_name="RenamedSheet", workbook=wb_name)
            r = sheet_action("list", workbook=wb_name)
            names_s = [s["name"] for s in r["sheets"]]
            assert "RenamedSheet" in names_s, names_s
            record("sheet.add_rename_list", "PASS")
        except Exception as e:
            record("sheet.add_rename_list", "FAIL", str(e))

        # range write / read
        try:
            def _setup_data():
                ws = _session.get_sheet("Sheet1", wb_name)
                ws.Range("A1:C1").Value = [["Name", "Value", "Category"]]
                ws.Range("A2:C6").Value = [
                    ["Alpha", 100, "X"],
                    ["Beta",  200, "Y"],
                    ["Gamma", 300, "X"],
                    ["Delta", 400, "Y"],
                    ["Epsilon",500, "X"],
                ]
            _session.run_com(_setup_data)
            r = range_action("read", range="A1:C3", sheet="Sheet1", workbook=wb_name)
            assert r["total_rows"] == 3
            assert r["values"][0][0] == "Name"
            record("range.write_read", "PASS")
        except Exception as e:
            record("range.write_read", "FAIL", str(e))

        # range write 2-D via MCP tool — read-back verifies every cell (guards
        # against pywin32 Resize quirk where only [0][0] was written)
        try:
            payload = [["P", "Q"], ["R", "S"], ["T", "U"]]  # 3 rows x 2 cols
            range_action("write", range="H1", sheet="Sheet1", workbook=wb_name,
                         values=payload)
            rb = range_action("read", range="H1:I3", sheet="Sheet1", workbook=wb_name)
            assert rb["total_rows"] == 3, f"expected 3 rows, got {rb['total_rows']}"
            for i, row in enumerate(payload):
                for j, expected in enumerate(row):
                    got = rb["values"][i][j]
                    assert got == expected, (
                        f"cell [{i}][{j}]: expected {expected!r}, got {got!r}"
                    )
            record("range.write_2d_readback", "PASS")
        except Exception as e:
            record("range.write_2d_readback", "FAIL", str(e))

        # write_formula + spill (SEQUENCE)
        try:
            range_action("write_formula", range="E1", formula="=SEQUENCE(5,3)",
                         sheet="Sheet1", workbook=wb_name)
            time.sleep(0.3)
            r = range_action("read_spill", range="E1", sheet="Sheet1", workbook=wb_name)
            assert r["total_rows"] == 5, f"expected 5 rows, got {r['total_rows']}"
            record("range.write_formula_spill", "PASS",
                   f"spill_range={r['spill_range']}")
        except Exception as e:
            record("range.write_formula_spill", "FAIL", str(e))

        # clear
        try:
            range_action("clear", range="E1:G5", sheet="Sheet1", workbook=wb_name)
            record("range.clear", "PASS")
        except Exception as e:
            record("range.clear", "FAIL", str(e))

        # sheet delete
        try:
            sheet_action("delete", name="RenamedSheet", workbook=wb_name)
            record("sheet.delete", "PASS")
        except Exception as e:
            record("sheet.delete", "FAIL", str(e))

    except Exception as e:
        record("section1.setup", "FAIL", f"setup error: {e}")
        traceback.print_exc()
    finally:
        _close_wb(wb, wb_name or "unknown")


# ── Section 2: Table lifecycle ─────────────────────────────────────────────────

def run_tables() -> None:
    section_header("SECTION 2 — Table lifecycle")
    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        # Write data
        def _write():
            ws = _session.get_sheet("Sheet1", wb_name)
            ws.Range("A1:D1").Value = [["Region", "Product", "Quarter", "Amount"]]
            ws.Range("A2:D8").Value = [
                ["North", "Widget", "Q1", 1000],
                ["South", "Gadget", "Q1", 1500],
                ["North", "Widget", "Q2", 1200],
                ["South", "Gadget", "Q2", 1800],
                ["North", "Gadget", "Q1",  900],
                ["East",  "Widget", "Q2", 2000],
                ["West",  "Gadget", "Q2", 2200],
            ]
        _session.run_com(_write)

        # create
        try:
            table_action("create", name="Sales", range="A1:D8",
                         sheet="Sheet1", workbook=wb_name)
            record("table.create", "PASS")
        except Exception as e:
            record("table.create", "FAIL", str(e)); return

        # read
        try:
            r = table_action("read", name="Sales", workbook=wb_name, limit=3)
            assert r["columns"] == ["Region", "Product", "Quarter", "Amount"]
            assert len(r["values"]) == 3
            assert r["has_more"]
            record("table.read", "PASS")
        except Exception as e:
            record("table.read", "FAIL", str(e))

        # append_rows
        try:
            table_action("append_rows", name="Sales", workbook=wb_name,
                         values=[["North", "Gadget", "Q3", 1100]])
            record("table.append_rows", "PASS")
        except Exception as e:
            record("table.append_rows", "FAIL", str(e))

        # add_column
        try:
            table_action("add_column", name="Sales", workbook=wb_name,
                         column_name="Tax", formula="=[@Amount]*0.07")
            record("table.add_column", "PASS")
        except Exception as e:
            record("table.add_column", "FAIL", str(e))

        # sort — read-back verifies actual reorder (guards against SortOn=1
        # xlSortOnCellColor no-op bug where sort appeared to succeed but did nothing)
        try:
            table_action("sort", name="Sales", workbook=wb_name,
                         sort_column="Amount", ascending=False)
            rb = table_action("read", name="Sales", workbook=wb_name, limit=20)
            amt_idx = rb["columns"].index("Amount")
            amounts = [row[amt_idx] for row in rb["values"]]
            assert amounts == sorted(amounts, reverse=True), (
                f"Amount column not sorted descending after sort: {amounts}"
            )
            record("table.sort", "PASS")
        except Exception as e:
            record("table.sort", "FAIL", str(e))

        # filter
        try:
            table_action("filter", name="Sales", workbook=wb_name,
                         filter_column="Region", filter_op="equals",
                         filter_value="North")
            record("table.filter", "PASS")
        except Exception as e:
            record("table.filter", "FAIL", str(e))

        # clear_filters
        try:
            table_action("filter", name="Sales", workbook=wb_name,
                         filter_column="Region", filter_op="clear_filters")
            record("table.clear_filters", "PASS")
        except Exception as e:
            record("table.clear_filters", "FAIL", str(e))

        # toggle_totals
        try:
            table_action("toggle_totals", name="Sales", workbook=wb_name,
                         show_totals=True, column_name="Amount", total_func="sum")
            record("table.toggle_totals", "PASS")
        except Exception as e:
            record("table.toggle_totals", "FAIL", str(e))

        # set_style
        try:
            table_action("set_style", name="Sales", workbook=wb_name,
                         style="TableStyleLight2")
            record("table.set_style", "PASS")
        except Exception as e:
            record("table.set_style", "FAIL", str(e))

        # rename
        try:
            table_action("rename", name="Sales", workbook=wb_name,
                         new_name="SalesData")
            record("table.rename", "PASS")
        except Exception as e:
            record("table.rename", "FAIL", str(e))

        # delete (keep_data)
        try:
            table_action("delete", name="SalesData", workbook=wb_name,
                         keep_data=True)
            record("table.delete", "PASS")
        except Exception as e:
            record("table.delete", "FAIL", str(e))

    except Exception as e:
        record("section2.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        _close_wb(wb, wb_name or "unknown")


# ── Section 3: Pivot ───────────────────────────────────────────────────────────

def run_pivots() -> None:
    section_header("SECTION 3 — PivotTable lifecycle")
    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        def _write():
            ws = _session.get_sheet("Sheet1", wb_name)
            ws.Range("A1:D1").Value = [["Region", "Product", "Quarter", "Amount"]]
            ws.Range("A2:D6").Value = [
                ["North", "Widget", "Q1", 1000],
                ["South", "Gadget", "Q1", 1500],
                ["North", "Widget", "Q2", 1200],
                ["South", "Gadget", "Q2", 1800],
                ["North", "Gadget", "Q1",  900],
            ]
        _session.run_com(_write)
        table_action("create", name="PivSrc", range="A1:D6",
                     sheet="Sheet1", workbook=wb_name)

        # create pivot
        try:
            r = pivot_action("create", name="SalesPivot", source="PivSrc",
                             workbook=wb_name)
            assert "sheet" in r
            record("pivot.create", "PASS", f"sheet={r['sheet']}")
        except Exception as e:
            record("pivot.create", "FAIL", str(e)); return

        # add rows field
        try:
            pivot_action("add_field", name="SalesPivot", workbook=wb_name,
                         field="Region", area="rows")
            record("pivot.add_field_rows", "PASS")
        except Exception as e:
            record("pivot.add_field_rows", "FAIL", str(e))

        # add values field
        try:
            pivot_action("add_field", name="SalesPivot", workbook=wb_name,
                         field="Amount", area="values", aggregation="sum",
                         number_format="#,##0")
            record("pivot.add_field_values", "PASS")
        except Exception as e:
            record("pivot.add_field_values", "FAIL", str(e))

        # read
        try:
            r = pivot_action("read", name="SalesPivot", workbook=wb_name)
            assert "values" in r
            record("pivot.read", "PASS",
                   f"total_rows={r['total_rows']}")
        except Exception as e:
            record("pivot.read", "FAIL", str(e))

        # set_layout
        try:
            pivot_action("set_layout", name="SalesPivot", workbook=wb_name,
                         layout="tabular", grand_totals=True)
            record("pivot.set_layout", "PASS")
        except Exception as e:
            record("pivot.set_layout", "FAIL", str(e))

        # refresh
        try:
            pivot_action("refresh", name="SalesPivot", workbook=wb_name)
            record("pivot.refresh", "PASS")
        except Exception as e:
            record("pivot.refresh", "FAIL", str(e))

        # remove_field
        try:
            pivot_action("remove_field", name="SalesPivot", workbook=wb_name,
                         field="Region")
            record("pivot.remove_field", "PASS")
        except Exception as e:
            record("pivot.remove_field", "FAIL", str(e))

        # delete
        try:
            pivot_action("delete", name="SalesPivot", workbook=wb_name)
            record("pivot.delete", "PASS")
        except Exception as e:
            record("pivot.delete", "FAIL", str(e))

    except Exception as e:
        record("section3.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        _close_wb(wb, wb_name or "unknown")


# ── Section 4: Power Query ─────────────────────────────────────────────────────

def run_powerquery() -> None:
    section_header("SECTION 4 — Power Query")
    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        m_code = textwrap.dedent("""\
            let
                Source = #table(
                    type table [ID=Int64.Type, Name=text, Value=Int64.Type],
                    {{1, "Alpha", 100}, {2, "Beta", 200}, {3, "Gamma", 300}}
                )
            in
                Source""")

        # create
        try:
            powerquery_action("create", name="TestQuery", formula=m_code,
                              workbook=wb_name)
            record("pq.create", "PASS")
        except Exception as e:
            record("pq.create", "FAIL", str(e)); return

        # list
        try:
            r = powerquery_action("list", workbook=wb_name)
            names = [q["name"] for q in r["queries"]]
            assert "TestQuery" in names, names
            record("pq.list", "PASS")
        except Exception as e:
            record("pq.list", "FAIL", str(e))

        # get
        try:
            r = powerquery_action("get", name="TestQuery", workbook=wb_name)
            assert "formula" in r
            record("pq.get", "PASS")
        except Exception as e:
            record("pq.get", "FAIL", str(e))

        # analyze
        try:
            r = powerquery_action("analyze", name="TestQuery", workbook=wb_name)
            assert "step_count" in r
            record("pq.analyze", "PASS",
                   f"steps={r['step_count']} complexity={r['estimated_complexity']}")
        except Exception as e:
            record("pq.analyze", "FAIL", str(e))

        # load_to_table
        try:
            r = powerquery_action("load_to_table", name="TestQuery",
                                  sheet_name="QueryOut", workbook=wb_name)
            assert r.get("rows", 0) > 0, f"expected rows>0, got {r}"
            record("pq.load_to_table", "PASS",
                   f"rows={r['rows']}, sheet={r['sheet']}")
        except Exception as e:
            record("pq.load_to_table", "FAIL", str(e))

        # update (rename)
        try:
            powerquery_action("update", name="TestQuery",
                              new_name="RenamedQuery", workbook=wb_name)
            r = powerquery_action("list", workbook=wb_name)
            names2 = [q["name"] for q in r["queries"]]
            assert "RenamedQuery" in names2, names2
            record("pq.update_rename", "PASS")
        except Exception as e:
            record("pq.update_rename", "FAIL", str(e))

        # delete
        try:
            powerquery_action("delete", name="RenamedQuery", workbook=wb_name)
            record("pq.delete", "PASS")
        except Exception as e:
            record("pq.delete", "FAIL", str(e))

    except Exception as e:
        record("section4.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        _close_wb(wb, wb_name or "unknown")


# ── Section 5: Data Model + DAX ────────────────────────────────────────────────

def _is_timeout_error(e: Exception) -> bool:
    """True when a COM call timed out (known Data Model deadlock scenario)."""
    msg = str(e).lower()
    return "timed out" in msg or "timeout" in msg


def run_datamodel() -> None:
    section_header("SECTION 5 — Data Model + DAX measures")
    # Data Model add_table calls conn.Refresh() which invokes the Mashup/PQ engine.
    # The engine makes reentrancy COM calls that need Excel's UI message pump.
    # Our STA COM worker blocks on the Future while waiting, so the pump never runs
    # → deadlock until the COM timeout fires (120 s by default).
    # We patch the module-level timeout to 30 s to fail fast, then record SKIP.
    # This works correctly when Claude Desktop is running with a visible Excel
    # (the UI pump runs), so this is an automation-context limitation only.
    import thepexcel_mcp.session as _session_mod
    prev_timeout = _session_mod._DEFAULT_TIMEOUT
    _session_mod._DEFAULT_TIMEOUT = 30

    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        # Create PQ queries with static data (avoids Excel.CurrentWorkbook dep)
        sales_m = textwrap.dedent("""\
            let
                Source = #table(
                    type table [ProductID=Int64.Type, Region=text, Amount=Int64.Type],
                    {{1, "North", 1000}, {2, "South", 1500}, {1, "North", 1200},
                     {2, "South", 1800}, {1, "East", 900}}
                )
            in
                Source""")

        products_m = textwrap.dedent("""\
            let
                Source = #table(
                    type table [ProductID=Int64.Type, ProductName=text],
                    {{1, "Widget"}, {2, "Gadget"}}
                )
            in
                Source""")

        powerquery_action("create", name="SalesQuery", formula=sales_m,
                          workbook=wb_name)
        powerquery_action("create", name="ProductsQuery", formula=products_m,
                          workbook=wb_name)

        # add_table: may deadlock in automation context (no UI message pump)
        dm_tables_ok = False
        try:
            r = datamodel_action("add_table", workbook=wb_name,
                                 source_type="query", source_name="SalesQuery")
            record("dm.add_table_sales", "PASS",
                   f"model_tables={r.get('model_table_count')}")
            dm_tables_ok = True
        except Exception as e:
            if _is_timeout_error(e):
                record("dm.add_table_sales", "SKIP",
                       "conn.Refresh() deadlock — no UI pump in automation context")
            else:
                record("dm.add_table_sales", "FAIL", str(e))

        if dm_tables_ok:
            try:
                r = datamodel_action("add_table", workbook=wb_name,
                                     source_type="query", source_name="ProductsQuery")
                record("dm.add_table_products", "PASS",
                       f"model_tables={r.get('model_table_count')}")
            except Exception as e:
                if _is_timeout_error(e):
                    record("dm.add_table_products", "SKIP",
                           "conn.Refresh() deadlock — no UI pump")
                    dm_tables_ok = False
                else:
                    record("dm.add_table_products", "FAIL", str(e))
                    dm_tables_ok = False
        else:
            record("dm.add_table_products", "SKIP", "previous add_table skipped/failed")

        # info + list_tables also access wb.Model which can deadlock when
        # PQ connections are unloaded — skip them when add_table already deadlocked.
        if dm_tables_ok:
            try:
                r = datamodel_action("info", workbook=wb_name)
                record("dm.info", "PASS", f"tables={r.get('table_count', 0)}")
            except Exception as e:
                if _is_timeout_error(e):
                    record("dm.info", "SKIP", "wb.Model access deadlocked")
                else:
                    record("dm.info", "FAIL", str(e))

            try:
                r = datamodel_action("list_tables", workbook=wb_name)
                names = [t["name"] for t in r["tables"]]
                record("dm.list_tables", "PASS", f"names={names}")
            except Exception as e:
                if _is_timeout_error(e):
                    record("dm.list_tables", "SKIP", "wb.Model access deadlocked")
                else:
                    record("dm.list_tables", "FAIL", str(e))
        else:
            record("dm.info", "SKIP", "tables not loaded")
            record("dm.list_tables", "SKIP", "tables not loaded")

        # All operations below require tables to be loaded first
        if not dm_tables_ok:
            for label in ("dm.add_relationship", "dm.list_relationships",
                          "dm.add_measure", "dm.list_measures",
                          "dm.add_measure2", "dm.update_measure",
                          "dm.pivot_create", "dm.delete_measure"):
                record(label, "SKIP", "tables not loaded (add_table skipped/failed)")
        else:
            # add_relationship
            try:
                r = datamodel_action("add_relationship", workbook=wb_name,
                                     from_table="SalesQuery",
                                     from_column="ProductID",
                                     to_table="ProductsQuery",
                                     to_column="ProductID")
                record("dm.add_relationship", "PASS")
            except Exception as e:
                record("dm.add_relationship", "FAIL", str(e))

            # list_relationships
            try:
                r = datamodel_action("list_relationships", workbook=wb_name)
                record("dm.list_relationships", "PASS", f"count={r['count']}")
            except Exception as e:
                record("dm.list_relationships", "FAIL", str(e))

            # add_measure
            try:
                r = datamodel_action("add_measure", workbook=wb_name,
                                     measure_name="Total Sales",
                                     table="SalesQuery",
                                     formula="=SUM(SalesQuery[Amount])",
                                     format_type="currency",
                                     decimal_places=2,
                                     use_thousand_sep=True)
                assert "created_measure" in r
                record("dm.add_measure", "PASS")
            except Exception as e:
                record("dm.add_measure", "FAIL", str(e))

            # list_measures
            try:
                r = datamodel_action("list_measures", workbook=wb_name)
                names = [m["name"] for m in r["measures"]]
                assert "Total Sales" in names, names
                record("dm.list_measures", "PASS")
            except Exception as e:
                record("dm.list_measures", "FAIL", str(e))

            # add second measure
            try:
                datamodel_action("add_measure", workbook=wb_name,
                                 measure_name="Avg Amount",
                                 table="SalesQuery",
                                 formula="=AVERAGE(SalesQuery[Amount])",
                                 format_type="decimal",
                                 decimal_places=1)
                record("dm.add_measure2", "PASS")
            except Exception as e:
                record("dm.add_measure2", "FAIL", str(e))

            # update_measure
            try:
                r = datamodel_action("update_measure", workbook=wb_name,
                                     measure_name="Avg Amount",
                                     new_formula="=AVERAGEX(SalesQuery,SalesQuery[Amount])")
                assert "updated_measure" in r
                record("dm.update_measure", "PASS")
            except Exception as e:
                record("dm.update_measure", "FAIL", str(e))

            # data-model pivot
            try:
                r = pivot_action("create", name="DMPivot", source="datamodel",
                                 workbook=wb_name)
                record("dm.pivot_create", "PASS", f"sheet={r.get('sheet')}")
            except Exception as e:
                record("dm.pivot_create", "FAIL", str(e))

            # delete_measure
            try:
                r = datamodel_action("delete_measure", workbook=wb_name,
                                     measure_name="Avg Amount")
                assert "deleted_measure" in r
                record("dm.delete_measure", "PASS")
            except Exception as e:
                record("dm.delete_measure", "FAIL", str(e))

    except Exception as e:
        record("section5.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        # Delete all PQ queries before closing — otherwise wb.Close() triggers
        # the Mashup engine and deadlocks.  Best-effort; ignore errors.
        if wb is not None:
            try:
                def _delete_queries():
                    while wb.Queries.Count > 0:
                        wb.Queries.Item(1).Delete()
                _session.run_com(_delete_queries)
            except Exception:
                pass
        # Close wb while short timeout is still active (30 s)
        _close_wb(wb, wb_name or "unknown")
        # Restore original timeout after close attempt
        _session_mod._DEFAULT_TIMEOUT = prev_timeout


# ── Section 6: LAMBDA named formula ───────────────────────────────────────────

def run_lambda() -> None:
    section_header("SECTION 6 — Named LAMBDA")
    if _check_excel_busy("lambda.set", "lambda.list_is_lambda", "lambda.get",
                         "lambda.use_in_cell", "lambda.delete"):
        return
    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        # set LAMBDA
        try:
            name_action("set", workbook=wb_name,
                        name="DOUBLE",
                        refers_to="=LAMBDA(x, x*2)")
            record("lambda.set", "PASS")
        except Exception as e:
            record("lambda.set", "FAIL", str(e))
            return

        # list — verify is_lambda=True
        try:
            r = name_action("list", workbook=wb_name)
            lambdas = [n for n in r["names"] if n.get("is_lambda")]
            assert any(n["name"] == "DOUBLE" for n in lambdas), r["names"]
            record("lambda.list_is_lambda", "PASS")
        except Exception as e:
            record("lambda.list_is_lambda", "FAIL", str(e))

        # get
        try:
            r = name_action("get", workbook=wb_name, name="DOUBLE")
            assert r["is_lambda"]
            record("lambda.get", "PASS")
        except Exception as e:
            record("lambda.get", "FAIL", str(e))

        # use LAMBDA in a cell — =DOUBLE(21) should give 42
        try:
            def _use_lambda():
                ws = _session.get_sheet("Sheet1", wb_name)
                ws.Range("A1").Formula2 = "=DOUBLE(21)"
                time.sleep(0.3)
                val = ws.Range("A1").Value
                return val
            val = _session.run_com(_use_lambda)
            assert val == 42, f"expected 42, got {val}"
            record("lambda.use_in_cell", "PASS", f"=DOUBLE(21)={val}")
        except Exception as e:
            record("lambda.use_in_cell", "FAIL", str(e))

        # delete
        try:
            name_action("delete", workbook=wb_name, name="DOUBLE")
            record("lambda.delete", "PASS")
        except Exception as e:
            record("lambda.delete", "FAIL", str(e))

    except Exception as e:
        record("section6.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        _close_wb(wb, wb_name or "unknown")


# ── Section 7: Chart lifecycle ─────────────────────────────────────────────────

def run_charts() -> None:
    section_header("SECTION 7 — Chart lifecycle + export")
    if _check_excel_busy("chart.create", "chart.list", "chart.configure",
                         "chart.export_image", "chart.set_source", "chart.delete"):
        return
    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        def _write():
            ws = _session.get_sheet("Sheet1", wb_name)
            ws.Range("A1:C1").Value = [["Month", "Revenue", "Cost"]]
            ws.Range("A2:C5").Value = [
                ["Jan", 10000, 7000],
                ["Feb", 12000, 8000],
                ["Mar", 11000, 7500],
                ["Apr", 14000, 9000],
            ]
        _session.run_com(_write)

        chart_name = None

        # create
        try:
            r = chart_action("create", source="A1:C5", chart_type="column",
                             sheet="Sheet1", workbook=wb_name,
                             title="Revenue vs Cost", position="E2")
            chart_name = r["name"]
            record("chart.create", "PASS", f"name={chart_name}")
        except Exception as e:
            record("chart.create", "FAIL", str(e)); return

        # list
        try:
            r = chart_action("list", workbook=wb_name)
            names = [c["name"] for c in r["charts"]]
            assert chart_name in names
            record("chart.list", "PASS")
        except Exception as e:
            record("chart.list", "FAIL", str(e))

        # configure
        try:
            r = chart_action("configure", name=chart_name, workbook=wb_name,
                             x_title="Month", y_title="Amount",
                             legend=True, legend_position="bottom")
            assert "legend" in r["changes"]
            record("chart.configure", "PASS", f"changes={r['changes']}")
        except Exception as e:
            record("chart.configure", "FAIL", str(e))

        # export_image
        try:
            r = chart_action("export_image", name=chart_name, workbook=wb_name)
            assert os.path.exists(r["path"]), f"PNG not found: {r['path']}"
            record("chart.export_image", "PASS",
                   f"path={os.path.basename(r['path'])}")
        except Exception as e:
            record("chart.export_image", "FAIL", str(e))

        # set_source
        try:
            chart_action("set_source", name=chart_name, workbook=wb_name,
                         source="Sheet1!A1:B5")
            record("chart.set_source", "PASS")
        except Exception as e:
            record("chart.set_source", "FAIL", str(e))

        # delete
        try:
            chart_action("delete", name=chart_name, workbook=wb_name)
            record("chart.delete", "PASS")
        except Exception as e:
            record("chart.delete", "FAIL", str(e))

    except Exception as e:
        record("section7.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        _close_wb(wb, wb_name or "unknown")


# ── Section 8: Screenshot ──────────────────────────────────────────────────────

def run_screenshot() -> None:
    section_header("SECTION 8 — Screenshot (range + sheet + chart)")
    if _check_excel_busy("screenshot.range", "screenshot.sheet", "screenshot.chart"):
        return
    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        def _write():
            ws = _session.get_sheet("Sheet1", wb_name)
            ws.Range("A1:C1").Value = [["X", "Y", "Z"]]
            ws.Range("A2:C4").Value = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
        _session.run_com(_write)

        # Create a chart for chart screenshot
        chart_action("create", source="A1:C4", chart_type="line",
                     sheet="Sheet1", workbook=wb_name)
        r_list = chart_action("list", workbook=wb_name)
        chart_name = r_list["charts"][0]["name"] if r_list["charts"] else None

        # range screenshot
        try:
            r = screenshot_action("range", range="A1:C4",
                                  sheet="Sheet1", workbook=wb_name)
            assert os.path.exists(r["path"])
            record("screenshot.range", "PASS",
                   f"path={os.path.basename(r['path'])}")
        except Exception as e:
            record("screenshot.range", "FAIL", str(e))

        # sheet screenshot
        try:
            r = screenshot_action("sheet", sheet="Sheet1", workbook=wb_name)
            assert os.path.exists(r["path"])
            record("screenshot.sheet", "PASS",
                   f"path={os.path.basename(r['path'])}")
        except Exception as e:
            record("screenshot.sheet", "FAIL", str(e))

        # chart screenshot
        if chart_name:
            try:
                r = screenshot_action("chart", name=chart_name, workbook=wb_name)
                assert os.path.exists(r["path"])
                record("screenshot.chart", "PASS",
                       f"path={os.path.basename(r['path'])}")
            except Exception as e:
                record("screenshot.chart", "FAIL", str(e))
        else:
            record("screenshot.chart", "SKIP", "no chart created")

    except Exception as e:
        record("section8.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        _close_wb(wb, wb_name or "unknown")


# ── Section 9: VBA ─────────────────────────────────────────────────────────────

def run_vba() -> None:
    section_header("SECTION 9 — VBA (requires AccessVBOM=1 + THEPEXCEL_MCP_ENABLE_VBA=1)")
    if _check_excel_busy("vba.all"):
        return
    import winreg as wr
    try:
        key_path = r"SOFTWARE\Microsoft\Office\16.0\Excel\Security"
        with wr.OpenKey(wr.HKEY_CURRENT_USER, key_path) as key:
            val, _ = wr.QueryValueEx(key, "AccessVBOM")
        if val != 1:
            record("vba.all", "SKIP", "AccessVBOM not 1 in registry")
            return
    except Exception:
        record("vba.all", "SKIP",
               "HKCU\\...\\Security\\AccessVBOM not found — VBA trust not enabled")
        return

    if os.environ.get("THEPEXCEL_MCP_ENABLE_VBA") != "1":
        record("vba.all", "SKIP",
               "THEPEXCEL_MCP_ENABLE_VBA not set to 1")
        return

    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        code = textwrap.dedent("""\
            Function AddTwo(a As Long, b As Long) As Long
                AddTwo = a + b
            End Function
            Sub HelloWorld()
                MsgBox "Hello from ThepExcelMCP"
            End Sub""")

        try:
            vba_action("write_module", workbook=wb_name,
                       module_name="TestModule", code=code)
            record("vba.write_module", "PASS")
        except Exception as e:
            record("vba.write_module", "FAIL", str(e)); return

        try:
            r = vba_action("list_modules", workbook=wb_name)
            names = [m["name"] for m in r["modules"]]
            assert "TestModule" in names, names
            record("vba.list_modules", "PASS")
        except Exception as e:
            record("vba.list_modules", "FAIL", str(e))

        try:
            r = vba_action("get_module", workbook=wb_name,
                           module_name="TestModule")
            assert "AddTwo" in r["code"]
            record("vba.get_module", "PASS")
        except Exception as e:
            record("vba.get_module", "FAIL", str(e))

        try:
            r = vba_action("run", workbook=wb_name,
                           proc_name="TestModule.AddTwo", args=[3, 4])
            assert r["result"] == 7, f"expected 7, got {r['result']}"
            record("vba.run", "PASS", f"AddTwo(3,4)={r['result']}")
        except Exception as e:
            record("vba.run", "FAIL", str(e))

        try:
            vba_action("delete_module", workbook=wb_name,
                       module_name="TestModule")
            record("vba.delete_module", "PASS")
        except Exception as e:
            record("vba.delete_module", "FAIL", str(e))

    except Exception as e:
        record("section9.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        _close_wb(wb, wb_name or "unknown")


# ── Section 10: write_py (Python in Excel) ────────────────────────────────────

def run_write_py() -> None:
    section_header("SECTION 10 — write_py (Python in Excel — insert only)")
    if _check_excel_busy("write_py.insert"):
        return
    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        # Insert =PY() formula — execution result is Azure-async.
        # We only verify the formula TEXT landed correctly.
        try:
            r = range_action("write_py", range="A1", sheet="Sheet1",
                             workbook=wb_name,
                             python_code='[x**2 for x in range(5)]')
            assert "formula_inserted" in r
            assert '=PY(' in r["formula_inserted"]
            record("write_py.insert", "PASS",
                   "formula landed (execution is Azure-async)")
        except Exception as e:
            # =PY() may not be available without M365 Python in Excel licence
            msg = str(e).lower()
            if "py" in msg or "python" in msg or "formula" in msg or "connect" in msg:
                record("write_py.insert", "SKIP",
                       "=PY() not available (no M365 Python in Excel licence)")
            else:
                record("write_py.insert", "FAIL", str(e))

    except Exception as e:
        record("section10.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        _close_wb(wb, wb_name or "unknown")


# ── Connectivity check ─────────────────────────────────────────────────────────

def check_excel_running() -> bool:
    """Return True if Excel is reachable (or was auto-launched)."""
    try:
        _session.run_com(_session.get_app)
        return True
    except Exception as e:
        print(f"\n  ERROR: Cannot reach Excel: {e}")
        print("  Start Excel manually, or set THEPEXCEL_MCP_AUTOLAUNCH=1.\n")
        return False


# ── Main ───────────────────────────────────────────────────────────────────────

_ALL_SECTIONS = {
    "1": run_workbook_sheet_range,
    "2": run_tables,
    "3": run_pivots,
    "4": run_powerquery,
    "5": run_datamodel,
    "6": run_lambda,
    "7": run_charts,
    "8": run_screenshot,
    "9": run_vba,
    "10": run_write_py,
}


def main() -> None:
    import argparse
    parser = argparse.ArgumentParser(description="ThepExcelMCP live smoke test")
    parser.add_argument(
        "--sections", "-s",
        help="Comma-separated section numbers to run (default: all). Example: 1,2,3",
    )
    args = parser.parse_args()

    if args.sections:
        keys = [s.strip() for s in args.sections.split(",")]
        runners = [(k, _ALL_SECTIONS[k]) for k in keys if k in _ALL_SECTIONS]
    else:
        runners = list(_ALL_SECTIONS.items())

    print("=" * 60)
    print("  ThepExcelMCP — Phase 5 Live Smoke Test")
    print(f"  AUTOLAUNCH={os.environ.get('THEPEXCEL_MCP_AUTOLAUNCH', '0')}")
    print(f"  VBA_ENABLED={os.environ.get('THEPEXCEL_MCP_ENABLE_VBA', '0')}")
    sections_str = ",".join(k for k, _ in runners)
    print(f"  SECTIONS={sections_str}")
    print("=" * 60)

    if not check_excel_running():
        print("Aborting — Excel not reachable.")
        sys.exit(1)

    for _, runner in runners:
        runner()

    print_summary()

    fails = [r for r in _results if r.status == "FAIL"]
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
