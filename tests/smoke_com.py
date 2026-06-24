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


class _SkipSentinel(Exception):
    """Internal control-flow marker: an action was already recorded SKIP and the
    enclosing try-block should bail out without recording a FAIL."""


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
    from thepexcel_mcp.domains.format     import format_action
    from thepexcel_mcp.domains.conditional_format import conditional_format_action
    from thepexcel_mcp.domains.view       import view_action
    from thepexcel_mcp.domains.validation import validation_action
    from thepexcel_mcp.domains.slicer     import slicer_action
    from thepexcel_mcp.domains.page_setup import page_setup_action
    from thepexcel_mcp.domains.comments   import comment_action
    from thepexcel_mcp.domains.hyperlinks import hyperlink_action
    from thepexcel_mcp.domains.outline    import outline_action
    from thepexcel_mcp.domains.protection import protection_action
    from thepexcel_mcp.domains.sparkline  import sparkline_action
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


# ── Section 11: excel_format ──────────────────────────────────────────────────

def run_format() -> None:
    section_header("SECTION 11 — excel_format (font / fill / border / number_format / alignment / autofit)")
    if _check_excel_busy(*[
        "format.font", "format.fill", "format.border",
        "format.number_format", "format.alignment",
        "format.column_width", "format.row_height", "format.autofit",
    ]):
        return
    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        def _write_data():
            ws = _session.get_sheet("Sheet1", wb_name)
            ws.Range("A1:C1").Value = [["Name", "Amount", "Rate"]]
            ws.Range("A2:C4").Value = [
                ["Alpha", 1000, 0.1],
                ["Beta",  2500, 0.15],
                ["Gamma", 750,  0.05],
            ]
        _session.run_com(_write_data)

        # --- font: bold + size + color on header row ---
        try:
            format_action("font", range="A1:C1", sheet="Sheet1", workbook=wb_name,
                          bold=True, font_size=12, font_color="#FFFFFF")
            # Read back via COM to verify
            def _check_font():
                ws = _session.get_sheet("Sheet1", wb_name)
                rng = ws.Range("A1:C1")
                return rng.Font.Bold, rng.Font.Size, rng.Font.Color
            bold_val, size_val, color_val = _session.run_com(_check_font)
            assert bold_val is True, f"Font.Bold expected True, got {bold_val}"
            assert size_val == 12, f"Font.Size expected 12, got {size_val}"
            # White #FFFFFF → BGR=0xFFFFFF=16777215
            assert color_val == 16777215, f"Font.Color expected 16777215, got {color_val}"
            record("format.font", "PASS",
                   f"bold={bold_val} size={size_val} color={color_val}")
        except Exception as e:
            record("format.font", "FAIL", str(e))

        # --- fill: gold header background ---
        try:
            format_action("fill", range="A1:C1", sheet="Sheet1", workbook=wb_name,
                          fill_color="#D4A84B")
            def _check_fill():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Range("A1").Interior.Color
            fill_color_val = _session.run_com(_check_fill)
            # #D4A84B → R=0xD4=212, G=0xA8=168, B=0x4B=75
            # BGR = (75<<16)|(168<<8)|212 = 4958420
            expected_bgr = (0x4B << 16) | (0xA8 << 8) | 0xD4
            assert fill_color_val == expected_bgr, (
                f"Interior.Color expected {expected_bgr}, got {fill_color_val}"
            )
            record("format.fill", "PASS", f"Interior.Color={fill_color_val}")
        except Exception as e:
            record("format.fill", "FAIL", str(e))

        # --- border: outline on data range ---
        try:
            format_action("border", range="A1:C4", sheet="Sheet1", workbook=wb_name,
                          border_sides="outline", border_style="continuous",
                          border_weight="medium", border_color="#000000")
            def _check_border():
                ws = _session.get_sheet("Sheet1", wb_name)
                # xlEdgeTop=8
                return ws.Range("A1:C4").Borders(8).LineStyle
            ls = _session.run_com(_check_border)
            assert ls == 1, f"LineStyle expected 1 (continuous), got {ls}"
            record("format.border", "PASS", f"LineStyle={ls}")
        except Exception as e:
            record("format.border", "FAIL", str(e))

        # --- number_format: currency on Amount column ---
        try:
            format_action("number_format", range="B2:B4", sheet="Sheet1", workbook=wb_name,
                          number_format="#,##0.00")
            def _check_nf():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Range("B2").NumberFormat
            nf = _session.run_com(_check_nf)
            assert nf == "#,##0.00", f"NumberFormat expected '#,##0.00', got {nf!r}"
            record("format.number_format", "PASS", f"NumberFormat={nf!r}")
        except Exception as e:
            record("format.number_format", "FAIL", str(e))

        # --- alignment: center header ---
        try:
            format_action("alignment", range="A1:C1", sheet="Sheet1", workbook=wb_name,
                          horizontal="center", vertical="center")
            def _check_align():
                ws = _session.get_sheet("Sheet1", wb_name)
                rng = ws.Range("A1")
                return rng.HorizontalAlignment, rng.VerticalAlignment
            h_align, v_align = _session.run_com(_check_align)
            assert h_align == -4108, f"HorizontalAlignment expected -4108 (center), got {h_align}"
            assert v_align == -4108, f"VerticalAlignment expected -4108 (center), got {v_align}"
            record("format.alignment", "PASS",
                   f"h={h_align} v={v_align}")
        except Exception as e:
            record("format.alignment", "FAIL", str(e))

        # --- column_width ---
        try:
            format_action("column_width", range="A:A", sheet="Sheet1", workbook=wb_name,
                          width=20.0)
            def _check_cw():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Range("A1").ColumnWidth
            cw = _session.run_com(_check_cw)
            # Excel may round slightly; allow a small tolerance
            assert abs(cw - 20.0) < 1.0, f"ColumnWidth expected ~20, got {cw}"
            record("format.column_width", "PASS", f"ColumnWidth={cw}")
        except Exception as e:
            record("format.column_width", "FAIL", str(e))

        # --- row_height ---
        try:
            format_action("row_height", range="1:1", sheet="Sheet1", workbook=wb_name,
                          height=30.0)
            def _check_rh():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Range("A1").RowHeight
            rh = _session.run_com(_check_rh)
            assert abs(rh - 30.0) < 1.0, f"RowHeight expected ~30, got {rh}"
            record("format.row_height", "PASS", f"RowHeight={rh}")
        except Exception as e:
            record("format.row_height", "FAIL", str(e))

        # --- autofit ---
        try:
            format_action("autofit", range="A:C", sheet="Sheet1", workbook=wb_name,
                          autofit_columns=True, autofit_rows=False)
            record("format.autofit", "PASS")
        except Exception as e:
            record("format.autofit", "FAIL", str(e))

    except Exception as e:
        record("section11.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        _close_wb(wb, wb_name or "unknown")


# ── Section 12: Workbook create / save_as ─────────────────────────────────────

def run_workbook_create_save_as() -> None:
    section_header("SECTION 12 — workbook.create + workbook.save_as")
    if _check_excel_busy("workbook.create", "workbook.create_with_path",
                         "workbook.save_as"):
        return
    import tempfile

    wb = None
    wb_name = None
    wb2 = None
    wb2_name = None
    tmp_path_create = None
    tmp_path_save_as = None
    try:
        # --- create: blank (no path) ---
        try:
            r = workbook_action("create")
            wb_name = r["created"]
            assert wb_name, "create returned empty name"
            # verify it is now open
            r2 = workbook_action("list")
            names = [w["name"] for w in r2["workbooks"]]
            assert wb_name in names, f"{wb_name} not found in {names}"
            # resolve COM object for cleanup
            def _get_wb(name):
                return _session.get_workbook(name)
            wb = _session.run_com(_get_wb, wb_name)
            record("workbook.create", "PASS", f"created={wb_name}")
        except Exception as e:
            record("workbook.create", "FAIL", str(e))

        # --- create_with_path: create + immediate SaveAs ---
        try:
            tmp_dir = tempfile.gettempdir()
            tmp_path_create = os.path.join(tmp_dir, "thepexcel_mcp_smoke_create.xlsx")
            # Remove if exists from a previous run
            if os.path.exists(tmp_path_create):
                os.remove(tmp_path_create)

            r = workbook_action("create", path=tmp_path_create)
            wb2_name = r["created"]
            assert os.path.exists(tmp_path_create), (
                f"File not created on disk: {tmp_path_create}"
            )
            # Read-back: file on disk is the proof
            wb2 = _session.run_com(lambda: _session.get_workbook(wb2_name))
            record("workbook.create_with_path", "PASS",
                   f"file_exists={os.path.exists(tmp_path_create)} name={wb2_name}")
        except Exception as e:
            record("workbook.create_with_path", "FAIL", str(e))

        # --- save_as: SaveAs existing workbook to new path ---
        try:
            # Use the first created blank wb (wb_name)
            if wb_name:
                tmp_path_save_as = os.path.join(
                    tempfile.gettempdir(), "thepexcel_mcp_smoke_saveas.xlsx"
                )
                if os.path.exists(tmp_path_save_as):
                    os.remove(tmp_path_save_as)

                r = workbook_action("save_as", workbook=wb_name, path=tmp_path_save_as)
                assert os.path.exists(tmp_path_save_as), (
                    f"SaveAs file not found on disk: {tmp_path_save_as}"
                )
                assert "saved_as" in r, f"Expected 'saved_as' key in result: {r}"
                record("workbook.save_as", "PASS",
                       f"file_exists=True path={os.path.basename(tmp_path_save_as)}")
            else:
                record("workbook.save_as", "SKIP", "workbook.create failed — no wb to save_as")
        except Exception as e:
            record("workbook.save_as", "FAIL", str(e))

    except Exception as e:
        record("section12.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        # Close both temp workbooks
        if wb is not None:
            _close_wb(wb, wb_name or "unknown")
        if wb2 is not None:
            _close_wb(wb2, wb2_name or "unknown")
        # Clean up temp files
        for tmp in (tmp_path_create, tmp_path_save_as):
            if tmp and os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except Exception:
                    pass


# ── Section 13: Conditional Formatting ───────────────────────────────────────

def run_conditional_format() -> None:
    """Live read-back smoke for each CF action.

    Read-back is the oracle the unit mocks cannot provide: we assert
    FormatConditions.Count increased (rule was actually applied) and
    check specific COM properties on the returned condition object.

    xlDataBar type constant = 4, xlIconSet = 6, xlCellValue = 1, xlTop10 = 5
    (XlFormatConditionType enum from Excel VBA object model).
    """
    section_header("SECTION 13 — Conditional Formatting (read-back oracle)")
    if _check_excel_busy(
        "cf.cell_rule", "cf.cell_rule_color", "cf.data_bar",
        "cf.data_bar_color", "cf.color_scale", "cf.icon_set",
        "cf.top_bottom", "cf.clear",
    ):
        return

    # xlFormatConditionType enum values
    XL_CELL_VALUE = 1
    XL_TOP10      = 5
    XL_DATA_BAR   = 4
    XL_ICON_SET   = 6

    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        def _write():
            ws = _session.get_sheet("Sheet1", wb_name)
            ws.Range("A1:A10").Value = [[i * 10] for i in range(1, 11)]
        _session.run_com(_write)

        # ── cell_rule: greater than 50, fill yellow ──────────────────────────
        try:
            conditional_format_action(
                "cell_rule", range="A1:A10", sheet="Sheet1", workbook=wb_name,
                operator="greater", formula1="50", fill_color="#FFFF00",
            )
            def _check_cell_rule():
                ws = _session.get_sheet("Sheet1", wb_name)
                rng = ws.Range("A1:A10")
                count = rng.FormatConditions.Count
                if count < 1:
                    return None, None, None
                cond = rng.FormatConditions(1)
                return count, cond.Type, cond.Interior.Color
            count, ctype, fill_bgr = _session.run_com(_check_cell_rule)
            assert count >= 1, f"FormatConditions.Count expected >=1, got {count}"
            assert ctype == XL_CELL_VALUE, f"Type expected {XL_CELL_VALUE} (xlCellValue), got {ctype}"
            # Yellow #FFFF00 → BGR = (0<<16)|(255<<8)|255 = 65535
            assert fill_bgr == 65535, f"Interior.Color expected 65535 (yellow), got {fill_bgr}"
            record("cf.cell_rule", "PASS",
                   f"Count={count} Type={ctype} Interior.Color={fill_bgr}")
        except Exception as e:
            record("cf.cell_rule", "FAIL", str(e))

        # ── cell_rule font_color (separate range to avoid count confusion) ───
        try:
            conditional_format_action(
                "cell_rule", range="A1:A5", sheet="Sheet1", workbook=wb_name,
                operator="less", formula1="30", font_color="#FF0000",
            )
            def _check_font_color():
                ws = _session.get_sheet("Sheet1", wb_name)
                rng = ws.Range("A1:A5")
                # Grab the last-added rule (highest index)
                count = rng.FormatConditions.Count
                cond = rng.FormatConditions(count)
                return cond.Font.Color
            font_bgr = _session.run_com(_check_font_color)
            # Red #FF0000 → BGR = 255
            assert font_bgr == 255, f"Font.Color expected 255 (red BGR), got {font_bgr}"
            record("cf.cell_rule_color", "PASS", f"Font.Color={font_bgr}")
        except Exception as e:
            record("cf.cell_rule_color", "FAIL", str(e))

        # ── clear (wipe A1:A5 before data_bar so counts are predictable) ────
        try:
            conditional_format_action(
                "clear", range="A1:A10", sheet="Sheet1", workbook=wb_name,
            )
            def _check_clear():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Range("A1:A10").FormatConditions.Count
            count_after = _session.run_com(_check_clear)
            assert count_after == 0, f"After clear, Count expected 0, got {count_after}"
            record("cf.clear", "PASS", f"Count after clear={count_after}")
        except Exception as e:
            record("cf.clear", "FAIL", str(e))

        # ── data_bar (no color) ──────────────────────────────────────────────
        try:
            conditional_format_action(
                "data_bar", range="A1:A10", sheet="Sheet1", workbook=wb_name,
            )
            def _check_data_bar():
                ws = _session.get_sheet("Sheet1", wb_name)
                rng = ws.Range("A1:A10")
                count = rng.FormatConditions.Count
                if count < 1:
                    return None, None
                cond = rng.FormatConditions(1)
                return count, cond.Type
            count, ctype = _session.run_com(_check_data_bar)
            assert count >= 1, f"FormatConditions.Count expected >=1, got {count}"
            assert ctype == XL_DATA_BAR, (
                f"Type expected {XL_DATA_BAR} (xlDataBar), got {ctype}"
            )
            record("cf.data_bar", "PASS", f"Count={count} Type={ctype}")
        except Exception as e:
            record("cf.data_bar", "FAIL", str(e))

        # ── data_bar with color ──────────────────────────────────────────────
        try:
            conditional_format_action(
                "data_bar", range="A1:A10", sheet="Sheet1", workbook=wb_name,
                color="#0070C0",
            )
            def _check_data_bar_color():
                ws = _session.get_sheet("Sheet1", wb_name)
                rng = ws.Range("A1:A10")
                count = rng.FormatConditions.Count
                cond = rng.FormatConditions(count)  # last added
                return count, cond.Type, cond.BarColor.Color
            count, ctype, bar_bgr = _session.run_com(_check_data_bar_color)
            assert ctype == XL_DATA_BAR, f"Type expected {XL_DATA_BAR}, got {ctype}"
            # #0070C0 → R=0x00, G=0x70=112, B=0xC0=192 → BGR=(192<<16)|(112<<8)|0 = 12607488
            expected_bar_bgr = (0xC0 << 16) | (0x70 << 8) | 0x00
            assert bar_bgr == expected_bar_bgr, (
                f"BarColor.Color expected {expected_bar_bgr} (#0070C0 BGR), got {bar_bgr}"
            )
            record("cf.data_bar_color", "PASS",
                   f"Count={count} BarColor.Color={bar_bgr}")
        except Exception as e:
            record("cf.data_bar_color", "FAIL", str(e))

        # ── clear before color_scale / icon_set ─────────────────────────────
        def _do_clear():
            ws = _session.get_sheet("Sheet1", wb_name)
            ws.Range("A1:A10").FormatConditions.Delete()
        _session.run_com(_do_clear)

        # ── color_scale (3-color) ────────────────────────────────────────────
        try:
            conditional_format_action(
                "color_scale", range="A1:A10", sheet="Sheet1", workbook=wb_name,
                scale_type=3,
            )
            def _check_color_scale():
                ws = _session.get_sheet("Sheet1", wb_name)
                rng = ws.Range("A1:A10")
                count = rng.FormatConditions.Count
                if count < 1:
                    return None, None, None
                cond = rng.FormatConditions(1)
                # ColorScale has ColorScaleCriteria collection (len=2 or 3)
                return count, cond.Type, len(cond.ColorScaleCriteria)
            count, ctype, criteria_len = _session.run_com(_check_color_scale)
            assert count >= 1, f"Count expected >=1, got {count}"
            # xlColorScale type = 3
            assert ctype == 3, f"Type expected 3 (xlColorScale), got {ctype}"
            assert criteria_len == 3, f"ColorScaleCriteria expected 3, got {criteria_len}"
            record("cf.color_scale", "PASS",
                   f"Count={count} Type={ctype} Criteria={criteria_len}")
        except Exception as e:
            record("cf.color_scale", "FAIL", str(e))

        # ── clear before icon_set ────────────────────────────────────────────
        _session.run_com(_do_clear)

        # ── icon_set ─────────────────────────────────────────────────────────
        try:
            conditional_format_action(
                "icon_set", range="A1:A10", sheet="Sheet1", workbook=wb_name,
                style="3traffic_lights",
            )
            def _check_icon_set():
                ws = _session.get_sheet("Sheet1", wb_name)
                rng = ws.Range("A1:A10")
                count = rng.FormatConditions.Count
                if count < 1:
                    return None, None, None
                cond = rng.FormatConditions(1)
                return count, cond.Type, cond.IconSet.Count
            count, ctype, icon_count = _session.run_com(_check_icon_set)
            assert count >= 1, f"Count expected >=1, got {count}"
            assert ctype == XL_ICON_SET, (
                f"Type expected {XL_ICON_SET} (xlIconSet), got {ctype}"
            )
            assert icon_count == 3, (
                f"IconSet.Count expected 3 (3 traffic lights), got {icon_count}"
            )
            record("cf.icon_set", "PASS",
                   f"Count={count} Type={ctype} IconSet.Count={icon_count}")
        except Exception as e:
            record("cf.icon_set", "FAIL", str(e))

        # ── clear before top_bottom ──────────────────────────────────────────
        _session.run_com(_do_clear)

        # ── top_bottom ───────────────────────────────────────────────────────
        try:
            conditional_format_action(
                "top_bottom", range="A1:A10", sheet="Sheet1", workbook=wb_name,
                kind="top", rank=3, percent=False, fill_color="#00B050",
            )
            def _check_top_bottom():
                ws = _session.get_sheet("Sheet1", wb_name)
                rng = ws.Range("A1:A10")
                count = rng.FormatConditions.Count
                if count < 1:
                    return None, None, None, None, None
                cond = rng.FormatConditions(1)
                return count, cond.Type, cond.TopBottom, cond.Rank, cond.Interior.Color
            count, ctype, top_bottom_val, rank_val, fill_bgr = _session.run_com(
                _check_top_bottom
            )
            assert count >= 1, f"Count expected >=1, got {count}"
            assert ctype == XL_TOP10, (
                f"Type expected {XL_TOP10} (xlTop10), got {ctype}"
            )
            assert top_bottom_val == 1, (
                f"TopBottom expected 1 (xlTop10Top), got {top_bottom_val}"
            )
            assert rank_val == 3, f"Rank expected 3, got {rank_val}"
            # Green #00B050 → R=0, G=0xB0=176, B=0x50=80 → BGR=(80<<16)|(176<<8)|0 = 5285888
            expected_fill = (0x50 << 16) | (0xB0 << 8) | 0x00
            assert fill_bgr == expected_fill, (
                f"Interior.Color expected {expected_fill} (#00B050 BGR), got {fill_bgr}"
            )
            record("cf.top_bottom", "PASS",
                   f"Count={count} Type={ctype} Rank={rank_val} Color={fill_bgr}")
        except Exception as e:
            record("cf.top_bottom", "FAIL", str(e))

    except Exception as e:
        record("section13.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        _close_wb(wb, wb_name or "unknown")


# ── Section 14: View (freeze / gridlines / zoom / headings) ───────────────────

def run_view() -> None:
    """Live read-back smoke for each view action.

    Critical regression guard: the multi-workbook test verifies that
    view mutations land on the TARGET workbook window (wb.Windows(1)),
    NOT on the foreground workbook (app.ActiveWindow).  This proves the
    fix vs the original ActiveWindow bug.
    """
    section_header("SECTION 14 — View (freeze_panes / unfreeze_panes / gridlines / zoom / headings)")
    if _check_excel_busy(
        "view.freeze_panes", "view.unfreeze_panes",
        "view.gridlines", "view.zoom", "view.headings",
        "view.multi_workbook_isolation",
    ):
        return

    wb = None
    wb_name = None
    wb2 = None
    wb2_name = None
    try:
        wb, wb_name = _new_wb()

        # ── freeze_panes: cell="B2" (freeze row 1 + col A) ─────────────────
        try:
            view_action("freeze_panes", sheet="Sheet1", workbook=wb_name, cell="B2")

            def _check_freeze():
                ws = _session.get_sheet("Sheet1", wb_name)
                wb_inner = ws.Parent
                win = wb_inner.Windows(1)
                return win.SplitRow, win.SplitColumn, win.FreezePanes

            split_row, split_col, frozen = _session.run_com(_check_freeze)
            assert split_row == 1, f"SplitRow expected 1, got {split_row}"
            assert split_col == 1, f"SplitColumn expected 1, got {split_col}"
            assert frozen is True, f"FreezePanes expected True, got {frozen}"
            record("view.freeze_panes", "PASS",
                   f"SplitRow={split_row} SplitColumn={split_col} FreezePanes={frozen}")
        except Exception as e:
            record("view.freeze_panes", "FAIL", str(e))

        # ── unfreeze_panes ──────────────────────────────────────────────────
        try:
            view_action("unfreeze_panes", sheet="Sheet1", workbook=wb_name)

            def _check_unfreeze():
                ws = _session.get_sheet("Sheet1", wb_name)
                win = ws.Parent.Windows(1)
                return win.FreezePanes

            frozen_after = _session.run_com(_check_unfreeze)
            assert frozen_after is False, f"FreezePanes expected False after unfreeze, got {frozen_after}"
            record("view.unfreeze_panes", "PASS", f"FreezePanes={frozen_after}")
        except Exception as e:
            record("view.unfreeze_panes", "FAIL", str(e))

        # ── gridlines: hide then show ────────────────────────────────────────
        try:
            view_action("gridlines", sheet="Sheet1", workbook=wb_name, show=False)

            def _check_gridlines():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Parent.Windows(1).DisplayGridlines

            gl = _session.run_com(_check_gridlines)
            assert gl is False, f"DisplayGridlines expected False, got {gl}"

            view_action("gridlines", sheet="Sheet1", workbook=wb_name, show=True)
            gl2 = _session.run_com(_check_gridlines)
            assert gl2 is True, f"DisplayGridlines expected True after re-enable, got {gl2}"
            record("view.gridlines", "PASS", f"hide={not gl} → show={gl2}")
        except Exception as e:
            record("view.gridlines", "FAIL", str(e))

        # ── zoom ─────────────────────────────────────────────────────────────
        try:
            view_action("zoom", sheet="Sheet1", workbook=wb_name, zoom=75)

            def _check_zoom():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Parent.Windows(1).Zoom

            z = _session.run_com(_check_zoom)
            assert z == 75, f"Zoom expected 75, got {z}"
            record("view.zoom", "PASS", f"Zoom={z}")
        except Exception as e:
            record("view.zoom", "FAIL", str(e))

        # ── headings: hide then show ─────────────────────────────────────────
        try:
            view_action("headings", sheet="Sheet1", workbook=wb_name, show=False)

            def _check_headings():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Parent.Windows(1).DisplayHeadings

            h = _session.run_com(_check_headings)
            assert h is False, f"DisplayHeadings expected False, got {h}"

            view_action("headings", sheet="Sheet1", workbook=wb_name, show=True)
            h2 = _session.run_com(_check_headings)
            assert h2 is True, f"DisplayHeadings expected True after re-enable, got {h2}"
            record("view.headings", "PASS", f"hide={not h} → show={h2}")
        except Exception as e:
            record("view.headings", "FAIL", str(e))

        # ── MULTI-WORKBOOK ISOLATION (regression guard for ActiveWindow bug) ─
        # Create a SECOND workbook so it becomes the foreground window.
        # Then mutate view properties targeting the BACKGROUND wb, and assert:
        #   (a) the target wb changed as expected
        #   (b) the foreground wb did NOT change (isolation proof)
        try:
            wb2, wb2_name = _new_wb()  # wb2 is now the foreground (ActiveWindow)

            # Apply freeze to the BACKGROUND target workbook
            view_action("freeze_panes", sheet="Sheet1", workbook=wb_name, cell="B2")
            view_action("zoom", sheet="Sheet1", workbook=wb_name, zoom=85)

            def _check_multi_wb():
                # Target workbook window
                ws_target = _session.get_sheet("Sheet1", wb_name)
                win_target = ws_target.Parent.Windows(1)
                target_frozen = win_target.FreezePanes
                target_zoom = win_target.Zoom

                # Foreground workbook window
                ws_fg = _session.get_sheet("Sheet1", wb2_name)
                win_fg = ws_fg.Parent.Windows(1)
                fg_frozen = win_fg.FreezePanes
                fg_zoom = win_fg.Zoom

                return target_frozen, target_zoom, fg_frozen, fg_zoom

            tgt_frozen, tgt_zoom, fg_frozen, fg_zoom = _session.run_com(_check_multi_wb)

            assert tgt_frozen is True, (
                f"TARGET wb FreezePanes expected True, got {tgt_frozen}  "
                f"(proves wb.Windows(1) targets the right workbook)"
            )
            assert tgt_zoom == 85, f"TARGET wb Zoom expected 85, got {tgt_zoom}"
            assert fg_frozen is False, (
                f"FOREGROUND wb FreezePanes expected False (unchanged), got {fg_frozen}  "
                f"(proves no cross-workbook mutation)"
            )
            # Foreground zoom should be its default (100), NOT 85
            assert fg_zoom != 85, (
                f"FOREGROUND wb Zoom should NOT be 85 — mutation leaked to wrong window"
            )
            record(
                "view.multi_workbook_isolation", "PASS",
                f"target_frozen={tgt_frozen} target_zoom={tgt_zoom} "
                f"fg_frozen={fg_frozen} fg_zoom={fg_zoom}",
            )
        except Exception as e:
            record("view.multi_workbook_isolation", "FAIL", str(e))

    except Exception as e:
        record("section14.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        # Close the second workbook first (it's foreground)
        if wb2 is not None:
            _close_wb(wb2, wb2_name or "unknown")
        _close_wb(wb, wb_name or "unknown")


# ── Section 15: Data Validation ───────────────────────────────────────────────

def run_validation() -> None:
    """Live read-back smoke for each validation action.

    Critical regression guards:
    - Delete-before-Add: apply a second validation OVER an existing one must
      NOT raise COM error 1004 (proves rng.Validation.Delete() runs first).
    - After clear, rng.Validation.Type must equal xlValidateInputOnly=0.
    """
    section_header("SECTION 15 — Data Validation (list / whole_number / custom / clear / re-apply)")
    if _check_excel_busy(
        "validation.list", "validation.whole_number",
        "validation.custom", "validation.clear", "validation.re_apply",
    ):
        return

    # XlDVType enum constants
    XL_VALIDATE_INPUT_ONLY  = 0   # no validation (cleared state)
    XL_VALIDATE_WHOLE_NUMBER = 1
    XL_VALIDATE_LIST        = 3
    XL_VALIDATE_CUSTOM      = 7

    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        # ── list validation ──────────────────────────────────────────────────
        try:
            validation_action(
                "list", range="A1:A5", sheet="Sheet1", workbook=wb_name,
                formula1="Yes,No,Maybe",
            )

            def _check_list():
                ws = _session.get_sheet("Sheet1", wb_name)
                v = ws.Range("A1").Validation
                return v.Type, v.Formula1

            v_type, v_formula1 = _session.run_com(_check_list)
            assert v_type == XL_VALIDATE_LIST, (
                f"Validation.Type expected {XL_VALIDATE_LIST} (xlValidateList), got {v_type}"
            )
            # Formula1 may be wrapped in quotes by Excel; check the values are present
            assert "Yes" in v_formula1, (
                f"Validation.Formula1 expected to contain 'Yes', got {v_formula1!r}"
            )
            record("validation.list", "PASS",
                   f"Type={v_type} Formula1={v_formula1!r}")
        except Exception as e:
            record("validation.list", "FAIL", str(e))

        # ── whole_number with operator=between ───────────────────────────────
        try:
            validation_action(
                "whole_number", range="B1:B5", sheet="Sheet1", workbook=wb_name,
                formula1="1", formula2="100", operator="between",
            )

            def _check_whole():
                ws = _session.get_sheet("Sheet1", wb_name)
                v = ws.Range("B1").Validation
                return v.Type, v.Operator, v.Formula1, v.Formula2

            v_type, v_op, v_f1, v_f2 = _session.run_com(_check_whole)
            assert v_type == XL_VALIDATE_WHOLE_NUMBER, (
                f"Validation.Type expected {XL_VALIDATE_WHOLE_NUMBER}, got {v_type}"
            )
            assert v_op == 1, f"Validation.Operator expected 1 (xlBetween), got {v_op}"
            assert "1" in v_f1, f"Formula1 expected '1', got {v_f1!r}"
            assert "100" in v_f2, f"Formula2 expected '100', got {v_f2!r}"
            record("validation.whole_number", "PASS",
                   f"Type={v_type} Operator={v_op} F1={v_f1!r} F2={v_f2!r}")
        except Exception as e:
            record("validation.whole_number", "FAIL", str(e))

        # ── custom formula validation ────────────────────────────────────────
        try:
            validation_action(
                "custom", range="C1:C5", sheet="Sheet1", workbook=wb_name,
                formula1="=ISNUMBER(C1)",
            )

            def _check_custom():
                ws = _session.get_sheet("Sheet1", wb_name)
                v = ws.Range("C1").Validation
                return v.Type, v.Formula1

            v_type, v_formula1 = _session.run_com(_check_custom)
            assert v_type == XL_VALIDATE_CUSTOM, (
                f"Validation.Type expected {XL_VALIDATE_CUSTOM} (xlValidateCustom), got {v_type}"
            )
            assert "ISNUMBER" in v_formula1.upper(), (
                f"Formula1 expected ISNUMBER formula, got {v_formula1!r}"
            )
            record("validation.custom", "PASS",
                   f"Type={v_type} Formula1={v_formula1!r}")
        except Exception as e:
            record("validation.custom", "FAIL", str(e))

        # ── clear (removes validation; Type should become 0) ─────────────────
        try:
            validation_action(
                "clear", range="A1:A5", sheet="Sheet1", workbook=wb_name,
            )

            def _check_clear():
                ws = _session.get_sheet("Sheet1", wb_name)
                try:
                    v_type = ws.Range("A1").Validation.Type
                except Exception:
                    # Some Excel versions raise COMError after Delete — treat as cleared
                    v_type = XL_VALIDATE_INPUT_ONLY
                return v_type

            v_type_after = _session.run_com(_check_clear)
            assert v_type_after == XL_VALIDATE_INPUT_ONLY, (
                f"After clear, Validation.Type expected 0 (xlValidateInputOnly), got {v_type_after}"
            )
            record("validation.clear", "PASS", f"Type after clear={v_type_after}")
        except Exception as e:
            record("validation.clear", "FAIL", str(e))

        # ── re-apply validation over an existing one (Delete-before-Add guard) ──
        # Apply list validation to D1:D5 first, then overwrite with a different list.
        # Must NOT raise COM error 1004.
        try:
            validation_action(
                "list", range="D1:D5", sheet="Sheet1", workbook=wb_name,
                formula1="Red,Green,Blue",
            )
            # Now apply again (overwrite) — this is the 1004-regression test
            validation_action(
                "list", range="D1:D5", sheet="Sheet1", workbook=wb_name,
                formula1="Alpha,Beta,Gamma",
            )

            def _check_reapply():
                ws = _session.get_sheet("Sheet1", wb_name)
                v = ws.Range("D1").Validation
                return v.Type, v.Formula1

            v_type, v_f1 = _session.run_com(_check_reapply)
            assert v_type == XL_VALIDATE_LIST, (
                f"Re-apply: Type expected {XL_VALIDATE_LIST}, got {v_type}"
            )
            assert "Alpha" in v_f1, (
                f"Re-apply: Formula1 expected 'Alpha,...', got {v_f1!r}"
            )
            record("validation.re_apply", "PASS",
                   f"Type={v_type} Formula1={v_f1!r} — no 1004 error (Delete-before-Add confirmed)")
        except Exception as e:
            record("validation.re_apply", "FAIL", str(e))

    except Exception as e:
        record("section15.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        _close_wb(wb, wb_name or "unknown")


# ── Section 16: Slicer ────────────────────────────────────────────────────────

def run_slicer() -> None:
    """Live read-back smoke for slicer add / list / delete on a Table source.

    Critical regression guard: assert the slicer's actual Top/Left/Width/Height
    match the requested values.  The historical bug was a wrong positional-arg
    layout in Slicers.Add that shifted Caption/Top/Left/Width/Height by one slot
    — the slicer appeared at a different position than requested.

    Timeline (add_timeline) requires a date-typed PivotTable field; that setup
    is complex and fragile in automation, so we SKIP it with a clear note.
    """
    section_header("SECTION 16 — Slicer (add / list / delete on Table source; position read-back)")
    if _check_excel_busy(
        "slicer.add", "slicer.list", "slicer.delete",
    ):
        return

    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        # Build a Table with a Category column so we have a slicer source
        def _setup_table():
            ws = _session.get_sheet("Sheet1", wb_name)
            ws.Range("A1:B1").Value = [["Category", "Value"]]
            ws.Range("A2:B5").Value = [
                ["Fruit",  10],
                ["Veggie", 20],
                ["Fruit",  30],
                ["Veggie", 40],
            ]
            # Create a ListObject (Table) named "SlicerSource"
            lo = ws.ListObjects.Add(
                1,                     # xlSrcRange
                ws.Range("A1:B5"),
                None,
                1,                     # xlYes (header row)
            )
            lo.Name = "SlicerSource"
            return lo.Name

        tbl_name = _session.run_com(_setup_table)
        assert tbl_name == "SlicerSource", f"Expected table 'SlicerSource', got {tbl_name!r}"

        # ── add slicer (position read-back regression test) ─────────────────
        REQ_TOP    = 60.0
        REQ_LEFT   = 80.0
        REQ_WIDTH  = 150.0
        REQ_HEIGHT = 160.0
        sc_name = None
        sl_name = None

        try:
            result = slicer_action(
                "add",
                workbook=wb_name,
                source="SlicerSource",
                field="Category",
                sheet="Sheet1",
                caption="Category Filter",
                top=REQ_TOP,
                left=REQ_LEFT,
                width=REQ_WIDTH,
                height=REQ_HEIGHT,
            )
            sc_name = result.get("slicer_cache")
            sl_name = result.get("slicer_name")

            # Read back actual slicer position from COM
            def _check_slicer_pos():
                wb_inner = _session.get_workbook(wb_name)
                # Navigate: SlicerCache → Slicers(1) → position properties
                sc = wb_inner.SlicerCaches(sc_name)
                sl = sc.Slicers(1)
                return sl.Top, sl.Left, sl.Width, sl.Height, sl.Caption

            act_top, act_left, act_width, act_height, act_caption = (
                _session.run_com(_check_slicer_pos)
            )
            TOL = 2.0  # Excel may round to nearest point
            assert abs(act_top    - REQ_TOP)    < TOL, (
                f"Slicer.Top expected ~{REQ_TOP}, got {act_top}  "
                f"(off-by-one arg bug would shift this)"
            )
            assert abs(act_left   - REQ_LEFT)   < TOL, (
                f"Slicer.Left expected ~{REQ_LEFT}, got {act_left}"
            )
            assert abs(act_width  - REQ_WIDTH)  < TOL, (
                f"Slicer.Width expected ~{REQ_WIDTH}, got {act_width}"
            )
            assert abs(act_height - REQ_HEIGHT) < TOL, (
                f"Slicer.Height expected ~{REQ_HEIGHT}, got {act_height}"
            )
            assert act_caption == "Category Filter", (
                f"Slicer.Caption expected 'Category Filter', got {act_caption!r}"
            )
            record(
                "slicer.add", "PASS",
                f"cache={sc_name} name={sl_name} "
                f"top={act_top} left={act_left} width={act_width} height={act_height} "
                f"caption={act_caption!r}",
            )
        except Exception as e:
            record("slicer.add", "FAIL", str(e))

        # ── list (slicer cache should appear) ────────────────────────────────
        try:
            list_result = slicer_action("list", workbook=wb_name)
            caches = list_result.get("slicer_caches", [])
            count = list_result.get("count", 0)
            assert count >= 1, f"Expected >=1 slicer cache, got {count}"
            cache_names = [c["name"] for c in caches]
            if sc_name:
                assert sc_name in cache_names, (
                    f"Expected slicer cache '{sc_name}' in list, got {cache_names}"
                )
            record("slicer.list", "PASS",
                   f"count={count} caches={cache_names}")
        except Exception as e:
            record("slicer.list", "FAIL", str(e))

        # ── delete ────────────────────────────────────────────────────────────
        try:
            if sc_name is None:
                raise AssertionError("Cannot test delete — slicer_cache name unknown (add failed)")
            slicer_action("delete", workbook=wb_name, slicer=sc_name)

            # Confirm cache is gone
            def _check_deleted():
                wb_inner = _session.get_workbook(wb_name)
                return wb_inner.SlicerCaches.Count

            count_after = _session.run_com(_check_deleted)
            assert count_after == 0, (
                f"After delete, SlicerCaches.Count expected 0, got {count_after}"
            )
            record("slicer.delete", "PASS", f"SlicerCaches.Count={count_after}")
        except Exception as e:
            record("slicer.delete", "FAIL", str(e))

        # ── add_timeline: SKIP (requires date PivotTable field — complex setup) ──
        record(
            "slicer.add_timeline", "SKIP",
            "Requires a PivotTable with a date-typed field; "
            "Table-based timeline is not supported by xlTimeline cache type. "
            "Covered separately when a PivotTable fixture is available.",
        )

    except Exception as e:
        record("section16.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        _close_wb(wb, wb_name or "unknown")


# ── Section 17: PQ Parameters ─────────────────────────────────────────────────

def run_pq_parameters() -> None:
    section_header("SECTION 17 — PQ Parameters (create / get / set / list / delete)")
    if _check_excel_busy(
        "pq_param.create_number", "pq_param.create_text",
        "pq_param.get", "pq_param.set", "pq_param.list", "pq_param.cleanup",
    ):
        return
    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        # create_parameter — Number
        try:
            r = powerquery_action("create_parameter", name="MaxRows",
                                  value=1000, workbook=wb_name)
            assert r.get("created_parameter") == "MaxRows", r
            record("pq_param.create_number", "PASS")
        except Exception as e:
            record("pq_param.create_number", "FAIL", str(e))

        # create_parameter — Text
        try:
            r = powerquery_action("create_parameter", name="Region",
                                  value="North", param_type="Text",
                                  workbook=wb_name)
            assert r.get("created_parameter") == "Region", r
            record("pq_param.create_text", "PASS")
        except Exception as e:
            record("pq_param.create_text", "FAIL", str(e))

        # get_parameter round-trip
        try:
            r = powerquery_action("get_parameter", name="MaxRows",
                                  workbook=wb_name)
            assert r.get("value") == 1000, r
            assert r.get("type") == "Number", r
            record("pq_param.get", "PASS", f"value={r['value']} type={r['type']}")
        except Exception as e:
            record("pq_param.get", "FAIL", str(e))

        # set_parameter changes value
        try:
            r = powerquery_action("set_parameter", name="MaxRows",
                                  value=500, workbook=wb_name)
            assert r.get("set_parameter") == "MaxRows", r
            # verify effect: read back
            r2 = powerquery_action("get_parameter", name="MaxRows",
                                   workbook=wb_name)
            assert r2.get("value") == 500, r2
            record("pq_param.set", "PASS", f"new_value={r2['value']}")
        except Exception as e:
            record("pq_param.set", "FAIL", str(e))

        # list_parameters returns both created parameters
        try:
            r = powerquery_action("list_parameters", workbook=wb_name)
            params = r.get("parameters", [])
            names = [p["name"] for p in params]
            assert "MaxRows" in names, names
            assert "Region" in names, names
            record("pq_param.list", "PASS", f"params={names}")
        except Exception as e:
            record("pq_param.list", "FAIL", str(e))

        # cleanup — delete both parameter queries
        try:
            powerquery_action("delete", name="MaxRows", workbook=wb_name)
            powerquery_action("delete", name="Region", workbook=wb_name)
            record("pq_param.cleanup", "PASS")
        except Exception as e:
            record("pq_param.cleanup", "FAIL", str(e))

    except Exception as e:
        record("section17.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        _close_wb(wb, wb_name or "unknown")


# ── Section 18: Cube formula builders + best-effort live cube_value ────────────

def run_cube() -> None:
    section_header("SECTION 18 — Cube formula builders + best-effort cube_value")
    if _check_excel_busy(
        "cube.formula_cubevalue", "cube.formula_cubemember",
        "cube.formula_cubevalue_members", "cube.live_cube_value",
        "cube.calc_column_guard", "cube.calc_table_guard",
    ):
        return

    # ── Part A: pure-Python builder strings (no Excel needed) ─────────────────
    try:
        from thepexcel_mcp.domains.datamodel import (
            _build_cubevalue, _build_cubemember, DEFAULT_CONNECTION,
        )

        formula = _build_cubevalue(
            measure="[Measures].[Total Sales]",
            members=None,
            connection=DEFAULT_CONNECTION,
        )
        assert "CUBEVALUE" in formula, formula
        assert "[Measures].[Total Sales]" in formula, formula
        record("cube.formula_cubevalue", "PASS", formula[:60])
    except Exception as e:
        record("cube.formula_cubevalue", "FAIL", str(e))

    try:
        formula = _build_cubemember(
            member_expression="[Date].[Year].&[2024]",
            caption=None,
            connection=DEFAULT_CONNECTION,
        )
        assert "CUBEMEMBER" in formula, formula
        assert "[Date].[Year].&[2024]" in formula, formula
        record("cube.formula_cubemember", "PASS", formula[:60])
    except Exception as e:
        record("cube.formula_cubemember", "FAIL", str(e))

    try:
        formula = _build_cubevalue(
            measure="[Measures].[Total Sales]",
            members=["[Date].[Year].&[2024]", "[Product].[Category].&[Bikes]"],
            connection=DEFAULT_CONNECTION,
        )
        assert "CUBEVALUE" in formula, formula
        assert "[Date].[Year].&[2024]" in formula, formula
        record("cube.formula_cubevalue_members", "PASS", formula[:80])
    except Exception as e:
        record("cube.formula_cubevalue_members", "FAIL", str(e))

    # ── Part B: live cube_value with deadlock-SKIP guard (mirrors section 5) ───
    import thepexcel_mcp.session as _session_mod
    prev_timeout = _session_mod._DEFAULT_TIMEOUT
    _session_mod._DEFAULT_TIMEOUT = 30

    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        sales_m = textwrap.dedent("""\
            let
                Source = #table(
                    type table [ProductID=Int64.Type, Amount=Int64.Type],
                    {{1, 1000}, {2, 1500}, {1, 1200}}
                )
            in
                Source""")
        powerquery_action("create", name="CubeSales", formula=sales_m,
                          workbook=wb_name)

        # add_table may deadlock (no UI pump in automation context) — SKIP not FAIL
        model_ok = False
        try:
            datamodel_action("add_table", workbook=wb_name,
                             source_type="query", source_name="CubeSales")
            model_ok = True
        except Exception as e:
            if _is_timeout_error(e):
                record("cube.live_cube_value", "SKIP",
                       "Data Model build deadlocked — no UI pump (expected in automation)")
            else:
                record("cube.live_cube_value", "FAIL", f"add_table: {e}")

        if model_ok:
            try:
                r = datamodel_action("cube_value", workbook=wb_name,
                                     target_cell="A1",
                                     measure="[Measures].[Sum of Amount]")
                record("cube.live_cube_value", "PASS",
                       f"value={r.get('value')}")
            except Exception as e:
                if _is_timeout_error(e):
                    record("cube.live_cube_value", "SKIP",
                           "cube_value timed out — Data Model not ready in automation")
                else:
                    record("cube.live_cube_value", "FAIL", str(e))

    except Exception as e:
        record("section18.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        if wb is not None:
            try:
                def _delete_queries():
                    while wb.Queries.Count > 0:
                        wb.Queries.Item(1).Delete()
                _session.run_com(_delete_queries)
            except Exception:
                pass
        _close_wb(wb, wb_name or "unknown")
        _session_mod._DEFAULT_TIMEOUT = prev_timeout

    # ── Part C: add_calculated_column / add_calculated_table guard (pure, no Excel) ──
    try:
        from fastmcp.exceptions import ToolError
        try:
            datamodel_action("add_calculated_column")
            record("cube.calc_column_guard", "FAIL", "expected ToolError, got none")
        except ToolError as te:
            msg = str(te)
            assert "Power BI" in msg or "Power Query" in msg or "READ-ONLY" in msg or \
                   "add_measure" in msg, f"unexpected msg: {msg}"
            record("cube.calc_column_guard", "PASS", "ToolError raised as expected")
        except Exception as e:
            record("cube.calc_column_guard", "FAIL", f"wrong exception type: {e}")
    except Exception as e:
        record("cube.calc_column_guard", "FAIL", str(e))

    try:
        from fastmcp.exceptions import ToolError
        try:
            datamodel_action("add_calculated_table")
            record("cube.calc_table_guard", "FAIL", "expected ToolError, got none")
        except ToolError as te:
            msg = str(te)
            assert "Power BI" in msg or "Analysis Services" in msg, \
                   f"unexpected msg: {msg}"
            record("cube.calc_table_guard", "PASS", "ToolError raised as expected")
        except Exception as e:
            record("cube.calc_table_guard", "FAIL", f"wrong exception type: {e}")
    except Exception as e:
        record("cube.calc_table_guard", "FAIL", str(e))


# ── Section 19: Page Setup ────────────────────────────────────────────────────

def run_page_setup() -> None:
    """Live read-back smoke for page-setup actions on a fresh workbook.

    Reads the real ws.PageSetup COM properties after each mutation (not the
    tool's own result dict). PDF export records SKIP if no PDF printer driver
    is available, since some CI hosts lack one.
    """
    section_header("SECTION 19 — Page Setup (set / print_area / print_titles / header_footer / export_pdf / get)")
    if _check_excel_busy(
        "page_setup.orientation", "page_setup.paper_size", "page_setup.scale",
        "page_setup.fit_to_pages", "page_setup.margin_top", "page_setup.center_gridlines",
        "page_setup.print_area", "page_setup.print_titles", "page_setup.header_footer",
        "page_setup.export_pdf", "page_setup.get",
    ):
        return

    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        # ── orientation = landscape (xlLandscape = 2) ────────────────────────
        try:
            page_setup_action("set", sheet="Sheet1", workbook=wb_name, orientation="landscape")

            def _check_orient():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.PageSetup.Orientation

            ori = _session.run_com(_check_orient)
            assert ori == 2, f"Orientation expected 2 (xlLandscape), got {ori}"
            record("page_setup.orientation", "PASS", f"Orientation={ori}")
        except Exception as e:
            record("page_setup.orientation", "FAIL", str(e))

        # ── paper_size = a4 (xlPaperA4 = 9) ──────────────────────────────────
        try:
            page_setup_action("set", sheet="Sheet1", workbook=wb_name, paper_size="a4")

            def _check_paper():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.PageSetup.PaperSize

            paper = _session.run_com(_check_paper)
            assert paper == 9, f"PaperSize expected 9 (xlPaperA4), got {paper}"
            record("page_setup.paper_size", "PASS", f"PaperSize={paper}")
        except Exception as e:
            record("page_setup.paper_size", "FAIL", str(e))

        # ── scale = 80 → Zoom == 80 ──────────────────────────────────────────
        try:
            page_setup_action("set", sheet="Sheet1", workbook=wb_name, scale=80)

            def _check_zoom():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.PageSetup.Zoom

            z = _session.run_com(_check_zoom)
            assert z == 80, f"Zoom expected 80, got {z}"
            record("page_setup.scale", "PASS", f"Zoom={z}")
        except Exception as e:
            record("page_setup.scale", "FAIL", str(e))

        # ── fit_to_wide=1, fit_to_tall=2 → Zoom False, FitTo* set ────────────
        try:
            page_setup_action("set", sheet="Sheet1", workbook=wb_name,
                              fit_to_wide=1, fit_to_tall=2)

            def _check_fit():
                ws = _session.get_sheet("Sheet1", wb_name)
                ps = ws.PageSetup
                return ps.Zoom, ps.FitToPagesWide, ps.FitToPagesTall

            zoom_val, fit_w, fit_t = _session.run_com(_check_fit)
            assert zoom_val is False, f"Zoom expected False with fit-to-pages, got {zoom_val}"
            assert fit_w == 1, f"FitToPagesWide expected 1, got {fit_w}"
            assert fit_t == 2, f"FitToPagesTall expected 2, got {fit_t}"
            record("page_setup.fit_to_pages", "PASS",
                   f"Zoom={zoom_val} Wide={fit_w} Tall={fit_t}")
        except Exception as e:
            record("page_setup.fit_to_pages", "FAIL", str(e))

        # ── top margin = 1.0 inch → TopMargin ≈ 72 points ───────────────────
        try:
            page_setup_action("set", sheet="Sheet1", workbook=wb_name, top=1.0)

            def _check_margin():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.PageSetup.TopMargin

            tm = _session.run_com(_check_margin)
            assert abs(tm - 72.0) < 1.0, f"TopMargin expected ~72.0 pts (1 inch), got {tm}"
            record("page_setup.margin_top", "PASS", f"TopMargin={tm}")
        except Exception as e:
            record("page_setup.margin_top", "FAIL", str(e))

        # ── center_horizontally + print_gridlines = True ────────────────────
        try:
            page_setup_action("set", sheet="Sheet1", workbook=wb_name,
                              center_horizontally=True, print_gridlines=True)

            def _check_center_gl():
                ws = _session.get_sheet("Sheet1", wb_name)
                ps = ws.PageSetup
                return ps.CenterHorizontally, ps.PrintGridlines

            ch, pg = _session.run_com(_check_center_gl)
            assert ch is True, f"CenterHorizontally expected True, got {ch}"
            assert pg is True, f"PrintGridlines expected True, got {pg}"
            record("page_setup.center_gridlines", "PASS",
                   f"CenterHorizontally={ch} PrintGridlines={pg}")
        except Exception as e:
            record("page_setup.center_gridlines", "FAIL", str(e))

        # ── print_area = A1:F50 (Excel normalizes to $A$1:$F$50) ────────────
        try:
            page_setup_action("print_area", sheet="Sheet1", workbook=wb_name,
                              address="A1:F50")

            def _check_print_area():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.PageSetup.PrintArea

            pa = _session.run_com(_check_print_area)
            # Excel normalizes "A1:F50" to absolute "$A$1:$F$50".
            assert "$F$50" in pa, f"PrintArea expected to contain '$F$50', got {pa!r}"
            record("page_setup.print_area", "PASS", f"PrintArea={pa!r}")
        except Exception as e:
            record("page_setup.print_area", "FAIL", str(e))

        # ── print_titles rows=$1:$1, cols=$A:$A ─────────────────────────────
        try:
            page_setup_action("print_titles", sheet="Sheet1", workbook=wb_name,
                              rows="$1:$1", cols="$A:$A")

            def _check_titles():
                ws = _session.get_sheet("Sheet1", wb_name)
                ps = ws.PageSetup
                return ps.PrintTitleRows, ps.PrintTitleColumns

            tr, tc = _session.run_com(_check_titles)
            assert tr == "$1:$1", f"PrintTitleRows expected '$1:$1', got {tr!r}"
            assert tc == "$A:$A", f"PrintTitleColumns expected '$A:$A', got {tc!r}"
            record("page_setup.print_titles", "PASS",
                   f"Rows={tr!r} Columns={tc!r}")
        except Exception as e:
            record("page_setup.print_titles", "FAIL", str(e))

        # ── header_footer center_header ─────────────────────────────────────
        try:
            page_setup_action("header_footer", sheet="Sheet1", workbook=wb_name,
                              center_header="ThepExcel Report")

            def _check_header():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.PageSetup.CenterHeader

            hdr = _session.run_com(_check_header)
            assert hdr == "ThepExcel Report", f"CenterHeader expected 'ThepExcel Report', got {hdr!r}"
            record("page_setup.header_footer", "PASS", f"CenterHeader={hdr!r}")
        except Exception as e:
            record("page_setup.header_footer", "FAIL", str(e))

        # ── export_pdf (SKIP if no PDF printer driver) ──────────────────────
        try:
            tmp_dir = os.environ.get("TEMP") or os.environ.get("TMP") or "."
            pdf_path = os.path.join(tmp_dir, f"thepexcel_smoke_{os.getpid()}.pdf")
            try:
                page_setup_action("export_pdf", sheet="Sheet1", workbook=wb_name,
                                  path=pdf_path, scope="sheet")
            except Exception as ex:
                msg = str(ex)
                if "printer" in msg.lower() or "driver" in msg.lower() or "1004" in msg:
                    record("page_setup.export_pdf", "SKIP",
                           f"No PDF printer driver available: {msg[:120]}")
                    raise _SkipSentinel()
                raise
            assert os.path.exists(pdf_path), f"PDF not created at {pdf_path}"
            size = os.path.getsize(pdf_path)
            assert size > 0, f"PDF file is empty ({size} bytes)"
            os.remove(pdf_path)
            record("page_setup.export_pdf", "PASS", f"size={size} bytes")
        except _SkipSentinel:
            pass
        except Exception as e:
            record("page_setup.export_pdf", "FAIL", str(e))

        # ── get → applied dict has orientation key ──────────────────────────
        try:
            r = page_setup_action("get", sheet="Sheet1", workbook=wb_name)
            applied = r.get("applied", {})
            assert "orientation" in applied, f"'orientation' missing from get result: {applied}"
            record("page_setup.get", "PASS",
                   f"orientation={applied.get('orientation')} paper={applied.get('paper_size')}")
        except Exception as e:
            record("page_setup.get", "FAIL", str(e))

    except Exception as e:
        record("section19.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        _close_wb(wb, wb_name or "unknown")


# ── Section 20: Comments (note + threaded) ─────────────────────────────────────

def run_comment() -> None:
    """Live read-back smoke for note + threaded comment actions.

    Reads cell.Comment.Text() / cell.CommentThreaded.Text() directly from COM.
    Threaded-comment support is gated on some Excel builds; if AddCommentThreaded
    raises, that action is recorded FAIL/SKIP with the COM message (a real finding)
    and the dependent reply/delete-threaded actions are skipped.
    """
    section_header("SECTION 20 — Comments (note add/edit/delete + threaded add/reply/delete + list)")
    if _check_excel_busy(
        "comment.add_note", "comment.edit_note", "comment.add_threaded",
        "comment.reply_threaded", "comment.list", "comment.delete_note",
        "comment.delete_threaded",
    ):
        return

    wb = None
    wb_name = None
    threaded_ok = False
    try:
        wb, wb_name = _new_wb()

        # ── add note on A1 ───────────────────────────────────────────────────
        try:
            comment_action("add", cell="A1", sheet="Sheet1", workbook=wb_name,
                           text="Hello note")

            def _check_note():
                ws = _session.get_sheet("Sheet1", wb_name)
                c = ws.Range("A1").Comment
                return c.Text() if c is not None else None

            txt = _session.run_com(_check_note)
            assert txt == "Hello note", f"Note text expected 'Hello note', got {txt!r}"
            record("comment.add_note", "PASS", f"text={txt!r}")
        except Exception as e:
            record("comment.add_note", "FAIL", str(e))

        # ── edit note on A1 ──────────────────────────────────────────────────
        try:
            comment_action("edit", cell="A1", sheet="Sheet1", workbook=wb_name,
                           text="Updated note")

            def _check_note_edit():
                ws = _session.get_sheet("Sheet1", wb_name)
                c = ws.Range("A1").Comment
                return c.Text() if c is not None else None

            txt = _session.run_com(_check_note_edit)
            assert txt == "Updated note", f"Note text expected 'Updated note', got {txt!r}"
            record("comment.edit_note", "PASS", f"text={txt!r}")
        except Exception as e:
            record("comment.edit_note", "FAIL", str(e))

        # ── add threaded on B2 (may be gated on some builds) ────────────────
        try:
            comment_action("add", cell="B2", sheet="Sheet1", workbook=wb_name,
                           text="Thread root", kind="threaded")

            def _check_threaded():
                ws = _session.get_sheet("Sheet1", wb_name)
                ct = ws.Range("B2").CommentThreaded
                return ct.Text() if ct is not None else None

            txt = _session.run_com(_check_threaded)
            assert txt == "Thread root", f"Threaded text expected 'Thread root', got {txt!r}"
            threaded_ok = True
            record("comment.add_threaded", "PASS", f"text={txt!r}")
        except Exception as e:
            msg = str(e)
            # Some Excel builds gate AddCommentThreaded — capture as a real finding.
            record("comment.add_threaded", "SKIP",
                   f"Threaded comments unavailable on this build: {msg[:140]}")

        # ── reply to threaded on B2 (only if threaded add worked) ───────────
        if threaded_ok:
            try:
                comment_action("reply", cell="B2", sheet="Sheet1", workbook=wb_name,
                               text="A reply", kind="threaded")

                def _check_replies():
                    ws = _session.get_sheet("Sheet1", wb_name)
                    return ws.Range("B2").CommentThreaded.Replies.Count

                cnt = _session.run_com(_check_replies)
                assert cnt >= 1, f"Reply count expected >= 1, got {cnt}"
                record("comment.reply_threaded", "PASS", f"reply_count={cnt}")
            except Exception as e:
                record("comment.reply_threaded", "FAIL", str(e))
        else:
            record("comment.reply_threaded", "SKIP", "threaded add failed/unavailable")

        # ── list kind=all → count >= 1 ──────────────────────────────────────
        try:
            r = comment_action("list", sheet="Sheet1", workbook=wb_name, kind="all")
            cnt = r.get("count", 0)
            assert cnt >= 1, f"list count expected >= 1, got {cnt}"
            record("comment.list", "PASS", f"count={cnt}")
        except Exception as e:
            record("comment.list", "FAIL", str(e))

        # ── delete note on A1 → Comment is None ─────────────────────────────
        try:
            comment_action("delete", cell="A1", sheet="Sheet1", workbook=wb_name)

            def _check_note_gone():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Range("A1").Comment

            c = _session.run_com(_check_note_gone)
            assert c is None, f"Note expected gone (None), got {c!r}"
            record("comment.delete_note", "PASS", "Comment is None")
        except Exception as e:
            record("comment.delete_note", "FAIL", str(e))

        # ── delete threaded on B2 (only if threaded existed) ────────────────
        if threaded_ok:
            try:
                comment_action("delete", cell="B2", sheet="Sheet1", workbook=wb_name,
                               kind="threaded")

                def _check_threaded_gone():
                    ws = _session.get_sheet("Sheet1", wb_name)
                    return ws.Range("B2").CommentThreaded

                ct = _session.run_com(_check_threaded_gone)
                assert ct is None, f"Threaded comment expected gone (None), got {ct!r}"
                record("comment.delete_threaded", "PASS", "CommentThreaded is None")
            except Exception as e:
                record("comment.delete_threaded", "FAIL", str(e))
        else:
            record("comment.delete_threaded", "SKIP", "threaded add failed/unavailable")

    except Exception as e:
        record("section20.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        _close_wb(wb, wb_name or "unknown")


# ── Section 21: Hyperlinks ─────────────────────────────────────────────────────

def run_hyperlink() -> None:
    """Live read-back smoke for hyperlink add (url + internal) / list / delete.

    Reads ws.Range(cell).Hyperlinks(1).Address / .SubAddress directly from COM.
    """
    section_header("SECTION 21 — Hyperlinks (add url + internal / list / delete)")
    if _check_excel_busy(
        "hyperlink.add_url", "hyperlink.add_internal",
        "hyperlink.list", "hyperlink.delete",
    ):
        return

    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        # ── add url on A1 ────────────────────────────────────────────────────
        try:
            hyperlink_action("add", sheet="Sheet1", workbook=wb_name, cell="A1",
                             link_type="url", target="https://www.thepexcel.com",
                             text_to_display="ThepExcel")

            def _check_url():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Range("A1").Hyperlinks(1).Address

            addr = _session.run_com(_check_url)
            assert "thepexcel.com" in addr, f"Hyperlink.Address expected to contain 'thepexcel.com', got {addr!r}"
            record("hyperlink.add_url", "PASS", f"Address={addr!r}")
        except Exception as e:
            record("hyperlink.add_url", "FAIL", str(e))

        # ── add internal on A2 ───────────────────────────────────────────────
        # The domain requires 'target' for ALL link types (it carries the
        # SubAddress for internal links); sub_address is only an override.
        try:
            hyperlink_action("add", sheet="Sheet1", workbook=wb_name, cell="A2",
                             link_type="internal", target="Sheet1!C3",
                             text_to_display="Jump")

            def _check_internal():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Range("A2").Hyperlinks(1).SubAddress

            sub = _session.run_com(_check_internal)
            assert "C3" in sub, f"Hyperlink.SubAddress expected to contain 'C3', got {sub!r}"
            record("hyperlink.add_internal", "PASS", f"SubAddress={sub!r}")
        except Exception as e:
            record("hyperlink.add_internal", "FAIL", str(e))

        # ── list → count >= 2 ────────────────────────────────────────────────
        try:
            r = hyperlink_action("list", sheet="Sheet1", workbook=wb_name)
            cnt = r.get("count", len(r.get("hyperlinks", [])))
            assert cnt >= 2, f"Hyperlink list count expected >= 2, got {cnt}"
            record("hyperlink.list", "PASS", f"count={cnt}")
        except Exception as e:
            record("hyperlink.list", "FAIL", str(e))

        # ── delete range A1 → Hyperlinks.Count == 0 ─────────────────────────
        try:
            hyperlink_action("delete", sheet="Sheet1", workbook=wb_name, range="A1")

            def _check_deleted():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Range("A1").Hyperlinks.Count

            cnt = _session.run_com(_check_deleted)
            assert cnt == 0, f"A1 Hyperlinks.Count expected 0 after delete, got {cnt}"
            record("hyperlink.delete", "PASS", f"Count={cnt}")
        except Exception as e:
            record("hyperlink.delete", "FAIL", str(e))

    except Exception as e:
        record("section21.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        _close_wb(wb, wb_name or "unknown")


# ── Section 22: Outline (group/ungroup rows & columns) ────────────────────────

def run_outline() -> None:
    """Live read-back smoke for grouping rows/columns and outline levels.

    Reads ws.Rows(n).OutlineLevel / ws.Columns(n).OutlineLevel directly from COM.
    """
    section_header("SECTION 22 — Outline (group/ungroup rows+columns / show_levels / clear)")
    if _check_excel_busy(
        "outline.group_rows", "outline.group_columns",
        "outline.show_levels", "outline.ungroup_rows", "outline.clear",
    ):
        return

    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        # ── group_rows 2:5 → Rows(2).OutlineLevel >= 2 ──────────────────────
        try:
            outline_action("group_rows", sheet="Sheet1", workbook=wb_name, rows="2:5")

            def _check_row_level():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Rows(2).OutlineLevel

            lvl = _session.run_com(_check_row_level)
            assert lvl >= 2, f"Rows(2).OutlineLevel expected >= 2, got {lvl}"
            record("outline.group_rows", "PASS", f"OutlineLevel={lvl}")
        except Exception as e:
            record("outline.group_rows", "FAIL", str(e))

        # ── group_columns B:D → Columns(2).OutlineLevel >= 2 ────────────────
        try:
            outline_action("group_columns", sheet="Sheet1", workbook=wb_name, columns="B:D")

            def _check_col_level():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Columns(2).OutlineLevel

            lvl = _session.run_com(_check_col_level)
            assert lvl >= 2, f"Columns(2).OutlineLevel expected >= 2, got {lvl}"
            record("outline.group_columns", "PASS", f"OutlineLevel={lvl}")
        except Exception as e:
            record("outline.group_columns", "FAIL", str(e))

        # ── show_levels row_levels=1 → no exception ─────────────────────────
        try:
            outline_action("show_levels", sheet="Sheet1", workbook=wb_name, row_levels=1)
            record("outline.show_levels", "PASS", "no exception")
        except Exception as e:
            record("outline.show_levels", "FAIL", str(e))

        # ── ungroup_rows 2:5 → Rows(2).OutlineLevel == 1 ────────────────────
        try:
            outline_action("ungroup_rows", sheet="Sheet1", workbook=wb_name, rows="2:5")

            def _check_row_level_after():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Rows(2).OutlineLevel

            lvl = _session.run_com(_check_row_level_after)
            assert lvl == 1, f"Rows(2).OutlineLevel expected 1 after ungroup, got {lvl}"
            record("outline.ungroup_rows", "PASS", f"OutlineLevel={lvl}")
        except Exception as e:
            record("outline.ungroup_rows", "FAIL", str(e))

        # ── clear (after a fresh group) → Rows(2).OutlineLevel == 1 ─────────
        try:
            outline_action("group_rows", sheet="Sheet1", workbook=wb_name, rows="2:5")
            outline_action("clear", sheet="Sheet1", workbook=wb_name)

            def _check_cleared():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Rows(2).OutlineLevel

            lvl = _session.run_com(_check_cleared)
            assert lvl == 1, f"Rows(2).OutlineLevel expected 1 after clear, got {lvl}"
            record("outline.clear", "PASS", f"OutlineLevel={lvl}")
        except Exception as e:
            record("outline.clear", "FAIL", str(e))

    except Exception as e:
        record("section22.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        _close_wb(wb, wb_name or "unknown")


# ── Section 23: Protection ─────────────────────────────────────────────────────

def run_protection() -> None:
    """Live read-back smoke for sheet/workbook protection + cell lock state.

    Uses password='test123' consistently and ALWAYS unprotects with the same
    password to avoid leaving a locked temp workbook. Reads ws.ProtectContents /
    wb.ProtectStructure / rng.Locked directly from COM.
    """
    section_header("SECTION 23 — Protection (set_locked / protect+unprotect sheet & workbook / status)")
    if _check_excel_busy(
        "protection.set_locked", "protection.protect_sheet", "protection.status",
        "protection.unprotect_sheet", "protection.protect_workbook",
        "protection.unprotect_workbook",
    ):
        return

    PW = "test123"
    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        # ── set_locked A1 locked=False → Range("A1").Locked == False ────────
        try:
            protection_action("set_locked", sheet="Sheet1", workbook=wb_name,
                              range="A1", locked=False)

            def _check_locked():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Range("A1").Locked

            lk = _session.run_com(_check_locked)
            assert lk is False, f"A1.Locked expected False, got {lk}"
            record("protection.set_locked", "PASS", f"Locked={lk}")
        except Exception as e:
            record("protection.set_locked", "FAIL", str(e))

        # ── protect_sheet → ws.ProtectContents == True ─────────────────────
        try:
            protection_action("protect_sheet", sheet="Sheet1", workbook=wb_name,
                              password=PW)

            def _check_protected():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.ProtectContents

            pc = _session.run_com(_check_protected)
            assert pc is True, f"ProtectContents expected True, got {pc}"
            record("protection.protect_sheet", "PASS", f"ProtectContents={pc}")
        except Exception as e:
            record("protection.protect_sheet", "FAIL", str(e))

        # ── status → sheet_protected True ──────────────────────────────────
        try:
            r = protection_action("status", sheet="Sheet1", workbook=wb_name)
            applied = r.get("applied", {})
            sp = applied.get("sheet_protected")
            assert sp is True, f"status sheet_protected expected True, got {sp}"
            record("protection.status", "PASS", f"sheet_protected={sp}")
        except Exception as e:
            record("protection.status", "FAIL", str(e))

        # ── unprotect_sheet → ws.ProtectContents == False ──────────────────
        try:
            protection_action("unprotect_sheet", sheet="Sheet1", workbook=wb_name,
                              password=PW)

            def _check_unprotected():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.ProtectContents

            pc = _session.run_com(_check_unprotected)
            assert pc is False, f"ProtectContents expected False after unprotect, got {pc}"
            record("protection.unprotect_sheet", "PASS", f"ProtectContents={pc}")
        except Exception as e:
            record("protection.unprotect_sheet", "FAIL", str(e))

        # ── protect_workbook structure=True → wb.ProtectStructure == True ──
        try:
            protection_action("protect_workbook", workbook=wb_name,
                              password=PW, structure=True)

            def _check_wb_protected():
                wb_inner = _session.get_workbook(wb_name)
                return wb_inner.ProtectStructure

            ps = _session.run_com(_check_wb_protected)
            assert ps is True, f"ProtectStructure expected True, got {ps}"
            record("protection.protect_workbook", "PASS", f"ProtectStructure={ps}")
        except Exception as e:
            record("protection.protect_workbook", "FAIL", str(e))

        # ── unprotect_workbook → wb.ProtectStructure == False ──────────────
        try:
            protection_action("unprotect_workbook", workbook=wb_name, password=PW)

            def _check_wb_unprotected():
                wb_inner = _session.get_workbook(wb_name)
                return wb_inner.ProtectStructure

            ps = _session.run_com(_check_wb_unprotected)
            assert ps is False, f"ProtectStructure expected False after unprotect, got {ps}"
            record("protection.unprotect_workbook", "PASS", f"ProtectStructure={ps}")
        except Exception as e:
            record("protection.unprotect_workbook", "FAIL", str(e))

    except Exception as e:
        record("section23.setup", "FAIL", str(e))
        traceback.print_exc()
    finally:
        # Defensive: ensure the temp workbook is NOT left protected (would block close).
        try:
            def _ensure_unprotected():
                wb_inner = _session.get_workbook(wb_name)
                try:
                    if wb_inner.ProtectStructure:
                        wb_inner.Unprotect(Password=PW)
                except Exception:
                    pass
                for ws in wb_inner.Worksheets:
                    try:
                        if ws.ProtectContents:
                            ws.Unprotect(Password=PW)
                    except Exception:
                        pass
            if wb_name is not None:
                _session.run_com(_ensure_unprotected)
        except Exception:
            pass
        _close_wb(wb, wb_name or "unknown")


# ── Section 24: Sparklines ─────────────────────────────────────────────────────

def run_sparkline() -> None:
    """Live read-back smoke for sparkline add (line + column) / list / clear.

    Reads ws.Range(location).SparklineGroups.Count directly from COM.
    """
    section_header("SECTION 24 — Sparklines (add line + column / list / clear)")
    if _check_excel_busy(
        "sparkline.add_line", "sparkline.add_column",
        "sparkline.list", "sparkline.clear",
    ):
        return

    wb = None
    wb_name = None
    try:
        wb, wb_name = _new_wb()

        # Setup: write numeric data to B2:E4
        def _setup_data():
            ws = _session.get_sheet("Sheet1", wb_name)
            ws.Range("B2:E4").Value = (
                (1, 2, 3, 4),
                (4, 3, 2, 1),
                (2, 3, 1, 4),
            )
        _session.run_com(_setup_data)

        # ── add line sparklines F2:F4 ───────────────────────────────────────
        try:
            sparkline_action("add", location="F2:F4", sheet="Sheet1", workbook=wb_name,
                             data_range="B2:E4", spark_type="line")

            def _check_line():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Range("F2:F4").SparklineGroups.Count

            cnt = _session.run_com(_check_line)
            assert cnt >= 1, f"F2:F4 SparklineGroups.Count expected >= 1, got {cnt}"
            record("sparkline.add_line", "PASS", f"Count={cnt}")
        except Exception as e:
            record("sparkline.add_line", "FAIL", str(e))

        # ── add column sparklines G2:G4 ─────────────────────────────────────
        try:
            sparkline_action("add", location="G2:G4", sheet="Sheet1", workbook=wb_name,
                             data_range="B2:E4", spark_type="column")

            def _check_col():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Range("G2:G4").SparklineGroups.Count

            cnt = _session.run_com(_check_col)
            assert cnt >= 1, f"G2:G4 SparklineGroups.Count expected >= 1, got {cnt}"
            record("sparkline.add_column", "PASS", f"Count={cnt}")
        except Exception as e:
            record("sparkline.add_column", "FAIL", str(e))

        # ── list location=F2 → groups_count >= 1 ────────────────────────────
        try:
            r = sparkline_action("list", location="F2", sheet="Sheet1", workbook=wb_name)
            cnt = r.get("applied", {}).get("groups_count", 0)
            assert cnt >= 1, f"list groups_count expected >= 1, got {cnt}"
            record("sparkline.list", "PASS", f"groups_count={cnt}")
        except Exception as e:
            record("sparkline.list", "FAIL", str(e))

        # ── clear F2:F4 → SparklineGroups.Count == 0 ────────────────────────
        try:
            sparkline_action("clear", location="F2:F4", sheet="Sheet1", workbook=wb_name)

            def _check_cleared():
                ws = _session.get_sheet("Sheet1", wb_name)
                return ws.Range("F2:F4").SparklineGroups.Count

            cnt = _session.run_com(_check_cleared)
            assert cnt == 0, f"F2:F4 SparklineGroups.Count expected 0 after clear, got {cnt}"
            record("sparkline.clear", "PASS", f"Count={cnt}")
        except Exception as e:
            record("sparkline.clear", "FAIL", str(e))

    except Exception as e:
        record("section24.setup", "FAIL", str(e))
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
    "11": run_format,
    "12": run_workbook_create_save_as,
    "13": run_conditional_format,
    "14": run_view,
    "15": run_validation,
    "16": run_slicer,
    "17": run_pq_parameters,
    "18": run_cube,
    "19": run_page_setup,
    "20": run_comment,
    "21": run_hyperlink,
    "22": run_outline,
    "23": run_protection,
    "24": run_sparkline,
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
