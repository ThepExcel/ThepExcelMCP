---
name: excel-god
description: "Strategy + live-orchestration guide for the thepexcel-excel MCP: triage, architecture framework, tool-call recipes, verification, gotchas."
scope_note: |
  Entry point for all live-Excel work via the thepexcel-excel MCP (26 tools). Triage
  into audit / build / extend, apply the architecture decision framework, then map to
  concrete tool-call sequences. Verify visually after building anything that renders
  (chart, pivot, formatted range).
out_of_scope: |
  File-only xlsx manipulation without a live session (use openpyxl directly).
  Power BI DAX/modelling.
  Office Scripts (cloud-only, no COM path — server redirects to VBA).
---

# Excel God — Excel Strategy & Live Orchestrator

Entry point for working with a **live open Excel Desktop session** via the
`thepexcel-excel` MCP server (26 tools). Triage, apply the architecture framework,
then map to concrete tool sequences. Always verify after building.

> **Prerequisite:** Excel must be open. Confirm with
> `excel_workbook(action="list")` before any other call.

---

## Step 0 — Triage

| Request shape | Route |
|---|---|
| Build a new report / dashboard | Build workflow (§ Workflow Recipes, recipe A) |
| Audit an existing workbook | Audit workflow (§ Workflow Recipes, recipe B) |
| Add measure / query / pivot to an existing workbook | Extend workflow (§ Workflow Recipes, recipe C) |
| Run a macro / do something native features cannot | VBA (§ VBA — Last Resort) |

---

## Tool / Layer Selection — which engine for which task

| Task shape | Engine | Why / not |
|---|---|---|
| Author a NEW .xlsx in bulk (hundreds–thousands of rows), no live features needed | **openpyxl (file-based)** | Fast, no open Excel, no agent-context cost; openpyxl writes Tables + real dates |
| Live features on an OPEN workbook — PivotTable, Chart, Power Query (M), Data Model + DAX, VBA, dynamic-array/LAMBDA, spill introspection, screenshot | **thepexcel-excel MCP** | openpyxl cannot do these — need live COM |
| Bulk data that then needs live enhancement | **Hybrid:** build the file offline with openpyxl → `excel_workbook(action="open")` → add pivots/charts/model via MCP | Best of both; avoids slow/large MCP bulk writes |
| Inspect / audit an existing live workbook | **thepexcel-excel MCP** (Recipe B) | |

> **Never** bulk-write a large 2-D array into an MCP-owned Excel from a SEPARATE Python/pywin32 process (GetActiveObject/ROT) — cross-apartment marshaling silently drops the write (only the top-left cell lands). Bulk data must go either through the MCP itself (in-apartment) or via an offline openpyxl file that the MCP then opens.

> **Data-write reliability guide:** (a) **Preferred for bulk/tabular data** — openpyxl offline → `excel_workbook(open)` (fastest, zero MCP context cost); (b) **Preferred for live edits** — `excel_table(append_rows / add_column)` + `write_formula` (reliable single-cell-path calls, read-back verify); (c) **Raw multi-cell `excel_range(write)`** — the `.Resize` bug is FIXED as of 2026-06-24; ALWAYS read-back verify the result before proceeding.

---

## Architecture Decision Framework

> Apply this before touching any tool.

### When to use each Excel layer

| Situation | Right layer | Wrong layer |
|---|---|---|
| Raw data from external source or repeating import | **Power Query** | Manual paste / manual transform |
| Multiple related tables, measures reused across pivots | **Data Model + DAX measures** | VLOOKUP-merged flat tables |
| Aggregation / cross-tab for presentation | **PivotTable** (source = Data Model) | Formulas on a flat sheet |
| Small derived column used only in THIS table | Table structured-reference formula or dynamic array | Thousand-row fill-down |
| Reusable formula logic called from multiple cells | **Named LAMBDA** | Nested IF chains or helper columns |
| Feature unavailable natively | **VBA** (last resort — 2-gate setup) | Any of the above |

### Calculated columns vs measures in DAX

- **Default = measure.** Measures are context-sensitive, work everywhere,
  compress better in the model.
- Calculated column only when you need row-level values visible in the table
  (e.g., a segmentation label), not for aggregations.

### Dynamic arrays and LAMBDA

- Prefer `UNIQUE`, `FILTER`, `SORT`, `XLOOKUP`, `SEQUENCE` over helper-column chains.
- Write dynamic-array formulas with `excel_range(action="write_formula")` — uses `Formula2`.
- Named LAMBDA for formula logic called from 3+ cells:
  `excel_name(action="set", name="TAX", refers_to="=LAMBDA(amount, rate, amount*rate/100)")`.

---

## Verification Principle — Verify EFFECT, not return value

**MUTATING ops can return `success` with ZERO real effect** — silent failures confirmed in
live testing: `excel_range write` wrote only `values[0][0]`; `excel_table sort` returned
success but never reordered. A `success` response is a proxy, not proof.

**Rule:** after ANY mutation (write, write_formula, sort, filter, add_column, set_layout,
measure/relationship adds, pivot field ops) — **READ BACK** the affected range/table/column
(or screenshot for visual objects) and confirm the actual change before reporting done.

| Mutation type | Verification method |
|---|---|
| `excel_range write / write_formula` | `excel_range(read, same range)` — confirm values |
| `excel_table sort / filter / add_column` | `excel_table(read, ...)` first rows — confirm order/values |
| `excel_pivot add_field / set_layout` | `excel_screenshot(sheet)` → inspect PNG |
| `excel_chart create / configure` | `excel_screenshot(chart)` → inspect PNG |
| `excel_datamodel add_measure / add_relationship` | `excel_datamodel(list_measures/list_relationships)` — confirm presence |

Exit-0 + right file size = proxy, not proof.

---

## Workflow Recipes

### Recipe A — Raw data to report (build from scratch)

```
1. excel_workbook(action="list")                          # confirm workbook
2. excel_powerquery(action="create", name, formula)       # M: load + clean data
3. excel_powerquery(action="load_to_datamodel", name)     # push to Data Model
   ⚠ DEADLOCK RISK (stdio/Claude Code): load_to_datamodel blocks the STA COM worker
     and bricks ALL subsequent MCP calls until Excel is force-killed. Use only with
     Claude Desktop + visible Excel. FALLBACK: use load_to_table instead → build
     the PivotTable from the worksheet Table (source=<table or range>).
   OR excel_powerquery(action="load_to_table", name, sheet_name)   # → Table on sheet (safe)
4. excel_datamodel(action="add_relationship", ...)        # link fact <-> dim tables
5. excel_datamodel(action="add_measure", ...)             # DAX measures
6. excel_pivot(action="create", source="datamodel", name, dest_sheet)
7. excel_pivot(action="add_field", ...)                   # rows / columns / values
8. excel_chart(action="create", source, chart_type)       # optional chart
9. excel_screenshot(action="sheet")  -->  inspect image   # verify (see Verification Principle above)
```

### Recipe B — Audit an existing workbook

```
1. excel_workbook(action="info")                          # sheet + object counts
2. excel_powerquery(action="list")                        # queries + model_connection flag
3. excel_table(action="list")                             # tables
4. excel_datamodel(action="list_measures")                # DAX measures
5. excel_pivot(action="list")                             # pivots + sources
6. excel_powerquery(action="analyze", name)               # M anti-pattern scan
7. excel_range(action="read", range=<sample>)             # spot-check data shape
```

Emit a structured finding: layer used, potential improvements, tech-debt flags.

### Recipe C — Extend an existing model

```
1. excel_workbook(action="info")  +  excel_datamodel(action="list_tables")   # orient
2. excel_powerquery(action="get", name) / excel_datamodel(action="list_measures")
3. Add query / measure / pivot field as needed
4. excel_datamodel(action="refresh") or excel_powerquery(action="refresh_all")
   ⚠ DEADLOCK RISK (stdio/Claude Code): model refresh blocks the STA COM worker —
     same as load_to_datamodel. Safe for Claude Desktop + visible Excel only. FALLBACK:
     skip the refresh call; if data was already in the model, adding measures/relationships
     to an EXISTING populated model is fine (no refresh involved). For a new query, use
     load_to_table and pivot from the worksheet Table instead.
5. excel_screenshot(action="sheet")  -->  inspect image
```

---

## Safety Rules (user's workbook = production)

1. **Never save without being asked.** `excel_workbook(action="save")` only on
   explicit request. `close` also does not auto-save — safe default.
2. **Confirm before destructive actions.** Deleting a sheet, query, measure, or
   relationship is irreversible within the session. Ask first.
3. **Prefer new sheets over overwriting.** Add a new sheet
   (`excel_sheet(action="add", name="Output")`) rather than overwriting existing data.
4. **PQ refresh can deadlock in pure automation.** `load_to_datamodel` and
   `load_to_table` need Excel's UI message pump. Works with a **visible Excel window**.
   In headless context warn the user: may time out at 120 s, requiring manual refresh.
5. **Snapshot before high-risk ops.** Use `excel_snapshot(action="snapshot")` before
   any destructive batch (bulk replace, delete queries, restructure model) so the user
   can `restore` without losing work.

---

## Gotchas (from architecture + live smoke tests)

| Gotcha | Detail |
|---|---|
| **Reads are paginated** | Default 100 rows. Check `has_more` / `next_offset` and loop if needed. |
| **Measure update = delete + re-add** | `update_measure` replaces formula by deleting and re-adding. Pass only fields to change. |
| **`write_py` async + experimental** | `=PY()` runs in Azure cloud. Cell shows `#BUSY!`. Cannot await — use `read` in a follow-up call. Not officially documented. |
| **Office Scripts not supported** | Cloud-only, no COM path. Use `excel_vba` instead — server returns actionable error. |
| **Data Model `add_table` / `load_to_datamodel` can DEADLOCK the server** | `WorkbookConnection.Refresh()` on a Mashup/PQ model connection needs Excel's UI message pump. In the Claude Code + stdio context it does NOT just time out — it blocks the single STA COM worker thread and **bricks ALL subsequent MCP calls until Excel is force-killed**. Avoid in that context; build the model interactively or via Claude Desktop. Measures/relationships on an EXISTING model are unaffected (no refresh). |
| **`close` with pending PQ** | Can deadlock 30 s. Server uses best-effort query deletion + timeout. |
| **VBA 2-gate opt-in** | `THEPEXCEL_MCP_ENABLE_VBA=1` env var AND AccessVBOM=1 registry setting. Workbook must be `.xlsm` or `.xlsb` to save VBA. |
| **Structured references via `excel_range`** | `"TableName[ColumnName]"` works via Excel's native Range() parser — no special handling needed. |
| **`excel_range write` multi-cell — FIXED** | The `.Resize` pywin32 dispatch quirk (only top-left cell written) is fixed in the current build; the corrected path uses `Range(Cells,Cells)`. Still always read-back verify — a `success` response is never proof of EFFECT. |
| **Bulk via MCP costs agent context** | `write` values pass through the tool call. For large data prefer the offline-openpyxl → open hybrid. |
| **External-process bulk write drops silently** | A separate pywin32 process cannot reliably write a big 2-D array into an MCP-owned Excel (cross-apartment). Use the MCP, or build offline + open. |
| **`excel_pivot(create)` — dest_sheet auto-creates** | Omit `dest_sheet` → a new `Pivot_<name>` sheet is created. Name an EXISTING sheet → the pivot is placed there. Name a sheet that does NOT exist yet → it is created automatically (fixed in the current build; older builds errored "sheet not found"). |
| **`excel_table(filter, filter_op="clear_filters")` requires filter_column** | Even when clearing all filters, `filter_column` must still be passed (the value is ignored but the param cannot be absent). |
| **`excel_name(list)` includes Excel internal hidden names** | Returns `_xlfn.*`, `_xlpm.*`, `_xleta.*` entries mixed with user-defined names. Skip any name starting with `_xl` when presenting results to the user. |
| **Excel must be RUNNING before any MCP call** | If `excel_workbook(list)` returns "Excel is not running", open the target file via the OS — it registers in the Running Object Table and the MCP finds it. Auto-launch only occurs when `THEPEXCEL_MCP_AUTOLAUNCH=1` is set. |
| **Screenshots and chart PNG require Excel VISIBLE / foreground** | `excel_screenshot` and `excel_chart(export_image)` use CopyPicture + PIL. A minimized or background Excel window yields a 0-byte PNG. Bring the target workbook to the foreground before capturing. |
| **Named LAMBDA: avoid cell-reference-looking parameter names** | Parameter names like `q1`, `x2`, `c3` silently fail. Use `val`, `rate`, `n`, `x` etc. A failed LAMBDA add can leave undeletable hidden `_xlpm.*` helper names that block later LAMBDA adds in the SAME workbook — retry in a fresh workbook. |
| **Tool count is 26** | The MCP ships 26 registered tools. Old docs or cached model knowledge may cite 11 or 24 — ignore those counts. |

---

## VBA — Last Resort

Use only when no native Excel feature covers the need.

```
1. Confirm both opt-in gates (env var + AccessVBOM registry)
2. excel_vba(action="write_module", module_name, code)
3. excel_vba(action="run", proc_name, args)
4. excel_vba(action="delete_module") when done — keep model clean
5. Remind user to save as .xlsm / .xlsb
```

---

## Examples

**Build a sales report from a CSV:**
Triage → Recipe A. Create PQ query loading the CSV, load to Data Model,
add `Total Sales = SUM(Sales[Amount])` measure, create PivotTable from
datamodel, screenshot + inspect to verify layout.

**Audit a received workbook:**
Triage → Recipe B. `workbook info` → `powerquery list` → `datamodel list_measures`
→ `powerquery analyze` on largest query — flag hardcoded paths, missing folding.

**Add a LAMBDA for reusable tax logic:**
`excel_name(action="set", name="TAX", refers_to="=LAMBDA(amount, rate, amount*rate/100)")`.
Then in cells: `=TAX([@Amount], 7)`.

---

## Tool Quick Reference (26 tools)

| Tool | Category | Key actions |
|---|---|---|
| `excel_workbook` | Core | list · info · open · save · close · create · save_as |
| `excel_sheet` | Core | list · add · rename · delete |
| `excel_range` | Core | read (paginated) · read_spill · write · write_formula · write_py · clear |
| `excel_powerquery` | Data/PQ/Model | list · get · create · update · delete · refresh · refresh_all · load_to_table · load_to_datamodel · analyze · analyze_raw · create_parameter · get_parameter · set_parameter · list_parameters |
| `excel_table` | Data/PQ/Model | list · create · read · append_rows · add_column · sort · filter · set_style · toggle_totals · rename · delete |
| `excel_pivot` | Data/PQ/Model | list · create · add_field · remove_field · move_field · set_layout · refresh · delete · read |
| `excel_datamodel` | Data/PQ/Model | info · list_tables · add_table · list_relationships · add_relationship · delete_relationship · list_measures · add_measure · update_measure · delete_measure · refresh · cube_formula · cube_value · cube_member · add_calculated_column · add_calculated_table |
| `excel_name` | Data/PQ/Model | list · get · set · delete (named ranges + LAMBDA) |
| `excel_chart` | Objects | list · create · configure · set_source · export_image · delete |
| `excel_shape` | Objects | add_image · add_textbox · add_shape · list · move · delete |
| `excel_slicer` | Objects | add · add_timeline · list · delete · connect |
| `excel_sparkline` | Objects | add · clear · list |
| `excel_format` | Formatting/View | font · fill · border · number_format · alignment · column_width · row_height · autofit |
| `excel_view` | Formatting/View | freeze_panes · unfreeze_panes · gridlines · zoom · headings |
| `excel_conditional_format` | Formatting/View | data_bar · color_scale · icon_set · cell_rule · top_bottom · clear |
| `excel_validation` | Formatting/View | list · whole_number · decimal · date · text_length · custom · clear |
| `excel_page_setup` | Page/Print | set · print_area · print_titles · header_footer · export_pdf · get |
| `excel_comment` | Annotation | add · edit · reply · delete · list · get |
| `excel_hyperlink` | Annotation | add · list · delete |
| `excel_outline` | Annotation | group_rows · group_columns · ungroup_rows · ungroup_columns · show_levels · clear |
| `excel_protection` | Annotation | protect_sheet · unprotect_sheet · protect_workbook · unprotect_workbook · set_locked · status |
| `excel_vba` | Safety | list_modules · get_module · write_module · delete_module · run |
| `excel_screenshot` | Safety | range · sheet · chart (PNG for visual verification) |
| `excel_find_replace` | Safety | find · count · replace |
| `excel_diff` | Safety | ranges · sheets |
| `excel_snapshot` | Safety | snapshot · list · restore · delete |

Full param specs: the `thepexcel-excel` MCP server tool docstrings (`src/thepexcel_mcp/server.py`).

---

## Reading Power Query M Code Without Excel (file-only audit)

Use when: **Excel is not open / unavailable** and the goal is to read, audit, or batch-extract
source URLs from `.xlsx` files — e.g., "what URLs are these workbooks pulling from?"
or "do any of our templates hardcode a stale endpoint?"

**Do NOT use this path when you need to EDIT M code** — writing back requires repacking the
DataMashup binary and recomputing its checksum/permissions. For edits, use the live Excel MCP
(`excel_powerquery(action="update")`).

### When to use file-only vs the live Excel MCP

| Goal | Method |
|---|---|
| Read / audit M code, extract source URLs, batch scan | File-only (`extract_mcode.py`) — no Excel needed |
| Create or update a query | Live Excel MCP (`excel_powerquery`) — Excel must be open |
| Inspect a workbook interactively | Live Excel MCP (Recipe B in § Workflow Recipes) |

### How M code is stored (nested structure)

```
file.xlsx  (ZIP)
└── customXml/item*.xml          ← UTF-16 encoded XML
    └── <DataMashup ...>BASE64</DataMashup>
        └── base64 decode → binary
            [int32 version][int32 package_len][package ZIP][metadata…]
            └── package ZIP  (slice exactly package_len bytes — trailing metadata corrupts zipfile)
                └── Formulas/Section1.m   ← all 'shared <QueryName> = ...' definitions
                                             Web.Contents("https://…") source URLs live here
```

**Legacy fallback** (older "From Web" / non-mashup workbooks): source URLs live as plaintext in
`xl/connections.xml` (dbPr connection string) and/or `xl/queryTables/queryTable*.xml`. No base64,
no inner ZIP. `extract_mcode.py` checks this automatically when no DataMashup part is found.

### Script: `extract_mcode.py`

Location: `skills/excel-god/scripts/extract_mcode.py` (relative to repo root)

Stdlib-only (zipfile, base64, struct, re, sys) — no pip install. UTF-8 output safe on Windows.

```bash
# Single file
python skills/excel-god/scripts/extract_mcode.py "path/to/file.xlsx"

# Batch (explicit list)
python skills/excel-god/scripts/extract_mcode.py file1.xlsx file2.xlsx file3.xlsx
```

Output per file:
1. Full `Section1.m` (all shared query definitions)
2. All `http/https` source URLs
3. Source-assignment lines for quick context (`Source = ...`, `Web.Contents(...)`, etc.)
