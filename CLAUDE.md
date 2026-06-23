---
purpose: |
  Let me control Excel with AI without limits — going far beyond Copilot, Claude-in-Excel or any official tool — so the Excel work behind my teaching and projects has no ceiling, and I stay the person known for mastering Excel more deeply than anyone.
---

# CLAUDE.md — ThepExcelMCP

> Windows-only MCP server. Controls live Excel Desktop via COM (pywin32).
> Gives AI agents full Excel capability beyond what openpyxl can do.

## Architecture

```
Claude / MCP Client
      │ stdio
FastMCP server (server.py)
      │ 14 action-dispatch tools
      ├── excel_workbook            → domains/workbook.py
      ├── excel_sheet               → domains/sheets.py
      ├── excel_range               → domains/ranges.py         (+ read_spill, spill metadata)
      ├── excel_powerquery          → domains/powerquery.py + analysis/pq_analyzer.py
      ├── excel_table               → domains/tables.py
      ├── excel_pivot               → domains/pivots.py
      ├── excel_datamodel           → domains/datamodel.py      ← Phase 2
      ├── excel_vba                 → domains/vba.py            ← Phase 3
      ├── excel_name                → domains/names.py          ← Phase 3
      ├── excel_format              → domains/format.py         ← Phase 5 (Tier-1 gap)
      ├── excel_chart               → domains/charts.py         ← Phase 4
      ├── excel_screenshot          → domains/screenshot.py     ← Phase 4
      ├── excel_view                → domains/view.py           ← Phase 5
      ├── excel_conditional_format  → domains/conditional_format.py ← Phase 5
      ├── excel_validation          → domains/validation.py     ← Phase 5
      └── excel_slicer              → domains/slicer.py         ← Phase 5
                                          │
                                    ExcelSession (session.py)
                                          │ run_com() → STA COM worker thread
                                          │ win32com
                                    Excel.Application (COM)
                                          │
                                    Running Excel Desktop
```

### STA COM worker thread (Phase 3)

All COM calls now run on a single dedicated background thread that owns the STA
apartment (`pythoncom.CoInitialize()` once at startup). Tool handlers submit
callables via `queue.Queue` and receive results via `concurrent.futures.Future`.

- **`_session.run_com(fn, *args)`** — submit to worker, block until result (120s timeout).
- **`excel_guard(app)`** — context manager, sets `DisplayAlerts=False` around risky ops.
- **`wait_calculation(app)`** — polls `CalculationState` with `PumpWaitingMessages()`.
- **ROT fallback** — `get_workbook()` scans the Running Object Table when a workbook
  is not found in the first Excel instance (handles multiple Excel instances).

### Key files

| File | Role |
|---|---|
| `src/thepexcel_mcp/server.py` | FastMCP app + 14 tool registrations |
| `src/thepexcel_mcp/session.py` | `ExcelSession` — STA worker thread, `run_com()`, `excel_guard`, `wait_calculation`, ROT fallback |
| `src/thepexcel_mcp/domains/workbook.py` | Workbook CRUD |
| `src/thepexcel_mcp/domains/sheets.py` | Sheet CRUD |
| `src/thepexcel_mcp/domains/ranges.py` | Range read (paginated, spill metadata) / read_spill / write / write_formula (Formula2) / clear |
| `src/thepexcel_mcp/domains/powerquery.py` | 11 PQ operations — create/update include M code warnings; `load_to_datamodel` Phase 2 |
| `src/thepexcel_mcp/domains/tables.py` | 11 Table (ListObject) operations |
| `src/thepexcel_mcp/domains/pivots.py` | 9 PivotTable operations; data-model CubeField support in `add_field` |
| `src/thepexcel_mcp/domains/datamodel.py` | **Phase 2** — 11 Data Model ops (model tables, relationships, DAX measures) |
| `src/thepexcel_mcp/domains/vba.py` | **Phase 3** — VBA module CRUD + macro execution (opt-in via env + AccessVBOM check) |
| `src/thepexcel_mcp/domains/names.py` | **Phase 3** — Named ranges, LAMBDA formulas, defined name CRUD |
| `src/thepexcel_mcp/domains/format.py` | **Phase 5** — Range formatting: font/fill/border/number_format/alignment/sizing |
| `src/thepexcel_mcp/domains/charts.py` | **Phase 4** — Chart CRUD + configure + export PNG |
| `src/thepexcel_mcp/domains/screenshot.py` | **Phase 4** — Range/sheet/chart capture as PNG (CopyPicture+PIL) |
| `src/thepexcel_mcp/domains/view.py` | **Phase 5** — Worksheet display: freeze panes, gridlines, zoom, headings |
| `src/thepexcel_mcp/domains/conditional_format.py` | **Phase 5** — Conditional formatting: cell_rule, data_bar, color_scale, icon_set, top_bottom, clear |
| `src/thepexcel_mcp/domains/validation.py` | **Phase 5** — Data validation: list, whole_number, decimal, date, text_length, custom, clear |
| `src/thepexcel_mcp/domains/slicer.py` | **Phase 5** — Slicers + timelines: add, add_timeline, list, delete, connect |
| `src/thepexcel_mcp/analysis/pq_analyzer.py` | M code static analyzer — copied verbatim from PoC |

### load_to_table pattern (important)

`powerquery.py::_load_to_table` uses a specific Connections.Add2 + Mashup-OLEDB sequence
to load a query into a worksheet Table. This exact form is required:
- `lCmdtype=2` (xlCmdSql)
- `Location=<query_name>` WITHOUT quotes in the connection string
- `CommandText = "SELECT * FROM [<query_name>]"`

Do NOT simplify this — it took real debugging in the PoC to get right.

## Tool registry (Phase 1 + 2 + 3 + 4 + 5)

| Tool | Actions |
|---|---|
| `excel_workbook` | `list`, `info`, `open`, `save`, `close`, `create` (Workbooks.Add + optional SaveAs), `save_as` (FileFormat inferred from extension) |
| `excel_sheet` | `list`, `add`, `rename`, `delete` |
| `excel_range` | `read` (paginated, 100-row default; spill metadata in response), `read_spill` (full spill range for dynamic array anchor), `write`, `write_formula` (Formula2/dynamic arrays), `write_py` (**Phase 4** — `=PY()` Formula2R1C1 insertion, experimental), `clear` |
| `excel_powerquery` | `list`, `get`, `create`, `update`, `delete`, `refresh`, `refresh_all`, `load_to_table`, `load_to_datamodel` (**Phase 2**), `analyze`, `analyze_raw`. `create`/`update` return `warnings` list from M code analyzer (non-blocking). `list` includes `model_connection` flag. |
| `excel_table` | `list`, `create`, `read` (paginated), `append_rows`, `add_column` (with formula), `sort`, `filter` (equals/contains/greater/less/clear_filters), `set_style`, `toggle_totals` (per-column aggregation), `rename`, `delete` (keep_data/unlist) |
| `excel_pivot` | `list`, `create` (range/table/datamodel source), `add_field` (with aggregation+number_format; CubeField fallback for datamodel pivots **Phase 2**), `remove_field`, `move_field`, `set_layout` (compact/outline/tabular + subtotals + grand_totals), `refresh`, `delete`, `read` (paginated via TableRange1) |
| `excel_datamodel` | **Phase 2** — `info`, `list_tables`, `add_table` (table\|query → model via CreateModelConnection=True), `list_relationships`, `add_relationship`, `delete_relationship`, `list_measures`, `add_measure` (DAX + format), `update_measure`, `delete_measure`, `refresh` |
| `excel_vba` | **Phase 3** — `list_modules`, `get_module`, `write_module`, `delete_module`, `run`. Opt-in: `THEPEXCEL_MCP_ENABLE_VBA=1` + AccessVBOM registry check per call. |
| `excel_name` | **Phase 3** — `list`, `get`, `set`, `delete`. Covers named ranges, constants, and LAMBDA formulas. `is_lambda` flag on list/get. |
| `excel_format` | **Phase 5** — `font` (name/size/bold/italic/underline/color), `fill` (bg color or clear), `border` (sides: all/outline/inside/top/bottom/left/right; style: continuous/dash/double/none; weight: thin/medium/thick), `number_format` (any Excel code), `alignment` (h/v align, wrap_text, merge/unmerge), `column_width`, `row_height`, `autofit`. All colors as `"#RRGGBB"` → converted to Excel BGR internally. |
| `excel_chart` | **Phase 4** — `list`, `create`, `configure`, `set_source`, `export_image`, `delete` |
| `excel_screenshot` | **Phase 4** — `range`, `sheet`, `chart` (PNG capture for LLM visual verification) |
| `excel_view` | **Phase 5** — `freeze_panes`, `unfreeze_panes`, `gridlines` (show/hide), `zoom` (10–400%), `headings` (show/hide). Targets workbook window, not Application.ActiveWindow. |
| `excel_conditional_format` | **Phase 5** — `data_bar`, `color_scale` (2- or 3-color), `icon_set`, `cell_rule` (operator + fill/font color), `top_bottom` (top/bottom N items or %), `clear` |
| `excel_validation` | **Phase 5** — `list` (dropdown), `whole_number`, `decimal`, `date`, `text_length`, `custom` (formula), `clear`. Always clears existing rules before adding (COM error 1004 avoidance). |
| `excel_slicer` | **Phase 5** — `add` (Table or PivotTable slicer), `add_timeline` (date-field timeline), `list`, `delete`, `connect` (link slicer cache to additional pivots) |

### Structured references

`excel_range(action="read", range="TableName[ColumnName]")` works via Excel's
native Range() parser — no special handling needed. Documented in both
`excel_range` and `excel_table` docstrings.

## Dev commands

```powershell
uv sync                                                          # install deps
uv run pytest -q                                                 # 397 unit tests (no Excel needed)
THEPEXCEL_MCP_AUTOLAUNCH=1 uv run python tests/smoke_com.py     # full live COM smoke
THEPEXCEL_MCP_AUTOLAUNCH=1 uv run python tests/smoke_com.py --sections 1,2,3,4  # partial (sections 1-12 available)
uv run thepexcel-mcp                                            # run stdio server
uv run python scripts/build_mcpb.py                             # build dist/thepexcel-mcp.mcpb
claude mcp add thepexcel-excel --scope user -- uv run --directory D:/ThepExcelMCP thepexcel-mcp  # register
```

## Constraints

- **Windows only** — COM requires Windows + same-user Excel Desktop.
- All COM calls run on a single dedicated STA worker thread (Phase 3). Tool handlers
  submit via `_session.run_com()`. Default timeout 120s; override via `THEPEXCEL_MCP_COM_TIMEOUT`.
- Auto-launch via `THEPEXCEL_MCP_AUTOLAUNCH=1` env var (launches visible Excel).
- All COM errors are caught and re-raised as `fastmcp.exceptions.ToolError` with actionable messages.
- VBA tool opt-in: `THEPEXCEL_MCP_ENABLE_VBA=1` + Excel's AccessVBOM trust setting.

## Future phases

- **Phase 5:** Live end-to-end smoke vs real Excel · packaging (uvx + MCPB bundle for Claude Desktop) · client registration

## Session Knowledge (durable facts — tools, paths, gotchas)
> Auto-maintained by /session-handoff. Reusable facts only — NOT session status.

- **2026-06-23 — Editable install ≠ hot reload.** `claude mcp add` runs an editable install, but the *running* stdio server keeps OLD code in memory after you edit a domain/server file. To test code changes THROUGH the live MCP tools you must **restart the MCP server** (Excel COM state survives; the Python process doesn't reload). pytest + `tests/smoke_com.py` validate code WITHOUT a restart (own Excel via AUTOLAUNCH), so use those for fast iteration; restart only to verify the MCP tool surface itself.
- **2026-06-23 — `excel_datamodel(add_table)` / `powerquery(load_to_datamodel)` DEADLOCK the STA COM worker** in the Claude Code stdio context — bricks ALL subsequent calls until Excel is force-killed (not a recoverable timeout). Fallback = `load_to_table` → pivot-from-table. (Also captured in the excel-god skill.)
- **2026-06-23 — Verify EFFECT, not the success report.** Several tools returned `success` while silently doing nothing (multi-cell write wrote 1 cell; table sort was a no-op). Always read-back the actual cell/format/file after a mutating call. Root cause of the write bug: pywin32 `.Resize(r,c)` returns a single offset cell (Count=1) → use `Range(Cells, Cells)` instead.
- **2026-06-23 — Tier-2 COM API gotchas (view / validation / slicer), each a success-but-noop bug that mocked unit tests PASSED but live read-back smoke (`smoke_com.py` sections 14-16) caught — reinforces verify-EFFECT above.** (1) **`Slicers.Add`**: `Level` is OLAP-only; on non-OLAP sources (Table / regular PivotTable) it must be OMITTED — call with keyword args (`SlicerDestination=/Caption=/Top=/Left=/Width=/Height=`), else COM `E_INVALIDARG (-2147024809)`. Passing `Level=""` or `1` both fail. (2) **View settings** (freeze/gridlines/zoom/headings) must bind to `wb.Windows(1)` (workbook-scoped), NOT `Application.ActiveWindow` — ActiveWindow = the *foreground* book, so a background target gets silently mutated on the wrong workbook. (3) **`Validation.Add`** raises 1004 when a rule already exists → always `Validation.Delete()` first.
- **2026-06-23 — Workflow shared-file race.** When fanning out parallel build agents, each NEW domain module + test file is conflict-free, but a SHARED file (here `smoke_com.py`, `server.py`) edited by multiple agents at once = last-writer-wins (only the conditional_format agent's `smoke_com.py` section survived; view/validation/slicer sections had to be added in a follow-up serial pass). Serialize shared-file edits into one integrate stage / one agent.
