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
      │ 9 action-dispatch tools
      ├── excel_workbook   → domains/workbook.py
      ├── excel_sheet      → domains/sheets.py
      ├── excel_range      → domains/ranges.py         (+ read_spill, spill metadata)
      ├── excel_powerquery → domains/powerquery.py + analysis/pq_analyzer.py
      ├── excel_table      → domains/tables.py
      ├── excel_pivot      → domains/pivots.py
      ├── excel_datamodel  → domains/datamodel.py      ← Phase 2
      ├── excel_vba        → domains/vba.py            ← Phase 3
      ├── excel_name       → domains/names.py          ← Phase 3
      ├── excel_chart      → domains/charts.py         ← Phase 4
      └── excel_screenshot → domains/screenshot.py     ← Phase 4
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
| `src/thepexcel_mcp/server.py` | FastMCP app + 9 tool registrations |
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
| `src/thepexcel_mcp/domains/charts.py` | **Phase 4** — Chart CRUD + configure + export PNG |
| `src/thepexcel_mcp/domains/screenshot.py` | **Phase 4** — Range/sheet/chart capture as PNG (CopyPicture+PIL) |
| `src/thepexcel_mcp/analysis/pq_analyzer.py` | M code static analyzer — copied verbatim from PoC |

### load_to_table pattern (important)

`powerquery.py::_load_to_table` uses a specific Connections.Add2 + Mashup-OLEDB sequence
to load a query into a worksheet Table. This exact form is required:
- `lCmdtype=2` (xlCmdSql)
- `Location=<query_name>` WITHOUT quotes in the connection string
- `CommandText = "SELECT * FROM [<query_name>]"`

Do NOT simplify this — it took real debugging in the PoC to get right.

## Tool registry (Phase 1 + 2 + 3)

| Tool | Actions |
|---|---|
| `excel_workbook` | `list`, `info`, `open`, `save`, `close` |
| `excel_sheet` | `list`, `add`, `rename`, `delete` |
| `excel_range` | `read` (paginated, 100-row default; spill metadata in response), `read_spill` (full spill range for dynamic array anchor), `write`, `write_formula` (Formula2/dynamic arrays), `write_py` (**Phase 4** — `=PY()` Formula2R1C1 insertion, experimental), `clear` |
| `excel_powerquery` | `list`, `get`, `create`, `update`, `delete`, `refresh`, `refresh_all`, `load_to_table`, `load_to_datamodel` (**Phase 2**), `analyze`, `analyze_raw`. `create`/`update` return `warnings` list from M code analyzer (non-blocking). `list` includes `model_connection` flag. |
| `excel_table` | `list`, `create`, `read` (paginated), `append_rows`, `add_column` (with formula), `sort`, `filter` (equals/contains/greater/less/clear_filters), `set_style`, `toggle_totals` (per-column aggregation), `rename`, `delete` (keep_data/unlist) |
| `excel_pivot` | `list`, `create` (range/table/datamodel source), `add_field` (with aggregation+number_format; CubeField fallback for datamodel pivots **Phase 2**), `remove_field`, `move_field`, `set_layout` (compact/outline/tabular + subtotals + grand_totals), `refresh`, `delete`, `read` (paginated via TableRange1) |
| `excel_datamodel` | **Phase 2** — `info`, `list_tables`, `add_table` (table\|query → model via CreateModelConnection=True), `list_relationships`, `add_relationship`, `delete_relationship`, `list_measures`, `add_measure` (DAX + format), `update_measure`, `delete_measure`, `refresh` |
| `excel_vba` | **Phase 3** — `list_modules`, `get_module`, `write_module`, `delete_module`, `run`. Opt-in: `THEPEXCEL_MCP_ENABLE_VBA=1` + AccessVBOM registry check per call. |
| `excel_name` | **Phase 3** — `list`, `get`, `set`, `delete`. Covers named ranges, constants, and LAMBDA formulas. `is_lambda` flag on list/get. |
| `excel_chart` | **Phase 4** — `list`, `create`, `configure`, `set_source`, `export_image`, `delete` |
| `excel_screenshot` | **Phase 4** — `range`, `sheet`, `chart` (PNG capture for LLM visual verification) |

### Structured references

`excel_range(action="read", range="TableName[ColumnName]")` works via Excel's
native Range() parser — no special handling needed. Documented in both
`excel_range` and `excel_table` docstrings.

## Dev commands

```powershell
uv sync                                                          # install deps
uv run pytest -q                                                 # 192 unit tests (no Excel needed)
THEPEXCEL_MCP_AUTOLAUNCH=1 uv run python tests/smoke_com.py     # full live COM smoke
THEPEXCEL_MCP_AUTOLAUNCH=1 uv run python tests/smoke_com.py --sections 1,2,3,4  # partial
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
