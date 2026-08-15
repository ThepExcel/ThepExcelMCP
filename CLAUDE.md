---
purpose: |
  Let me control Excel with AI without limits — going far beyond Copilot, Claude-in-Excel or any official tool — so the Excel work behind my teaching and projects has no ceiling, and I stay the person known for mastering Excel more deeply than anyone.
---

# CLAUDE.md — ThepExcelMCP

> Windows-only MCP server (stdio). Controls live Excel Desktop via COM (pywin32) —
> 26 action-dispatch tools giving AI agents full Excel capability beyond openpyxl.

Maintainer test discipline (units-green ≠ effect for COM tools):
`.claude/skills/thepexcelmcp-verify-effect-testing/SKILL.md` — re-verify any section its
banner flags PARTIAL before relying on it.

## Architecture

Client ─stdio─ FastMCP (`src/thepexcel_mcp/server.py`, 26 tool registrations) →
`ExcelSession` (`src/thepexcel_mcp/session.py`) → `Excel.Application` COM → running Excel
Desktop. All COM calls run on one dedicated STA worker thread; handlers submit via
`_session.run_com(fn)` (blocks, 120s default, `THEPEXCEL_MCP_COM_TIMEOUT`). Helpers:
`excel_guard(app)` (DisplayAlerts=False around risky ops), `wait_calculation(app)`,
ROT fallback in `get_workbook()` for multiple Excel instances.

## Tools → files → actions

All domain modules live in `src/thepexcel_mcp/domains/`. Full per-action semantics are in
each tool's docstring (server.py); API landmines in `docs/com-gotchas.md`.

| Tool | File | Actions |
|---|---|---|
| `excel_workbook` | workbook.py | list, info, open, save, close, create, save_as |
| `excel_sheet` | sheets.py | list, add, rename, delete |
| `excel_range` | ranges.py | read (paginated, spill metadata), read_spill, write, write_formula (Formula2), write_py (experimental `=PY()`), clear |
| `excel_powerquery` | powerquery.py + analysis/pq_analyzer.py | list, get, create, update, delete, refresh, refresh_all, load_to_table, load_to_datamodel, analyze, analyze_raw, create/get/set/list_parameters |
| `excel_table` | tables.py | list, create, read, append_rows, add_column, sort, filter, set_style, toggle_totals, rename, delete |
| `excel_pivot` | pivots.py | list, create (range/table/datamodel), add_field, remove_field, move_field, set_layout, refresh, delete, read |
| `excel_datamodel` | datamodel.py | info, list_tables, add_table, list/add/delete_relationship, list/add/update/delete_measure, refresh, cube_value, cube_member, cube_formula, add_calculated_column/table (COM-impossible → guard ToolError) |
| `excel_vba` | vba.py | list_modules, get_module, write_module, delete_module, run (opt-in) |
| `excel_name` | names.py | list, get, set, delete (named ranges, constants, LAMBDA) |
| `excel_format` | format.py | font, fill, border, number_format, alignment, column_width, row_height, autofit (colors `"#RRGGBB"`) |
| `excel_chart` | charts.py | list, create, configure, set_source, export_image, delete |
| `excel_screenshot` | screenshot.py | range, sheet, chart (PNG for visual verification) |
| `excel_view` | view.py | freeze_panes, unfreeze_panes, gridlines, zoom, headings |
| `excel_conditional_format` | conditional_format.py | data_bar, color_scale, icon_set, cell_rule, top_bottom, clear |
| `excel_validation` | validation.py | list, whole_number, decimal, date, text_length, custom, clear |
| `excel_slicer` | slicer.py | add, add_timeline, list, delete, connect |
| `excel_page_setup` | page_setup.py | set, print_area, print_titles, header_footer, export_pdf, get |
| `excel_comment` | comments.py | add, edit, reply, delete, list, get (kind=note\|threaded\|all) |
| `excel_hyperlink` | hyperlinks.py | add (url/internal/email/file), list, delete |
| `excel_outline` | outline.py | group/ungroup rows & columns, show_levels, clear |
| `excel_protection` | protection.py | protect/unprotect sheet & workbook, set_locked, status |
| `excel_sparkline` | sparkline.py | add, clear, list |
| `excel_shape` | shapes.py | add_image, add_textbox, add_shape, list, delete, move |
| `excel_find_replace` | find_replace.py | find, count, replace (scope=range\|sheet\|workbook) |
| `excel_diff` | diff.py | ranges, sheets (pure read) |
| `excel_snapshot` | snapshot.py | snapshot, list, restore, delete (SaveCopyAs; restore opens a NEW workbook — non-destructive by design) |

### load_to_table pattern (do NOT simplify)

`powerquery.py::_load_to_table` uses a specific Connections.Add2 + Mashup-OLEDB sequence:
`lCmdtype=2` (xlCmdSql) · `Location=<query_name>` WITHOUT quotes in the connection string
· `CommandText = "SELECT * FROM [<query_name>]"`. This exact form is required — it took
real debugging to get right.

Structured references work natively: `excel_range(action="read",
range="TableName[ColumnName]")` — no special handling.

## Dev commands

```powershell
uv sync --frozen                                                 # install exact locked deps
uv run pytest -q                                                 # unit tests (no Excel needed)
uv run ruff check src scripts --select E9,F                     # static checks
THEPEXCEL_MCP_AUTOLAUNCH=1 uv run python tests/smoke_com.py     # full live COM smoke
THEPEXCEL_MCP_AUTOLAUNCH=1 uv run python tests/smoke_com.py --sections 1,2,3,4  # partial (sections 1-28)
THEPEXCEL_MCP_AUTOLAUNCH=1 uv run python tests/bench_com.py     # live perf bench (stderr; not pytest-collected)
uv run --isolated --with mcp==2.0.0 python tests/protocol_smoke_v2.py  # real stdio handshake
uv run thepexcel-mcp                                            # run stdio server
uv run python scripts/build_mcpb.py                             # build dist/thepexcel-mcp.mcpb
claude mcp add thepexcel-excel --scope user -- uv run --directory D:/ThepExcelMCP thepexcel-mcp  # register
```

## Constraints

- **Windows only** — COM requires Windows + same-user Excel Desktop.
- **Never print to stdout** — stdout is the stdio transport. Diagnostics go to stderr
  (slow COM calls over `THEPEXCEL_MCP_SLOW_LOG_S` seconds, default 5.0, already do).
- Auto-launch is ON by default (`get_app()` opens a visible Excel + blank workbook if
  none is running); opt out with a falsy `THEPEXCEL_MCP_AUTOLAUNCH`. It also self-heals
  a corrupt win32com gen_py cache.
- Perf layer: `get_app()` caches the Application handle; paginated reads marshal only the
  page-bounded sub-range; `bulk_guard(app)` around bulk write loops — NEVER around PQ
  refresh / datamodel / cube paths (deadlock + async-calc). `THEPEXCEL_MCP_EARLYBIND=1`
  (default ON) re-wraps the same IDispatch early-bound; `=0` is the kill switch.
  `value_mode="raw"` gives faster `.Value2` reads. Details: `docs/com-gotchas.md`
  § Performance layer.
- All COM errors are re-raised as `fastmcp.exceptions.ToolError` with actionable messages.
- VBA tool opt-in: `THEPEXCEL_MCP_ENABLE_VBA=1` + Excel's AccessVBOM trust setting.
- Progressive tool discovery opt-in: `THEPEXCEL_MCP_TOOL_DISCOVERY=bm25` hides the
  catalog behind `search_tools`/`call_tool`; default `full`.
- **Public repo — synthetic data only (HARDLINE).** `samples/`, tests, docstrings, and
  doc examples must use synthetic/anonymized data: never a real customer or company
  name, product catalog, model codes, client figures, or third-party business data
  ("Alpha/Beta", "Widget/Gadget", "North/South"). Keep internal working notes and
  private-tooling vocabulary out of committed files.

## Known traps (read before touching COM code)

- `excel_datamodel(add_table)` / `excel_powerquery(load_to_datamodel)` can **DEADLOCK
  the STA worker** in the headless stdio context — bricks ALL subsequent calls until
  Excel is force-killed. Fallback: `load_to_table` → pivot-from-table.
- Editable install ≠ hot reload — restart the MCP server to test tool-surface changes;
  pytest + `tests/smoke_com.py` validate code without a restart.
- **Verify EFFECT, not the success report** — read back the actual cell/format/file
  after every mutating call; several COM APIs return success while doing nothing.
- `get_app()` attaches to the FIRST ROT Excel instance — when testing, create and
  target an explicit test workbook; never assume the active workbook is disposable.
- Excel tools silently vanish from the client's tool list → suspect MCP protocol
  version skew, NOT Excel: `docs/mcp-rc-2026-07-28-migration.md`.
- Full per-API gotcha library (data model, PQ parameters, per-tool landmines, perf
  facts, testing discipline): `docs/com-gotchas.md`.

Roadmap: all planned phases and tiers have shipped — history in `docs/ROADMAP.md`.
