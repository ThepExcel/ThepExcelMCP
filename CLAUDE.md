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
      │ 20 action-dispatch tools
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
      ├── excel_slicer              → domains/slicer.py         ← Phase 5
      ├── excel_page_setup          → domains/page_setup.py     ← Tier-4
      ├── excel_comment             → domains/comments.py       ← Tier-4
      ├── excel_hyperlink           → domains/hyperlinks.py     ← Tier-4
      ├── excel_outline             → domains/outline.py        ← Tier-4
      ├── excel_protection          → domains/protection.py     ← Tier-4
      └── excel_sparkline           → domains/sparkline.py      ← Tier-4
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
| `src/thepexcel_mcp/server.py` | FastMCP app + 20 tool registrations |
| `src/thepexcel_mcp/session.py` | `ExcelSession` — STA worker thread, `run_com()`, `excel_guard`, `wait_calculation`, ROT fallback |
| `src/thepexcel_mcp/domains/workbook.py` | Workbook CRUD |
| `src/thepexcel_mcp/domains/sheets.py` | Sheet CRUD |
| `src/thepexcel_mcp/domains/ranges.py` | Range read (paginated, spill metadata) / read_spill / write / write_formula (Formula2) / clear |
| `src/thepexcel_mcp/domains/powerquery.py` | 11 PQ operations + 4 parameter ops (Tier-3: create/get/set/list_parameters, `is_parameter`/`is_function` flags) — create/update include M code warnings; `load_to_datamodel` Phase 2 |
| `src/thepexcel_mcp/domains/tables.py` | 11 Table (ListObject) operations |
| `src/thepexcel_mcp/domains/pivots.py` | 9 PivotTable operations; data-model CubeField support in `add_field` |
| `src/thepexcel_mcp/domains/datamodel.py` | **Phase 2** — 11 Data Model ops (model tables, relationships, DAX measures) + Tier-3 cube formulas (`cube_value`/`cube_member`/`cube_formula`) + calc-column/table guards |
| `src/thepexcel_mcp/domains/vba.py` | **Phase 3** — VBA module CRUD + macro execution (opt-in via env + AccessVBOM check) |
| `src/thepexcel_mcp/domains/names.py` | **Phase 3** — Named ranges, LAMBDA formulas, defined name CRUD |
| `src/thepexcel_mcp/domains/format.py` | **Phase 5** — Range formatting: font/fill/border/number_format/alignment/sizing |
| `src/thepexcel_mcp/domains/charts.py` | **Phase 4** — Chart CRUD + configure + export PNG |
| `src/thepexcel_mcp/domains/screenshot.py` | **Phase 4** — Range/sheet/chart capture as PNG (CopyPicture+PIL) |
| `src/thepexcel_mcp/domains/view.py` | **Phase 5** — Worksheet display: freeze panes, gridlines, zoom, headings |
| `src/thepexcel_mcp/domains/conditional_format.py` | **Phase 5** — Conditional formatting: cell_rule, data_bar, color_scale, icon_set, top_bottom, clear |
| `src/thepexcel_mcp/domains/validation.py` | **Phase 5** — Data validation: list, whole_number, decimal, date, text_length, custom, clear |
| `src/thepexcel_mcp/domains/slicer.py` | **Phase 5** — Slicers + timelines: add, add_timeline, list, delete, connect |
| `src/thepexcel_mcp/domains/page_setup.py` | **Tier-4** — Page setup: orientation/paper/margins/fit-to-page/scale, print_area, print_titles, header_footer, **export_pdf** (ExportAsFixedFormat), get |
| `src/thepexcel_mcp/domains/comments.py` | **Tier-4** — Cell comments: legacy notes + threaded comments (add/edit/reply/delete/list/get); `kind=note\|threaded\|all` |
| `src/thepexcel_mcp/domains/hyperlinks.py` | **Tier-4** — Hyperlinks: add (url/internal/email/file), list, delete (keeps cell text) |
| `src/thepexcel_mcp/domains/outline.py` | **Tier-4** — Row/column grouping: group/ungroup rows&cols, show_levels, clear |
| `src/thepexcel_mcp/domains/protection.py` | **Tier-4** — Protection: protect/unprotect sheet&workbook, set_locked (cell lock/FormulaHidden), status |
| `src/thepexcel_mcp/domains/sparkline.py` | **Tier-4** — In-cell sparklines: add (line/column/win_loss), clear, list; ReferenceStyle=xlA1 guard |
| `src/thepexcel_mcp/analysis/pq_analyzer.py` | M code static analyzer — copied verbatim from PoC |

### load_to_table pattern (important)

`powerquery.py::_load_to_table` uses a specific Connections.Add2 + Mashup-OLEDB sequence
to load a query into a worksheet Table. This exact form is required:
- `lCmdtype=2` (xlCmdSql)
- `Location=<query_name>` WITHOUT quotes in the connection string
- `CommandText = "SELECT * FROM [<query_name>]"`

Do NOT simplify this — it took real debugging in the PoC to get right.

## Tool registry (Phase 1 + 2 + 3 + 4 + 5 + Tier-4)

| Tool | Actions |
|---|---|
| `excel_workbook` | `list`, `info`, `open`, `save`, `close`, `create` (Workbooks.Add + optional SaveAs), `save_as` (FileFormat inferred from extension) |
| `excel_sheet` | `list`, `add`, `rename`, `delete` |
| `excel_range` | `read` (paginated, 100-row default; spill metadata in response), `read_spill` (full spill range for dynamic array anchor), `write`, `write_formula` (Formula2/dynamic arrays), `write_py` (**Phase 4** — `=PY()` Formula2R1C1 insertion, experimental), `clear` |
| `excel_powerquery` | `list`, `get`, `create`, `update`, `delete`, `refresh`, `refresh_all`, `load_to_table`, `load_to_datamodel` (**Phase 2**), `analyze`, `analyze_raw`, `create_parameter`, `get_parameter`, `set_parameter`, `list_parameters` (**Tier-3**). `create`/`update` return `warnings` list from M code analyzer (non-blocking). `list` includes `model_connection` + `is_parameter`/`is_function` flags. PQ parameters = literal + meta-record M via pure `wb.Queries` (no deadlock surface); value swap preserves the meta record verbatim. |
| `excel_table` | `list`, `create`, `read` (paginated), `append_rows`, `add_column` (with formula), `sort`, `filter` (equals/contains/greater/less/clear_filters), `set_style`, `toggle_totals` (per-column aggregation), `rename`, `delete` (keep_data/unlist) |
| `excel_pivot` | `list`, `create` (range/table/datamodel source), `add_field` (with aggregation+number_format; CubeField fallback for datamodel pivots **Phase 2**), `remove_field`, `move_field`, `set_layout` (compact/outline/tabular + subtotals + grand_totals), `refresh`, `delete`, `read` (paginated via TableRange1) |
| `excel_datamodel` | **Phase 2** — `info`, `list_tables`, `add_table` (table\|query → model via CreateModelConnection=True), `list_relationships`, `add_relationship`, `delete_relationship`, `list_measures`, `add_measure` (DAX + format), `update_measure`, `delete_measure`, `refresh`. **Tier-3** — `cube_value`/`cube_member` (build + write CUBEVALUE/CUBEMEMBER to a cell, async-resolve + verify-EFFECT; numeric resolution needs an existing model), `cube_formula` (build-only string, no COM), `add_calculated_column`/`add_calculated_table` (guards — COM-impossible per Microsoft Learn; raise actionable ToolError pointing to the Power Query workaround) |
| `excel_vba` | **Phase 3** — `list_modules`, `get_module`, `write_module`, `delete_module`, `run`. Opt-in: `THEPEXCEL_MCP_ENABLE_VBA=1` + AccessVBOM registry check per call. |
| `excel_name` | **Phase 3** — `list`, `get`, `set`, `delete`. Covers named ranges, constants, and LAMBDA formulas. `is_lambda` flag on list/get. |
| `excel_format` | **Phase 5** — `font` (name/size/bold/italic/underline/color), `fill` (bg color or clear), `border` (sides: all/outline/inside/top/bottom/left/right; style: continuous/dash/double/none; weight: thin/medium/thick), `number_format` (any Excel code), `alignment` (h/v align, wrap_text, merge/unmerge), `column_width`, `row_height`, `autofit`. All colors as `"#RRGGBB"` → converted to Excel BGR internally. |
| `excel_chart` | **Phase 4** — `list`, `create`, `configure`, `set_source`, `export_image`, `delete` |
| `excel_screenshot` | **Phase 4** — `range`, `sheet`, `chart` (PNG capture for LLM visual verification) |
| `excel_view` | **Phase 5** — `freeze_panes`, `unfreeze_panes`, `gridlines` (show/hide), `zoom` (10–400%), `headings` (show/hide). Targets workbook window, not Application.ActiveWindow. |
| `excel_conditional_format` | **Phase 5** — `data_bar`, `color_scale` (2- or 3-color), `icon_set`, `cell_rule` (operator + fill/font color), `top_bottom` (top/bottom N items or %), `clear` |
| `excel_validation` | **Phase 5** — `list` (dropdown), `whole_number`, `decimal`, `date`, `text_length`, `custom` (formula), `clear`. Always clears existing rules before adding (COM error 1004 avoidance). |
| `excel_slicer` | **Phase 5** — `add` (Table or PivotTable slicer), `add_timeline` (date-field timeline), `list`, `delete`, `connect` (link slicer cache to additional pivots) |
| `excel_page_setup` | **Tier-4** — `set` (orientation/paper_size/fit_to_wide+tall/scale/margins-in-inches/center/print_gridlines/black_and_white), `print_area`, `print_titles` (rows/cols), `header_footer` (L/C/R header+footer, &P &N &D &T &F &A codes), `export_pdf` (scope=sheet\|workbook, verify file-on-disk), `get`. Margins converted via InchesToPoints; fit-to-page sets Zoom=False first. |
| `excel_comment` | **Tier-4** — `add`, `edit` (note only), `reply` (threaded only), `delete`, `list`, `get`. `kind="note"` (legacy ws.Comments) or `"threaded"` (ws.CommentsThreaded); `kind="all"` for list/get. Text() is a METHOD; delete-before-add guard; verify-effect read-back on add/delete. |
| `excel_hyperlink` | **Tier-4** — `add` (link_type=url\|internal\|email\|file; internal needs Address=""), `list`, `delete` (removes link, keeps cell text). Anchor must be a Range; read-back of `.Address`/`.SubAddress`/`.ScreenTip`. |
| `excel_outline` | **Tier-4** — `group_rows`, `group_columns`, `ungroup_rows`, `ungroup_columns` (rows="2:5"/columns="B:D"), `show_levels` (row_levels/column_levels 1–8), `clear`. Read-back via OutlineLevel (ungrouped=1, grouped≥2); max 8 levels. |
| `excel_protection` | **Tier-4** — `protect_sheet` (+ pass-through allow_* flags or `allow` dict), `unprotect_sheet`, `protect_workbook` (structure/windows), `unprotect_workbook`, `set_locked` (Locked/FormulaHidden — only effective under protection), `status`. Password omitted when None; wrong-password 1004 → ToolError; verify-effect via ProtectContents/ProtectStructure. |
| `excel_sparkline` | **Tier-4** — `add` (location=destination, data_range=source, spark_type=line\|column\|win_loss, optional marker/color), `clear`, `list`. `location` required all actions. Application.ReferenceStyle forced to xlA1 around Add (saved/restored); color "#RRGGBB"→BGR; verify-effect via SparklineGroups.Count. |

### Structured references

`excel_range(action="read", range="TableName[ColumnName]")` works via Excel's
native Range() parser — no special handling needed. Documented in both
`excel_range` and `excel_table` docstrings.

## Dev commands

```powershell
uv sync                                                          # install deps
uv run pytest -q                                                 # 758 unit tests (no Excel needed)
THEPEXCEL_MCP_AUTOLAUNCH=1 uv run python tests/smoke_com.py     # full live COM smoke
THEPEXCEL_MCP_AUTOLAUNCH=1 uv run python tests/smoke_com.py --sections 1,2,3,4  # partial (sections 1-24 available)
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
- **2026-06-24 — MCP RC 2026-07-28 ("sessionless core") needs NO code migration here; risk is operational only.** This server is stdio-only (`mcp.run(transport="stdio")`, server.py:1614); `Mcp-Session-Id` is HTTP-transport-only (SEP-2567) and was never used — all protocol/handshake handling lives inside `fastmcp` (grep `src/` for session/initialize = 0). The ONE real risk is a time-gated client/server version-skew window (~late-Jul→Sep 2026): if a sessionless-only Claude client ships before an RC-ready `fastmcp` is `uv sync`'d + the stdio process restarted, **the Excel tools silently vanish from Claude's tool list — and it looks NOTHING like a COM/Excel error.** Diagnostic: tools disappear post-July-2026 → suspect MCP protocol skew, NOT Excel → `uv sync` an RC-ready fastmcp + restart. Watch fastmcp releases from Jun 30 / Jul 27; buy signal = notes naming "2026-07-28 / SEP-2575 / sessionless". `smoke_com.py` does NOT exercise the protocol layer (owns its own Excel) — needs a manual client-handshake check after any fastmcp bump. Full plan + runbook: `docs/mcp-rc-2026-07-28-migration.md`.
- **2026-06-24 — Tier-3 COM facts (PQ parameters · cube formulas · calc-column guards; shipped `3e1329a`).** (1) **PQ parameter** = a normal query whose M is `<literal> meta [IsParameterQuery=true, Type="Number"|"Text", IsParameterQueryRequired=...]`; `Queries.Add` stores M verbatim so it round-trips — pure `wb.Queries`, **no data-model/deadlock surface**. GOTCHA: parse/rewrite must split the leading literal from the meta record **structurally** (walk the quoted string honoring `""` escapes), NEVER by string-searching `' meta '` — a Text value can itself contain `' meta '` and corrupt the formula (`_split_literal_and_meta` in powerquery.py). Value guard = `value is None`, not `if not value` (0/`''`/False are valid params). (2) **Cube formulas** (`cube_value`/`cube_member`) write CUBEVALUE/CUBEMEMBER then `Application.CalculateUntilAsyncQueriesDone()` to resolve `#GETTING_DATA`; **numeric resolution needs an EXISTING in-workbook model + the named measure** (headless model-build deadlocks) — `cube_formula` is build-only (no COM). Connection literal = `"ThisWorkbookDataModel"`. (3) **DAX calculated column/table = impossible via COM** (`ModelTableColumns` READ-ONLY, no `.Add`; calc table = Power BI/AS only — Microsoft Learn); `add_calculated_column`/`add_calculated_table` are pure-Python guards raising a ToolError to the Power Query workaround.
- **2026-06-24 — Tier-4 COM facts (page_setup · comment · hyperlink · outline · protection · sparkline = 6 new tools, 14→20; shipped this session). 271 new unit tests (758 total) + live smoke §19-24 = 37/37 PASS (real read-back).** (1) **Sparkline**: `Range.SparklineGroups.Add(Type, SourceData)` REQUIRES `Application.ReferenceStyle = xlA1 (1)` — save+force+restore around the call or it can error. `XlSparkType`: line=1, column=2, win_loss=`xlSparkColumnStacked100`=3. **`_list` MUST scope to the LOCATION range, never `UsedRange`** — sparkline destination cells are value-EMPTY so UsedRange excludes them → `UsedRange.SparklineGroups.Count`=0 while `Range(location).SparklineGroups.Count`=1 (live-smoke-caught bug units missed). color="#RRGGBB"→BGR. (2) **Comments — two separate systems**: legacy NOTES (`cell.Comment`/`AddComment`, `.Author` is a plain string) vs THREADED (`cell.CommentThreaded`/`AddCommentThreaded`, `.Author.Name` is the string, `.AddReply(text)`). `Comment.Text()`/`CommentThreaded.Text()` are **METHODS** — call `.Text()` to read, `.Text(x)` to write, never `.Text` attribute. Delete-before-add (AddComment raises 1004 if one exists). Threaded text NOT editable via COM → `edit` is note-only. Threaded comments DO work on Win11 Excel desktop (not gated). (3) **Page setup**: margins are in POINTS → `Application.InchesToPoints()`. Fit-to-page needs `Zoom = False` set BEFORE `FitToPagesWide/Tall`. `PrintArea` normalizes to absolute `$A$1:$F$50`. `ExportAsFixedFormat(Type=0 PDF, Filename, ...)` on ws (sheet) or wb (workbook) — verify file-on-disk (`os.path.exists`+size); "Microsoft Print to PDF" satisfies the driver requirement. **`XlPaperSize`: A4=9, Letter=1, A3=8, Legal=5 — WebSearch HALLUCINATED these (A4=2 etc.); only a direct MS-Learn WebFetch was correct** (G-FRESHNESS: verify enums at source, not search snippets). (4) **Protection**: `protect_workbook` verify-effect guard must be `if structure and not wb.ProtectStructure` (NOT unconditional) — else a legit `structure=False` call always raises (red-team-caught correctness bug). `Password=None` must OMIT the keyword (passing `""` can raise 1004); wrong-password Unprotect → 1004. `set_locked` (`.Locked`/`.FormulaHidden`) only effective under sheet protection. (5) **Outline**: `ws.Rows("2:5").Group()` / `ws.Columns("B:D").Group()`; OutlineLevel ungrouped=**1**, grouped≥2 (not 0); max 8 levels; `ws.Cells.ClearOutline()`. (6) **Process win**: build→adversarial-verify pipeline (6 build + 6 red-team agents) + **serial single-agent integration** (server.py/smoke shared files) = NO shared-file race this time (Tier-2 lesson applied). Red-team caught real defects (protection guard, multiple success-but-noop test weaknesses); live smoke caught the one bug units structurally couldn't (sparkline UsedRange scope) — both layers earned their keep.
