# CLAUDE.md — ThepExcelMCP

> Windows-only MCP server. Controls live Excel Desktop via COM (pywin32).
> Gives AI agents full Excel capability beyond what openpyxl can do.

## Architecture

```
Claude / MCP Client
      │ stdio
FastMCP server (server.py)
      │ 4 action-dispatch tools
      ├── excel_workbook  → domains/workbook.py
      ├── excel_sheet     → domains/sheets.py
      ├── excel_range     → domains/ranges.py
      └── excel_powerquery→ domains/powerquery.py + analysis/pq_analyzer.py
                                          │
                                    ExcelSession (session.py)
                                          │ win32com
                                    Excel.Application (COM)
                                          │
                                    Running Excel Desktop
```

### Key files

| File | Role |
|---|---|
| `src/thepexcel_mcp/server.py` | FastMCP app + 4 tool registrations |
| `src/thepexcel_mcp/session.py` | `ExcelSession` — COM attach/launch, workbook/sheet routing, `ToolError` wrapping |
| `src/thepexcel_mcp/domains/workbook.py` | Workbook CRUD |
| `src/thepexcel_mcp/domains/sheets.py` | Sheet CRUD |
| `src/thepexcel_mcp/domains/ranges.py` | Range read (paginated) / write / write_formula (Formula2) / clear |
| `src/thepexcel_mcp/domains/powerquery.py` | 10 PQ operations — create/update include M code warnings |
| `src/thepexcel_mcp/domains/tables.py` | 11 Table (ListObject) operations |
| `src/thepexcel_mcp/domains/pivots.py` | 9 PivotTable operations |
| `src/thepexcel_mcp/analysis/pq_analyzer.py` | M code static analyzer — copied verbatim from PoC |

### load_to_table pattern (important)

`powerquery.py::_load_to_table` uses a specific Connections.Add2 + Mashup-OLEDB sequence
to load a query into a worksheet Table. This exact form is required:
- `lCmdtype=2` (xlCmdSql)
- `Location=<query_name>` WITHOUT quotes in the connection string
- `CommandText = "SELECT * FROM [<query_name>]"`

Do NOT simplify this — it took real debugging in the PoC to get right.

## Tool registry (Phase 1)

| Tool | Actions |
|---|---|
| `excel_workbook` | `list`, `info`, `open`, `save`, `close` |
| `excel_sheet` | `list`, `add`, `rename`, `delete` |
| `excel_range` | `read` (paginated, 100-row default), `write`, `write_formula` (Formula2/dynamic arrays), `clear` |
| `excel_powerquery` | `list`, `get`, `create`, `update`, `delete`, `refresh`, `refresh_all`, `load_to_table`, `analyze`, `analyze_raw`. `create`/`update` return `warnings` list from M code analyzer (non-blocking). |
| `excel_table` | `list`, `create`, `read` (paginated), `append_rows`, `add_column` (with formula), `sort`, `filter` (equals/contains/greater/less/clear_filters), `set_style`, `toggle_totals` (per-column aggregation), `rename`, `delete` (keep_data/unlist) |
| `excel_pivot` | `list`, `create` (range/table/datamodel source), `add_field` (with aggregation+number_format), `remove_field`, `move_field`, `set_layout` (compact/outline/tabular + subtotals + grand_totals), `refresh`, `delete`, `read` (paginated via TableRange1) |

### Structured references

`excel_range(action="read", range="TableName[ColumnName]")` works via Excel's
native Range() parser — no special handling needed. Documented in both
`excel_range` and `excel_table` docstrings.

## Dev commands

```powershell
uv sync                          # install deps
uv run pytest                    # unit tests (no Excel needed)
uv run python tests/smoke_com.py # manual COM smoke (requires live Excel)
uv run thepexcel-mcp             # run stdio server
uv run python -c "from thepexcel_mcp.server import mcp; print('OK')"  # import check
```

## Constraints

- **Windows only** — COM requires Windows + same-user Excel Desktop.
- `pythoncom.CoInitialize()` is called per tool call (FastMCP may use multiple threads).
- Auto-launch via `THEPEXCEL_MCP_AUTOLAUNCH=1` env var (launches visible Excel).
- All COM errors are caught and re-raised as `fastmcp.exceptions.ToolError` with actionable messages.

## Future phases

- **Phase 2:** Data Model / DAX (measures, relationships via `Model.ModelMeasures`)
- **Phase 3:** VBA execution (`Application.Run`)
- **Phase 4:** Charts
