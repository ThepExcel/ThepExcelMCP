# ThepExcelMCP

Windows desktop MCP server that controls a **live running Excel instance** via COM (pywin32).

Gives AI agents full Excel capability — Power Query, PivotTables, Data Model/DAX, VBA, dynamic arrays — beyond what openpyxl can do because it drives the actual Excel process.

## Requirements

- Windows only (COM requires Windows + same-user Excel Desktop)
- Python 3.11+
- Microsoft Excel (desktop) must be running before tool calls

## Quick Start

```powershell
# Install
uv sync

# Run server (stdio — for MCP client config)
uv run thepexcel-mcp

# Or directly
uv run python -m thepexcel_mcp.server
```

## Claude Desktop Config

```json
{
  "mcpServers": {
    "thepexcel": {
      "command": "uv",
      "args": ["--directory", "D:/ThepExcelMCP", "run", "thepexcel-mcp"]
    }
  }
}
```

Set `THEPEXCEL_MCP_AUTOLAUNCH=1` to auto-launch Excel if not running.

## Tools (Phase 3 — current)

| Tool | Actions |
|------|---------|
| `excel_workbook` | list, info, open, save, close |
| `excel_sheet` | list, add, rename, delete |
| `excel_range` | read (paginated + spill metadata), read_spill, write, write_formula, clear |
| `excel_powerquery` | list, get, create, update, delete, refresh, refresh_all, load_to_table, load_to_datamodel, analyze, analyze_raw |
| `excel_table` | list, create, read, append_rows, add_column, sort, filter, set_style, toggle_totals, rename, delete |
| `excel_pivot` | list, create, add_field, remove_field, move_field, set_layout, refresh, delete, read |
| `excel_datamodel` | info, list_tables, add_table, list_relationships, add_relationship, delete_relationship, list_measures, add_measure, update_measure, delete_measure, refresh |
| `excel_vba` | list_modules, get_module, write_module, delete_module, run |
| `excel_name` | list, get, set, delete (named ranges + LAMBDA formulas) |

### Optional env vars

| Variable | Default | Description |
|---|---|---|
| `THEPEXCEL_MCP_AUTOLAUNCH` | unset | Set to `1` to auto-launch Excel |
| `THEPEXCEL_MCP_ENABLE_VBA` | unset | Set to `1` to enable `excel_vba` tool |
| `THEPEXCEL_MCP_COM_TIMEOUT` | `120` | Per-call COM timeout in seconds |

### VBA setup

`excel_vba` requires both:
1. `THEPEXCEL_MCP_ENABLE_VBA=1` environment variable
2. Excel: File → Options → Trust Center → Trust Center Settings → Macro Settings → "Trust access to the VBA project object model"

## Development

```powershell
uv sync
uv run pytest                     # 167 unit tests (no Excel needed)
uv run python tests/smoke_com.py  # manual COM smoke (requires live Excel)
```
