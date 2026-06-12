# ThepExcelMCP

Windows desktop MCP server that controls a **live running Excel instance** via COM (pywin32).

Gives AI agents full Excel capability — Power Query, PivotTables, Data Model/DAX, VBA, dynamic arrays, charts, screenshots — beyond what openpyxl can do because it drives the actual Excel process.

## Requirements

- Windows 10/11 only (COM requires Windows + same-user Excel Desktop)
- Python 3.11+ with [uv](https://docs.astral.sh/uv/) installed
- Microsoft 365 / Excel desktop (must be running before first tool call, or set `THEPEXCEL_MCP_AUTOLAUNCH=1`)

## Install

```powershell
cd D:\ThepExcelMCP
uv sync
```

That's it — no pip, no virtualenv setup. `uv sync` installs `fastmcp`, `pywin32`, and `pillow` into an isolated env.

## Register with Claude Code (CLI)

```powershell
claude mcp add thepexcel-excel --scope user -- uv run --directory D:\ThepExcelMCP thepexcel-mcp
```

Verify it appears and starts cleanly:

```powershell
claude mcp list
```

To enable VBA and auto-launch Excel:

```powershell
claude mcp add thepexcel-excel --scope user `
  -e THEPEXCEL_MCP_AUTOLAUNCH=1 `
  -e THEPEXCEL_MCP_ENABLE_VBA=1 `
  -- uv run --directory D:\ThepExcelMCP thepexcel-mcp
```

## Register with Claude Desktop

Edit `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "thepexcel-excel": {
      "command": "uv",
      "args": ["run", "--directory", "D:\\ThepExcelMCP", "thepexcel-mcp"],
      "env": {
        "THEPEXCEL_MCP_AUTOLAUNCH": "1"
      }
    }
  }
}
```

Restart Claude Desktop after saving. The server starts on demand when you first call an Excel tool.

## Tools

| Tool | Actions |
|------|---------|
| `excel_workbook` | list, info, open, save, close |
| `excel_sheet` | list, add, rename, delete |
| `excel_range` | read (paginated + spill metadata), read_spill, write, write_formula, write_py, clear |
| `excel_powerquery` | list, get, create, update, delete, refresh, refresh_all, load_to_table, load_to_datamodel, analyze, analyze_raw |
| `excel_table` | list, create, read, append_rows, add_column, sort, filter, set_style, toggle_totals, rename, delete |
| `excel_pivot` | list, create, add_field, remove_field, move_field, set_layout, refresh, delete, read |
| `excel_datamodel` | info, list_tables, add_table, list_relationships, add_relationship, delete_relationship, list_measures, add_measure, update_measure, delete_measure, refresh |
| `excel_vba` | list_modules, get_module, write_module, delete_module, run |
| `excel_name` | list, get, set, delete (named ranges + LAMBDA formulas) |
| `excel_chart` | list, create, configure, set_source, export_image, delete |
| `excel_screenshot` | range, sheet, chart |

## Environment variables

| Variable | Default | Description |
|---|---|---|
| `THEPEXCEL_MCP_AUTOLAUNCH` | unset | Set to `1` to auto-launch Excel if not running |
| `THEPEXCEL_MCP_ENABLE_VBA` | unset | Set to `1` to enable `excel_vba` tool |
| `THEPEXCEL_MCP_COM_TIMEOUT` | `120` | Per-call COM timeout in seconds |

## VBA setup

`excel_vba` requires both:

1. `THEPEXCEL_MCP_ENABLE_VBA=1` environment variable
2. Excel trust setting: File → Options → Trust Center → Trust Center Settings → Macro Settings → check **"Trust access to the VBA project object model"**

## Known limitations

- **Windows only** — COM is not available on macOS/Linux.
- **Data Model `add_table` requires Excel UI** — `conn.Refresh()` for Mashup/Power Query model connections deadlocks when called from a headless COM automation context (no UI message pump). Works correctly when Claude Desktop is running with a visible Excel window. Smoke-test records SKIP for this case.
- **`=PY()` Python in Excel** — `write_py` inserts the formula but execution is Azure-async. Requires M365 Python in Excel subscription.
- **VBA `Run` return value** — only scalar return values (Long/String/Double) are supported via `Application.Run`. Sub procedures return None.

## Development

```powershell
uv sync
uv run pytest -q                             # 192 unit tests (no Excel needed)
THEPEXCEL_MCP_AUTOLAUNCH=1 uv run python tests/smoke_com.py  # live COM smoke test
```
