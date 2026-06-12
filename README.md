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

## Tools (Phase 0)

| Tool | Actions |
|------|---------|
| `excel_workbook` | list, info, open, save, close |
| `excel_sheet` | list, add, rename, delete |
| `excel_range` | read (paginated), write, write_formula, clear |
| `excel_powerquery` | list, get, create, update, delete, refresh, refresh_all, load_to_table, analyze, analyze_raw |

## Development

```powershell
uv sync
uv run pytest          # analyzer unit tests (no Excel needed)
uv run python tests/smoke_com.py  # manual COM smoke (requires live Excel)
```
