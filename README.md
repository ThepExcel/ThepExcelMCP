# ThepExcelMCP

A Windows MCP server that drives a **live running Excel Desktop instance** via COM (pywin32), giving AI agents full Excel capability that file-only libraries cannot provide. When an agent calls a tool here, it is talking to the actual Excel process — Power Query queries refresh against live data sources, PivotTables rebuild with real aggregations, charts render to PNG for visual verification, DAX measures execute in the in-memory Data Model, and dynamic-array spill ranges resolve correctly. This is the difference between editing an XML zip file and actually using Excel.

## How it runs — local, not a hosted service

ThepExcelMCP is a **stdio MCP server that runs as a process on your own Windows
machine** and controls the Excel running there. There is nothing to sign up for, no
account, and no data leaves your computer. Use it from any MCP client that can
launch a local stdio server on that same machine:

- ✅ **Claude Code**, **Claude Desktop**, and **Codex CLI** run locally and can launch it.
- ❌ Cloud-only agent surfaces that can't run a local process on your machine — e.g.
  **Claude Cowork** today — can't reach your local Excel, so they can't use it.

## Why COM, not a file library?

Libraries such as `openpyxl` read and write the `.xlsx` file on disk, but they
cannot *operate Excel*. ThepExcelMCP drives the real application, so an AI can
**author a complete solution from a blank workbook** — not just tweak a file that
already exists:

- **Build a whole data pipeline end to end, in one tool.** Hand the AI a blank
  workbook and it can *write* a Power Query in M from scratch, merge and reshape
  several sources, load the result into a worksheet Table, then build a PivotTable
  and a chart on top — the entire flow, start to finish. Refreshing an existing
  query is the trivial part; this *creates* the query, the table, and the pivot.
- **Power Query, fully editable.** Read and rewrite the M of any existing query,
  create and set query parameters, refresh against live data sources, and load to a
  worksheet Table or straight into the Data Model. (A file library cannot evaluate
  M code at all.)
- **Data Model & DAX (Power Pivot) that actually calculate.** Add model tables and
  relationships, write and update DAX measures, and have Excel's in-memory engine
  compute them — plus CUBEVALUE / CUBEMEMBER helpers.
- **PivotTables that really pivot.** From a range, a table, or the Data Model; add,
  move, and remove fields with real aggregations; set layout and refresh — every
  value computed by Excel, not approximated.
- **Live recalculation.** Dynamic-array spill formulas (`XLOOKUP`, `SORT`,
  `FILTER`), volatile functions, and ordinary formulas are evaluated by Excel's
  calculation engine, and the spilled results can be read back.
- **Python in Excel.** Insert `=PY()` formulas that run in Excel's Python runtime.
- **Charts, formatting, and the visual layer** — and capture any range, sheet, or
  chart as a PNG so the AI can *see* what it just built and verify it.
- **VBA and the full object model.** Run macros and reach anything COM exposes.
- **A safety net.** Take a non-destructive snapshot before risky automation, and
  diff two ranges or sheets to see exactly what changed.

In short: a file library edits a spreadsheet *file*; ThepExcelMCP *uses Excel*.
Every call runs through a single STA COM worker thread, so Excel's own rules for
calculation, formatting, and event handling apply exactly as they would for a
human sitting at the keyboard.

## Requirements

- Windows 10 or 11
- Python 3.11 or later with [uv](https://docs.astral.sh/uv/) installed
- Microsoft 365 Excel Desktop (Excel must be **running** before the first tool call, or set `THEPEXCEL_MCP_AUTOLAUNCH=1`)

## Platform support

This server drives a **real Excel process through Windows COM** (`pywin32`), so it is **Windows only by design**:

- ✅ **Windows 10 / 11** — fully supported.
- ❌ **macOS** — not supported. Excel for Mac does not expose the COM automation API, and `pywin32` does not exist on macOS. There is no workaround; this is a hard platform limitation, not a missing feature.
- ⚠️ **WSL (Windows Subsystem for Linux)** — the server itself **cannot run inside WSL**, because the Linux Python in WSL has no `pywin32` and no path to Windows COM. But if your MCP client (e.g. Claude Code) runs *inside* WSL, you can still use the server by running it as a **native Windows process** across the WSL ↔ Windows interop boundary:
  1. Clone and `uv sync` the repo on the **Windows** filesystem (e.g. `C:\Tools\ThepExcelMCP`), using Windows `uv` — not inside the WSL filesystem.
  2. From your WSL client, register the server so the *command* is the Windows `uv.exe` pointing at that Windows checkout:
     ```bash
     claude mcp add thepexcel-excel --scope user -- \
       uv.exe run --directory 'C:\Tools\ThepExcelMCP' thepexcel-mcp
     ```
     The stdio pipes bridge the WSL→Windows process boundary; the server process and Excel both run natively on Windows. Excel must be open in your **Windows** desktop session (WSLg/Linux GUI Excel will not work).

  This WSL path runs the same code on Windows under the hood — it is offered as guidance, not a separately tested install mode. The supported, no-surprises setup is plain Windows.

## Install

```bash
git clone https://github.com/ThepExcel/ThepExcelMCP.git
cd ThepExcelMCP
uv sync
```

`uv sync` installs `fastmcp`, `pywin32`, and `pillow` into an isolated virtual environment. No pip or manual virtualenv setup required.

## Register with Claude Code (CLI)

```bash
claude mcp add thepexcel-excel --scope user -- uv run --directory /path/to/ThepExcelMCP thepexcel-mcp
```

On Windows the path will look like:

```powershell
claude mcp add thepexcel-excel --scope user -- uv run --directory C:\path\to\ThepExcelMCP thepexcel-mcp
```

Verify the server appears and starts cleanly:

```bash
claude mcp list
```

To enable VBA and auto-launch Excel at startup, pass environment variables:

```powershell
claude mcp add thepexcel-excel --scope user `
  -e THEPEXCEL_MCP_AUTOLAUNCH=1 `
  -e THEPEXCEL_MCP_ENABLE_VBA=1 `
  -- uv run --directory C:\path\to\ThepExcelMCP thepexcel-mcp
```

## Register with Claude Desktop

Edit `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "thepexcel-excel": {
      "command": "uv",
      "args": ["run", "--directory", "C:\\path\\to\\ThepExcelMCP", "thepexcel-mcp"],
      "env": {
        "THEPEXCEL_MCP_AUTOLAUNCH": "1"
      }
    }
  }
}
```

Restart Claude Desktop after saving. The server starts on demand when you first call an Excel tool.

## Register with Codex CLI

Codex talks to MCP servers over the same stdio transport, so the server runs unchanged. Add it with the `codex mcp add` command:

```powershell
codex mcp add thepexcel-excel -- uv run --directory C:\path\to\ThepExcelMCP thepexcel-mcp
```

Or edit `~/.codex/config.toml` directly:

```toml
[mcp_servers.thepexcel-excel]
command = "uv"
args = ["run", "--directory", "C:\\path\\to\\ThepExcelMCP", "thepexcel-mcp"]

[mcp_servers.thepexcel-excel.env]
THEPEXCEL_MCP_AUTOLAUNCH = "1"
```

The `[mcp_servers.thepexcel-excel.env]` table is optional — drop it if you would rather start Excel yourself before the first call.

## Tools

There are **26 tools** in total, organized by category.

### Core

| Tool | Actions | Description |
|------|---------|-------------|
| `excel_workbook` | list, info, open, save, close, create, save_as | Manage open workbooks: discover, open, create, save, save-as, and close. |
| `excel_sheet` | list, add, rename, delete | Manage worksheets within a workbook: list, add, rename, delete. |
| `excel_range` | read, read_spill, write, write_formula, write_py, clear | Read and write cell ranges, including dynamic-array spill and Python-in-Excel formulas. |

### Data / Power Query / Data Model

| Tool | Actions | Description |
|------|---------|-------------|
| `excel_powerquery` | list, get, create, update, delete, refresh, refresh_all, load_to_table, load_to_datamodel, analyze, analyze_raw, create_parameter, get_parameter, set_parameter, list_parameters | Manage Power Query (M code) queries: CRUD, refresh, load to table/model, analyze, and parameter ops. |
| `excel_table` | list, create, read, append_rows, add_column, sort, filter, set_style, toggle_totals, rename, delete | Manage Excel Tables (ListObjects): create, read, append, add columns, sort, filter, style, totals, rename, delete. |
| `excel_pivot` | list, create, add_field, remove_field, move_field, set_layout, refresh, delete, read | Manage PivotTables: create, add/remove/move fields, set layout, refresh, read, and delete. |
| `excel_datamodel` | info, list_tables, add_table, list_relationships, add_relationship, delete_relationship, list_measures, add_measure, update_measure, delete_measure, refresh, cube_formula, cube_value, cube_member, add_calculated_column, add_calculated_table | Manage the Excel Data Model (Power Pivot): tables, relationships, DAX measures, and CUBE formula helpers. |
| `excel_name` | list, get, set, delete | Manage defined names: named ranges, constants, and LAMBDA formulas. |

### Objects

| Tool | Actions | Description |
|------|---------|-------------|
| `excel_chart` | list, create, configure, set_source, export_image, delete | Create and manage embedded charts: create, configure, change source, export PNG, delete. |
| `excel_shape` | add_image, add_textbox, add_shape, list, move, delete | Add and manage drawing objects (images, text boxes, AutoShapes): add, list, move, delete. |
| `excel_slicer` | add, add_timeline, list, delete, connect | Manage slicers and date timelines for PivotTables and Tables: add, add_timeline, list, delete, connect. |
| `excel_sparkline` | add, clear, list | Add, clear, and list in-cell sparkline mini-charts (line, column, win/loss). |

### Formatting / View

| Tool | Actions | Description |
|------|---------|-------------|
| `excel_format` | font, fill, border, number_format, alignment, column_width, row_height, autofit | Apply cell formatting: font, fill, border, number format, alignment, column/row sizing, autofit. |
| `excel_view` | freeze_panes, unfreeze_panes, gridlines, zoom, headings | Control worksheet display settings: freeze panes, gridlines, zoom, row/column headings. |
| `excel_conditional_format` | data_bar, color_scale, icon_set, cell_rule, top_bottom, clear | Add or remove conditional formatting rules: data bar, color scale, icon set, cell rule, top/bottom, clear. |
| `excel_validation` | list, whole_number, decimal, date, text_length, custom, clear | Add or remove data validation on ranges: dropdown list, number/decimal/date/text-length/custom constraints, clear. |

### Page / Print

| Tool | Actions | Description |
|------|---------|-------------|
| `excel_page_setup` | set, print_area, print_titles, header_footer, export_pdf, get | Configure page setup and export to PDF: orientation, paper size, margins, fit-to-page, print area, titles, headers/footers, PDF export. |

### Annotation

| Tool | Actions | Description |
|------|---------|-------------|
| `excel_comment` | add, edit, reply, delete, list, get | Add, edit, reply to, delete, list, and get cell comments (legacy notes and threaded). |
| `excel_hyperlink` | add, list, delete | Add, list, and delete worksheet hyperlinks (URL, internal, email, file). |
| `excel_outline` | group_rows, group_columns, ungroup_rows, ungroup_columns, show_levels, clear | Group/ungroup rows and columns, control outline levels, and clear all groupings. |
| `excel_protection` | protect_sheet, unprotect_sheet, protect_workbook, unprotect_workbook, set_locked, status | Protect/unprotect worksheets and workbooks, set cell lock/formula-hidden flags, and query protection status. |

### Safety

| Tool | Actions | Description |
|------|---------|-------------|
| `excel_vba` | list_modules, get_module, write_module, delete_module, run | Manage VBA modules and run macros (opt-in, requires `THEPEXCEL_MCP_ENABLE_VBA=1` and AccessVBOM trust). |
| `excel_screenshot` | range, sheet, chart | Capture a range, sheet, or chart as PNG for visual verification by the AI agent. |
| `excel_find_replace` | find, count, replace | Find, count, or replace text across a range, sheet, or entire workbook with match-case and whole-cell options. |
| `excel_diff` | ranges, sheets | Diff two ranges or two whole sheets cell-by-cell on values, formulas, or both (pure read, no mutation). |
| `excel_snapshot` | snapshot, list, restore, delete | Non-destructive workbook safety copies: snapshot (SaveCopyAs), list, restore (open copy as new workbook), delete. |

## Environment Variables

| Variable | Default | Description |
|---|---|---|
| `THEPEXCEL_MCP_AUTOLAUNCH` | unset | Set to `1` to auto-launch a visible Excel instance if none is running. |
| `THEPEXCEL_MCP_ENABLE_VBA` | unset | Set to `1` to enable the `excel_vba` tool. Off by default for security. |
| `THEPEXCEL_MCP_COM_TIMEOUT` | `120` | Per-call COM timeout in seconds. Increase for slow data refreshes. |

## VBA Setup

`excel_vba` requires both gates to be open:

1. Set the environment variable `THEPEXCEL_MCP_ENABLE_VBA=1` when registering the server.
2. Enable the trust setting in Excel: **File → Options → Trust Center → Trust Center Settings → Macro Settings → check "Trust access to the VBA project object model".**

Without both, any `excel_vba` call returns a clear error explaining which gate is missing.

## Using the Bundled Skill

The repository includes a Claude skill at `skills/excel-god/` — a strategy and orchestration guide that helps an AI agent decide which tools to call, in what order, for common Excel tasks (building dashboards, cleaning data with Power Query, setting up a Data Model, etc.).

To use it, copy the `excel-god` folder into your skill directory:

```bash
# project-level
cp -r skills/excel-god /path/to/your/project/.claude/skills/

# user-level (available in all projects)
cp -r skills/excel-god ~/.claude/skills/
```

Then invoke it with `/excel-god` in a Claude Code session.

## Troubleshooting / Known Limitations

**Excel must be running before the first call.**
The server connects to the first Excel instance visible in the Windows Running Object Table. If Excel is not open, all tool calls will fail with a COM error. Either open Excel manually first, or set `THEPEXCEL_MCP_AUTOLAUNCH=1`.

**Screenshots and chart image export require a visible Excel window.**
`excel_screenshot` and `excel_chart(export_image)` use Excel's `CopyPicture` API, which needs the window to be rendered on screen. If Excel is minimized or its window is hidden behind other apps, the captured PNG may come out empty. Keep the Excel window visible when using these tools.

**Data Model `add_table` and Power Query `load_to_datamodel` can deadlock in a headless stdio context.**
Loading data into the in-memory Data Model triggers a Mashup/Power Query refresh that requires Excel's UI message pump. In a Claude Code CLI session (stdio transport, no visible window), this operation can deadlock the COM worker thread and brick all subsequent tool calls until Excel is force-killed. With Claude Desktop and a fully visible Excel window, it works correctly. Fallback: use `excel_powerquery(load_to_table)` to load into a worksheet Table, then build a PivotTable from that table with `excel_pivot(create, source="<table name>")` (this exact flow — cross-file Power Query merge → table → pivot → chart → slicer — is verified end-to-end).

**Named LAMBDA via `excel_name`.**
Parameter names in a LAMBDA must not look like cell references. Avoid names like `q1`, `x2`, `c3` — use descriptive names such as `val`, `rate`, `n`, or `x` instead. If a LAMBDA add fails partway through, Excel may leave hidden `_xlpm.*` helper names in the workbook that block later LAMBDA adds. If this happens, use a fresh workbook.

**`=PY()` Python in Excel.**
`excel_range(write_py)` inserts a Python-in-Excel formula, but execution is handled asynchronously by Microsoft's cloud service. It requires an M365 Python in Excel subscription and an active internet connection. The tool inserts the formula; it does not wait for or verify the cloud result.

**VBA `run` return values.**
`excel_vba(run)` returns scalar values (Long, String, Double) from VBA Functions. Sub procedures return None. Complex return types (arrays, objects) are not supported.

**Windows only.**
COM is a Windows technology. This server cannot run on macOS or Linux.

## Development

```bash
uv sync                                                       # install dependencies
uv run pytest -q                                              # 927 unit tests, no Excel needed
THEPEXCEL_MCP_AUTOLAUNCH=1 uv run python tests/smoke_com.py  # live COM smoke test (requires Windows + Excel)
```

To run a subset of the smoke test:

```bash
THEPEXCEL_MCP_AUTOLAUNCH=1 uv run python tests/smoke_com.py --sections 1,2,3,4
```

Sections 1–28 cover all tool categories. The full suite runs in roughly 5–10 minutes depending on Excel startup time.

## Acknowledgments

This project stands on the work of others in the Excel-automation and MCP community:

- **[sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel)** (MIT) — the primary reference implementation. A number of the Excel COM automation sequences here were studied and ported from this C# project to Python. If you want a mature C#/.NET COM-based Excel MCP server, theirs is excellent.
- **[lingfan36/ai-office-mcp](https://github.com/lingfan36/ai-office-mcp)** — design inspiration for the snapshot/undo and range-diff tooling.
- **[haris-musa/excel-mcp-server](https://github.com/haris-musa/excel-mcp-server)** — a reference for MCP tool API design conventions (file-based, `openpyxl`).

Upstream license texts are reproduced in [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).

## Contributing

Work on a branch and open a pull request — `main` is protected. Please use
synthetic data only (this is a public repo) and enable the pre-push safety hook.
See [CONTRIBUTING.md](CONTRIBUTING.md).

## License

MIT License. Copyright (c) 2026 ThepExcel <thepexcel@gmail.com>.
