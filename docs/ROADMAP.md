# ThepExcelMCP — Roadmap & Design Decisions

> Synthesized 2026-06-12 from 3 research streams (landscape / COM feasibility / MCP design).
> Full research: vault `Inbox/2026-06-12_thepexcel-mcp-landscape.md`, `_excel-com-feasibility.md`, `_mcp-design-and-gaps.md`.

## Positioning

Windows desktop MCP server controlling a LIVE Excel instance via COM — does what cloud assistants structurally cannot:

| Feature | Claude for Excel | Copilot in Excel | openpyxl MCPs | **ThepExcelMCP** |
|---|:---:|:---:|:---:|:---:|
| VBA execution | No (policy) | No (AutoSave conflicts with .xlsm — structural) | No | **Yes** |
| Power Query M authoring | No | Connect-only | No | **Yes** |
| Data Model / DAX measures | No | Advisory only | No | **Yes** |
| Live workbook (no upload) | No | No | No | **Yes** |

The VBA + Power Query + Data Model cluster requires desktop COM — no cloud competitor can follow.

**Primary reference implementation:** [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel) (C#, MIT, 23 tools / 230 ops) — COM sequences ported to Python. Lineage: supersedes `D:/ThepExcel-PQ-MCP` PoC (its PQ code lives on in `domains/powerquery.py`).

## Architecture decisions

- **Stack:** Python 3.11+ / FastMCP / pywin32. stdio transport (Claude Desktop, Claude Code, Cursor, Copilot all spawn stdio).
- **Tool design:** action-dispatch — few coarse domain tools (`excel_table(action=...)`) not one-tool-per-COM-method. Ceiling ~12 tools (LLM reliability degrades >30).
- **Errors:** `ToolError` with actionable text listing alternatives ("Sheet 'X' not found. Available: ..."), never raw COM exceptions.
- **Reads:** paginated (default 100 rows, `has_more`/`next_offset`), 500-char cell truncation.
- **COM threading (Phase 3):** Excel is STA — all COM calls must run on ONE dedicated worker thread (queue-fed); `CoInitialize` there; `CalculationState` polling with `PumpWaitingMessages` (sleep-only loops deadlock the STA queue); `DisplayAlerts/Interactive=False` guards around prompts.

## Feasibility verdicts (researched 2026-06-12, MS Learn-verified)

| Feature | Verdict | Note |
|---|---|---|
| Power Query M CRUD + load (table/model) | ✅ Shipped (P0/P2) | `Workbook.Queries` + Mashup-OLEDB `Connections.Add2` |
| Tables / Pivots (incl. data-model pivots) | ✅ Shipped (P1) | data-model pivot via `ThisWorkbookDataModel` connection, measures = CubeFields |
| Data Model / DAX measures | ✅ Shipped (P2) | `ModelMeasures.Add`; no Formula setter → update = delete+re-add; `FormatInformation` required |
| Dynamic arrays / LAMBDA | ✅ P3 | `Range.Formula2`, `Names.Add` for named LAMBDA, spill introspection via `HasSpill`/`SpillingRange` |
| VBA inject + run | ✅ P3 | `VBE.VBComponents` + `Application.Run` (returns values). Requires user-enabled `AccessVBOM` trust setting (HKCU…\Security\AccessVBOM=1) — pre-flight check + opt-in env `THEPEXCEL_MCP_ENABLE_VBA` |
| Python in Excel | 🟡 P4, write-only | COM can insert `=PY()` but execution is async in Azure — cannot await results. Insert-formula convenience only |
| Office Scripts | ❌ Skip | OneDrive-stored, internet-required, no COM path. Server replies "use VBA instead" |

## Phases

- **P0 ✅** Scaffold + PQ port — `excel_workbook` `excel_sheet` `excel_range` `excel_powerquery`
- **P1 ✅** `excel_table` (11 actions) + `excel_pivot` (9 actions)
- **P2 ✅** `excel_datamodel` (11 actions) + `load_to_datamodel`
- **P3 ✅** STA worker-thread hardening · `excel_vba` · LAMBDA/named formulas (`excel_name`) + spill introspection (`read_spill`, spill metadata in `read`)
- **P4 ✅** `excel_chart` (list/create/configure/set_source/export_image/delete) · `excel_screenshot` (range/sheet/chart → PNG, CopyPicture+PIL) · `excel_range(action="write_py")` (`=PY()` Formula2R1C1 insertion, experimental)
- **P5** Live end-to-end smoke vs real Excel · packaging (uvx + MCPB bundle for Claude Desktop) · client registration
- **Later** Companion SKILL (Excel best practices: when to use Table vs Model, M patterns, DAX patterns) — separate from tools; snapshot/undo + excel_diff (pattern from lingfan36/ai-office-mcp); progressive-disclosure meta-tool if tool context grows

## Moat ideas nobody covers (from landscape scan)

Python-in-Excel insertion, live spill-range introspection, LAMBDA authoring, Power Pivot relationship graph export, Thai-aware error/doc text, ThepExcel teaching-aligned best-practice skill layer.
