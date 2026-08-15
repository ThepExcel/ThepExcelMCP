# COM gotchas — hard-won API facts

Durable Excel-COM facts and traps learned from live testing. Read the section for the
domain you are touching before editing it. Mocked unit tests PASS straight through most
of these bugs — only live read-back catches them (discipline:
`.claude/skills/thepexcelmcp-verify-effect-testing/SKILL.md`).

New durable facts land here, dateless and imperative — not in CLAUDE.md.

## Testing discipline

- Editable install ≠ hot reload: the running stdio server keeps OLD code in memory after
  you edit a file. Restart the MCP server to test the tool surface; `pytest` +
  `tests/smoke_com.py` validate code without a restart (they own their Excel via
  AUTOLAUNCH).
- Verify EFFECT, not the success report: several tools returned `success` while doing
  nothing. Canonical example: pywin32 `.Resize(r, c)` returns a single offset cell
  (Count=1) — use `Range(Cells, Cells)` instead.
- `get_app()` uses `GetActiveObject` = the FIRST ROT Excel instance. If the user has
  their own Excel open, a `workbook=None` call targets THEIR active workbook — when
  testing, always create and target an explicit test workbook.
- `tests/_excel_lifecycle.py::ExcelInstanceGuard` snapshots EXCEL.EXE PIDs before the
  first `get_app()` and quits only PIDs the harness itself spawned — a pre-existing
  Excel is never touched. A harness-launched EXCEL.EXE can linger ~a minute after
  `Quit()` (the STA worker still holds a COM reference).
- Manual Excel launch (when `THEPEXCEL_MCP_AUTOLAUNCH=0`): `DispatchEx`,
  `.Visible=True`, then `.Workbooks.Add()` — the Add is REQUIRED; Excel registers in
  the ROT only once it owns a document.
- Known pre-existing live-smoke failures (NOT regressions): the Data Model
  relationship/measure cascade (§5/§8) and `cube.live_cube_value` (§18) — the same
  headless-stdio data-model limitation family.
- `smoke_com.py` does not exercise the protocol layer. After any fastmcp bump, run
  `tests/protocol_smoke_v2.py` plus a manual client handshake
  (`docs/mcp-rc-2026-07-28-migration.md`).

## Data Model / cube

- `excel_datamodel(add_table)` / `excel_powerquery(load_to_datamodel)` DEADLOCK the STA
  worker in the headless stdio context — bricks ALL subsequent calls until Excel is
  force-killed (not a recoverable timeout). Fallback: `load_to_table` →
  pivot-from-table.
- Cube formulas: write CUBEVALUE/CUBEMEMBER, then
  `Application.CalculateUntilAsyncQueriesDone()` to resolve `#GETTING_DATA`. Numeric
  resolution needs an EXISTING in-workbook model + the named measure (headless
  model-build deadlocks). `cube_formula` is build-only (no COM). Connection literal =
  `"ThisWorkbookDataModel"`.
- DAX calculated column/table are IMPOSSIBLE via COM (`ModelTableColumns` is read-only,
  no `.Add`; calculated tables are Power BI/AS-only). `add_calculated_column` /
  `add_calculated_table` are pure-Python guards raising a ToolError that points to the
  Power Query workaround.

## Power Query

- A PQ parameter is a normal query whose M is
  `<literal> meta [IsParameterQuery=true, Type=..., IsParameterQueryRequired=...]`;
  `Queries.Add` stores M verbatim so it round-trips — pure `wb.Queries`, no
  data-model/deadlock surface.
- Split the leading literal from the meta record STRUCTURALLY (walk the quoted string
  honoring `""` escapes — `_split_literal_and_meta` in powerquery.py), NEVER by
  string-searching `' meta '`: a Text value can itself contain `' meta '`.
- Parameter value guard is `value is None`, not `if not value` — `0`, `''`, `False` are
  valid parameter values.

## Ranges / reads / writes

- `.Value` / `.Formula` return a SCALAR for one cell, a flat tuple for one row/col, and
  tuple-of-tuples for N×M — normalize to 2-D before comparing or mapping (`_to_2d` in
  diff.py).
- Paginated reads marshal only the page's `ws.Range(ws.Cells(...), ws.Cells(...))`
  block — never whole-range `.Value` + Python slice (quadratic). Preserve the legacy
  quirk: a truly-empty 1×1 range reports `total_rows=0`, not one empty row.
- `append_rows`: one block `.Value` write capped at the WIDEST SUPPLIED row — writing
  None across unsupplied columns would clear calculated-column formulas that
  `ListRows.Add` auto-fills. Use one `ListObject.Resize` only when the expansion band is
  empty; otherwise fall back to insertion-preserving `ListRows.Add`; verify the final
  row count.
- LAMBDA (`excel_name`): cell-ref-like parameter names (`q1`, `x2`) are rejected by
  Excel, and a failed LAMBDA add leaves undeletable `_xlpm.*` orphans that block later
  adds in the same workbook. On a clean workbook, named LAMBDA works fine, incl.
  multi-arg.

## View / validation / slicer

- View settings (freeze/gridlines/zoom/headings) must bind to `wb.Windows(1)`
  (workbook-scoped), NOT `Application.ActiveWindow` — ActiveWindow is the foreground
  book, so a background target gets silently mutated on the wrong workbook.
- `Validation.Add` raises 1004 when a rule already exists — always
  `Validation.Delete()` first.
- `Slicers.Add`: `Level` is OLAP-only — on Table / regular-pivot sources it must be
  OMITTED (call with keyword args), else COM `E_INVALIDARG`. Same class:
  `SlicerCaches.Add2` fails on an EMPTY-STRING Name — omit the Name positional
  (`Add2(src, field, SlicerCacheType=...)`).

## Sparkline

- `SparklineGroups.Add(Type, SourceData)` REQUIRES `Application.ReferenceStyle = xlA1`
  — save, force, restore around the call.
- `XlSparkType`: line=1, column=2, win_loss=3 (`xlSparkColumnStacked100`).
- List by scoping to the LOCATION range, never `UsedRange` — sparkline destination
  cells are value-empty, so UsedRange excludes them (`UsedRange...Count`=0 while
  `Range(location)...Count`=1).

## Comments

- Two separate systems: legacy NOTES (`cell.Comment` / `AddComment`, `.Author` is a
  plain string) vs THREADED (`cell.CommentThreaded` / `AddCommentThreaded`,
  `.Author.Name`, `.AddReply(text)`).
- `Comment.Text()` / `CommentThreaded.Text()` are METHODS — call `.Text()` to read,
  `.Text(x)` to write; never the attribute.
- Delete-before-add: `AddComment` raises 1004 if a comment exists. Threaded text is NOT
  editable via COM → `edit` is note-only. Threaded comments do work on Win11 Excel
  desktop.

## Page setup

- Margins are in POINTS → `Application.InchesToPoints()`.
- Set `Zoom = False` BEFORE `FitToPagesWide/Tall`.
- `XlPaperSize`: A4=9, Letter=1, A3=8, Legal=5 — verify COM enums against Microsoft
  Learn directly; web-search snippets have returned wrong values for these.
- `ExportAsFixedFormat(Type=0)` on ws (sheet) or wb (workbook); verify the file on disk
  (exists + size). "Microsoft Print to PDF" satisfies the driver requirement.

## Protection

- `protect_workbook` verify-effect must be conditional
  (`if structure and not wb.ProtectStructure`) — an unconditional check raises on a
  legitimate `structure=False` call.
- `Password=None` must OMIT the keyword (passing `""` can raise 1004); wrong-password
  Unprotect → 1004.
- `.Locked` / `.FormulaHidden` only take effect while sheet protection is on.

## Outline

- `ws.Rows("2:5").Group()` / `ws.Columns("B:D").Group()`. `OutlineLevel`: ungrouped=1
  (not 0), grouped ≥2; max 8 levels; clear via `ws.Cells.ClearOutline()`.

## Shapes

- `AddPicture` needs an ABSOLUTE path (relative → E_FAIL); Width/Height=`-1` = native
  size; `msoTriState`: msoFalse=0, msoTrue=-1; `LinkToFile=msoFalse` REQUIRES
  `SaveWithDocument=msoTrue`.
- `AddTextbox(Orientation=1)` then `shape.TextFrame2.TextRange.Text` (TextFrame is the
  legacy fallback).
- `msoAutoShapeType`: rectangle=1, rounded_rectangle=5, oval=9, diamond=4, triangle=7,
  hexagon=10, right_arrow=33, star=92, cloud=179, heart=21.
- `Shape.Type` (MsoShapeType: AutoShape=1, Picture=13, TextBox=17) ≠
  `Shape.AutoShapeType` (meaningful only when Type==1).
- Positions are POINTS; an anchor cell's `.Left`/`.Top` are already points.
  Verify-effect reads the ACTUAL post-Add geometry — AddPicture resolves `-1` to real
  dims; never echo `-1` back.

## Find / replace

- `Range.Find` returns **None** on no-match — never call `.Address` on None.
- `FindNext` wraps silently — save the first hit's `.Address` and stop when it cycles
  back, or the loop hangs live Excel.
- `Range.Replace` returns a Boolean, NOT a count — derive counts from the Find loop.
  `Replace` has no `LookIn` param (always operates on formulas).
- Enums: xlFormulas=-4123, xlValues=-4163, xlWhole=1, xlPart=2, xlByRows=1.
- Verify-effect = post-replace remaining count == 0.

## Diff

- Pure read — no deadlock surface. Map the diff-cell A1 address from the LEFT range's
  `.Row`/`.Column` origin (not 1,1). Shape mismatch → diff the overlap and emit
  `shape_note`; never silently drop extra rows/cols. Cap at `max_diffs` but always
  report true `total_diffs` + `truncated`.

## Snapshot

- `wb.SaveCopyAs(path)` is the safe primitive: streams a copy to disk WITHOUT touching
  the live workbook's `Saved` flag / `Name` / `FullName`. Never use `SaveAs` for
  backups — SaveAs REBINDS the open workbook to the new path.
- `restore` = `Workbooks.Open(copy)` as a separate NEW workbook. An in-place revert
  (close-current + reopen) is deliberately NOT built — that is the destructive path.
- `delete` only `os.remove`s the registry-tracked copy under
  `%TEMP%/thepexcel_mcp/snapshots` (server-constructed path — no arbitrary delete).
  Deleting a copy still open in Excel → `WinError 32`, surfaced honestly as a
  ToolError.
- The snapshot registry is in-memory / session-scoped (lost on server restart); files
  persist on disk. Format is preserved by extension (`.xlsm` keeps macros).

## Performance layer (session.py)

- `get_app()` caches the Application handle on the STA worker; liveness probe =
  `.Name`; any exception drops the cache and re-resolves through the unchanged
  attach → self-heal → auto-launch chain.
- gen_py self-heal: a corrupt win32com cache throws
  `AttributeError: ... 'CLSIDToClassMap'` on Dispatch/GetActiveObject and masquerades
  as "Excel not running" — it is cleared and retried once.
- Early binding (`THEPEXCEL_MCP_EARLYBIND=1`, default ON): reads the typelib off the
  already-attached late-bound Application and re-wraps the SAME IDispatch pointer via
  `gencache.EnsureModule` + `Dispatch` — never activates COM, so it cannot spawn a
  duplicate Excel. Kill switch `=0` explicitly constructs a dynamic wrapper
  (GetActiveObject can return gen_py even when rewrap is skipped). ~1.09x faster
  property reads.
- `bulk_guard(app)` suppresses ScreenUpdating/EnableEvents + manual calc around bulk
  write loops (table append_rows, border all/inside, find_replace replace,
  load_to_table sheet-rebuild). It restores the SAVED calculation mode (the user may
  run manual calc deliberately) and forces one `Calculate()` only if it changed the
  mode. NEVER wrap PQ refresh / datamodel / cube paths — deadlock + async-calc
  semantics.
- `value_mode="raw"` on range/table/pivot reads returns `.Value2` (~1.22x on large
  mixed reads); typed `.Value` stays the default (date/currency semantics differ).
- Slow COM calls (> `THEPEXCEL_MCP_SLOW_LOG_S`, default 5.0s) log one line to STDERR.
  stdout is the stdio transport — never print there.
