---
name: thepexcelmcp-verify-effect-testing
description: How to add or change a COM tool in ThepExcelMCP (the thepexcel-excel MCP server) WITHOUT shipping a "success-but-noop" bug — the project's defining test discipline. Use when extending/editing any `domains/*.py` tool or `server.py`, writing a `tests/smoke_com.py` section, or when a tool "returns success but nothing happened", mocked pytest passes but the live tool is a no-op, or you need to prove a mutation actually changed the workbook. Server-maintainer scope. NOT for consuming the tools (→ excel-god) or the COM-fact catalog itself (→ repo CLAUDE.md Session Knowledge + docs/ROADMAP.md).
---

> ⚠️ Absorbed from FableLegacy 2026-07-31 (source repo archived). Audit verdict PARTIAL — stale: §7 wrong — EARLYBIND now default ON with kill-switch; counts drifted. Re-verify those sections before relying on them; everything else re-verified 2026-07-31 — evidence: D:/claude-master/research/2026-07-31-fablelegacy-absorption/g5-content-tools.md


# ThepExcelMCP — verify-EFFECT testing discipline (COM-tool maintainer)

Target repo: `D:/ThepExcelMCP` (PUBLIC remote `ThepExcel/ThepExcelMCP`, verified 2026-07-17).
A Windows-only MCP server that drives **live Excel Desktop** through COM (pywin32). This skill
is the one thing a Sonnet-class maintainer extending the tools will get wrong: **a COM call can
return success while changing nothing.** Mocked unit tests cannot catch that. Read this before
touching any tool.

## 1. What this is / when NOT to use it

This is the **maintainer test discipline** — the repeatable procedure for adding or editing a
COM tool such that you don't ship a mutation that silently does nothing. It synthesizes a method
that today lives scattered across 13 dated `CLAUDE.md § Session Knowledge` log entries, comments
inside `tests/smoke_com.py`, and code comments — none of which state the *procedure* in one place.
It does NOT re-list every COM enum or gotcha; those are a catalog, canonical in the repo.

| If you need… | Use instead |
|---|---|
| To USE the 26 tools (triage / build / consumer-side "verify visually after building") | `excel-god` skill — bundled `D:/ThepExcelMCP/skills/excel-god/SKILL.md`; private twin `D:/claude-master/skills/excel-god/` (the two have **diverged** — see §8) |
| The COM-fact catalog (enum values, Resize quirk, per-tool gotchas, deadlock root cause) | `D:/ThepExcelMCP/CLAUDE.md § Session Knowledge` (13 entries) + `docs/ROADMAP.md § Known pywin32 COM gotchas` / `§ Known limitations` — canonical, not restated here |
| Why "tools vanished from Claude" post-July-2026 | Not a COM/effect bug → MCP protocol skew; `docs/mcp-rc-2026-07-28-migration.md` + Session Knowledge 2026-06-24 |
| The Data Model deadlock as a runtime symptom | `docs/ROADMAP.md § Known limitations` (headless-stdio `WorkbookConnection.Refresh()` bricks the STA worker) — this skill only tells you how to keep it out of your tests (§5) |
| `excel_range write_py` one-off-read fleet rule | Auto-loaded canon `D:/claude-master/rules/windows-env.md § Tool tips` — not restated |
| cp874 / Thai-subprocess rules | Auto-loaded canon `rules/windows-env.md` — not restated |

## 2. The core theorem — units green ≠ effect

**The expected default for a brand-new COM call is: pytest units pass, the tool returns
`success`, and the workbook did not change.** This is not a surprise; it is the baseline you must
actively disprove. It has recurred across at least six tools (§3 table).

**Why mocked units are structurally blind to it.** The unit tests inject a mock session
(`tests/conftest.py` → `make_mock_session`): `run_com(fn, ...)` is a transparent passthrough and
every COM object is a `MagicMock`. A `MagicMock` **records** the attribute/method access and
returns another truthy mock — so a tool that calls the *wrong* COM method, writes to a
single-cell offset instead of a block, or passes an argument Excel silently ignores, still
returns without raising. The unit proves **dispatch** (you called *something* with *some* shape)
and the arg-shaping / guard / error-path logic around it. It can never prove **effect** (a real
Excel workbook actually changed). No amount of mock assertions closes this gap — only a live
read-back does.

**Corollary: mutating tools verify their own effect in production too.** Tools read the actual
state back after acting and never trust the COM return code — e.g. `range_action("write")`
builds the target with `ws.Range(ws.Cells(r1,c1), ws.Cells(r2,c2))` explicitly (never
`.Resize`/`.Offset`, which dispatch as single-cell indexed properties under pywin32) and shape
tools read ACTUAL geometry post-`Add` rather than echoing requested values. Origin: commit
`1d63a55` (2026-06-23), subject "fix(range,table): correct multi-cell write + table sort +
read-back smoke asserts" (commit subject verified via `git log` this session; the underlying
bug/fix is documented in `docs/ROADMAP.md § gotchas`, though the commit itself is not cited there).

### The no-op bug catalog (compact — full detail at the pointers)

Each is a "units passed / live read-back caught it" case. Study the *shape*, not just the fact.

| Tool / call | The silent no-op | One-line cause | Canon |
|---|---|---|---|
| multi-cell `range write` | wrote only cell `[0][0]` for a week | `rng.Resize(r,c)` returns a single offset cell (`Count==1`) under pywin32 IDispatch | Session Knowledge 2026-06-23; ROADMAP § gotchas |
| `table sort` | returned success, order unchanged | `SortOn=1` = `xlSortOnCellColor` (no-op on uniform color) | Session Knowledge 2026-06-23 |
| `slicer add` (Table src) | `E_INVALIDARG` / empty result | `Level` is OLAP-only — must be OMITTED, not `""` or `1` | Session Knowledge 2026-06-23 |
| `view` settings | mutated the wrong workbook | bound to `Application.ActiveWindow` (foreground book) not `wb.Windows(1)` | Session Knowledge 2026-06-23 |
| `validation add` | 1004 on re-add | must `Validation.Delete()` before `Add` | Session Knowledge 2026-06-23 |
| `sparkline _list` | reported 0 sparklines that existed | scoped to `UsedRange`; sparkline cells are value-EMPTY so UsedRange excludes them — scope to the LOCATION range | Session Knowledge 2026-06-24 |
| `slicer` cache | `E_INVALIDARG` on empty `Name` | `SlicerCaches.Add2(..., "")` — omit the empty positional, same class as the Level slot | Session Knowledge 2026-06-24 |

**The unifying rule (three faces, all in the catalog):** *omit optional COM args, never pass
empty/placeholder values*; *bind to the target object, not the app-global*; *delete-before-add
where COM raises 1004 on duplicates*.

## 3. The two/three-layer test model — the mitigation

| Layer | File | Proves | Needs Excel? |
|---|---|---|---|
| Mocked units | `tests/test_<domain>.py` | dispatch, arg-shaping, guards, error → `ToolError` mapping | No (mock session) |
| **Live read-back smoke** | `tests/smoke_com.py` (one section per domain) | **EFFECT** — mutate, then read the state back and assert on the read | Yes (owns its own via `THEPEXCEL_MCP_AUTOLAUNCH`) |
| Static red-team | ad-hoc adversarial review of the diff | correctness units can't judge (guard conditions, off-by-one, wrong-object) | No |

The two-layer pattern (red-team static review + live read-back smoke) is the established default
for COM-tool batches — proven repeatedly (Session Knowledge 2026-06-24 Tier-4/4b "Process win").
**Neither the mock nor the red-team substitutes for the live read-back.** Budget a smoke section
for every new tool; that is where the no-op dies.

## 4. Runbook — add or change a COM tool without a no-op

Ordered. Do not skip the read-back step even when "it obviously works".

1. **Get the COM fact from the SOURCE, not a search snippet.** WebSearch has HALLUCINATED enum
   values (claimed `XlPaperSize` A4=2; truth A4=9/Letter=1/A3=8/Legal=5) — only a direct
   MS-Learn fetch was correct (Session Knowledge 2026-06-24). Verify against a live call or MS
   Learn; tag anything transcribed **DOC-CLAIM** in your commit.
2. **Write the domain fn** in `src/thepexcel_mcp/domains/<domain>.py`. Apply the unifying rule:
   build blocks with `ws.Range(ws.Cells(...), ws.Cells(...))` (never `.Resize`/`.Offset`); omit
   empty optional COM args; bind to `wb.Windows(1)` / the specific object, not `Application.*`;
   delete-before-add where 1004 fires. Wrap COM exceptions to `ToolError` via `_session.wrap`.
3. **Verify effect INSIDE the tool** where cheap — read back the cell value / shape geometry /
   file-on-disk existence+size and return the ACTUAL, never the requested value.
4. **COM discipline that a no-op test would still let you break:**
   - All COM object creation/use must happen INSIDE a `run_com` callable (single STA worker;
     touching a COM object off-worker → `RPC_E_WRONG_THREAD`).
   - **NEVER wrap PQ-refresh / datamodel / cube paths in `bulk_guard`** — deadlock + async-calc
     semantics (docstring `session.py`). `bulk_guard` is for genuinely-bulk write loops only.
   - Poll calc with `wait_calculation` (it pumps the STA message queue); a bare `sleep` loop
     deadlocks.
5. **Add mocked units** (`tests/test_<domain>.py`) — dispatch, arg edge cases, guard conditions,
   error paths. These are necessary but *not sufficient* (§2).
6. **Add a live read-back smoke section** (§5) — the only proof of effect.
7. **Run both** (§6). Fix domain bugs the smoke section surfaces (they will not show in units).
8. **To test through the live MCP tool surface, RESTART the server** — editable install ≠ hot
   reload (§7). pytest + smoke import the code fresh; the *running* stdio server keeps stale code.

## 5. Writing the smoke section (the read-back oracle)

The pattern that catches no-ops: **mutate via the `*_action()` tool, then READ the observable
state back and assert the read changed.** Never assert on the tool's own return code alone.

Canonical templates already in `tests/smoke_com.py` (Section 1 / Section 2):

```python
# EFFECT proof for a multi-cell write — read every cell back (guards the Resize quirk)
payload = [["P", "Q"], ["R", "S"], ["T", "U"]]
range_action("write", range="H1", sheet="Sheet1", workbook=wb_name, values=payload)
rb = range_action("read", range="H1:I3", sheet="Sheet1", workbook=wb_name)
assert rb["total_rows"] == 3
for i, row in enumerate(payload):
    for j, expected in enumerate(row):
        assert rb["values"][i][j] == expected      # ← the read is the oracle

# EFFECT proof for a sort — read back and assert the ORDER (guards the SortOn no-op)
table_action("sort", name="Sales", workbook=wb_name, sort_column="Amount", ascending=False)
rb = table_action("read", name="Sales", workbook=wb_name, limit=20)
amounts = [row[rb["columns"].index("Amount")] for row in rb["values"]]
assert amounts == sorted(amounts, reverse=True)
```

Section scaffolding rules (all observed in `smoke_com.py` this session):

- **Never touch a pre-existing workbook.** Create a throwaway with `_new_wb()` (`Workbooks.Add()`
  on the worker); close it WITHOUT saving in a `finally:` via `_close_wb(wb, name)`.
- **Record per-check**, not per-section: `record("<domain>.<action>", "PASS"|"FAIL"|"SKIP", detail)`.
- **Register the section** in the `_ALL_SECTIONS` dict at the bottom (28 sections as of
  2026-07-17; re-check the dict, don't trust that count). It gives you `--sections N` selection.
- **Deadlock-safe by construction:** anything datamodel-adjacent must SKIP under headless stdio,
  not run — call `_check_excel_busy(...)` to early-SKIP when a prior section left Excel busy, and
  delete PQ queries before closing a workbook (a `wb.Close()` with pending PQ connections
  deadlocks; the harness uses a 30s timeout and sets `_excel_busy`). The 12 Data Model smoke
  items are deliberately SKIPped — that is expected, not a failure (ROADMAP § Known limitations).
- **`smoke_com.py` does NOT exercise the MCP protocol layer** (it owns its own Excel via direct
  import) — a green smoke run says nothing about the client handshake (relevant only to the
  protocol-skew watch item, §8; out of scope for effect testing).

## 6. Running the tests (commands — all UNRUN this session; authored against verified files)

From `D:/ThepExcelMCP`:

```bash
uv run pytest -q                                   # mocked units — no Excel needed, fast
uv run python tests/smoke_com.py                   # LIVE read-back smoke — needs Windows+Excel, ~5-10 min
uv run python tests/smoke_com.py --sections 1,2    # subset by _ALL_SECTIONS key
```

- **CI runs units only** (`.github/workflows/ci.yml`: `windows-latest`, `uv sync` + `uv run
  pytest -q`; `setup-uv` pinned `@v8.2.0` deliberately — comment in the file). `smoke_com.py` is
  intentionally NOT in CI (needs real Excel) — so **CI green does not prove effect**; a human
  runs the live smoke.
- Test counts drift and must never be asserted as facts. Re-verify:
  `grep -rh 'def test_' tests/*.py | wc -l` (1001 defs as of 2026-07-17; `CLAUDE.md` says 1008
  via parametrization; the README badge said 927 — stale). Tool count ground truth:
  `grep -c '@mcp.tool' src/thepexcel_mcp/server.py` (26 as of 2026-07-17).

## 7. Machine-local state (hostname-attributed; authored on `SiraPC`, verified 2026-07-17)

| Item | Fact | Re-verify |
|---|---|---|
| **Editable ≠ hot reload** | `claude mcp add` is an editable install, but the *running* stdio server keeps OLD code in memory after you edit a `domains/*.py`/`server.py` file. Testing THROUGH the live MCP surface requires restarting the server; pytest + `smoke_com.py` import fresh so they don't need it (Session Knowledge 2026-06-23; bit the team twice) | restart the `thepexcel-excel` server / start a fresh Claude session after any code edit |
| Register the server | `claude mcp add thepexcel-excel --scope user -- uv run --directory D:/ThepExcelMCP thepexcel-mcp` (entry point `thepexcel_mcp.server:main`, `pyproject.toml [project.scripts]`) | `claude mcp list` |
| Auto-launch (default ON) | `THEPEXCEL_MCP_AUTOLAUNCH` (default 1) → `get_app()` attaches to running Excel or launches a visible one + blank workbook + self-heals a corrupt `gen_py` cache (commit `53f9436`, 2026-07-06). Smoke relies on this owning its own Excel | `echo "$THEPEXCEL_MCP_AUTOLAUNCH"` (unset = default ON) |
| Other env toggles (names only, not secrets) | `THEPEXCEL_MCP_ENABLE_VBA` (default off), `THEPEXCEL_MCP_COM_TIMEOUT` (120s), `THEPEXCEL_MCP_SLOW_LOG_S` (5.0s → stderr), `THEPEXCEL_MCP_EARLYBIND` (experimental, default off) | `README.md § Environment variables`; `session.py` |
| `uv.lock` is gitignored | `.gitignore` — reproducible resolution rests on the pyproject pin `fastmcp>=3.4.2,<4` only; the migration doc's "uv.lock is already committed" is WRONG (queued doc fix, §8) | `git -C D:/ThepExcelMCP check-ignore uv.lock` |
| Windows COM env only | pywin32 STA worker; Excel Desktop must be installable/runnable. No Linux/cloud path — smoke cannot run in CI or a cloud session | `hostname` (this was authored on `SiraPC`; laptop state UNVERIFIED from here) |

## 8. Open decisions (claude-master `handoff/inbox.yaml` — ids only; do not re-raise)

- `2026-06-23-thepexcelmcp-mcp-rc-deadline` (deadline 2026-07-28) — the "sessionless core"
  protocol-skew watch; NO code migration, operational only. Relevant to effect-testing solely as
  a reminder that `smoke_com.py` does not cover the protocol layer.
- `2026-06-26-drift-thepexcelmcp-pii-hook-branch-protection` — the 2026-06-25 PII leak /
  history-rewrite story lives ONLY in the private inbox, not the repo's own docs.
- `2026-07-16` repo-consolidation entry — the bundled public `skills/excel-god/SKILL.md` has
  diverged from the private claude-master copy (line counts differed 2026-07-17); "add a compare
  step to the release checklist" is queued. Consult the private twin, not only the bundled one.

## 9. Provenance and maintenance

- **Source:** Phase-1 brief `D:/FableLegacy/briefs/ThepExcelMCP.md` (2026-07-17) + live read-only
  re-verification on `SiraPC`, 2026-07-17. Phase-1.5 verdict: no overlap; scope to
  server-maintainer, point consumer symptoms to `excel-god`. Author: Fable-authoring session.
- **VERIFIED this session (read):** `src/thepexcel_mcp/session.py` (worker/guards/`bulk_guard`/
  `wait_calculation`/`get_app`/`wrap`/`_enum_rot_workbooks` defs) · `domains/ranges.py` write
  path (explicit `Range(Cells,Cells)`, the Resize-quirk comment) · `tests/conftest.py`
  (`make_mock_session` passthrough + `MagicMock`) · `tests/smoke_com.py` header, helpers
  (`_new_wb`/`_close_wb`/`_check_excel_busy`/`record`), Section 1+2 read-back oracles,
  `_ALL_SECTIONS` (28 keys) · `tests/test_snapshot.py` (`Save`/`SaveAs`/`Close` `assert_not_called`)
  · `docs/ROADMAP.md § Known limitations` + `§ Known pywin32 COM gotchas` · `CLAUDE.md § Session
  Knowledge` (13 entries) · `.github/workflows/ci.yml` (units-only, `windows-latest`,
  `setup-uv@v8.2.0`) · `pyproject.toml [project.scripts]` · both `excel-god` SKILL.md headers ·
  tool count (26) + test-def count (1001) via grep.
- **UNRUN:** every command in §4–§6 (`uv run pytest -q`, `smoke_com.py`, `claude mcp add/list`) —
  authored against verified files; an authoring session must not mutate the target repo or spin
  up Excel.
- **DOC-CLAIM (transcribed from repo docs/log, not observed running):** the specific no-op
  histories in §2's catalog · the Data-Model-deadlock root cause · the `XlPaperSize` hallucination
  story · README badge staleness · `uv.lock` migration-doc error.
- **Volatile — never state as fact; re-verify:** tool count (`grep -c '@mcp.tool' server.py`);
  test count (`grep -rh 'def test_' tests/*.py | wc -l`); smoke section count (`_ALL_SECTIONS`
  dict); git HEAD / push state (`git -C D:/ThepExcelMCP log --oneline -1`); whether the two
  `excel-god` copies still diverge (`diff D:/ThepExcelMCP/skills/excel-god/SKILL.md
  D:/claude-master/skills/excel-god/SKILL.md`).
- **No `scripts/` shipped, deliberately:** the runnable tools already live in the target repo
  (`tests/smoke_com.py`, `tests/test_*.py`); duplicating a test runner here would fork the canon.
- **Queued upstream side-actions** (execute via the ThepExcelMCP repo's own PR-only change control,
  not from FableLegacy):
  1. `docs/mcp-rc-2026-07-28-migration.md` action-item #1 wrongly says "uv.lock is already
     committed" — it is gitignored (`.gitignore`). One-line doc fix.
  2. Same migration doc cites `server.py:1614` for the stdio line; it is now ~:2676 after the
     2026-07-08 perf pass. Anchor on a quoted string, not a line number.
  3. `README.md` test badge (927) is stale vs actual (~1001–1008). Refresh or make it dynamic.

## 10. Acceptance task

**Task `add-com-tool-noop-safe`** — given ONLY this skill, a fresh Sonnet-class session must
answer, in writing (no repo access needed):

(a) A new `excel_foo(action="set")` mutating tool has passing `tests/test_foo.py` and returns
`{"status": "success"}`. State the single-sentence theorem that says this is NOT evidence the
tool works, and explain precisely WHY the mocked units cannot catch a no-op (name the mock
mechanism from `tests/conftest.py` and what a `MagicMock` returns). (§2)

(b) List the three test layers, which one alone proves EFFECT, and where CI stops. (§3, §6)

(c) Write the shape of the `smoke_com.py` section that would prove the tool's effect — the
mutate-then-read-back oracle, the throwaway-workbook lifecycle, where you register the section,
and the ONE category of action that must SKIP rather than run under headless stdio and why. (§5)

(d) Name two of the six historical "success-but-noop" bugs and the one-line COM cause of each,
and state the unifying three-faced rule they share. (§2 catalog)

PASS = (a)–(d) match §§2–6. Log PASS/FAIL + date in `D:/FableLegacy/reviews/acceptance-log.md`.
