# Performance Round 2 — constant-factor polish (post-tactical-pass)

> Status: PLAN. Authored 2026-07-24 as a follow-up to the tactical perf pass in
> commit `1af1527` (2026-07-08), which took the big-O wins (page-scoped reads,
> block writes, bulk_guard, app-cache, spill binary-search).

## Verdict up front (honest ceiling)

The tactical pass already captured the big-O wins. What remains is **constant-factor
territory**. The end-to-end latency stack is:

```
LLM turn (seconds) ≫ MCP stdio (ms) ≫ COM work (~5–50 ms typical) ≫ queue/Future (µs)
```

For heavy tools (PQ refresh, `export_pdf`, screenshot via CopyPicture+clipboard+PIL,
pivot refresh) the wall-clock is **Excel's own compute** — no marshaling fix touches it.
Realistic prize for everything below: **~1.3–2× on property-heavy paths** (format loops,
list enumerations, find/replace scans, large reads) and a few ms off the fixed per-call
floor. This is a polish pass, not a frontier. Cap total investment at ~2–3 focused
sessions; let **Phase 0's numbers** decide whether the risky phases ship at all.

**Riskiest assumption across the whole effort:** that late-bound IDispatch overhead is a
material fraction of call time. It may be 25% of a hot loop or it may be 3% — pywin32's
dynamic dispatch resolves DISPIDs against a cached in-process ITypeInfo, and the
cross-process `Invoke` marshal is identical under early binding. **Phase 0 measures this
in ~30 lines before any early-binding code is written.**

## Constraints (locked)

- Windows + STA single-worker-thread is fixed — no asyncio/threadpool/multiprocess COM.
- Correctness > speed. Do not remove verify-effect read-backs. This drives the user's
  real open workbook.
- 1008 unit tests stay green; public tool API (names/args/return shapes) stable; new
  semantics are additive/opt-in.
- Never wrap PQ-refresh / datamodel / cube paths in `bulk_guard` or any calc-suppression.
- No heavy new dependencies (pywin32 / fastmcp / pillow only). Import stays cheap on
  non-Windows.

## Ranked by impact-per-effort

| # | Phase | Effort | Expected gain | Confidence |
|---|---|---|---|---|
| 1 | **P0 bench harness** | S | Gates/kills everything else + permanent regression baseline | Certain |
| 2 | **P2 spill-probe trim** | XS | 2–4 COM calls off every read page (~0.5–2 ms; highest-frequency path) | Certain |
| 3 | **P1 early binding (safe construction)** | M | 0–40% on property-heavy loops; the only "biggest lever" candidate — bench decides | Medium |
| 4 | **P3 Value2 opt-in** | S–M | Large date/currency-heavy reads only | Medium |
| 5 | P4 per-call floor | — | Deliberately none (see below) | — |

**Sequence:** P0 → P2 (ship immediately) → P1 gated on P0 numbers → P3 gated likewise.
If P0 shows late-bind overhead <10% of hot-loop time: ship only P2, **delete
`_maybe_earlybind` entirely** (dead experimental surface), declare the server perf-complete.

---

## Phase 0 — Measurement harness (DO FIRST, gates everything)

Add `tests/bench_com.py` (stdlib-only, live-Excel, same AUTOLAUNCH pattern as
`smoke_com.py`). Scenarios, print µs/op medians to stderr:
- (a) 1,000× property read (`ws.Cells(1,1).Value`) late-bound **vs** early-bound.
- (b) block read 10k×10 `.Value` **vs** `.Value2`.
- (c) `excel_format` border `all` on 50×10 (property-heavy write loop).
- (d) fixed per-call floor: `run_com(lambda: get_app() and None)` × 100.

Mitigate benchmark noise: fresh AUTOLAUNCH instance + median-of-N.

## Phase 2 — Trim redundant COM calls on the read hot path (certain, zero-risk)

In `ranges.py:_read` the spill-metadata probe (`ranges.py:215-226`) runs on **every**
read: `anchor.HasSpill` + `_is_spill_anchor` (Formula2 read) + maybe `_spill_parent_address`
(SpillParent) = 2–4 round-trips per page, re-paid on every page for identical metadata.

1. **Skip the probe entirely when `offset > 0`** — pagination re-pays it for the same
   anchor/answer. Semantically clean.
2. **Probe `SpillParent` only when the requested range is 1×1** — the "part of someone
   else's spill" hint is only meaningful for a single-cell read.

Same trim applies to `_read_spill` pages. Leave the empty-1×1 `rng.Value` check (only
fires on 1×1).

- Verify: `pytest tests/test_ranges*` + smoke §1; assert the offset-0 response is
  byte-for-byte unchanged. Document in the docstring that `has_spill`/`spill_parent` keys
  appear only on offset-0 / 1×1 reads.

## Phase 1 — Early binding, solved correctly

`_maybe_earlybind` (`session.py:221-242`) uses the **wrong primitive**: `gencache.
EnsureDispatch` *activates* COM and can spawn a duplicate Excel (hence the Hwnd check +
opt-in). The safe construction never activates anything:

1. Generate the makepy cache from the **attached object's own typelib** (zero activation,
   version-exact for the running Office build — no hard-coded `1.9`):
   ```python
   ti = app_late._oleobj_.GetTypeInfo()
   tlb, idx = ti.GetContainingTypeLib()
   la = tlb.GetLibAttr()
   gencache.EnsureModule(la[0], la[1], la[3], la[4])  # guid, lcid, major, minor
   ```
2. Re-wrap the **same pointer**: `app = win32com.client.Dispatch(app_late._oleobj_)` —
   `Dispatch` on an existing IDispatch consults gencache and returns the generated
   early-bound class around the same running instance. Child objects (Workbooks → Workbook
   → Range) come back early-bound automatically. No Hwnd check needed.

Wire into `get_app()`, cache the wrapped handle as today, keep the gen_py self-heal.
Change `THEPEXCEL_MCP_EARLYBIND` from opt-in to a **kill-switch** (default ON) **only
after** Phase 0 shows ≥15% on scenario (a)/(c) *and* full smoke §1–28 passes early-bound.

**Pre-mortem (this is the one phase that can regress correctness):**
- Early-bound wrappers return **typed** values where dynamic returned raw variants (genuine
  `bool` vs `-1/0` ints, typed enums). Grep for `== -1` / msoTriState-style comparisons and
  `int()`-sensitive JSON before flipping.
- Generated-class signatures enforce named params where dynamic tolerated omission — a call
  that "worked" late-bound can raise TypeError early-bound (the historical
  `Slicers.Add`/`Add2` omit-the-arg class).
- gen_py becomes a hard runtime dependency per user machine; Office updates invalidate it —
  the existing self-heal + kill-switch is the mitigation.
- Verify: 1008 unit tests (mocked — structurally unaffected; they CANNOT see
  binding-semantics drift, so smoke is the real gate) + **full** `smoke_com.py` §1–28
  early-bound + A/B in bench. If gain <10% → drop it, delete `_maybe_earlybind`.

## Phase 3 — `.Value2` lane (additive, opt-in) — only if P0(b) shows a gap

Public additive param `excel_range(action="read", value_kind="value2")` (default `"value"`,
shape otherwise identical). `.Value2` skips VT_DATE/VT_CY boxing in Excel *and* the per-cell
`pywintypes.datetime` construction in Python — real Python-side time on large date/currency
reads; serial numbers are also friendlier to JSON. Block `.Value` **is** already the true
per-COM optimum (one Invoke, one SAFEARRAY) — nothing faster exists on the wire; Value2 is
the only remaining axis. Docstring must state the serial-number trade explicitly.

## Phase 4 — Per-call floor: measured, then deliberately left alone

Floor = queue+Future (µs) + `.Name` probe (~0.1–0.3 ms) + get_workbook/get_sheet (1–3
calls) = <2–3 ms against tool calls that cost the LLM seconds. **Do nothing.**

Explicitly **rejected**: caching Workbook/Sheet handles keyed by name. A stale workbook
handle that still answers COM calls → a mutation landing on the **wrong live workbook** =
the nightmare class for this server. The Application `.Name` probe is the cheap liveness
contract that makes the app cache safe; there is no equally cheap orthogonal probe that
proves a *Workbook* handle still binds to what the user thinks. Saving ~1 ms is not worth
that failure mode. Same verdict on removing the `.Name` probe.

## Explicitly NOT doing (scope discipline)

- Multi-action batching / composite "script" tool — saves the LLM round-trip, an API-surface
  question, not perf; API is locked.
- File-sidecar reads (SaveCopyAs → openpyxl for huge reads) — stale-data hazard vs an
  unsaved live workbook; violates the server's reason to exist.
- xlwings or any new dependency — same `Range(Cells,Cells)` + block-Value patterns already
  adopted.
- Any change to PQ/datamodel/cube paths — documented deadlock surface.
- `wait_calculation` poll tuning (50 ms → adaptive) — saves ≤50 ms on second-long ops; noise.
