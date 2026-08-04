"""Live-Excel perf bench harness — constant-factor measurement (Phase 0).

Requires a live Excel instance (or THEPEXCEL_MCP_AUTOLAUNCH=1), same
AUTOLAUNCH + own-scratch-workbook idiom as smoke_com.py: all bench work runs
inside a single Workbooks.Add() scratch workbook and is closed WITHOUT saving
at the end — no pre-existing workbook is ever touched.

Run:
    THEPEXCEL_MCP_AUTOLAUNCH=1 uv run python tests/bench_com.py

Not collected by pytest — it needs live Excel and the filename doesn't match
the `test_*.py` pattern in pyproject.toml's testpaths (same convention as
smoke_com.py). Meant to be run standalone; also guarded by
`if __name__ == "__main__"`.

Instance teardown contract: a throwaway Excel instance LAUNCHED BY THIS
HARNESS (get_app() auto-launch, when GetActiveObject fails to attach to an
already-running Excel) is quit on exit, on top of the scratch-workbook
cleanup below. A pre-existing Excel instance is never touched — see
tests/_excel_lifecycle.py for the PID-ownership check.

Scenarios (see
docs/superpowers/plans/2026-07-24-perf-round2-constant-factor.md Phase 0),
each median-of-N (N>=5), printed as µs/op (or ms) to stderr:
  (a) 1,000x single-property read `ws.Cells(1,1).Value`, late-bound vs
      early-bound (via `_rewrap_earlybound` — see below).
  (b) block read of a 10,000x10 range: `.Value` vs `.Value2`.
  (c) `excel_format` border=all on a 50x10 range (property-heavy write loop),
      wall-clock, through the actual `format_action` tool entry point.
  (d) fixed per-call floor: `_session.run_com(...)` round-trip x100.

Run this script WITHOUT THEPEXCEL_MCP_EARLYBIND set. It forces both bindings
itself: `_session.get_app()` gives the ordinary (late-bound, when the flag is
unset) cached Application, and `_rewrap_earlybound()` is called directly on
that same handle to get the early-bound counterpart for comparison — the
comparison does not depend on the server's opt-in flag.
"""

from __future__ import annotations

import atexit
import datetime
import os
import sys
import time

from _excel_lifecycle import ExcelInstanceGuard

try:
    from thepexcel_mcp.session import ExcelSession, _rewrap_earlybound
    from thepexcel_mcp.domains.format import format_action
except ImportError as e:
    print(f"Import failed: {e}", file=sys.stderr)
    sys.exit(1)

_session = ExcelSession()
_instance_guard = ExcelInstanceGuard(_session)


def _log(msg: str) -> None:
    print(msg, file=sys.stderr)


def _median(samples: list) -> float:
    s = sorted(samples)
    n = len(s)
    mid = n // 2
    return s[mid] if n % 2 else (s[mid - 1] + s[mid]) / 2


def _col_letter(n: int) -> str:
    """1 -> A, 10 -> J, 27 -> AA, ..."""
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


# ── Bootstrap (mirrors smoke_com.py's own-instance pattern) ────────────────────

def check_excel_running() -> bool:
    """Return True if Excel is reachable (or was auto-launched).

    Snapshots pre-existing Excel PIDs BEFORE this first COM touch, then
    claims ownership right after — see tests/_excel_lifecycle.py.
    """
    _instance_guard.snapshot()
    try:
        _session.run_com(_session.get_app)
    except Exception as e:
        _log(f"ERROR: Cannot reach Excel: {e}")
        _log("Start Excel manually, or set THEPEXCEL_MCP_AUTOLAUNCH=1.")
        return False
    _instance_guard.claim()
    return True


def _new_wb() -> str:
    def _do():
        app = _session.get_app()
        wb = app.Workbooks.Add()
        return wb.Name
    return _session.run_com(_do)


def _add_sheet(wb_name: str, name: str) -> None:
    def _do():
        wb = _session.get_workbook(wb_name)
        ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        ws.Name = name
    _session.run_com(_do)


def _close_wb(wb_name: str) -> None:
    def _do():
        wb = _session.get_workbook(wb_name)
        wb.Close(SaveChanges=False)
    try:
        _session.run_com(_do)
        _log(f"[cleanup] closed scratch workbook '{wb_name}' (not saved)")
    except Exception as e:
        _log(f"[cleanup] warning: could not close '{wb_name}': {e}")


def _seed_block_data(wb_name: str, sheet_name: str, rows: int, cols: int) -> None:
    """Seed a rows x cols block with a mix of number / text / date values,
    cycling by column so scenario (b) exercises Value's per-cell boxing
    (VT_DATE / VT_CY-shaped) instead of an all-numeric range."""
    def _do():
        ws = _session.get_sheet(sheet_name, wb_name)
        base_date = datetime.datetime(2024, 1, 1)
        block = []
        for r in range(rows):
            row = []
            for c in range(cols):
                m = c % 3
                if m == 0:
                    row.append(r * 1.5 + c)
                elif m == 1:
                    row.append(f"text-{r}-{c}")
                else:
                    row.append(base_date + datetime.timedelta(days=r % 365))
            block.append(tuple(row))
        target = ws.Range(ws.Cells(1, 1), ws.Cells(rows, cols))
        target.Value = tuple(block)
    _session.run_com(_do)


# ── (a) property read: late-bound vs early-bound ────────────────────────────

def bench_property_read(wb_name: str, n: int = 1000, reps: int = 5) -> None:
    def _do_late():
        ws = _session.get_sheet("Sheet1", wb_name)
        cell = ws.Cells(1, 1)
        cell.Value = 42
        samples = []
        for _ in range(reps):
            t0 = time.perf_counter()
            for _ in range(n):
                _ = cell.Value
            samples.append((time.perf_counter() - t0) / n)
        return _median(samples)

    def _do_early():
        app = _session.get_app()
        early = _rewrap_earlybound(app)  # zero-activation rewrap, same pointer
        wb = early.Workbooks(wb_name)
        ws = wb.Sheets("Sheet1")
        cell = ws.Cells(1, 1)
        cell.Value = 42
        samples = []
        for _ in range(reps):
            t0 = time.perf_counter()
            for _ in range(n):
                _ = cell.Value
            samples.append((time.perf_counter() - t0) / n)
        return _median(samples)

    late_s = _session.run_com(_do_late)
    late_us = late_s * 1e6
    try:
        early_s = _session.run_com(_do_early)
        early_us = early_s * 1e6
        ratio = late_s / early_s if early_s else float("nan")
        _log(
            f"(a) property read x{n} (median of {reps}): "
            f"late={late_us:.2f}us/op early={early_us:.2f}us/op "
            f"ratio(late/early)={ratio:.2f}x"
        )
    except Exception as e:
        _log(f"(a) property read x{n}: late={late_us:.2f}us/op early=FAILED ({e})")


# ── (b) block read: .Value vs .Value2 ───────────────────────────────────────

def bench_block_read(wb_name: str, sheet_name: str, rows: int = 10_000, cols: int = 10, reps: int = 5) -> None:
    def _read_value():
        ws = _session.get_sheet(sheet_name, wb_name)
        rng = ws.Range(ws.Cells(1, 1), ws.Cells(rows, cols))
        samples = []
        for _ in range(reps):
            t0 = time.perf_counter()
            _ = rng.Value
            samples.append(time.perf_counter() - t0)
        return _median(samples)

    def _read_value2():
        ws = _session.get_sheet(sheet_name, wb_name)
        rng = ws.Range(ws.Cells(1, 1), ws.Cells(rows, cols))
        samples = []
        for _ in range(reps):
            t0 = time.perf_counter()
            _ = rng.Value2
            samples.append(time.perf_counter() - t0)
        return _median(samples)

    value_s = _session.run_com(_read_value)
    value2_s = _session.run_com(_read_value2)
    ratio = value_s / value2_s if value2_s else float("nan")
    _log(
        f"(b) block read {rows}x{cols} (median of {reps}): "
        f"Value={value_s * 1000:.2f}ms Value2={value2_s * 1000:.2f}ms "
        f"ratio(Value/Value2)={ratio:.2f}x"
    )


# ── (c) excel_format border=all, wall-clock through the tool entry point ───

def bench_border_format(wb_name: str, sheet_name: str, rows: int = 50, cols: int = 10, reps: int = 5) -> None:
    range_str = f"A1:{_col_letter(cols)}{rows}"
    samples = []
    for _ in range(reps):
        t0 = time.perf_counter()
        format_action(
            "border",
            range=range_str,
            sheet=sheet_name,
            workbook=wb_name,
            border_sides="all",
            border_style="continuous",
            border_weight="thin",
            border_color="#000000",
        )
        samples.append(time.perf_counter() - t0)
    ms = _median(samples) * 1000
    _log(f"(c) format border=all {rows}x{cols} (median of {reps}): {ms:.2f}ms wall-clock")


# ── (d) fixed per-call floor ─────────────────────────────────────────────────

def bench_percall_floor(n: int = 100, reps: int = 5) -> None:
    # Warm the cache once outside the timed loop.
    _session.run_com(lambda: (_session.get_app(), None)[1])
    samples = []
    for _ in range(reps):
        t0 = time.perf_counter()
        for _ in range(n):
            _session.run_com(lambda: (_session.get_app(), None)[1])
        samples.append((time.perf_counter() - t0) / n)
    per_call_us = _median(samples) * 1e6
    _log(f"(d) per-call floor x{n} (median of {reps} reps): {per_call_us:.2f}us/call")


# ── Main ─────────────────────────────────────────────────────────────────────

def main() -> None:
    _log("=" * 60)
    _log("  ThepExcelMCP perf bench (Phase 0 — constant-factor round 2)")
    _log(f"  AUTOLAUNCH={os.environ.get('THEPEXCEL_MCP_AUTOLAUNCH', '0')}")
    _log(f"  EARLYBIND={os.environ.get('THEPEXCEL_MCP_EARLYBIND', '0')} (should be unset for this script)")
    _log("=" * 60)

    atexit.register(_instance_guard.teardown)

    if not check_excel_running():
        sys.exit(1)

    wb_name = None
    try:
        wb_name = _new_wb()
        _log(f"Bench scratch workbook: {wb_name}")

        _add_sheet(wb_name, "Bench2")
        _seed_block_data(wb_name, "Bench2", rows=10_000, cols=10)
        _add_sheet(wb_name, "Bench3")

        bench_property_read(wb_name)
        bench_block_read(wb_name, "Bench2")
        bench_border_format(wb_name, "Bench3")
        bench_percall_floor()

        _log("=" * 60)
        _log("Bench complete.")
        _log("=" * 60)
    finally:
        if wb_name:
            _close_wb(wb_name)
        _instance_guard.teardown()


if __name__ == "__main__":
    main()
