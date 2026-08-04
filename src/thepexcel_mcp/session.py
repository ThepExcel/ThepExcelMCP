"""ExcelSession — STA COM worker thread + workbook/sheet routing helpers.

Architecture (Phase 3 hardening)
---------------------------------
Excel COM objects are STA (Single-Threaded Apartment). FastMCP may dispatch
tool handlers on arbitrary threads, causing RPC_E_WRONG_THREAD crashes.

Solution: a single background thread owns the STA apartment.
  - Thread calls pythoncom.CoInitialize() once at startup.
  - All COM creation AND usage happens only on that thread.
  - Public API: run_com(fn, *args, **kwargs) submits a callable to the
    thread via queue.Queue and returns the result (or re-raises an exception)
    via concurrent.futures.Future.
  - Per-call timeout defaults to 120 s (THEPEXCEL_MCP_COM_TIMEOUT env var).

ExcelSession methods still expose get_app / get_workbook / get_sheet as
convenience helpers — but they now run INSIDE the worker (called from
domain callables passed to run_com). The old per-call CoInitialize() is gone.

ROT fallback
------------
win32com.GetActiveObject returns only the first registered Excel instance.
When a requested workbook is not found there, _enum_rot_workbooks() scans
the Running Object Table to find it in another instance.

excel_guard context manager
----------------------------
Sets app.DisplayAlerts = False inside save/close/delete operations (guards
against modal dialogs on the worker thread) and restores True afterward.
Interactive=False is NOT set globally — that would lock the user out of
their own Excel.

wait_calculation
----------------
Polls app.CalculationState with pythoncom.PumpWaitingMessages() to avoid
deadlocking the STA message queue. Never a bare sleep loop.

THEPEXCEL_MCP_EARLYBIND (default ON — kill-switch, flipped 2026-07-24)
-----------------------------------------------------------------------
get_app() swaps the late-bound Application handle for an early-bound
(win32com.client.gencache) wrapper of the SAME running instance, via a
zero-activation rewrap (see ``_rewrap_earlybound``): the makepy cache is
generated from the already-attached object's own typelib (no COM activation,
so it cannot spin up a duplicate Excel process), then the same underlying
IDispatch pointer is re-wrapped early-bound. Set THEPEXCEL_MCP_EARLYBIND to a
falsy value (0/false/no/off) to force late binding (kill-switch); any
rewrap failure also silently falls back to an explicitly dynamic wrapper.
Flipped to default-ON on 2026-07-24. Corrected 2026-08-04 benchmarks that
force both wrapper types around the same IDispatch pointer measured modest
gains (9% faster property reads in the latest run); live effect smoke is
the compatibility oracle. Corrupt makepy recovery clears both the on-disk
cache and loaded ``win32com.gen_py.*`` modules so retry works in-process. See
docs/superpowers/plans/2026-07-24-perf-round2-constant-factor.md.
"""

from __future__ import annotations

import contextlib
import importlib
import os
import queue
import shutil
import sys
import threading
import time
from concurrent.futures import Future
from typing import Any, Callable

import pythoncom
import win32com
import win32com.client
from fastmcp.exceptions import ToolError

# Default per-call timeout in seconds; override via env var.
_DEFAULT_TIMEOUT = int(os.environ.get("THEPEXCEL_MCP_COM_TIMEOUT", "120"))

# Slow-call diagnostic threshold in seconds; override via env var. Logged to
# stderr only (stdout is the stdio transport — never write diagnostics there).
_SLOW_LOG_THRESHOLD = float(os.environ.get("THEPEXCEL_MCP_SLOW_LOG_S", "5.0"))

# xlCalculationStateIdle = 0 (Excel.XlCalculationState enum)
_XL_CALCULATION_IDLE = 0
_CALC_POLL_INTERVAL = 0.05  # seconds between CalculationState polls

# XlCalculation.xlCalculationManual — used by bulk_guard(). Never hardcode the
# "automatic" counterpart on restore; bulk_guard restores whatever the caller's
# Calculation mode actually was before the guard ran.
_XL_CALCULATION_MANUAL = -4135

# Cached Application handle. Accessed only from the single STA worker thread
# (see _COMWorker._run), so no lock is needed despite being module-level.
_cached_app: win32com.client.CDispatch | None = None


class _COMWorker:
    """Background STA thread that owns all COM object lifetimes.

    Lazily started on first run_com() call (so import is cheap on non-Windows).
    """

    def __init__(self) -> None:
        self._queue: queue.Queue[tuple[Future, Callable, tuple, dict] | None] = queue.Queue()
        self._thread: threading.Thread | None = None
        self._lock = threading.Lock()

    def _ensure_started(self) -> None:
        if self._thread is not None and self._thread.is_alive():
            return
        with self._lock:
            if self._thread is not None and self._thread.is_alive():
                return
            t = threading.Thread(target=self._run, name="excel-com-worker", daemon=True)
            t.start()
            self._thread = t

    def _run(self) -> None:
        """Main loop — runs entirely on the dedicated STA thread."""
        pythoncom.CoInitialize()
        try:
            while True:
                item = self._queue.get()
                if item is None:
                    break  # shutdown signal
                future, fn, args, kwargs = item
                start = time.perf_counter()
                try:
                    result = fn(*args, **kwargs)
                    future.set_result(result)
                except Exception as exc:
                    future.set_exception(exc)
                finally:
                    duration = time.perf_counter() - start
                    if duration > _SLOW_LOG_THRESHOLD:
                        name = getattr(fn, "__qualname__", repr(fn))
                        print(
                            f"[thepexcel-mcp] slow COM call: {name} took {duration:.1f}s",
                            file=sys.stderr,
                        )
        finally:
            pythoncom.CoUninitialize()

    def submit(self, fn: Callable, *args: Any, **kwargs: Any) -> Any:
        """Submit *fn* to the COM worker and return its result (blocking).

        Raises ToolError on timeout, re-raises any exception from fn.
        """
        self._ensure_started()
        future: Future = Future()
        self._queue.put((future, fn, args, kwargs))
        try:
            return future.result(timeout=_DEFAULT_TIMEOUT)
        except TimeoutError:
            raise ToolError(
                f"Excel COM call timed out after {_DEFAULT_TIMEOUT}s. "
                "Excel may be busy or showing a dialog. "
                "Dismiss any open dialogs, then retry."
            )


# Module-level singleton worker
_worker = _COMWorker()


@contextlib.contextmanager
def excel_guard(app):
    """Context manager: suppress Excel modal dialogs during risky COM calls.

    Sets DisplayAlerts=False and restores True on exit.
    Use around save / close / delete operations only — not globally.
    (Interactive=False is intentionally NOT set here; that locks the user out.)
    """
    app.DisplayAlerts = False
    try:
        yield
    finally:
        app.DisplayAlerts = True


@contextlib.contextmanager
def bulk_guard(app):
    """Context manager: suppress screen/event churn for a batch of COM writes.

    Saves ScreenUpdating / EnableEvents / Calculation, sets them to
    False / False / xlCalculationManual, then restores the SAVED values on
    exit (never hardcodes xlCalculationAutomatic — respects whatever mode the
    caller/user actually had). If we put Calculation into Manual (i.e. it
    wasn't already Manual), forces one Calculate() before restoring the mode
    so writes made during the guard aren't left uncalculated.

    Use only around genuinely bulk write loops (table append_rows, an
    all/inside border loop, find/replace, the load_to_table sheet-rebuild
    step) — never around PQ refresh / datamodel / cube calls, which have
    their own documented deadlock and async-calculation semantics.
    """
    prev_screen_updating = app.ScreenUpdating
    prev_enable_events = app.EnableEvents
    prev_calculation = app.Calculation
    app.ScreenUpdating = False
    app.EnableEvents = False
    app.Calculation = _XL_CALCULATION_MANUAL
    try:
        yield
    finally:
        if prev_calculation != _XL_CALCULATION_MANUAL:
            try:
                app.Calculate()
            except Exception:
                pass
        app.Calculation = prev_calculation
        app.EnableEvents = prev_enable_events
        app.ScreenUpdating = prev_screen_updating


def wait_calculation(app, timeout: float = 60.0) -> None:
    """Block until Excel finishes calculating (CalculationState == Idle).

    Pumps the STA message queue between polls to avoid deadlock.
    Raises ToolError on timeout.
    """
    deadline = time.monotonic() + timeout
    while app.CalculationState != _XL_CALCULATION_IDLE:
        pythoncom.PumpWaitingMessages()
        if time.monotonic() > deadline:
            raise ToolError(
                "Excel is still calculating after "
                f"{timeout:.0f}s. Retry when calculation completes."
            )
        time.sleep(_CALC_POLL_INTERVAL)


def _rewrap_earlybound(app: win32com.client.CDispatch) -> win32com.client.CDispatch:
    """Construct an early-bound wrapper of *app* — zero-activation rewrap.

    The old approach (`gencache.EnsureDispatch("Excel.Application")`) *activates*
    COM — it can spin up a brand new Excel process instead of attaching to the
    running one (hence the old Hwnd-match-or-discard dance). This construction
    never activates anything:

    1. Read the typelib straight off the already-attached late-bound object
       (``app._oleobj_.GetTypeInfo()`` → ``GetContainingTypeLib()`` →
       ``GetLibAttr()``) — version-exact for whatever Office build is actually
       running, no hard-coded typelib version.
    2. ``gencache.EnsureModule(guid, lcid, major, minor)`` generates/loads the
       makepy module for that exact typelib (no COM activation).
    3. ``win32com.client.Dispatch(app._oleobj_)`` re-wraps the SAME underlying
       IDispatch pointer — this consults gencache and returns the generated
       early-bound class around the same running instance. Child objects
       (Workbooks -> Workbook -> Range) come back early-bound automatically.

    Raises on any failure (missing typelib info, gencache errors, etc.) —
    callers decide the fallback; this function never silently returns the
    late-bound app itself, so a caller checking "did I get early-bound back"
    can rely on a clean exception instead of an ambiguous same-type return.
    """
    ti = app._oleobj_.GetTypeInfo()
    tlb, _idx = ti.GetContainingTypeLib()
    la = tlb.GetLibAttr()  # (guid, lcid, ?, major, minor, ...) — verified via a
    # live Application on this machine 2026-07-24: (guid, 0, 1, 1, 9, 8) with
    # gencache module name ...x0x1x9 → la[3]/la[4] are indeed major/minor.
    win32com.client.gencache.EnsureModule(la[0], la[1], la[3], la[4])
    return win32com.client.Dispatch(app._oleobj_)


def _force_latebound(app: win32com.client.CDispatch) -> win32com.client.CDispatch:
    """Wrap *app* dynamically around the same underlying COM pointer.

    ``GetActiveObject`` may return a generated ``gen_py`` wrapper whenever a
    makepy cache exists, so merely skipping the early-bind rewrap does not
    actually force late binding. The kill-switch and benchmark both need an
    explicit dynamic wrapper to make the selected mode real and measurable.
    """
    if type(app).__module__ == "win32com.client.dynamic":
        return app
    return win32com.client.dynamic.Dispatch(app._oleobj_)


def _maybe_earlybind(app: win32com.client.CDispatch) -> win32com.client.CDispatch:
    """Default ON as of 2026-07-24; THEPEXCEL_MCP_EARLYBIND is now a KILL-SWITCH.

    Swap a late-bound Application handle for an early-bound
    (win32com.client.gencache) wrapper of the SAME running instance via
    ``_rewrap_earlybound`` (zero-activation rewrap — see its docstring).
    The kill-switch explicitly constructs a dynamic wrapper because
    ``GetActiveObject`` itself may return a generated wrapper when ``gen_py``
    exists. Any early-bind failure also falls back to that explicit dynamic
    wrapper. This call never activates COM or spawns a process.

    Corrected 2026-08-04 benchmarks measured modest gains (9% faster property
    reads in the latest run). The live smoke suite remains
    the effect oracle — see
    docs/superpowers/plans/2026-07-24-perf-round2-constant-factor.md.
    Set THEPEXCEL_MCP_EARLYBIND=0 (or false/no/off) to force late binding.
    """
    flag = os.environ.get("THEPEXCEL_MCP_EARLYBIND", "1").strip().lower()
    if flag in ("0", "false", "no", "off"):
        try:
            return _force_latebound(app)
        except Exception:
            return app  # keep Excel reachable if dynamic wrapping itself fails
    try:
        return _rewrap_earlybound(app)
    except Exception:
        # Never replace a handle that already works with an unproven one on an
        # error path: *app* is the object GetActiveObject/_launch just returned
        # and is known reachable. Constructing a fresh dynamic wrapper here
        # (the 2026-08-04 behaviour) turns a recoverable rewrap failure into a
        # hard break when that wrapper's .Range/.Name access fails.
        return app


def _clear_gen_py_cache() -> None:
    """Clear a corrupt pywin32 makepy cache for an in-process retry.

    Deleting the on-disk ``gen_py`` directory alone is insufficient: the
    broken generated module remains in ``sys.modules`` and the immediate
    retry imports that same object again. Evict generated children and
    invalidate import caches so recovery does not require restarting the
    Python process.

    Deliberately does NOT call ``gencache.Rebuild()``: that regenerates the
    makepy cache for EVERY registered typelib on the machine, not just the one
    this process broke — a machine-wide side effect taken on a recovery path,
    and the suspected trigger of the 2026-08-04 live-COM incident. Dispatch
    regenerates the exact typelib lazily on next use, so the eviction above is
    the whole cure.
    """
    shutil.rmtree(win32com.__gen_path__, ignore_errors=True)
    for module_name in list(sys.modules):
        if module_name.startswith("win32com.gen_py."):
            sys.modules.pop(module_name, None)
    importlib.invalidate_caches()


def _enum_rot_workbooks(name: str):
    """Scan the Running Object Table for a workbook named *name*.

    GetActiveObject returns only the first Excel instance registered in the ROT.
    When a workbook lives in a second/third instance, enumerate all ROT monikers
    looking for the workbook by name.

    Returns the Workbook COM object if found, else None.
    """
    try:
        rot = pythoncom.GetRunningObjectTable()
        enum = rot.EnumRunning()
        while True:
            monikers = enum.Next(1)
            if not monikers:
                break
            moniker = monikers[0]
            ctx = pythoncom.CreateBindCtx(0)
            try:
                display = moniker.GetDisplayName(ctx, None)
            except Exception:
                continue
            # Excel workbook monikers look like "C:\path\to\Book.xlsx"
            if not display.endswith((".xlsx", ".xlsm", ".xlsb", ".xls", ".xlam")):
                continue
            try:
                obj = rot.GetObject(moniker)
                # QueryInterface to IDispatch so win32com can wrap it
                dispatch = win32com.client.Dispatch(
                    obj.QueryInterface(pythoncom.IID_IDispatch)
                )
                # dispatch is the Workbook itself in Excel's ROT entries
                if dispatch.Name == name or dispatch.Name.lower() == name.lower():
                    return dispatch
                # Might be the Application; try Workbooks collection
                try:
                    return dispatch.Workbooks(name)
                except Exception:
                    pass
            except Exception:
                continue
    except Exception:
        pass
    return None


class ExcelSession:
    """Thin wrapper — resolves Excel COM objects via the STA worker thread.

    Domain callables passed to run_com() receive no pre-resolved objects;
    they call get_app() / get_workbook() / get_sheet() themselves INSIDE
    the worker (since those helpers access COM objects).

    Usage in domain modules
    -----------------------
    All action functions must be called via _session.run_com(...):

        def some_action(workbook, ...):
            def _do(wb):
                return ...
            return _session.run_com(lambda: _do(_session.get_workbook(workbook)))

    Or, equivalently, wrap the entire domain function body:

        def workbook_action(action, workbook=None):
            def _impl():
                ...use get_app/get_workbook/get_sheet here...
            return _session.run_com(_impl)
    """

    def run_com(self, fn: Callable, *args: Any, **kwargs: Any) -> Any:
        """Execute *fn* on the dedicated STA COM worker thread."""
        return _worker.submit(fn, *args, **kwargs)

    def get_app(self) -> win32com.client.CDispatch:
        """Return the running Excel.Application COM object, auto-launching one
        if none is found.

        Auto-launch is ON by default so tools "just work" when Excel is closed;
        opt out with a falsy ``THEPEXCEL_MCP_AUTOLAUNCH`` (0/false/no/off).

        Also self-heals a corrupt win32com early-binding cache: a stale
        ``gen_py`` cache raises ``AttributeError: module '...' has no attribute
        'CLSIDToClassMap'`` on Dispatch/GetActiveObject *even when Excel is
        running fine* — which otherwise looks like "Excel is not running". On
        that error we clear the cache and retry once.

        Caches the resolved Application handle across calls (module-level,
        worker-thread-only — safe without a lock since only the single STA
        worker thread ever calls this). Each call does a cheap liveness probe
        (``app.Name``) against the cache first; any exception drops the stale
        cache and falls through to the normal attach/self-heal/auto-launch path.

        Must be called from within the COM worker thread (i.e. inside a
        callable passed to run_com).
        """
        global _cached_app

        if _cached_app is not None:
            try:
                _ = _cached_app.Name  # cheap liveness probe
                return _cached_app
            except Exception:
                _cached_app = None  # stale handle (Excel closed/crashed) — re-resolve

        def _attach() -> win32com.client.CDispatch:
            return win32com.client.GetActiveObject("Excel.Application")

        def _launch() -> win32com.client.CDispatch:
            app = win32com.client.Dispatch("Excel.Application")
            app.Visible = True
            # A freshly-launched Excel has no workbook; add a blank one so
            # get_workbook()/ActiveWorkbook is immediately usable.
            try:
                if app.Workbooks.Count == 0:
                    app.Workbooks.Add()
            except Exception:
                pass
            return app

        app: win32com.client.CDispatch | None = None

        # 1) Attach to a running instance. A corrupt gen_py cache can make this
        #    throw AttributeError even when Excel IS running → clear + retry so
        #    we reattach to the user's Excel instead of spawning a duplicate.
        try:
            app = _attach()
        except AttributeError:
            _clear_gen_py_cache()
            try:
                app = _attach()
            except Exception:
                app = None  # genuinely not running → fall through to auto-launch
        except Exception:
            app = None  # not running / ROT miss → fall through to auto-launch

        if app is None:
            # 2) Auto-launch (default on).
            autolaunch = os.environ.get("THEPEXCEL_MCP_AUTOLAUNCH", "1").strip().lower()
            if autolaunch in ("0", "false", "no", "off"):
                raise ToolError(
                    "Excel is not running and auto-launch is disabled "
                    "(THEPEXCEL_MCP_AUTOLAUNCH is falsy). Open Excel yourself, or "
                    "re-enable auto-launch (unset the var or set it to 1) so the "
                    "tool can open Excel for you."
                )
            try:
                app = _launch()
            except AttributeError:
                _clear_gen_py_cache()
                app = _launch()

        app = _maybe_earlybind(app)
        _cached_app = app
        return app

    def get_workbook(self, workbook: str | None = None) -> win32com.client.CDispatch:
        """Return a Workbook COM object.

        Args:
            workbook: Workbook name (e.g. "Sales.xlsx"). Pass None for active workbook.

        Must be called from within the COM worker thread.
        """
        app = self.get_app()
        if workbook:
            try:
                return app.Workbooks(workbook)
            except Exception:
                # ROT fallback: check other Excel instances
                wb = _enum_rot_workbooks(workbook)
                if wb is not None:
                    return wb
                available = [app.Workbooks(i + 1).Name for i in range(app.Workbooks.Count)]
                raise ToolError(
                    f"Workbook '{workbook}' not found. "
                    f"Open workbooks: {available or ['(none)']}"
                )
        wb = app.ActiveWorkbook
        if wb is None:
            raise ToolError("No active workbook — open an Excel file first.")
        return wb

    def get_sheet(
        self, name: str | None, workbook: str | None = None
    ) -> win32com.client.CDispatch:
        """Return a Worksheet COM object.

        Args:
            name: Sheet name. Pass None for active sheet.
            workbook: Workbook name. Pass None for active workbook.

        Must be called from within the COM worker thread.
        """
        wb = self.get_workbook(workbook)
        if name is None:
            return wb.ActiveSheet
        try:
            return wb.Sheets(name)
        except Exception:
            available = [wb.Sheets(i + 1).Name for i in range(wb.Sheets.Count)]
            raise ToolError(
                f"Sheet '{name}' not found. Available: {available}"
            )

    @staticmethod
    def wrap(exc: Exception, context: str = "") -> ToolError:
        """Convert a raw COM exception to an actionable ToolError."""
        prefix = f"{context}: " if context else ""
        return ToolError(f"{prefix}{exc}")
