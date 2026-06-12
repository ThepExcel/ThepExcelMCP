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
"""

from __future__ import annotations

import contextlib
import os
import queue
import threading
import time
from concurrent.futures import Future
from typing import Any, Callable

import pythoncom
import win32com.client
from fastmcp.exceptions import ToolError

# Default per-call timeout in seconds; override via env var.
_DEFAULT_TIMEOUT = int(os.environ.get("THEPEXCEL_MCP_COM_TIMEOUT", "120"))

# xlCalculationStateIdle = 0 (Excel.XlCalculationState enum)
_XL_CALCULATION_IDLE = 0
_CALC_POLL_INTERVAL = 0.05  # seconds between CalculationState polls


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
                try:
                    result = fn(*args, **kwargs)
                    future.set_result(result)
                except Exception as exc:
                    future.set_exception(exc)
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
        """Return the running Excel.Application COM object.

        Must be called from within the COM worker thread (i.e. inside a
        callable passed to run_com).
        """
        try:
            return win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            if os.environ.get("THEPEXCEL_MCP_AUTOLAUNCH") == "1":
                app = win32com.client.Dispatch("Excel.Application")
                app.Visible = True
                return app
            raise ToolError(
                "Excel is not running — open Excel first, then retry. "
                "Or set THEPEXCEL_MCP_AUTOLAUNCH=1 to auto-launch."
            )

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
