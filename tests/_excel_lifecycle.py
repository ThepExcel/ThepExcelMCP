"""Shared Excel-instance lifecycle helper for the live COM harnesses
(smoke_com.py, bench_com.py).

Both harnesses call ``_session.get_app()``, which attaches to Excel via
``GetActiveObject`` or auto-launches a new visible instance if attach fails.
On this machine ``GetActiveObject`` intermittently misses the user's already
-running Excel, so a harness run can silently spawn a duplicate throwaway
"Book1" Excel that is never ``Quit()``'d and lingers after the run.

SAFETY INVARIANT (read twice): this module must NEVER quit an Excel instance
the harness did not itself launch. Ownership is decided purely by PID-
snapshot diffing -- an EXCEL.EXE process already running *before* the
harness's first COM call belongs to the user (or some other pre-existing
session) and is never closed or terminated. Only a PID that first appears
*after* the harness reaches Excel (i.e. one the harness itself caused to
spawn) is eligible for teardown. When ownership can't be determined, the
instance is left alone.

All COM access (Hwnd lookup, Workbooks.Close, Quit) happens on the harness's
own ``ExcelSession`` STA worker thread via ``session.run_com(...)`` -- raw
COM objects never cross the thread boundary, only plain ints/None.
"""

from __future__ import annotations

import subprocess
import time


def running_excel_pids() -> set[int]:
    """Return the PIDs of every currently-running EXCEL.EXE process.

    Dependency-free: shells out to `tasklist` and parses its CSV output
    (pywin32 + stdlib only -- no psutil).
    """
    try:
        proc = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=10,
        )
    except Exception:
        return set()

    pids: set[int] = set()
    for line in proc.stdout.splitlines():
        parts = [p.strip('"') for p in line.strip().split(",")]
        if len(parts) >= 2 and parts[0].upper() == "EXCEL.EXE":
            try:
                pids.add(int(parts[1]))
            except ValueError:
                pass
    return pids


class ExcelInstanceGuard:
    """Snapshot pre-existing Excel PIDs, claim the harness's own instance,
    and quit ONLY that instance on teardown -- never anything from the
    pre-snapshot.

    Usage::

        guard = ExcelInstanceGuard(_session)
        guard.snapshot()             # BEFORE the harness's first get_app()
        ...                          # (harness reaches Excel here)
        guard.claim()                # once, right after reaching Excel
        ...
        guard.teardown()             # in a top-level finally / atexit
    """

    def __init__(self, session) -> None:
        self._session = session
        self._pre_pids: set[int] = set()
        self._owned_pid: int | None = None
        self._claimed = False

    def snapshot(self) -> None:
        """Record the set of EXCEL.EXE PIDs already running. Call this
        BEFORE any code path that might call ``_session.get_app()``."""
        self._pre_pids = running_excel_pids()

    def claim(self) -> None:
        """Resolve which Excel instance the harness is talking to, and
        whether it's one the harness caused to spawn (i.e. its PID was
        NOT in the pre-snapshot). Call once, after the harness has
        successfully reached Excel."""
        if self._claimed:
            return
        self._claimed = True

        def _do():
            app = self._session.get_app()
            try:
                import win32process

                _, pid = win32process.GetWindowThreadProcessId(app.Hwnd)
                return pid
            except Exception:
                return None

        try:
            pid = self._session.run_com(_do)
        except Exception:
            pid = None

        if pid is not None and pid not in self._pre_pids:
            self._owned_pid = pid

    def teardown(self) -> None:
        """Quit the harness-owned instance, if any. A no-op when the
        harness attached to a pre-existing Excel instead of spawning its
        own. Safe to call multiple times.

        SAFETY: ``get_app()`` returns a CACHED app handle -- if the owned
        instance died mid-run (or the cache otherwise went stale), calling
        ``get_app()`` here can silently re-attach to a DIFFERENT Excel (the
        user's real one) via GetActiveObject. To never Close/Quit the wrong
        instance, this re-resolves the CURRENT app's PID and only proceeds
        with Close/Quit when it still equals the claimed ``_owned_pid``. Any
        mismatch or unknown identity aborts with no COM mutation at all.
        """
        pid = self._owned_pid
        if pid is None:
            return

        def _do() -> bool:
            app = self._session.get_app()
            try:
                import win32process

                _, cur_pid = win32process.GetWindowThreadProcessId(app.Hwnd)
            except Exception:
                return False  # can't confirm identity -> do NOT quit anything
            if cur_pid != pid:
                return False  # re-resolved to a DIFFERENT instance -> leave it alone

            # Confirmed: still the same owned instance -> safe to close+quit.
            try:
                app.DisplayAlerts = False
            except Exception:
                pass
            try:
                while app.Workbooks.Count > 0:
                    app.Workbooks(1).Close(SaveChanges=False)
            except Exception:
                pass
            try:
                app.Quit()
            except Exception:
                pass
            return True

        try:
            quit_attempted = self._session.run_com(_do)
        except Exception:
            quit_attempted = False

        if not quit_attempted:
            # Identity check failed or mismatched -- nothing was touched.
            # Leave _owned_pid as-is; there's nothing more this call can
            # safely do (a last-resort kill would target a PID we could no
            # longer positively identify as ours).
            return

        # Confirm it's actually gone; poll briefly before the last-resort kill.
        for _ in range(10):
            if pid not in running_excel_pids():
                self._owned_pid = None
                return
            time.sleep(0.5)

        if pid in running_excel_pids():
            # Last resort -- terminate this SAME owned pid only, never a pid
            # from the pre-existing snapshot. Gated on quit_attempted (the
            # identity check just confirmed, moments ago, that this pid WAS
            # the owned instance) for the same reasoning as above.
            try:
                import win32api
                import win32con

                handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
                win32api.TerminateProcess(handle, 0)
                win32api.CloseHandle(handle)
            except Exception:
                pass

        self._owned_pid = None
