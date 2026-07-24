"""Unit tests for tests/_excel_lifecycle.py — ExcelInstanceGuard's identity
re-check in teardown().

Locks the safety fix: teardown() must positively confirm the CURRENT app's
PID still equals the claimed `_owned_pid` immediately before any
Close/Quit/TerminateProcess. A cached ``get_app()`` handle can go stale
mid-run and silently re-attach to a DIFFERENT Excel instance (possibly the
user's real one) — on any PID mismatch (or unresolvable identity), teardown
must abort with zero COM mutation.

All COM/process boundaries are mocked — no real Excel needed, no pywin32
process handles opened.
"""

from __future__ import annotations

from unittest.mock import MagicMock, patch

from _excel_lifecycle import ExcelInstanceGuard


def _make_session(app_hwnd=123):
    """A MagicMock session whose run_com() is a transparent passthrough
    (fn executed synchronously) and whose get_app() returns a MagicMock
    app with the given Hwnd."""
    app = MagicMock()
    app.Hwnd = app_hwnd
    session = MagicMock()
    session.run_com.side_effect = lambda fn, *a, **kw: fn(*a, **kw)
    session.get_app.return_value = app
    return session, app


def test_teardown_aborts_on_pid_mismatch():
    """The resolved current-app PID (12345) no longer matches the claimed
    owned PID (999) — e.g. the owned instance died and get_app()'s cache
    re-resolved to a different (possibly the user's) Excel. teardown must
    NOT Close/Quit/Terminate anything, and must not even reach the
    post-quit gone-check poll."""
    session, app = _make_session()
    guard = ExcelInstanceGuard(session)
    guard._owned_pid = 999  # simulate a prior successful claim()

    with patch("win32process.GetWindowThreadProcessId", return_value=(0, 12345)):
        with patch("_excel_lifecycle.running_excel_pids") as mock_running_pids:
            with patch("win32api.OpenProcess") as mock_open_process:
                guard.teardown()

    assert app.Quit.call_count == 0, "teardown must NOT Quit() on PID mismatch"
    assert app.Workbooks.call_count == 0, "teardown must NOT touch Workbooks on PID mismatch"
    mock_open_process.assert_not_called()
    mock_running_pids.assert_not_called()  # confirms the abort short-circuits before the gone-check poll
    assert guard._owned_pid == 999  # left untouched — no claim of successful teardown


def test_teardown_aborts_when_identity_unresolvable():
    """Hwnd lookup itself raises (e.g. Hwnd invalid / process gone) — identity
    can't be confirmed at all, so teardown must treat that the same as a
    mismatch: no mutation."""
    session, app = _make_session()
    guard = ExcelInstanceGuard(session)
    guard._owned_pid = 999

    with patch("win32process.GetWindowThreadProcessId", side_effect=RuntimeError("boom")):
        with patch("_excel_lifecycle.running_excel_pids") as mock_running_pids:
            guard.teardown()

    assert app.Quit.call_count == 0
    assert app.Workbooks.call_count == 0
    mock_running_pids.assert_not_called()
    assert guard._owned_pid == 999


def test_teardown_proceeds_on_pid_match():
    """Contrast case: when the resolved current PID DOES match the owned
    PID, teardown closes any open workbooks and Quits — proving the
    identity check isn't just permanently aborting."""
    session, app = _make_session()
    app.Workbooks.Count = 0  # no workbooks to close
    guard = ExcelInstanceGuard(session)
    guard._owned_pid = 555

    with patch("win32process.GetWindowThreadProcessId", return_value=(0, 555)):
        with patch("_excel_lifecycle.running_excel_pids", return_value=set()):
            guard.teardown()

    assert app.Quit.call_count == 1
    assert guard._owned_pid is None


def test_teardown_noop_when_nothing_owned():
    """No prior claim() (or claim() found no owned instance) -> teardown is
    a pure no-op, touching neither COM nor the process table."""
    session, app = _make_session()
    guard = ExcelInstanceGuard(session)
    assert guard._owned_pid is None

    guard.teardown()

    session.run_com.assert_not_called()
    session.get_app.assert_not_called()
