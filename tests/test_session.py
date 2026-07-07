"""Unit tests for session.py — app caching, bulk_guard, early-bind, slow-call log.

No Excel required — win32com calls are patched. These tests exercise the
Windows-only ``thepexcel_mcp.session`` module directly (not via a domain
module's ``_session`` mock), since the behavior under test lives in
ExcelSession.get_app() / bulk_guard() / _maybe_earlybind() themselves.
"""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock

from fastmcp.exceptions import ToolError

import thepexcel_mcp.session as session_mod
from thepexcel_mcp.session import ExcelSession, bulk_guard, _maybe_earlybind


# ── get_app() caching ──────────────────────────────────────────────────────────

class _DeadApp:
    """A COM handle whose .Name access raises — simulates a closed/crashed Excel."""

    @property
    def Name(self):
        raise Exception("RPC server unavailable")


class TestGetAppCaching:
    def setup_method(self):
        session_mod._cached_app = None

    def teardown_method(self):
        session_mod._cached_app = None

    def test_returns_cached_app_without_reattaching(self, monkeypatch):
        cached = MagicMock()
        cached.Name = "Excel"
        session_mod._cached_app = cached

        mock_get_active = MagicMock(side_effect=AssertionError("must not attach when cache is alive"))
        monkeypatch.setattr(session_mod.win32com.client, "GetActiveObject", mock_get_active)

        result = ExcelSession().get_app()

        assert result is cached
        mock_get_active.assert_not_called()

    def test_stale_cache_is_dropped_and_reattached(self, monkeypatch):
        session_mod._cached_app = _DeadApp()

        fresh = MagicMock()
        fresh.Name = "Excel"
        monkeypatch.setattr(session_mod.win32com.client, "GetActiveObject", MagicMock(return_value=fresh))
        monkeypatch.delenv("THEPEXCEL_MCP_EARLYBIND", raising=False)

        result = ExcelSession().get_app()

        assert result is fresh
        assert session_mod._cached_app is fresh

    def test_fresh_attach_is_cached_for_next_call(self, monkeypatch):
        fresh = MagicMock()
        fresh.Name = "Excel"
        monkeypatch.setattr(session_mod.win32com.client, "GetActiveObject", MagicMock(return_value=fresh))
        monkeypatch.delenv("THEPEXCEL_MCP_EARLYBIND", raising=False)

        sess = ExcelSession()
        first = sess.get_app()
        assert first is fresh
        assert session_mod._cached_app is fresh

        mock_get_active_2 = MagicMock(side_effect=AssertionError("must not re-attach — cache should serve this call"))
        monkeypatch.setattr(session_mod.win32com.client, "GetActiveObject", mock_get_active_2)

        second = sess.get_app()
        assert second is fresh
        mock_get_active_2.assert_not_called()

    def test_autolaunch_disabled_raises_when_not_attached(self, monkeypatch):
        monkeypatch.setattr(
            session_mod.win32com.client, "GetActiveObject",
            MagicMock(side_effect=Exception("not running")),
        )
        monkeypatch.setenv("THEPEXCEL_MCP_AUTOLAUNCH", "0")

        with pytest.raises(ToolError, match="auto-launch is disabled"):
            ExcelSession().get_app()


# ── bulk_guard ─────────────────────────────────────────────────────────────────

class TestBulkGuard:
    def test_sets_manual_calc_and_restores_original_on_exit(self):
        app = MagicMock()
        app.ScreenUpdating = True
        app.EnableEvents = True
        app.Calculation = -4105  # arbitrary "not manual" value (never hardcode xlCalculationAutomatic)

        with bulk_guard(app):
            assert app.ScreenUpdating is False
            assert app.EnableEvents is False
            assert app.Calculation == -4135  # xlCalculationManual

        assert app.ScreenUpdating is True
        assert app.EnableEvents is True
        assert app.Calculation == -4105  # restored to the SAVED value, not a hardcoded constant
        app.Calculate.assert_called_once()

    def test_no_forced_calculate_when_already_manual(self):
        app = MagicMock()
        app.ScreenUpdating = True
        app.EnableEvents = True
        app.Calculation = -4135  # already manual — guard shouldn't force a recalc

        with bulk_guard(app):
            pass

        app.Calculate.assert_not_called()
        assert app.Calculation == -4135

    def test_restores_on_exception(self):
        app = MagicMock()
        app.ScreenUpdating = True
        app.EnableEvents = True
        app.Calculation = -4105

        with pytest.raises(ValueError):
            with bulk_guard(app):
                raise ValueError("boom")

        assert app.ScreenUpdating is True
        assert app.EnableEvents is True
        assert app.Calculation == -4105


# ── _maybe_earlybind ───────────────────────────────────────────────────────────

class TestMaybeEarlybind:
    def test_flag_off_returns_same_app_untouched(self, monkeypatch):
        monkeypatch.delenv("THEPEXCEL_MCP_EARLYBIND", raising=False)
        app = MagicMock()
        assert _maybe_earlybind(app) is app

    def test_flag_on_hwnd_match_adopts_early_bound(self, monkeypatch):
        monkeypatch.setenv("THEPEXCEL_MCP_EARLYBIND", "1")
        app = MagicMock()
        app.Hwnd = 123
        early = MagicMock()
        early.Hwnd = 123
        monkeypatch.setattr(
            session_mod.win32com.client.gencache, "EnsureDispatch",
            MagicMock(return_value=early),
        )
        assert _maybe_earlybind(app) is early

    def test_flag_on_hwnd_mismatch_falls_back_to_late_bound(self, monkeypatch):
        monkeypatch.setenv("THEPEXCEL_MCP_EARLYBIND", "1")
        app = MagicMock()
        app.Hwnd = 123
        early = MagicMock()
        early.Hwnd = 999  # different instance — EnsureDispatch spun up a new one
        monkeypatch.setattr(
            session_mod.win32com.client.gencache, "EnsureDispatch",
            MagicMock(return_value=early),
        )
        assert _maybe_earlybind(app) is app

    def test_flag_on_ensure_dispatch_raises_falls_back(self, monkeypatch):
        monkeypatch.setenv("THEPEXCEL_MCP_EARLYBIND", "1")
        app = MagicMock()
        monkeypatch.setattr(
            session_mod.win32com.client.gencache, "EnsureDispatch",
            MagicMock(side_effect=Exception("boom")),
        )
        assert _maybe_earlybind(app) is app


# ── run_com slow-call diagnostic ──────────────────────────────────────────────

class TestSlowCallLogging:
    def test_slow_call_logs_one_line_to_stderr(self, monkeypatch, capsys):
        monkeypatch.setattr(session_mod, "_SLOW_LOG_THRESHOLD", 0.0)
        worker = session_mod._COMWorker()

        def _fast_fn():
            return "ok"

        result = worker.submit(_fast_fn)

        assert result == "ok"
        captured = capsys.readouterr()
        assert "slow COM call" in captured.err
        assert captured.out == ""  # never write diagnostics to stdout (stdio transport)

    def test_fast_call_under_threshold_logs_nothing(self, capsys):
        worker = session_mod._COMWorker()  # default threshold 5.0s

        def _fast_fn():
            return "ok"

        result = worker.submit(_fast_fn)

        assert result == "ok"
        captured = capsys.readouterr()
        assert "slow COM call" not in captured.err
