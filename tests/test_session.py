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
from thepexcel_mcp.session import (
    ExcelSession,
    _clear_gen_py_cache,
    _force_latebound,
    bulk_guard,
    _maybe_earlybind,
    _rewrap_earlybound,
)


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
        monkeypatch.setattr(session_mod, "_maybe_earlybind", lambda app: app)

        result = ExcelSession().get_app()

        assert result is fresh
        assert session_mod._cached_app is fresh

    def test_fresh_attach_is_cached_for_next_call(self, monkeypatch):
        fresh = MagicMock()
        fresh.Name = "Excel"
        monkeypatch.setattr(session_mod.win32com.client, "GetActiveObject", MagicMock(return_value=fresh))
        monkeypatch.setattr(session_mod, "_maybe_earlybind", lambda app: app)

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

    def test_corrupt_gen_py_attach_recovers_in_same_process(self, monkeypatch):
        fresh = MagicMock()
        fresh.Name = "Excel"
        get_active = MagicMock(
            side_effect=[AttributeError("missing CLSIDToPackageMap"), fresh]
        )
        clear_cache = MagicMock()
        monkeypatch.setattr(session_mod.win32com.client, "GetActiveObject", get_active)
        monkeypatch.setattr(session_mod, "_clear_gen_py_cache", clear_cache)
        monkeypatch.setattr(session_mod, "_maybe_earlybind", lambda app: app)

        result = ExcelSession().get_app()

        assert result is fresh
        assert get_active.call_count == 2
        clear_cache.assert_called_once()


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


# ── _rewrap_earlybound / _maybe_earlybind ──────────────────────────────────────

def _wire_typelib(app: MagicMock, la=("GUID", 0, 1, 1, 9, 8)) -> MagicMock:
    """Wire app._oleobj_.GetTypeInfo()...GetLibAttr() to return *la*."""
    tlb = MagicMock()
    tlb.GetLibAttr.return_value = la
    ti = MagicMock()
    ti.GetContainingTypeLib.return_value = (tlb, 0)
    app._oleobj_.GetTypeInfo.return_value = ti
    return tlb


class TestRewrapEarlybound:
    def test_success_rewraps_same_pointer_via_gencache(self, monkeypatch):
        app = MagicMock()
        _wire_typelib(app, la=("GUID", 0, 1, 1, 9, 8))
        early = MagicMock()
        ensure_module = MagicMock()
        dispatch = MagicMock(return_value=early)
        monkeypatch.setattr(session_mod.win32com.client.gencache, "EnsureModule", ensure_module)
        monkeypatch.setattr(session_mod.win32com.client, "Dispatch", dispatch)

        result = _rewrap_earlybound(app)

        assert result is early
        # guid, lcid, major, minor == la[0], la[1], la[3], la[4]
        ensure_module.assert_called_once_with("GUID", 0, 1, 9)
        dispatch.assert_called_once_with(app._oleobj_)

    def test_typelib_lookup_failure_raises(self):
        app = MagicMock()
        app._oleobj_.GetTypeInfo.side_effect = Exception("no typeinfo")
        with pytest.raises(Exception, match="no typeinfo"):
            _rewrap_earlybound(app)

    def test_gencache_failure_raises(self, monkeypatch):
        app = MagicMock()
        _wire_typelib(app)
        monkeypatch.setattr(
            session_mod.win32com.client.gencache, "EnsureModule",
            MagicMock(side_effect=Exception("boom")),
        )
        with pytest.raises(Exception, match="boom"):
            _rewrap_earlybound(app)


class TestMaybeEarlybind:
    def test_kill_switch_constructs_dynamic_wrapper(self, monkeypatch):
        """Skipping the rewrap is insufficient when GetActiveObject already
        returned a generated wrapper; the kill-switch must force dynamic."""
        monkeypatch.setenv("THEPEXCEL_MCP_EARLYBIND", "0")
        app = MagicMock()
        late = MagicMock()
        dispatch = MagicMock(return_value=late)
        monkeypatch.setattr(session_mod.win32com.client.dynamic, "Dispatch", dispatch)

        assert _maybe_earlybind(app) is late
        dispatch.assert_called_once_with(app._oleobj_)

    def test_force_latebound_keeps_existing_dynamic_wrapper(self):
        class DynamicApp:
            pass

        DynamicApp.__module__ = "win32com.client.dynamic"
        app = DynamicApp()
        assert _force_latebound(app) is app

    def test_default_unset_attempts_early_bind(self, monkeypatch):
        """Default (env var unset) is ON as of 2026-07-24 — a successful
        rewrap is adopted with no env var set at all."""
        monkeypatch.delenv("THEPEXCEL_MCP_EARLYBIND", raising=False)
        app = MagicMock()
        _wire_typelib(app)
        early = MagicMock()
        monkeypatch.setattr(session_mod.win32com.client.gencache, "EnsureModule", MagicMock())
        monkeypatch.setattr(session_mod.win32com.client, "Dispatch", MagicMock(return_value=early))

        assert _maybe_earlybind(app) is early

    def test_flag_on_success_adopts_early_bound(self, monkeypatch):
        monkeypatch.setenv("THEPEXCEL_MCP_EARLYBIND", "1")
        app = MagicMock()
        _wire_typelib(app)
        early = MagicMock()
        monkeypatch.setattr(session_mod.win32com.client.gencache, "EnsureModule", MagicMock())
        monkeypatch.setattr(session_mod.win32com.client, "Dispatch", MagicMock(return_value=early))

        assert _maybe_earlybind(app) is early

    def test_flag_on_rewrap_failure_keeps_the_working_handle(self, monkeypatch):
        """On rewrap failure, return the handle that already works — do NOT
        construct a fresh dynamic wrapper. The 2026-08-04 live-COM incident:
        the replacement wrapper's .Range/.Name access failed against a stale
        makepy cache, turning a recoverable rewrap failure into a hard break.
        """
        monkeypatch.setenv("THEPEXCEL_MCP_EARLYBIND", "1")
        app = MagicMock()
        app._oleobj_.GetTypeInfo.side_effect = Exception("boom")
        dynamic_dispatch = MagicMock(return_value=MagicMock())
        monkeypatch.setattr(
            session_mod.win32com.client.dynamic, "Dispatch", dynamic_dispatch
        )

        assert _maybe_earlybind(app) is app
        dynamic_dispatch.assert_not_called()

    def test_flag_on_never_raises_out(self, monkeypatch):
        """Any failure inside the rewrap must be swallowed — _maybe_earlybind
        is called on the get_app() hot path and must never itself raise."""
        monkeypatch.setenv("THEPEXCEL_MCP_EARLYBIND", "1")
        app = MagicMock()
        _wire_typelib(app)
        monkeypatch.setattr(
            session_mod.win32com.client.gencache, "EnsureModule",
            MagicMock(side_effect=Exception("gencache exploded")),
        )
        monkeypatch.setattr(
            session_mod.win32com.client.dynamic,
            "Dispatch",
            MagicMock(side_effect=Exception("dynamic exploded")),
        )
        assert _maybe_earlybind(app) is app


def test_clear_gen_py_cache_evicts_generated_modules_without_machine_wide_rebuild(
    monkeypatch,
):
    """Eviction is the cure; gencache.Rebuild() must NOT be called.

    Rebuild() regenerates every registered typelib on the machine — a
    machine-wide side effect on a recovery path, and the suspected trigger of
    the 2026-08-04 live-COM incident. Dispatch regenerates the one typelib
    this process needs lazily.
    """
    fake_module = "win32com.gen_py.fake_excel_typelib"
    monkeypatch.setitem(session_mod.sys.modules, fake_module, object())
    remove_tree = MagicMock()
    rebuild = MagicMock()
    invalidate = MagicMock()
    monkeypatch.setattr(session_mod.shutil, "rmtree", remove_tree)
    monkeypatch.setattr(session_mod.win32com.client.gencache, "Rebuild", rebuild)
    monkeypatch.setattr(session_mod.importlib, "invalidate_caches", invalidate)

    _clear_gen_py_cache()

    remove_tree.assert_called_once_with(session_mod.win32com.__gen_path__, ignore_errors=True)
    assert fake_module not in session_mod.sys.modules
    invalidate.assert_called_once()
    rebuild.assert_not_called()


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
