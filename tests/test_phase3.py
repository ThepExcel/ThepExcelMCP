"""Phase 3 unit tests — no Excel required.

Covers:
- STA COM worker thread (queue/Future/timeout machinery)
- excel_vba: env gate, AccessVBOM pre-flight, action dispatch, module helpers
- excel_name: action dispatch, _is_lambda heuristic, _name_info scope detection
- excel_range read_spill: action dispatch
"""

from __future__ import annotations

import os
import time

import pytest
from unittest.mock import MagicMock, patch, PropertyMock

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


# ── STA COM Worker ─────────────────────────────────────────────────────────────

class TestCOMWorker:
    """Queue/Future machinery — no real COM calls needed."""

    def test_submit_returns_result(self):
        from thepexcel_mcp.session import _COMWorker
        w = _COMWorker()
        result = w.submit(lambda: 42)
        assert result == 42

    def test_submit_propagates_exception(self):
        from thepexcel_mcp.session import _COMWorker
        w = _COMWorker()
        with pytest.raises(ValueError, match="boom"):
            w.submit(lambda: (_ for _ in ()).throw(ValueError("boom")))

    def test_submit_passes_args(self):
        from thepexcel_mcp.session import _COMWorker
        w = _COMWorker()
        result = w.submit(lambda a, b: a + b, 3, 4)
        assert result == 7

    def test_submit_with_kwargs(self):
        from thepexcel_mcp.session import _COMWorker
        w = _COMWorker()
        result = w.submit(lambda x=0: x * 2, x=5)
        assert result == 10

    def test_timeout_raises_tool_error(self, monkeypatch):
        """When Future.result() raises TimeoutError, run_com raises ToolError."""
        from thepexcel_mcp.session import _COMWorker
        from concurrent.futures import Future
        w = _COMWorker()
        # Monkey-patch submit to simulate a Future that times out
        fake_future = MagicMock(spec=Future)
        fake_future.result.side_effect = TimeoutError()
        original_put = w._queue.put

        def fake_put(item):
            # Don't actually put it; just set the future directly
            future, fn, args, kwargs = item
            future.result = lambda timeout=None: (_ for _ in ()).throw(TimeoutError())

        monkeypatch.setattr(w._queue, "put", fake_put)
        monkeypatch.setattr(w, "_ensure_started", lambda: None)

        with pytest.raises(ToolError, match="timed out"):
            w.submit(lambda: time.sleep(999))

    def test_run_com_on_session(self):
        """ExcelSession.run_com delegates to the worker."""
        from thepexcel_mcp.session import ExcelSession
        sess = ExcelSession()
        result = sess.run_com(lambda: "hello")
        assert result == "hello"

    def test_run_com_re_raises_tool_error(self):
        from thepexcel_mcp.session import ExcelSession
        sess = ExcelSession()
        with pytest.raises(ToolError, match="test error"):
            sess.run_com(lambda: (_ for _ in ()).throw(ToolError("test error")))


class TestExcelGuard:
    """excel_guard context manager sets/restores DisplayAlerts."""

    def test_sets_false_then_restores(self):
        from thepexcel_mcp.session import excel_guard
        mock_app = MagicMock()
        with excel_guard(mock_app):
            assert mock_app.DisplayAlerts is False
        assert mock_app.DisplayAlerts is True

    def test_restores_on_exception(self):
        from thepexcel_mcp.session import excel_guard
        mock_app = MagicMock()
        try:
            with excel_guard(mock_app):
                raise RuntimeError("boom")
        except RuntimeError:
            pass
        assert mock_app.DisplayAlerts is True


class TestWaitCalculation:
    """wait_calculation raises ToolError on timeout."""

    def test_immediate_idle_returns(self):
        from thepexcel_mcp.session import wait_calculation
        mock_app = MagicMock()
        mock_app.CalculationState = 0  # idle immediately
        wait_calculation(mock_app, timeout=1.0)  # should not raise

    def test_timeout_raises(self):
        from thepexcel_mcp.session import wait_calculation
        mock_app = MagicMock()
        mock_app.CalculationState = 1  # always calculating
        with pytest.raises(ToolError, match="still calculating"):
            wait_calculation(mock_app, timeout=0.1)


# ── VBA domain ─────────────────────────────────────────────────────────────────

class TestVBAEnvGate:
    """Every action raises ToolError when THEPEXCEL_MCP_ENABLE_VBA is unset."""

    def _call(self, **kwargs):
        from thepexcel_mcp.domains.vba import vba_action
        return vba_action(**kwargs)

    def test_list_modules_blocked(self):
        with patch.dict(os.environ, {}, clear=False):
            os.environ.pop("THEPEXCEL_MCP_ENABLE_VBA", None)
            with pytest.raises(ToolError, match="VBA tool disabled"):
                self._call(action="list_modules")

    def test_run_blocked(self):
        with patch.dict(os.environ, {}, clear=False):
            os.environ.pop("THEPEXCEL_MCP_ENABLE_VBA", None)
            with pytest.raises(ToolError, match="VBA tool disabled"):
                self._call(action="run", proc_name="Macro1")

    def test_unknown_action_still_blocked_by_env_gate(self):
        with patch.dict(os.environ, {}, clear=False):
            os.environ.pop("THEPEXCEL_MCP_ENABLE_VBA", None)
            with pytest.raises(ToolError, match="VBA tool disabled"):
                self._call(action="frobnicate")


class TestVBAAccessVBOM:
    """AccessVBOM pre-flight raises when registry value != 1."""

    def _make_winreg_mock(self, value):
        """Build a mock winreg module that returns the given AccessVBOM value."""
        mock_wr = MagicMock()

        class _FakeKey:
            def __enter__(self): return self
            def __exit__(self, *_): pass

        mock_wr.HKEY_CURRENT_USER = 0x80000001
        mock_wr.OpenKey.return_value = _FakeKey()
        mock_wr.QueryValueEx.return_value = (value, 1)
        return mock_wr

    def test_access_vbom_0_raises(self):
        mock_wr = self._make_winreg_mock(0)
        with patch.dict(os.environ, {"THEPEXCEL_MCP_ENABLE_VBA": "1"}):
            with patch("thepexcel_mcp.domains.vba.winreg", mock_wr):
                with pytest.raises(ToolError, match="Trust access"):
                    from thepexcel_mcp.domains.vba import vba_action
                    vba_action(action="list_modules")

    def test_access_vbom_missing_raises(self):
        """FileNotFoundError from OpenKey → pre-flight raises."""
        mock_wr = MagicMock()
        mock_wr.HKEY_CURRENT_USER = 0x80000001
        mock_wr.OpenKey.side_effect = FileNotFoundError
        with patch.dict(os.environ, {"THEPEXCEL_MCP_ENABLE_VBA": "1"}):
            with patch("thepexcel_mcp.domains.vba.winreg", mock_wr):
                with pytest.raises(ToolError, match="Trust access"):
                    from thepexcel_mcp.domains.vba import vba_action
                    vba_action(action="list_modules")


class TestVBAActionDispatch:
    """Argument validation fires before run_com (env gate + vbom bypassed via mock)."""

    def _call(self, **kwargs):
        from thepexcel_mcp.domains.vba import vba_action
        mock_session = make_mock_session()
        mock_session.get_workbook.side_effect = ToolError("no excel")
        with patch.dict(os.environ, {"THEPEXCEL_MCP_ENABLE_VBA": "1"}):
            with patch("thepexcel_mcp.domains.vba._access_vbom_preflight"):
                with patch("thepexcel_mcp.domains.vba._session", mock_session):
                    return vba_action(**kwargs)

    def test_unknown_action(self):
        with pytest.raises(ToolError, match="Unknown action"):
            self._call(action="frobnicate")

    def test_get_module_missing_name(self):
        with pytest.raises(ToolError, match="requires 'module_name'"):
            self._call(action="get_module")

    def test_write_module_missing_name(self):
        with pytest.raises(ToolError, match="requires 'module_name'"):
            self._call(action="write_module", code="Sub X() End Sub")

    def test_write_module_missing_code(self):
        with pytest.raises(ToolError, match="requires 'code'"):
            self._call(action="write_module", module_name="M1")

    def test_delete_module_missing_name(self):
        with pytest.raises(ToolError, match="requires 'module_name'"):
            self._call(action="delete_module")

    def test_run_missing_proc_name(self):
        with pytest.raises(ToolError, match="requires 'proc_name'"):
            self._call(action="run")

    def test_list_modules_hits_session(self):
        with pytest.raises(ToolError, match="no excel"):
            self._call(action="list_modules")


class TestVBAHelpers:
    """_is_lambda_module and _find_component helpers."""

    def test_is_lambda_detection(self):
        from thepexcel_mcp.domains.vba import _MODULE_TYPE
        # Type 1 = standard
        assert _MODULE_TYPE[1] == "standard"
        assert _MODULE_TYPE[2] == "class"
        assert _MODULE_TYPE[3] == "form"
        assert _MODULE_TYPE[100] == "document"

    def test_find_component_raises_with_list(self):
        from thepexcel_mcp.domains.vba import _find_component
        mock_vbp = MagicMock()
        mock_vbp.VBComponents.Count = 1
        mock_c = MagicMock()
        mock_c.Name = "Module1"
        mock_vbp.VBComponents.Item.return_value = mock_c
        with pytest.raises(ToolError, match="Module1"):
            _find_component(mock_vbp, "NonExistent")

    def test_find_component_found(self):
        from thepexcel_mcp.domains.vba import _find_component
        mock_vbp = MagicMock()
        mock_vbp.VBComponents.Count = 1
        mock_c = MagicMock()
        mock_c.Name = "Module1"
        mock_vbp.VBComponents.Item.return_value = mock_c
        result = _find_component(mock_vbp, "module1")  # case-insensitive
        assert result is mock_c


# ── Names domain ───────────────────────────────────────────────────────────────

class TestNameActionDispatch:
    """Argument validation in name_action before COM calls."""

    def _call(self, **kwargs):
        mock_session = make_mock_session()
        mock_session.get_workbook.side_effect = ToolError("no excel")
        with patch("thepexcel_mcp.domains.names._session", mock_session):
            from thepexcel_mcp.domains.names import name_action
            return name_action(**kwargs)

    def test_unknown_action(self):
        with pytest.raises(ToolError, match="Unknown action"):
            self._call(action="frobnicate")

    def test_get_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="get")

    def test_set_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="set", refers_to="=A1")

    def test_set_missing_refers_to(self):
        with pytest.raises(ToolError, match="requires 'refers_to'"):
            self._call(action="set", name="X")

    def test_set_refers_to_no_equals(self):
        with pytest.raises(ToolError, match="must start with '='"):
            self._call(action="set", name="X", refers_to="A1")

    def test_delete_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="delete")

    def test_list_hits_session(self):
        with pytest.raises(ToolError, match="no excel"):
            self._call(action="list")

    def test_get_hits_session(self):
        with pytest.raises(ToolError, match="no excel"):
            self._call(action="get", name="MyRange")


class TestIsLambda:
    """_is_lambda heuristic correctly identifies LAMBDA formulas."""

    def setup_method(self):
        from thepexcel_mcp.domains.names import _is_lambda
        self._fn = _is_lambda

    def test_lambda_formula(self):
        assert self._fn("=LAMBDA(x, x*2)") is True

    def test_lambda_multiline(self):
        assert self._fn("=LAMBDA(x, y,\n  x + y)") is True

    def test_lambda_case_insensitive(self):
        assert self._fn("=lambda(x, x)") is True

    def test_named_range(self):
        assert self._fn("=Sheet1!$A$1:$B$10") is False

    def test_constant(self):
        assert self._fn("=42") is False

    def test_empty_string(self):
        assert self._fn("") is False


class TestNameInfo:
    """_name_info extracts correct fields from a mock Name COM object."""

    def test_workbook_scope(self):
        from thepexcel_mcp.domains.names import _name_info
        mock_n = MagicMock()
        mock_n.Name = "TotalSales"
        mock_n.RefersTo = "=Sheet1!$D$10"
        # Parent has no Type attr → treated as workbook scope
        mock_parent = MagicMock(spec=[])  # no attributes, spec=[] means hasattr returns False
        mock_n.Parent = mock_parent
        info = _name_info(mock_n)
        assert info["name"] == "TotalSales"
        assert info["refers_to"] == "=Sheet1!$D$10"
        assert info["is_lambda"] is False

    def test_lambda_detected(self):
        from thepexcel_mcp.domains.names import _name_info
        mock_n = MagicMock()
        mock_n.Name = "DOUBLE"
        mock_n.RefersTo = "=LAMBDA(x, x*2)"
        mock_n.Parent = MagicMock(spec=[])
        info = _name_info(mock_n)
        assert info["is_lambda"] is True

    def test_sheet_scope(self):
        from thepexcel_mcp.domains.names import _name_info
        mock_n = MagicMock()
        mock_n.Name = "Sheet1!LocalName"
        mock_n.RefersTo = "=Sheet1!$A$1"
        mock_parent = MagicMock()
        mock_parent.Type = 1  # Worksheet type exists
        mock_parent.Name = "Sheet1"
        mock_n.Parent = mock_parent
        info = _name_info(mock_n)
        assert info["name"] == "LocalName"
        assert info["scope"] == "Sheet1"


class TestFindName:
    """_find_name raises ToolError with available list on miss."""

    def test_not_found_shows_available(self):
        from thepexcel_mcp.domains.names import _find_name
        mock_wb = MagicMock()
        mock_wb.Names.Count = 1
        mock_n = MagicMock()
        mock_n.Name = "ExistingName"
        mock_wb.Names.Item.return_value = mock_n
        with pytest.raises(ToolError, match="ExistingName"):
            _find_name(mock_wb, "Missing")

    def test_case_insensitive_match(self):
        from thepexcel_mcp.domains.names import _find_name
        mock_wb = MagicMock()
        mock_wb.Names.Count = 1
        mock_n = MagicMock()
        mock_n.Name = "SalesTotal"
        mock_wb.Names.Item.return_value = mock_n
        result = _find_name(mock_wb, "salestotal")
        assert result is mock_n


# ── Range read_spill ───────────────────────────────────────────────────────────

class TestRangeReadSpill:
    """read_spill action validation and dispatch."""

    def _call(self, **kwargs):
        mock_session = make_mock_session()
        mock_session.get_workbook.side_effect = ToolError("no excel")
        mock_session.get_sheet.side_effect = ToolError("no excel")
        with patch("thepexcel_mcp.domains.ranges._session", mock_session):
            from thepexcel_mcp.domains.ranges import range_action
            return range_action(**kwargs)

    def test_unknown_action_includes_read_spill(self):
        with pytest.raises(ToolError, match="read_spill"):
            self._call(action="frobnicate", range="A1")

    def test_read_spill_hits_session(self):
        with pytest.raises(ToolError, match="no excel"):
            self._call(action="read_spill", range="A1")

    def test_read_spill_valid_action(self):
        """read_spill action dispatches without pre-validation error."""
        # It should reach the COM call (which raises "no excel" from mock)
        with pytest.raises(ToolError):
            self._call(action="read_spill", range="E1")

    def test_write_still_requires_values(self):
        with pytest.raises(ToolError, match="requires 'values'"):
            self._call(action="write", range="A1")


class TestReadSpillHelper:
    """_read_spill logic with a mock COM range."""

    def _make_range_mock(self, has_spill: bool, spill_values=None):
        """Build a mock COM range that simulates a spill cell."""
        mock_rng = MagicMock()
        anchor = MagicMock()
        anchor.HasSpill = has_spill
        if has_spill:
            spill_rng = MagicMock()
            spill_rng.Address = "$E$1:$E$5"
            spill_rng.Value = tuple((v,) for v in (spill_values or [1, 2, 3]))
            anchor.SpillingRange = spill_rng
            # _read_spill now uses _spill_range_address → ws.Range(addr)
            # so anchor.Parent.Range() must return the same spill_rng
            anchor.Parent.Range.return_value = spill_rng
        mock_rng.Cells.return_value = anchor
        return mock_rng

    def test_no_spill_raises(self):
        from thepexcel_mcp.domains.ranges import _read_spill
        mock_session = make_mock_session()
        rng_mock = self._make_range_mock(has_spill=False)
        with patch("thepexcel_mcp.domains.ranges._resolve_range", return_value=rng_mock):
            with patch("thepexcel_mcp.domains.ranges._session", mock_session):
                with pytest.raises(ToolError, match="no spill"):
                    _read_spill("E1", None, None, 0, 100)

    def test_spill_returns_values(self):
        from thepexcel_mcp.domains.ranges import _read_spill
        mock_session = make_mock_session()
        rng_mock = self._make_range_mock(has_spill=True, spill_values=[10, 20, 30])
        # Anchor address
        rng_mock.Cells.return_value.Address = "$E$1"
        with patch("thepexcel_mcp.domains.ranges._resolve_range", return_value=rng_mock):
            with patch("thepexcel_mcp.domains.ranges._session", mock_session):
                result = _read_spill("E1", None, None, 0, 100)
        assert result["spill_range"] == "$E$1:$E$5"
        assert result["total_rows"] == 3
        assert result["has_more"] is False
