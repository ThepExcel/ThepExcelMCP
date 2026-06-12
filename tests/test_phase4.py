"""Phase 4 unit tests — no Excel required.

Covers:
- chart type mapping (_CHART_TYPE completeness)
- chart_action arg validation (unknown action, missing required args)
- write_py escaping logic
- screenshot_action arg validation
"""

from __future__ import annotations

import pytest
from unittest.mock import patch, MagicMock

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


# ── Chart type mapping ─────────────────────────────────────────────────────────

class TestChartTypeMapping:
    def test_column_maps_to_51(self):
        from thepexcel_mcp.domains.charts import _CHART_TYPE
        assert _CHART_TYPE["column"] == 51

    def test_bar_maps_to_57(self):
        from thepexcel_mcp.domains.charts import _CHART_TYPE
        assert _CHART_TYPE["bar"] == 57

    def test_line_maps_to_4(self):
        from thepexcel_mcp.domains.charts import _CHART_TYPE
        assert _CHART_TYPE["line"] == 4

    def test_pie_maps_to_5(self):
        from thepexcel_mcp.domains.charts import _CHART_TYPE
        assert _CHART_TYPE["pie"] == 5

    def test_scatter_maps_to_negative_4169(self):
        from thepexcel_mcp.domains.charts import _CHART_TYPE
        assert _CHART_TYPE["scatter"] == -4169

    def test_doughnut_maps_to_negative_4120(self):
        from thepexcel_mcp.domains.charts import _CHART_TYPE
        assert _CHART_TYPE["doughnut"] == -4120

    def test_all_values_are_ints(self):
        from thepexcel_mcp.domains.charts import _CHART_TYPE
        for k, v in _CHART_TYPE.items():
            assert isinstance(v, int), f"chart type '{k}' value is not int: {v!r}"


# ── Chart action dispatch ──────────────────────────────────────────────────────

class TestChartActionDispatch:
    def _call(self, **kwargs):
        mock_session = make_mock_session()
        mock_session.get_workbook.side_effect = ToolError("no excel")
        with patch("thepexcel_mcp.domains.charts._session", mock_session):
            from thepexcel_mcp.domains.charts import chart_action
            return chart_action(**kwargs)

    def test_unknown_action(self):
        with pytest.raises(ToolError, match="Unknown action"):
            self._call(action="frobnicate")

    def test_create_missing_source(self):
        with pytest.raises(ToolError, match="requires 'source'"):
            self._call(action="create", chart_type="column")

    def test_create_missing_chart_type(self):
        with pytest.raises(ToolError, match="requires 'chart_type'"):
            self._call(action="create", source="A1:C10")

    def test_create_unknown_chart_type(self):
        with pytest.raises(ToolError, match="Unknown chart_type"):
            self._call(action="create", source="A1:C10", chart_type="laser_beam")

    def test_configure_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="configure", title="My Chart")

    def test_set_source_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="set_source", source="A1:D10")

    def test_set_source_missing_source(self):
        with pytest.raises(ToolError, match="requires 'source'"):
            self._call(action="set_source", name="Chart1")

    def test_export_image_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="export_image")

    def test_delete_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="delete")


# ── Screenshot action dispatch ─────────────────────────────────────────────────

class TestScreenshotActionDispatch:
    def _call(self, **kwargs):
        mock_session = make_mock_session()
        mock_session.get_workbook.side_effect = ToolError("no excel")
        with patch("thepexcel_mcp.domains.screenshot._session", mock_session):
            from thepexcel_mcp.domains.screenshot import screenshot_action
            return screenshot_action(**kwargs)

    def test_unknown_action(self):
        with pytest.raises(ToolError, match="Unknown action"):
            self._call(action="frobnicate")

    def test_range_missing_range(self):
        with pytest.raises(ToolError, match="requires 'range'"):
            self._call(action="range")

    def test_sheet_no_error_without_name(self):
        # sheet action does NOT require name (uses active sheet when omitted)
        # dispatch should pass to run_com; ToolError from get_workbook is fine
        with pytest.raises(ToolError):
            self._call(action="sheet")

    def test_chart_missing_name(self):
        with pytest.raises(ToolError, match="requires 'name'"):
            self._call(action="chart")


# ── write_py escaping ──────────────────────────────────────────────────────────

class TestWritePyEscaping:
    """Test the _build_py_formula helper in isolation."""

    def test_simple_code_no_quotes(self):
        from thepexcel_mcp.domains.ranges import _build_py_formula
        result = _build_py_formula("x = 1 + 2")
        assert result == '=PY("x = 1 + 2",0)'

    def test_code_with_double_quotes_escaped(self):
        from thepexcel_mcp.domains.ranges import _build_py_formula
        result = _build_py_formula('s = "hello"')
        # Double-quotes inside code must be doubled for Excel formula string
        assert result == '=PY("s = ""hello""",0)'

    def test_code_with_newline_preserved(self):
        from thepexcel_mcp.domains.ranges import _build_py_formula
        result = _build_py_formula("x = 1\ny = 2")
        # Newlines are passed as-is (Excel's Formula2R1C1 handles them)
        assert "x = 1\ny = 2" in result

    def test_empty_code_raises(self):
        from thepexcel_mcp.domains.ranges import _build_py_formula
        with pytest.raises(ToolError, match="python_code"):
            _build_py_formula("")

    def test_range_action_dispatch_write_py(self):
        mock_session = make_mock_session()
        mock_session.get_workbook.side_effect = ToolError("no excel")
        with patch("thepexcel_mcp.domains.ranges._session", mock_session):
            from thepexcel_mcp.domains.ranges import range_action
            with pytest.raises(ToolError, match="requires 'python_code'"):
                range_action(action="write_py", range="A1")
