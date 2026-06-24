"""Unit tests for the Data Model domain — no Excel required.

Tests cover:
- Format type mapping (_FORMAT_MAP and _get_format_obj fallback logic)
- Argument validation (_require guards in datamodel_action dispatch)
- Unknown action rejection
- _find_model_table / _find_model_column error messages
- _list_measure_names helper
- add_table source_type validation
"""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock, patch

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


# ── Import guard: all COM imports happen inside the module, not at collection ──

def _action(**kwargs):
    """Call datamodel_action with a patched _session that never reaches COM.

    run_com is a transparent passthrough so validation and helper logic runs
    synchronously in the test process (no STA worker, no Excel needed).
    """
    mock_session = make_mock_session()
    mock_session.get_workbook.side_effect = ToolError("no excel")
    with patch("thepexcel_mcp.domains.datamodel._session", mock_session):
        from thepexcel_mcp.domains.datamodel import datamodel_action
        return datamodel_action(**kwargs)


# ── Format map ─────────────────────────────────────────────────────────────────

class TestFormatMap:
    """Verify _FORMAT_MAP keys resolve to correct property names."""

    def _prop(self, key: str) -> str:
        from thepexcel_mcp.domains.datamodel import _FORMAT_MAP
        return _FORMAT_MAP.get(key.lower(), "MISSING")

    def test_general(self):
        assert self._prop("general") == "ModelFormatGeneral"

    def test_decimal_aliases(self):
        assert self._prop("decimal") == "ModelFormatDecimalNumber"
        assert self._prop("number") == "ModelFormatDecimalNumber"

    def test_currency(self):
        assert self._prop("currency") == "ModelFormatCurrency"

    def test_percent_aliases(self):
        assert self._prop("percent") == "ModelFormatPercentageNumber"
        assert self._prop("percentage") == "ModelFormatPercentageNumber"

    def test_whole_aliases(self):
        assert self._prop("whole") == "ModelFormatWholeNumber"
        assert self._prop("integer") == "ModelFormatWholeNumber"

    def test_boolean(self):
        assert self._prop("boolean") == "ModelFormatBoolean"

    def test_date(self):
        assert self._prop("date") == "ModelFormatDate"

    def test_scientific(self):
        assert self._prop("scientific") == "ModelFormatScientificNumber"


class TestGetFormatObj:
    """_get_format_obj returns the right property from a mock model."""

    def _call(self, key):
        from thepexcel_mcp.domains.datamodel import _get_format_obj
        mock_model = MagicMock()
        # Make getattr return a sentinel per property name
        mock_model.ModelFormatGeneral = "fmt_general"
        mock_model.ModelFormatDecimalNumber = "fmt_decimal"
        mock_model.ModelFormatCurrency = "fmt_currency"
        mock_model.ModelFormatPercentageNumber = "fmt_percent"
        mock_model.ModelFormatWholeNumber = "fmt_whole"
        mock_model.ModelFormatBoolean = "fmt_bool"
        mock_model.ModelFormatDate = "fmt_date"
        mock_model.ModelFormatScientificNumber = "fmt_sci"
        return _get_format_obj(mock_model, key)

    def test_none_falls_back_to_general(self):
        assert self._call(None) == "fmt_general"

    def test_unknown_falls_back_to_general(self):
        assert self._call("weird_type") == "fmt_general"

    def test_currency(self):
        assert self._call("currency") == "fmt_currency"

    def test_decimal(self):
        assert self._call("decimal") == "fmt_decimal"

    def test_percent(self):
        assert self._call("percent") == "fmt_percent"

    def test_whole(self):
        assert self._call("whole") == "fmt_whole"


# ── Argument validation in dispatch ────────────────────────────────────────────

class TestActionDispatch:
    """Verify _require guards fire before any COM call reaches the session."""

    def test_unknown_action(self):
        with pytest.raises(ToolError, match="Unknown action"):
            _action(action="nonexistent")

    def test_info_hits_session(self):
        # info has no _require guard, goes straight to _get_model → session
        with pytest.raises(ToolError, match="no excel"):
            _action(action="info")

    def test_list_tables_hits_session(self):
        with pytest.raises(ToolError, match="no excel"):
            _action(action="list_tables")

    def test_add_table_missing_source_type(self):
        with pytest.raises(ToolError, match="requires 'source_type'"):
            _action(action="add_table", source_name="T1")

    def test_add_table_missing_source_name(self):
        with pytest.raises(ToolError, match="requires 'source_name'"):
            _action(action="add_table", source_type="table")

    def test_add_relationship_missing_from_table(self):
        with pytest.raises(ToolError, match="requires 'from_table'"):
            _action(action="add_relationship",
                    from_column="ID", to_table="P", to_column="ID")

    def test_add_relationship_missing_from_column(self):
        with pytest.raises(ToolError, match="requires 'from_column'"):
            _action(action="add_relationship",
                    from_table="S", to_table="P", to_column="ID")

    def test_add_relationship_missing_to_table(self):
        with pytest.raises(ToolError, match="requires 'to_table'"):
            _action(action="add_relationship",
                    from_table="S", from_column="ID", to_column="ID")

    def test_add_relationship_missing_to_column(self):
        with pytest.raises(ToolError, match="requires 'to_column'"):
            _action(action="add_relationship",
                    from_table="S", from_column="ID", to_table="P")

    def test_delete_relationship_missing_index(self):
        with pytest.raises(ToolError, match="requires 'relationship_index'"):
            _action(action="delete_relationship")

    def test_add_measure_missing_name(self):
        with pytest.raises(ToolError, match="requires 'measure_name'"):
            _action(action="add_measure", table="T", formula="=SUM(T[A])")

    def test_add_measure_missing_table(self):
        with pytest.raises(ToolError, match="requires 'table'"):
            _action(action="add_measure", measure_name="M", formula="=SUM(T[A])")

    def test_add_measure_missing_formula(self):
        with pytest.raises(ToolError, match="requires 'formula'"):
            _action(action="add_measure", measure_name="M", table="T")

    def test_update_measure_missing_name(self):
        with pytest.raises(ToolError, match="requires 'measure_name'"):
            _action(action="update_measure")

    def test_delete_measure_missing_name(self):
        with pytest.raises(ToolError, match="requires 'measure_name'"):
            _action(action="delete_measure")


# ── add_table source_type validation ──────────────────────────────────────────

class TestAddTableSourceType:
    """source_type must be 'table' or 'query' (validated after session call)."""

    def _action_with_wb(self, source_type, source_name):
        """Patch session to succeed (return mock wb), let domain logic run."""
        mock_session = make_mock_session()
        mock_wb = MagicMock()
        mock_wb.Sheets.Count = 0
        mock_wb.Queries.Count = 0
        mock_wb.Connections.Count = 0
        mock_session.get_workbook.return_value = mock_wb
        with patch("thepexcel_mcp.domains.datamodel._session", mock_session):
            from thepexcel_mcp.domains.datamodel import datamodel_action
            return datamodel_action(action="add_table",
                                    source_type=source_type,
                                    source_name=source_name)

    def test_invalid_source_type(self):
        with pytest.raises(ToolError, match="source_type.*invalid"):
            self._action_with_wb("range", "SalesTable")

    def test_table_not_found(self):
        with pytest.raises(ToolError, match="not found"):
            self._action_with_wb("table", "NonExistent")

    def test_query_not_found(self):
        with pytest.raises(ToolError, match="not found"):
            self._action_with_wb("query", "NonExistent")


# ── _find_model_table error message ───────────────────────────────────────────

class TestFindModelTable:
    """_find_model_table raises ToolError with available list on miss."""

    def test_table_not_found_shows_available(self):
        from thepexcel_mcp.domains.datamodel import _find_model_table
        mock_model = MagicMock()
        # Simulate model with 1 table named "Sales"
        mock_model.ModelTables.side_effect = Exception("not found")
        mock_model.ModelTables.Count = 1
        mock_item = MagicMock()
        mock_item.Name = "Sales"
        mock_model.ModelTables.Item.return_value = mock_item

        with pytest.raises(ToolError) as exc_info:
            _find_model_table(mock_model, "Products")
        assert "Products" in str(exc_info.value)
        assert "not found" in str(exc_info.value).lower()


# ── _list_measure_names helper ─────────────────────────────────────────────────

class TestListMeasureNames:
    """_list_measure_names returns empty list on exception."""

    def test_empty_on_error(self):
        from thepexcel_mcp.domains.datamodel import _list_measure_names
        mock_model = MagicMock()
        mock_model.ModelMeasures.Count = 0
        assert _list_measure_names(mock_model) == []

    def test_returns_names(self):
        from thepexcel_mcp.domains.datamodel import _list_measure_names
        mock_model = MagicMock()
        mock_model.ModelMeasures.Count = 2
        m1 = MagicMock()
        m1.Name = "Total Sales"
        m2 = MagicMock()
        m2.Name = "Avg Price"
        mock_model.ModelMeasures.Item.side_effect = lambda i: [m1, m2][i - 1]
        assert _list_measure_names(mock_model) == ["Total Sales", "Avg Price"]


# ── _format_info helper ────────────────────────────────────────────────────────

class TestFormatInfo:
    """_format_info extracts a human-readable format string from a measure."""

    def _call(self, cls_name: str) -> str:
        from thepexcel_mcp.domains.datamodel import _format_info
        mock_m = MagicMock()
        mock_fmt = MagicMock()
        mock_fmt.__class__.__name__ = cls_name
        mock_m.FormatInformation = mock_fmt
        return _format_info(mock_m)

    def test_general(self):
        assert self._call("ModelFormatGeneral") == "general"

    def test_decimal(self):
        assert self._call("ModelFormatDecimalNumber") == "decimal"

    def test_currency(self):
        assert self._call("ModelFormatCurrency") == "currency"

    def test_percent(self):
        assert self._call("ModelFormatPercentageNumber") == "percent"

    def test_whole(self):
        assert self._call("ModelFormatWholeNumber") == "whole"

    def test_exception_returns_general(self):
        from thepexcel_mcp.domains.datamodel import _format_info
        mock_m = MagicMock()
        mock_m.FormatInformation = property(lambda self: (_ for _ in ()).throw(Exception("boom")))
        # Should not raise
        result = _format_info(mock_m)
        assert isinstance(result, str)


# ── Cube formula builder (pure Python, no COM) ────────────────────────────────

class TestCubeFormulaBuilders:
    """Exact-match assertions for _build_cubevalue / _build_cubemember.

    All expected strings were verified in scratch/tier3/exp_cube_formula.py
    Layer A (10/10 offline exact-match checks).
    """

    def test_cubevalue_measure_only_friendly(self):
        from thepexcel_mcp.domains.datamodel import _build_cubevalue
        assert _build_cubevalue("Total Sales") == (
            '=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Total Sales]")'
        )

    def test_cubevalue_measure_only_already_mdx(self):
        from thepexcel_mcp.domains.datamodel import _build_cubevalue
        assert _build_cubevalue("[Measures].[Total Sales]") == (
            '=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Total Sales]")'
        )

    def test_cubevalue_measure_plus_one_member(self):
        from thepexcel_mcp.domains.datamodel import _build_cubevalue
        assert _build_cubevalue(
            "Total Sales", ["[Products].[Category].[Bikes]"]
        ) == (
            '=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Total Sales]",'
            '"[Products].[Category].[Bikes]")'
        )

    def test_cubevalue_measure_plus_two_members(self):
        from thepexcel_mcp.domains.datamodel import _build_cubevalue
        assert _build_cubevalue(
            "Total Sales",
            ["[Products].[Category].[Bikes]", "[Date].[CalendarYear].[2024]"],
        ) == (
            '=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Total Sales]",'
            '"[Products].[Category].[Bikes]","[Date].[CalendarYear].[2024]")'
        )

    def test_cubevalue_custom_connection(self):
        from thepexcel_mcp.domains.datamodel import _build_cubevalue
        assert _build_cubevalue("Total Sales", connection="MyOlapCube") == (
            '=CUBEVALUE("MyOlapCube","[Measures].[Total Sales]")'
        )

    def test_cubemember_no_caption(self):
        from thepexcel_mcp.domains.datamodel import _build_cubemember
        assert _build_cubemember("[Products].[Category].[Bikes]") == (
            '=CUBEMEMBER("ThisWorkbookDataModel","[Products].[Category].[Bikes]")'
        )

    def test_cubemember_with_caption(self):
        from thepexcel_mcp.domains.datamodel import _build_cubemember
        assert _build_cubemember(
            "[Products].[Category].[Bikes]", caption="Bikes"
        ) == (
            '=CUBEMEMBER("ThisWorkbookDataModel","[Products].[Category].[Bikes]","Bikes")'
        )

    def test_cubemember_measure_member(self):
        from thepexcel_mcp.domains.datamodel import _build_cubemember
        assert _build_cubemember("[Measures].[Total Sales]") == (
            '=CUBEMEMBER("ThisWorkbookDataModel","[Measures].[Total Sales]")'
        )

    def test_cubevalue_inner_quote_escaped(self):
        from thepexcel_mcp.domains.datamodel import _build_cubevalue
        assert _build_cubevalue('[Measures].[Say "Hi"]') == (
            '=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Say ""Hi""]")'
        )


# ── cube_formula dispatch (pure Python, no COM) ───────────────────────────────

class TestCubeFormulaDispatch:
    """cube_formula action is pure Python — no COM call, no session needed."""

    def _action(self, **kwargs):
        # cube_formula never calls run_com, so session mock is irrelevant.
        mock_session = make_mock_session()
        with patch("thepexcel_mcp.domains.datamodel._session", mock_session):
            from thepexcel_mcp.domains.datamodel import datamodel_action
            return datamodel_action(**kwargs)

    def test_cubevalue_kind_default(self):
        result = self._action(action="cube_formula", measure="Total Sales")
        assert result["formula"] == (
            '=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Total Sales]")'
        )

    def test_cubevalue_kind_explicit(self):
        result = self._action(action="cube_formula", kind="cubevalue", measure="Revenue")
        assert result["formula"].startswith("=CUBEVALUE(")

    def test_cubemember_kind(self):
        result = self._action(
            action="cube_formula",
            kind="cubemember",
            member_expression="[Products].[Category].[Bikes]",
        )
        assert result["formula"] == (
            '=CUBEMEMBER("ThisWorkbookDataModel","[Products].[Category].[Bikes]")'
        )

    def test_cubemember_with_caption(self):
        result = self._action(
            action="cube_formula",
            kind="cubemember",
            member_expression="[Products].[Category].[Bikes]",
            caption="Bikes",
        )
        assert result["formula"] == (
            '=CUBEMEMBER("ThisWorkbookDataModel","[Products].[Category].[Bikes]","Bikes")'
        )

    def test_cubevalue_missing_measure_raises(self):
        with pytest.raises(ToolError, match="requires 'measure'"):
            self._action(action="cube_formula", kind="cubevalue")

    def test_cubemember_missing_member_expression_raises(self):
        with pytest.raises(ToolError, match="requires 'member_expression'"):
            self._action(action="cube_formula", kind="cubemember")

    def test_invalid_kind_raises(self):
        with pytest.raises(ToolError, match="kind.*invalid"):
            self._action(action="cube_formula", kind="badkind", measure="M")


# ── cube_value / cube_member COM write (mocked) ───────────────────────────────

class TestCubeValueCOMWrite:
    """cube_value: asserts cell.Formula is set to exact built string.

    The COM layer is fully mocked — no Excel needed. Error-marker read-back
    triggers ToolError. Normal read-back returns the result dict.

    CAVEAT (from dispatcher docstring): live numeric resolution requires an
    existing in-workbook Data Model. This is not unit-testable headless.
    """

    def _make_session_and_cell(self, text_return: str):
        """Return (mock_session, mock_cell) pair with cell.Text = text_return."""
        mock_session = make_mock_session()
        mock_app = MagicMock()
        mock_ws = MagicMock()
        mock_cell = MagicMock()
        mock_cell.Value = 1350
        mock_cell.Text = text_return
        mock_cell.Formula = None  # will be set by the domain helper
        mock_ws.Range.return_value = mock_cell
        mock_session.get_app.return_value = mock_app
        mock_session.get_sheet.return_value = mock_ws
        return mock_session, mock_cell

    def _action(self, mock_session, **kwargs):
        with patch("thepexcel_mcp.domains.datamodel._session", mock_session):
            from thepexcel_mcp.domains.datamodel import datamodel_action
            return datamodel_action(**kwargs)

    def test_cell_formula_set_to_built_string(self):
        mock_session, mock_cell = self._make_session_and_cell("1,350")
        result = self._action(
            mock_session,
            action="cube_value",
            target_cell="B2",
            measure="Total Sales",
        )
        # VERIFY-EFFECT: cell.Formula must be the exact CUBEVALUE string
        expected = '=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Total Sales]")'
        assert mock_cell.Formula == expected
        assert result["formula"] == expected
        assert result["cell"] == "B2"

    def test_cell_formula_with_members(self):
        mock_session, mock_cell = self._make_session_and_cell("250")
        self._action(
            mock_session,
            action="cube_value",
            target_cell="C3",
            measure="Total Sales",
            members=["[Products].[Category].[Bikes]"],
        )
        expected = (
            '=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Total Sales]",'
            '"[Products].[Category].[Bikes]")'
        )
        assert mock_cell.Formula == expected

    def test_error_marker_name_raises(self):
        mock_session, mock_cell = self._make_session_and_cell("#NAME?")
        with pytest.raises(ToolError, match="#NAME"):
            self._action(
                mock_session,
                action="cube_value",
                target_cell="B2",
                measure="NoSuchMeasure",
            )

    def test_error_marker_getting_data_raises(self):
        mock_session, mock_cell = self._make_session_and_cell("#GETTING_DATA")
        with pytest.raises(ToolError, match="GETTING_DATA"):
            self._action(
                mock_session,
                action="cube_value",
                target_cell="B2",
                measure="Total Sales",
            )

    def test_missing_target_cell_raises(self):
        mock_session, _ = self._make_session_and_cell("1,350")
        with pytest.raises(ToolError, match="requires 'target_cell'"):
            self._action(mock_session, action="cube_value", measure="Total Sales")

    def test_missing_measure_raises(self):
        mock_session, _ = self._make_session_and_cell("1,350")
        with pytest.raises(ToolError, match="requires 'measure'"):
            self._action(mock_session, action="cube_value", target_cell="B2")


class TestCubeMemberCOMWrite:
    """cube_member: asserts cell.Formula is set to exact built CUBEMEMBER string."""

    def _make_session_and_cell(self, text_return: str):
        mock_session = make_mock_session()
        mock_app = MagicMock()
        mock_ws = MagicMock()
        mock_cell = MagicMock()
        mock_cell.Value = "Bikes"
        mock_cell.Text = text_return
        mock_cell.Formula = None
        mock_ws.Range.return_value = mock_cell
        mock_session.get_app.return_value = mock_app
        mock_session.get_sheet.return_value = mock_ws
        return mock_session, mock_cell

    def _action(self, mock_session, **kwargs):
        with patch("thepexcel_mcp.domains.datamodel._session", mock_session):
            from thepexcel_mcp.domains.datamodel import datamodel_action
            return datamodel_action(**kwargs)

    def test_cell_formula_set_to_built_string(self):
        mock_session, mock_cell = self._make_session_and_cell("Bikes")
        result = self._action(
            mock_session,
            action="cube_member",
            target_cell="A1",
            member_expression="[Products].[Category].[Bikes]",
        )
        expected = (
            '=CUBEMEMBER("ThisWorkbookDataModel","[Products].[Category].[Bikes]")'
        )
        assert mock_cell.Formula == expected
        assert result["formula"] == expected
        assert result["cell"] == "A1"

    def test_cell_formula_with_caption(self):
        mock_session, mock_cell = self._make_session_and_cell("Bikes")
        self._action(
            mock_session,
            action="cube_member",
            target_cell="A1",
            member_expression="[Products].[Category].[Bikes]",
            caption="Bikes",
        )
        expected = (
            '=CUBEMEMBER("ThisWorkbookDataModel","[Products].[Category].[Bikes]","Bikes")'
        )
        assert mock_cell.Formula == expected

    def test_error_marker_raises(self):
        mock_session, mock_cell = self._make_session_and_cell("#NAME?")
        with pytest.raises(ToolError, match="#NAME"):
            self._action(
                mock_session,
                action="cube_member",
                target_cell="A1",
                member_expression="[Bad].[Member].[X]",
            )

    def test_missing_target_cell_raises(self):
        mock_session, _ = self._make_session_and_cell("Bikes")
        with pytest.raises(ToolError, match="requires 'target_cell'"):
            self._action(
                mock_session,
                action="cube_member",
                member_expression="[Products].[Category].[Bikes]",
            )

    def test_missing_member_expression_raises(self):
        mock_session, _ = self._make_session_and_cell("Bikes")
        with pytest.raises(ToolError, match="requires 'member_expression'"):
            self._action(mock_session, action="cube_member", target_cell="A1")


# ── Calculated column / table guard ──────────────────────────────────────────

class TestCalcColumnGuard:
    """add_calculated_column / add_calculated_table raise informative ToolErrors.

    Pure Python — no COM call, no Excel. Verified against constants from
    scratch/tier3/exp_calc_column.py (design_confirmation 5/5 checks).
    """

    def _action(self, **kwargs):
        mock_session = make_mock_session()
        with patch("thepexcel_mcp.domains.datamodel._session", mock_session):
            from thepexcel_mcp.domains.datamodel import datamodel_action
            return datamodel_action(**kwargs)

    def test_add_calculated_column_raises(self):
        with pytest.raises(ToolError) as exc_info:
            self._action(action="add_calculated_column")
        msg = str(exc_info.value)
        assert "READ-ONLY" in msg
        assert "Power Query" in msg
        assert "add_measure" in msg

    def test_add_calculated_table_raises(self):
        with pytest.raises(ToolError) as exc_info:
            self._action(action="add_calculated_table")
        msg = str(exc_info.value)
        assert "Power BI" in msg
        assert "Analysis Services" in msg

    def test_unknown_action_lists_new_names(self):
        with pytest.raises(ToolError) as exc_info:
            self._action(action="nonexistent_xyz")
        msg = str(exc_info.value)
        assert "cube_value" in msg
        assert "cube_member" in msg
        assert "cube_formula" in msg
        assert "add_calculated_column" in msg
        assert "add_calculated_table" in msg
