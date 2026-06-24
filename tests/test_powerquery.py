"""Unit tests for Power Query parameter management (powerquery.py Tier-3 additions).

No Excel required — pure-Python helpers tested directly; COM-touching branches
are intercepted via make_mock_session (run_com is a transparent passthrough).

Test coverage:
  - Pure-Python helpers: _build_parameter_m, _m_escape_text, _is_parameter,
    _is_function, _parse_parameter, _set_parameter_formula
  - Dispatch: create_parameter / get_parameter / set_parameter / list_parameters
    — missing-arg guards, result shape, and unknown-action error list
  - _list enhancement: is_parameter and is_function flags appear in entries
"""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock, patch

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


# ── Pure-Python helper tests ──────────────────────────────────────────────────

class TestMEscapeText:
    def test_no_quotes(self):
        from thepexcel_mcp.domains.powerquery import _m_escape_text
        assert _m_escape_text("Bangkok") == "Bangkok"

    def test_single_embedded_quote(self):
        from thepexcel_mcp.domains.powerquery import _m_escape_text
        assert _m_escape_text('A "B" C') == 'A ""B"" C'

    def test_multiple_quotes(self):
        from thepexcel_mcp.domains.powerquery import _m_escape_text
        assert _m_escape_text('"x"') == '""x""'


class TestBuildParameterM:
    def test_number_template(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        m = _build_parameter_m(0.07, "Number")
        assert "0.07 meta" in m
        assert "IsParameterQuery=true" in m
        assert 'Type="Number"' in m
        assert not m.startswith('"')

    def test_text_template(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        m = _build_parameter_m("BKK", "Text")
        assert '"BKK" meta' in m
        assert 'Type="Text"' in m

    def test_text_quote_escaping(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        m = _build_parameter_m('A "B" C', "Text")
        assert m.startswith('"A ""B"" C" meta')

    def test_type_inferred_from_int(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        m = _build_parameter_m(10)
        assert 'Type="Number"' in m
        assert "10 meta" in m

    def test_type_inferred_from_str(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        m = _build_parameter_m("hello")
        assert 'Type="Text"' in m

    def test_required_true(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        m = _build_parameter_m(1, "Number", required=True)
        assert "IsParameterQueryRequired=true" in m

    def test_required_false(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        m = _build_parameter_m(1, "Number", required=False)
        assert "IsParameterQueryRequired=false" in m


class TestIsParameter:
    def test_number_param_detected(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m, _is_parameter
        assert _is_parameter(_build_parameter_m(0.07, "Number"))

    def test_text_param_detected(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m, _is_parameter
        assert _is_parameter(_build_parameter_m("BKK", "Text"))

    def test_plain_query_not_detected(self):
        from thepexcel_mcp.domains.powerquery import _is_parameter
        assert not _is_parameter("let Source = 1 in Source")

    def test_none_safe(self):
        from thepexcel_mcp.domains.powerquery import _is_parameter
        assert not _is_parameter("")


class TestIsFunction:
    def test_leading_sig_detected(self):
        from thepexcel_mcp.domains.powerquery import _is_function
        assert _is_function("(x as number) => x * 2")

    def test_param_query_not_function(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m, _is_function
        assert not _is_function(_build_parameter_m(0.07, "Number"))

    def test_let_each_not_function(self):
        # "each" inside a let is NOT a top-level function signature
        from thepexcel_mcp.domains.powerquery import _is_function
        assert not _is_function("let f = each _ * 2 in f")

    def test_empty_safe(self):
        from thepexcel_mcp.domains.powerquery import _is_function
        assert not _is_function("")


class TestParseParameter:
    def test_number_roundtrip(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m, _parse_parameter
        m = _build_parameter_m(0.07, "Number")
        p = _parse_parameter(m)
        assert abs(p["value"] - 0.07) < 1e-9
        assert p["type"] == "Number"

    def test_text_roundtrip(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m, _parse_parameter
        m = _build_parameter_m("BKK", "Text")
        p = _parse_parameter(m)
        assert p["value"] == "BKK"
        assert p["type"] == "Text"

    def test_escaped_quote_decoded(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m, _parse_parameter
        m = _build_parameter_m('A "B" C', "Text")
        p = _parse_parameter(m)
        assert p["value"] == 'A "B" C'


class TestSetParameterFormula:
    def test_number_literal_replaced(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m, _set_parameter_formula
        old = _build_parameter_m(0.07, "Number")
        new = _set_parameter_formula(old, 0.10)
        assert new.startswith("0.1 meta")
        assert 'Type="Number"' in new
        assert "IsParameterQueryRequired=true" in new

    def test_meta_record_preserved(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m, _set_parameter_formula
        old = _build_parameter_m(5, "Number")
        new = _set_parameter_formula(old, 99)
        assert "IsParameterQuery=true" in new

    def test_text_re_quoted(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m, _set_parameter_formula
        old = _build_parameter_m("BKK", "Text")
        new = _set_parameter_formula(old, "CNX")
        assert new.startswith('"CNX" meta')
        assert 'Type="Text"' in new

    def test_type_override(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m, _set_parameter_formula, _parse_parameter
        old = _build_parameter_m(42, "Number")
        # Override to Text — should re-quote the value
        new = _set_parameter_formula(old, "hello", ptype="Text")
        p = _parse_parameter(new)
        assert p["value"] == "hello"


# ── _list enhancement: is_parameter / is_function flags ──────────────────────

class TestListIsParameterFlag:
    def _make_wb(self, queries: list[tuple[str, str]]) -> MagicMock:
        """Build a fake workbook whose Queries collection yields given (name, formula) pairs."""
        wb = MagicMock()
        wb.Connections.Count = 0
        q_items = []
        for name, formula in queries:
            q = MagicMock()
            q.Name = name
            q.Formula = formula
            q.Description = ""
            q_items.append(q)
        wb.Queries.Count = len(q_items)
        wb.Queries.Item.side_effect = lambda i: q_items[i - 1]
        return wb

    def test_param_query_flagged(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        mock_session = make_mock_session()
        param_formula = _build_parameter_m(0.07, "Number")
        wb = self._make_wb([("pTax", param_formula)])
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            result = powerquery_action(action="list")

        entry = result["queries"][0]
        assert entry["is_parameter"] is True
        assert entry["is_function"] is False

    def test_function_query_flagged(self):
        mock_session = make_mock_session()
        func_formula = "(x as number) => x * 2"
        wb = self._make_wb([("fnDouble", func_formula)])
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            result = powerquery_action(action="list")

        entry = result["queries"][0]
        assert entry["is_function"] is True
        assert entry["is_parameter"] is False

    def test_plain_query_both_false(self):
        mock_session = make_mock_session()
        wb = self._make_wb([("Sales", "let Source = 1 in Source")])
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            result = powerquery_action(action="list")

        entry = result["queries"][0]
        assert entry["is_parameter"] is False
        assert entry["is_function"] is False


# ── Dispatch: create_parameter ────────────────────────────────────────────────

class TestCreateParameterDispatch:
    def _make_wb_empty(self) -> MagicMock:
        wb = MagicMock()
        wb.Queries.Count = 0
        q_mock = MagicMock()
        # Simulate formula round-trip: Queries.Add returns a query whose Formula
        # is the same as what was passed in (COM stores it verbatim).
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        wb.Queries.Add.side_effect = lambda name, formula: _make_query_mock(name, formula)
        return wb

    def test_missing_name_raises(self):
        mock_session = make_mock_session()
        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            with pytest.raises(ToolError, match="name"):
                powerquery_action(action="create_parameter", value=0.07)

    def test_missing_value_raises(self):
        mock_session = make_mock_session()
        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            with pytest.raises(ToolError, match="value"):
                powerquery_action(action="create_parameter", name="pTax")

    def test_duplicate_name_raises(self):
        mock_session = make_mock_session()
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        wb = MagicMock()
        wb.Queries.Count = 1
        existing = MagicMock()
        existing.Name = "pTax"
        wb.Queries.Item.return_value = existing
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            with pytest.raises(ToolError, match="already exists"):
                powerquery_action(action="create_parameter", name="pTax", value=0.07)

    def test_create_number_result_shape(self):
        mock_session = make_mock_session()
        wb = MagicMock()
        wb.Queries.Count = 0
        wb.Queries.Add.side_effect = lambda name, formula: _make_query_mock(name, formula)
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            result = powerquery_action(action="create_parameter", name="pRate", value=0.07, param_type="Number")

        assert result["created_parameter"] == "pRate"
        assert result["type"] == "Number"
        assert abs(result["value"] - 0.07) < 1e-9

    def test_create_text_result_shape(self):
        mock_session = make_mock_session()
        wb = MagicMock()
        wb.Queries.Count = 0
        wb.Queries.Add.side_effect = lambda name, formula: _make_query_mock(name, formula)
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            result = powerquery_action(action="create_parameter", name="pCity", value="BKK", param_type="Text")

        assert result["created_parameter"] == "pCity"
        assert result["type"] == "Text"
        assert result["value"] == "BKK"


# ── Dispatch: get_parameter ───────────────────────────────────────────────────

class TestGetParameterDispatch:
    def test_missing_name_raises(self):
        mock_session = make_mock_session()
        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            with pytest.raises(ToolError, match="name"):
                powerquery_action(action="get_parameter")

    def test_non_parameter_query_raises(self):
        mock_session = make_mock_session()
        wb = MagicMock()
        # _query_obj uses wb.Queries(name) which MagicMock supports via __call__
        plain_q = _make_query_mock("SalesData", "let Source = 1 in Source")
        wb.Queries.side_effect = lambda name: plain_q
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            with pytest.raises(ToolError, match="not a parameter query"):
                powerquery_action(action="get_parameter", name="SalesData")

    def test_get_number_param(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        mock_session = make_mock_session()
        wb = MagicMock()
        param_q = _make_query_mock("pTax", _build_parameter_m(0.07, "Number"))
        wb.Queries.side_effect = lambda name: param_q
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            result = powerquery_action(action="get_parameter", name="pTax")

        assert result["name"] == "pTax"
        assert abs(result["value"] - 0.07) < 1e-9
        assert result["type"] == "Number"
        assert "formula" in result

    def test_get_text_param(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        mock_session = make_mock_session()
        wb = MagicMock()
        param_q = _make_query_mock("pCity", _build_parameter_m("BKK", "Text"))
        wb.Queries.side_effect = lambda name: param_q
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            result = powerquery_action(action="get_parameter", name="pCity")

        assert result["value"] == "BKK"
        assert result["type"] == "Text"


# ── Dispatch: set_parameter ───────────────────────────────────────────────────

class TestSetParameterDispatch:
    def test_missing_name_raises(self):
        mock_session = make_mock_session()
        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            with pytest.raises(ToolError, match="name"):
                powerquery_action(action="set_parameter", value=0.10)

    def test_missing_value_raises(self):
        mock_session = make_mock_session()
        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            with pytest.raises(ToolError, match="value"):
                powerquery_action(action="set_parameter", name="pTax")

    def test_non_parameter_query_raises(self):
        mock_session = make_mock_session()
        wb = MagicMock()
        plain_q = _make_query_mock("SalesData", "let Source = 1 in Source")
        wb.Queries.side_effect = lambda name: plain_q
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            with pytest.raises(ToolError, match="not a parameter query"):
                powerquery_action(action="set_parameter", name="SalesData", value=99)

    def test_set_updates_value(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        mock_session = make_mock_session()
        wb = MagicMock()
        # Mutable mock: Formula setter persists the new value
        param_q = _make_query_mock_mutable("pTax", _build_parameter_m(0.07, "Number"))
        wb.Queries.side_effect = lambda name: param_q
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            result = powerquery_action(action="set_parameter", name="pTax", value=0.10)

        assert result["set_parameter"] == "pTax"
        assert abs(result["value"] - 0.10) < 1e-9
        assert "formula" in result

    def test_meta_preserved_after_set(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        mock_session = make_mock_session()
        wb = MagicMock()
        param_q = _make_query_mock_mutable("pCity", _build_parameter_m("BKK", "Text"))
        wb.Queries.side_effect = lambda name: param_q
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            result = powerquery_action(action="set_parameter", name="pCity", value="CNX")

        assert result["value"] == "CNX"
        assert "IsParameterQuery" in result["formula"]


# ── Dispatch: list_parameters ─────────────────────────────────────────────────

class TestListParametersDispatch:
    def _make_wb_with_queries(self, queries: list[tuple[str, str]]) -> MagicMock:
        wb = MagicMock()
        q_items = []
        for name, formula in queries:
            q = _make_query_mock(name, formula)
            q_items.append(q)
        wb.Queries.Count = len(q_items)
        wb.Queries.Item.side_effect = lambda i: q_items[i - 1]
        return wb

    def test_returns_only_parameters(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        mock_session = make_mock_session()
        wb = self._make_wb_with_queries([
            ("pTax", _build_parameter_m(0.07, "Number")),
            ("SalesData", "let Source = 1 in Source"),
            ("pCity", _build_parameter_m("BKK", "Text")),
        ])
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            result = powerquery_action(action="list_parameters")

        assert result["count"] == 2
        names = {p["name"] for p in result["parameters"]}
        assert names == {"pTax", "pCity"}

    def test_empty_workbook_returns_empty(self):
        mock_session = make_mock_session()
        wb = self._make_wb_with_queries([])
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            result = powerquery_action(action="list_parameters")

        assert result["count"] == 0
        assert result["parameters"] == []

    def test_entry_shape(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        mock_session = make_mock_session()
        wb = self._make_wb_with_queries([
            ("pRate", _build_parameter_m(0.05, "Number")),
        ])
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            result = powerquery_action(action="list_parameters")

        p = result["parameters"][0]
        assert "name" in p
        assert "value" in p
        assert "type" in p


# ── Unknown-action error includes new action names ────────────────────────────

class TestUnknownActionListsNewActions:
    def test_all_new_actions_in_error(self):
        mock_session = make_mock_session()
        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            with pytest.raises(ToolError) as exc_info:
                powerquery_action(action="nonexistent_action_xyz")
        msg = str(exc_info.value)
        for name in ("create_parameter", "get_parameter", "set_parameter", "list_parameters"):
            assert name in msg, f"Expected '{name}' in error message, got: {msg}"

    def test_existing_actions_still_in_error(self):
        mock_session = make_mock_session()
        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            with pytest.raises(ToolError) as exc_info:
                powerquery_action(action="bogus")
        msg = str(exc_info.value)
        for name in ("list", "get", "create", "update", "delete", "analyze_raw"):
            assert name in msg


# ── Fix 1: text values containing ' meta ' survive round-trip ────────────────

class TestMetaInTextValue:
    """Values that contain the substring ' meta ' must survive build→set→parse."""

    def test_parse_text_with_meta_substring(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m, _parse_parameter
        m = _build_parameter_m("has meta inside", "Text")
        p = _parse_parameter(m)
        assert p["value"] == "has meta inside"
        assert p["type"] == "Text"
        # meta record preserved intact
        assert "IsParameterQuery=true" in m

    def test_set_text_with_meta_substring(self):
        from thepexcel_mcp.domains.powerquery import _build_parameter_m, _set_parameter_formula, _parse_parameter
        old = _build_parameter_m("has meta inside", "Text")
        new = _set_parameter_formula(old, "also meta here")
        p = _parse_parameter(new)
        assert p["value"] == "also meta here"
        assert "IsParameterQuery=true" in new
        # Verify the meta record is still structurally valid (quotes balanced)
        assert new.startswith('"also meta here" meta')

    def test_roundtrip_meta_in_value(self):
        """build → set → parse should preserve value containing ' meta ' exactly."""
        from thepexcel_mcp.domains.powerquery import _build_parameter_m, _set_parameter_formula, _parse_parameter
        original = "region meta data"
        old = _build_parameter_m(original, "Text")
        # First parse is correct
        assert _parse_parameter(old)["value"] == original
        # Set to a different value also containing meta
        new_val = "meta prefix label"
        new = _set_parameter_formula(old, new_val)
        assert _parse_parameter(new)["value"] == new_val
        # Meta record fields still present
        assert "IsParameterQueryRequired=" in new

    def test_meta_in_value_quotes_balanced(self):
        """The formula produced for a 'meta'-containing text value must have balanced quotes."""
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        m = _build_parameter_m("has meta inside", "Text")
        # The literal part is the first token; count unescaped quotes
        # Strip leading quote, find closing quote ignoring ""
        literal_part = m.split(" meta [")[0]  # safe: meta [ is not in the value
        # Actually use the helper to confirm structure
        from thepexcel_mcp.domains.powerquery import _split_literal_and_meta
        literal, meta_part = _split_literal_and_meta(m)
        assert literal == '"has meta inside"'
        assert "meta [IsParameterQuery" in meta_part


# ── Fix 2: falsy values (0, 0.0, '') are accepted ────────────────────────────

class TestFalsyValues:
    """value=0, value=0.0, and value='' must NOT be rejected by the guard."""

    def test_create_value_zero_accepted(self):
        """value=0 must not raise — it's a legitimate parameter value."""
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        mock_session = make_mock_session()
        wb = MagicMock()
        wb.Queries.Count = 0
        wb.Queries.Add.side_effect = lambda name, formula: _make_query_mock(name, formula)
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            # Should NOT raise
            result = powerquery_action(action="create_parameter", name="MaxRows", value=0)
        assert result["created_parameter"] == "MaxRows"
        assert result["value"] == 0

    def test_create_value_empty_string_accepted(self):
        """value='' must not raise — empty string is a valid Text parameter."""
        mock_session = make_mock_session()
        wb = MagicMock()
        wb.Queries.Count = 0
        wb.Queries.Add.side_effect = lambda name, formula: _make_query_mock(name, formula)
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            result = powerquery_action(action="create_parameter", name="pEmpty",
                                       value="", param_type="Text")
        assert result["created_parameter"] == "pEmpty"
        assert result["value"] == ""

    def test_set_value_zero_accepted(self):
        """value=0 in set_parameter must not raise."""
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        mock_session = make_mock_session()
        wb = MagicMock()
        param_q = _make_query_mock_mutable("MaxRows", _build_parameter_m(100, "Number"))
        wb.Queries.side_effect = lambda name: param_q
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            result = powerquery_action(action="set_parameter", name="MaxRows", value=0)
        assert result["set_parameter"] == "MaxRows"
        assert result["value"] == 0

    def test_build_zero_value(self):
        """value=0 must produce a bare '0 meta ...' literal (not quoted)."""
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        m = _build_parameter_m(0, "Number")
        assert m.startswith("0 meta")

    def test_build_empty_string_value(self):
        """value='' must produce a '\"\" meta ...' literal (empty quoted string)."""
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        m = _build_parameter_m("", "Text")
        assert m.startswith('"" meta')

    def test_create_value_none_still_raises(self):
        """value=None must still raise (it is the sentinel for 'not provided')."""
        mock_session = make_mock_session()
        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            with pytest.raises(ToolError, match="value"):
                powerquery_action(action="create_parameter", name="p", value=None)

    def test_set_value_none_still_raises(self):
        """value=None must still raise."""
        mock_session = make_mock_session()
        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            with pytest.raises(ToolError, match="value"):
                powerquery_action(action="set_parameter", name="p", value=None)


# ── Fix 4: verify-effect compares against requested value ────────────────────

class TestVerifyEffectAgainstRequestedValue:
    """_set_parameter must detect a mismatch between requested and read-back value."""

    def test_verify_effect_passes_when_value_matches(self):
        """Normal set with matching read-back must not raise."""
        from thepexcel_mcp.domains.powerquery import _build_parameter_m
        mock_session = make_mock_session()
        wb = MagicMock()
        param_q = _make_query_mock_mutable("pRate", _build_parameter_m(0.07, "Number"))
        wb.Queries.side_effect = lambda name: param_q
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            result = powerquery_action(action="set_parameter", name="pRate", value=0.10)
        assert abs(result["value"] - 0.10) < 1e-9

    def test_verify_effect_detects_mismatch(self):
        """If COM silently writes a different value, verify-effect must raise ToolError."""
        from thepexcel_mcp.domains.powerquery import _build_parameter_m, _set_parameter
        # Build a query mock whose Formula setter ignores the write (simulates silent noop)
        initial_formula = _build_parameter_m(0.07, "Number")

        class _SilentNoopQuery:
            Name = "pRate"
            Description = ""

            @property
            def Formula(self):
                # Always returns the original — setter is a no-op
                return initial_formula

            @Formula.setter
            def Formula(self, value):
                pass  # silently ignores the write

        mock_session = make_mock_session()
        wb = MagicMock()
        wb.Queries.side_effect = lambda name: _SilentNoopQuery()
        mock_session.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.powerquery._session", mock_session):
            from thepexcel_mcp.domains.powerquery import powerquery_action
            with pytest.raises(ToolError, match="verify-effect mismatch"):
                powerquery_action(action="set_parameter", name="pRate", value=0.99)


# ── Shared mock builders ──────────────────────────────────────────────────────

def _make_query_mock(name: str, formula: str) -> MagicMock:
    """A query mock whose Formula is fixed (read-only simulation)."""
    q = MagicMock()
    q.Name = name
    q.Formula = formula
    q.Description = ""
    return q


def _make_query_mock_mutable(name: str, initial_formula: str) -> MagicMock:
    """A query mock that simulates COM Formula property with getter+setter.

    When q.Formula = new_val is called, subsequent q.Formula reads return new_val.
    This lets set_parameter's verify-effect read-back see the updated formula.
    """
    state = {"formula": initial_formula}

    class _MutableQuery:
        Name = name
        Description = ""

        @property
        def Formula(self):
            return state["formula"]

        @Formula.setter
        def Formula(self, value):
            state["formula"] = value

    return _MutableQuery()
