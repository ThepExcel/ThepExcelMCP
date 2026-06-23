"""Unit tests for excel_validation (validation.py).

No Excel required — all COM calls are intercepted via make_mock_session.
Follows the same pattern as test_format_workbook.py:
  - patch _resolve_range + _session + excel_guard
  - assert SPECIFIC COM property/method was set to the SPECIFIC value
"""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock, call, patch

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


# ── Mock builders ──────────────────────────────────────────────────────────────

def _make_range_mock(address: str = "$B$2:$B$10"):
    """Minimal Range mock with a Validation sub-object."""
    rng = MagicMock()
    rng.Address = address
    rng.Application = MagicMock()

    # Validation is a separate COM object; give it its own mock
    val = MagicMock()
    rng.Validation = val
    return rng


def _call_validation(action, rng_mock, mock_session, **kwargs):
    """Patch _resolve_range + _session + excel_guard, then call validation_action."""
    with patch("thepexcel_mcp.domains.validation._resolve_range", return_value=rng_mock):
        with patch("thepexcel_mcp.domains.validation._session", mock_session):
            with patch("thepexcel_mcp.domains.validation.excel_guard") as eg:
                eg.return_value.__enter__ = MagicMock(return_value=None)
                eg.return_value.__exit__ = MagicMock(return_value=False)
                from thepexcel_mcp.domains.validation import validation_action
                return validation_action(
                    action,
                    range="B2:B10",
                    sheet=None,
                    workbook=None,
                    **kwargs,
                )


# ── Dispatch validation ────────────────────────────────────────────────────────

class TestValidationActionDispatch:
    def test_unknown_action_raises(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="Unknown action"):
            _call_validation("bogus", rng, mock_session)

    def test_all_valid_action_names_enumerated_in_error(self):
        """Unknown action error message should list all valid actions."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError) as exc_info:
            _call_validation("unknown", rng, mock_session)
        msg = str(exc_info.value)
        for name in ("list", "whole_number", "decimal", "date", "text_length", "custom", "clear"):
            assert name in msg


# ── clear action ──────────────────────────────────────────────────────────────

class TestClearAction:
    def test_delete_called(self):
        """clear must call Validation.Delete() on the range."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_validation("clear", rng, mock_session)
        rng.Validation.Delete.assert_called_once()

    def test_clear_returns_cleared_true(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_validation("clear", rng, mock_session)
        assert result["validation"] == "clear"
        assert result["applied"]["cleared"] is True


# ── list action ───────────────────────────────────────────────────────────────

class TestListAction:
    def test_requires_formula1(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="formula1"):
            _call_validation("list", rng, mock_session)

    def test_delete_called_before_add(self):
        """Delete must be called before Add to avoid COM error 1004."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("list", rng, mock_session, formula1="Yes,No,Maybe")
        # Delete called first on the Validation object
        rng.Validation.Delete.assert_called_once()

    def test_add_called_with_list_type(self):
        """Validation.Add must be called with Type=3 (xlValidateList)."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("list", rng, mock_session, formula1="Yes,No,Maybe")
        rng.Validation.Add.assert_called_once()
        kwargs = rng.Validation.Add.call_args.kwargs
        assert kwargs["Type"] == 3, f"Expected Type=3 (xlValidateList), got {kwargs['Type']}"
        assert kwargs["Formula1"] == "Yes,No,Maybe"
        assert "Operator" not in kwargs, "Operator must be OMITTED for list type"

    def test_add_called_with_alert_stop(self):
        """AlertStyle must be xlValidAlertStop=1."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("list", rng, mock_session, formula1="a,b,c")
        kwargs = rng.Validation.Add.call_args.kwargs
        assert kwargs["AlertStyle"] == 1, f"Expected AlertStyle=1 (xlValidAlertStop), got {kwargs['AlertStyle']}"

    def test_in_cell_dropdown_set(self):
        """After Add, InCellDropdown must be set on the Validation object."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("list", rng, mock_session, formula1="a,b,c", in_cell_dropdown=True)
        assert rng.Validation.InCellDropdown is True

    def test_in_cell_dropdown_false(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("list", rng, mock_session, formula1="a,b,c", in_cell_dropdown=False)
        assert rng.Validation.InCellDropdown is False

    def test_ignore_blank_set(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("list", rng, mock_session, formula1="a,b,c", ignore_blank=True)
        assert rng.Validation.IgnoreBlank is True

    def test_show_error_set(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("list", rng, mock_session, formula1="a,b,c", show_error=True)
        assert rng.Validation.ShowError is True

    def test_range_ref_formula(self):
        """Formula1 may be a range reference like =$A$1:$A$5."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_validation("list", rng, mock_session, formula1="=$A$1:$A$5")
        assert result["applied"]["formula1"] == "=$A$1:$A$5"

    def test_result_shape(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_validation("list", rng, mock_session, formula1="Yes,No")
        assert result["validation"] == "list"
        assert "range" in result
        assert result["applied"]["type"] == "list"


# ── custom action ─────────────────────────────────────────────────────────────

class TestCustomAction:
    def test_requires_formula1(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="formula1"):
            _call_validation("custom", rng, mock_session)

    def test_delete_called_before_add(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("custom", rng, mock_session, formula1="=ISNUMBER(B2)")
        rng.Validation.Delete.assert_called_once()

    def test_add_called_with_custom_type(self):
        """Type must be 7 (xlValidateCustom); Operator must be OMITTED."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("custom", rng, mock_session, formula1="=ISNUMBER(B2)")
        kwargs = rng.Validation.Add.call_args.kwargs
        assert kwargs["Type"] == 7, f"Expected Type=7 (xlValidateCustom), got {kwargs['Type']}"
        assert kwargs["Formula1"] == "=ISNUMBER(B2)"
        assert "Operator" not in kwargs, "Operator must be OMITTED for custom type"

    def test_result_shape(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_validation("custom", rng, mock_session, formula1="=LEN(B2)<=50")
        assert result["validation"] == "custom"
        assert result["applied"]["type"] == "custom"
        assert result["applied"]["formula1"] == "=LEN(B2)<=50"


# ── whole_number action ───────────────────────────────────────────────────────

class TestWholeNumberAction:
    def test_requires_formula1(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="formula1"):
            _call_validation("whole_number", rng, mock_session, operator="between")

    def test_requires_operator(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="operator"):
            _call_validation("whole_number", rng, mock_session, formula1="1")

    def test_unknown_operator_raises(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="operator"):
            _call_validation("whole_number", rng, mock_session,
                             formula1="1", operator="nope")

    def test_between_requires_formula2(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="formula2"):
            _call_validation("whole_number", rng, mock_session,
                             formula1="1", operator="between")

    def test_delete_called_before_add(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("whole_number", rng, mock_session,
                         formula1="5", formula2="10", operator="between")
        rng.Validation.Delete.assert_called_once()

    def test_add_with_between_operator(self):
        """Type=1 (xlValidateWholeNumber), Operator=1 (xlBetween), both formulas passed."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("whole_number", rng, mock_session,
                         formula1="5", formula2="10", operator="between")
        kwargs = rng.Validation.Add.call_args.kwargs
        assert kwargs["Type"] == 1, f"Expected 1 (xlValidateWholeNumber), got {kwargs['Type']}"
        assert kwargs["Operator"] == 1, f"Expected 1 (xlBetween), got {kwargs['Operator']}"
        assert kwargs["Formula1"] == "5"
        assert kwargs["Formula2"] == "10"

    def test_add_with_greater_operator(self):
        """Operator=5 (xlGreater); Formula2 must NOT be passed."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("whole_number", rng, mock_session,
                         formula1="0", operator="greater")
        kwargs = rng.Validation.Add.call_args.kwargs
        assert kwargs["Operator"] == 5, f"Expected 5 (xlGreater), got {kwargs['Operator']}"
        assert "Formula2" not in kwargs

    def test_add_with_less_equal_operator(self):
        """Operator=8 (xlLessEqual)."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("whole_number", rng, mock_session,
                         formula1="100", operator="less_equal")
        kwargs = rng.Validation.Add.call_args.kwargs
        assert kwargs["Operator"] == 8, f"Expected 8 (xlLessEqual), got {kwargs['Operator']}"

    def test_result_shape(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_validation("whole_number", rng, mock_session,
                                  formula1="1", formula2="99", operator="between")
        assert result["validation"] == "whole_number"
        assert result["applied"]["type"] == "whole_number"
        assert result["applied"]["operator"] == "between"
        assert result["applied"]["formula1"] == "1"
        assert result["applied"]["formula2"] == "99"


# ── decimal action ────────────────────────────────────────────────────────────

class TestDecimalAction:
    def test_type_is_2(self):
        """xlValidateDecimal = 2."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("decimal", rng, mock_session,
                         formula1="0.0", formula2="1.0", operator="between")
        kwargs = rng.Validation.Add.call_args.kwargs
        assert kwargs["Type"] == 2, f"Expected 2 (xlValidateDecimal), got {kwargs['Type']}"

    def test_equal_operator(self):
        """Operator=3 (xlEqual)."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("decimal", rng, mock_session,
                         formula1="3.14", operator="equal")
        kwargs = rng.Validation.Add.call_args.kwargs
        assert kwargs["Operator"] == 3


# ── date action ───────────────────────────────────────────────────────────────

class TestDateAction:
    def test_type_is_4(self):
        """xlValidateDate = 4."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("date", rng, mock_session,
                         formula1="2026-01-01", operator="greater_equal")
        kwargs = rng.Validation.Add.call_args.kwargs
        assert kwargs["Type"] == 4, f"Expected 4 (xlValidateDate), got {kwargs['Type']}"

    def test_not_between_operator(self):
        """Operator=2 (xlNotBetween) — requires formula2."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("date", rng, mock_session,
                         formula1="2026-01-01", formula2="2026-12-31",
                         operator="not_between")
        kwargs = rng.Validation.Add.call_args.kwargs
        assert kwargs["Operator"] == 2
        assert kwargs["Formula2"] == "2026-12-31"


# ── text_length action ────────────────────────────────────────────────────────

class TestTextLengthAction:
    def test_type_is_6(self):
        """xlValidateTextLength = 6."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("text_length", rng, mock_session,
                         formula1="50", operator="less_equal")
        kwargs = rng.Validation.Add.call_args.kwargs
        assert kwargs["Type"] == 6, f"Expected 6 (xlValidateTextLength), got {kwargs['Type']}"

    def test_not_equal_operator(self):
        """Operator=4 (xlNotEqual)."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_validation("text_length", rng, mock_session,
                         formula1="0", operator="not_equal")
        kwargs = rng.Validation.Add.call_args.kwargs
        assert kwargs["Operator"] == 4


# ── Operator enum coverage ────────────────────────────────────────────────────

class TestOperatorEnum:
    """Verify each operator string maps to the correct COM integer."""

    @pytest.mark.parametrize("op_str, expected_int", [
        ("between",       1),
        ("not_between",   2),
        ("equal",         3),
        ("not_equal",     4),
        ("greater",       5),
        ("less",          6),
        ("greater_equal", 7),
        ("less_equal",    8),
    ])
    def test_operator_mapping(self, op_str, expected_int):
        from thepexcel_mcp.domains.validation import _OPERATORS
        assert _OPERATORS[op_str] == expected_int, (
            f"Operator '{op_str}' should map to {expected_int}, "
            f"got {_OPERATORS[op_str]}"
        )


# ── XlDVType enum coverage ────────────────────────────────────────────────────

class TestDVTypeEnum:
    """Verify the XlDVType constants match the official Microsoft docs."""

    def test_whole_number_is_1(self):
        from thepexcel_mcp.domains.validation import _XL_VALIDATE_WHOLE_NUMBER
        assert _XL_VALIDATE_WHOLE_NUMBER == 1

    def test_decimal_is_2(self):
        from thepexcel_mcp.domains.validation import _XL_VALIDATE_DECIMAL
        assert _XL_VALIDATE_DECIMAL == 2

    def test_list_is_3(self):
        from thepexcel_mcp.domains.validation import _XL_VALIDATE_LIST
        assert _XL_VALIDATE_LIST == 3

    def test_date_is_4(self):
        from thepexcel_mcp.domains.validation import _XL_VALIDATE_DATE
        assert _XL_VALIDATE_DATE == 4

    def test_text_length_is_6(self):
        from thepexcel_mcp.domains.validation import _XL_VALIDATE_TEXT_LENGTH
        assert _XL_VALIDATE_TEXT_LENGTH == 6

    def test_custom_is_7(self):
        from thepexcel_mcp.domains.validation import _XL_VALIDATE_CUSTOM
        assert _XL_VALIDATE_CUSTOM == 7
