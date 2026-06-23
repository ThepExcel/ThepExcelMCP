"""Range data validation: dropdowns, constraints, custom formulas.

Supported types
---------------
list            xlValidateList=3      — dropdown from a,b,c or =$A$1:$A$10
whole_number    xlValidateWholeNumber=1
decimal         xlValidateDecimal=2
date            xlValidateDate=4
text_length     xlValidateTextLength=6
custom          xlValidateCustom=7    — arbitrary Excel formula
clear           — removes all validation from the range

Delete-before-Add (critical)
----------------------------
Calling rng.Validation.Add(...) on a range that already has validation raises
COM error 1004.  We always call rng.Validation.Delete() first.

Operator is required for number/date/text_length types, and must be OMITTED
(not even passed as a keyword argument) for list and custom types.

Color note: this module does not handle colors; import _hex_to_bgr from format
only if needed in future extensions.
"""

from __future__ import annotations

from fastmcp.exceptions import ToolError

from ..session import ExcelSession, excel_guard

_session = ExcelSession()

# XlDVType — confirmed from Microsoft Learn docs
_XL_VALIDATE_WHOLE_NUMBER = 1
_XL_VALIDATE_DECIMAL      = 2
_XL_VALIDATE_LIST         = 3
_XL_VALIDATE_DATE         = 4
_XL_VALIDATE_TEXT_LENGTH  = 6
_XL_VALIDATE_CUSTOM       = 7

# XlDVAlertStyle
_XL_VALID_ALERT_STOP = 1  # xlValidAlertStop

# XlFormatConditionOperator (shared with conditional format) — confirmed from docs
_OPERATORS = {
    "between":       1,  # xlBetween
    "not_between":   2,  # xlNotBetween
    "equal":         3,  # xlEqual
    "not_equal":     4,  # xlNotEqual
    "greater":       5,  # xlGreater
    "less":          6,  # xlLess
    "greater_equal": 7,  # xlGreaterEqual
    "less_equal":    8,  # xlLessEqual
}

# Types that require an operator
_OPERATOR_REQUIRED_TYPES = {
    "whole_number", "decimal", "date", "text_length",
}

# Map action name → XlDVType int
_ACTION_TO_TYPE = {
    "whole_number": _XL_VALIDATE_WHOLE_NUMBER,
    "decimal":      _XL_VALIDATE_DECIMAL,
    "date":         _XL_VALIDATE_DATE,
    "text_length":  _XL_VALIDATE_TEXT_LENGTH,
}


# ── Public entry point ─────────────────────────────────────────────────────────

def validation_action(
    action: str,
    range: str,
    sheet: str | None = None,
    workbook: str | None = None,
    # list / custom / constraint formula
    formula1: str | None = None,
    formula2: str | None = None,
    # operator for number/date/text_length
    operator: str | None = None,
    # list-specific: show dropdown in cell
    in_cell_dropdown: bool = True,
    # common
    ignore_blank: bool = True,
    show_error: bool = True,
) -> dict:
    """Dispatch a data-validation action.

    Actions
    -------
    list            Add a dropdown list validation.
    whole_number    Restrict entry to whole integers.
    decimal         Restrict entry to decimal numbers.
    date            Restrict entry to dates.
    text_length     Restrict text length.
    custom          Validate against an arbitrary formula.
    clear           Remove all validation from the range.
    """
    valid = {
        "list", "whole_number", "decimal", "date",
        "text_length", "custom", "clear",
    }
    if action not in valid:
        raise ToolError(
            f"Unknown action '{action}'. Valid: {', '.join(sorted(valid))}."
        )
    return _session.run_com(
        _dispatch,
        action, range, sheet, workbook,
        formula1, formula2, operator,
        in_cell_dropdown, ignore_blank, show_error,
    )


# ── Worker-thread dispatcher ───────────────────────────────────────────────────

def _dispatch(
    action, range_str, sheet, workbook,
    formula1, formula2, operator,
    in_cell_dropdown, ignore_blank, show_error,
) -> dict:
    """Executed on the STA COM worker thread."""
    rng = _resolve_range(range_str, sheet, workbook)
    app = rng.Application

    with excel_guard(app):
        if action == "clear":
            return _clear(rng)
        if action == "list":
            return _list(rng, formula1, in_cell_dropdown, ignore_blank, show_error)
        if action == "custom":
            return _custom(rng, formula1, ignore_blank, show_error)
        # whole_number / decimal / date / text_length
        return _constraint(rng, action, formula1, formula2, operator, ignore_blank, show_error)


# ── Range resolver (mirrors format.py verbatim) ───────────────────────────────

def _resolve_range(range_str: str, sheet: str | None, workbook: str | None):
    if "!" in range_str:
        sheet_part, cell_part = range_str.split("!", 1)
        ws = _session.get_sheet(sheet_part.strip("'"), workbook)
        try:
            return ws.Range(cell_part)
        except Exception as e:
            raise _session.wrap(e, f"Invalid range '{cell_part}' on sheet '{sheet_part}'")
    ws = _session.get_sheet(sheet, workbook)
    try:
        return ws.Range(range_str)
    except Exception as e:
        raise _session.wrap(e, f"Invalid range '{range_str}'")


# ── Action implementations ────────────────────────────────────────────────────

def _clear(rng) -> dict:
    """Remove all validation from the range."""
    try:
        rng.Validation.Delete()
    except Exception as e:
        raise _session.wrap(e, "Validation.Delete failed")
    return {"validation": "clear", "range": rng.Address, "applied": {"cleared": True}}


def _list(rng, formula1: str | None, in_cell_dropdown: bool, ignore_blank: bool, show_error: bool) -> dict:
    """Add dropdown list validation."""
    if not formula1:
        raise ToolError(
            "list action requires 'formula1': comma-separated values "
            "(e.g. 'Yes,No,Maybe') or a range reference (e.g. '=$A$1:$A$5')."
        )
    try:
        # Delete-before-Add: always required to avoid COM error 1004
        rng.Validation.Delete()
        # Operator must be OMITTED for list type
        rng.Validation.Add(
            Type=_XL_VALIDATE_LIST,
            AlertStyle=_XL_VALID_ALERT_STOP,
            Formula1=formula1,
        )
        v = rng.Validation
        v.InCellDropdown = in_cell_dropdown
        v.IgnoreBlank = ignore_blank
        v.ShowError = show_error
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "list validation failed")
    return {
        "validation": "list",
        "range": rng.Address,
        "applied": {
            "type": "list",
            "formula1": formula1,
            "in_cell_dropdown": in_cell_dropdown,
        },
    }


def _custom(rng, formula1: str | None, ignore_blank: bool, show_error: bool) -> dict:
    """Add custom-formula validation."""
    if not formula1:
        raise ToolError(
            "custom action requires 'formula1': an Excel formula that evaluates "
            "to TRUE for valid data (e.g. '=ISNUMBER(A1)')."
        )
    try:
        rng.Validation.Delete()
        # Operator must be OMITTED for custom type
        rng.Validation.Add(
            Type=_XL_VALIDATE_CUSTOM,
            AlertStyle=_XL_VALID_ALERT_STOP,
            Formula1=formula1,
        )
        v = rng.Validation
        v.IgnoreBlank = ignore_blank
        v.ShowError = show_error
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "custom validation failed")
    return {
        "validation": "custom",
        "range": rng.Address,
        "applied": {
            "type": "custom",
            "formula1": formula1,
        },
    }


def _constraint(
    rng, action: str, formula1: str | None, formula2: str | None,
    operator: str | None, ignore_blank: bool, show_error: bool,
) -> dict:
    """Add whole_number / decimal / date / text_length validation."""
    if not formula1:
        raise ToolError(
            f"{action} action requires 'formula1' (the constraint value or first bound)."
        )
    if operator is None:
        raise ToolError(
            f"{action} action requires 'operator'. "
            f"Valid: {', '.join(sorted(_OPERATORS))}."
        )
    op_int = _OPERATORS.get(operator)
    if op_int is None:
        raise ToolError(
            f"Unknown operator '{operator}'. "
            f"Valid: {', '.join(sorted(_OPERATORS))}."
        )
    # formula2 is required when operator is between / not_between
    if operator in ("between", "not_between") and not formula2:
        raise ToolError(
            f"operator='{operator}' requires 'formula2' (the upper bound)."
        )
    type_int = _ACTION_TO_TYPE[action]
    try:
        rng.Validation.Delete()
        if operator in ("between", "not_between"):
            rng.Validation.Add(
                Type=type_int,
                AlertStyle=_XL_VALID_ALERT_STOP,
                Operator=op_int,
                Formula1=formula1,
                Formula2=formula2,
            )
        else:
            rng.Validation.Add(
                Type=type_int,
                AlertStyle=_XL_VALID_ALERT_STOP,
                Operator=op_int,
                Formula1=formula1,
            )
        v = rng.Validation
        v.IgnoreBlank = ignore_blank
        v.ShowError = show_error
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"{action} validation failed")
    applied: dict = {
        "type": action,
        "operator": operator,
        "formula1": formula1,
    }
    if formula2 is not None:
        applied["formula2"] = formula2
    return {"validation": action, "range": rng.Address, "applied": applied}
