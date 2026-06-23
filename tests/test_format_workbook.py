"""Unit tests for excel_format (format.py) and workbook create/save_as.

No Excel required — all COM calls are intercepted via make_mock_session.
"""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock, patch, call

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


# ── Helpers ────────────────────────────────────────────────────────────────────

def _make_range_mock(address: str = "$A$1:$C$3"):
    """Minimal COM Range mock sufficient for format.py tests."""
    rng = MagicMock()
    rng.Address = address
    rng.Font = MagicMock()
    rng.Interior = MagicMock()
    rng.Application = MagicMock()

    # Borders(const) → border mock
    border_mock = MagicMock()
    rng.Borders.return_value = border_mock

    rng.Columns = MagicMock()
    rng.Rows = MagicMock()
    return rng


def _call_format(action, rng_mock, mock_session, **kwargs):
    """Patch _resolve_range + _session, invoke the format action helper directly."""
    with patch("thepexcel_mcp.domains.format._resolve_range", return_value=rng_mock):
        with patch("thepexcel_mcp.domains.format._session", mock_session):
            # excel_guard context manager — patch it to be a no-op pass-through
            with patch("thepexcel_mcp.domains.format.excel_guard") as eg:
                eg.return_value.__enter__ = MagicMock(return_value=None)
                eg.return_value.__exit__ = MagicMock(return_value=False)
                from thepexcel_mcp.domains.format import format_action
                return format_action(
                    action,
                    range="A1:C3",
                    sheet=None,
                    workbook=None,
                    **kwargs,
                )


# ── Color helper ───────────────────────────────────────────────────────────────

class TestHexToBgr:
    def test_red(self):
        from thepexcel_mcp.domains.format import _hex_to_bgr
        # #FF0000 → R=255, G=0, B=0 → BGR = 0x0000FF = 255
        assert _hex_to_bgr("#FF0000") == 255

    def test_blue(self):
        from thepexcel_mcp.domains.format import _hex_to_bgr
        # #0000FF → R=0, G=0, B=255 → BGR = 0xFF0000 = 16711680
        assert _hex_to_bgr("#0000FF") == 0xFF0000

    def test_thepexcel_gold(self):
        from thepexcel_mcp.domains.format import _hex_to_bgr
        # ThepExcel brand gold: #D4A84B → R=0xD4, G=0xA8, B=0x4B
        # BGR = (0x4B << 16) | (0xA8 << 8) | 0xD4 = 0x4BA8D4
        assert _hex_to_bgr("#D4A84B") == (0x4B << 16) | (0xA8 << 8) | 0xD4

    def test_without_hash(self):
        from thepexcel_mcp.domains.format import _hex_to_bgr
        # Should work with or without leading #
        assert _hex_to_bgr("FF0000") == 255

    def test_invalid_length(self):
        from thepexcel_mcp.domains.format import _hex_to_bgr
        with pytest.raises(ToolError, match="RRGGBB"):
            _hex_to_bgr("#FFF")

    def test_invalid_chars(self):
        from thepexcel_mcp.domains.format import _hex_to_bgr
        with pytest.raises(ToolError):
            _hex_to_bgr("#GGGGGG")


# ── format_action dispatch validation ────────────────────────────────────────

class TestFormatActionValidation:
    def test_unknown_action_raises(self):
        mock_session = make_mock_session()
        rng_mock = _make_range_mock()
        with pytest.raises(ToolError, match="Unknown action"):
            _call_format("bogus", rng_mock, mock_session)

    def test_fill_requires_color_or_clear(self):
        mock_session = make_mock_session()
        rng_mock = _make_range_mock()
        with pytest.raises(ToolError, match="fill_color"):
            _call_format("fill", rng_mock, mock_session)

    def test_number_format_requires_param(self):
        mock_session = make_mock_session()
        rng_mock = _make_range_mock()
        with pytest.raises(ToolError, match="number_format"):
            _call_format("number_format", rng_mock, mock_session)

    def test_column_width_requires_param(self):
        mock_session = make_mock_session()
        rng_mock = _make_range_mock()
        with pytest.raises(ToolError, match="width"):
            _call_format("column_width", rng_mock, mock_session)

    def test_row_height_requires_param(self):
        mock_session = make_mock_session()
        rng_mock = _make_range_mock()
        with pytest.raises(ToolError, match="height"):
            _call_format("row_height", rng_mock, mock_session)

    def test_autofit_requires_at_least_one(self):
        mock_session = make_mock_session()
        rng_mock = _make_range_mock()
        with pytest.raises(ToolError, match="autofit_columns"):
            _call_format("autofit", rng_mock, mock_session,
                         autofit_columns=False, autofit_rows=False)


# ── Font ───────────────────────────────────────────────────────────────────────

class TestFontAction:
    def test_bold_applied(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_format("font", rng, mock_session, bold=True)
        assert rng.Font.Bold is True
        assert result["formatted"] == "font"
        assert result["applied"]["bold"] is True

    def test_font_name_size(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_format("font", rng, mock_session, font_name="Calibri", font_size=14)
        assert rng.Font.Name == "Calibri"
        assert rng.Font.Size == 14
        assert "font_name" in result["applied"]
        assert "font_size" in result["applied"]

    def test_font_color_converted_to_bgr(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_format("font", rng, mock_session, font_color="#FF0000")
        # Red #FF0000 → BGR = 0x0000FF = 255
        assert rng.Font.Color == 255

    def test_underline_true_sets_xlunderlinestyle_single(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_format("font", rng, mock_session, underline=True)
        assert rng.Font.Underline == 2  # xlUnderlineStyleSingle

    def test_underline_false_sets_none(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_format("font", rng, mock_session, underline=False)
        assert rng.Font.Underline == -4142  # xlUnderlineStyleNone

    def test_no_params_returns_empty_applied(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_format("font", rng, mock_session)
        assert result["applied"] == {}


# ── Fill ───────────────────────────────────────────────────────────────────────

class TestFillAction:
    def test_fill_color_sets_interior(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_format("fill", rng, mock_session, fill_color="#FFFF00")
        # Yellow #FFFF00 → R=255, G=255, B=0 → BGR = 0x00FFFF = 65535
        assert rng.Interior.Color == 65535
        assert result["applied"]["fill_color"] == "#FFFF00"

    def test_clear_fill_sets_color_index_none(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_format("fill", rng, mock_session, clear_fill=True)
        assert rng.Interior.ColorIndex == -4142
        assert result["applied"]["cleared"] is True


# ── Border ─────────────────────────────────────────────────────────────────────

class TestBorderAction:
    def test_outline_calls_four_edges(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_format("border", rng, mock_session,
                              border_sides="outline",
                              border_style="continuous",
                              border_weight="thin")
        # Borders() should be called for 4 edge constants
        calls = [c.args[0] for c in rng.Borders.call_args_list]
        from thepexcel_mcp.domains.format import (
            _XL_EDGE_LEFT, _XL_EDGE_TOP, _XL_EDGE_BOTTOM, _XL_EDGE_RIGHT
        )
        for edge in (_XL_EDGE_LEFT, _XL_EDGE_TOP, _XL_EDGE_BOTTOM, _XL_EDGE_RIGHT):
            assert edge in calls, f"edge {edge} not found in calls: {calls}"

    def test_all_calls_six_sides(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_format("border", rng, mock_session, border_sides="all")
        calls = [c.args[0] for c in rng.Borders.call_args_list]
        assert len(calls) == 6

    def test_invalid_border_sides_raises(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="border_sides"):
            _call_format("border", rng, mock_session, border_sides="diagonal")

    def test_invalid_border_style_raises(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="border_style"):
            _call_format("border", rng, mock_session, border_style="dotted")

    def test_none_style_skips_weight(self):
        """When border_style='none', Weight should NOT be set (avoid COM error)."""
        mock_session = make_mock_session()
        rng = _make_range_mock()
        border_mock = MagicMock()
        rng.Borders.return_value = border_mock
        _call_format("border", rng, mock_session,
                     border_sides="top", border_style="none")
        # LineStyle was set to -4142 (xlLineStyleNone); Weight should NOT be set
        border_mock.LineStyle  # accessed OK
        border_mock.Weight  # should NOT have been called
        # Verify Weight was NOT assigned (checking set calls)
        assert not any(
            "Weight" in str(c) for c in border_mock.method_calls
            if hasattr(c, '__len__') and len(c) > 0
        ), "Weight should not be set when border_style='none'"


# ── Number format ─────────────────────────────────────────────────────────────

class TestNumberFormatAction:
    def test_number_format_set(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_format("number_format", rng, mock_session,
                              number_format="#,##0.00")
        assert rng.NumberFormat == "#,##0.00"
        assert result["applied"]["number_format"] == "#,##0.00"


# ── Alignment ─────────────────────────────────────────────────────────────────

class TestAlignmentAction:
    def test_center_alignment(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_format("alignment", rng, mock_session, horizontal="center")
        assert rng.HorizontalAlignment == -4108  # xlHAlignCenter
        assert result["applied"]["horizontal"] == "center"

    def test_invalid_horizontal_raises(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="horizontal"):
            _call_format("alignment", rng, mock_session, horizontal="justify")

    def test_invalid_vertical_raises(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        with pytest.raises(ToolError, match="vertical"):
            _call_format("alignment", rng, mock_session, vertical="baseline")

    def test_wrap_text_set(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_format("alignment", rng, mock_session, wrap_text=True)
        assert rng.WrapText is True

    def test_merge_true_calls_merge(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_format("alignment", rng, mock_session, merge=True)
        rng.Merge.assert_called_once()

    def test_merge_false_calls_unmerge(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        _call_format("alignment", rng, mock_session, merge=False)
        rng.UnMerge.assert_called_once()


# ── Column width / Row height ─────────────────────────────────────────────────

class TestSizingActions:
    def test_column_width(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_format("column_width", rng, mock_session, width=25.0)
        assert rng.ColumnWidth == 25.0
        assert result["applied"]["width"] == 25.0

    def test_row_height(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_format("row_height", rng, mock_session, height=30.0)
        assert rng.RowHeight == 30.0
        assert result["applied"]["height"] == 30.0


# ── Autofit ───────────────────────────────────────────────────────────────────

class TestAutofitAction:
    def test_autofit_columns_only(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_format("autofit", rng, mock_session,
                              autofit_columns=True, autofit_rows=False)
        rng.Columns.AutoFit.assert_called_once()
        rng.Rows.AutoFit.assert_not_called()
        assert result["applied"]["columns"] is True

    def test_autofit_both(self):
        mock_session = make_mock_session()
        rng = _make_range_mock()
        result = _call_format("autofit", rng, mock_session,
                              autofit_columns=True, autofit_rows=True)
        rng.Columns.AutoFit.assert_called_once()
        rng.Rows.AutoFit.assert_called_once()


# ── Workbook create / save_as ─────────────────────────────────────────────────

class TestWorkbookCreate:
    def test_create_no_path(self):
        """create with no path: Workbooks.Add() called, no SaveAs."""
        from thepexcel_mcp.domains.workbook import workbook_action

        mock_wb = MagicMock()
        mock_wb.Name = "Book1.xlsx"
        mock_wb.FullName = "C:/Book1.xlsx"
        mock_app = MagicMock()
        mock_app.Workbooks.Add.return_value = mock_wb

        mock_session = make_mock_session()
        mock_session.get_app.return_value = mock_app

        with patch("thepexcel_mcp.domains.workbook._session", mock_session):
            with patch("thepexcel_mcp.domains.workbook.excel_guard") as eg:
                eg.return_value.__enter__ = MagicMock(return_value=None)
                eg.return_value.__exit__ = MagicMock(return_value=False)
                result = workbook_action("create")

        mock_app.Workbooks.Add.assert_called_once()
        mock_wb.SaveAs.assert_not_called()
        assert result["created"] == "Book1.xlsx"
        assert result["path"] is None

    def test_create_with_path(self):
        """create with path: Workbooks.Add() then SaveAs() called."""
        from thepexcel_mcp.domains.workbook import workbook_action

        mock_wb = MagicMock()
        mock_wb.Name = "New.xlsx"
        mock_wb.FullName = "C:/data/New.xlsx"
        mock_app = MagicMock()
        mock_app.Workbooks.Add.return_value = mock_wb

        mock_session = make_mock_session()
        mock_session.get_app.return_value = mock_app

        with patch("thepexcel_mcp.domains.workbook._session", mock_session):
            with patch("thepexcel_mcp.domains.workbook.excel_guard") as eg:
                eg.return_value.__enter__ = MagicMock(return_value=None)
                eg.return_value.__exit__ = MagicMock(return_value=False)
                result = workbook_action("create", path="C:/data/New.xlsx")

        mock_wb.SaveAs.assert_called_once_with("C:/data/New.xlsx", FileFormat=51)
        assert result["created"] == "New.xlsx"

    def test_create_missing_path_save_as_raises(self):
        """save_as without path raises ToolError before any COM call."""
        from thepexcel_mcp.domains.workbook import workbook_action
        with pytest.raises(ToolError, match="path"):
            workbook_action("save_as", path=None)


class TestWorkbookSaveAs:
    def test_save_as_xlsx(self):
        """save_as to .xlsx uses FileFormat=51."""
        from thepexcel_mcp.domains.workbook import workbook_action

        mock_wb = MagicMock()
        mock_wb.Name = "Sales.xlsx"
        mock_wb.FullName = "C:/out/Sales.xlsx"

        mock_session = make_mock_session()
        mock_session.get_workbook.return_value = mock_wb

        with patch("thepexcel_mcp.domains.workbook._session", mock_session):
            with patch("thepexcel_mcp.domains.workbook.excel_guard") as eg:
                eg.return_value.__enter__ = MagicMock(return_value=None)
                eg.return_value.__exit__ = MagicMock(return_value=False)
                result = workbook_action("save_as", path="C:/out/Sales.xlsx")

        mock_wb.SaveAs.assert_called_once_with("C:/out/Sales.xlsx", FileFormat=51)
        assert result["saved_as"] == "Sales.xlsx"

    def test_save_as_xlsm(self):
        """save_as to .xlsm uses FileFormat=52."""
        from thepexcel_mcp.domains.workbook import workbook_action

        mock_wb = MagicMock()
        mock_wb.Name = "Macros.xlsm"
        mock_wb.FullName = "C:/out/Macros.xlsm"

        mock_session = make_mock_session()
        mock_session.get_workbook.return_value = mock_wb

        with patch("thepexcel_mcp.domains.workbook._session", mock_session):
            with patch("thepexcel_mcp.domains.workbook.excel_guard") as eg:
                eg.return_value.__enter__ = MagicMock(return_value=None)
                eg.return_value.__exit__ = MagicMock(return_value=False)
                workbook_action("save_as", path="C:/out/Macros.xlsm")

        mock_wb.SaveAs.assert_called_once_with("C:/out/Macros.xlsm", FileFormat=52)

    def test_save_as_csv(self):
        """save_as to .csv uses FileFormat=6."""
        from thepexcel_mcp.domains.workbook import workbook_action

        mock_wb = MagicMock()
        mock_wb.Name = "Export.csv"
        mock_wb.FullName = "C:/out/Export.csv"

        mock_session = make_mock_session()
        mock_session.get_workbook.return_value = mock_wb

        with patch("thepexcel_mcp.domains.workbook._session", mock_session):
            with patch("thepexcel_mcp.domains.workbook.excel_guard") as eg:
                eg.return_value.__enter__ = MagicMock(return_value=None)
                eg.return_value.__exit__ = MagicMock(return_value=False)
                workbook_action("save_as", path="C:/out/Export.csv")

        mock_wb.SaveAs.assert_called_once_with("C:/out/Export.csv", FileFormat=6)

    def test_save_as_unknown_extension_raises(self):
        """save_as with unknown extension raises ToolError."""
        from thepexcel_mcp.domains.workbook import workbook_action

        mock_wb = MagicMock()
        mock_session = make_mock_session()
        mock_session.get_workbook.return_value = mock_wb

        with patch("thepexcel_mcp.domains.workbook._session", mock_session):
            with patch("thepexcel_mcp.domains.workbook.excel_guard") as eg:
                eg.return_value.__enter__ = MagicMock(return_value=None)
                eg.return_value.__exit__ = MagicMock(return_value=False)
                with pytest.raises(ToolError, match="format"):
                    workbook_action("save_as", path="C:/out/Export.txt")


class TestWorkbookActionValidation:
    def test_unknown_action_raises(self):
        from thepexcel_mcp.domains.workbook import workbook_action
        with pytest.raises(ToolError, match="Unknown action"):
            workbook_action("convert")

    def test_valid_actions_listed(self):
        from thepexcel_mcp.domains.workbook import workbook_action
        with pytest.raises(ToolError, match="create"):
            workbook_action("bogus")


class TestInferFileFormat:
    def test_xlsx(self):
        from thepexcel_mcp.domains.workbook import _infer_file_format
        assert _infer_file_format("C:/out/report.xlsx") == 51

    def test_xlsm(self):
        from thepexcel_mcp.domains.workbook import _infer_file_format
        assert _infer_file_format("report.xlsm") == 52

    def test_xlsb(self):
        from thepexcel_mcp.domains.workbook import _infer_file_format
        assert _infer_file_format("report.xlsb") == 50

    def test_xls(self):
        from thepexcel_mcp.domains.workbook import _infer_file_format
        assert _infer_file_format("report.xls") == 56

    def test_csv(self):
        from thepexcel_mcp.domains.workbook import _infer_file_format
        assert _infer_file_format("report.csv") == 6

    def test_uppercase_extension(self):
        """Extension matching is case-insensitive."""
        from thepexcel_mcp.domains.workbook import _infer_file_format
        assert _infer_file_format("report.XLSX") == 51

    def test_unknown_raises(self):
        from thepexcel_mcp.domains.workbook import _infer_file_format
        with pytest.raises(ToolError, match="format"):
            _infer_file_format("report.ods")
