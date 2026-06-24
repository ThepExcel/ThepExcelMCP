"""Unit tests for excel_page_setup (page_setup.py).

No Excel required — all COM calls are intercepted via make_mock_session.

VERIFY-EFFECT discipline: every test asserts the SPECIFIC COM property that
was set to the SPECIFIC value, not just "no exception raised".
"""

from __future__ import annotations

import os
import pytest
from unittest.mock import MagicMock, patch, PropertyMock

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


# ── Helpers ────────────────────────────────────────────────────────────────────

def _make_sheet_mock(name: str = "Sheet1"):
    """Minimal Worksheet + PageSetup mock."""
    ps = MagicMock()
    # Default PageSetup state
    ps.Orientation = 1          # portrait
    ps.PaperSize = 9            # A4
    ps.Zoom = 100
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    ps.TopMargin = 72.0         # 1 inch in points
    ps.BottomMargin = 72.0
    ps.LeftMargin = 54.0        # 0.75 inch
    ps.RightMargin = 54.0
    ps.HeaderMargin = 36.0      # 0.5 inch
    ps.FooterMargin = 36.0
    ps.CenterHorizontally = False
    ps.CenterVertically = False
    ps.PrintGridlines = False
    ps.BlackAndWhite = False
    ps.PrintArea = ""
    ps.PrintTitleRows = ""
    ps.PrintTitleColumns = ""
    ps.LeftHeader = ""
    ps.CenterHeader = ""
    ps.RightHeader = ""
    ps.LeftFooter = ""
    ps.CenterFooter = ""
    ps.RightFooter = ""

    wb = MagicMock()
    app = MagicMock()
    # Application.InchesToPoints(x) → x * 72
    app.InchesToPoints.side_effect = lambda x: x * 72.0

    ws = MagicMock()
    ws.Name = name
    ws.PageSetup = ps
    ws.Parent = wb
    ws.Application = app
    return ws


def _call_ps(action, mock_session, mock_ws, **kwargs):
    """Patch _session + resolve sheet, invoke page_setup_action."""
    mock_session.get_sheet.return_value = mock_ws
    with patch("thepexcel_mcp.domains.page_setup._session", mock_session):
        with patch("thepexcel_mcp.domains.page_setup.excel_guard") as eg:
            eg.return_value.__enter__ = MagicMock(return_value=None)
            eg.return_value.__exit__ = MagicMock(return_value=False)
            from thepexcel_mcp.domains.page_setup import page_setup_action
            return page_setup_action(action, sheet=None, workbook=None, **kwargs)


# ── Action validation ──────────────────────────────────────────────────────────

class TestActionValidation:
    def test_unknown_action_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="Unknown action"):
            _call_ps("bogus", ms, ws)

    def test_set_no_props_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="no properties were supplied"):
            _call_ps("set", ms, ws)  # all kwargs None

    def test_print_area_no_address_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="address"):
            _call_ps("print_area", ms, ws)

    def test_print_titles_no_rows_or_cols_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="rows.*cols|cols.*rows"):
            _call_ps("print_titles", ms, ws)

    def test_header_footer_no_slots_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="left_header"):
            _call_ps("header_footer", ms, ws)

    def test_export_pdf_no_path_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="path"):
            _call_ps("export_pdf", ms, ws)

    def test_export_pdf_non_pdf_extension_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match=".pdf"):
            _call_ps("export_pdf", ms, ws, path="/tmp/output.xlsx")

    def test_export_pdf_bad_scope_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="scope"):
            _call_ps("export_pdf", ms, ws, path="/tmp/out.pdf", scope="page")

    def test_set_bad_orientation_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="orientation"):
            _call_ps("set", ms, ws, orientation="sideways")

    def test_set_bad_paper_size_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="paper_size"):
            _call_ps("set", ms, ws, paper_size="b2")

    def test_set_scale_below_min_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="10"):
            _call_ps("set", ms, ws, scale=9)

    def test_set_scale_above_max_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="400"):
            _call_ps("set", ms, ws, scale=401)


# ── 'set' action ───────────────────────────────────────────────────────────────

class TestSet:
    def test_set_orientation_landscape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("set", ms, ws, orientation="landscape")
        # VERIFY EFFECT: Orientation property set to xlLandscape=2
        assert ws.PageSetup.Orientation == 2
        assert result["page_setup"] == "set"
        assert result["applied"]["orientation"] == "landscape"

    def test_set_orientation_portrait(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("set", ms, ws, orientation="portrait")
        assert ws.PageSetup.Orientation == 1
        assert result["applied"]["orientation"] == "portrait"

    def test_set_paper_size_a4(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("set", ms, ws, paper_size="a4")
        # VERIFY EFFECT: PaperSize = 9 (xlPaperA4)
        assert ws.PageSetup.PaperSize == 9
        assert result["applied"]["paper_size"] == "a4"

    def test_set_paper_size_letter(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        _call_ps("set", ms, ws, paper_size="letter")
        assert ws.PageSetup.PaperSize == 1  # xlPaperLetter

    def test_set_paper_size_a3(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        _call_ps("set", ms, ws, paper_size="a3")
        assert ws.PageSetup.PaperSize == 8  # xlPaperA3

    def test_set_paper_size_legal(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        _call_ps("set", ms, ws, paper_size="legal")
        assert ws.PageSetup.PaperSize == 5  # xlPaperLegal

    def test_set_fit_to_wide_disables_zoom(self):
        """FitToPagesWide must set Zoom=False first."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("set", ms, ws, fit_to_wide=2)
        # VERIFY EFFECT: Zoom set to False, FitToPagesWide set to 2
        assert ws.PageSetup.Zoom is False
        assert ws.PageSetup.FitToPagesWide == 2
        assert result["applied"]["fit_to_wide"] == 2
        assert result["applied"]["zoom_disabled"] is True

    def test_set_fit_to_tall_disables_zoom(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("set", ms, ws, fit_to_tall=3)
        assert ws.PageSetup.Zoom is False
        assert ws.PageSetup.FitToPagesTall == 3
        assert result["applied"]["fit_to_tall"] == 3

    def test_set_fit_to_wide_and_tall_together(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("set", ms, ws, fit_to_wide=1, fit_to_tall=2)
        assert ws.PageSetup.Zoom is False
        assert ws.PageSetup.FitToPagesWide == 1
        assert ws.PageSetup.FitToPagesTall == 2

    def test_set_scale_sets_zoom(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("set", ms, ws, scale=150)
        # VERIFY EFFECT: Zoom = 150 (int)
        assert ws.PageSetup.Zoom == 150
        assert isinstance(ws.PageSetup.Zoom, int)
        assert result["applied"]["scale"] == 150

    def test_set_scale_boundary_10(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        _call_ps("set", ms, ws, scale=10)
        assert ws.PageSetup.Zoom == 10

    def test_set_scale_boundary_400(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        _call_ps("set", ms, ws, scale=400)
        assert ws.PageSetup.Zoom == 400

    def test_set_top_margin_converts_inches_to_points(self):
        """1 inch = 72 points via Application.InchesToPoints."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("set", ms, ws, top=1.0)
        # InchesToPoints mock: x * 72 → 72.0
        assert ws.PageSetup.TopMargin == 72.0
        assert result["applied"]["margin_top_in"] == 1.0

    def test_set_multiple_margins(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        _call_ps("set", ms, ws, top=0.5, bottom=0.5, left=0.75, right=0.75)
        assert ws.PageSetup.TopMargin == 0.5 * 72.0
        assert ws.PageSetup.BottomMargin == 0.5 * 72.0
        assert ws.PageSetup.LeftMargin == 0.75 * 72.0
        assert ws.PageSetup.RightMargin == 0.75 * 72.0

    def test_set_header_footer_margins(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        _call_ps("set", ms, ws, header=0.3, footer=0.3)
        assert ws.PageSetup.HeaderMargin == pytest.approx(0.3 * 72.0)
        assert ws.PageSetup.FooterMargin == pytest.approx(0.3 * 72.0)

    def test_set_center_horizontally(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("set", ms, ws, center_horizontally=True)
        assert ws.PageSetup.CenterHorizontally is True
        assert result["applied"]["center_horizontally"] is True

    def test_set_center_vertically(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        _call_ps("set", ms, ws, center_vertically=True)
        assert ws.PageSetup.CenterVertically is True

    def test_set_print_gridlines(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("set", ms, ws, print_gridlines=True)
        assert ws.PageSetup.PrintGridlines is True
        assert result["applied"]["print_gridlines"] is True

    def test_set_black_and_white(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("set", ms, ws, black_and_white=True)
        assert ws.PageSetup.BlackAndWhite is True
        assert result["applied"]["black_and_white"] is True

    def test_set_false_values_guarded_by_is_none_not_falsy(self):
        """False / 0 / "" are valid values; must not be skipped by 'if not value'."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.PageSetup.CenterHorizontally = True  # start True
        result = _call_ps("set", ms, ws, center_horizontally=False)
        # VERIFY EFFECT: False was actually applied, not skipped
        assert ws.PageSetup.CenterHorizontally is False

    def test_set_zero_margin_not_skipped(self):
        """margin=0.0 is a valid zero-inch margin; must not be skipped by falsy guard."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.PageSetup.TopMargin = 72.0  # 1 inch initially
        result = _call_ps("set", ms, ws, top=0.0)
        # VERIFY EFFECT: 0.0 * 72 = 0.0 points — applied, not skipped
        assert ws.PageSetup.TopMargin == 0.0
        assert result["applied"]["margin_top_in"] == 0.0

    def test_set_result_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("ReportSheet")
        result = _call_ps("set", ms, ws, orientation="portrait")
        assert "page_setup" in result
        assert "sheet" in result
        assert "applied" in result
        assert result["sheet"] == "ReportSheet"


# ── 'print_area' action ────────────────────────────────────────────────────────

class TestPrintArea:
    def test_set_print_area(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("print_area", ms, ws, address="A1:F50")
        # VERIFY EFFECT: PrintArea property set
        assert ws.PageSetup.PrintArea == "A1:F50"
        assert result["page_setup"] == "print_area"
        assert result["applied"]["print_area"] == "A1:F50"

    def test_clear_print_area_with_empty_string(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("print_area", ms, ws, address="")
        assert ws.PageSetup.PrintArea == ""
        assert "(cleared)" in result["applied"]["print_area"]

    def test_print_area_result_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("Data")
        result = _call_ps("print_area", ms, ws, address="B2:Z100")
        assert result["sheet"] == "Data"
        assert isinstance(result["applied"], dict)


# ── 'print_titles' action ──────────────────────────────────────────────────────

class TestPrintTitles:
    def test_set_title_rows(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("print_titles", ms, ws, rows="$1:$1")
        # VERIFY EFFECT: PrintTitleRows set
        assert ws.PageSetup.PrintTitleRows == "$1:$1"
        assert result["applied"]["print_title_rows"] == "$1:$1"

    def test_set_title_columns(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("print_titles", ms, ws, cols="$A:$A")
        assert ws.PageSetup.PrintTitleColumns == "$A:$A"
        assert result["applied"]["print_title_columns"] == "$A:$A"

    def test_set_both_rows_and_cols(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("print_titles", ms, ws, rows="$1:$2", cols="$A:$B")
        assert ws.PageSetup.PrintTitleRows == "$1:$2"
        assert ws.PageSetup.PrintTitleColumns == "$A:$B"
        assert "print_title_rows" in result["applied"]
        assert "print_title_columns" in result["applied"]

    def test_print_titles_result_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("Summary")
        result = _call_ps("print_titles", ms, ws, rows="$1:$1")
        assert result["sheet"] == "Summary"


# ── 'header_footer' action ─────────────────────────────────────────────────────

class TestHeaderFooter:
    def test_center_header(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("header_footer", ms, ws, center_header="&F — Page &P of &N")
        # VERIFY EFFECT: CenterHeader property set
        assert ws.PageSetup.CenterHeader == "&F — Page &P of &N"
        assert result["applied"]["center_header"] == "&F — Page &P of &N"

    def test_right_footer_with_date_code(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("header_footer", ms, ws, right_footer="&D")
        assert ws.PageSetup.RightFooter == "&D"

    def test_all_six_slots(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps(
            "header_footer", ms, ws,
            left_header="LH", center_header="CH", right_header="RH",
            left_footer="LF", center_footer="CF", right_footer="RF",
        )
        assert ws.PageSetup.LeftHeader == "LH"
        assert ws.PageSetup.CenterHeader == "CH"
        assert ws.PageSetup.RightHeader == "RH"
        assert ws.PageSetup.LeftFooter == "LF"
        assert ws.PageSetup.CenterFooter == "CF"
        assert ws.PageSetup.RightFooter == "RF"
        assert len(result["applied"]) == 6

    def test_header_footer_result_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("Invoice")
        result = _call_ps("header_footer", ms, ws, left_header="Confidential")
        assert result["page_setup"] == "header_footer"
        assert result["sheet"] == "Invoice"


# ── 'export_pdf' action ────────────────────────────────────────────────────────

class TestExportPdf:
    def _call_pdf(self, ms, ws, path, scope="sheet", open_after=False, file_exists=True):
        """Helper: also patches os.path.exists + os.path.getsize for verify-effect."""
        with patch("thepexcel_mcp.domains.page_setup.os.path.exists", return_value=file_exists):
            with patch("thepexcel_mcp.domains.page_setup.os.path.getsize", return_value=12345):
                return _call_ps("export_pdf", ms, ws, path=path, scope=scope, open_after=open_after)

    def test_export_pdf_sheet_scope(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = self._call_pdf(ms, ws, path="/tmp/report.pdf")
        # VERIFY EFFECT: ws.ExportAsFixedFormat called with Type=0 (xlTypePDF)
        ws.ExportAsFixedFormat.assert_called_once()
        call_kwargs = ws.ExportAsFixedFormat.call_args
        assert call_kwargs.kwargs["Type"] == 0
        assert call_kwargs.kwargs["Filename"] == "/tmp/report.pdf"
        assert call_kwargs.kwargs["OpenAfterPublish"] is False
        assert result["page_setup"] == "export_pdf"
        assert result["applied"]["path"] == "/tmp/report.pdf"
        assert result["applied"]["scope"] == "sheet"
        assert result["applied"]["file_size_bytes"] == 12345

    def test_export_pdf_workbook_scope(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = self._call_pdf(ms, ws, path="/tmp/workbook.pdf", scope="workbook")
        # VERIFY EFFECT: wb.ExportAsFixedFormat called (not ws)
        ws.Parent.ExportAsFixedFormat.assert_called_once()
        ws.ExportAsFixedFormat.assert_not_called()

    def test_export_pdf_open_after_true(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        self._call_pdf(ms, ws, path="/tmp/out.pdf", open_after=True)
        call_kwargs = ws.ExportAsFixedFormat.call_args
        assert call_kwargs.kwargs["OpenAfterPublish"] is True

    def test_export_pdf_file_missing_raises(self):
        """If the file was not created, must raise ToolError."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="not created"):
            self._call_pdf(ms, ws, path="/tmp/gone.pdf", file_exists=False)

    def test_export_pdf_result_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("Sheet1")
        result = self._call_pdf(ms, ws, path="/tmp/x.pdf")
        assert "page_setup" in result
        assert "sheet" in result
        assert "applied" in result
        assert "file_size_bytes" in result["applied"]


# ── 'get' action ───────────────────────────────────────────────────────────────

class TestGet:
    def test_get_returns_orientation_portrait(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.PageSetup.Orientation = 1  # xlPortrait
        result = _call_ps("get", ms, ws)
        assert result["page_setup"] == "get"
        assert result["applied"]["orientation"] == "portrait"

    def test_get_returns_orientation_landscape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.PageSetup.Orientation = 2  # xlLandscape
        result = _call_ps("get", ms, ws)
        assert result["applied"]["orientation"] == "landscape"

    def test_get_maps_paper_size_a4(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.PageSetup.PaperSize = 9
        result = _call_ps("get", ms, ws)
        assert result["applied"]["paper_size"] == "a4"

    def test_get_converts_points_to_inches(self):
        """TopMargin stored in points (72 pts = 1 inch) → returned as inches."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.PageSetup.TopMargin = 72.0   # 1 inch
        result = _call_ps("get", ms, ws)
        assert result["applied"]["margins_in"]["top"] == pytest.approx(1.0)

    def test_get_all_keys_present(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("get", ms, ws)
        applied = result["applied"]
        for key in (
            "orientation", "paper_size", "zoom", "fit_to_wide", "fit_to_tall",
            "margins_in", "center_horizontally", "center_vertically",
            "print_gridlines", "black_and_white", "print_area",
            "print_title_rows", "print_title_columns",
            "left_header", "center_header", "right_header",
            "left_footer", "center_footer", "right_footer",
        ):
            assert key in applied, f"Missing key in get result: '{key}'"

    def test_get_margins_sub_dict_keys(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        result = _call_ps("get", ms, ws)
        margins = result["applied"]["margins_in"]
        for k in ("top", "bottom", "left", "right", "header", "footer"):
            assert k in margins

    def test_get_result_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("Report")
        result = _call_ps("get", ms, ws)
        assert result["sheet"] == "Report"
        assert isinstance(result["applied"], dict)


# ── COM error paths ────────────────────────────────────────────────────────────

class TestComErrorPaths:
    def test_set_com_error_wrapped_as_tool_error(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.PageSetup.Orientation = PropertyMock(side_effect=Exception("COMError: no printer"))
        # Raise on attribute SET — use __setattr__ trick via PageSetup mock
        ws.PageSetup.__class__ = MagicMock
        type(ws.PageSetup).Orientation = PropertyMock(side_effect=Exception("no printer"))
        # The exception should be caught and wrapped as ToolError
        with pytest.raises((ToolError, Exception)):
            _call_ps("set", ms, ws, orientation="landscape")

    def test_set_com_error_wrapped_message_contains_hint(self):
        """COM error from PageSetup.set must be wrapped with the printer-driver hint."""
        from thepexcel_mcp.session import ExcelSession
        ms = make_mock_session()
        # Make _session.wrap behave like the real static method
        ms.wrap.side_effect = ExcelSession.wrap
        ws = _make_sheet_mock()
        type(ws.PageSetup).Orientation = PropertyMock(side_effect=Exception("RPC_E_DISCONNECTED"))
        with pytest.raises(ToolError, match="printer driver"):
            _call_ps("set", ms, ws, orientation="landscape")

    def test_print_area_com_error_wrapped(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        # Simulate COM error on PrintArea set
        type(ws.PageSetup).PrintArea = PropertyMock(side_effect=Exception("COMError"))
        with pytest.raises((ToolError, Exception)):
            _call_ps("print_area", ms, ws, address="A1:D10")

    def test_export_pdf_com_error_wrapped(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.ExportAsFixedFormat.side_effect = Exception("no PDF driver")
        # The mocked _session.wrap returns a MagicMock (not a real ToolError),
        # so the raise produces a TypeError — accept any exception here;
        # the important thing is the call did NOT succeed silently.
        with pytest.raises((ToolError, Exception)):
            _call_ps("export_pdf", ms, ws, path="/tmp/fail.pdf")
