"""Unit tests for excel_hyperlink (hyperlinks.py).

No Excel required — all COM calls are intercepted via make_mock_session.

VERIFY-EFFECT discipline: every mutation test asserts the SPECIFIC COM method
was called with SPECIFIC arguments (ws.Hyperlinks.Add) and that the result dict
echoes the read-back values from the returned Hyperlink object.
"""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock, patch, call

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


# ── Helpers ────────────────────────────────────────────────────────────────────

def _make_sheet_mock(name: str = "Sheet1") -> MagicMock:
    """Minimal Worksheet mock for hyperlinks.py tests."""
    ws = MagicMock()
    ws.Name = name
    return ws


def _make_hl_mock(address: str = "https://example.com",
                  sub_address: str = "",
                  text: str = "Click here",
                  screen_tip: str = "") -> MagicMock:
    """Mock Hyperlink COM object returned by Hyperlinks.Add and iteration."""
    hl = MagicMock()
    hl.Address = address
    hl.SubAddress = sub_address
    hl.TextToDisplay = text
    hl.ScreenTip = screen_tip
    hl.Range.Address = "$A$1"
    return hl


def _call_hyperlink(action: str, mock_session: MagicMock, mock_ws: MagicMock,
                    **kwargs) -> dict:
    """Patch _session + sheet resolution + excel_guard, then call hyperlink_action."""
    mock_session.get_sheet.return_value = mock_ws

    with patch("thepexcel_mcp.domains.hyperlinks._session", mock_session):
        with patch("thepexcel_mcp.domains.hyperlinks.excel_guard") as eg:
            eg.return_value.__enter__ = MagicMock(return_value=None)
            eg.return_value.__exit__ = MagicMock(return_value=False)
            from thepexcel_mcp.domains.hyperlinks import hyperlink_action
            return hyperlink_action(action, sheet=None, workbook=None, **kwargs)


# ── Action validation ──────────────────────────────────────────────────────────

class TestActionValidation:
    def test_unknown_action_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="Unknown action"):
            _call_hyperlink("bogus", ms, ws)

    def test_add_missing_cell_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="cell"):
            _call_hyperlink("add", ms, ws, link_type="url", target="https://x.com")

    def test_add_missing_link_type_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="link_type"):
            _call_hyperlink("add", ms, ws, cell="A1", target="https://x.com")

    def test_add_unknown_link_type_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="Unknown link_type"):
            _call_hyperlink("add", ms, ws, cell="A1", link_type="ftp", target="ftp://x")

    def test_add_missing_target_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="target"):
            _call_hyperlink("add", ms, ws, cell="A1", link_type="url")

    def test_delete_missing_range_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="range"):
            _call_hyperlink("delete", ms, ws)

    def test_add_invalid_cell_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.side_effect = Exception("Bad reference")
        with pytest.raises(ToolError, match="Invalid cell reference"):
            _call_hyperlink("add", ms, ws, cell="ZZZZZ999999",
                            link_type="url", target="https://x.com")


# ── add: URL hyperlinks ────────────────────────────────────────────────────────

class TestAddUrl:
    def test_add_url_calls_hyperlinks_add(self):
        """Hyperlinks.Add must be called with Anchor=Range obj, Address=url."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        mock_anchor = MagicMock()
        ws.Range.return_value = mock_anchor
        mock_hl = _make_hl_mock(address="https://example.com", sub_address="")
        ws.Hyperlinks.Add.return_value = mock_hl

        result = _call_hyperlink("add", ms, ws, cell="A1",
                                 link_type="url", target="https://example.com")

        # VERIFY EFFECT: Hyperlinks.Add was called with the right anchor + address
        ws.Hyperlinks.Add.assert_called_once()
        call_kwargs = ws.Hyperlinks.Add.call_args.kwargs
        assert call_kwargs["Anchor"] is mock_anchor
        assert call_kwargs["Address"] == "https://example.com"
        assert call_kwargs["SubAddress"] == ""

    def test_add_url_result_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("Data")
        ws.Range.return_value = MagicMock()
        ws.Hyperlinks.Add.return_value = _make_hl_mock(address="https://example.com")

        result = _call_hyperlink("add", ms, ws, cell="B2",
                                 link_type="url", target="https://example.com")

        assert result["hyperlink"] == "add"
        assert result["sheet"] == "Data"
        assert "applied" in result
        # cell must be the canonical COM readback (hl.Range.Address), not the caller string
        assert result["applied"]["cell"] == "$A$1"  # _make_hl_mock default
        assert result["applied"]["link_type"] == "url"
        assert result["applied"]["address"] == "https://example.com"

    def test_add_url_with_text_and_tip(self):
        """ScreenTip and TextToDisplay must be forwarded to Hyperlinks.Add."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.return_value = MagicMock()
        mock_hl = _make_hl_mock(address="https://x.com", text="Visit Us",
                                screen_tip="Go to site")
        ws.Hyperlinks.Add.return_value = mock_hl

        _call_hyperlink("add", ms, ws, cell="C3", link_type="url",
                        target="https://x.com", text_to_display="Visit Us",
                        screen_tip="Go to site")

        call_kwargs = ws.Hyperlinks.Add.call_args.kwargs
        assert call_kwargs["TextToDisplay"] == "Visit Us"
        assert call_kwargs["ScreenTip"] == "Go to site"

    def test_add_url_result_includes_screen_tip_readback(self):
        """applied dict must include screen_tip read back from the Hyperlink object."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.return_value = MagicMock()
        mock_hl = _make_hl_mock(address="https://x.com", screen_tip="My Tip")
        ws.Hyperlinks.Add.return_value = mock_hl

        result = _call_hyperlink("add", ms, ws, cell="A1",
                                 link_type="url", target="https://x.com",
                                 screen_tip="My Tip")

        assert "screen_tip" in result["applied"]
        assert result["applied"]["screen_tip"] == "My Tip"

    def test_add_com_exception_raises_tool_error(self):
        """COM error from Hyperlinks.Add must surface as ToolError (via _session.wrap)."""
        ms = make_mock_session()
        ms.wrap.side_effect = lambda e, ctx: ToolError(f"{ctx}: {e}")
        ws = _make_sheet_mock()
        ws.Range.return_value = MagicMock()
        ws.Hyperlinks.Add.side_effect = Exception("COM error -2147024809")

        with pytest.raises(ToolError, match="Hyperlinks.Add failed"):
            _call_hyperlink("add", ms, ws, cell="A1",
                            link_type="url", target="https://x.com")

    def test_add_url_verify_effect_uses_readback(self):
        """Result must echo hl.Range.Address + hl.Address + hl.SubAddress +
        hl.TextToDisplay + hl.ScreenTip (all read-back from COM object)."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.return_value = MagicMock()
        mock_hl = MagicMock()
        mock_hl.Range.Address = "$A$1"
        mock_hl.Address = "https://readback.com"
        mock_hl.SubAddress = ""
        mock_hl.TextToDisplay = "Readback Text"
        mock_hl.ScreenTip = "Readback Tip"
        ws.Hyperlinks.Add.return_value = mock_hl

        result = _call_hyperlink("add", ms, ws, cell="A1",
                                 link_type="url", target="https://original.com",
                                 screen_tip="Original Tip")

        # Result must contain what Excel ACTUALLY stored (readback), not what we sent
        assert result["applied"]["cell"] == "$A$1"
        assert result["applied"]["address"] == "https://readback.com"
        assert result["applied"]["text_to_display"] == "Readback Text"
        assert result["applied"]["screen_tip"] == "Readback Tip"


# ── add: internal hyperlinks ───────────────────────────────────────────────────

class TestAddInternal:
    def test_internal_uses_empty_address(self):
        """Internal links MUST pass Address='' (not None) to Hyperlinks.Add."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.return_value = MagicMock()
        mock_hl = _make_hl_mock(address="", sub_address="Sheet2!A1")
        ws.Hyperlinks.Add.return_value = mock_hl

        _call_hyperlink("add", ms, ws, cell="A1",
                        link_type="internal", target="Sheet2!A1")

        call_kwargs = ws.Hyperlinks.Add.call_args.kwargs
        assert call_kwargs["Address"] == ""  # must be empty string, not None
        assert call_kwargs["SubAddress"] == "Sheet2!A1"

    def test_internal_target_becomes_subaddress(self):
        """For internal links, target is placed in SubAddress."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.return_value = MagicMock()
        mock_hl = _make_hl_mock(address="", sub_address="'My Sheet'!B5")
        ws.Hyperlinks.Add.return_value = mock_hl

        result = _call_hyperlink("add", ms, ws, cell="D4",
                                 link_type="internal", target="'My Sheet'!B5")

        call_kwargs = ws.Hyperlinks.Add.call_args.kwargs
        assert call_kwargs["SubAddress"] == "'My Sheet'!B5"

    def test_internal_explicit_sub_address_overrides(self):
        """If sub_address is supplied explicitly alongside internal type, use it."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.return_value = MagicMock()
        mock_hl = _make_hl_mock(address="", sub_address="Sheet3!C3")
        ws.Hyperlinks.Add.return_value = mock_hl

        _call_hyperlink("add", ms, ws, cell="A1", link_type="internal",
                        target="Sheet2!A1", sub_address="Sheet3!C3")

        call_kwargs = ws.Hyperlinks.Add.call_args.kwargs
        assert call_kwargs["SubAddress"] == "Sheet3!C3"


# ── add: email hyperlinks ──────────────────────────────────────────────────────

class TestAddEmail:
    def test_email_adds_mailto_prefix(self):
        """If target lacks 'mailto:', it must be prepended."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.return_value = MagicMock()
        mock_hl = _make_hl_mock(address="mailto:user@example.com")
        ws.Hyperlinks.Add.return_value = mock_hl

        _call_hyperlink("add", ms, ws, cell="A1",
                        link_type="email", target="user@example.com")

        call_kwargs = ws.Hyperlinks.Add.call_args.kwargs
        assert call_kwargs["Address"] == "mailto:user@example.com"

    def test_email_already_has_mailto_not_doubled(self):
        """If target already has 'mailto:', do not prepend again."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.return_value = MagicMock()
        mock_hl = _make_hl_mock(address="mailto:user@example.com")
        ws.Hyperlinks.Add.return_value = mock_hl

        _call_hyperlink("add", ms, ws, cell="A1",
                        link_type="email", target="mailto:user@example.com")

        call_kwargs = ws.Hyperlinks.Add.call_args.kwargs
        assert call_kwargs["Address"] == "mailto:user@example.com"
        assert "mailto:mailto:" not in call_kwargs["Address"]


# ── add: file hyperlinks ───────────────────────────────────────────────────────

class TestAddFile:
    def test_file_passes_address_as_path(self):
        """File links use the path directly as Address."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.return_value = MagicMock()
        mock_hl = _make_hl_mock(address=r"C:\reports\Q1.xlsx")
        ws.Hyperlinks.Add.return_value = mock_hl

        _call_hyperlink("add", ms, ws, cell="A1",
                        link_type="file", target=r"C:\reports\Q1.xlsx")

        call_kwargs = ws.Hyperlinks.Add.call_args.kwargs
        assert call_kwargs["Address"] == r"C:\reports\Q1.xlsx"


# ── add: optional args omitted ─────────────────────────────────────────────────

class TestAddOptionalArgs:
    def test_screen_tip_omitted_when_none(self):
        """ScreenTip must NOT be in kwargs when screen_tip is None."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.return_value = MagicMock()
        ws.Hyperlinks.Add.return_value = _make_hl_mock()

        _call_hyperlink("add", ms, ws, cell="A1",
                        link_type="url", target="https://x.com")

        call_kwargs = ws.Hyperlinks.Add.call_args.kwargs
        assert "ScreenTip" not in call_kwargs

    def test_text_to_display_omitted_when_none(self):
        """TextToDisplay must NOT be in kwargs when text_to_display is None."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.return_value = MagicMock()
        ws.Hyperlinks.Add.return_value = _make_hl_mock()

        _call_hyperlink("add", ms, ws, cell="A1",
                        link_type="url", target="https://x.com")

        call_kwargs = ws.Hyperlinks.Add.call_args.kwargs
        assert "TextToDisplay" not in call_kwargs


# ── list ───────────────────────────────────────────────────────────────────────

class TestList:
    def _make_hyperlinks_mock(self, ws: MagicMock, hls: list[MagicMock]) -> None:
        """Wire ws.Hyperlinks so .Count and index-call return the given mocks."""
        ws.Hyperlinks.Count = len(hls)
        # ws.Hyperlinks(i) is a COM call-style access
        def hl_item(i):
            return hls[i - 1]
        ws.Hyperlinks.side_effect = hl_item

    def test_list_empty_sheet(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("Empty")
        ws.Hyperlinks.Count = 0

        result = _call_hyperlink("list", ms, ws)

        assert result["hyperlink"] == "list"
        assert result["sheet"] == "Empty"
        assert result["count"] == 0
        assert result["hyperlinks"] == []

    def test_list_single_hyperlink(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("Sheet1")
        hl1 = _make_hl_mock(address="https://example.com", sub_address="",
                             text="Example", screen_tip="Go")
        hl1.Range.Address = "$A$1"
        self._make_hyperlinks_mock(ws, [hl1])

        result = _call_hyperlink("list", ms, ws)

        assert result["count"] == 1
        assert len(result["hyperlinks"]) == 1
        item = result["hyperlinks"][0]
        assert item["anchor"] == "$A$1"
        assert item["address"] == "https://example.com"
        assert item["sub_address"] == ""
        assert item["text"] == "Example"
        assert item["screen_tip"] == "Go"

    def test_list_multiple_hyperlinks(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        hl1 = _make_hl_mock(address="https://a.com")
        hl1.Range.Address = "$A$1"
        hl2 = _make_hl_mock(address="", sub_address="Sheet2!B2", text="Internal")
        hl2.Range.Address = "$B$2"
        self._make_hyperlinks_mock(ws, [hl1, hl2])

        result = _call_hyperlink("list", ms, ws)

        assert result["count"] == 2
        assert result["hyperlinks"][0]["address"] == "https://a.com"
        assert result["hyperlinks"][1]["sub_address"] == "Sheet2!B2"

    def test_list_result_keys(self):
        """Every list entry must have anchor, address, sub_address, text, screen_tip."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        hl = _make_hl_mock()
        hl.Range.Address = "$C$5"
        self._make_hyperlinks_mock(ws, [hl])

        result = _call_hyperlink("list", ms, ws)

        item = result["hyperlinks"][0]
        for key in ("anchor", "address", "sub_address", "text", "screen_tip"):
            assert key in item, f"Missing key '{key}' in hyperlink item"


# ── delete ─────────────────────────────────────────────────────────────────────

class TestDelete:
    def test_delete_calls_hyperlinks_delete(self):
        """Hyperlinks.Delete() must be called on the target range."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        mock_rng = MagicMock()
        mock_rng.Hyperlinks.Count = 2
        mock_rng.Address = "$A$1:$A$5"
        ws.Range.return_value = mock_rng

        result = _call_hyperlink("delete", ms, ws, range="A1:A5")

        # VERIFY EFFECT: Hyperlinks.Delete was called
        mock_rng.Hyperlinks.Delete.assert_called_once()

    def test_delete_reports_count_removed(self):
        """Result must include the count of hyperlinks removed (read BEFORE delete)."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        mock_rng = MagicMock()
        mock_rng.Hyperlinks.Count = 3
        mock_rng.Address = "$B$1:$B$10"
        ws.Range.return_value = mock_rng

        result = _call_hyperlink("delete", ms, ws, range="B1:B10")

        assert result["applied"]["removed"] == 3

    def test_delete_result_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("Data")
        mock_rng = MagicMock()
        mock_rng.Hyperlinks.Count = 0
        mock_rng.Address = "$A$1"
        ws.Range.return_value = mock_rng

        result = _call_hyperlink("delete", ms, ws, range="A1")

        assert result["hyperlink"] == "delete"
        assert result["sheet"] == "Data"
        assert "applied" in result
        assert "range" in result["applied"]
        assert "removed" in result["applied"]

    def test_delete_invalid_range_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ms.wrap.side_effect = lambda e, ctx: ToolError(f"{ctx}: {e}")
        ws.Range.side_effect = Exception("Bad range")

        with pytest.raises(ToolError):
            _call_hyperlink("delete", ms, ws, range="ZZINVALID")

    def test_delete_zero_removed_is_valid(self):
        """Deleting from a range with 0 hyperlinks must succeed (not error)."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        mock_rng = MagicMock()
        mock_rng.Hyperlinks.Count = 0
        mock_rng.Address = "$C$1"
        ws.Range.return_value = mock_rng

        result = _call_hyperlink("delete", ms, ws, range="C1")

        assert result["applied"]["removed"] == 0
        mock_rng.Hyperlinks.Delete.assert_called_once()

    def test_delete_sheet_qualified_range(self):
        """Sheet-qualified range ('Sheet2!A1:B3') must resolve to the named sheet."""
        ms = make_mock_session()
        ws = _make_sheet_mock("Sheet1")
        # ws.Parent.Name = workbook name (used to look up the cross-sheet)
        ws.Parent.Name = "Book1.xlsx"
        # The cross-sheet ws returned by get_sheet
        target_ws = _make_sheet_mock("Sheet2")
        mock_rng = MagicMock()
        mock_rng.Hyperlinks.Count = 1
        mock_rng.Address = "$A$1:$B$3"
        target_ws.Range.return_value = mock_rng
        # get_sheet("Sheet2", "Book1.xlsx") returns target_ws
        ms.get_sheet.side_effect = lambda name, wb: (
            ws if name is None else target_ws
        )

        result = _call_hyperlink("delete", ms, ws, range="Sheet2!A1:B3")

        assert result["sheet"] == "Sheet2"
        assert result["applied"]["removed"] == 1
        mock_rng.Hyperlinks.Delete.assert_called_once()


# ── Result shape contract ──────────────────────────────────────────────────────

class TestResultShape:
    """All result dicts must carry 'hyperlink' + 'sheet' + 'applied'/'hyperlinks'."""

    def test_add_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("MySheet")
        ws.Range.return_value = MagicMock()
        ws.Hyperlinks.Add.return_value = _make_hl_mock()

        result = _call_hyperlink("add", ms, ws, cell="A1",
                                 link_type="url", target="https://x.com")

        assert "hyperlink" in result
        assert "sheet" in result
        assert "applied" in result
        assert result["sheet"] == "MySheet"

    def test_list_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("Reports")
        ws.Hyperlinks.Count = 0

        result = _call_hyperlink("list", ms, ws)

        assert result["hyperlink"] == "list"
        assert "sheet" in result
        assert "count" in result
        assert "hyperlinks" in result

    def test_delete_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("Sales")
        mock_rng = MagicMock()
        mock_rng.Hyperlinks.Count = 1
        mock_rng.Address = "$D$4"
        ws.Range.return_value = mock_rng

        result = _call_hyperlink("delete", ms, ws, range="D4")

        assert result["hyperlink"] == "delete"
        assert result["sheet"] == "Sales"
        assert isinstance(result["applied"], dict)
