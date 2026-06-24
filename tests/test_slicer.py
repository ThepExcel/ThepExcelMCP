"""Unit tests for slicer.py — excel_slicer domain.

No Excel required — all COM calls intercepted via make_mock_session.
The conftest.run_com passthrough executes domain functions synchronously.

Key verification discipline (VERIFY-EFFECT):
  - Assert the SPECIFIC COM method was called with the SPECIFIC value.
  - Do NOT just assert "no exception raised".
"""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock, patch, call

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


# ── Helpers ────────────────────────────────────────────────────────────────────

def _make_wb_mock(
    pivottable_names: list[str] | None = None,
    listobject_names: list[str] | None = None,
    slicer_cache_names: list[str] | None = None,
):
    """Build a minimal Workbook mock with sheets containing PivotTables/Tables."""
    wb = MagicMock()
    wb.Name = "Book1.xlsx"
    wb.Application = MagicMock()

    # One worksheet that holds the requested pivots and tables
    ws = MagicMock()
    ws.Name = "Sheet1"

    # PivotTables mock — call-style access: ws.PivotTables() returns collection
    pt_objects = []
    for pt_name in (pivottable_names or []):
        pt = MagicMock()
        pt.Name = pt_name
        pt_objects.append(pt)

    pt_collection = MagicMock()
    pt_collection.Count = len(pt_objects)

    def _pt_getitem(idx):
        if isinstance(idx, int):
            return pt_objects[idx - 1]
        for p in pt_objects:
            if p.Name == idx:
                return p
        raise Exception(f"PivotTable '{idx}' not found")

    # ws.PivotTables() is called with no args to get the collection
    # AND with an index to get a specific pivot — handle both
    ws.PivotTables.side_effect = lambda *args: (
        pt_collection if not args else _pt_getitem(args[0])
    )
    pt_collection.__call__ = ws.PivotTables.side_effect

    # ListObjects mock — attribute access: ws.ListObjects is a collection
    lo_objects = []
    for lo_name in (listobject_names or []):
        lo = MagicMock()
        lo.Name = lo_name
        lo_objects.append(lo)

    lo_collection = MagicMock()
    lo_collection.Count = len(lo_objects)

    def _lo_getitem(idx):
        if isinstance(idx, int):
            return lo_objects[idx - 1]
        raise Exception(f"ListObject '{idx}' not found")

    lo_collection.side_effect = _lo_getitem
    # ws.ListObjects(j) returns lo_objects[j-1]
    ws.ListObjects = lo_collection

    # Workbook sheets
    wb.Sheets.Count = 1
    wb.Sheets.side_effect = lambda idx: ws

    # SlicerCaches
    sc_objects = []
    for scn in (slicer_cache_names or []):
        sc = MagicMock()
        sc.Name = scn
        sc.SourceName = f"source_of_{scn}"
        sc.SlicerCacheType = 1  # xlSlicer
        sc_objects.append(sc)

    sc_collection = MagicMock()
    sc_collection.Count = len(sc_objects)

    def _sc_getitem(idx):
        if isinstance(idx, int):
            return sc_objects[idx - 1]
        for s in sc_objects:
            if s.Name == idx:
                return s
        raise Exception(f"SlicerCache '{idx}' not found")

    sc_collection.side_effect = _sc_getitem

    # Add2 returns a stable SlicerCache mock (same object every time — tests can
    # inspect Slicers.Add calls on it via sc_collection.Add2.return_value)
    new_sc = MagicMock()
    new_sc.Name = "Slicer_MockField"
    new_sc.PivotTables = MagicMock()
    new_slicer = MagicMock()
    new_slicer.Name = "mock_slicer"
    new_sc.Slicers.Add.return_value = new_slicer
    sc_collection.Add2.return_value = new_sc

    wb.SlicerCaches = sc_collection
    return wb, ws, sc_objects, lo_objects, pt_objects


def _call_slicer(action, mock_session, wb_mock, ws_mock=None, **kwargs):
    """Patch _session + workbook + sheet, call slicer_action."""
    with patch("thepexcel_mcp.domains.slicer._session", mock_session):
        with patch("thepexcel_mcp.domains.slicer.excel_guard") as eg:
            eg.return_value.__enter__ = MagicMock(return_value=None)
            eg.return_value.__exit__ = MagicMock(return_value=False)
            mock_session.get_workbook.return_value = wb_mock
            if ws_mock is not None:
                mock_session.get_sheet.return_value = ws_mock
            from thepexcel_mcp.domains.slicer import slicer_action
            return slicer_action(action, **kwargs)


# ── Validation ────────────────────────────────────────────────────────────────

class TestSlicerActionValidation:
    def test_unknown_action_raises(self):
        mock_session = make_mock_session()
        wb, ws, _, _, _ = _make_wb_mock()
        with pytest.raises(ToolError, match="Unknown action"):
            _call_slicer("bogus", mock_session, wb, ws)

    def test_add_requires_source(self):
        mock_session = make_mock_session()
        wb, ws, _, _, _ = _make_wb_mock(listobject_names=["SalesT"])
        with pytest.raises(ToolError, match="source"):
            _call_slicer("add", mock_session, wb, ws, field="Region")

    def test_add_requires_field(self):
        mock_session = make_mock_session()
        wb, ws, _, _, _ = _make_wb_mock(listobject_names=["SalesT"])
        with pytest.raises(ToolError, match="field"):
            _call_slicer("add", mock_session, wb, ws, source="SalesT")

    def test_delete_requires_slicer_name(self):
        mock_session = make_mock_session()
        wb, ws, _, _, _ = _make_wb_mock()
        with pytest.raises(ToolError, match="slicer"):
            _call_slicer("delete", mock_session, wb, ws)

    def test_connect_requires_slicer_name(self):
        mock_session = make_mock_session()
        wb, ws, _, _, _ = _make_wb_mock(slicer_cache_names=["Slicer_Region"])
        with pytest.raises(ToolError, match="slicer"):
            _call_slicer("connect", mock_session, wb, ws, pivots=["PT1"])

    def test_connect_requires_pivots(self):
        mock_session = make_mock_session()
        wb, ws, _, _, _ = _make_wb_mock(slicer_cache_names=["Slicer_Region"])
        with pytest.raises(ToolError, match="pivots"):
            _call_slicer("connect", mock_session, wb, ws, slicer="Slicer_Region")


# ── Source resolver ────────────────────────────────────────────────────────────

class TestResolveSource:
    def test_resolves_listobject(self):
        from thepexcel_mcp.domains.slicer import _resolve_source
        wb, ws, _, lo_objects, _ = _make_wb_mock(listobject_names=["SalesT"])
        result = _resolve_source(wb, "SalesT")
        assert result.Name == "SalesT"

    def test_resolves_pivottable(self):
        from thepexcel_mcp.domains.slicer import _resolve_source
        wb, ws, _, _, pt_objects = _make_wb_mock(pivottable_names=["PivotTable1"])
        result = _resolve_source(wb, "PivotTable1")
        assert result.Name == "PivotTable1"

    def test_not_found_raises(self):
        from thepexcel_mcp.domains.slicer import _resolve_source
        wb, ws, _, _, _ = _make_wb_mock()
        with pytest.raises(ToolError, match="not found"):
            _resolve_source(wb, "NonExistent")


# ── Add (table slicer) ────────────────────────────────────────────────────────

class TestAddSlicer:
    def test_add_calls_add2_with_listobject(self):
        """add on a Table source must pass the ListObject object to Add2."""
        mock_session = make_mock_session()
        wb, ws, _, lo_objects, _ = _make_wb_mock(listobject_names=["SalesT"])
        result = _call_slicer(
            "add", mock_session, wb, ws,
            source="SalesT", field="Region",
        )
        # Add2 must have been called on SlicerCaches
        wb.SlicerCaches.Add2.assert_called_once()
        call_args = wb.SlicerCaches.Add2.call_args
        # First positional arg must be the ListObject (resolved object, not string)
        src_arg = call_args[0][0]
        assert src_arg.Name == "SalesT", (
            "Add2 must receive the actual ListObject, not a string"
        )

    def test_add_passes_xlslicer_cache_type(self):
        """Regular add must pass xlSlicer=1 as cache type."""
        mock_session = make_mock_session()
        wb, ws, _, lo_objects, _ = _make_wb_mock(listobject_names=["SalesT"])
        _call_slicer("add", mock_session, wb, ws, source="SalesT", field="Region")
        call_args = wb.SlicerCaches.Add2.call_args
        # cache_type is positional[3] when a name is given, else the
        # SlicerCacheType keyword (no-name path omits the Name slot)
        cache_type = call_args.kwargs.get(
            "SlicerCacheType",
            call_args.args[3] if len(call_args.args) > 3 else None,
        )
        assert cache_type == 1, f"Expected xlSlicer=1, got {cache_type}"

    def test_add_omits_empty_name_arg(self):
        """REGRESSION: with no name given, Add2 must OMIT the Name slot.

        Passing Name="" (or None) to SlicerCaches.Add2 raises COM
        E_INVALIDARG (-2147024809) against a live Excel — a defect a mock
        cannot surface, so assert the call SHAPE directly: no empty-string /
        None Name, and SlicerCacheType conveyed as a keyword instead.
        """
        mock_session = make_mock_session()
        wb, ws, _, lo_objects, _ = _make_wb_mock(listobject_names=["SalesT"])
        _call_slicer("add", mock_session, wb, ws, source="SalesT", field="Region")
        call_args = wb.SlicerCaches.Add2.call_args
        positional = call_args.args
        assert "" not in positional, "Add2 must not receive an empty-string Name"
        assert len(positional) < 3 or positional[2] not in ("", None), (
            "Add2 Name slot must be omitted (not empty-string/None) when unnamed"
        )
        assert call_args.kwargs.get("SlicerCacheType") == 1, (
            "cache type must be passed as the SlicerCacheType keyword on the no-name path"
        )

    def test_add_calls_slicers_add_on_returned_cache(self):
        """After Add2, Slicers.Add must be called on the returned SlicerCache."""
        mock_session = make_mock_session()
        wb, ws, _, lo_objects, _ = _make_wb_mock(listobject_names=["SalesT"])
        _call_slicer(
            "add", mock_session, wb, ws,
            source="SalesT", field="Region", caption="Region Filter",
        )
        # Get the SlicerCache that Add2 returned
        returned_sc = wb.SlicerCaches.Add2.return_value
        returned_sc.Slicers.Add.assert_called_once()

    def test_add_passes_custom_position(self):
        """Slicers.Add must be called with keyword args (no Level positional slot).

        Level is OLAP-only.  For Table/non-OLAP sources, passing Level as a
        positional arg raises COM E_INVALIDARG (-2147024809).  The fix uses
        keyword args to skip Level entirely.  This test verifies each kwarg slot
        so a regression to positional-with-Level is caught immediately.
        """
        mock_session = make_mock_session()
        wb, ws, _, lo_objects, _ = _make_wb_mock(listobject_names=["SalesT"])
        _call_slicer(
            "add", mock_session, wb, ws,
            source="SalesT", field="Region",
            top=10.0, left=20.0, width=200.0, height=100.0,
        )
        returned_sc = wb.SlicerCaches.Add2.return_value
        add_call = returned_sc.Slicers.Add.call_args
        # Must use keyword args only — no positional args (Level must be absent)
        kwargs = add_call[1]
        assert add_call[0] == (), (
            f"Slicers.Add must have NO positional args (got {add_call[0]}); "
            "Level must be omitted for non-OLAP slicers"
        )
        assert kwargs.get("SlicerDestination") is ws, (
            f"SlicerDestination mismatch: {kwargs.get('SlicerDestination')}"
        )
        # Caption defaults to the field name when not explicitly supplied
        assert kwargs.get("Caption") == "Region", (
            f"Caption expected 'Region', got {kwargs.get('Caption')!r}"
        )
        assert kwargs.get("Top")    == 10.0,  f"Top mismatch: {kwargs.get('Top')}"
        assert kwargs.get("Left")   == 20.0,  f"Left mismatch: {kwargs.get('Left')}"
        assert kwargs.get("Width")  == 200.0, f"Width mismatch: {kwargs.get('Width')}"
        assert kwargs.get("Height") == 100.0, f"Height mismatch: {kwargs.get('Height')}"

    def test_add_result_shape(self):
        """Result dict must contain expected keys."""
        mock_session = make_mock_session()
        wb, ws, _, lo_objects, _ = _make_wb_mock(listobject_names=["SalesT"])
        result = _call_slicer(
            "add", mock_session, wb, ws,
            source="SalesT", field="Region",
        )
        assert result["added"] == "slicer"
        assert "slicer_cache" in result
        assert result["source"] == "SalesT"
        assert result["field"] == "Region"

    def test_add_pivot_source(self):
        """add on a PivotTable source must pass the PivotTable object to Add2."""
        mock_session = make_mock_session()
        wb, ws, _, _, pt_objects = _make_wb_mock(pivottable_names=["PivotTable1"])
        _call_slicer(
            "add", mock_session, wb, ws,
            source="PivotTable1", field="Category",
        )
        call_args = wb.SlicerCaches.Add2.call_args
        src_arg = call_args[0][0]
        assert src_arg.Name == "PivotTable1"


# ── Add timeline ──────────────────────────────────────────────────────────────

class TestAddTimeline:
    def test_add_timeline_passes_xltimeline_cache_type(self):
        """add_timeline must pass xlTimeline=2 as the SlicerCacheType."""
        mock_session = make_mock_session()
        wb, ws, _, _, pt_objects = _make_wb_mock(pivottable_names=["SalesPivot"])
        _call_slicer(
            "add_timeline", mock_session, wb, ws,
            source="SalesPivot", field="OrderDate",
        )
        call_args = wb.SlicerCaches.Add2.call_args
        cache_type = call_args.kwargs.get(
            "SlicerCacheType",
            call_args.args[3] if len(call_args.args) > 3 else None,
        )
        assert cache_type == 2, f"Expected xlTimeline=2, got {cache_type}"

    def test_add_timeline_result_shape(self):
        mock_session = make_mock_session()
        wb, ws, _, _, pt_objects = _make_wb_mock(pivottable_names=["SalesPivot"])
        result = _call_slicer(
            "add_timeline", mock_session, wb, ws,
            source="SalesPivot", field="OrderDate",
        )
        assert result["added"] == "timeline"
        assert result["field"] == "OrderDate"

    def test_add_timeline_requires_source(self):
        mock_session = make_mock_session()
        wb, ws, _, _, _ = _make_wb_mock(pivottable_names=["SalesPivot"])
        with pytest.raises(ToolError, match="source"):
            _call_slicer("add_timeline", mock_session, wb, ws, field="OrderDate")

    def test_add_timeline_requires_field(self):
        mock_session = make_mock_session()
        wb, ws, _, _, _ = _make_wb_mock(pivottable_names=["SalesPivot"])
        with pytest.raises(ToolError, match="field"):
            _call_slicer("add_timeline", mock_session, wb, ws, source="SalesPivot")


# ── List ──────────────────────────────────────────────────────────────────────

class TestListSlicers:
    def test_list_returns_all_caches(self):
        mock_session = make_mock_session()
        wb, ws, sc_objects, _, _ = _make_wb_mock(
            slicer_cache_names=["Slicer_Region", "Slicer_Category"]
        )
        result = _call_slicer("list", mock_session, wb, ws)
        assert result["count"] == 2
        names = [s["name"] for s in result["slicer_caches"]]
        assert "Slicer_Region" in names
        assert "Slicer_Category" in names

    def test_list_empty_workbook(self):
        mock_session = make_mock_session()
        wb, ws, _, _, _ = _make_wb_mock()
        result = _call_slicer("list", mock_session, wb, ws)
        assert result["count"] == 0
        assert result["slicer_caches"] == []

    def test_list_includes_slicer_type(self):
        mock_session = make_mock_session()
        wb, ws, sc_objects, _, _ = _make_wb_mock(slicer_cache_names=["Slicer_Date"])
        # Make it look like a timeline
        sc_objects[0].SlicerCacheType = 2  # xlTimeline
        result = _call_slicer("list", mock_session, wb, ws)
        assert result["slicer_caches"][0]["slicer_type"] == "timeline"

    def test_list_slicer_type_str_for_regular(self):
        mock_session = make_mock_session()
        wb, ws, sc_objects, _, _ = _make_wb_mock(slicer_cache_names=["Slicer_Region"])
        sc_objects[0].SlicerCacheType = 1  # xlSlicer
        result = _call_slicer("list", mock_session, wb, ws)
        assert result["slicer_caches"][0]["slicer_type"] == "slicer"


# ── Delete ────────────────────────────────────────────────────────────────────

class TestDeleteSlicer:
    def test_delete_calls_cache_delete(self):
        """SlicerCache.Delete() must be called on the named cache."""
        mock_session = make_mock_session()
        wb, ws, sc_objects, _, _ = _make_wb_mock(slicer_cache_names=["Slicer_Region"])
        result = _call_slicer("delete", mock_session, wb, ws, slicer="Slicer_Region")
        sc_objects[0].Delete.assert_called_once()
        assert result["deleted"] == "Slicer_Region"

    def test_delete_not_found_raises(self):
        mock_session = make_mock_session()
        wb, ws, _, _, _ = _make_wb_mock()
        with pytest.raises(ToolError, match="not found"):
            _call_slicer("delete", mock_session, wb, ws, slicer="Slicer_Missing")


# ── Connect ───────────────────────────────────────────────────────────────────

class TestConnectSlicer:
    def test_connect_calls_addpivottable(self):
        """connect must call SlicerPivotTables.AddPivotTable with the resolved pivot."""
        mock_session = make_mock_session()
        wb, ws, sc_objects, _, pt_objects = _make_wb_mock(
            slicer_cache_names=["Slicer_Region"],
            pivottable_names=["PivotTable1"],
        )
        result = _call_slicer(
            "connect", mock_session, wb, ws,
            slicer="Slicer_Region",
            pivots=["PivotTable1"],
        )
        # AddPivotTable must have been called on the cache's PivotTables collection
        sc_objects[0].PivotTables.AddPivotTable.assert_called_once()
        call_arg = sc_objects[0].PivotTables.AddPivotTable.call_args[0][0]
        assert call_arg.Name == "PivotTable1"
        assert "PivotTable1" in result["connected"]

    def test_connect_result_shape(self):
        mock_session = make_mock_session()
        wb, ws, sc_objects, _, pt_objects = _make_wb_mock(
            slicer_cache_names=["Slicer_Region"],
            pivottable_names=["PT1", "PT2"],
        )
        result = _call_slicer(
            "connect", mock_session, wb, ws,
            slicer="Slicer_Region",
            pivots=["PT1", "PT2"],
        )
        assert set(result["connected"]) == {"PT1", "PT2"}
        assert result["slicer_cache"] == "Slicer_Region"
        assert result["errors"] == []

    def test_connect_bad_pivot_recorded_as_error(self):
        """A pivot name that doesn't exist should be recorded in errors, not crash."""
        mock_session = make_mock_session()
        wb, ws, sc_objects, _, pt_objects = _make_wb_mock(
            slicer_cache_names=["Slicer_Region"],
            pivottable_names=["PT1"],
        )
        result = _call_slicer(
            "connect", mock_session, wb, ws,
            slicer="Slicer_Region",
            pivots=["PT1", "NonExistentPT"],
        )
        assert "PT1" in result["connected"]
        assert any(e["pivot"] == "NonExistentPT" for e in result["errors"])
