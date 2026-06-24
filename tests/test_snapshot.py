"""Unit tests for excel_snapshot (snapshot.py).

No Excel required — all COM calls are intercepted via make_mock_session, and the
module's `os` filesystem calls are patched per test.

SAFETY-CRITICAL discipline: these tests verify that the snapshot path is
non-destructive — `snapshot` uses SaveCopyAs (never SaveAs/Save/Close),
`restore` uses Workbooks.Open (never close/overwrite of the live workbook).
VERIFY-EFFECT discipline: every test asserts the SPECIFIC COM call / registry
state, not just "no exception raised".
"""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock, patch

from fastmcp.exceptions import ToolError
from conftest import make_mock_session

import thepexcel_mcp.domains.snapshot as snap_mod


# ── Fixtures ────────────────────────────────────────────────────────────────────

@pytest.fixture(autouse=True)
def _reset_registry():
    """Clear the module-level snapshot registry + counter before/after each test
    so registry state never leaks across tests."""
    snap_mod._SNAPSHOTS.clear()
    snap_mod._counter = 0
    yield
    snap_mod._SNAPSHOTS.clear()
    snap_mod._counter = 0


# ── Helpers ────────────────────────────────────────────────────────────────────

def _make_wb_mock(name: str = "Sales.xlsx"):
    """Minimal Workbook mock sufficient for snapshot.py tests."""
    wb = MagicMock()
    wb.Name = name
    wb.Application = MagicMock()
    return wb


def _patch_guard():
    """Return a patch context for excel_guard that is a no-op context manager."""
    eg = patch("thepexcel_mcp.domains.snapshot.excel_guard")
    return eg


def _call(action, mock_session, **kwargs):
    """Invoke snapshot_action with _session + excel_guard patched."""
    with patch("thepexcel_mcp.domains.snapshot._session", mock_session):
        with _patch_guard() as eg:
            eg.return_value.__enter__ = MagicMock(return_value=None)
            eg.return_value.__exit__ = MagicMock(return_value=False)
            return snap_mod.snapshot_action(action, **kwargs)


# ── Action validation ──────────────────────────────────────────────────────────

class TestActionValidation:
    def test_unknown_action_raises(self):
        ms = make_mock_session()
        with pytest.raises(ToolError, match="Unknown action"):
            _call("bogus", ms)

    def test_restore_requires_snapshot_id(self):
        ms = make_mock_session()
        with pytest.raises(ToolError, match="snapshot_id"):
            _call("restore", ms)

    def test_delete_requires_snapshot_id(self):
        ms = make_mock_session()
        with pytest.raises(ToolError, match="snapshot_id"):
            _call("delete", ms)


# ── snapshot ────────────────────────────────────────────────────────────────────

class TestSnapshot:
    def test_savecopyas_called_with_computed_path(self):
        ms = make_mock_session()
        wb = _make_wb_mock("Sales.xlsx")
        ms.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.snapshot.os.makedirs"), \
             patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=True), \
             patch("thepexcel_mcp.domains.snapshot.os.path.getsize", return_value=4096):
            result = _call("snapshot", ms)

        # VERIFY EFFECT: SaveCopyAs was called once with the computed path,
        # and that path matches the one returned.
        wb.SaveCopyAs.assert_called_once()
        called_path = wb.SaveCopyAs.call_args[0][0]
        assert called_path == result["path"]
        # path must preserve the .xlsx extension and contain the id
        assert result["path"].endswith(".xlsx")
        assert result["id"] in result["path"]

    def test_snapshot_never_calls_destructive_save_or_close(self):
        """SAFETY: snapshot must use SaveCopyAs ONLY — never Save/SaveAs/Close."""
        ms = make_mock_session()
        wb = _make_wb_mock("Sales.xlsx")
        ms.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.snapshot.os.makedirs"), \
             patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=True), \
             patch("thepexcel_mcp.domains.snapshot.os.path.getsize", return_value=4096):
            _call("snapshot", ms)

        wb.Save.assert_not_called()
        wb.SaveAs.assert_not_called()
        wb.Close.assert_not_called()

    def test_snapshot_preserves_xlsm_extension(self):
        """Macro-enabled workbook must keep .xlsm so macros survive the copy."""
        ms = make_mock_session()
        wb = _make_wb_mock("Macros.xlsm")
        ms.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.snapshot.os.makedirs"), \
             patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=True), \
             patch("thepexcel_mcp.domains.snapshot.os.path.getsize", return_value=1024):
            result = _call("snapshot", ms)

        assert result["path"].endswith(".xlsm")

    def test_snapshot_defaults_unsaved_workbook_to_xlsx(self):
        """A never-saved 'Book1' (no extension) defaults to .xlsx."""
        ms = make_mock_session()
        wb = _make_wb_mock("Book1")
        ms.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.snapshot.os.makedirs"), \
             patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=True), \
             patch("thepexcel_mcp.domains.snapshot.os.path.getsize", return_value=512):
            result = _call("snapshot", ms)

        assert result["path"].endswith(".xlsx")

    def test_snapshot_unknown_extension_falls_back_xlsx(self):
        ms = make_mock_session()
        wb = _make_wb_mock("data.weird")
        ms.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.snapshot.os.makedirs"), \
             patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=True), \
             patch("thepexcel_mcp.domains.snapshot.os.path.getsize", return_value=512):
            result = _call("snapshot", ms)

        assert result["path"].endswith(".xlsx")

    def test_verify_effect_raises_when_file_missing(self):
        """If SaveCopyAs 'succeeds' but the file is not on disk → ToolError."""
        ms = make_mock_session()
        wb = _make_wb_mock("Sales.xlsx")
        ms.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.snapshot.os.makedirs"), \
             patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=False):
            with pytest.raises(ToolError, match="missing or empty"):
                _call("snapshot", ms)

    def test_verify_effect_raises_when_file_empty(self):
        ms = make_mock_session()
        wb = _make_wb_mock("Sales.xlsx")
        ms.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.snapshot.os.makedirs"), \
             patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=True), \
             patch("thepexcel_mcp.domains.snapshot.os.path.getsize", return_value=0):
            with pytest.raises(ToolError, match="missing or empty"):
                _call("snapshot", ms)

    def test_snapshot_registers_in_registry(self):
        ms = make_mock_session()
        wb = _make_wb_mock("Sales.xlsx")
        ms.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.snapshot.os.makedirs"), \
             patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=True), \
             patch("thepexcel_mcp.domains.snapshot.os.path.getsize", return_value=2048):
            result = _call("snapshot", ms)

        # registry round-trips id → path
        sid = result["id"]
        assert sid in snap_mod._SNAPSHOTS
        assert snap_mod._SNAPSHOTS[sid]["path"] == result["path"]
        assert snap_mod._SNAPSHOTS[sid]["workbook"] == "Sales.xlsx"

    def test_snapshot_result_shape(self):
        ms = make_mock_session()
        wb = _make_wb_mock("Sales.xlsx")
        ms.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.snapshot.os.makedirs"), \
             patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=True), \
             patch("thepexcel_mcp.domains.snapshot.os.path.getsize", return_value=4096):
            result = _call("snapshot", ms)

        for key in ("id", "path", "workbook", "size_bytes", "created"):
            assert key in result
        assert result["size_bytes"] == 4096

    def test_two_snapshots_get_distinct_ids(self):
        ms = make_mock_session()
        wb = _make_wb_mock("Sales.xlsx")
        ms.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.snapshot.os.makedirs"), \
             patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=True), \
             patch("thepexcel_mcp.domains.snapshot.os.path.getsize", return_value=4096):
            r1 = _call("snapshot", ms)
            r2 = _call("snapshot", ms)

        assert r1["id"] != r2["id"]
        assert len(snap_mod._SNAPSHOTS) == 2


# ── list ────────────────────────────────────────────────────────────────────────

class TestList:
    def test_list_empty(self):
        ms = make_mock_session()
        result = _call("list", ms)
        assert result["snapshots"] == []
        assert result["count"] == 0

    def test_list_reports_existing_snapshot(self):
        ms = make_mock_session()
        wb = _make_wb_mock("Sales.xlsx")
        ms.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.snapshot.os.makedirs"), \
             patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=True), \
             patch("thepexcel_mcp.domains.snapshot.os.path.getsize", return_value=4096):
            snap = _call("snapshot", ms)
            # list also probes os.path.exists / getsize → still patched True/4096
            result = _call("list", ms)

        assert result["count"] == 1
        entry = result["snapshots"][0]
        assert entry["id"] == snap["id"]
        assert entry["exists"] is True
        assert entry["size_bytes"] == 4096
        assert entry["workbook"] == "Sales.xlsx"

    def test_list_reports_missing_file_as_not_exists(self):
        """A registered snapshot whose file was deleted externally → exists False."""
        ms = make_mock_session()
        snap_mod._SNAPSHOTS["snap_1_Foo"] = {
            "path": "C:/tmp/Foo__snap_1_Foo.xlsx",
            "workbook": "Foo.xlsx",
            "created": "2026-06-24T00:00:00",
            "format": ".xlsx",
        }
        with patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=False):
            result = _call("list", ms)

        entry = result["snapshots"][0]
        assert entry["exists"] is False
        assert entry["size_bytes"] == 0


# ── restore ──────────────────────────────────────────────────────────────────────

class TestRestore:
    def _register(self, path="C:/tmp/Sales__snap_1_Sales.xlsx"):
        snap_mod._SNAPSHOTS["snap_1_Sales"] = {
            "path": path,
            "workbook": "Sales.xlsx",
            "created": "2026-06-24T00:00:00",
            "format": ".xlsx",
        }

    def test_restore_opens_registered_path_as_new_workbook(self):
        ms = make_mock_session()
        self._register()

        opened = MagicMock()
        opened.Name = "Sales__snap_1_Sales.xlsx"

        app = MagicMock()
        app.Workbooks.Open.return_value = opened
        # Workbooks set: index access returns the opened wb (present check)
        app.Workbooks.Count = 1
        app.Workbooks.side_effect = lambda i: opened
        ms.get_app.return_value = app

        with patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=True):
            result = _call("restore", ms, snapshot_id="snap_1_Sales")

        # VERIFY EFFECT: Workbooks.Open called with the registered path
        app.Workbooks.Open.assert_called_once_with(
            "C:/tmp/Sales__snap_1_Sales.xlsx"
        )
        assert result["restored_from"] == "snap_1_Sales"
        assert result["opened_workbook"] == "Sales__snap_1_Sales.xlsx"
        assert "not modified" in result["note"]

    def test_restore_never_closes_or_overwrites_any_workbook(self):
        """SAFETY: restore must only Open — never Close/SaveAs on the app or wb."""
        ms = make_mock_session()
        self._register()

        opened = MagicMock()
        opened.Name = "Sales__snap_1_Sales.xlsx"

        app = MagicMock()
        app.Workbooks.Open.return_value = opened
        app.Workbooks.Count = 1
        app.Workbooks.side_effect = lambda i: opened
        ms.get_app.return_value = app

        with patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=True):
            _call("restore", ms, snapshot_id="snap_1_Sales")

        opened.Close.assert_not_called()
        opened.SaveAs.assert_not_called()
        app.ActiveWorkbook.Close.assert_not_called()

    def test_restore_unknown_id_raises(self):
        ms = make_mock_session()
        with pytest.raises(ToolError, match="unknown snapshot_id"):
            _call("restore", ms, snapshot_id="nope")

    def test_restore_missing_file_raises(self):
        ms = make_mock_session()
        self._register()
        with patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=False):
            with pytest.raises(ToolError, match="missing on disk"):
                _call("restore", ms, snapshot_id="snap_1_Sales")

    def test_restore_verify_effect_raises_when_not_present(self):
        """Open 'succeeds' but the opened name isn't in Workbooks → ToolError."""
        ms = make_mock_session()
        self._register()

        opened = MagicMock()
        opened.Name = "Sales__snap_1_Sales.xlsx"

        other = MagicMock()
        other.Name = "Different.xlsx"

        app = MagicMock()
        app.Workbooks.Open.return_value = opened
        app.Workbooks.Count = 1
        app.Workbooks.side_effect = lambda i: other  # opened name NOT present
        ms.get_app.return_value = app

        with patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=True):
            with pytest.raises(ToolError, match="not present"):
                _call("restore", ms, snapshot_id="snap_1_Sales")


# ── delete ───────────────────────────────────────────────────────────────────────

class TestDelete:
    def _register(self, path="C:/tmp/Sales__snap_1_Sales.xlsx"):
        snap_mod._SNAPSHOTS["snap_1_Sales"] = {
            "path": path,
            "workbook": "Sales.xlsx",
            "created": "2026-06-24T00:00:00",
            "format": ".xlsx",
        }

    def test_delete_removes_file_and_clears_registry(self):
        ms = make_mock_session()
        self._register()

        # exists True before remove (to trigger os.remove), False after (verify)
        exists_seq = iter([True, False])
        with patch("thepexcel_mcp.domains.snapshot.os.path.exists",
                   side_effect=lambda p: next(exists_seq)), \
             patch("thepexcel_mcp.domains.snapshot.os.remove") as rm:
            result = _call("delete", ms, snapshot_id="snap_1_Sales")

        rm.assert_called_once_with("C:/tmp/Sales__snap_1_Sales.xlsx")
        assert result["deleted"] == "snap_1_Sales"
        assert "snap_1_Sales" not in snap_mod._SNAPSHOTS

    def test_delete_unknown_id_raises(self):
        ms = make_mock_session()
        with pytest.raises(ToolError, match="unknown snapshot_id"):
            _call("delete", ms, snapshot_id="nope")

    def test_delete_file_already_gone_still_clears_registry(self):
        """If the file is already gone, delete must still succeed + clear entry."""
        ms = make_mock_session()
        self._register()

        with patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=False), \
             patch("thepexcel_mcp.domains.snapshot.os.remove") as rm:
            result = _call("delete", ms, snapshot_id="snap_1_Sales")

        rm.assert_not_called()  # nothing to remove
        assert result["deleted"] == "snap_1_Sales"
        assert "snap_1_Sales" not in snap_mod._SNAPSHOTS

    def test_delete_verify_effect_raises_if_file_persists(self):
        """os.remove 'succeeds' but file still exists → ToolError."""
        ms = make_mock_session()
        self._register()

        # exists True before AND after remove → verify-effect must fire
        with patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=True), \
             patch("thepexcel_mcp.domains.snapshot.os.remove"):
            with pytest.raises(ToolError, match="still exists"):
                _call("delete", ms, snapshot_id="snap_1_Sales")


# ── Registry round-trip ──────────────────────────────────────────────────────────

class TestRegistryRoundTrip:
    def test_snapshot_then_list_then_delete(self):
        ms = make_mock_session()
        wb = _make_wb_mock("Sales.xlsx")
        ms.get_workbook.return_value = wb

        with patch("thepexcel_mcp.domains.snapshot.os.makedirs"), \
             patch("thepexcel_mcp.domains.snapshot.os.path.exists", return_value=True), \
             patch("thepexcel_mcp.domains.snapshot.os.path.getsize", return_value=4096):
            snap = _call("snapshot", ms)
            sid = snap["id"]
            listed = _call("list", ms)
            assert any(e["id"] == sid for e in listed["snapshots"])

        # delete it
        exists_seq = iter([True, False])
        with patch("thepexcel_mcp.domains.snapshot.os.path.exists",
                   side_effect=lambda p: next(exists_seq)), \
             patch("thepexcel_mcp.domains.snapshot.os.remove"):
            _call("delete", ms, snapshot_id=sid)

        assert sid not in snap_mod._SNAPSHOTS
        # list now empty
        empty = _call("list", ms)
        assert empty["count"] == 0
