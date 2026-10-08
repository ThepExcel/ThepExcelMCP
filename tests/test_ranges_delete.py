"""Row/column deletion dispatch, targeting, ordering and partial-failure tests."""

from unittest.mock import MagicMock, patch

import pytest
from fastmcp.exceptions import ToolError

from conftest import make_mock_session
from thepexcel_mcp.domains import ranges


def _selection(axis, intervals, fail_at=None):
    ws = MagicMock()
    ws.Name = "Data"
    ws.Parent.Name = "Test.xlsx"
    ws.Parent.Application.DisplayAlerts = True
    ws.Cells.side_effect = lambda row, col: (row, col)
    areas = []
    for start, end in intervals:
        area = MagicMock()
        area.Row = area.Column = start
        area.Rows.Count = area.Columns.Count = end - start + 1
        areas.append(area)
    rng = MagicMock()
    rng.Parent = ws
    rng.Areas.Count = len(areas)
    rng.Areas.Item.side_effect = lambda index: areas[index - 1]
    calls, targets = [], []

    def make_target(first, last):
        start, end = (first[0], last[0]) if axis == "rows" else (first[1], last[1])
        target = MagicMock()
        target.Address = f"original:{start}:{end}"

        def delete():
            assert ws.Parent.Application.DisplayAlerts is False
            calls.append((start, end))
            if start == fail_at:
                raise RuntimeError("sheet is protected")

        target.Delete.side_effect = delete
        block = MagicMock()
        block.EntireRow = block.EntireColumn = target
        targets.append(target)
        return block

    ws.Range.side_effect = make_target
    return rng, ws, calls, targets


@pytest.mark.parametrize("axis", ["rows", "columns"])
@pytest.mark.parametrize("intervals,expected", [
    ([(3, 3)], [(3, 3)]),
    ([(2, 4)], [(2, 4)]),
    ([(2, 4), (9, 9)], [(9, 9), (2, 4)]),
    ([(9, 9), (2, 4), (3, 6), (9, 9)], [(9, 9), (2, 6)]),
    ([(2, 4), (5, 6)], [(2, 6)]),
])
def test_delete_original_intervals_once_in_descending_order(axis, intervals, expected):
    rng, ws, calls, targets = _selection(axis, intervals)
    session = make_mock_session()
    with patch.object(ranges, "_session", session), patch.object(
        ranges, "_resolve_range", return_value=rng,
    ) as resolve:
        result = ranges.range_action(
            "delete_" + axis, "selection", sheet="Data", workbook="Test.xlsx",
        )["deleted"]
    resolve.assert_called_once_with("selection", "Data", "Test.xlsx")
    session.run_com.assert_called_once()
    assert calls == expected
    assert result == {
        "workbook": "Test.xlsx", "sheet": "Data", "axis": axis,
        "intervals": [
            {"start": start, "end": end, "range": f"original:{start}:{end}"}
            for start, end in reversed(expected)
        ],
        "count": sum(end - start + 1 for start, end in expected),
    }
    assert ws.Parent.Application.DisplayAlerts is True
    for target in targets:
        target.Delete.assert_called_once_with()
    ws.Parent.Save.assert_not_called()
    ws.Parent.Close.assert_not_called()
    rng.Delete.assert_not_called()
    rng.ClearContents.assert_not_called()


@pytest.mark.parametrize("axis", ["rows", "columns"])
def test_error_reports_completed_intervals_without_retry(axis):
    rng, ws, calls, _ = _selection(axis, [(2, 4), (9, 9)], fail_at=2)
    session = make_mock_session()
    session.wrap.side_effect = lambda exc, context: ToolError(f"{context}: {exc}")
    with patch.object(ranges, "_session", session), patch.object(
        ranges, "_resolve_range", return_value=rng,
    ), pytest.raises(ToolError, match=f"Already deleted 1 {axis}") as caught:
        ranges.range_action("delete_" + axis, "selection")
    assert "original:9:9" in str(caught.value)
    assert "not rolled back" in str(caught.value)
    assert "sheet is protected" in str(caught.value)
    assert calls == [(9, 9), (2, 4)]
    assert ws.Parent.Application.DisplayAlerts is True


def test_all_areas_are_resolved_before_any_deletion():
    rng, ws, calls, _ = _selection("rows", [(2, 4), (9, 9)])
    rng.Areas.Item.side_effect = [rng.Areas.Item(1), RuntimeError("invalid area")]
    session = make_mock_session()
    session.wrap.side_effect = lambda exc, context: ToolError(f"{context}: {exc}")
    with patch.object(ranges, "_session", session), patch.object(
        ranges, "_resolve_range", return_value=rng,
    ), pytest.raises(ToolError, match="Already deleted 0 rows"):
        ranges.range_action("delete_rows", "selection")
    assert calls == []
    ws.Range.assert_not_called()


@pytest.mark.parametrize("address", ["", "   "])
@pytest.mark.parametrize("action", ["delete_rows", "delete_columns"])
def test_empty_range_fails_before_com(action, address):
    session = make_mock_session()
    with patch.object(ranges, "_session", session), pytest.raises(
        ToolError, match="non-empty range",
    ):
        ranges.range_action(action, address)
    session.run_com.assert_not_called()


@pytest.mark.parametrize("action", ["delete_rows", "delete_columns"])
def test_invalid_range_never_deletes(action):
    session = make_mock_session()
    with patch.object(ranges, "_session", session), patch.object(
        ranges, "_resolve_range", side_effect=ToolError("Invalid range"),
    ), pytest.raises(ToolError, match="Invalid range"):
        ranges.range_action(action, "not a range")
    session.get_app.assert_not_called()


@pytest.mark.parametrize("address,sheet,expected_sheet,cell_part", [
    ("2:4,9:9", None, None, "2:4,9:9"),
    ("B:D,G:G", "Data", "Data", "B:D,G:G"),
    ("'Other Sheet'!A2:B4,C9", "Data", "Other Sheet", "A2:B4,C9"),
])
def test_delete_preserves_sheet_qualifier_and_workbook_target(
    address, sheet, expected_sheet, cell_part,
):
    rng, ws, _, _ = _selection("rows", [(2, 4)])
    session = make_mock_session()
    session.get_sheet.return_value = ws
    original_range = ws.Range.side_effect
    ws.Range.side_effect = lambda *args: rng if len(args) == 1 else original_range(*args)
    with patch.object(ranges, "_session", session):
        ranges.range_action("delete_rows", address, sheet=sheet, workbook="Test.xlsx")
    session.get_sheet.assert_called_once_with(expected_sheet, "Test.xlsx")
    ws.Range.assert_any_call(cell_part.split(",")[0])


@pytest.mark.parametrize("address,expected", [
    ("2:4,9:9", ["2:4", "9:9"]),
    ("'Data, Other'!A1,B2", ["'Data, Other'!A1", "B2"]),
    ("Table1[[#Data],[Amount]],G1", ["Table1[[#Data],[Amount]]", "G1"]),
])
def test_split_selection_is_independent_of_excel_locale(address, expected):
    assert ranges._split_delete_areas(address) == expected


def test_invalid_later_area_is_resolved_before_mutation():
    rng, ws, calls, _ = _selection("rows", [(2, 4)])
    ws.Range.side_effect = RuntimeError("Invalid later area")
    session = make_mock_session()
    session.wrap.side_effect = lambda exc, context: ToolError(f"{context}: {exc}")
    with patch.object(ranges, "_session", session), patch.object(
        ranges, "_resolve_range", return_value=rng,
    ), pytest.raises(ToolError, match="Already deleted 0 rows"):
        ranges.range_action("delete_rows", "2:4,A0")
    assert calls == []


@pytest.mark.parametrize("address", ["2:4,", ",2:4", "2:4,,9:9"])
def test_empty_union_component_is_rejected(address):
    with pytest.raises(ToolError, match="empty area"):
        ranges._split_delete_areas(address)


@pytest.mark.parametrize("action", ["delete_rows", "delete_columns"])
def test_raw_value_mode_is_still_read_only(action):
    session = make_mock_session()
    with patch.object(ranges, "_session", session), pytest.raises(
        ToolError, match="only valid for read",
    ):
        ranges.range_action(action, "A1", value_mode="raw")
    session.run_com.assert_not_called()
