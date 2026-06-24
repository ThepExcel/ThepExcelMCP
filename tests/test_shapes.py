"""Unit tests for excel_shape (shapes.py).

No Excel required — all COM calls are intercepted via make_mock_session.

VERIFY-EFFECT discipline: every test asserts the SPECIFIC COM method/property
that was called/set to the SPECIFIC value, not just "no exception raised".
Mocks are initialised to the OPPOSITE of the expected value where possible
so that a passing assertion actually proves the code ran.
"""

from __future__ import annotations

import os
from unittest.mock import MagicMock, patch, call

import pytest

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


# ── Helpers ────────────────────────────────────────────────────────────────────

def _make_ws(name: str = "Sheet1") -> MagicMock:
    """Minimal Worksheet mock for shapes.py tests."""
    ws = MagicMock()
    ws.Name = name
    ws.Application = MagicMock()
    return ws


def _make_shape(shape_name: str = "Picture 1", auto_type: int = -2, shp_type: int = 13) -> MagicMock:
    """A minimal Shape COM mock."""
    shp = MagicMock()
    shp.Name = shape_name
    shp.Type = shp_type
    shp.AutoShapeType = auto_type
    shp.Left = 100.0
    shp.Top  = 50.0
    shp.Width  = 200.0
    shp.Height = 100.0
    return shp


def _call_shape(action, mock_session, mock_ws, **kwargs):
    """Patch _session + get_sheet, then invoke shape_action."""
    mock_session.get_sheet.return_value = mock_ws
    with patch("thepexcel_mcp.domains.shapes._session", mock_session):
        with patch("thepexcel_mcp.domains.shapes.excel_guard") as eg:
            eg.return_value.__enter__ = MagicMock(return_value=None)
            eg.return_value.__exit__ = MagicMock(return_value=False)
            from thepexcel_mcp.domains.shapes import shape_action
            return shape_action(action, sheet=None, workbook=None, **kwargs)


# ── Action validation ──────────────────────────────────────────────────────────

class TestActionValidation:
    def test_unknown_action_raises(self):
        ms = make_mock_session()
        ws = _make_ws()
        with pytest.raises(ToolError, match="Unknown action"):
            _call_shape("bogus", ms, ws)

    def test_add_image_requires_filename(self):
        ms = make_mock_session()
        ws = _make_ws()
        # cell anchor provided so position resolves, but no filename
        ws.Range.return_value = MagicMock(Left=10.0, Top=20.0)
        with pytest.raises(ToolError, match="filename"):
            _call_shape("add_image", ms, ws, cell="A1")

    def test_add_image_nonexistent_file_raises(self):
        ms = make_mock_session()
        ws = _make_ws()
        ws.Range.return_value = MagicMock(Left=0.0, Top=0.0)
        with pytest.raises(ToolError, match="file not found"):
            _call_shape("add_image", ms, ws, cell="A1", filename="/nonexistent/path/img.png")

    def test_add_textbox_requires_width(self):
        ms = make_mock_session()
        ws = _make_ws()
        ws.Shapes = MagicMock()
        ws.Shapes.Count = 0
        ws.Shapes.AddTextbox = MagicMock(return_value=_make_shape())
        with pytest.raises(ToolError, match="width"):
            _call_shape("add_textbox", ms, ws, left=0.0, top=0.0, height=50.0)

    def test_add_textbox_requires_height(self):
        ms = make_mock_session()
        ws = _make_ws()
        with pytest.raises(ToolError, match="height"):
            _call_shape("add_textbox", ms, ws, left=0.0, top=0.0, width=100.0)

    def test_add_shape_requires_shape_type(self):
        ms = make_mock_session()
        ws = _make_ws()
        with pytest.raises(ToolError, match="shape_type"):
            _call_shape("add_shape", ms, ws, left=0.0, top=0.0, width=100.0, height=50.0)

    def test_add_shape_unknown_type_raises(self):
        ms = make_mock_session()
        ws = _make_ws()
        with pytest.raises(ToolError, match="Unknown shape_type"):
            _call_shape("add_shape", ms, ws, left=0.0, top=0.0, width=100.0, height=50.0, shape_type="flying_saucer")

    def test_delete_requires_name(self):
        ms = make_mock_session()
        ws = _make_ws()
        with pytest.raises(ToolError, match="name"):
            _call_shape("delete", ms, ws)

    def test_move_requires_name(self):
        ms = make_mock_session()
        ws = _make_ws()
        with pytest.raises(ToolError, match="name"):
            _call_shape("move", ms, ws, left=0.0, top=0.0)

    def test_missing_position_raises(self):
        """No cell and no left/top should raise."""
        ms = make_mock_session()
        ws = _make_ws()
        with pytest.raises(ToolError, match="cell"):
            _call_shape("add_textbox", ms, ws, width=100.0, height=50.0)


# ── add_image ──────────────────────────────────────────────────────────────────

class TestAddImage:
    def _make_ws_with_shapes(self, count_before: int = 0) -> MagicMock:
        ws = _make_ws()
        shp = _make_shape("Picture 1", shp_type=13)
        ws.Shapes = MagicMock()
        ws.Shapes.Count = count_before
        def _add_picture(**kwargs):
            ws.Shapes.Count = count_before + 1
            return shp
        ws.Shapes.AddPicture = MagicMock(side_effect=_add_picture)
        ws.Range.return_value = MagicMock(Left=10.0, Top=20.0)
        return ws, shp

    def test_calls_addpicture_with_embed_flags(self, tmp_path):
        """AddPicture must be called with LinkToFile=0 (msoFalse), SaveWithDocument=-1 (msoTrue)."""
        img = tmp_path / "test.png"
        img.write_bytes(b"\x89PNG")

        ms = make_mock_session()
        ws, shp = self._make_ws_with_shapes()

        _call_shape("add_image", ms, ws, cell="A1", filename=str(img))

        call_kwargs = ws.Shapes.AddPicture.call_args.kwargs
        assert call_kwargs["LinkToFile"] == 0,   "msoFalse=0 required for embedded image"
        assert call_kwargs["SaveWithDocument"] == -1, "msoTrue=-1 required when linking=False"
        assert call_kwargs["Filename"] == str(img)

    def test_cell_anchor_position(self, tmp_path):
        """Left/Top must come from Range.Left/Range.Top when cell is provided."""
        img = tmp_path / "img.png"
        img.write_bytes(b"\x89PNG")

        ms = make_mock_session()
        ws, shp = self._make_ws_with_shapes()
        ws.Range.return_value = MagicMock(Left=42.0, Top=17.0)

        _call_shape("add_image", ms, ws, cell="C5", filename=str(img))

        call_kwargs = ws.Shapes.AddPicture.call_args.kwargs
        assert call_kwargs["Left"] == 42.0
        assert call_kwargs["Top"] == 17.0

    def test_explicit_position(self, tmp_path):
        img = tmp_path / "img2.png"
        img.write_bytes(b"\x89PNG")

        ms = make_mock_session()
        ws, shp = self._make_ws_with_shapes()

        _call_shape("add_image", ms, ws, left=55.0, top=33.0, filename=str(img))

        kw = ws.Shapes.AddPicture.call_args.kwargs
        assert kw["Left"] == 55.0
        assert kw["Top"] == 33.0

    def test_native_size_passes_minus1(self, tmp_path):
        """Omitting width/height should pass Width=-1, Height=-1 (native size)."""
        img = tmp_path / "img3.png"
        img.write_bytes(b"\x89PNG")

        ms = make_mock_session()
        ws, shp = self._make_ws_with_shapes()

        _call_shape("add_image", ms, ws, left=0.0, top=0.0, filename=str(img))

        kw = ws.Shapes.AddPicture.call_args.kwargs
        assert kw["Width"] == -1.0
        assert kw["Height"] == -1.0

    def test_explicit_size_is_passed(self, tmp_path):
        img = tmp_path / "img4.png"
        img.write_bytes(b"\x89PNG")

        ms = make_mock_session()
        ws, shp = self._make_ws_with_shapes()

        _call_shape("add_image", ms, ws, left=0.0, top=0.0, filename=str(img), width=120.0, height=80.0)

        kw = ws.Shapes.AddPicture.call_args.kwargs
        assert kw["Width"] == 120.0
        assert kw["Height"] == 80.0

    def test_result_shape(self, tmp_path):
        img = tmp_path / "img5.png"
        img.write_bytes(b"\x89PNG")

        ms = make_mock_session()
        ws, shp = self._make_ws_with_shapes()
        # Set shape geometry so we can verify readback (not the requested -1/-1)
        shp.Left   = 0.0
        shp.Top    = 0.0
        shp.Width  = 320.0  # native width Excel would set
        shp.Height = 240.0  # native height Excel would set

        result = _call_shape("add_image", ms, ws, left=0.0, top=0.0, filename=str(img))

        assert result["shape"] == "add_image"
        assert "sheet" in result
        assert "applied" in result
        assert result["applied"]["name"] == "Picture 1"
        assert result["applied"]["shapes_count"] == 1
        # width/height must reflect actual shape dimensions, not the -1 sentinel
        assert result["applied"]["width"]  == 320.0
        assert result["applied"]["height"] == 240.0

    def test_verify_effect_count_not_increased_raises(self, tmp_path):
        """If Shapes.Count does not increase, should raise ToolError."""
        img = tmp_path / "img6.png"
        img.write_bytes(b"\x89PNG")

        ms = make_mock_session()
        ws = _make_ws()
        shp = _make_shape()
        ws.Shapes = MagicMock()
        ws.Shapes.Count = 2  # before
        ws.Shapes.AddPicture = MagicMock(return_value=shp)
        # Count stays 2 (no-op) — simulating a silent failure
        ws.Range.return_value = MagicMock(Left=0.0, Top=0.0)

        with pytest.raises(ToolError, match="Count did not increase"):
            _call_shape("add_image", ms, ws, left=0.0, top=0.0, filename=str(img))


# ── add_textbox ────────────────────────────────────────────────────────────────

class TestAddTextbox:
    def _make_ws_with_textbox(self, count_before: int = 0) -> tuple[MagicMock, MagicMock]:
        ws = _make_ws()
        shp = _make_shape("TextBox 1", shp_type=17)
        ws.Shapes = MagicMock()
        ws.Shapes.Count = count_before

        def _add_tb(**kwargs):
            ws.Shapes.Count = count_before + 1
            return shp
        ws.Shapes.AddTextbox = MagicMock(side_effect=_add_tb)
        return ws, shp

    def test_calls_addtextbox_with_horizontal_orientation(self):
        ms = make_mock_session()
        ws, shp = self._make_ws_with_textbox()

        _call_shape("add_textbox", ms, ws, left=0.0, top=0.0, width=200.0, height=50.0)

        kw = ws.Shapes.AddTextbox.call_args.kwargs
        assert kw["Orientation"] == 1, "msoTextOrientationHorizontal=1"

    def test_textbox_dimensions_passed(self):
        ms = make_mock_session()
        ws, shp = self._make_ws_with_textbox()

        _call_shape("add_textbox", ms, ws, left=10.0, top=20.0, width=150.0, height=60.0)

        kw = ws.Shapes.AddTextbox.call_args.kwargs
        assert kw["Left"]   == 10.0
        assert kw["Top"]    == 20.0
        assert kw["Width"]  == 150.0
        assert kw["Height"] == 60.0

    def test_text_is_set_via_textframe2(self):
        """TextFrame2.TextRange.Text setter must be called; read-back is verified."""
        ms = make_mock_session()
        ws, shp = self._make_ws_with_textbox()

        # Use a simple container to simulate readable/writable Text property
        # without fighting MagicMock's property machinery.
        store: dict = {"text": ""}

        text_range = MagicMock()
        type(text_range).Text = property(
            fget=lambda self: store["text"],
            fset=lambda self, v: store.update({"text": v}),
        )
        shp.TextFrame2.TextRange = text_range

        result = _call_shape(
            "add_textbox", ms, ws, left=0.0, top=0.0, width=100.0, height=40.0, text="Hello"
        )

        # Text must have been written and read back
        assert store["text"] == "Hello"
        assert result["applied"]["text"] == "Hello"

    def test_no_text_param_skips_text_setter(self):
        ms = make_mock_session()
        ws, shp = self._make_ws_with_textbox()

        result = _call_shape("add_textbox", ms, ws, left=0.0, top=0.0, width=100.0, height=40.0)

        # text key should NOT be in applied when no text was passed
        assert "text" not in result["applied"]

    def test_result_shape(self):
        ms = make_mock_session()
        ws, shp = self._make_ws_with_textbox()

        result = _call_shape("add_textbox", ms, ws, left=5.0, top=5.0, width=100.0, height=50.0)

        assert result["shape"] == "add_textbox"
        assert "sheet" in result
        assert result["applied"]["shapes_count"] == 1
        assert result["applied"]["name"] == "TextBox 1"

    def test_verify_effect_count_not_increased_raises(self):
        ms = make_mock_session()
        ws = _make_ws()
        shp = _make_shape("TextBox 1", shp_type=17)
        ws.Shapes = MagicMock()
        ws.Shapes.Count = 1
        ws.Shapes.AddTextbox = MagicMock(return_value=shp)  # Count stays 1

        with pytest.raises(ToolError, match="Count did not increase"):
            _call_shape("add_textbox", ms, ws, left=0.0, top=0.0, width=100.0, height=50.0)


# ── add_shape ─────────────────────────────────────────────────────────────────

class TestAddShape:
    def _make_ws_with_shape(self, auto_type: int, count_before: int = 0) -> tuple[MagicMock, MagicMock]:
        ws = _make_ws()
        shp = _make_shape("Rectangle 1", auto_type=auto_type, shp_type=1)
        ws.Shapes = MagicMock()
        ws.Shapes.Count = count_before

        def _add_s(**kwargs):
            ws.Shapes.Count = count_before + 1
            return shp
        ws.Shapes.AddShape = MagicMock(side_effect=_add_s)
        return ws, shp

    def test_rectangle_type_int_1(self):
        ms = make_mock_session()
        ws, shp = self._make_ws_with_shape(auto_type=1)

        _call_shape("add_shape", ms, ws, left=0.0, top=0.0, width=100.0, height=50.0, shape_type="rectangle")

        kw = ws.Shapes.AddShape.call_args.kwargs
        assert kw["Type"] == 1, "msoShapeRectangle=1"

    def test_oval_type_int_9(self):
        ms = make_mock_session()
        ws, shp = self._make_ws_with_shape(auto_type=9)
        shp.AutoShapeType = 9

        _call_shape("add_shape", ms, ws, left=0.0, top=0.0, width=100.0, height=50.0, shape_type="oval")

        kw = ws.Shapes.AddShape.call_args.kwargs
        assert kw["Type"] == 9, "msoShapeOval=9"

    def test_rounded_rectangle_type_int_5(self):
        ms = make_mock_session()
        ws, shp = self._make_ws_with_shape(auto_type=5)
        shp.AutoShapeType = 5

        _call_shape("add_shape", ms, ws, left=0.0, top=0.0, width=100.0, height=50.0, shape_type="rounded_rectangle")

        kw = ws.Shapes.AddShape.call_args.kwargs
        assert kw["Type"] == 5, "msoShapeRoundedRectangle=5"

    def test_right_arrow_type_int_33(self):
        ms = make_mock_session()
        ws, shp = self._make_ws_with_shape(auto_type=33)
        shp.AutoShapeType = 33

        _call_shape("add_shape", ms, ws, left=0.0, top=0.0, width=100.0, height=50.0, shape_type="right_arrow")

        kw = ws.Shapes.AddShape.call_args.kwargs
        assert kw["Type"] == 33, "msoShapeRightArrow=33"

    def test_raw_int_string_shape_type(self):
        """Passing shape_type="4" (diamond) should work as raw int."""
        ms = make_mock_session()
        ws, shp = self._make_ws_with_shape(auto_type=4)
        shp.AutoShapeType = 4

        _call_shape("add_shape", ms, ws, left=0.0, top=0.0, width=80.0, height=80.0, shape_type="4")

        kw = ws.Shapes.AddShape.call_args.kwargs
        assert kw["Type"] == 4

    def test_verify_effect_autoshapetype_mismatch_raises(self):
        """If AutoShapeType doesn't match the request, raise ToolError."""
        ms = make_mock_session()
        ws, shp = self._make_ws_with_shape(auto_type=99)  # wrong type returned
        shp.AutoShapeType = 99

        with pytest.raises(ToolError, match="AutoShapeType"):
            _call_shape("add_shape", ms, ws, left=0.0, top=0.0, width=100.0, height=50.0, shape_type="rectangle")

    def test_verify_effect_count_not_increased_raises(self):
        ms = make_mock_session()
        ws = _make_ws()
        shp = _make_shape("Rect 1", auto_type=1, shp_type=1)
        ws.Shapes = MagicMock()
        ws.Shapes.Count = 3  # will not increase
        ws.Shapes.AddShape = MagicMock(return_value=shp)

        with pytest.raises(ToolError, match="Count did not increase"):
            _call_shape("add_shape", ms, ws, left=0.0, top=0.0, width=100.0, height=50.0, shape_type="rectangle")

    def test_dimensions_passed_to_addshape(self):
        ms = make_mock_session()
        ws, shp = self._make_ws_with_shape(auto_type=1)

        _call_shape("add_shape", ms, ws, left=10.0, top=20.0, width=150.0, height=75.0, shape_type="rectangle")

        kw = ws.Shapes.AddShape.call_args.kwargs
        assert kw["Left"]   == 10.0
        assert kw["Top"]    == 20.0
        assert kw["Width"]  == 150.0
        assert kw["Height"] == 75.0

    def test_result_shape(self):
        ms = make_mock_session()
        ws, shp = self._make_ws_with_shape(auto_type=1)

        result = _call_shape("add_shape", ms, ws, left=0.0, top=0.0, width=100.0, height=50.0, shape_type="rectangle")

        assert result["shape"] == "add_shape"
        assert "applied" in result
        assert result["applied"]["type_int"] == 1
        assert result["applied"]["shapes_count"] == 1


# ── list ──────────────────────────────────────────────────────────────────────

class TestList:
    def test_list_returns_all_shapes(self):
        ms = make_mock_session()
        ws = _make_ws("DataSheet")

        shp1 = _make_shape("Picture 1", shp_type=13)
        shp2 = _make_shape("Rectangle 2", auto_type=1, shp_type=1)
        ws.Shapes = MagicMock()
        ws.Shapes.Count = 2
        ws.Shapes.Item.side_effect = lambda i: shp1 if i == 1 else shp2

        result = _call_shape("list", ms, ws)

        assert result["shape"] == "list"
        assert result["applied"]["count"] == 2
        assert len(result["applied"]["shapes"]) == 2

    def test_list_includes_name_and_geometry(self):
        ms = make_mock_session()
        ws = _make_ws()

        shp = _make_shape("MyBox", shp_type=17)
        shp.Left = 30.0; shp.Top = 40.0; shp.Width = 200.0; shp.Height = 80.0
        ws.Shapes = MagicMock()
        ws.Shapes.Count = 1
        ws.Shapes.Item.return_value = shp

        result = _call_shape("list", ms, ws)

        entry = result["applied"]["shapes"][0]
        assert entry["name"] == "MyBox"
        assert entry["left"] == 30.0
        assert entry["top"] == 40.0
        assert entry["width"] == 200.0
        assert entry["height"] == 80.0

    def test_list_empty_sheet_returns_zero_count(self):
        ms = make_mock_session()
        ws = _make_ws()
        ws.Shapes = MagicMock()
        ws.Shapes.Count = 0

        result = _call_shape("list", ms, ws)

        assert result["applied"]["count"] == 0
        assert result["applied"]["shapes"] == []

    def test_list_autoshape_includes_auto_shape_type(self):
        """AutoShape (Type=1) must include auto_shape_type in the entry."""
        ms = make_mock_session()
        ws = _make_ws()

        shp = _make_shape("Oval 1", auto_type=9, shp_type=1)
        ws.Shapes = MagicMock()
        ws.Shapes.Count = 1
        ws.Shapes.Item.return_value = shp

        result = _call_shape("list", ms, ws)
        entry = result["applied"]["shapes"][0]
        assert "auto_shape_type" in entry
        assert entry["auto_shape_type"] == 9


# ── delete ────────────────────────────────────────────────────────────────────

class TestDelete:
    def test_delete_calls_shape_delete(self):
        ms = make_mock_session()
        ws = _make_ws()

        shp = _make_shape("Picture 1")
        ws.Shapes = MagicMock()
        ws.Shapes.Count = 0  # after delete
        ws.Shapes.return_value = shp  # ws.Shapes(name) returns shp

        # After delete, re-enumeration finds nothing
        def _item(i):
            raise IndexError("no items")  # empty collection
        ws.Shapes.Item.side_effect = _item

        result = _call_shape("delete", ms, ws, name="Picture 1")

        shp.Delete.assert_called_once()
        assert result["shape"] == "delete"
        assert result["applied"]["deleted"] is True
        assert result["applied"]["name"] == "Picture 1"

    def test_verify_effect_shape_still_present_raises(self):
        """If shape still in collection after Delete, raise ToolError."""
        ms = make_mock_session()
        ws = _make_ws()

        shp = _make_shape("Stuck Shape")
        ws.Shapes = MagicMock()
        ws.Shapes.Count = 1
        ws.Shapes.return_value = shp

        # Item still returns the shape → still present
        ws.Shapes.Item.return_value = shp

        with pytest.raises(ToolError, match="still present"):
            _call_shape("delete", ms, ws, name="Stuck Shape")


# ── move ──────────────────────────────────────────────────────────────────────

class TestMove:
    def _make_movable_shape(self, initial_left=100.0, initial_top=50.0) -> tuple[MagicMock, MagicMock]:
        ws = _make_ws()
        shp = _make_shape("MyShape")
        shp.Left = initial_left
        shp.Top  = initial_top

        def _set_left(val): shp.Left = val
        def _set_top(val):  shp.Top  = val

        # ws.Shapes(name) returns our shape
        ws.Shapes = MagicMock()
        ws.Shapes.return_value = shp
        return ws, shp

    def test_move_sets_left_and_top(self):
        ms = make_mock_session()
        ws, shp = self._make_movable_shape(initial_left=999.0, initial_top=999.0)

        # MagicMock tracks attribute assignments directly; the production code
        # sets shp.Left/shp.Top then reads them back for the tolerance check.
        result = _call_shape("move", ms, ws, name="MyShape", left=30.0, top=60.0)

        # Shape properties must have been assigned
        assert shp.Left == 30.0
        assert shp.Top  == 60.0
        assert result["shape"] == "move"
        assert result["applied"]["left"]  == 30.0
        assert result["applied"]["top"]   == 60.0

    def test_move_with_cell_anchor(self):
        ms = make_mock_session()
        ws, shp = self._make_movable_shape()
        ws.Range.return_value = MagicMock(Left=72.0, Top=36.0)

        result = _call_shape("move", ms, ws, name="MyShape", cell="C3")

        assert shp.Left == 72.0
        assert shp.Top  == 36.0

    def test_move_optional_resize(self):
        ms = make_mock_session()
        ws, shp = self._make_movable_shape()

        _call_shape("move", ms, ws, name="MyShape", left=10.0, top=10.0, width=300.0, height=150.0)

        assert shp.Width  == 300.0
        assert shp.Height == 150.0

    def test_move_result_shape(self):
        ms = make_mock_session()
        ws, shp = self._make_movable_shape()

        result = _call_shape("move", ms, ws, name="MyShape", left=5.0, top=8.0)

        assert "shape" in result
        assert "applied" in result
        assert result["applied"]["name"] == "MyShape"

    def test_move_verify_effect_tolerance_exceeded_raises(self):
        """If readback Left differs by more than 1pt from requested, raise ToolError.

        Simulates a locked/constrained shape that ignores position updates by using
        a MagicMock whose Left/Top properties always read back the original values
        regardless of what was written (like a grouped or protected shape).
        """
        ms = make_mock_session()
        ws = _make_ws()

        # Use a real object with read-only-style Left/Top that ignore writes
        store = {"left": 999.0, "top": 888.0}

        class _StubShape:
            Name = "MyShape"

            @property
            def Left(self):
                return store["left"]  # always returns original — ignores sets

            @Left.setter
            def Left(self, v):
                pass  # deliberately ignores the write

            @property
            def Top(self):
                return store["top"]

            @Top.setter
            def Top(self, v):
                pass

        shp = _StubShape()
        ws.Shapes = MagicMock()
        ws.Shapes.return_value = shp

        with pytest.raises(ToolError, match="Left="):
            _call_shape("move", ms, ws, name="MyShape", left=10.0, top=20.0)


# ── Positioning helper ─────────────────────────────────────────────────────────

class TestPositioning:
    def test_cell_takes_priority_over_explicit(self):
        """When both cell and left/top are supplied, cell wins."""
        ms = make_mock_session()
        ws = _make_ws()

        shp = _make_shape("Textbox 1", shp_type=17)
        ws.Shapes = MagicMock()
        ws.Shapes.Count = 0

        def _add_tb(**kwargs):
            ws.Shapes.Count = 1
            return shp
        ws.Shapes.AddTextbox = MagicMock(side_effect=_add_tb)
        ws.Range.return_value = MagicMock(Left=88.0, Top=99.0)

        _call_shape("add_textbox", ms, ws, cell="D5", left=0.0, top=0.0, width=100.0, height=50.0)

        kw = ws.Shapes.AddTextbox.call_args.kwargs
        # Cell anchor (88, 99) must win over explicit (0, 0)
        assert kw["Left"] == 88.0
        assert kw["Top"]  == 99.0

    def test_invalid_cell_raises_toolerror(self):
        ms = make_mock_session()
        ws = _make_ws()
        ws.Range.side_effect = Exception("Invalid reference")
        ws.Shapes = MagicMock()
        ws.Shapes.Count = 0
        ws.Shapes.AddTextbox = MagicMock(return_value=_make_shape())

        with pytest.raises(ToolError, match="Invalid cell reference"):
            _call_shape("add_textbox", ms, ws, cell="ZZZ99999", width=100.0, height=50.0)
