"""Drawing object management via ws.Shapes COM API.

Supports adding pictures (embedded images), text boxes, AutoShapes, and
managing existing shapes: list, move, delete.

COM API reference (all values confirmed from Microsoft Learn)
-------------------------------------------------------------
Shapes.AddPicture(Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
  - LinkToFile  : MsoTriState — msoFalse=0 (embed), msoTrue=-1 (link)
  - SaveWithDocument : MsoTriState — msoTrue=-1 (required when LinkToFile=msoFalse)
  - Width/Height=-1  : preserves native file dimensions
  - Filename MUST be an absolute path; COM raises E_FAIL on relative paths.

Shapes.AddTextbox(Orientation, Left, Top, Width, Height)
  - Orientation: MsoTextOrientation — msoTextOrientationHorizontal=1

Shapes.AddShape(Type, Left, Top, Width, Height)
  - Type: MsoAutoShapeType — see _SHAPE_TYPE_MAP for a subset we expose.

Shape.Type (MsoShapeType, read-only) — describes the KIND of shape:
  msoAutoShape=1, msoCallout=2, msoFreeform=5, msoPicture=13, msoTextBox=17

Shape positioning
-----------------
All Left/Top/Width/Height values are in POINTS (1 inch = 72 pt).
Anchor cell positioning: ws.Range(cell).Left and .Top return points already
referenced to the upper-left corner of the sheet — we read them and pass
directly to the Add* methods.

TextFrame2 vs TextFrame
-----------------------
TextFrame2 is the modern (Office 2007+) text API. We try it first; on failure
fall back to the legacy TextFrame.Characters().Text.

VERIFY-EFFECT contract
-----------------------
- add_image  : ws.Shapes.Count increased; shape.Name is present in collection.
- add_textbox: text read-back via TextFrame2.TextRange.Text (or TextFrame fallback).
- add_shape  : ws.Shapes.Count increased; shape.AutoShapeType matches requested.
- delete     : shape name no longer in ws.Shapes after deletion (Count decreased).
- move       : shape.Left and shape.Top match the requested values after the set.

GOTCHA: Wrap all mutations in excel_guard to suppress dialogs during delete/move.
"""

from __future__ import annotations

import os

from fastmcp.exceptions import ToolError

from ..session import ExcelSession, excel_guard

_session = ExcelSession()

# MsoTriState (Microsoft Learn: office.msotristate)
_MSO_FALSE = 0   # msoFalse — embed / don't link
_MSO_TRUE  = -1  # msoTrue  — save embedded picture with the document

# MsoTextOrientation (Microsoft Learn: office.msotextorientation)
_MSO_TEXT_ORIENTATION_HORIZONTAL = 1  # msoTextOrientationHorizontal

# MsoAutoShapeType subset (Microsoft Learn: office.msoautoshapetype)
# We expose a small curated set; raw ints pass through for any unlisted value.
_SHAPE_TYPE_MAP: dict[str, int] = {
    "rectangle":          1,   # msoShapeRectangle
    "rounded_rectangle":  5,   # msoShapeRoundedRectangle
    "oval":               9,   # msoShapeOval
    "right_arrow":        33,  # msoShapeRightArrow
    "diamond":            4,   # msoShapeDiamond
    "triangle":           7,   # msoShapeIsoscelesTriangle
    "hexagon":            10,  # msoShapeHexagon
    "star":               92,  # msoShape5pointStar
    "cloud":              179, # msoShapeCloud
    "heart":              21,  # msoShapeHeart
}

# MsoShapeType (Microsoft Learn: office.msoshapetype) — for Shape.Type read-back
_MSO_SHAPE_TYPE_NAMES: dict[int, str] = {
    1:  "AutoShape",
    2:  "Callout",
    5:  "Freeform",
    9:  "Line",
    13: "Picture",
    17: "TextBox",
}


# ── Public entry point ─────────────────────────────────────────────────────────

def shape_action(
    action: str,
    sheet: str | None = None,
    workbook: str | None = None,
    # --- positioning (anchor cell OR explicit points) ---
    cell: str | None = None,
    left: float | None = None,
    top: float | None = None,
    width: float | None = None,
    height: float | None = None,
    # --- add_image ---
    filename: str | None = None,
    # --- add_textbox / add_shape optional text ---
    text: str | None = None,
    # --- add_shape ---
    shape_type: str | None = None,
    # --- delete / move ---
    name: str | None = None,
) -> dict:
    """Dispatch a shape management action.

    Actions
    -------
    add_image
        Embed an image from *filename* (absolute path) into the sheet.
        Position via *cell* anchor or explicit *left*/*top* in points.
        *width*/*height* in points; omit (or pass -1) to keep native dimensions.

    add_textbox
        Add a horizontal text box at the given position.
        *width* and *height* are required (in points).
        Optional *text* sets the initial content.

    add_shape
        Add an AutoShape of type *shape_type*.
        Named types: rectangle, rounded_rectangle, oval, right_arrow, diamond,
        triangle, hexagon, star, cloud, heart.  Or pass a raw integer as a string
        (e.g. shape_type="1" for rectangle).
        *width* and *height* are required (in points).
        Optional *text* sets label text.

    list
        Enumerate shapes on the sheet: name, shape_type, left, top, width, height.

    delete
        Delete the shape named *name* from the sheet.

    move
        Reposition (and optionally resize) the shape named *name*.
        Position via *cell* anchor or explicit *left*/*top* in points.
        Optional *width*/*height* resize the shape.
    """
    valid = {"add_image", "add_textbox", "add_shape", "list", "delete", "move"}
    if action not in valid:
        raise ToolError(
            f"Unknown action '{action}'. Valid: {', '.join(sorted(valid))}."
        )
    return _session.run_com(
        _dispatch,
        action, sheet, workbook,
        cell, left, top, width, height,
        filename, text, shape_type, name,
    )


# ── Worker-thread dispatcher ───────────────────────────────────────────────────

def _dispatch(
    action: str,
    sheet: str | None,
    workbook: str | None,
    cell: str | None,
    left: float | None,
    top: float | None,
    width: float | None,
    height: float | None,
    filename: str | None,
    text: str | None,
    shape_type: str | None,
    name: str | None,
) -> dict:
    """Executed on the STA COM worker thread."""
    ws = _session.get_sheet(sheet, workbook)
    app = ws.Application

    with excel_guard(app):
        if action == "add_image":
            return _add_image(ws, cell, left, top, width, height, filename)
        if action == "add_textbox":
            return _add_textbox(ws, cell, left, top, width, height, text)
        if action == "add_shape":
            return _add_shape(ws, cell, left, top, width, height, shape_type, text)
        if action == "list":
            return _list(ws)
        if action == "delete":
            return _delete(ws, name)
        # action == "move"
        return _move(ws, name, cell, left, top, width, height)


# ── Positioning helpers ────────────────────────────────────────────────────────

def _resolve_position(
    ws,
    cell: str | None,
    left: float | None,
    top: float | None,
) -> tuple[float, float]:
    """Return (left_pts, top_pts) from cell anchor or explicit values.

    Cell anchor takes priority when supplied.  Reading Range.Left / .Top
    returns points referenced to the sheet upper-left — ready for AddPicture
    and friends.
    """
    if cell is not None:
        try:
            rng = ws.Range(cell)
        except Exception as e:
            raise ToolError(f"Invalid cell reference '{cell}': {e}")
        return float(rng.Left), float(rng.Top)
    if left is None or top is None:
        raise ToolError(
            "Supply either 'cell' (e.g. 'B3') OR both 'left' and 'top' in points."
        )
    return float(left), float(top)


# ── Action implementations ─────────────────────────────────────────────────────

def _add_image(
    ws,
    cell: str | None,
    left: float | None,
    top: float | None,
    width: float | None,
    height: float | None,
    filename: str | None,
) -> dict:
    """Embed an image from filename into the worksheet.

    AddPicture requires:
      - An absolute file path (COM raises E_FAIL on relative paths).
      - LinkToFile=msoFalse (0) to embed; SaveWithDocument=msoTrue (-1) required
        when LinkToFile=msoFalse per the MsoTriState contract.
      - Width/Height=-1 preserves the native image dimensions.

    VERIFY-EFFECT: Shapes.Count increases by 1; shape.Name is non-empty.
    """
    if not filename:
        raise ToolError("add_image requires 'filename' (absolute path to the image file).")
    if not os.path.isfile(filename):
        raise ToolError(
            f"add_image: file not found: '{filename}'. "
            "Provide an absolute path to an existing image file."
        )

    left_pts, top_pts = _resolve_position(ws, cell, left, top)
    # -1 means "use native size" per the Excel AddPicture spec
    w_pts = float(width) if width is not None and width != -1 else -1.0
    h_pts = float(height) if height is not None and height != -1 else -1.0

    try:
        count_before = ws.Shapes.Count
        shape = ws.Shapes.AddPicture(
            Filename=filename,
            LinkToFile=_MSO_FALSE,
            SaveWithDocument=_MSO_TRUE,
            Left=left_pts,
            Top=top_pts,
            Width=w_pts,
            Height=h_pts,
        )
        count_after = ws.Shapes.Count
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "Shapes.AddPicture failed")

    # VERIFY-EFFECT: count must have grown
    if count_after <= count_before:
        raise ToolError(
            f"AddPicture reported no error but Shapes.Count did not increase "
            f"({count_before} → {count_after}). Check that the file is a supported "
            "image format and the path is accessible."
        )

    # Read actual dimensions back from the shape object.  When Width/Height=-1
    # (native size) the requested value is -1 but the shape now has real pixel
    # dimensions; returning -1 would be misleading.
    shape_name  = shape.Name
    actual_left = shape.Left
    actual_top  = shape.Top
    actual_w    = shape.Width
    actual_h    = shape.Height

    return {
        "shape": "add_image",
        "sheet": ws.Name,
        "applied": {
            "name": shape_name,
            "filename": filename,
            "left": actual_left,
            "top": actual_top,
            "width": actual_w,
            "height": actual_h,
            "shapes_count": count_after,
        },
    }


def _add_textbox(
    ws,
    cell: str | None,
    left: float | None,
    top: float | None,
    width: float | None,
    height: float | None,
    text: str | None,
) -> dict:
    """Add a horizontal text box.

    Uses AddTextbox(Orientation=1 [msoTextOrientationHorizontal], ...).
    Text is set via TextFrame2.TextRange.Text (modern API) with a fallback
    to the legacy TextFrame.Characters().Text path.

    VERIFY-EFFECT: text read back from TextFrame2 must match what was set.
    Width and height are required (text boxes need explicit sizing).
    """
    left_pts, top_pts = _resolve_position(ws, cell, left, top)
    if width is None:
        raise ToolError("add_textbox requires 'width' in points.")
    if height is None:
        raise ToolError("add_textbox requires 'height' in points.")
    w_pts = float(width)
    h_pts = float(height)

    try:
        count_before = ws.Shapes.Count
        shape = ws.Shapes.AddTextbox(
            Orientation=_MSO_TEXT_ORIENTATION_HORIZONTAL,
            Left=left_pts,
            Top=top_pts,
            Width=w_pts,
            Height=h_pts,
        )
        count_after = ws.Shapes.Count
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "Shapes.AddTextbox failed")

    if count_after <= count_before:
        raise ToolError(
            f"AddTextbox reported no error but Shapes.Count did not increase "
            f"({count_before} → {count_after})."
        )

    # Set text content — try modern TextFrame2 first, legacy fallback
    readback_text: str | None = None
    if text is not None:
        try:
            shape.TextFrame2.TextRange.Text = text
            readback_text = shape.TextFrame2.TextRange.Text
        except Exception:
            try:
                shape.TextFrame.Characters().Text = text
                readback_text = shape.TextFrame.Characters().Text
            except Exception as e2:
                raise _session.wrap(e2, "Setting textbox text failed")

        # VERIFY-EFFECT: text must round-trip
        if readback_text != text:
            raise ToolError(
                f"Textbox text was set but read-back differs: "
                f"expected {text!r}, got {readback_text!r}."
            )

    applied: dict = {
        "name": shape.Name,
        "left": left_pts,
        "top": top_pts,
        "width": w_pts,
        "height": h_pts,
        "shapes_count": count_after,
    }
    if text is not None:
        applied["text"] = readback_text

    return {
        "shape": "add_textbox",
        "sheet": ws.Name,
        "applied": applied,
    }


def _add_shape(
    ws,
    cell: str | None,
    left: float | None,
    top: float | None,
    width: float | None,
    height: float | None,
    shape_type: str | None,
    text: str | None,
) -> dict:
    """Add an AutoShape.

    shape_type accepts a name from _SHAPE_TYPE_MAP or a raw integer string.
    Width and height are required.

    VERIFY-EFFECT: Shapes.Count increases; shape.AutoShapeType matches.
    """
    if shape_type is None:
        raise ToolError(
            f"add_shape requires 'shape_type'. Named types: "
            f"{', '.join(sorted(_SHAPE_TYPE_MAP))}. "
            "Or pass a raw msoAutoShapeType integer as a string."
        )

    # Resolve shape_type to an int
    type_int: int | None = _SHAPE_TYPE_MAP.get(shape_type.lower())
    if type_int is None:
        # Try interpreting as a raw integer
        try:
            type_int = int(shape_type)
        except ValueError:
            raise ToolError(
                f"Unknown shape_type '{shape_type}'. "
                f"Named types: {', '.join(sorted(_SHAPE_TYPE_MAP))}. "
                "Or pass a raw msoAutoShapeType integer as a string."
            )

    left_pts, top_pts = _resolve_position(ws, cell, left, top)
    if width is None:
        raise ToolError("add_shape requires 'width' in points.")
    if height is None:
        raise ToolError("add_shape requires 'height' in points.")
    w_pts = float(width)
    h_pts = float(height)

    try:
        count_before = ws.Shapes.Count
        shape = ws.Shapes.AddShape(
            Type=type_int,
            Left=left_pts,
            Top=top_pts,
            Width=w_pts,
            Height=h_pts,
        )
        count_after = ws.Shapes.Count
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "Shapes.AddShape failed")

    if count_after <= count_before:
        raise ToolError(
            f"AddShape reported no error but Shapes.Count did not increase "
            f"({count_before} → {count_after})."
        )

    # VERIFY-EFFECT: AutoShapeType must match
    auto_shape_type = shape.AutoShapeType
    if auto_shape_type != type_int:
        raise ToolError(
            f"AddShape: requested AutoShapeType={type_int} but got {auto_shape_type}."
        )

    # Optional text on the shape
    readback_text: str | None = None
    if text is not None:
        try:
            shape.TextFrame2.TextRange.Text = text
            readback_text = shape.TextFrame2.TextRange.Text
        except Exception:
            try:
                shape.TextFrame.Characters().Text = text
                readback_text = shape.TextFrame.Characters().Text
            except Exception as e2:
                raise _session.wrap(e2, "Setting shape text failed")

        if readback_text != text:
            raise ToolError(
                f"Shape text set but read-back differs: "
                f"expected {text!r}, got {readback_text!r}."
            )

    applied: dict = {
        "name": shape.Name,
        "shape_type": shape_type,
        "type_int": type_int,
        "auto_shape_type": auto_shape_type,
        "left": left_pts,
        "top": top_pts,
        "width": w_pts,
        "height": h_pts,
        "shapes_count": count_after,
    }
    if readback_text is not None:
        applied["text"] = readback_text

    return {
        "shape": "add_shape",
        "sheet": ws.Name,
        "applied": applied,
    }


def _list(ws) -> dict:
    """Enumerate shapes on the sheet.

    Uses .Count and .Item(i) (1-based) which is more reliable than Python
    for-in iteration over a COM dispatch collection.
    Shape.Type gives the MsoShapeType; .AutoShapeType gives MsoAutoShapeType
    (only meaningful when Type==msoAutoShape=1, else returns -2).
    """
    try:
        shapes_col = ws.Shapes
        count = shapes_col.Count
    except Exception as e:
        raise _session.wrap(e, "shapes list: Shapes collection unavailable")

    shapes_info = []
    for i in range(1, count + 1):
        try:
            shp = shapes_col.Item(i)
            type_int = shp.Type
            type_name = _MSO_SHAPE_TYPE_NAMES.get(type_int, str(type_int))
            entry: dict = {
                "name":       shp.Name,
                "type":       type_name,
                "type_int":   type_int,
                "left":       shp.Left,
                "top":        shp.Top,
                "width":      shp.Width,
                "height":     shp.Height,
            }
            # AutoShapeType is only meaningful for AutoShape (type 1)
            if type_int == 1:
                try:
                    entry["auto_shape_type"] = shp.AutoShapeType
                except Exception:
                    pass
            shapes_info.append(entry)
        except Exception as ex:
            shapes_info.append({"index": i, "error": str(ex)})

    return {
        "shape": "list",
        "sheet": ws.Name,
        "applied": {
            "count": count,
            "shapes": shapes_info,
        },
    }


def _delete(ws, name: str | None) -> dict:
    """Delete a shape by name.

    VERIFY-EFFECT: name no longer in ws.Shapes after deletion.
    """
    if name is None:
        raise ToolError("delete action requires 'name' (the shape's name).")

    try:
        ws.Shapes(name).Delete()
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"Shapes.Delete('{name}') failed — shape may not exist")

    # VERIFY-EFFECT: confirm name is gone
    try:
        count_after = ws.Shapes.Count
        still_present = any(
            ws.Shapes.Item(i).Name == name
            for i in range(1, count_after + 1)
        )
    except Exception as e:
        raise _session.wrap(e, "delete verify-effect: re-enumeration failed")

    if still_present:
        raise ToolError(
            f"Delete was called but shape '{name}' is still present in the collection."
        )

    return {
        "shape": "delete",
        "sheet": ws.Name,
        "applied": {
            "name": name,
            "deleted": True,
            "shapes_count": count_after,
        },
    }


def _move(
    ws,
    name: str | None,
    cell: str | None,
    left: float | None,
    top: float | None,
    width: float | None,
    height: float | None,
) -> dict:
    """Reposition (and optionally resize) a named shape.

    Position from cell anchor or explicit left/top in points.
    VERIFY-EFFECT: read back .Left and .Top after setting.
    """
    if name is None:
        raise ToolError("move action requires 'name' (the shape's name).")

    left_pts, top_pts = _resolve_position(ws, cell, left, top)

    try:
        shp = ws.Shapes(name)
    except Exception as e:
        raise _session.wrap(e, f"Shape '{name}' not found")

    try:
        shp.Left = left_pts
        shp.Top  = top_pts
        if width is not None:
            shp.Width  = float(width)
        if height is not None:
            shp.Height = float(height)

        # VERIFY-EFFECT: read back position
        actual_left  = shp.Left
        actual_top   = shp.Top
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"move shape '{name}' failed")

    # Tolerance: floating-point round-trip through COM can differ by a fraction
    _TOL = 1.0  # 1 point tolerance
    if abs(actual_left - left_pts) > _TOL or abs(actual_top - top_pts) > _TOL:
        raise ToolError(
            f"move: set Left={left_pts}, Top={top_pts} but read back "
            f"Left={actual_left}, Top={actual_top}. Move may have been constrained."
        )

    applied: dict = {
        "name": name,
        "left": actual_left,
        "top": actual_top,
    }
    if width is not None:
        applied["width"] = shp.Width
    if height is not None:
        applied["height"] = shp.Height

    return {
        "shape": "move",
        "sheet": ws.Name,
        "applied": applied,
    }
