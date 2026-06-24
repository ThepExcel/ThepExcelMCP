"""Worksheet page setup: orientation, paper size, margins, fit-to-page, headers/footers, PDF export.

Operates on ws.PageSetup (a per-worksheet COM object) and
ws.ExportAsFixedFormat / wb.ExportAsFixedFormat.

COM gotchas
-----------
1. Zoom vs FitToPages:  When either FitToPagesWide or FitToPagesTall is set,
   Excel requires Zoom=False to activate fit-to-page mode.  Setting both Zoom
   and FitTo* leaves Zoom active and ignores FitTo*.

2. Margins are stored INTERNALLY in POINTS; the API surfaces inches.
   Conversion: PageSetup.TopMargin = Application.InchesToPoints(value_in_inches).
   Reading back gives points; we convert back to inches for 'get'.

3. PageSetup property writes can raise if no printer driver is installed.
   The error is not fatal — catch it and re-raise with an actionable hint.

4. PrintArea uses absolute references ($A$1:$F$50).  Passing "" clears the
   print area (equivalent to PageSetup.PrintArea = "").

5. ExportAsFixedFormat Type=0 = xlTypePDF; Quality=0 = xlQualityStandard.
   After the call we assert os.path.exists(path) — if the file is missing the
   export silently failed (no printer driver, path permission, etc.).

6. xl code strings in headers/footers (e.g. "&P", "&N", "&D", "&T", "&F",
   "&A") are passed through verbatim; we do not validate or escape them.

Excel COM property → verified constant source (Microsoft Learn):
  Orientation: xlPortrait=1, xlLandscape=2
  PaperSize:   xlPaperLetter=1, xlPaperA4=9, xlPaperA3=8, xlPaperLegal=5
  ExportAsFixedFormat: Type=0 (xlTypePDF), Quality=0 (xlQualityStandard)
"""

from __future__ import annotations

import os

from fastmcp.exceptions import ToolError

from ..session import ExcelSession, excel_guard

_session = ExcelSession()

# XlPageOrientation (Microsoft Learn: excel.xlpageorientation)
_XL_PORTRAIT  = 1   # xlPortrait
_XL_LANDSCAPE = 2   # xlLandscape

# XlPaperSize subset (Microsoft Learn: excel.xlpapersize)
_PAPER_SIZES = {
    "letter": 1,   # xlPaperLetter  — 8.5 × 11 in
    "a4":     9,   # xlPaperA4      — 210 × 297 mm
    "a3":     8,   # xlPaperA3      — 297 × 420 mm
    "legal":  5,   # xlPaperLegal   — 8.5 × 14 in
}

# ExportAsFixedFormat constants
_XL_TYPE_PDF            = 0   # xlTypePDF
_XL_QUALITY_STANDARD    = 0   # xlQualityStandard

# Margin property names on PageSetup
_MARGIN_PROPS = {
    "top":     "TopMargin",
    "bottom":  "BottomMargin",
    "left":    "LeftMargin",
    "right":   "RightMargin",
    "header":  "HeaderMargin",
    "footer":  "FooterMargin",
}

# Header/footer slot property names on PageSetup
_HF_PROPS = {
    "left_header":    "LeftHeader",
    "center_header":  "CenterHeader",
    "right_header":   "RightHeader",
    "left_footer":    "LeftFooter",
    "center_footer":  "CenterFooter",
    "right_footer":   "RightFooter",
}


# ── Public entry point ─────────────────────────────────────────────────────────

def page_setup_action(
    action: str,
    sheet: str | None = None,
    workbook: str | None = None,
    # --- set ---
    orientation: str | None = None,
    paper_size: str | None = None,
    fit_to_wide: int | None = None,
    fit_to_tall: int | None = None,
    scale: int | None = None,
    # margins in inches
    top: float | None = None,
    bottom: float | None = None,
    left: float | None = None,
    right: float | None = None,
    header: float | None = None,
    footer: float | None = None,
    center_horizontally: bool | None = None,
    center_vertically: bool | None = None,
    print_gridlines: bool | None = None,
    black_and_white: bool | None = None,
    # --- print_area ---
    address: str | None = None,
    # --- print_titles ---
    rows: str | None = None,
    cols: str | None = None,
    # --- header_footer ---
    left_header: str | None = None,
    center_header: str | None = None,
    right_header: str | None = None,
    left_footer: str | None = None,
    center_footer: str | None = None,
    right_footer: str | None = None,
    # --- export_pdf ---
    path: str | None = None,
    scope: str = "sheet",
    open_after: bool = False,
) -> dict:
    """Dispatch a page-setup action.

    Actions
    -------
    set
        Set any subset of page-setup properties.  Only supplied (non-None)
        properties are written; omitted ones are left unchanged.
        orientation: "portrait" | "landscape"
        paper_size:  "a4" | "letter" | "a3" | "legal"
        fit_to_wide: int — FitToPagesWide (also sets Zoom=False)
        fit_to_tall: int — FitToPagesTall (also sets Zoom=False)
        scale:       int 10–400 → Zoom
        top/bottom/left/right/header/footer: float in inches → converted to points
        center_horizontally / center_vertically: bool
        print_gridlines / black_and_white: bool
    print_area
        Set PrintArea to address (e.g. "A1:F50").  Pass address="" to clear.
    print_titles
        Set PrintTitleRows (e.g. rows="$1:$3") and/or PrintTitleColumns
        (e.g. cols="$A:$B").
    header_footer
        Set header/footer slots.  Excel codes: &P=page, &N=pages, &D=date,
        &T=time, &F=filename, &A=sheet name.
    export_pdf
        Export to PDF via ExportAsFixedFormat.
        path:       required, full path ending in .pdf
        scope:      "sheet" (default) | "workbook"
        open_after: bool (default False)
    get
        Read back current page-setup properties as a dict.
    """
    valid = {"set", "print_area", "print_titles", "header_footer", "export_pdf", "get"}
    if action not in valid:
        raise ToolError(
            f"Unknown action '{action}'. Valid: {', '.join(sorted(valid))}."
        )
    return _session.run_com(
        _dispatch,
        action, sheet, workbook,
        orientation, paper_size, fit_to_wide, fit_to_tall, scale,
        top, bottom, left, right, header, footer,
        center_horizontally, center_vertically,
        print_gridlines, black_and_white,
        address,
        rows, cols,
        left_header, center_header, right_header,
        left_footer, center_footer, right_footer,
        path, scope, open_after,
    )


# ── Worker-thread dispatcher ───────────────────────────────────────────────────

def _dispatch(
    action, sheet, workbook,
    orientation, paper_size, fit_to_wide, fit_to_tall, scale,
    top, bottom, left, right, header, footer,
    center_horizontally, center_vertically,
    print_gridlines, black_and_white,
    address,
    rows, cols,
    left_header, center_header, right_header,
    left_footer, center_footer, right_footer,
    path, scope, open_after,
) -> dict:
    """Executed on the STA COM worker thread."""
    ws = _session.get_sheet(sheet, workbook)
    app = ws.Application

    with excel_guard(app):
        if action == "set":
            return _set(
                ws, app,
                orientation, paper_size, fit_to_wide, fit_to_tall, scale,
                top, bottom, left, right, header, footer,
                center_horizontally, center_vertically,
                print_gridlines, black_and_white,
            )
        if action == "print_area":
            return _print_area(ws, address)
        if action == "print_titles":
            return _print_titles(ws, rows, cols)
        if action == "header_footer":
            return _header_footer(
                ws,
                left_header, center_header, right_header,
                left_footer, center_footer, right_footer,
            )
        if action == "export_pdf":
            return _export_pdf(ws, app, path, scope, open_after)
        # action == "get"
        return _get(ws, app)


# ── Action implementations ─────────────────────────────────────────────────────

def _set(
    ws, app,
    orientation, paper_size, fit_to_wide, fit_to_tall, scale,
    top, bottom, left, right, header, footer,
    center_horizontally, center_vertically,
    print_gridlines, black_and_white,
) -> dict:
    """Set any subset of PageSetup properties — only non-None values are applied."""
    ps = ws.PageSetup
    applied: dict = {}

    try:
        if orientation is not None:
            ori_map = {"portrait": _XL_PORTRAIT, "landscape": _XL_LANDSCAPE}
            ori_int = ori_map.get(orientation.lower())
            if ori_int is None:
                raise ToolError(
                    f"Unknown orientation '{orientation}'. Valid: portrait, landscape."
                )
            ps.Orientation = ori_int
            applied["orientation"] = orientation

        if paper_size is not None:
            ps_int = _PAPER_SIZES.get(paper_size.lower())
            if ps_int is None:
                raise ToolError(
                    f"Unknown paper_size '{paper_size}'. "
                    f"Valid: {', '.join(sorted(_PAPER_SIZES))}."
                )
            ps.PaperSize = ps_int
            applied["paper_size"] = paper_size

        # fit_to_wide / fit_to_tall — must set Zoom=False first
        if fit_to_wide is not None or fit_to_tall is not None:
            ps.Zoom = False          # required: disable scale-zoom when using fit-to-pages
            if fit_to_wide is not None:
                ps.FitToPagesWide = int(fit_to_wide)
                applied["fit_to_wide"] = fit_to_wide
            if fit_to_tall is not None:
                ps.FitToPagesTall = int(fit_to_tall)
                applied["fit_to_tall"] = fit_to_tall
            applied["zoom_disabled"] = True

        if scale is not None:
            if not (10 <= scale <= 400):
                raise ToolError(
                    f"scale must be between 10 and 400 (got {scale})."
                )
            ps.Zoom = int(scale)
            applied["scale"] = scale

        # Margins: convert inches → points via Application.InchesToPoints
        margins = {
            "top": top, "bottom": bottom, "left": left,
            "right": right, "header": header, "footer": footer,
        }
        for key, value in margins.items():
            if value is not None:
                points = app.InchesToPoints(float(value))
                setattr(ps, _MARGIN_PROPS[key], points)
                applied[f"margin_{key}_in"] = value

        if center_horizontally is not None:
            ps.CenterHorizontally = bool(center_horizontally)
            applied["center_horizontally"] = center_horizontally

        if center_vertically is not None:
            ps.CenterVertically = bool(center_vertically)
            applied["center_vertically"] = center_vertically

        if print_gridlines is not None:
            ps.PrintGridlines = bool(print_gridlines)
            applied["print_gridlines"] = print_gridlines

        if black_and_white is not None:
            ps.BlackAndWhite = bool(black_and_white)
            applied["black_and_white"] = black_and_white

    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(
            e,
            "PageSetup.set failed — ensure a printer driver is installed "
            "(even a PDF/virtual printer), as Excel requires one for PageSetup writes",
        )

    if not applied:
        raise ToolError(
            "set action: no properties were supplied. Pass at least one of: "
            "orientation, paper_size, fit_to_wide, fit_to_tall, scale, "
            "top/bottom/left/right/header/footer, "
            "center_horizontally, center_vertically, print_gridlines, black_and_white."
        )

    return {
        "page_setup": "set",
        "sheet": ws.Name,
        "applied": applied,
    }


def _print_area(ws, address: str | None) -> dict:
    """Set or clear the print area."""
    if address is None:
        raise ToolError(
            "print_area action requires 'address' "
            "(e.g. 'A1:F50', or '' to clear the print area)."
        )
    try:
        ws.PageSetup.PrintArea = address
    except Exception as e:
        raise _session.wrap(e, "print_area failed")
    return {
        "page_setup": "print_area",
        "sheet": ws.Name,
        "applied": {"print_area": address or "(cleared)"},
    }


def _print_titles(ws, rows: str | None, cols: str | None) -> dict:
    """Set print title rows and/or columns."""
    if rows is None and cols is None:
        raise ToolError(
            "print_titles action requires at least one of 'rows' or 'cols'. "
            "Examples: rows='$1:$1', cols='$A:$A'."
        )
    applied: dict = {}
    try:
        if rows is not None:
            ws.PageSetup.PrintTitleRows = rows
            applied["print_title_rows"] = rows
        if cols is not None:
            ws.PageSetup.PrintTitleColumns = cols
            applied["print_title_columns"] = cols
    except Exception as e:
        raise _session.wrap(e, "print_titles failed")
    return {
        "page_setup": "print_titles",
        "sheet": ws.Name,
        "applied": applied,
    }


def _header_footer(
    ws,
    left_header, center_header, right_header,
    left_footer, center_footer, right_footer,
) -> dict:
    """Set header/footer slots.

    Supported Excel codes (pass them inside the string):
      &P  — current page number
      &N  — total pages
      &D  — current date
      &T  — current time
      &F  — file name
      &A  — sheet (tab) name
    """
    slots = {
        "left_header":   left_header,
        "center_header": center_header,
        "right_header":  right_header,
        "left_footer":   left_footer,
        "center_footer": center_footer,
        "right_footer":  right_footer,
    }
    supplied = {k: v for k, v in slots.items() if v is not None}
    if not supplied:
        raise ToolError(
            "header_footer action requires at least one of: "
            "left_header, center_header, right_header, "
            "left_footer, center_footer, right_footer."
        )
    ps = ws.PageSetup
    applied: dict = {}
    try:
        for key, value in supplied.items():
            setattr(ps, _HF_PROPS[key], value)
            applied[key] = value
    except Exception as e:
        raise _session.wrap(e, "header_footer failed")
    return {
        "page_setup": "header_footer",
        "sheet": ws.Name,
        "applied": applied,
    }


def _export_pdf(ws, app, path: str | None, scope: str, open_after: bool) -> dict:
    """Export sheet or workbook to PDF via ExportAsFixedFormat."""
    if path is None:
        raise ToolError("export_pdf action requires 'path' (full path ending in .pdf).")
    if not path.lower().endswith(".pdf"):
        raise ToolError(f"export_pdf: 'path' must end with .pdf (got '{path}').")
    if scope not in ("sheet", "workbook"):
        raise ToolError(
            f"export_pdf: 'scope' must be 'sheet' or 'workbook' (got '{scope}')."
        )

    try:
        if scope == "workbook":
            target = ws.Parent  # the Workbook
        else:
            target = ws

        target.ExportAsFixedFormat(
            Type=_XL_TYPE_PDF,
            Filename=path,
            Quality=_XL_QUALITY_STANDARD,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=open_after,
        )
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(
            e,
            "export_pdf failed — ensure a PDF printer driver is installed "
            "and the destination path is writable",
        )

    # VERIFY EFFECT: the file must exist after export
    if not os.path.exists(path):
        raise ToolError(
            f"export_pdf: ExportAsFixedFormat reported no error but the file "
            f"'{path}' was not created. Check path permissions and printer driver."
        )

    file_size = os.path.getsize(path)
    return {
        "page_setup": "export_pdf",
        "sheet": ws.Name,
        "applied": {
            "path": path,
            "scope": scope,
            "file_size_bytes": file_size,
        },
    }


def _get(ws, app) -> dict:
    """Read back current PageSetup properties."""
    ps = ws.PageSetup
    # Points → inches helper; Application.InchesToPoints is the forward direction;
    # reverse: points / 72 = inches (Excel defines 1 inch = 72 points).
    _pts_to_in = lambda pts: round(pts / 72.0, 4)

    try:
        orientation_int = ps.Orientation
        orientation_str = {_XL_PORTRAIT: "portrait", _XL_LANDSCAPE: "landscape"}.get(
            orientation_int, str(orientation_int)
        )
        paper_int = ps.PaperSize
        paper_str = {v: k for k, v in _PAPER_SIZES.items()}.get(paper_int, str(paper_int))

        zoom_val = ps.Zoom   # False when fit-to-page is active; int when scale-zoom is set

        return {
            "page_setup": "get",
            "sheet": ws.Name,
            "applied": {
                "orientation": orientation_str,
                "paper_size": paper_str,
                "zoom": zoom_val,
                "fit_to_wide": ps.FitToPagesWide,
                "fit_to_tall": ps.FitToPagesTall,
                "margins_in": {
                    "top":    _pts_to_in(ps.TopMargin),
                    "bottom": _pts_to_in(ps.BottomMargin),
                    "left":   _pts_to_in(ps.LeftMargin),
                    "right":  _pts_to_in(ps.RightMargin),
                    "header": _pts_to_in(ps.HeaderMargin),
                    "footer": _pts_to_in(ps.FooterMargin),
                },
                "center_horizontally": ps.CenterHorizontally,
                "center_vertically":   ps.CenterVertically,
                "print_gridlines":     ps.PrintGridlines,
                "black_and_white":     ps.BlackAndWhite,
                "print_area":          ps.PrintArea,
                "print_title_rows":    ps.PrintTitleRows,
                "print_title_columns": ps.PrintTitleColumns,
                "left_header":    ps.LeftHeader,
                "center_header":  ps.CenterHeader,
                "right_header":   ps.RightHeader,
                "left_footer":    ps.LeftFooter,
                "center_footer":  ps.CenterFooter,
                "right_footer":   ps.RightFooter,
            },
        }
    except Exception as e:
        raise _session.wrap(e, "page_setup.get failed")
