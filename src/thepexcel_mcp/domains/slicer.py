"""Slicer and Timeline operations on workbook-level SlicerCaches.

Slicers connect a visual filter control to a PivotTable or Table (ListObject).
Timelines are date-field slicers that use SlicerCaches.Add2 with the xlTimeline
cache type and require a date-typed field in the source PivotTable.

Source resolution
-----------------
``add`` and ``add_timeline`` accept a *source* name (str).  Resolution order:
  1. Search every worksheet for a PivotTable whose Name matches.
  2. Search every worksheet for a ListObject whose Name matches.
The resolved COM object is passed directly to ``SlicerCaches.Add2`` — the API
requires the actual object, not a string (a string would be interpreted as a
WorkbookConnection name, which is different).

XlSlicerCacheType enum values (Excel VBA)
-----------------------------------------
  xlSlicer   = 1  — regular field slicer
  xlTimeline = 2  — date-field timeline slicer
"""

from __future__ import annotations

from fastmcp.exceptions import ToolError

from ..session import ExcelSession, excel_guard

_session = ExcelSession()

# XlSlicerCacheType (Excel VBA enum)
_XL_SLICER    = 1   # xlSlicer
_XL_TIMELINE  = 2   # xlTimeline

# Sensible position / size defaults (points)
_DEFAULT_TOP    = 50.0
_DEFAULT_LEFT   = 50.0
_DEFAULT_WIDTH  = 144.0   # ~2 inches
_DEFAULT_HEIGHT = 144.0   # ~2 inches


# ── Public entry point ─────────────────────────────────────────────────────────

def slicer_action(
    action: str,
    workbook: str | None = None,
    # add / add_timeline
    source: str | None = None,
    field: str | None = None,
    sheet: str | None = None,
    caption: str | None = None,
    name: str | None = None,
    top: float = _DEFAULT_TOP,
    left: float = _DEFAULT_LEFT,
    width: float = _DEFAULT_WIDTH,
    height: float = _DEFAULT_HEIGHT,
    # delete / connect
    slicer: str | None = None,
    # connect
    pivots: list[str] | None = None,
) -> dict:
    """Dispatch a slicer or timeline action."""
    valid = {"add", "add_timeline", "list", "delete", "connect"}
    if action not in valid:
        raise ToolError(
            f"Unknown action '{action}'. Valid: {', '.join(sorted(valid))}."
        )
    return _session.run_com(
        _dispatch,
        action, workbook,
        source, field, sheet, caption, name,
        top, left, width, height,
        slicer, pivots,
    )


# ── Worker-thread dispatcher ───────────────────────────────────────────────────

def _dispatch(
    action, workbook,
    source, field, sheet, caption, name,
    top, left, width, height,
    slicer_name, pivots,
) -> dict:
    """Executed on the STA COM worker thread."""
    wb = _session.get_workbook(workbook)
    app = wb.Application

    with excel_guard(app):
        if action == "add":
            return _add(wb, source, field, sheet, caption, name,
                        top, left, width, height, _XL_SLICER)
        if action == "add_timeline":
            return _add(wb, source, field, sheet, caption, name,
                        top, left, width, height, _XL_TIMELINE)
        if action == "list":
            return _list(wb)
        if action == "delete":
            return _delete(wb, slicer_name)
        # action == "connect"
        return _connect(wb, slicer_name, pivots or [])


# ── Source resolver ────────────────────────────────────────────────────────────

def _resolve_source(wb, source_name: str):
    """Return the COM PivotTable or ListObject whose name matches *source_name*.

    Searches all sheets.  Raises ToolError if not found.
    """
    for i in range(1, wb.Sheets.Count + 1):
        ws = wb.Sheets(i)
        # Search PivotTables
        try:
            pt_count = ws.PivotTables().Count
            for j in range(1, pt_count + 1):
                pt = ws.PivotTables(j)
                if pt.Name == source_name:
                    return pt
        except Exception:
            pass
        # Search ListObjects (Tables)
        try:
            lo_count = ws.ListObjects.Count
            for j in range(1, lo_count + 1):
                lo = ws.ListObjects(j)
                if lo.Name == source_name:
                    return lo
        except Exception:
            pass

    raise ToolError(
        f"Source '{source_name}' not found. "
        "Provide a PivotTable name or Table (ListObject) name from this workbook."
    )


def _resolve_slicer_cache(wb, slicer_name: str):
    """Return the SlicerCache whose Name matches *slicer_name*."""
    try:
        return wb.SlicerCaches(slicer_name)
    except Exception:
        pass
    # Auto-generated names use "Slicer_<field>" — try to find by iterating
    try:
        count = wb.SlicerCaches.Count
        for i in range(1, count + 1):
            sc = wb.SlicerCaches(i)
            if sc.Name == slicer_name:
                return sc
    except Exception:
        pass
    raise ToolError(
        f"SlicerCache '{slicer_name}' not found. "
        "Use slicer action='list' to see available slicer names."
    )


# ── Action implementations ────────────────────────────────────────────────────

def _add(wb, source, field, sheet, caption, name,
         top, left, width, height, cache_type) -> dict:
    """Create a slicer (cache_type=xlSlicer) or timeline (cache_type=xlTimeline)."""
    if not source:
        raise ToolError("'source' is required: provide a PivotTable or Table name.")
    if not field:
        raise ToolError("'field' is required: provide the column/field name to filter on.")

    # Resolve destination sheet (active sheet if not given)
    ws_dest = _session.get_sheet(sheet, wb.Name)

    # Resolve source object
    src_obj = _resolve_source(wb, source)

    try:
        # Add2(Source, SourceField, Name, SlicerCacheType)
        # Name is optional — pass it only when explicitly supplied
        if name:
            sc = wb.SlicerCaches.Add2(src_obj, field, name, cache_type)
        else:
            sc = wb.SlicerCaches.Add2(src_obj, field, "", cache_type)

        # Add the visual slicer to the destination sheet.
        # Slicers.Add signature: (SlicerDestination, [Level], [Name], [Caption],
        #                         [Top], [Left], [Width], [Height])
        # CRITICAL: Level is OLAP-only.  For non-OLAP sources (Table / regular
        # PivotTable), Level must be completely OMITTED — passing "" or 1 raises
        # COM error E_INVALIDARG (-2147024809).  Use keyword args so pywin32
        # skips the Level slot without sending a VARIANT for it.
        slicer_caption = caption or field
        sl = sc.Slicers.Add(
            SlicerDestination=ws_dest,
            Caption=slicer_caption,
            Top=top,
            Left=left,
            Width=width,
            Height=height,
        )

        cache_type_str = "timeline" if cache_type == _XL_TIMELINE else "slicer"
        return {
            "added": cache_type_str,
            "slicer_cache": sc.Name,
            "slicer_name": sl.Name,
            "source": source,
            "field": field,
            "sheet": ws_dest.Name,
            "caption": slicer_caption,
            "position": {"top": top, "left": left, "width": width, "height": height},
        }
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"add {'timeline' if cache_type == _XL_TIMELINE else 'slicer'} failed")


def _list(wb) -> dict:
    """List all slicer caches in the workbook."""
    try:
        result = []
        count = wb.SlicerCaches.Count
        for i in range(1, count + 1):
            sc = wb.SlicerCaches(i)
            # SlicerCacheType: 1=xlSlicer, 2=xlTimeline
            try:
                sc_type = sc.SlicerCacheType
                type_str = "timeline" if sc_type == _XL_TIMELINE else "slicer"
            except Exception:
                type_str = "unknown"
            try:
                src_name = sc.SourceName
            except Exception:
                src_name = ""
            result.append({
                "name": sc.Name,
                "source": src_name,
                "slicer_type": type_str,
            })
        return {"slicer_caches": result, "count": len(result)}
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "list slicers failed")


def _delete(wb, slicer_name: str | None) -> dict:
    """Delete a SlicerCache (removes the cache and all associated slicers)."""
    if not slicer_name:
        raise ToolError("'slicer' (slicer cache name) is required for delete.")
    sc = _resolve_slicer_cache(wb, slicer_name)
    try:
        deleted_name = sc.Name
        sc.Delete()
        return {"deleted": deleted_name}
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"delete slicer cache '{slicer_name}' failed")


def _connect(wb, slicer_name: str | None, pivot_names: list[str]) -> dict:
    """Connect a slicer cache to one or more additional PivotTables.

    All PivotTables sharing the same PivotCache can be connected to the same
    slicer.  Calling AddPivotTable on a pivot that is already connected is a
    no-op (Excel silently ignores it).
    """
    if not slicer_name:
        raise ToolError("'slicer' (slicer cache name) is required for connect.")
    if not pivot_names:
        raise ToolError("'pivots' list is required for connect (list of PivotTable names).")

    sc = _resolve_slicer_cache(wb, slicer_name)
    connected = []
    errors = []
    pvts = sc.PivotTables

    for pname in pivot_names:
        try:
            pt_obj = _resolve_source(wb, pname)
            pvts.AddPivotTable(pt_obj)
            connected.append(pname)
        except Exception as e:
            errors.append({"pivot": pname, "error": str(e)})

    return {
        "connected": connected,
        "slicer_cache": slicer_name,
        "errors": errors,
    }
