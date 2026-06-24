"""Workbook snapshots — non-destructive on-disk safety copies.

This domain lets an agent create a point-in-time COPY of a workbook on disk
and later re-OPEN that copy as a SEPARATE new workbook. It is deliberately
SAFE: there is NO code path that closes, overwrites, reverts, or rebinds the
user's live workbook.

Why SaveCopyAs (the safe primitive)
-----------------------------------
``wb.SaveCopyAs(path)`` writes a copy of the workbook to *path* WITHOUT
changing the open workbook's dirty (``Saved``) flag, its ``Name``, or its
``FullName``/path. The open file the user is editing is untouched — Excel
simply streams a copy to disk. This is fundamentally different from
``SaveAs``, which REBINDS the open workbook to the new path (changing Name +
path) and would silently move the user's live editing target. We never use
SaveAs here. The output format is inferred by Excel from the path extension,
so we preserve the source extension (.xlsx / .xlsm / .xlsb) — keeping macros
intact for .xlsm.

restore is non-destructive
--------------------------
``restore`` calls ``app.Workbooks.Open(path)`` on the snapshot copy. That
opens the copy as an ADDITIONAL workbook alongside everything already open.
The user's original workbook is never closed or modified — to "revert" you
copy what you need out of the freshly opened snapshot copy by hand.

Snapshots are tracked in a module-level, session-scoped registry
(``_SNAPSHOTS``): it lives for the life of the MCP server process only and is
not persisted. ``list`` reports each entry plus whether the file still exists
on disk.
"""

from __future__ import annotations

import datetime
import os
import tempfile

from fastmcp.exceptions import ToolError

from ..session import ExcelSession, excel_guard

_session = ExcelSession()

# Session-scoped snapshot registry: id → {path, workbook, created, format}.
# Lives for the MCP server process lifetime only; not persisted to disk.
_SNAPSHOTS: dict[str, dict] = {}

# Monotonic counter for snapshot id generation.
_counter = 0

# Extensions Excel can round-trip via SaveCopyAs (format inferred from path).
_KNOWN_EXTS = (".xlsx", ".xlsm", ".xlsb")


def _snapshot_dir() -> str:
    """Return (creating if needed) the snapshot temp directory."""
    base = tempfile.gettempdir() or os.environ.get("TEMP") or "."
    snap_dir = os.path.join(base, "thepexcel_mcp", "snapshots")
    os.makedirs(snap_dir, exist_ok=True)
    return snap_dir


def _wb_stem_and_ext(wb) -> tuple[str, str]:
    """Derive (stem, ext) from wb.Name, preserving a known macro-capable ext.

    Defaults to .xlsx when the workbook has no extension (e.g. a never-saved
    "Book1"). Unknown extensions fall back to .xlsx so the copy is always a
    valid Excel file.
    """
    name = wb.Name or "Workbook"
    stem, ext = os.path.splitext(name)
    ext = ext.lower()
    if ext not in _KNOWN_EXTS:
        ext = ".xlsx"
    if not stem:
        stem = "Workbook"
    return stem, ext


# ── Public entry point ─────────────────────────────────────────────────────────

def snapshot_action(
    action: str,
    workbook: str | None = None,
    snapshot_id: str | None = None,
) -> dict:
    """Dispatch a workbook-snapshot action.

    SAFETY: This tool NEVER closes, overwrites, or reverts the workbook you are
    editing. ``snapshot`` only writes an on-disk copy; ``restore`` only OPENS
    that copy as a new, separate workbook. There is no in-place / reopen path.

    Actions
    -------
    snapshot
        Create a point-in-time copy of *workbook* (active when None) on disk via
        ``wb.SaveCopyAs``. The live workbook's name/path/dirty-flag are NOT
        changed. The source extension is preserved (so .xlsm keeps macros).
        Returns {id, path, workbook, size_bytes, created}.
    list
        Return every registered snapshot with {id, workbook, path, created,
        exists, size_bytes}. Read-only.
    restore
        Open the snapshot identified by *snapshot_id* as a NEW, separate
        workbook (``Workbooks.Open``) — your original workbook is left exactly
        as it is. Copy what you need out of the restored copy by hand. Requires
        *snapshot_id*.
    delete
        Remove the snapshot file from disk and drop it from the registry.
        Requires *snapshot_id*.
    """
    valid = {"snapshot", "list", "restore", "delete"}
    if action not in valid:
        raise ToolError(
            f"Unknown action '{action}'. Valid: {', '.join(sorted(valid))}."
        )
    if action in ("restore", "delete") and not snapshot_id:
        raise ToolError(f"action='{action}' requires the 'snapshot_id' parameter.")
    return _session.run_com(_dispatch, action, workbook, snapshot_id)


# ── Worker-thread dispatcher ───────────────────────────────────────────────────

def _dispatch(action: str, workbook: str | None, snapshot_id: str | None) -> dict:
    """Executed on the STA COM worker thread."""
    if action == "snapshot":
        return _snapshot(workbook)
    if action == "list":
        return _list()
    if action == "restore":
        return _restore(snapshot_id)
    return _delete(snapshot_id)  # action == "delete"


# ── Action implementations ─────────────────────────────────────────────────────

def _snapshot(workbook: str | None) -> dict:
    """Write an on-disk copy of the workbook via SaveCopyAs (non-destructive)."""
    global _counter
    wb = _session.get_workbook(workbook)
    stem, ext = _wb_stem_and_ext(wb)

    _counter += 1
    snap_id = f"snap_{_counter}_{stem}"
    path = os.path.join(_snapshot_dir(), f"{stem}__{snap_id}{ext}")

    try:
        with excel_guard(wb.Application):
            # SaveCopyAs streams a copy to disk WITHOUT touching the open
            # workbook's Saved flag, Name, or path — the user's live file is
            # never rebound or modified.
            wb.SaveCopyAs(path)
    except Exception as e:
        raise _session.wrap(e, f"snapshot: SaveCopyAs to '{path}' failed")

    # VERIFY EFFECT: the copy must exist and be non-empty on disk.
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        raise ToolError(
            f"snapshot: SaveCopyAs reported no error but the copy at '{path}' "
            f"is missing or empty. Check the temp path is writable."
        )

    created = datetime.datetime.now().isoformat()
    size = os.path.getsize(path)
    _SNAPSHOTS[snap_id] = {
        "path": path,
        "workbook": wb.Name,
        "created": created,
        "format": ext,
    }
    return {
        "id": snap_id,
        "path": path,
        "workbook": wb.Name,
        "size_bytes": size,
        "created": created,
    }


def _list() -> dict:
    """Return every registered snapshot. Read-only — no COM, no mutation."""
    snapshots = []
    for snap_id, meta in _SNAPSHOTS.items():
        path = meta["path"]
        exists = os.path.exists(path)
        snapshots.append(
            {
                "id": snap_id,
                "workbook": meta["workbook"],
                "path": path,
                "created": meta["created"],
                "exists": exists,
                "size_bytes": os.path.getsize(path) if exists else 0,
            }
        )
    return {"snapshots": snapshots, "count": len(snapshots)}


def _restore(snapshot_id: str) -> dict:
    """Open the snapshot copy as a SEPARATE new workbook.

    NON-DESTRUCTIVE: this does NOT revert, overwrite, or close the user's
    current workbook. It calls Workbooks.Open on the snapshot file, which adds
    a new workbook alongside everything already open. To "restore" content, the
    caller copies what they need out of the opened copy manually.
    """
    meta = _SNAPSHOTS.get(snapshot_id)
    if meta is None:
        raise ToolError(
            f"restore: unknown snapshot_id '{snapshot_id}'. "
            f"Known: {list(_SNAPSHOTS) or ['(none)']}."
        )
    path = meta["path"]
    if not os.path.exists(path):
        raise ToolError(
            f"restore: snapshot file for '{snapshot_id}' is missing on disk "
            f"('{path}'). It may have been deleted externally."
        )

    app = _session.get_app()
    try:
        with excel_guard(app):
            opened = app.Workbooks.Open(path)
    except Exception as e:
        raise _session.wrap(e, f"restore: Workbooks.Open('{path}') failed")

    opened_name = opened.Name

    # VERIFY EFFECT: the opened workbook is now present in the Workbooks set.
    present = any(
        app.Workbooks(i + 1).Name == opened_name
        for i in range(app.Workbooks.Count)
    )
    if not present:
        raise ToolError(
            f"restore: Workbooks.Open reported success but '{opened_name}' "
            f"is not present in the open workbooks set."
        )

    return {
        "restored_from": snapshot_id,
        "opened_workbook": opened_name,
        "path": path,
        "note": (
            "Opened as a NEW workbook — your original workbook was not "
            "modified. Copy what you need from the restored copy."
        ),
    }


def _delete(snapshot_id: str) -> dict:
    """Remove the snapshot file from disk and drop it from the registry."""
    meta = _SNAPSHOTS.get(snapshot_id)
    if meta is None:
        raise ToolError(
            f"delete: unknown snapshot_id '{snapshot_id}'. "
            f"Known: {list(_SNAPSHOTS) or ['(none)']}."
        )
    path = meta["path"]
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception as e:
        raise _session.wrap(e, f"delete: could not remove '{path}'")

    _SNAPSHOTS.pop(snapshot_id, None)

    # VERIFY EFFECT: the file must be gone from disk.
    if os.path.exists(path):
        raise ToolError(
            f"delete: os.remove reported no error but '{path}' still exists."
        )

    return {"deleted": snapshot_id}
