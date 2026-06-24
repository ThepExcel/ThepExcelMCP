"""Cell comment (note + threaded) operations.

Two separate comment systems in modern Excel
--------------------------------------------
kind="note"     — Legacy yellow sticky-note. Stored in ws.Comments collection.
                  Range property: cell.Comment (None when absent).
                  Add:    cell.AddComment(text)   — raises 1004 if one exists → delete first.
                  Edit:   cell.Comment.Text(text) — METHOD (call it), replaces text.
                  Read:   cell.Comment.Text()      — METHOD with no args returns the string.
                  Author: cell.Comment.Author      — STRING property (not an object).
                  Delete: cell.Comment.Delete().

kind="threaded" — Modern threaded comment. Stored in ws.CommentsThreaded collection.
                  Range property: cell.CommentThreaded (None when absent).
                  Add:    cell.AddCommentThreaded(text).
                  Edit:   cell.CommentThreaded.Text(text) — METHOD, same signature as note.
                  Read:   cell.CommentThreaded.Text()     — METHOD with no args.
                  Author: cell.CommentThreaded.Author.Name — Author OBJECT, use .Name.
                  Reply:  cell.CommentThreaded.AddReply(text).
                  Delete: cell.CommentThreaded.Delete().

GOTCHAS
-------
- cell.AddComment raises COM 1004 when a note already exists → always Delete first.
- cell.Comment is None (not an error) when no note is present — guard with `is not None`.
- cell.CommentThreaded is None (VBA Nothing) when no threaded comment is present.
- Comment.Text() and CommentThreaded.Text() are METHODS, not properties — always call them.
- CommentThreaded.Author is an object; CommentThreaded.Author.Name is the string.
- Comment.Author is already a string (no .Name needed).
- Threaded comment text is not editable via COM (no direct text-property setter) beyond
  the Text() method; replies cannot be edited via COM — inform caller to use AddReply.
- Wrap ALL mutations in excel_guard (suppresses modal dialogs).
"""

from __future__ import annotations

from fastmcp.exceptions import ToolError

from ..session import ExcelSession, excel_guard

_session = ExcelSession()

_VALID_KINDS = {"note", "threaded"}
_VALID_ACTIONS = {"add", "edit", "reply", "delete", "list", "get"}


# ── Public entry point ─────────────────────────────────────────────────────────

def comment_action(
    action: str,
    cell: str | None = None,
    sheet: str | None = None,
    workbook: str | None = None,
    text: str | None = None,
    kind: str = "note",
) -> dict:
    """Dispatch a cell-comment action.

    Parameters
    ----------
    action   : add | edit | reply | delete | list | get
    cell     : Single-cell address (e.g. "A1", "B5"). Required for all actions
               except 'list'.
    sheet    : Sheet name; None = active sheet.
    workbook : Workbook name; None = active workbook.
    text     : Comment text (required for add, edit, reply).
    kind     : "note" (default) — legacy yellow sticky-note (ws.Comments)
               "threaded"      — modern threaded comment (ws.CommentsThreaded)
               "all"           — only for 'list' action (returns both collections)

    Actions
    -------
    add      Add a new note or threaded comment.
    edit     Replace text of an existing note (kind=note only;
             threaded text replacement → use 'add' to delete+re-add).
    reply    Add a reply to a threaded comment (kind=threaded only).
    delete   Delete the note or threaded comment (no-op-safe: raises if absent).
    list     List all notes and/or threaded comments on the sheet.
    get      Read the note and/or threaded comment on a single cell.
    """
    if action not in _VALID_ACTIONS:
        raise ToolError(
            f"Unknown action '{action}'. Valid: {', '.join(sorted(_VALID_ACTIONS))}."
        )
    if kind not in {*_VALID_KINDS, "all"}:
        raise ToolError(
            f"Unknown kind '{kind}'. Valid: note, threaded, all (list-only)."
        )
    if action != "list" and not cell:
        raise ToolError(
            f"action='{action}' requires a 'cell' parameter (e.g. cell='A1')."
        )
    return _session.run_com(
        _dispatch,
        action, cell, sheet, workbook, text, kind,
    )


# ── Worker-thread dispatcher ───────────────────────────────────────────────────

def _dispatch(
    action: str,
    cell: str | None,
    sheet: str | None,
    workbook: str | None,
    text: str | None,
    kind: str,
) -> dict:
    """Executed on the STA COM worker thread."""
    ws = _session.get_sheet(sheet, workbook)
    app = ws.Application

    if action == "list":
        with excel_guard(app):
            return _list(ws, kind)

    # Resolve the single-cell range for all non-list actions
    try:
        rng = ws.Range(cell)
    except Exception as e:
        raise ToolError(f"Invalid cell reference '{cell}': {e}")

    with excel_guard(app):
        if action == "add":
            return _add(rng, text, kind)
        if action == "edit":
            return _edit(rng, text, kind)
        if action == "reply":
            return _reply(rng, text, kind)
        if action == "delete":
            return _delete(rng, kind)
        # action == "get"
        return _get(rng, kind)


# ── Action implementations ────────────────────────────────────────────────────

def _add(rng, text: str | None, kind: str) -> dict:
    """Add a note or threaded comment to a cell.

    For notes: delete-before-add (COM 1004 guard).
    For threaded: AddCommentThreaded raises if one exists → same guard.
    """
    if text is None:
        raise ToolError("add action requires 'text'.")

    cell_addr = rng.Address

    try:
        if kind == "note":
            # Delete-before-add: AddComment raises 1004 if a comment already exists.
            if rng.Comment is not None:
                rng.Comment.Delete()
            rng.AddComment(text)
            # VERIFY EFFECT: read back the text
            read_back = rng.Comment.Text()
            return {
                "comment": "add",
                "kind": "note",
                "cell": cell_addr,
                "applied": {"text": read_back},
            }
        else:  # threaded
            if rng.CommentThreaded is not None:
                rng.CommentThreaded.Delete()
            rng.AddCommentThreaded(text)
            # VERIFY EFFECT: read back
            read_back = rng.CommentThreaded.Text()
            return {
                "comment": "add",
                "kind": "threaded",
                "cell": cell_addr,
                "applied": {"text": read_back},
            }
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"add {kind} comment on {cell_addr} failed")


def _edit(rng, text: str | None, kind: str) -> dict:
    """Replace the text of an existing note.

    kind=note only — threaded comment text replacement is not supported via COM
    (the Text() method changes the comment but there is no reliable way to edit
    a threaded comment while preserving author/date metadata). Use add (which
    deletes and re-creates) or reply for additional context.
    """
    if kind == "threaded":
        raise ToolError(
            "edit is not supported for kind='threaded' via COM. "
            "Use action='add' to replace the threaded comment (delete + re-add), "
            "or action='reply' to add a reply."
        )
    if text is None:
        raise ToolError("edit action requires 'text'.")

    cell_addr = rng.Address
    try:
        if rng.Comment is None:
            raise ToolError(
                f"No note found on cell {cell_addr}. "
                "Use action='add' to create one first."
            )
        # Comment.Text() is a METHOD that replaces text when called with an argument.
        rng.Comment.Text(text)
        # VERIFY EFFECT: read back
        read_back = rng.Comment.Text()
        return {
            "comment": "edit",
            "kind": "note",
            "cell": cell_addr,
            "applied": {"text": read_back},
        }
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"edit note on {cell_addr} failed")


def _reply(rng, text: str | None, kind: str) -> dict:
    """Add a reply to a threaded comment.

    kind=threaded only — legacy notes do not support threading.
    """
    if kind == "note":
        raise ToolError(
            "reply is not supported for kind='note'. "
            "Legacy notes do not have threading. "
            "Use kind='threaded' with action='add' to create a threaded comment."
        )
    if text is None:
        raise ToolError("reply action requires 'text'.")

    cell_addr = rng.Address
    try:
        ct = rng.CommentThreaded
        if ct is None:
            raise ToolError(
                f"No threaded comment found on cell {cell_addr}. "
                "Use action='add' with kind='threaded' to create one first."
            )
        ct.AddReply(text)
        # VERIFY EFFECT: count replies to confirm the reply was appended
        reply_count = ct.Replies.Count
        return {
            "comment": "reply",
            "kind": "threaded",
            "cell": cell_addr,
            "applied": {"text": text, "reply_count": reply_count},
        }
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"reply to threaded comment on {cell_addr} failed")


def _delete(rng, kind: str) -> dict:
    """Delete the note or threaded comment on a cell.

    Raises an actionable ToolError if no comment is present (never a silent no-op).
    """
    cell_addr = rng.Address
    try:
        if kind == "note":
            if rng.Comment is None:
                raise ToolError(
                    f"No note found on cell {cell_addr} to delete."
                )
            rng.Comment.Delete()
            # VERIFY EFFECT: confirm the comment is actually gone.
            if rng.Comment is not None:
                raise ToolError(
                    f"Delete() called on note at {cell_addr} but comment still present. "
                    "The delete may have been blocked by a protection setting."
                )
            return {
                "comment": "delete",
                "kind": "note",
                "cell": cell_addr,
                "applied": {"deleted": True},
            }
        else:  # threaded
            if rng.CommentThreaded is None:
                raise ToolError(
                    f"No threaded comment found on cell {cell_addr} to delete."
                )
            rng.CommentThreaded.Delete()
            # VERIFY EFFECT: confirm the threaded comment is actually gone.
            if rng.CommentThreaded is not None:
                raise ToolError(
                    f"Delete() called on threaded comment at {cell_addr} but comment still present. "
                    "The delete may have been blocked by a protection setting."
                )
            return {
                "comment": "delete",
                "kind": "threaded",
                "cell": cell_addr,
                "applied": {"deleted": True},
            }
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"delete {kind} comment on {cell_addr} failed")


def _list(ws, kind: str) -> dict:
    """Enumerate notes and/or threaded comments on the sheet.

    Returns a list of dicts: {cell, text, author}.
    For notes: comment.Author is a string.
    For threaded: ct.Author.Name is the string (Author is an object).
    """
    try:
        results: list[dict] = []

        if kind in ("note", "all"):
            for cmt in ws.Comments:
                try:
                    results.append({
                        "kind": "note",
                        "cell": cmt.Parent.Address,
                        "text": cmt.Text(),
                        "author": cmt.Author,
                    })
                except Exception:
                    # Skip any malformed comment rather than crashing the whole list
                    pass

        if kind in ("threaded", "all"):
            for ct in ws.CommentsThreaded:
                try:
                    results.append({
                        "kind": "threaded",
                        "cell": ct.Parent.Address,
                        "text": ct.Text(),
                        "author": ct.Author.Name,
                    })
                except Exception:
                    pass

        return {
            "comment": "list",
            "sheet": ws.Name,
            "kind": kind,
            "comments": results,
            "count": len(results),
        }
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "list comments failed")


def _get(rng, kind: str) -> dict:
    """Read note and/or threaded comment from a single cell.

    Returns text + author for each kind present. Missing comments are
    represented as None in the result (not an error).
    """
    cell_addr = rng.Address
    try:
        result: dict = {
            "comment": "get",
            "cell": cell_addr,
        }

        if kind in ("note", "all"):
            cmt = rng.Comment
            if cmt is not None:
                result["note"] = {
                    "text": cmt.Text(),
                    "author": cmt.Author,
                }
            else:
                result["note"] = None

        if kind in ("threaded", "all"):
            ct = rng.CommentThreaded
            if ct is not None:
                result["threaded"] = {
                    "text": ct.Text(),
                    "author": ct.Author.Name,
                }
            else:
                result["threaded"] = None

        return result
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, f"get comment on {cell_addr} failed")
