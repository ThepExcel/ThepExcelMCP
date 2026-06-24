"""Unit tests for excel_comment (comments.py).

No Excel required — all COM calls are intercepted via make_mock_session.

VERIFY-EFFECT discipline: every test asserts the SPECIFIC COM method/property
that was called with the SPECIFIC value, not just "no exception raised".

Key COM facts confirmed from Microsoft Learn docs:
- Comment.Text()     — a METHOD (call with no args to read, with text arg to set)
- Comment.Author     — a STRING property (legacy notes)
- CommentThreaded.Text()  — also a METHOD
- CommentThreaded.Author  — an OBJECT; use .Author.Name for the string
- cell.Comment             — None when no legacy note is present
- cell.CommentThreaded     — None when no threaded comment is present
- cell.AddComment(text)    — raises 1004 if note already exists → delete first
- AddCommentThreaded(text) — also raises if threaded comment exists → delete first
"""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock, patch, call

from fastmcp.exceptions import ToolError
from conftest import make_mock_session


# ── Helpers ────────────────────────────────────────────────────────────────────

def _make_sheet_mock(name: str = "Sheet1"):
    ws = MagicMock()
    ws.Name = name
    return ws


def _make_note_mock(text: str = "Hello", author: str = "Alice"):
    """Mock a legacy Comment (note) object."""
    cmt = MagicMock()
    cmt.Text.return_value = text
    cmt.Author = author           # Author is a plain string for legacy notes
    # Parent.Address simulates comment.Parent.Address in list
    cmt.Parent = MagicMock()
    cmt.Parent.Address = "$A$1"
    return cmt


def _make_threaded_mock(text: str = "Thread text", author_name: str = "Bob"):
    """Mock a CommentThreaded object."""
    ct = MagicMock()
    ct.Text.return_value = text
    ct.Author = MagicMock()
    ct.Author.Name = author_name  # Author is an object; .Name is the string
    ct.Parent = MagicMock()
    ct.Parent.Address = "$B$2"
    ct.Replies = MagicMock()
    ct.Replies.Count = 0
    return ct


def _call_comment(action, mock_session, mock_ws, **kwargs):
    """Patch _session and excel_guard, invoke comment_action."""
    mock_session.get_sheet.return_value = mock_ws
    mock_ws.Application = MagicMock()

    with patch("thepexcel_mcp.domains.comments._session", mock_session):
        with patch("thepexcel_mcp.domains.comments.excel_guard") as eg:
            eg.return_value.__enter__ = MagicMock(return_value=None)
            eg.return_value.__exit__ = MagicMock(return_value=False)
            from thepexcel_mcp.domains.comments import comment_action
            return comment_action(action, **kwargs)


# ── Action validation ──────────────────────────────────────────────────────────

class TestActionValidation:
    def test_unknown_action_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="Unknown action"):
            _call_comment("bogus", ms, ws, cell="A1")

    def test_unknown_kind_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="Unknown kind"):
            _call_comment("add", ms, ws, cell="A1", text="hi", kind="badkind")

    def test_missing_cell_raises_for_add(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="cell"):
            _call_comment("add", ms, ws, text="hi")  # no cell

    def test_missing_cell_raises_for_delete(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="cell"):
            _call_comment("delete", ms, ws)  # no cell

    def test_missing_cell_ok_for_list(self):
        """list does not require a cell parameter."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Comments = []
        ws.CommentsThreaded = []
        # Should not raise for list without cell
        result = _call_comment("list", ms, ws, kind="all")
        assert result["comment"] == "list"

    def test_invalid_cell_reference_raises(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Range.side_effect = Exception("Bad ref")
        with pytest.raises(ToolError, match="Invalid cell reference"):
            _call_comment("add", ms, ws, cell="ZZZBAD", text="hi")

    def test_add_requires_text(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rng = MagicMock()
        rng.Address = "$A$1"
        rng.Comment = None
        ws.Range.return_value = rng
        with pytest.raises(ToolError, match="text"):
            _call_comment("add", ms, ws, cell="A1")  # text=None

    def test_edit_requires_text(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        cmt = _make_note_mock()
        rng = MagicMock()
        rng.Address = "$A$1"
        rng.Comment = cmt
        ws.Range.return_value = rng
        with pytest.raises(ToolError, match="text"):
            _call_comment("edit", ms, ws, cell="A1", kind="note")  # text=None

    def test_reply_requires_text(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ct = _make_threaded_mock()
        rng = MagicMock()
        rng.Address = "$B$2"
        rng.CommentThreaded = ct
        ws.Range.return_value = rng
        with pytest.raises(ToolError, match="text"):
            _call_comment("reply", ms, ws, cell="B2", kind="threaded")  # text=None


# ── add (note) ────────────────────────────────────────────────────────────────

class TestAddNote:
    def test_add_note_calls_addcomment(self):
        """AddComment must be called with the text; verify-effect reads back."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        cmt = _make_note_mock("My note")
        rng = MagicMock()
        rng.Address = "$A$1"
        rng.Comment = None          # no existing note
        ws.Range.return_value = rng
        # After AddComment, rng.Comment should return cmt
        def _add_comment_side_effect(text):
            rng.Comment = cmt
        rng.AddComment.side_effect = _add_comment_side_effect

        result = _call_comment("add", ms, ws, cell="A1", text="My note", kind="note")

        rng.AddComment.assert_called_once_with("My note")
        # VERIFY EFFECT: read-back via Comment.Text()
        assert result["comment"] == "add"
        assert result["kind"] == "note"
        assert result["applied"]["text"] == "My note"

    def test_add_note_deletes_existing_before_add(self):
        """If a note already exists, Delete() must be called before AddComment."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        existing_cmt = _make_note_mock("old text")
        new_cmt = _make_note_mock("new text")
        rng = MagicMock()
        rng.Address = "$A$1"
        rng.Comment = existing_cmt  # note already exists
        ws.Range.return_value = rng

        deleted = []
        def _delete():
            deleted.append(True)
            rng.Comment = new_cmt   # after delete, replace with new mock
        existing_cmt.Delete.side_effect = _delete

        def _add_comment_side_effect(text):
            pass
        rng.AddComment.side_effect = _add_comment_side_effect

        _call_comment("add", ms, ws, cell="A1", text="new text", kind="note")

        # Delete must have been called before AddComment
        assert deleted, "existing note Delete() was not called"
        existing_cmt.Delete.assert_called_once()
        rng.AddComment.assert_called_once_with("new text")

    def test_add_note_result_shape(self):
        """Result must carry comment, kind, cell, applied keys."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        cmt = _make_note_mock("Test")
        rng = MagicMock()
        rng.Address = "$C$3"
        rng.Comment = None
        ws.Range.return_value = rng

        def _side(text):
            rng.Comment = cmt
        rng.AddComment.side_effect = _side

        result = _call_comment("add", ms, ws, cell="C3", text="Test", kind="note")

        assert result["comment"] == "add"
        assert result["kind"] == "note"
        assert result["cell"] == "$C$3"
        assert "applied" in result
        assert "text" in result["applied"]


# ── add (threaded) ────────────────────────────────────────────────────────────

class TestAddThreaded:
    def test_add_threaded_calls_addcommentthreaded(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ct = _make_threaded_mock("Threaded note")
        rng = MagicMock()
        rng.Address = "$B$2"
        rng.CommentThreaded = None  # no existing threaded comment
        ws.Range.return_value = rng

        def _add_ct(text):
            rng.CommentThreaded = ct
        rng.AddCommentThreaded.side_effect = _add_ct

        result = _call_comment("add", ms, ws, cell="B2", text="Threaded note", kind="threaded")

        rng.AddCommentThreaded.assert_called_once_with("Threaded note")
        assert result["kind"] == "threaded"
        assert result["applied"]["text"] == "Threaded note"

    def test_add_threaded_deletes_existing_before_add(self):
        """Existing threaded comment must be deleted before AddCommentThreaded."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        existing_ct = _make_threaded_mock("old")
        new_ct = _make_threaded_mock("new")
        rng = MagicMock()
        rng.Address = "$B$2"
        rng.CommentThreaded = existing_ct
        ws.Range.return_value = rng

        deleted = []
        def _delete():
            deleted.append(True)
            rng.CommentThreaded = new_ct
        existing_ct.Delete.side_effect = _delete
        rng.AddCommentThreaded.return_value = None

        _call_comment("add", ms, ws, cell="B2", text="new", kind="threaded")

        assert deleted, "existing threaded comment Delete() was not called"
        rng.AddCommentThreaded.assert_called_once_with("new")

    def test_add_threaded_result_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ct = _make_threaded_mock("msg")
        rng = MagicMock()
        rng.Address = "$D$4"
        rng.CommentThreaded = None
        ws.Range.return_value = rng

        def _side(text):
            rng.CommentThreaded = ct
        rng.AddCommentThreaded.side_effect = _side

        result = _call_comment("add", ms, ws, cell="D4", text="msg", kind="threaded")

        assert result["comment"] == "add"
        assert result["kind"] == "threaded"
        assert result["cell"] == "$D$4"
        assert result["applied"]["text"] == "msg"


# ── edit ──────────────────────────────────────────────────────────────────────

class TestEdit:
    def test_edit_note_calls_text_method_with_arg(self):
        """Comment.Text(text) must be called to set; then read back with Text()."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        cmt = _make_note_mock("original")
        rng = MagicMock()
        rng.Address = "$A$1"
        rng.Comment = cmt
        ws.Range.return_value = rng
        # After Text("new"), the mock returns the new value
        cmt.Text.return_value = "new text"

        result = _call_comment("edit", ms, ws, cell="A1", text="new text", kind="note")

        # VERIFY EFFECT: Text() called twice — once to set, once to read
        assert cmt.Text.call_count == 2
        # First call had the text arg (set)
        cmt.Text.assert_any_call("new text")
        # Second call had no args (read-back)
        cmt.Text.assert_any_call()
        assert result["comment"] == "edit"
        assert result["applied"]["text"] == "new text"

    def test_edit_note_raises_when_no_comment(self):
        """Editing a non-existent note must raise ToolError, not crash."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rng = MagicMock()
        rng.Address = "$A$1"
        rng.Comment = None
        ws.Range.return_value = rng

        with pytest.raises(ToolError, match="No note found"):
            _call_comment("edit", ms, ws, cell="A1", text="something", kind="note")

    def test_edit_threaded_raises_unsupported(self):
        """edit action for kind=threaded must raise a clear ToolError."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="not supported for kind='threaded'"):
            _call_comment("edit", ms, ws, cell="A1", text="x", kind="threaded")


# ── reply ─────────────────────────────────────────────────────────────────────

class TestReply:
    def test_reply_threaded_calls_addreply(self):
        """AddReply(text) must be called; verify-effect checks Replies.Count.

        Count starts at 0; AddReply side-effect increments it to 1 so that
        the assertion proves AddReply actually ran (not a pre-set stub value).
        """
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ct = _make_threaded_mock()
        ct.Replies.Count = 0  # starts at zero — before any reply
        rng = MagicMock()
        rng.Address = "$B$2"
        rng.CommentThreaded = ct
        ws.Range.return_value = rng

        # Side effect: AddReply increments Count, proving the call happened.
        def _add_reply_side_effect(text):
            ct.Replies.Count = 1
        ct.AddReply.side_effect = _add_reply_side_effect

        result = _call_comment("reply", ms, ws, cell="B2", text="my reply", kind="threaded")

        ct.AddReply.assert_called_once_with("my reply")
        # VERIFY EFFECT: reply_count must reflect the post-AddReply Count (1, not 0)
        assert result["comment"] == "reply"
        assert result["kind"] == "threaded"
        assert result["applied"]["text"] == "my reply"
        assert result["applied"]["reply_count"] == 1

    def test_reply_threaded_raises_when_no_comment(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rng = MagicMock()
        rng.Address = "$B$2"
        rng.CommentThreaded = None
        ws.Range.return_value = rng

        with pytest.raises(ToolError, match="No threaded comment found"):
            _call_comment("reply", ms, ws, cell="B2", text="reply", kind="threaded")

    def test_reply_note_raises_unsupported(self):
        """reply action for kind=note must raise a clear ToolError."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        with pytest.raises(ToolError, match="not supported for kind='note'"):
            _call_comment("reply", ms, ws, cell="A1", text="x", kind="note")


# ── delete ────────────────────────────────────────────────────────────────────

class TestDelete:
    def test_delete_note_calls_delete(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        cmt = _make_note_mock()
        rng = MagicMock()
        rng.Address = "$A$1"
        rng.Comment = cmt
        ws.Range.return_value = rng

        # Side effect: Delete() clears rng.Comment — simulates COM removing the comment.
        def _delete_side_effect():
            rng.Comment = None
        cmt.Delete.side_effect = _delete_side_effect

        result = _call_comment("delete", ms, ws, cell="A1", kind="note")

        cmt.Delete.assert_called_once()
        # VERIFY EFFECT: deleted=True only after confirming Comment is None.
        assert result["comment"] == "delete"
        assert result["kind"] == "note"
        assert result["applied"]["deleted"] is True

    def test_delete_note_raises_when_absent(self):
        """Deleting a non-existent note must raise ToolError, not silently no-op."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rng = MagicMock()
        rng.Address = "$A$1"
        rng.Comment = None
        ws.Range.return_value = rng

        with pytest.raises(ToolError, match="No note found"):
            _call_comment("delete", ms, ws, cell="A1", kind="note")

    def test_delete_threaded_calls_delete(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ct = _make_threaded_mock()
        rng = MagicMock()
        rng.Address = "$B$2"
        rng.CommentThreaded = ct
        ws.Range.return_value = rng

        # Side effect: Delete() clears rng.CommentThreaded — simulates COM removing the comment.
        def _delete_side_effect():
            rng.CommentThreaded = None
        ct.Delete.side_effect = _delete_side_effect

        result = _call_comment("delete", ms, ws, cell="B2", kind="threaded")

        ct.Delete.assert_called_once()
        # VERIFY EFFECT: deleted=True only after confirming CommentThreaded is None.
        assert result["comment"] == "delete"
        assert result["kind"] == "threaded"
        assert result["applied"]["deleted"] is True

    def test_delete_threaded_raises_when_absent(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rng = MagicMock()
        rng.Address = "$B$2"
        rng.CommentThreaded = None
        ws.Range.return_value = rng

        with pytest.raises(ToolError, match="No threaded comment found"):
            _call_comment("delete", ms, ws, cell="B2", kind="threaded")


# ── list ──────────────────────────────────────────────────────────────────────

class TestList:
    def _make_list_note(self, addr="$A$1", text="Note text", author="Alice"):
        cmt = MagicMock()
        cmt.Text.return_value = text
        cmt.Author = author
        cmt.Parent = MagicMock()
        cmt.Parent.Address = addr
        return cmt

    def _make_list_ct(self, addr="$B$2", text="Thread", author_name="Bob"):
        ct = MagicMock()
        ct.Text.return_value = text
        ct.Author = MagicMock()
        ct.Author.Name = author_name
        ct.Parent = MagicMock()
        ct.Parent.Address = addr
        return ct

    def test_list_notes_returns_notes(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("MySheet")
        cmt = self._make_list_note("$A$1", "Hello", "Alice")
        ws.Comments = [cmt]
        ws.CommentsThreaded = []

        result = _call_comment("list", ms, ws, kind="note")

        assert result["comment"] == "list"
        assert result["kind"] == "note"
        assert result["sheet"] == "MySheet"
        assert result["count"] == 1
        notes = result["comments"]
        assert notes[0]["kind"] == "note"
        assert notes[0]["cell"] == "$A$1"
        assert notes[0]["text"] == "Hello"
        assert notes[0]["author"] == "Alice"

    def test_list_threaded_returns_threaded(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ct = self._make_list_ct("$C$3", "Thread msg", "Bob")
        ws.Comments = []
        ws.CommentsThreaded = [ct]

        result = _call_comment("list", ms, ws, kind="threaded")

        assert result["count"] == 1
        items = result["comments"]
        assert items[0]["kind"] == "threaded"
        assert items[0]["cell"] == "$C$3"
        assert items[0]["text"] == "Thread msg"
        assert items[0]["author"] == "Bob"

    def test_list_all_returns_both(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        cmt = self._make_list_note("$A$1", "Note", "Alice")
        ct = self._make_list_ct("$B$2", "Thread", "Bob")
        ws.Comments = [cmt]
        ws.CommentsThreaded = [ct]

        result = _call_comment("list", ms, ws, kind="all")

        assert result["count"] == 2
        kinds = {c["kind"] for c in result["comments"]}
        assert kinds == {"note", "threaded"}

    def test_list_empty_sheet_returns_zero(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ws.Comments = []
        ws.CommentsThreaded = []

        result = _call_comment("list", ms, ws, kind="all")

        assert result["count"] == 0
        assert result["comments"] == []

    def test_list_author_name_for_threaded(self):
        """Threaded comment author is Author.Name (object), not Author (string)."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ct = self._make_list_ct(author_name="Carol")
        ws.Comments = []
        ws.CommentsThreaded = [ct]

        result = _call_comment("list", ms, ws, kind="threaded")

        # Should use ct.Author.Name, NOT ct.Author (which is an object)
        assert result["comments"][0]["author"] == "Carol"

    def test_list_author_for_notes_is_string(self):
        """Legacy comment Author is a string property (no .Name needed)."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        cmt = self._make_list_note(author="Dave")
        ws.Comments = [cmt]
        ws.CommentsThreaded = []

        result = _call_comment("list", ms, ws, kind="note")

        assert result["comments"][0]["author"] == "Dave"


# ── get ───────────────────────────────────────────────────────────────────────

class TestGet:
    def test_get_note_returns_text_and_author(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        cmt = _make_note_mock("Get text", "Alice")
        rng = MagicMock()
        rng.Address = "$A$1"
        rng.Comment = cmt
        ws.Range.return_value = rng

        result = _call_comment("get", ms, ws, cell="A1", kind="note")

        assert result["comment"] == "get"
        assert result["cell"] == "$A$1"
        assert result["note"]["text"] == "Get text"
        assert result["note"]["author"] == "Alice"

    def test_get_note_absent_returns_none(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rng = MagicMock()
        rng.Address = "$A$1"
        rng.Comment = None
        ws.Range.return_value = rng

        result = _call_comment("get", ms, ws, cell="A1", kind="note")

        assert result["note"] is None

    def test_get_threaded_returns_text_and_author_name(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        ct = _make_threaded_mock("Thread text", "Bob")
        rng = MagicMock()
        rng.Address = "$B$2"
        rng.CommentThreaded = ct
        ws.Range.return_value = rng

        result = _call_comment("get", ms, ws, cell="B2", kind="threaded")

        assert result["threaded"]["text"] == "Thread text"
        assert result["threaded"]["author"] == "Bob"

    def test_get_threaded_absent_returns_none(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rng = MagicMock()
        rng.Address = "$B$2"
        rng.CommentThreaded = None
        ws.Range.return_value = rng

        result = _call_comment("get", ms, ws, cell="B2", kind="threaded")

        assert result["threaded"] is None

    def test_get_all_returns_both(self):
        """kind='all' on get returns both note and threaded keys."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        cmt = _make_note_mock("Note here", "Alice")
        ct = _make_threaded_mock("Thread here", "Bob")
        rng = MagicMock()
        rng.Address = "$A$1"
        rng.Comment = cmt
        rng.CommentThreaded = ct
        ws.Range.return_value = rng

        result = _call_comment("get", ms, ws, cell="A1", kind="all")

        assert result["note"]["text"] == "Note here"
        assert result["threaded"]["text"] == "Thread here"

    def test_get_all_when_cell_empty(self):
        """kind='all' on a cell with no comments returns both keys as None."""
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rng = MagicMock()
        rng.Address = "$Z$99"
        rng.Comment = None
        rng.CommentThreaded = None
        ws.Range.return_value = rng

        result = _call_comment("get", ms, ws, cell="Z99", kind="all")

        assert result["note"] is None
        assert result["threaded"] is None


# ── Result shape ───────────────────────────────────────────────────────────────

class TestResultShape:
    """Every result dict must carry the 'comment' key + relevant sub-keys."""

    def test_add_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        cmt = _make_note_mock("t")
        rng = MagicMock()
        rng.Address = "$A$1"
        rng.Comment = None
        ws.Range.return_value = rng

        def _side(text):
            rng.Comment = cmt
        rng.AddComment.side_effect = _side

        result = _call_comment("add", ms, ws, cell="A1", text="t", kind="note")
        assert "comment" in result
        assert "kind" in result
        assert "cell" in result
        assert "applied" in result

    def test_delete_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        cmt = _make_note_mock()
        rng = MagicMock()
        rng.Address = "$A$1"
        rng.Comment = cmt
        ws.Range.return_value = rng
        # Side effect: simulate COM removing the comment on Delete().
        def _side():
            rng.Comment = None
        cmt.Delete.side_effect = _side
        result = _call_comment("delete", ms, ws, cell="A1", kind="note")
        assert result["comment"] == "delete"
        assert result["applied"]["deleted"] is True

    def test_list_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock("S")
        ws.Comments = []
        ws.CommentsThreaded = []
        result = _call_comment("list", ms, ws, kind="all")
        assert "comment" in result
        assert "sheet" in result
        assert "kind" in result
        assert "comments" in result
        assert "count" in result
        assert isinstance(result["comments"], list)

    def test_get_shape(self):
        ms = make_mock_session()
        ws = _make_sheet_mock()
        rng = MagicMock()
        rng.Address = "$A$1"
        rng.Comment = None
        ws.Range.return_value = rng
        result = _call_comment("get", ms, ws, cell="A1", kind="note")
        assert "comment" in result
        assert "cell" in result
