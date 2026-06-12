"""Shared test helpers.

make_mock_session
-----------------
Creates a MagicMock ExcelSession whose run_com() is a transparent passthrough:
    run_com(fn, *args, **kwargs) → fn(*args, **kwargs)

This lets existing unit tests continue to work after the STA worker-thread
refactor (Phase 3): tests patch _session with this mock, and domain code that
calls _session.run_com(fn, ...) still executes fn synchronously in the test
process — no real COM worker thread started, no Excel needed.
"""

from __future__ import annotations

from unittest.mock import MagicMock


def make_mock_session(**side_effects) -> MagicMock:
    """Return a MagicMock session with run_com as a transparent passthrough.

    Additional keyword args set side_effect on named methods, e.g.:
        make_mock_session(get_workbook=ToolError("no excel"))
    """
    mock = MagicMock()
    mock.run_com.side_effect = lambda fn, *args, **kwargs: fn(*args, **kwargs)
    for method_name, effect in side_effects.items():
        getattr(mock, method_name).side_effect = (
            effect if callable(effect) else (lambda e=effect: (_ for _ in ()).throw(e))
        )
    return mock
