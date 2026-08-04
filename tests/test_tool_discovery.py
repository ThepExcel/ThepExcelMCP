"""Unit tests for the optional progressive tool-discovery mode."""

from __future__ import annotations

import pytest
from fastmcp.server.transforms.search import BM25SearchTransform

from thepexcel_mcp.server import _tool_discovery_transforms


@pytest.mark.parametrize("mode", ["", "full", "off", "0", "false", "no"])
def test_full_catalog_modes_have_no_transform(mode):
    assert _tool_discovery_transforms(mode) == []


@pytest.mark.parametrize("mode", ["bm25", "search", "on", "1", "true", "yes"])
def test_search_modes_enable_bm25(mode):
    transforms = _tool_discovery_transforms(mode)

    assert len(transforms) == 1
    assert isinstance(transforms[0], BM25SearchTransform)


def test_invalid_discovery_mode_fails_fast():
    with pytest.raises(ValueError, match="full.*bm25"):
        _tool_discovery_transforms("vector-magic")
