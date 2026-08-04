"""Packaging checks for the reproducible MCPB builder."""

from __future__ import annotations

import importlib.util
from pathlib import Path
import zipfile

import pytest


_SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "build_mcpb.py"
_SPEC = importlib.util.spec_from_file_location("thepexcel_build_mcpb", _SCRIPT)
assert _SPEC is not None and _SPEC.loader is not None
build_mcpb = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(build_mcpb)


def _redirect_builder(monkeypatch, tmp_path):
    dist = tmp_path / "dist"
    monkeypatch.setattr(build_mcpb, "REPO", tmp_path)
    monkeypatch.setattr(build_mcpb, "DIST", dist)
    monkeypatch.setattr(build_mcpb, "OUT", dist / "thepexcel-mcp.mcpb")


def test_build_fails_before_writing_when_lock_is_missing(monkeypatch, tmp_path):
    _redirect_builder(monkeypatch, tmp_path)
    (tmp_path / "manifest.json").write_text("{}", encoding="utf-8")
    (tmp_path / "pyproject.toml").write_text("[project]", encoding="utf-8")

    with pytest.raises(FileNotFoundError, match="uv.lock"):
        build_mcpb.build()

    assert not build_mcpb.OUT.exists()


def test_build_includes_lock_and_source(monkeypatch, tmp_path):
    _redirect_builder(monkeypatch, tmp_path)
    for name, content in {
        "manifest.json": "{}",
        "pyproject.toml": "[project]",
        "uv.lock": "version = 1",
    }.items():
        (tmp_path / name).write_text(content, encoding="utf-8")
    package = tmp_path / "src" / "thepexcel_mcp"
    package.mkdir(parents=True)
    (package / "__init__.py").write_text("", encoding="utf-8")

    build_mcpb.build()

    with zipfile.ZipFile(build_mcpb.OUT) as archive:
        assert set(archive.namelist()) == {
            "manifest.json",
            "pyproject.toml",
            "uv.lock",
            "src/thepexcel_mcp/__init__.py",
        }
