"""Build script for thepexcel-mcp.mcpb bundle.

Usage:
    uv run python scripts/build_mcpb.py

Output: dist/thepexcel-mcp.mcpb  (ZIP archive)

Bundle contents (per MCPB uv-type spec):
  manifest.json
  pyproject.toml
  uv.lock
  src/thepexcel_mcp/**   (all .py source files)

Excluded:
  __pycache__/, *.pyc, .venv/, tests/, docs/, scripts/, dist/, *.mcpb
"""

from __future__ import annotations

import pathlib
import zipfile
import sys

REPO = pathlib.Path(__file__).parent.parent.resolve()
DIST = REPO / "dist"
OUT  = DIST / "thepexcel-mcp.mcpb"

# Files/dirs to always include at repo root
ROOT_INCLUDES = ["manifest.json", "pyproject.toml", "uv.lock"]

# Directories to bundle recursively (relative to REPO)
SRC_DIRS = ["src"]

# Patterns to skip inside src
SKIP_SUFFIXES = {".pyc"}
SKIP_DIRS    = {"__pycache__", ".venv", ".git"}


def _should_include(rel: pathlib.Path) -> bool:
    for part in rel.parts:
        if part in SKIP_DIRS:
            return False
    return rel.suffix not in SKIP_SUFFIXES


def build() -> None:
    DIST.mkdir(exist_ok=True)

    with zipfile.ZipFile(OUT, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        # Root-level files
        for name in ROOT_INCLUDES:
            path = REPO / name
            if path.exists():
                zf.write(path, arcname=name)
                print(f"  + {name}")
            else:
                print(f"  ! MISSING: {name}", file=sys.stderr)

        # Source tree
        for src_dir_name in SRC_DIRS:
            src_dir = REPO / src_dir_name
            if not src_dir.is_dir():
                print(f"  ! MISSING dir: {src_dir_name}", file=sys.stderr)
                continue
            for filepath in sorted(src_dir.rglob("*")):
                if not filepath.is_file():
                    continue
                rel = filepath.relative_to(REPO)
                if _should_include(rel):
                    zf.write(filepath, arcname=str(rel).replace("\\", "/"))
                    print(f"  + {str(rel)}")

    size_kb = OUT.stat().st_size / 1024
    print(f"\nBuilt: {OUT}  ({size_kb:.1f} KB)")


if __name__ == "__main__":
    print(f"Building thepexcel-mcp.mcpb from {REPO}")
    build()
