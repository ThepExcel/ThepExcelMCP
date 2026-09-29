"""Measure MCP server cold-start cost (no Excel needed).

    uv run python scripts/bench_startup.py [--runs 3]

Reports, per run: wall time of `uv run --directory <repo> python -c "import
thepexcel_mcp.server"` (the documented launch path minus the stdio loop) and
the same import via the venv python directly (isolates uv's own overhead).
Then prints the 15 slowest modules from `python -X importtime` of one import.
All output goes to stderr/stdout of this script only — it never starts a server.
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
VENV_PY = REPO / ".venv" / "Scripts" / "python.exe"
IMPORT = "import thepexcel_mcp.server"


def _timed(cmd: list[str]) -> float:
    env = os.environ.copy()
    env.pop("VIRTUAL_ENV", None)
    t0 = time.perf_counter()
    subprocess.run(cmd, check=True, env=env, cwd=REPO,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return time.perf_counter() - t0


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--runs", type=int, default=3)
    runs = ap.parse_args().runs

    via_uv = ["uv", "run", "--directory", str(REPO), "python", "-c", IMPORT]
    direct = [str(VENV_PY), "-c", IMPORT]
    for i in range(runs):
        print(f"run {i + 1}: uv run {_timed(via_uv):.2f}s | venv python {_timed(direct):.2f}s")

    res = subprocess.run([str(VENV_PY), "-X", "importtime", "-c", IMPORT],
                         capture_output=True, text=True, cwd=REPO)
    rows = []
    for line in res.stderr.splitlines():
        if not line.startswith("import time:") or "cumulative" in line:
            continue
        # "import time:  self_us |  cum_us | name"
        self_us, cum_us, name = (p.strip() for p in line.split(":", 1)[1].split("|"))
        rows.append((int(cum_us), int(self_us), name))
    total = max((r[0] for r in rows), default=0)
    print(f"\nimporttime total (largest cumulative): {total / 1e6:.2f}s — top 15 by cumulative:")
    for cum, self_, name in sorted(rows, reverse=True)[:15]:
        print(f"  {cum / 1e3:8.0f} ms cum  {self_ / 1e3:7.0f} ms self  {name}")


if __name__ == "__main__":
    sys.exit(main())
