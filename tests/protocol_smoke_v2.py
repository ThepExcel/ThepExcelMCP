"""Cross-version stdio handshake smoke: MCP SDK v2 client -> current server.

Run outside the project environment so the v2 client does not conflict with
FastMCP's current ``mcp<2`` dependency::

    uv run --isolated --with mcp==2.0.0 python tests/protocol_smoke_v2.py

This does not touch Excel. It proves protocol negotiation, tool listing, and
the optional BM25 discovery tool through the real stdio transport.
"""

from __future__ import annotations

import asyncio
import os
import tempfile
from pathlib import Path

from mcp import Client, StdioServerParameters
from mcp.client.stdio import stdio_client


REPO = Path(__file__).resolve().parents[1]
SERVER_EXE = REPO / ".venv" / "Scripts" / "thepexcel-mcp.exe"


async def _list_tools(discovery: str) -> tuple[list, str]:
    env = os.environ.copy()
    env["THEPEXCEL_MCP_TOOL_DISCOVERY"] = discovery
    params = StdioServerParameters(command=str(SERVER_EXE), args=[], env=env)
    with tempfile.TemporaryFile(mode="w+", encoding="utf-8") as errlog:
        async with Client(
            stdio_client(params, errlog=errlog),
            read_timeout_seconds=30,
            raise_exceptions=True,
        ) as client:
            response = await client.list_tools()
            protocol = str(getattr(client.session, "protocol_version", "unknown"))
            if discovery == "bm25":
                search = await client.call_tool(
                    "search_tools", {"query": "format cells with colors and borders"}
                )
                text = "\n".join(
                    getattr(item, "text", "") for item in search.content
                )
                assert "excel_format" in text
            return response.tools, protocol


async def main() -> None:
    if not SERVER_EXE.exists():
        raise SystemExit(f"Server executable missing: {SERVER_EXE}; run uv sync first")

    full_tools, full_protocol = await _list_tools("full")
    full_names = {tool.name for tool in full_tools}
    assert len(full_names) == 26, sorted(full_names)
    assert {"excel_workbook", "excel_range", "excel_snapshot"} <= full_names

    search_tools, search_protocol = await _list_tools("bm25")
    search_names = {tool.name for tool in search_tools}
    assert search_names == {
        "excel_workbook",
        "excel_range",
        "search_tools",
        "call_tool",
    }

    print(
        "protocol smoke PASS:",
        f"full={len(full_names)} tools/{full_protocol}",
        f"bm25={len(search_names)} tools/{search_protocol}",
    )


if __name__ == "__main__":
    asyncio.run(main())
