# MCP RC 2026-07-28 ("Sessionless Core") — ThepExcelMCP Exposure Audit & Migration Plan

> **Audited:** 2026-06-24 · **Re-verified:** 2026-08-04 with an MCP SDK 2.0.0 stdio client
> **Verdict:** **LOW exposure · ZERO code migration required.** The one real risk is an *operational*, time-gated client/server version-skew window (late-July → Sept 2026), not an architectural one.

## 2026-08-04 status

- Dependency hardening is complete: `fastmcp>=3.4.5,<4`, a committed `uv.lock`, and
  `uv sync --frozen` in CI/bundle builds.
- `tests/protocol_smoke_v2.py` now launches the real stdio server from an isolated
  `mcp==2.0.0` client. It verifies negotiation, all 26 tools in full mode, and the
  optional BM25 discovery surface. The tested pair negotiated protocol `2025-11-25`.
- FastMCP 3.4.5 currently resolves `mcp 1.28.0`; native sessionless-protocol adoption
  remains a deliberate future dependency upgrade, but v2-client interoperability no
  longer depends on a manual tool-list check.

---

## TL;DR

ThepExcelMCP is a **stdio-only** FastMCP server. The breaking parts of the 2026-07-28 RC that everyone is alarmed about — removal of the `Mcp-Session-Id` header — are **exclusively a Streamable HTTP transport concern**. stdio never used them. Every protocol/handshake detail the RC *does* change for all transports lives **inside `fastmcp`**, not in this repo's code. A grep of `src/` for `Mcp-Session-Id | session_id | initialize | client-id | per-request metadata` finds **zero** protocol-handling code.

**No `@mcp.tool` handler changes. No server.py changes. No "stateless COM redesign."** The only action is **operational**: keep `fastmcp` current and restart the process when an RC-ready release ships — with a dated watch so the client/server skew window doesn't silently brick the server.

---

## What the RC actually changes

| SEP | Change | Scope | Touches ThepExcelMCP? |
|---|---|---|---|
| **SEP-2567** | Removes `Mcp-Session-Id` header + protocol session concept (clean break) | **Streamable HTTP only** | ❌ No — stdio never had it |
| **SEP-2575** | Removes `initialize`/`initialized` handshake; metadata moves inline; adds `Mcp-Method`/`Mcp-Name` routing headers | All transports incl. stdio | ⚠️ Indirectly — but handled 100% inside `fastmcp` |
| **SEP-2596** | Guarantees ≥12-month deprecation window; this release's deprecations are annotation-only | All | ✅ Protective — old methods keep working ≥1 year |
| — | "not initialized" error code `-32002` → `-32602` | All | ❌ No — this repo never pattern-matches these codes |

Timeline: RC locked **2026-05-21**, final spec **2026-07-28**. `mcp` python-SDK v2 beta ≈ **June 30**, stable **July 27**. As of this audit the official SDK has *not* activated negotiation (`v2.0.0a2`: "2026-07-28 is modeled but not yet negotiable").

Sources: [RC blog](https://blog.modelcontextprotocol.io/posts/2026-07-28-release-candidate/) · [SEP-2567 PR #2567](https://github.com/modelcontextprotocol/modelcontextprotocol/pull/2567) · [MCP transports spec 2025-11-25](https://modelcontextprotocol.io/specification/2025-11-25/basic/transports) · [fastmcp releases](https://github.com/jlowin/fastmcp/releases) · [mcp python-sdk releases](https://github.com/modelcontextprotocol/python-sdk/releases)

---

## Why exposure is LOW (verified)

1. **Transport = stdio only** — `mcp.run(transport="stdio")` in [`server.py`](../src/thepexcel_mcp/server.py). No HTTP/SSE/uvicorn/starlette deps anywhere.
2. **`Mcp-Session-Id` was never in stdio** — it's an HTTP session-affinity header (SEP-2567). Removing it cannot affect a transport that never sent it.
3. **All protocol/handshake handling is inside `fastmcp`** — this repo registers 26 `@mcp.tool` handlers + calls `mcp.run()`. No handler receives a context object, session id, client identity, or per-request protocol metadata (grep-confirmed: 0 hits).
4. **`ExcelSession` ≠ MCP session** — [`session.py`](../src/thepexcel_mcp/session.py)'s `ExcelSession` is a single STA COM worker thread holding one live `Excel.Application` handle. Statefulness lives in the **process**, not in any protocol session. The RC's removal of *protocol* sessions is orthogonal to it.
5. **The spec built a ramp, not a cliff** — a stdio backwards-compat probe (`server/discover`) exists, and SEP-2596 guarantees ≥12 months before any removed method actually breaks.

### Impact map (per area)

| Area | Affected by RC? | Severity | Note |
|---|---|---|---|
| `ExcelSession` STA worker / single COM apartment | No | none | Process-local state, not protocol state |
| ROT fallback (multi-Excel instance lookup) | No | none | Local mechanism, no session metadata |
| `add_table`/`load_to_datamodel` COM deadlock | No | medium\* | Pre-existing COM constraint, unchanged by RC |
| 120s COM timeout | No | none | Local safeguard |
| `THEPEXCEL_MCP_AUTOLAUNCH` | No | none | Read once per process |
| MCPB `manifest.json` + stdio `uv` launch | No | none | stdio transport unaffected |
| `@mcp.tool` registration + dispatch | No | none | Pure-Python signatures; fastmcp owns marshalling |
| client identity / session id / per-request meta | No | none | **0 code references** |
| Process-level statefulness (1 process : 1 Excel) | Yes | low | By design; document as hard constraint |
| `fastmcp` protocol-version compatibility | Yes | low | The one real lever — see risk below |

\* The deadlock severity is "medium" as a *standalone* operational hazard, but it is **not introduced or worsened by the RC**.

---

## The ONE real risk (operational, time-gated)

**Client/server protocol skew during late-July → September 2026.**

The MCP lifecycle spec puts the disconnect decision on the **client**: *"If the client does not support the version in the server's response, it SHOULD disconnect."* Claude Code issue #768 (reported by the red-team pass; exact URL not re-verified here) is proof-of-shape that a client can **hard-fail on the `protocolVersion` field with strict validation _before any server communication_** — i.e. the server logs show no incoming request at all.

**The brick scenario:** if a Claude Desktop/Code build flips to the sessionless 2026-07-28 client protocol **before** an RC-ready `fastmcp` has been `uv sync`'d **and the stdio process restarted**, the server could silently vanish from Claude's tool list.

- Current locked state (2026-08-04): `fastmcp 3.4.5` / `mcp 1.28.0`; the isolated SDK-v2 probe successfully negotiates `2025-11-25` and lists/calls tools.
- Mitigating lags: client adoption lags spec publication by weeks–months; SEP-2596's 12-month deprecation window; the `server/discover` compat probe.
- **Failure signature:** the `thepexcel-excel` tools disappear from Claude (tools not callable) — and it looks **nothing like** a COM/Excel error. Easy to misdiagnose as an Excel problem and burn an hour.

This keeps the verdict at **LOW, not NONE** (correcting the impact-map's over-rosy "none" and its wrong assumption that skew resolves via "graceful negotiation").

---

## Action items (active + dated — NOT passive "just track fastmcp")

1. **[DONE]** Pin FastMCP to a tested floor+ceiling and commit `uv.lock` for deterministic installs and bundles.
2. **[DONE]** Put a real SDK-v2 stdio handshake in CI (`tests/protocol_smoke_v2.py`), including tool-list and BM25 search-tool calls.
3. **Continue the release watch.** The buy signal is a FastMCP release whose notes name MCP 2026-07-28 / SEP-2575 / sessionless / `mcp>=2.0.0`; then make a deliberate lockfile upgrade and rerun the runbook below.
4. **Keep the diagnostic:** if `thepexcel-excel` disappears from the client tool list without a COM error, suspect protocol skew first, then upgrade/restart the stdio process.

### Migration runbook (execute only when the buy signal fires)

```
1. Update the FastMCP constraint deliberately, then `uv lock --upgrade-package fastmcp`
2. uv sync --frozen                         # install the reviewed lock exactly
3. THEPEXCEL_MCP_AUTOLAUNCH=1 uv run python tests/smoke_com.py   # all tools still pass?
4. uv run --isolated --with mcp==2.0.0 python tests/protocol_smoke_v2.py
5. CLIENT-HANDSHAKE CHECK (manual final mile):
   restart Claude Desktop/Code → confirm "thepexcel-excel" appears in the tool list
   AND one tool call succeeds (e.g. excel_workbook list).
   ⚠ smoke_com.py does NOT exercise the MCP protocol handshake — it owns its own
     Excel and bypasses the stdio layer. exit-0 on smoke is NOT proof of protocol health.
6. If step 5 fails → the running stdio process is stale: kill + respawn
   (re-registration via `claude mcp add` is NOT needed). Editable install ≠ hot reload.
```

## What NOT to do

- ❌ Do **not** pre-emptively "migrate," redesign the COM session for statelessness, add routing headers, or change error-code matching — none apply to a stdio server, and all protocol concerns are fastmcp's job.
- ❌ Do **not** change any `@mcp.tool` handler or `server.py` — zero handler changes are required.
- ❌ Do **not** treat `smoke_com.py` passing as proof the protocol layer is healthy.

---

## Why this was initially over-scoped

An early assessment mistakenly flagged this as "🔴 high severity, migrate before July 28" by conflating ThepExcelMCP's COM `ExcelSession` with the MCP *protocol* session. The four "required actions" it listed are all mis-scoped:

1. "audit initialize/initialized handshake code" — this repo has none.
2. "update error code -32002 → -32602" — never matched here.
3. "design stateless-compatible COM session strategy" — COM state is process-level, not protocol-level.
4. "add routing headers to responses" — HTTP-only, and fastmcp owns transport.

**Bottom line:** treat the 2026-07-28 RC as an operational watch (keep `fastmcp` current, restart on an RC-ready release), not a code migration.
