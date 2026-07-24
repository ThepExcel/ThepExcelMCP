# Security Policy

ThepExcelMCP is a **local, stdio MCP server** that drives a real Excel Desktop
process through Windows COM. It has no network service and no telemetry — it runs
as a child process of your MCP client, on your machine, against your Excel. That
shape rules out whole classes of remote vulnerabilities, but a few local ones are
worth understanding and reporting.

## Supported versions

This project ships from `main`; the latest published commit on `main` is the only
supported version. Please reproduce any issue against a current `uv sync` before
reporting.

## Reporting a vulnerability

**Please do not open a public issue for a security problem.**

- Preferred: use GitHub's **[Report a vulnerability](https://github.com/ThepExcel/ThepExcelMCP/security/advisories/new)**
  (Security → Advisories) for a private disclosure.
- Or email **thepexcel@gmail.com** with `SECURITY` in the subject.

Include: what you did, what happened, what you expected, the affected file/tool,
and a minimal repro. We aim to acknowledge within a few business days. Please give
a reasonable window to release a fix before any public disclosure.

## Security model & things to know

The risk here is **local code execution and file access via an AI agent**, not a
network attacker. Keep these in mind:

- **The agent drives your live Excel — including the workbook you already have
  open.** A prompt-injected or mistaken agent can modify or save real workbooks.
  Use `excel_snapshot` before risky bulk operations, and review agent actions.
- **VBA execution is double-gated and OFF by default.** The `excel_vba` tool runs
  only when **both** `THEPEXCEL_MCP_ENABLE_VBA=1` **and** Excel's own *"Trust
  access to the VBA project object model"* setting are enabled. Enabling it lets an
  agent write and run arbitrary VBA (i.e. arbitrary code) in your Excel session.
  Leave it off unless you specifically need it and trust the agent.
- **File paths are honored as given.** Tools such as `excel_workbook(open/save_as)`,
  `excel_shape(add_image)`, and `excel_page_setup(export_pdf)` read/write paths the
  agent supplies. The server does not sandbox the filesystem — it has the
  permissions of the user running it.
- **`=PY()` (Python in Excel)** is inserted as a formula and executed by
  Microsoft's cloud service under your M365 identity — not by this server.
- **No secrets belong in this repo.** It is public. See
  [CONTRIBUTING.md](CONTRIBUTING.md) for the synthetic-data-only rule and the
  pre-push safety hook that helps enforce it.

## Out of scope

- Vulnerabilities in Excel, Windows, `pywin32`, or `fastmcp` themselves — report
  those to their maintainers (we'll happily bump a dependency once fixed).
- The inherent capability of an AI agent to change data when you point it at real
  workbooks. That is the tool working as designed; mitigate with snapshots, diffs,
  and review.
