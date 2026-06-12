"""Manual COM smoke test — requires a live Excel instance.

NOT collected by pytest (see pyproject.toml collect_ignore).
Run manually: uv run python tests/smoke_com.py

Only performs READ-ONLY checks. Does not mutate any open workbook.
"""

from __future__ import annotations

import sys

# Read-only: list workbooks and queries
try:
    from thepexcel_mcp.domains.workbook import workbook_action
    from thepexcel_mcp.domains.powerquery import powerquery_action
    from thepexcel_mcp.domains.sheets import sheet_action
except ImportError as e:
    print(f"Import failed: {e}")
    sys.exit(1)


def main():
    print("=== ThepExcelMCP COM Smoke Test ===\n")

    print("[1] Listing open workbooks...")
    try:
        result = workbook_action("list")
        print(f"    Workbooks: {result}")
    except Exception as e:
        print(f"    ERROR: {e}")
        print("    (Is Excel running?)")
        return

    workbooks = result.get("workbooks", [])
    if not workbooks:
        print("    No workbooks open — skipping further checks.")
        return

    active = next((w["name"] for w in workbooks if w.get("active")), workbooks[0]["name"])
    print(f"\n[2] Workbook info for active: '{active}'")
    try:
        info = workbook_action("info", workbook=active)
        print(f"    Sheets: {info['sheets']}")
        print(f"    Queries: {info['query_count']}, Tables: {info['table_count']}")
    except Exception as e:
        print(f"    ERROR: {e}")

    print(f"\n[3] Listing sheets in '{active}'...")
    try:
        sheets = sheet_action("list", workbook=active)
        print(f"    {sheets}")
    except Exception as e:
        print(f"    ERROR: {e}")

    print(f"\n[4] Listing Power Query queries in '{active}'...")
    try:
        queries = powerquery_action("list", workbook=active)
        print(f"    Found {queries['count']} queries")
        for q in queries["queries"]:
            print(f"    - {q['name']} ({q['line_count']} lines)")
    except Exception as e:
        print(f"    ERROR: {e}")

    print("\n=== Smoke test complete (read-only, no mutations) ===")


if __name__ == "__main__":
    main()
