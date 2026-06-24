"""Extract Power Query M code (and source URLs) from .xlsx files WITHOUT Excel.

Handles:
  - Modern workbooks: DataMashup binary (customXml/item*.xml, UTF-16, base64 -> inner ZIP ->
    Formulas/Section1.m).  All 'shared <query> = ...' definitions are printed.
  - Legacy workbooks: xl/connections.xml and xl/queryTables/queryTable*.xml (plaintext URLs,
    no base64).  Falls back automatically when no DataMashup part exists.

Usage:
    uv run skills/excel-god/scripts/extract_mcode.py file1.xlsx [file2.xlsx ...]

Output (per file):
    - Full M code (Section1.m)
    - All http/https source URLs, with Google/published ones highlighted
    - Source-assignment lines (Source = ..., Web.Contents(...), etc.)

Requires only stdlib — no dependencies.  Prints UTF-8 (safe on Windows with uv).

CAVEAT: Reading M code is file-only and Excel-free.  WRITING M back requires repacking the
DataMashup binary and recomputing its checksum/permissions — for that, use the live Excel MCP
(excel_powerquery action="update") instead of this script.
"""
# /// script
# requires-python = ">=3.9"
# ///

import sys
import re
import io
import base64
import zipfile
import struct

sys.stdout.reconfigure(encoding="utf-8")

URL_RE = re.compile(r"https?://[^\s\"')>]+")


# ---------------------------------------------------------------------------
# Modern path: DataMashup binary inside customXml/item*.xml
# ---------------------------------------------------------------------------

def _find_datamashup_b64(xlsx_zip: zipfile.ZipFile) -> str | None:
    """Return base64 DataMashup payload, or None if not found."""
    candidates = [n for n in xlsx_zip.namelist()
                  if n.lower().endswith(".xml")]
    # Try customXml/* and any DataMashup-named parts first (fast path)
    priority = [n for n in candidates
                if "customxml" in n.lower() or "datamashup" in n.lower()]
    rest = [n for n in candidates if n not in priority]

    for name in priority + rest:
        try:
            raw = xlsx_zip.read(name)
        except Exception:
            continue
        # Try UTF-16 first (the spec); fall back to UTF-16-LE and UTF-8
        text = None
        for enc in ("utf-16", "utf-16-le", "utf-8"):
            try:
                cand = raw.decode(enc)
                if "DataMashup" in cand:
                    text = cand
                    break
            except Exception:
                continue
        if text is None:
            # Try raw bytes scan (some encodings look like plain ASCII)
            if b"DataMashup" not in raw:
                continue
            m = re.search(
                rb"<DataMashup[^>]*>([A-Za-z0-9+/=\s]+)</DataMashup>", raw
            )
            if m:
                return m.group(1).decode("ascii", errors="ignore")
            continue
        m = re.search(
            r"<DataMashup[^>]*>([A-Za-z0-9+/=\s]+)</DataMashup>", text
        )
        if m:
            return m.group(1)
    return None


def _decode_datamashup(b64: str) -> tuple[list[str], str | None]:
    """
    Decode DataMashup base64 -> binary -> inner package ZIP -> .m files.

    Binary layout:
        [int32 version (4 bytes)]
        [int32 package_len (4 bytes)]
        [package ZIP (package_len bytes)]
        [metadata ...]     <- trailing bytes; MUST NOT include in ZipFile

    Returns (list_of_m_texts, error_message_or_None).
    """
    try:
        blob = base64.b64decode("".join(b64.split()))
    except Exception as e:
        return [], f"base64-decode failed: {e}"

    if len(blob) < 8:
        return [], "blob too short (< 8 bytes)"

    # version = struct.unpack("<i", blob[0:4])[0]  # informational only
    plen = struct.unpack("<i", blob[4:8])[0]

    if plen <= 0 or 8 + plen > len(blob):
        # Fallback: try to locate the PK signature and read from there
        idx = blob.find(b"PK\x03\x04")
        if idx < 0:
            return [], f"invalid package_len ({plen}) and no PK magic found"
        pkg = blob[idx:]
    else:
        pkg = blob[8 : 8 + plen]

    try:
        inner = zipfile.ZipFile(io.BytesIO(pkg))
    except Exception as e:
        return [], f"inner ZIP open failed: {e}"

    m_texts = []
    for name in inner.namelist():
        if name.lower().endswith(".m"):
            try:
                m_texts.append(inner.read(name).decode("utf-8", errors="ignore"))
            except Exception:
                pass

    if not m_texts:
        return [], f"inner ZIP opened but no .m files found (contents: {inner.namelist()})"

    return m_texts, None


# ---------------------------------------------------------------------------
# Legacy path: xl/connections.xml + xl/queryTables/queryTable*.xml
# ---------------------------------------------------------------------------

def _extract_legacy_urls(xlsx_zip: zipfile.ZipFile) -> list[str]:
    """Extract URLs from connections.xml and queryTable*.xml (no DataMashup)."""
    found: list[str] = []
    legacy_parts = [n for n in xlsx_zip.namelist()
                    if "connections" in n.lower() or "querytable" in n.lower()]
    for name in legacy_parts:
        try:
            text = xlsx_zip.read(name).decode("utf-8", errors="ignore")
        except Exception:
            continue
        found.extend(URL_RE.findall(text))
    return sorted(set(found))


# ---------------------------------------------------------------------------
# Per-file orchestration
# ---------------------------------------------------------------------------

def extract_file(path: str) -> None:
    print("=" * 80)
    print(f"FILE: {path}")
    print()

    try:
        xlsx_zip = zipfile.ZipFile(path)
    except Exception as e:
        print(f"  [ERROR] Cannot open as ZIP: {e}")
        return

    # --- Modern path ---
    b64 = _find_datamashup_b64(xlsx_zip)
    if b64 is not None:
        m_texts, err = _decode_datamashup(b64)
        if err:
            print(f"  [ERROR] DataMashup decode: {err}")
        elif not m_texts:
            print("  [WARN] DataMashup found but no .m sections extracted.")
        else:
            full_m = "\n\n".join(m_texts)
            print("  ── M Code (Section1.m) " + "─" * 56)
            print(full_m)
            print()

            # Source URLs
            all_urls = sorted(set(URL_RE.findall(full_m)))
            google_urls = [u for u in all_urls
                           if "docs.google.com" in u
                           or "spreadsheets" in u
                           or "pub?" in u
                           or "output=csv" in u]

            print("  ── Source URLs ─────────────────────────────────────────")
            if google_urls:
                print("  [Google/published]")
                for u in google_urls:
                    print(f"    {u}")
            if all_urls:
                non_google = [u for u in all_urls if u not in google_urls]
                if non_google:
                    print("  [Other]")
                    for u in non_google:
                        print(f"    {u}")
            else:
                print("  (no http/https URLs found in M code)")
            print()

            # Source-assignment lines (quick context)
            print("  ── Source / connection lines ────────────────────────────")
            keywords = ("Source", "Web.Contents", "Csv.Document",
                        "Excel.Workbook", "OData.Feed", "Sql.Database",
                        "SharePoint.Files", "File.Contents")
            shown = False
            for line in full_m.splitlines():
                if any(kw in line for kw in keywords):
                    print(f"    {line.strip()[:200]}")
                    shown = True
            if not shown:
                print("  (none)")
            print()
            return  # done — don't fall through to legacy

    # --- Legacy path (no DataMashup) ---
    legacy_urls = _extract_legacy_urls(xlsx_zip)
    if legacy_urls:
        print("  [INFO] No DataMashup part — legacy connections found.")
        print("  ── Legacy connection URLs ───────────────────────────────────")
        for u in legacy_urls:
            print(f"    {u}")
        print()
    else:
        print("  [INFO] No Power Query (DataMashup) or legacy connections found in this file.")
        print("         This workbook may not have any Power Query queries.")
        print()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: uv run extract_mcode.py file1.xlsx [file2.xlsx ...]")
        sys.exit(1)

    for path in sys.argv[1:]:
        extract_file(path)
