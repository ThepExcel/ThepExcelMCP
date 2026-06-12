"""Unit tests for the M code analyzer — no Excel required."""

import pytest

from thepexcel_mcp.analysis.pq_analyzer import analyze_mcode


# ── Helpers ────────────────────────────────────────────────────────────────────

def _issues_by_rule(result, rule):
    return [i for i in result.issues if i.rule == rule]


# ── Basic contract ──────────────────────────────────────────────────────────────

def test_clean_query_no_issues():
    formula = """let
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
    Filtered = Table.SelectRows(Source, each [Status] = "Active")
in
    Filtered"""
    result = analyze_mcode("CleanQuery", formula)
    assert result.query_name == "CleanQuery"
    assert result.step_count >= 2
    assert result.issues == []
    assert result.summary == "No issues found"


def test_step_count():
    formula = """let
    A = 1,
    B = 2,
    C = 3
in C"""
    result = analyze_mcode("Steps", formula)
    assert result.step_count == 3


def test_complexity_low():
    formula = "let A = 1 in A"
    result = analyze_mcode("Q", formula)
    assert result.estimated_complexity == "low"


def test_complexity_medium():
    steps = "\n".join(f"    S{i} = {i}," for i in range(8))
    formula = f"let\n{steps}\n    Last = 99\nin Last"
    result = analyze_mcode("Q", formula)
    assert result.estimated_complexity in ("medium", "high")


def test_complexity_high():
    steps = "\n".join(f"    S{i} = {i}," for i in range(16))
    formula = f"let\n{steps}\n    Last = 99\nin Last"
    result = analyze_mcode("Q", formula)
    assert result.estimated_complexity == "high"


# ── Rule: query-folding ─────────────────────────────────────────────────────────

def test_query_folding_detected():
    formula = """let
    Source = Sql.Database("server", "db"),
    Added = Table.AddColumn(Source, "X", each 1)
in Added"""
    result = analyze_mcode("Folding", formula)
    issues = _issues_by_rule(result, "query-folding")
    assert len(issues) >= 1
    assert issues[0].severity == "info"


# ── Rule: unnecessary-buffer ────────────────────────────────────────────────────

def test_unnecessary_buffer_no_join():
    formula = """let
    Source = Excel.CurrentWorkbook(){[Name="T"]}[Content],
    Buffered = Table.Buffer(Source)
in Buffered"""
    result = analyze_mcode("Buf", formula)
    issues = _issues_by_rule(result, "unnecessary-buffer")
    assert len(issues) == 1
    assert issues[0].severity == "warning"


def test_buffer_with_join_no_warning():
    formula = """let
    A = Table.Buffer(Excel.CurrentWorkbook(){[Name="T1"]}[Content]),
    B = Excel.CurrentWorkbook(){[Name="T2"]}[Content],
    Joined = Table.NestedJoin(A, "ID", B, "ID", "Extra", JoinKind.Inner)
in Joined"""
    result = analyze_mcode("Join", formula)
    issues = _issues_by_rule(result, "unnecessary-buffer")
    assert issues == []


# ── Rule: multiple-type-changes ────────────────────────────────────────────────

def test_multiple_type_changes():
    formula = """let
    S = Excel.CurrentWorkbook(){[Name="T"]}[Content],
    T1 = Table.TransformColumnTypes(S, {{"A", type text}}),
    T2 = Table.TransformColumnTypes(T1, {{"B", Int64.Type}}),
    T3 = Table.TransformColumnTypes(T2, {{"C", type date}})
in T3"""
    result = analyze_mcode("Types", formula)
    issues = _issues_by_rule(result, "multiple-type-changes")
    assert len(issues) == 1


def test_two_type_changes_no_warning():
    formula = """let
    S = Excel.CurrentWorkbook(){[Name="T"]}[Content],
    T1 = Table.TransformColumnTypes(S, {{"A", type text}}),
    T2 = Table.TransformColumnTypes(T1, {{"B", Int64.Type}})
in T2"""
    result = analyze_mcode("Types2", formula)
    issues = _issues_by_rule(result, "multiple-type-changes")
    assert issues == []


# ── Rule: select-vs-remove ─────────────────────────────────────────────────────

def test_select_vs_remove_many_cols():
    cols = ", ".join(f'"Col{i}"' for i in range(6))
    formula = f"""let
    S = Excel.CurrentWorkbook(){{[Name="T"]}}[Content],
    R = Table.RemoveColumns(S, {{{cols}}})
in R"""
    result = analyze_mcode("Remove", formula)
    issues = _issues_by_rule(result, "select-vs-remove")
    assert len(issues) == 1


# ── Rule: unbuffered-multi-join ────────────────────────────────────────────────

def test_unbuffered_multi_join():
    formula = """let
    A = Excel.CurrentWorkbook(){[Name="T1"]}[Content],
    B = Excel.CurrentWorkbook(){[Name="T2"]}[Content],
    C = Excel.CurrentWorkbook(){[Name="T3"]}[Content],
    J1 = Table.NestedJoin(A, "ID", B, "ID", "X", JoinKind.Inner),
    J2 = Table.NestedJoin(J1, "ID", C, "ID", "Y", JoinKind.Inner)
in J2"""
    result = analyze_mcode("MultiJoin", formula)
    issues = _issues_by_rule(result, "unbuffered-multi-join")
    assert len(issues) == 1
    assert issues[0].severity == "warning"


# ── Rule: hardcoded-path ───────────────────────────────────────────────────────

def test_hardcoded_windows_path():
    formula = r"""let
    Source = Excel.Workbook(File.Contents("C:\Users\data\file.xlsx"), null, true)
in Source"""
    result = analyze_mcode("HardPath", formula)
    issues = _issues_by_rule(result, "hardcoded-path")
    assert len(issues) == 1


# ── Rule: too-many-steps ───────────────────────────────────────────────────────

def test_too_many_steps():
    steps = "\n".join(f"    S{i} = {i}," for i in range(21))
    formula = f"let\n{steps}\n    Last = 99\nin Last"
    result = analyze_mcode("Big", formula)
    issues = _issues_by_rule(result, "too-many-steps")
    assert len(issues) == 1
    assert issues[0].severity == "warning"


# ── Rule: simple-range ─────────────────────────────────────────────────────────

def test_list_generate_flagged():
    formula = """let
    Nums = List.Generate(() => 1, each _ <= 10, each _ + 1)
in Nums"""
    result = analyze_mcode("Gen", formula)
    issues = _issues_by_rule(result, "simple-range")
    assert len(issues) == 1


# ── Rule: record-field-in-filter ───────────────────────────────────────────────

def test_record_field_in_filter():
    formula = """let
    S = Excel.CurrentWorkbook(){[Name="T"]}[Content],
    F = Table.SelectRows(S, each Record.Field(_, "Status") = "Active")
in F"""
    result = analyze_mcode("RecField", formula)
    issues = _issues_by_rule(result, "record-field-in-filter")
    assert len(issues) == 1
    assert issues[0].severity == "warning"


# ── Rule: no-error-handling ────────────────────────────────────────────────────

def test_web_contents_without_try():
    formula = """let
    Source = Web.Contents("https://api.example.com/data"),
    Parsed = Json.Document(Source)
in Parsed"""
    result = analyze_mcode("Web", formula)
    issues = _issues_by_rule(result, "no-error-handling")
    assert len(issues) == 1
    assert issues[0].severity == "info"


def test_web_contents_with_try_no_warning():
    formula = """let
    Source = try Web.Contents("https://api.example.com/data") otherwise null,
    Parsed = if Source = null then null else Json.Document(Source)
in Parsed"""
    result = analyze_mcode("WebSafe", formula)
    issues = _issues_by_rule(result, "no-error-handling")
    assert issues == []


# ── Summary property ───────────────────────────────────────────────────────────

def test_summary_with_mixed_issues():
    formula = r"""let
    S = Web.Contents("https://example.com"),
    B = Table.Buffer(S),
    A = Table.AddColumn(B, "X", each 1)
in A"""
    result = analyze_mcode("Mixed", formula)
    summary = result.summary
    # At minimum Web + Buffer + AddColumn issues should fire
    assert summary != "No issues found"
    assert "(" in summary  # e.g. "1 warning(s), 2 info(s)"
