"""Power Query M code analyzer — detects anti-patterns and suggests optimizations.

Rule-based analysis that catches common M code performance issues.
Works on the M code string directly (no Excel connection needed).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


@dataclass
class AnalysisIssue:
    """A single optimization finding."""

    severity: str  # "warning" | "info" | "error"
    rule: str
    message: str
    line: int | None = None
    suggestion: str = ""


@dataclass
class AnalysisResult:
    """Complete analysis result for a query."""

    query_name: str
    issues: list[AnalysisIssue] = field(default_factory=list)
    step_count: int = 0
    has_source: bool = False
    estimated_complexity: str = "low"  # low | medium | high

    @property
    def summary(self) -> str:
        warnings = sum(1 for i in self.issues if i.severity == "warning")
        errors = sum(1 for i in self.issues if i.severity == "error")
        infos = sum(1 for i in self.issues if i.severity == "info")
        parts = []
        if errors:
            parts.append(f"{errors} error(s)")
        if warnings:
            parts.append(f"{warnings} warning(s)")
        if infos:
            parts.append(f"{infos} info(s)")
        return ", ".join(parts) if parts else "No issues found"


def analyze_mcode(query_name: str, formula: str) -> AnalysisResult:
    """Analyze M code for common anti-patterns and optimization opportunities."""
    result = AnalysisResult(query_name=query_name)
    lines = formula.split("\n")

    # Count let/in steps
    step_pattern = re.compile(r"^\s+#?\"?[\w\s]+\"?\s*=", re.MULTILINE)
    steps = step_pattern.findall(formula)
    result.step_count = len(steps)
    result.has_source = "Source" in formula

    # Complexity estimate
    if result.step_count > 15:
        result.estimated_complexity = "high"
    elif result.step_count > 7:
        result.estimated_complexity = "medium"

    # ── Rule checks ────────────────────────────────────────────

    _check_no_folding_operations(formula, lines, result)
    _check_table_buffer_misuse(formula, lines, result)
    _check_unnecessary_type_conversions(formula, lines, result)
    _check_select_vs_remove(formula, lines, result)
    _check_nested_joins(formula, lines, result)
    _check_hardcoded_paths(formula, lines, result)
    _check_large_step_count(result)
    _check_list_generate_vs_numbers(formula, lines, result)
    _check_record_field_access_in_filter(formula, lines, result)
    _check_missing_error_handling(formula, lines, result)

    return result


def _find_line(lines: list[str], pattern: str) -> int | None:
    """Find the first line number containing the pattern."""
    for i, line in enumerate(lines, 1):
        if pattern in line:
            return i
    return None


def _check_no_folding_operations(
    formula: str, lines: list[str], result: AnalysisResult
) -> None:
    """Detect operations that prevent query folding."""
    no_fold_funcs = [
        "Table.AddColumn",
        "Table.TransformColumns",
        "Table.FillDown",
        "Table.FillUp",
        "Table.Combine",
        "List.Generate",
        "List.Accumulate",
    ]
    for func in no_fold_funcs:
        if func in formula:
            result.issues.append(
                AnalysisIssue(
                    severity="info",
                    rule="query-folding",
                    message=f"`{func}` prevents query folding after this step.",
                    line=_find_line(lines, func),
                    suggestion=(
                        "Move filter/select steps BEFORE this operation "
                        "to maximize query folding."
                    ),
                )
            )


def _check_table_buffer_misuse(
    formula: str, lines: list[str], result: AnalysisResult
) -> None:
    """Detect Table.Buffer used without a clear reason (e.g., no join)."""
    buffer_count = formula.count("Table.Buffer")
    join_count = formula.count("Table.NestedJoin") + formula.count("Table.Join")
    if buffer_count > 0 and join_count == 0:
        result.issues.append(
            AnalysisIssue(
                severity="warning",
                rule="unnecessary-buffer",
                message=(
                    "Table.Buffer found without any join. "
                    "Buffering without reuse wastes memory."
                ),
                line=_find_line(lines, "Table.Buffer"),
                suggestion="Remove Table.Buffer unless the table is referenced multiple times.",
            )
        )


def _check_unnecessary_type_conversions(
    formula: str, lines: list[str], result: AnalysisResult
) -> None:
    """Detect multiple Table.TransformColumnTypes in sequence."""
    type_steps = [
        i
        for i, line in enumerate(lines, 1)
        if "Table.TransformColumnTypes" in line
    ]
    if len(type_steps) > 2:
        result.issues.append(
            AnalysisIssue(
                severity="warning",
                rule="multiple-type-changes",
                message=(
                    f"Found {len(type_steps)} type conversion steps. "
                    "Consider consolidating into one."
                ),
                line=type_steps[0],
                suggestion=(
                    "Merge all type changes into a single "
                    "Table.TransformColumnTypes call at the end."
                ),
            )
        )


def _check_select_vs_remove(
    formula: str, lines: list[str], result: AnalysisResult
) -> None:
    """Suggest Table.SelectColumns when removing many columns."""
    remove_match = re.search(
        r"Table\.RemoveColumns\s*\([^,]+,\s*\{([^}]+)\}", formula
    )
    if remove_match:
        removed_cols = remove_match.group(1).count(",") + 1
        if removed_cols > 5:
            result.issues.append(
                AnalysisIssue(
                    severity="info",
                    rule="select-vs-remove",
                    message=(
                        f"Removing {removed_cols} columns. "
                        "Table.SelectColumns may be clearer."
                    ),
                    line=_find_line(lines, "Table.RemoveColumns"),
                    suggestion=(
                        "Use Table.SelectColumns to keep only needed columns "
                        "instead of removing many."
                    ),
                )
            )


def _check_nested_joins(
    formula: str, lines: list[str], result: AnalysisResult
) -> None:
    """Detect multiple nested joins that might benefit from buffering."""
    join_count = formula.count("Table.NestedJoin")
    buffer_count = formula.count("Table.Buffer")
    if join_count >= 2 and buffer_count == 0:
        result.issues.append(
            AnalysisIssue(
                severity="warning",
                rule="unbuffered-multi-join",
                message=(
                    f"Found {join_count} nested joins without Table.Buffer. "
                    "Repeated source evaluation may be slow."
                ),
                line=_find_line(lines, "Table.NestedJoin"),
                suggestion=(
                    "Buffer frequently joined tables with Table.Buffer "
                    "to avoid re-evaluation."
                ),
            )
        )


def _check_hardcoded_paths(
    formula: str, lines: list[str], result: AnalysisResult
) -> None:
    """Detect hardcoded file paths."""
    path_patterns = [
        r'[A-Z]:\\[^\\"]+',  # Windows paths
        r"//[a-zA-Z][a-zA-Z0-9._-]+/",  # UNC paths
    ]
    for pattern in path_patterns:
        match = re.search(pattern, formula)
        if match:
            result.issues.append(
                AnalysisIssue(
                    severity="info",
                    rule="hardcoded-path",
                    message=f"Hardcoded path detected: {match.group()[:50]}...",
                    line=_find_line(lines, match.group()[:20]),
                    suggestion=(
                        "Consider using a parameter or "
                        "Excel.CurrentWorkbook() for portability."
                    ),
                )
            )
            break


def _check_large_step_count(result: AnalysisResult) -> None:
    """Warn about overly complex queries."""
    if result.step_count > 20:
        result.issues.append(
            AnalysisIssue(
                severity="warning",
                rule="too-many-steps",
                message=f"Query has {result.step_count} steps — consider splitting.",
                suggestion=(
                    "Break into smaller helper queries for maintainability. "
                    "Use reference queries or function queries."
                ),
            )
        )


def _check_list_generate_vs_numbers(
    formula: str, lines: list[str], result: AnalysisResult
) -> None:
    """Suggest {start..end} syntax instead of List.Generate for simple ranges."""
    if "List.Generate" in formula and "List.Numbers" not in formula:
        result.issues.append(
            AnalysisIssue(
                severity="info",
                rule="simple-range",
                message="List.Generate may be replaceable with {start..end} or List.Numbers.",
                line=_find_line(lines, "List.Generate"),
                suggestion="For simple numeric ranges, use `{1..100}` or `List.Numbers(start, count)`.",
            )
        )


def _check_record_field_access_in_filter(
    formula: str, lines: list[str], result: AnalysisResult
) -> None:
    """Detect Record.Field in Table.SelectRows (slower than [Column] syntax)."""
    if "Record.Field" in formula and "Table.SelectRows" in formula:
        result.issues.append(
            AnalysisIssue(
                severity="warning",
                rule="record-field-in-filter",
                message="Record.Field used in filter — slower than direct field access.",
                line=_find_line(lines, "Record.Field"),
                suggestion="Use `each [ColumnName]` instead of `each Record.Field(_, \"ColumnName\")`.",
            )
        )


def _check_missing_error_handling(
    formula: str, lines: list[str], result: AnalysisResult
) -> None:
    """Suggest error handling for web/database sources."""
    risky_sources = ["Web.Contents", "Sql.Database", "OData.Feed", "Odbc.Query"]
    for src in risky_sources:
        if src in formula and "try" not in formula:
            result.issues.append(
                AnalysisIssue(
                    severity="info",
                    rule="no-error-handling",
                    message=f"`{src}` without `try` — refresh errors won't be caught.",
                    line=_find_line(lines, src),
                    suggestion=f"Wrap `{src}(...)` in `try ... otherwise ...` for graceful error handling.",
                )
            )
            break
