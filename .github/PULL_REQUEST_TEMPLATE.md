<!-- Thanks for contributing to ThepExcelMCP! Please fill this in. -->

## What & why

<!-- What does this PR change, and why? Link any related issue: Closes #123 -->

## Type of change

- [ ] Bug fix
- [ ] New feature / new tool or action
- [ ] Docs only
- [ ] Refactor / internal (no behavior change)

## Checklist

- [ ] I branched off `main` and this is a PR (not a direct push).
- [ ] `uv run pytest -q` passes locally (unit tests, no Excel needed).
- [ ] For COM-behavior changes: verified against a real Excel via
      `tests/smoke_com.py` (or explained why it couldn't be, e.g. no Windows).
- [ ] Added/updated unit tests for the change.
- [ ] **No customer / private data.** Only synthetic, anonymized data in code,
      tests, docstrings, and docs (public repo — see
      [CONTRIBUTING.md](../CONTRIBUTING.md)).
- [ ] The pre-push safety hook is enabled (`git config core.hooksPath .githooks`).
- [ ] Docs/README/docstrings updated if the tool surface or behavior changed.

## Notes for reviewers

<!-- Anything that needs a closer look, tradeoffs, follow-ups, or known gaps. -->
