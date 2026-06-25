# Contributing

## Branch workflow

`main` is the published, always-installable branch. **Do not commit directly to
`main`** — work on a branch and open a pull request:

```bash
git checkout -b feat/my-change       # or fix/... , docs/...
# ... make changes, commit ...
git push -u origin feat/my-change
gh pr create --fill                  # then review the diff and merge
```

The pull request is the review gate before anything reaches the public default
branch — use it to eyeball the full diff (especially for anything that could
contain data that should not be public). `main` is protected: it requires a PR
and rejects force-pushes and branch deletion.

## No customer / private data (HARDLINE)

This is a **public** repository. `samples/`, tests, docstrings, and docs must use
**synthetic, anonymized data only** — never a real customer or company name, a
real product catalog or model codes, client figures, or any third-party business
data. Keep internal working notes (inbox/handoff/memory/vault references) out of
committed files. See the "Public repo — synthetic data only" constraint in
[CLAUDE.md](CLAUDE.md).

### Enable the pre-push safety hook

A pre-push hook scans the lines you are about to push for likely customer/PII
data and blocks the push if it finds any. Enable it once per clone:

```bash
git config core.hooksPath .githooks
```

Optionally, create a local `.sensitive-terms.local` file (it is gitignored) with
one real client/company name per line — the hook hard-blocks any push that
re-introduces those terms:

```
# .sensitive-terms.local  — NEVER commit this file
Some Client Co
AnotherBrand
```

To override a verified false positive for a single push:

```bash
ALLOW_SENSITIVE=1 git push ...
```

## Dev quickstart

```bash
uv sync
uv run pytest -q                                              # unit tests, no Excel needed
THEPEXCEL_MCP_AUTOLAUNCH=1 uv run python tests/smoke_com.py  # live COM smoke (Windows + Excel)
```
