# Working agreement for this repository

Short rules for anyone — human or agent — changing `pivot-diag`.

## What this tool is

`pivot-diag` audits the pivot table configurations inside Excel workbooks
(`.xlsx` / `.xlsm`) and reports the problems that break a refresh: overlapping
pivots, shared sources, broken links. It is **read-only** and **stdlib only**.

## Ground rules

- **Never modify the workbook.** The tool opens files to read them; a change that
  writes to the input is a bug, not a feature.
- **No new dependencies.** Stdlib only, on purpose. The CLI must run from a clean
  checkout with nothing installed.
- **Deterministic output.** Same workbook in, same report out. No network, no LLM,
  no clock in the output.
- **Tests come with the change.** `python -m pytest -q` must pass, and a new
  reported problem class needs a fixture workbook that triggers it.
- **No secrets in Git**, and no absolute paths in code, tests or docs — a clone
  must run anywhere.
- **Do not bypass the secret scan.** `git commit --no-verify` is never a fix for a
  gitleaks hit; rotate the credential and rewrite the commit.
- **The README is a promise.** Every documented command must work on a fresh
  clone, and a badge must point at CI rather than hardcode a test count.

## Checks that must pass

```
pip install -e ".[dev]"
python -m pytest -q
```
