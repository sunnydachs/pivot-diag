# pivot-diag

**Audit pivot table configurations in Excel workbooks — overlaps, shared sources, broken links. Read-only, stdlib only.**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.11%2B-blue.svg)](pyproject.toml)
[![Tests](https://img.shields.io/badge/tests-27%20passing-brightgreen.svg)](tests/)

English | [日本語](README.ja.md)

<!-- Sync with README.ja.md as of commit <sync-point> -->

`pivot-diag` audits the pivot table configurations inside Excel workbooks (.xlsx/.xlsm) and lists problems **before** they break a refresh:

- ⚠️ **OVERLAPPING LOCATIONS** — two pivot table placements intersect (the zone where Excel raises "A PivotTable report cannot overlap another PivotTable report")
- ⚠️ **OVERLAPPING SOURCES** — source ranges of multiple pivots intersect on the same sheet (double counting, refresh-order issues)
- · **SHARED SOURCE** — pivots sharing the exact same source range (fine if intentional; consider separate caches for independent refresh)
- ❌ **MISSING SOURCE SHEET** — a cacheSource points to a sheet that no longer exists (renamed or deleted)
- · **informational** — named-range/external sources, whole-column refs (`A:E`), which are out of scope

Everything is **deterministic**: the workbook is parsed directly as OOXML (`pivotCacheDefinition` / `pivotTable` XML parts) with the standard library only. No LLM, no third-party packages, and the tool is **read-only** — it never modifies your files.

## Status: practice/portfolio project — competition unverified

This tool was inspired by a recurring pain in pivot-heavy workbooks: refresh breaks and nobody knows which pivot configuration caused it.

**Honest positioning**: the competition landscape for this idea has **not** been verified (evaluation was inconclusive due to search infrastructure issues). VBA snippets and add-ins doing similar checks probably exist. This is a practice/portfolio utility built for my own workbooks — use it as such.

## How it works

xlsx/xlsm files are OOXML zips. Pivot definitions live in two places: `xl/pivotCache/pivotCacheDefinitionN.xml` (source ranges) and `xl/pivotTables/pivotTableN.xml` (placement and cache references). pivot-diag parses these parts directly with `zipfile` + `ElementTree` — deliberately not via openpyxl, so the results don't depend on a third-party reader's quirks.

| Check | Rule |
|---|---|
| OVERLAPPING LOCATIONS | two pivot placements intersect on the same worksheet |
| OVERLAPPING SOURCES | source ranges intersect on the same sheet (identical ranges are reported as SHARED SOURCE instead) |
| MISSING SOURCE SHEET | the sheet referenced by a cacheSource no longer exists — broken link (renamed/deleted) |
| UNPARSEABLE REF | whole-column refs (`A:E`) and similar — reported as informational, never guessed |

## Install

```bash
pip install git+https://github.com/sunnydachs/pivot-diag.git
```

No third-party dependencies (standard library only).

## Usage

```bash
pivot-diag report.xlsx          # audit one workbook
pivot-diag ./reports/           # every *.xlsx / *.xlsm in a directory
pivot-diag report.xlsx --json
```

Example output:

```
pivot-diag — read-only pivot configuration audit

clean.xlsx: 2 pivot(s), 2 bounded source range(s)
  - SalesByRegion @P1!A3:B8 ← source Data!A1:D7
  - SalesByProduct @P2!E3:F8 ← source Data!A1:D7
  no issues found.

overlap.xlsx: 2 pivot(s), 2 bounded source range(s)
  - P1 @Data!A3:B8 ← source Data!A1:D7
  - P2 @Data!A3:B8 ← source Data!C1:E20
  ⚠️ OVERLAPPING LOCATIONS
    pivot locations overlap: 'P1' (Data!A3:B8) and 'P2' (Data!A3:B8) — Excel raises 'A PivotTable report cannot overlap another PivotTable report'
  ⚠️ OVERLAPPING SOURCES
    pivots 'P1' (source A1:D7 on 'Data') and 'P2' (source C1:E20 on 'Data') have overlapping source ranges — double-counting and refresh-order issues

summary: 3 file(s) scanned, 3 finding(s)
```

## Known limitations

- **Read-only.** No repair, no relocation — the report points at configurations for you to fix.
- Property-level (pivot field) validation is out of scope; the granularity is placement, source range, and cache references.
- `.xls` (legacy format) is not supported — OOXML only.
- Whole-column source refs (`A:E`) and named-range sources are reported as informational (not parsed).

## Development

```bash
pip install -e ".[dev]"
python -m pytest tests/ -q   # 27 tests, fully offline
```

Test fixtures assemble OOXML zips directly (no Excel needed) via `tests/fixture_builder.py`. Validated with three scenarios: a clean workbook (no findings), overlapping source ranges (both overlap diagnostics fired), and a renamed source sheet (broken-link detection).

## Provenance

Inspired by a recurring pain observed in pivot-heavy workbooks: refresh breaks and nobody knows which pivot configuration caused it. An independent, general-purpose implementation. Part of a small family of consistency checkers: [doc-drift](https://github.com/sunnydachs/doc-drift) (docs vs code), [plan-drift](https://github.com/sunnydachs/plan-drift) (tracking plan vs code), and this tool (pivot configuration vs itself).

## License

[MIT](LICENSE)
