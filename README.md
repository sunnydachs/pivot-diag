# pivot-diag

**Audit pivot table configurations in Excel workbooks — overlaps, shared sources, broken links. Read-only, stdlib only.**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.11%2B-blue.svg)](pyproject.toml)
[![Tests](https://img.shields.io/badge/tests-27%20passing-brightgreen.svg)](tests/)

Excel がピボットテーブルの構成問題を教えてくれるのは「重複しています」エラーが出た瞬間だけです。`pivot-diag` はワークブック内の全ピボットテーブルの構成（配置・ソース範囲・キャッシュ参照）を読み取り、**壊れる前に**問題を列挙します。

- ⚠️ **OVERLAPPING LOCATIONS** — ピボットの配置範囲同士が重なる（Excel の「PivotTable レポートを重複させることができません」エラーの領域）
- ⚠️ **OVERLAPPING SOURCES** — 複数ピボットのソース範囲が同じシート上で重なる（二重カウント・リフレッシュ順序問題）
- · **SHARED SOURCE** — 同一ソースを共有（意図的なら問題なし。独立更新したい場合は cache 分けを）
- ❌ **MISSING SOURCE SHEET** — ソース参照先のシートがワークブックに存在しない（改名・削除で壊れたリンク）
- · **informational** — named range / external ソース、A:E 等の非対応形式

すべて **決定的** — OOXML（pivotCacheDefinition / pivotTable XML）を直接解析し、同じ入力なら常に同じレポート。LLM 不使用、依存ゼロ（標準ライブラリのみ）、**読み取り専用**。

## Status: practice/portfolio project — competition unverified

このツールは、ピボットテーブルが多数あるワークブックで「どのピボットがどの範囲を参照しているか分からず、リフレッシュを壊してしまう」という悩みから作りました。

**正直な位置づけ**: このアイデアの競合状況は**未検証**です（評価時に検索基盤が不安定で、競合調査が不完全なまま）。VBA やアドインで同種のチェックをするツールは存在する可能性が高いため、「需要検証済みの製品」としては扱わないでください。手元のワークブックで実際に使うための、練習を兼ねたユーティリティです。

## Install

```bash
pip install git+https://github.com/sunnydachs/pivot-diag.git
```

依存はありません（標準ライブラリのみ）。

## Usage

```bash
pivot-diag report.xlsx          # 1 ファイルを診断
pivot-diag ./reports/           # ディレクトリ内の *.xlsx / *.xlsm を一括診断
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

## 検出のしくみ

xlsx/xlsm は OOXML zip であり、ピボットの定義は `xl/pivotCache/pivotCacheDefinitionN.xml`（ソース範囲）と `xl/pivotTables/pivotTableN.xml`（配置・cacheId）に分かれています。pivot-diag はこれらを直接解析します（openpyxl 等は不使用 — ライブラリの読み込み quirks に左右されないため）。

| 診断 | 判定基準 |
|---|---|
| OVERLAPPING LOCATIONS | 配置範囲（location ref）が同じシート上で交差 |
| OVERLAPPING SOURCES | ソース範囲が同じシート上で交差（完全一致は SHARED SOURCE として区別） |
| SHARED SOURCE | 複数ピボットが同一ソース範囲を参照 |
| MISSING SOURCE SHEET | cacheSource の参照先シートがワークブックに存在しない |
| UNPARSEABLE REF | 全列参照（A:E）等の MVP 非対応形式 |

## Known limitations

- **読み取り専用**。修正・再配置はしません（レポートの指摘箇所を手で直す運用を想定）
- 全列参照（`A:E`）や名前付き範囲ソースは構文解析できず `informational` 扱い
- プロパティ（pivot field）レベルの詳細検証は未対応。配置・ソース範囲・キャッシュ参照の粒度
- `.xls`（旧形式）は対象外（OOXML のみ）

## Development

```bash
pip install -e ".[dev]"
python -m pytest tests/ -q   # 27 tests, fully offline
```

テストフィクスチャは tests/fixture_builder.py が OOXML zip を直接組み立てます（Excel 不要）。

## Provenance

Inspired by a recurring pain observed in pivot-heavy workbooks: refresh breaks and nobody knows which pivot configuration caused it. An independent, general-purpose implementation. This is part of a small family of drift/consistency checkers ([doc-drift](https://github.com/sunnydachs/doc-drift), [plan-drift](https://github.com/sunnydachs/plan-drift)).

## License

[MIT](LICENSE)
