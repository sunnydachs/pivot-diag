# pivot-diag

**Excel ワークブック内のピボットテーブル構成を診断する — 重複・共有ソース・リンク切れ。読み取り専用・標準ライブラリのみ。**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.11%2B-blue.svg)](pyproject.toml)
[![Tests](https://img.shields.io/badge/tests-27%20passing-brightgreen.svg)](tests/)

[English](README.md) | 日本語

<!-- README.md（英語）と同期。更新時は両方を直すこと -->

`pivot-diag` は Excel ワークブック（.xlsx/.xlsm）内のピボットテーブル構成（配置・ソース範囲・キャッシュ参照）を読み取り、**リフレッシュが壊れる前に**問題を列挙します。

- ⚠️ **OVERLAPPING LOCATIONS** — ピボットの配置範囲同士が重なる（Excel の「PivotTable レポートを重複させることができません」エラーの領域）
- ⚠️ **OVERLAPPING SOURCES** — 複数ピボットのソース範囲が同じシート上で重なる（二重カウント・リフレッシュ順序問題）
- · **SHARED SOURCE** — 同一ソース範囲の共有（意図的なら問題なし。独立更新したい場合は cache 分けを）
- ❌ **MISSING SOURCE SHEET** — cacheSource の参照先シートが存在しない（改名・削除で壊れたリンク）
- · **informational** — named range・外部ソース、全列参照（A:E）等の非対応形式

判定はすべて**決定的**。ワークブックを OOXML zip として直接解析（`pivotCacheDefinition` / `pivotTable` XML）し、標準ライブラリのみで動作します。LLM 不使用、サードパーティ依存なし、**読み取り専用** — ファイルは一切変更しません。

## Status: practice/portfolio — competition 未検証

このツールは、ピボットテーブルが多数あるワークブックで「リフレッシュすると壊れる。どのピボットの構成が原因か分からない」という繰り返される悩みから作りました。

**正直な位置づけ**: このアイデアの競合状況は**未検証**です（評価時に検索基盤が不安定で、競合調査が不完全なまま）。同種のチェックをする VBA やアドインは既に存在する可能性が高いため、「需要検証済みの製品」としては扱わないでください。手元のワークブックで実際に使うための、練習を兼ねたユーティリティです。

## しくみ

xlsx/xlsm は OOXML zip です。ピボットの定義は `xl/pivotCache/pivotCacheDefinitionN.xml`（ソース範囲）と `xl/pivotTables/pivotTableN.xml`（配置・cacheId）に分かれています。pivot-diag はこの2種のパーツを `zipfile` + `ElementTree` で直接解析します — わざと openpyxl を経由せず、結果がライブラリの読み込み quirks に左右されないようにしています。

| 診断 | 規則 |
|---|---|
| OVERLAPPING LOCATIONS | 配置範囲（location ref）が同じシート上で交差 |
| OVERLAPPING SOURCES | ソース範囲が同じシート上で交差（完全一致は SHARED SOURCE として区別） |
| MISSING SOURCE SHEET | cacheSource の参照先シートがワークブックに存在しない |
| UNPARSEABLE REF | 全列参照（A:E）等の MVP 非対応形式 |

## インストール

```bash
pip install git+https://github.com/sunnydachs/pivot-diag.git
```

依存はありません（標準ライブラリのみ）。

## 使い方

```bash
pivot-diag report.xlsx          # 1 ファイルを診断
pivot-diag ./reports/           # ディレクトリ内の *.xlsx / *.xlsm を一括診断
pivot-diag report.xlsx --json
```

出力例:

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

- **読み取り専用**。修正・再配置はしません（レポートの指摘箇所を手で直す運用を想定）
- プロパティ（pivot field）レベルの詳細検証は未対応。配置・ソース範囲・キャッシュ参照の粒度
- `.xls`（旧形式）は対象外（OOXML のみ）
- 全列参照（`A:E`）や名前付き範囲ソースは構文解析できず `informational` 扱い

## 開発

```bash
pip install -e ".[dev]"
python -m pytest tests/ -q   # 27 tests, fully offline
```

テストフィクスチャは tests/fixture_builder.py が OOXML zip を直接組み立てます（Excel 不要）。実検証は3シナリオで実施: クリーンブック（findings 0）、ソース範囲重複（両方の overlap 診断が発火）、ソースシート改名（リンク切れ検出）。

## Provenance

ピボットが多いワークブックで繰り返し観測される悩み — 「リフレッシュが壊れても、どのピボット構成が原因か分からない」— から作った、独立した汎用実装です。一貫性チェッカー・ファミリーの1つ: [doc-drift](https://github.com/sunnydachs/doc-drift)（docs vs code）、[plan-drift](https://github.com/sunnydachs/plan-drift)（tracking plan vs code）、そして本ツール（ピボット構成の自己診断）。

## License

[MIT](LICENSE)
