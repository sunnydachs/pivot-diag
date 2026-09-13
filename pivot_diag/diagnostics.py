"""diagnostics — ピボット定義のリストから問題を検出する（純関数）。

検出する診断（元の悩み: 「どのピボットが重複してリフレッシュを壊すか分からない」）:
  overlapping_locations — ピボットの配置範囲（location）同士が同じシート上で重なる
                          （Excel が「重複しています」エラーを出す領域）
  overlapping_sources   — 複数ピボットのソース範囲が同じシート上で重なる
                          （データ二重カウント・リフレッシュ順序問題の温床）
  shared_source         — 複数ピボットが同一ソース範囲を共有（情報表示）
  external_source       — worksheetSource に明示的なセル範囲が無い（named range 等）
  unparseable_ref       — ソース範囲が A:E 等の MVP 非対応形式
  missing_source        — cacheSource が読み取れない
"""
from itertools import combinations

from pivot_diag.ranges import parse_ref, ranges_overlap


def diagnose(extracted: dict) -> dict:
    """extract_pivots の出力から診断を生成する（純関数）。

    返り値: {findings: [{status, detail}], pivots_total, sources_ok}
    """
    pivots = extracted["pivots"]
    findings = []

    # 0) ピボット配置（location）の重なり — Excel のハードエラー領域
    locations = []
    for p in pivots:
        if p.get("location_ref") and p.get("worksheet"):
            lr = parse_ref(p["location_ref"], default_sheet=p["worksheet"])
            if lr.get("ok"):
                locations.append((p, lr))
    for (p1, l1), (p2, l2) in combinations(locations, 2):
        if ranges_overlap(l1, l2):
            findings.append({
                "status": "overlapping_locations",
                "detail": f"pivot locations overlap: {p1['name']!r} "
                          f"({l1['sheet']}!{p1['location_ref']}) and {p2['name']!r} "
                          f"({l2['sheet']}!{p2['location_ref']}) — Excel raises "
                          f"'A PivotTable report cannot overlap another PivotTable report'",
            })


    # 各ピボットのソース範囲をパース
    parsed = []
    for p in pivots:
        src = p.get("source") or {}
        ref = src.get("ref")
        if src.get("type") is None:
            findings.append({
                "status": "missing_source",
                "detail": f"pivot {p['name']!r}: cacheSource not found — "
                          f"cache definition may be broken or external",
            })
            parsed.append((p, None))
            continue
        if ref is None:
            findings.append({
                "status": "external_source",
                "detail": f"pivot {p['name']!r}: source has no explicit range "
                          f"(name={src.get('name')!r}, sheet={src.get('sheet')!r}) — "
                          f"table/named-range source, not a cell range",
            })
            parsed.append((p, None))
            continue
        pr = parse_ref(ref, default_sheet=src.get("sheet"))
        if not pr.get("ok"):
            findings.append({
                "status": "unparseable_ref",
                "detail": f"pivot {p['name']!r}: source ref {ref!r} could not be parsed",
            })
            parsed.append((p, None))
            continue
        parsed.append((p, pr))

    ok = [(p, pr) for p, pr in parsed if pr is not None]

    # ソースシートの存在確認（リンク切れ: シート改名・削除で壊れた参照）
    sheet_names = set(extracted.get("sheets") or [])
    for p, pr in ok:
        if (pr["sheet"] or "") and pr["sheet"] not in sheet_names:
            findings.append({
                "status": "missing_source_sheet",
                "detail": f"pivot {p['name']!r}: source sheet {pr['sheet']!r} does not "
                          f"exist in the workbook — broken link (sheet renamed or deleted?)",
            })

    # ソース重複・共有の判定
    for (p1, r1), (p2, r2) in combinations(ok, 2):
        same = (r1["sheet"], r1["min_row"], r1["min_col"],
                r1["max_row"], r1["max_col"]) == \
               (r2["sheet"], r2["min_row"], r2["min_col"],
                r2["max_row"], r2["max_col"])
        if same:
            findings.append({
                "status": "shared_source",
                "detail": f"pivots {p1['name']!r} and {p2['name']!r} share the exact "
                          f"source range on sheet {r1['sheet']!r} "
                          f"({p1['source']['ref']}) — fine if intentional; consider "
                          f"separate caches if they must refresh independently",
            })
        elif ranges_overlap(r1, r2):
            findings.append({
                "status": "overlapping_sources",
                "detail": f"pivots {p1['name']!r} (source {p1['source']['ref']} on "
                          f"{r1['sheet']!r}) and {p2['name']!r} (source "
                          f"{p2['source']['ref']} on {r2['sheet']!r}) have overlapping "
                          f"source ranges — double-counting and refresh-order issues",
            })

    return {
        "findings": findings,
        "pivots_total": len(pivots),
        "sources_ok": len(ok),
    }
