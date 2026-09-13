"""diagnostics の分類テスト（extracted dict を手組みで与える純関数テスト）。"""
from pivot_diag.diagnostics import diagnose


def _pivot(name, loc_sheet, loc_ref, src_sheet, src_ref):
    return {
        "name": name, "worksheet": loc_sheet, "location_ref": loc_ref,
        "source": {"type": "worksheet", "sheet": src_sheet, "ref": src_ref},
    }


def _extract(*pivots, sheets=None):
    return {"pivots": list(pivots), "sheets": sheets or []}


def test_no_issues_clean_workbook():
    r = diagnose(_extract(
        _pivot("P1", "P1", "A3:B8", "Data", "A1:D7"),
        _pivot("P2", "P2", "A3:B8", "Data", "F1:H7"),
        sheets=["Data"],
    ))
    assert r["findings"] == []
    assert r["pivots_total"] == 2
    assert r["sources_ok"] == 2


def test_overlapping_sources_detected():
    r = diagnose(_extract(
        _pivot("P1", "P1", "A3:B8", "Data", "A1:D7"),
        _pivot("P2", "P2", "A3:B8", "Data", "C1:E20"),   # C列が重なる
        sheets=["Data"],
    ))
    ov = [f for f in r["findings"] if f["status"] == "overlapping_sources"]
    assert len(ov) == 1
    assert "P1" in ov[0]["detail"] and "P2" in ov[0]["detail"]


def test_overlapping_locations_detected():
    r = diagnose(_extract(
        _pivot("P1", "P1", "A3:C10", "Data", "A1:D7"),
        _pivot("P2", "P1", "B5:D12", "Data", "F1:H7"),   # 配置が重なる（別ソース）
        sheets=["Data"],
    ))
    ov = [f for f in r["findings"] if f["status"] == "overlapping_locations"]
    assert len(ov) == 1


def test_identical_sources_are_shared_not_overlapping():
    """完全に同じ範囲なら shared_source（重複カウントではなく情報）。"""
    r = diagnose(_extract(
        _pivot("P1", "P1", "A3:B8", "Data", "A1:D7"),
        _pivot("P2", "P2", "A3:B8", "Data", "A1:D7"),
        sheets=["Data"],
    ))
    statuses = {f["status"] for f in r["findings"]}
    assert "shared_source" in statuses
    assert "overlapping_sources" not in statuses


def test_missing_source_detected():
    r = diagnose(_extract({"name": "Broken", "worksheet": "P1",
                           "location_ref": "A3:B8", "source": {"type": None}}))
    assert r["findings"][0]["status"] == "missing_source"


def test_external_or_named_source_is_informational():
    r = diagnose(_extract(_pivot("P1", "P1", "A3:B8", None, None)))
    # ref 無し（named range 等）は external_source として情報表示
    assert r["findings"][0]["status"] == "external_source"


def test_truly_unparseable_ref_is_informational():
    """完全に解釈不能な ref のみ unparseable_ref になる（A:E は whole-column として正常パース）。"""
    r = diagnose(_extract(_pivot("P1", "P1", "A3:B8", "Data", "###invalid###"), sheets=["Data"]))
    up = [f for f in r["findings"] if f["status"] == "unparseable_ref"]
    assert len(up) == 1


def test_different_sheets_same_range_not_overlap():
    r = diagnose(_extract(
        _pivot("P1", "P1", "A3:B8", "Data", "A1:D7"),
        _pivot("P2", "P2", "A3:B8", "Archive", "A1:D7"),   # シート違い
        sheets=["Data", "Archive"],
    ))
    assert r["findings"] == []


def test_missing_source_sheet_detected():
    """cacheSource の参照先シートが workbook に存在しない（改名・削除）場合。"""
    ex = {
        "sheets": ["Archive"],   # "Data" シートは改名されて存在しない
        "pivots": [_pivot("P1", "P1", "A3:B8", "Data", "A1:D7")],
    }
    r = diagnose(ex)
    ms = [f for f in r["findings"] if f["status"] == "missing_source_sheet"]
    assert len(ms) == 1
    assert "'Data'" in ms[0]["detail"]


def test_existing_source_sheet_no_finding():
    ex = {
        "sheets": ["Data"],
        "pivots": [_pivot("P1", "P1", "A3:B8", "Data", "A1:D7")],
    }
    sheet_findings = [f for f in diagnose(ex)["findings"]
                      if f["status"] == "missing_source_sheet"]
    assert sheet_findings == []
