"""合成 OOXML フィクスチャ（fixture_builder）からの抽出テスト。"""
from fixture_builder import build_workbook

from pivot_diag.ooxml import extract_pivots


def _fixture(tmp_path, pivots, sheets=("P1", "Data")):
    path = tmp_path / "fixture.xlsx"
    build_workbook(path, pivots, sheets=sheets)
    return extract_pivots(str(path))


def test_extract_two_pivots(tmp_path):
    pivots = [
        {"name": "PivotByRegion", "location": "A3:B8", "cache_id": 1,
         "src_sheet": "Data", "src_ref": "A1:D7"},
        {"name": "PivotByProduct", "location": "A3:B8", "cache_id": 1,
         "src_sheet": "Data", "src_ref": "A1:D7"},
    ]
    r = _fixture(tmp_path, pivots)
    assert r["sheets"] == ["P1", "Data"]
    assert len(r["pivots"]) == 2
    p = r["pivots"][0]
    assert p["name"] == "PivotByRegion"
    assert p["location_ref"] == "A3:B8"
    assert p["worksheet"] == "P1"          # sheet1 (=P1) の rels から逆引き
    assert p["source"]["sheet"] == "Data"
    assert p["source"]["ref"] == "A1:D7"


def test_extract_single_pivot_no_props(tmp_path):
    pivots = [{"name": "Solo", "location": "A3:A6", "cache_id": 1,
               "src_sheet": "Data", "src_ref": "A1:B7"}]
    r = _fixture(tmp_path, pivots, sheets=("Data",))
    assert r["pivots"][0]["worksheet"] == "Data"
    assert r["pivots"][0]["source"]["ref"] == "A1:B7"
