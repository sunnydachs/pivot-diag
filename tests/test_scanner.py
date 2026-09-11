"""scanner / CLI 結合テスト（fixture_builder で実 OOXML 形式の xlsx を組み立て）。"""
import json

from fixture_builder import build_workbook

from pivot_diag.cli import main, scan


def _make_workbook(tmp_path):
    # Data!A1:D7 を共有する2ピボット（配置は P1 シート上で重ならない位置）
    build_workbook(
        tmp_path / "clean.xlsx",
        [
            {"name": "PivotByRegion", "location": "A3:B8", "cache_id": 1,
             "src_sheet": "Data", "src_ref": "A1:D7"},
            {"name": "PivotByProduct", "location": "E3:F8", "cache_id": 1,
             "src_sheet": "Data", "src_ref": "A1:D7"},
        ],
        sheets=("P1", "Data"),
    )


def test_scan_clean_fixture_reports_shared_source(tmp_path):
    _make_workbook(tmp_path)
    results = scan(str(tmp_path / "clean.xlsx"))
    r = results[0]
    assert r["file"] == "clean.xlsx"
    assert r["pivots_total"] == 2
    statuses = {f["status"] for f in r["findings"]}
    assert statuses == {"shared_source"}   # 同一ソース共有は情報表示


def test_scan_detects_overlapping_sources(tmp_path):
    # ソース範囲が C 列で重なる2ピボット
    build_workbook(
        tmp_path / "overlap.xlsx",
        [
            {"name": "P1", "location": "A3:B8", "cache_id": 1,
             "src_sheet": "Data", "src_ref": "A1:D7"},
            {"name": "P2", "location": "A3:B8", "cache_id": 1,
             "src_sheet": "Data", "src_ref": "C1:E20"},
        ],
        sheets=("Data",),
    )
    results = scan(str(tmp_path / "overlap.xlsx"))
    statuses = {f["status"] for f in results[0]["findings"]}
    assert "overlapping_sources" in statuses


def test_scan_directory_skips_non_xlsx(tmp_path):
    _make_workbook(tmp_path)
    (tmp_path / "notes.txt").write_text("not a workbook")
    results = scan(str(tmp_path))
    assert len(results) == 1
    assert results[0]["file"] == "clean.xlsx"


def test_cli_main_json(tmp_path, capsys):
    _make_workbook(tmp_path)
    rc = main([str(tmp_path / "clean.xlsx"), "--json"])
    assert rc == 0
    data = json.loads(capsys.readouterr().out)
    assert data[0]["pivots_total"] == 2
    assert data[0]["findings"][0]["status"] == "shared_source"


def test_cli_text_render(tmp_path, capsys):
    _make_workbook(tmp_path)
    main([str(tmp_path / "clean.xlsx")])
    out = capsys.readouterr().out
    assert "SHARED SOURCE" in out
    assert "summary:" in out
