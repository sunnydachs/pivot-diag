from pivot_diag.ranges import col_to_index, parse_ref, ranges_overlap


def test_col_to_index():
    assert col_to_index("A") == 1
    assert col_to_index("Z") == 26
    assert col_to_index("AA") == 27


def test_parse_ref_with_sheet():
    r = parse_ref("Sheet1!A1:E20")
    assert r["ok"] is True
    assert r["sheet"] == "Sheet1"
    assert (r["min_row"], r["min_col"], r["max_row"], r["max_col"]) == (1, 1, 20, 5)


def test_parse_ref_quoted_sheet():
    r = parse_ref("'My Sheet'!B2:D4")
    assert r["sheet"] == "My Sheet"
    assert (r["min_row"], r["min_col"]) == (2, 2)


def test_parse_ref_default_sheet():
    r = parse_ref("A1:C3", default_sheet="Data")
    assert r["sheet"] == "Data"


def test_parse_ref_single_cell():
    r = parse_ref("Sheet1!B2")
    assert (r["min_row"], r["min_col"], r["max_row"], r["max_col"]) == (2, 2, 2, 2)


def test_parse_ref_reversed_bounds_are_normalized():
    r = parse_ref("Sheet1!E20:A1")
    assert (r["min_row"], r["min_col"], r["max_row"], r["max_col"]) == (1, 1, 20, 5)


def test_parse_ref_whole_column_now_supported():
    """全列参照は used range に対して展開される（issue #1）。"""
    r = parse_ref("Sheet1!A:E")
    assert r["ok"] is True
    assert r.get("whole_column") is True


def test_ranges_overlap_same_sheet():
    a = parse_ref("Sheet1!A1:C10")
    b = parse_ref("Sheet1!B5:D20")
    c = parse_ref("Sheet1!D20:E30")
    assert ranges_overlap(a, b) is True
    assert ranges_overlap(a, c) is False


def test_ranges_overlap_different_sheet_is_false():
    a = parse_ref("Sheet1!A1:C10")
    b = parse_ref("Sheet2!A1:C10")
    assert ranges_overlap(a, b) is False


def test_ranges_overlap_invalid_is_false():
    assert ranges_overlap(parse_ref("Sheet1!A1:B2"), {"ok": False}) is False


# ── 全列参照 (issue #1: raknaos 指摘) ──

def test_parse_whole_column_ref():
    r = parse_ref("Data!A:E", default_sheet="Data")
    assert r["ok"] is True
    assert r["sheet"] == "Data"
    assert r["min_col"] == 1 and r["max_col"] == 5
    assert r["min_row"] == 1 and r["max_row"] == 999999
    assert r.get("whole_column") is True


def test_whole_column_overlaps_bounded_range():
    """全列参照 A:E は bounded range A1:D7 と列範囲が重なるため overlap = True。"""
    wc = parse_ref("Data!A:E")
    bounded = parse_ref("Data!A1:D7")
    assert ranges_overlap(wc, bounded) is True


def test_whole_column_no_overlap_different_cols():
    wc = parse_ref("Data!F:J")
    bounded = parse_ref("Data!A1:D7")
    assert ranges_overlap(wc, bounded) is False


def test_whole_column_different_sheet_no_overlap():
    wc = parse_ref("Sheet2!A:E")
    bounded = parse_ref("Data!A1:D7")
    assert ranges_overlap(wc, bounded) is False
