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


def test_parse_ref_whole_column_is_invalid_for_mvp():
    r = parse_ref("Sheet1!A:E")
    assert r["ok"] is False


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
