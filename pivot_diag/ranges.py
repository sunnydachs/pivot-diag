"""ranges — Excel セル範囲参照（"Sheet1!A1:E20" 等）のパースと重なり判定（純関数）。

対応: Sheet!A1:E20 / A1:E20 / A1（単一セル）/ ' quoted sheet'!A1:E20
非対応（invalid として扱う）: 全列参照 A:E、全行参照 1:5、名前付き範囲
（いずれも MVP の対象外。無理に解釈せず invalid として報告する）
"""
import re

_CELL = r"\$?([A-Za-z]{1,3})\$?([0-9]{1,7})"
_REF_RE = re.compile(
    rf"^(?:(?:'(?P<qs>[^']+)'|(?P<raw>[^'!:]+))!)?{_CELL}(?::{_CELL})?$"
)
# 全列参照 (A:E) 用（issue #1: raknaos 指摘）
_WC_RE = re.compile(
    rf"^(?:(?:'(?P<qs>[^']+)'|(?P<raw>[^'!:]+))!)?\$?(?P<c1>[A-Za-z]{{1,3}}):\$?(?P<c2>[A-Za-z]{{1,3}})$"
)
MAX_ROW = 999999  # 全列参照の実効的な行数上限（重なり判定用の境界値）


def col_to_index(letters: str) -> int:
    """A=1, B=2, ... Z=26, AA=27（1-based）。"""
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def parse_ref(ref: str, default_sheet: str = None) -> dict:
    """範囲参照をパースする（純関数）。

    返り値: {ok: True, sheet, min_row, min_col, max_row, max_col}
    失敗時: {ok: False, error: "..."}
    全列参照 (A:E) は used range 全体とみなして行範囲を 1〜MAX_ROW に展開する。
    """
    if not ref or not ref.strip():
        return {"ok": False, "error": "empty ref"}
    clean = ref.strip().replace("$", "")
    m = _REF_RE.match(clean)
    if m:
        sheet = m.group("qs") or m.group("raw") or default_sheet
        c1, r1 = col_to_index(m.group(3)), int(m.group(4))
        c2 = col_to_index(m.group(5)) if m.group(5) else c1
        r2 = int(m.group(6)) if m.group(6) else r1
        return {
            "ok": True,
            "sheet": sheet,
            "min_row": min(r1, r2), "max_row": max(r1, r2),
            "min_col": min(c1, c2), "max_col": max(c1, c2),
        }
    # 全列参照 fallback (issue #1)
    wc = _WC_RE.match(clean)
    if wc:
        sheet = wc.group("qs") or wc.group("raw") or default_sheet
        c1, c2 = col_to_index(wc.group("c1")), col_to_index(wc.group("c2"))
        return {
            "ok": True,
            "sheet": sheet,
            "min_row": 1, "max_row": MAX_ROW,
            "min_col": min(c1, c2), "max_col": max(c1, c2),
            "whole_column": True,
        }
    return {"ok": False, "error": f"unparseable ref: {ref!r}"}


def ranges_overlap(a: dict, b: dict) -> bool:
    """同じシート上の2つのパース済み範囲が重なるか（純関数）。

    どちらかが ok=False の場合は False（判定不能なものを重なり扱いしない）。
    """
    if not (a.get("ok") and b.get("ok")):
        return False
    if (a["sheet"] or "") != (b["sheet"] or ""):
        return False
    return not (
        a["max_row"] < b["min_row"] or b["max_row"] < a["min_row"]
        or a["max_col"] < b["min_col"] or b["max_col"] < a["min_col"]
    )
