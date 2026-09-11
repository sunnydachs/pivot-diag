"""ooxml — xlsx/xlsm（OOXML zip）からピボット定義を読み出す。

読み取るパーツ（Excel の実ファイル構造）:
  xl/workbook.xml                          … シート名 → r:id
  xl/_rels/workbook.xml.rels               … r:id → worksheet パス
  xl/worksheets/_rels/sheetN.xml.rels      … worksheet → pivotTable パーツ（所属特定）
  xl/pivotTables/pivotTableN.xml           … name / location ref / cacheId
  xl/pivotTables/_rels/pivotTableN.xml.rels… … pivotTable → pivotCacheDefinition
  xl/pivotCache/pivotCacheDefinitionN.xml  … cacheSource（worksheetSource ref/sheet）

XML 解析は ElementTree（標準ライブラリ）。サードパーティ依存は無し。
"""
import posixpath
import zipfile
import xml.etree.ElementTree as ET

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _read(zf: zipfile.ZipFile, part: str):
    try:
        return ET.fromstring(zf.read(part))
    except KeyError:
        return None
    except ET.ParseError:
        return None


def _norm(base_dir: str, target: str) -> str:
    """base_dir（ディレクトリ）基準で target をパッケージ内絶対パスに正規化する。"""
    return posixpath.normpath(posixpath.join(base_dir, target)).lstrip("/")


def read_relationships(zf: zipfile.ZipFile, source_part: str) -> dict:
    """source_part に対応する .rels の {Id: Target絶対パス} を返す（無ければ {}）。

    Target は OPC 仕様どおり「ソースパーツのディレクトリ」基準で解決する
    （_rels ディレクトリは格納場所にすぎず、基準にはならない）。
    """
    d, base = posixpath.split(source_part)
    base_dir = d if d else "."
    rels_part = posixpath.join(d, "_rels", base + ".rels") if d else f"_rels/{base}.rels"
    root = _read(zf, rels_part)
    if root is None:
        return {}
    out = {}
    for rel in root.findall(f"{NS_PKG_REL}Relationship"):
        target = rel.get("Target", "")
        if rel.get("TargetMode") == "External":
            continue
        out[rel.get("Id")] = _norm(base_dir, target)
    return out


def read_sheet_parts(zf: zipfile.ZipFile) -> dict:
    """{sheet_name: worksheet_part_path} を返す（workbook.xml + rels）。"""
    wb = _read(zf, "xl/workbook.xml")
    if wb is None:
        return {}
    rels = read_relationships(zf, "xl/workbook.xml")
    out = {}
    for sh in wb.findall(f".//{NS_MAIN}sheet"):
        name = sh.get("name")
        rid = sh.get(f"{NS_R}id")
        target = rels.get(rid)
        if name and target:
            out[name] = target
    return out


def read_pivot_tables(zf: zipfile.ZipFile, sheet_parts: dict) -> list:
    """全 pivotTable パーツを読み、所属シートつきで返す。

    返り値: [{name, location_ref, cache_id, part, worksheet}]
    worksheet は「そのピボットが置かれているシート名」— pivotTable パーツ自体は
    シート名を持たないため、worksheet 側 rels（→ pivotTable パーツ）の逆引きで
    対応づける。該当シートが見つからない場合は None。
    """
    out = []
    parts = sorted(n for n in zf.namelist()
                   if n.startswith("xl/pivotTables/pivotTable")
                   and n.endswith(".xml"))
    for part in parts:
        root = _read(zf, part)
        if root is None:
            continue
        loc = root.find(f"{NS_MAIN}location")
        cache_id = root.get("cacheId")

        # 所属シート: (a) シート側 rels 逆引き (b) pivotTable 側 rels → worksheet 逆引き
        worksheet = None
        for sheet_name, sheet_part in sheet_parts.items():
            rels = read_relationships(zf, sheet_part)
            if part in set(rels.values()):
                worksheet = sheet_name
                break
        if worksheet is None:
            pt_rels = read_relationships(zf, part)
            for target in pt_rels.values():
                for sheet_name, sheet_part in sheet_parts.items():
                    if target == sheet_part:
                        worksheet = sheet_name
                        break

        out.append({
            "name": root.get("name") or part,
            "location_ref": loc.get("ref") if loc is not None else None,
            "cache_id": cache_id,
            "part": part,
            "worksheet": worksheet,
        })
    return out


def read_pivot_cache_parts(zf: zipfile.ZipFile, pivots: list) -> dict:
    """{pivotTable part: cache part} を pivotTable 側 rels から解決し、
    ついでに全 cacheSource も返す。

    返り値: (cache_part_by_pivot: {part: cache_part},
             cache_sources: {cache_part: {"type", "sheet", "ref", "name"}})
    """
    cache_parts = sorted(n for n in zf.namelist()
                         if n.startswith("xl/pivotCache/pivotCacheDefinition")
                         and n.endswith(".xml"))
    cache_by_id = {}
    for cp in cache_parts:
        root = _read(zf, cp)
        if root is not None and root.get("cacheId") is not None:
            cache_by_id[root.get("cacheId")] = cp

    cache_part_by_pivot = {}
    for p in pivots:
        part, cache_id = p["part"], p.get("cache_id")
        rels = read_relationships(zf, part)
        cache_part = next((t for t in rels.values() if "pivotCacheDefinition" in t), None)
        if cache_part is None and cache_id is not None:
            cache_part = cache_by_id.get(cache_id)
        cache_part_by_pivot[part] = cache_part

    cache_sources = {}
    for cp in cache_parts:
        root = _read(zf, cp)
        src = root.find(f"{NS_MAIN}cacheSource") if root is not None else None
        if src is None:
            cache_sources[cp] = {"type": None}
            continue
        ws = src.find(f"{NS_MAIN}worksheetSource")
        cache_sources[cp] = {
            "type": src.get("type"),
            "sheet": ws.get("sheet") if ws is not None else None,
            "ref": ws.get("ref") if ws is not None else None,
            "name": ws.get("name") if ws is not None else None,
        }
    return cache_part_by_pivot, cache_sources


def extract_pivots(xlsx_path: str) -> dict:
    """xlsx を開いてピボット定義を収集する。

    返り値: {sheets: [name...], pivots: [{name, worksheet, location_ref,
              source: {"type", "sheet", "ref", "name"}, part, cache_part}],
             cache_sources, parts_error}
    """
    pivots, cache_sources, sheets = [], {}, []
    parts_error = None
    with zipfile.ZipFile(xlsx_path) as zf:
        sheets = list(read_sheet_parts(zf).keys())
        pivots = read_pivot_tables(zf, read_sheet_parts(zf))
        cache_part_by_pivot, cache_sources = read_pivot_cache_parts(zf, pivots)

    for p in pivots:
        cp = cache_part_by_pivot.get(p["part"])
        p["cache_part"] = cp
        p["source"] = cache_sources.get(cp, {"type": None})

    return {
        "sheets": sheets,
        "pivots": pivots,
        "cache_sources": cache_sources,
        "parts_error": parts_error,
    }
