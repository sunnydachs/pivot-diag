"""テスト用の合成 xlsx ビルダー（stdlib の zipfile だけで OOXML を組み立てる）。

Excel が書き出す構造（workbook / sheets / pivotTables / pivotCache / rels）を
最小構成で再現する。openpyxl が生成できないピボット構造を柔軟に作るための
自作フィクスチャで、実装の検証に使う。
"""
import zipfile

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
{pivot_overrides}
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""


def workbook_xml(sheets):
    rows = "".join(
        f'<sheet name="{name}" sheetId="{i+1}" r:id="rId{i+1}"/>'
        for i, name in enumerate(sheets)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>{rows}</sheets>
</workbook>'''


def workbook_rels(sheets):
    rows = "".join(
        f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i+1}.xml"/>'
        for i in range(len(sheets))
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rows}</Relationships>'''


def sheet_xml():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c></row></sheetData>
</worksheet>'''


def pivot_table_xml(name, location_ref, cache_id, cache_rid="rId1"):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 name="{name}" cacheId="{cache_id}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">
<location ref="{location_ref}" firstHeaderRow="1" firstDataRow="2" firstDataCol="0"/>
<pivotFields count="1"><pivotField showAll="0"/></pivotFields>
<rowFields count="1"><field x="0"/></rowFields>
<rowItems count="1"><i><x/></i></rowItems>
<colItems count="1"><i/></colItems>
<dataFields count="1"><dataField name="Sum of Amount" fld="1" baseField="0" baseItem="0"/></dataFields>
<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>
</pivotTableDefinition>'''


def pivot_table_rels(cache_target, sheet_target="../worksheets/sheet1.xml"):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="{cache_target}"/>
</Relationships>'''


def pivot_cache_xml(cache_id, src_sheet, src_ref):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 r:id="rId1" refreshOnLoad="1" refreshedBy="fixture" createdVersion="8" minRefreshableVersion="3" recordCount="6" cacheId="{cache_id}">
<cacheSource type="worksheet"><worksheetSource ref="{src_ref}" sheet="{src_sheet}"/></cacheSource>
<cacheFields count="2"><cacheField name="Region" numFmtId="0"><sharedItems count="1"><s v="North"/></sharedItems></cacheField><cacheField name="Amount" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="100" maxValue="300"/></cacheField></cacheFields>
</pivotCacheDefinition>'''


def build_workbook(path, pivots, sheets=("Data", "P1")):
    """pivots: [{name, location, cache_id, src_sheet, src_ref}] で xlsx を組む。

    全ピボットの cache は1つ（cacheId=1）を共有する構成で生成する。
    """
    n_sheets = len(sheets)
    pivot_overrides = "".join(
        f'<Override PartName="/xl/pivotCache/pivotCacheDefinition{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>'
        f'<Override PartName="/xl/pivotTables/pivotTable{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>'
        for i in range(1, len(pivots) + 1)
    )
    content_types = CONTENT_TYPES.replace("{pivot_overrides}", pivot_overrides)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", workbook_xml(sheets))
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels(sheets))
        for i in range(n_sheets):
            z.writestr(f"xl/worksheets/sheet{i+1}.xml", sheet_xml())
            # sheet1 のみピボットを持つ体で rels を張る
            if i == 0:
                rows = "".join(
                    f'<Relationship Id="rIdPT{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable{i+1}.xml"/>'
                    for i in range(len(pivots))
                )
                z.writestr(f"xl/worksheets/_rels/sheet{i+1}.xml.rels", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rows}</Relationships>')
        for i, p in enumerate(pivots, 1):
            z.writestr(f"xl/pivotTables/pivotTable{i}.xml",
                       pivot_table_xml(p["name"], p["location"], p["cache_id"]))
            # キャッシュはピボットごとに独立して書く（ソース範囲の違いを再現するため）
            z.writestr(f"xl/pivotTables/_rels/pivotTable{i}.xml.rels",
                       pivot_table_rels(cache_target=f"../pivotCache/pivotCacheDefinition{i}.xml"))
            z.writestr(f"xl/pivotCache/pivotCacheDefinition{i}.xml",
                       pivot_cache_xml(p["cache_id"], p["src_sheet"], p["src_ref"]))
