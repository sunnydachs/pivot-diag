"""scanner + cli — ワークブック/ディレクトリを走査して診断レポートを出す（読み取り専用）。

usage:
  pivot-diag report.xlsx            # 1 ファイルを診断
  pivot-diag ./                     # ディレクトリ内の *.xlsx/*.xlsm を一括診断
  pivot-diag report.xlsx --json
"""
import argparse
import json
import sys
from pathlib import Path

from pivot_diag.diagnostics import diagnose
from pivot_diag.ooxml import extract_pivots

MARK = {
    "overlapping_sources": "⚠️ OVERLAPPING SOURCES",
    "shared_source": "· SHARED SOURCE (informational)",
    "external_source": "· EXTERNAL/NAMED SOURCE (informational)",
    "unparseable_ref": "· UNPARSEABLE REF (informational)",
    "missing_source": "❌ MISSING SOURCE",
}


def scan_file(path: Path) -> dict:
    """1 ファイルを診断。壊れた zip / OOXML でない場合は skipped として返す。"""
    try:
        extracted = extract_pivots(str(path))
    except zipfile.BadZipFile:
        return {"file": path.name, "skipped": "not a valid OOXML zip"}
    except OSError as e:
        return {"file": path.name, "skipped": f"unreadable: {e}"}
    if not extracted["pivots"]:
        return {"file": path.name, "skipped": "no pivot tables"}
    diag = diagnose(extracted)
    return {
        "file": path.name,
        "sheets": extracted["sheets"],
        "pivots": [{
            "name": p["name"],
            "worksheet": p.get("worksheet"),
            "location": p.get("location_ref"),
            "source_sheet": (p.get("source") or {}).get("sheet"),
            "source_ref": (p.get("source") or {}).get("ref"),
        } for p in extracted["pivots"]],
        **diag,
    }


def scan(target: str) -> list:
    t = Path(target)
    if t.is_file():
        files = [t]
    else:
        files = sorted(list(t.rglob("*.xlsx")) + list(t.rglob("*.xlsm")))
        files = [f for f in files if not any(part.startswith(("$", "~"))
                                             for part in f.parts) and "$" not in f.name]
    return [scan_file(f) for f in files]


def render(results: list, json_output: bool) -> str:
    if json_output:
        return json.dumps(results, ensure_ascii=False, indent=2)

    lines = ["pivot-diag — read-only pivot configuration audit", ""]
    for r in results:
        if r.get("skipped"):
            lines.append(f"{r['file']}: skipped ({r['skipped']})")
            continue
        lines.append(f"{r['file']}: {r['pivots_total']} pivot(s), "
                     f"{r['sources_ok']} bounded source range(s)")
        for p in r["pivots"]:
            loc = f"@{p['worksheet']}!{p['location']}" if p.get("worksheet") else ""
            lines.append(f"  - {p['name']} {loc} "
                         f"← source {p.get('source_sheet')}!{p.get('source_ref')}")
        if not r["findings"]:
            lines.append("  no issues found.")
        for f in r["findings"]:
            lines.append(f"  {MARK.get(f['status'], f['status'])}")
            lines.append(f"    {f['detail']}")
        lines.append("")
    n_issues = sum(len(r.get("findings", [])) for r in results
                   if not r.get("skipped"))
    lines.append(f"summary: {len(results)} file(s) scanned, {n_issues} finding(s)")
    return "\n".join(lines)


def main(argv=None) -> int:
    import zipfile  # scan_file で使う
    ap = argparse.ArgumentParser(
        prog="pivot-diag",
        description="Audit pivot table configurations in Excel workbooks. Read-only.")
    ap.add_argument("target", help=".xlsx/.xlsm file or directory")
    ap.add_argument("--json", action="store_true", help="machine-readable output")
    args = ap.parse_args(argv)
    results = scan(args.target)
    print(render(results, json_output=args.json))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
