"""Static structure of corpus workbooks: parts, sizes, dimensions, tables.

Reads only ZIP directory entries and small XML parts; never a worksheet body
beyond its first 4 KiB (for <dimension>).

usage: corpus_structure.py out.json file.xlsx [...]
"""

import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def rels(archive, part):
    import posixpath

    name = posixpath.join(
        posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels"
    )
    if name not in archive.namelist():
        return {}
    out = {}
    for r in ET.fromstring(archive.read(name)):
        target = r.get("Target", "")
        if not target.startswith("/"):
            target = posixpath.normpath(posixpath.join(posixpath.dirname(part), target))
        else:
            target = target.lstrip("/")
        out[r.get("Id")] = (r.get("Type", "").rsplit("/", 1)[-1], target)
    return out


def main():
    out = Path(sys.argv[1])
    report = {}
    for file in map(Path, sys.argv[2:]):
        with zipfile.ZipFile(file) as archive:
            info = {
                i.filename: (i.file_size, i.compress_size) for i in archive.infolist()
            }
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            wb_rels = rels(archive, "xl/workbook.xml")
            sheets = []
            for sheet in workbook.findall("m:sheets/m:sheet", NS):
                rid = next(v for k, v in sheet.attrib.items() if k.endswith("}id"))
                kind, part = wb_rels.get(rid, ("?", None))
                entry = {
                    "name": sheet.get("name"),
                    "state": sheet.get("state", "visible"),
                    "part": part,
                }
                if part in info:
                    entry["xmlBytes"], entry["zipBytes"] = info[part]
                    with archive.open(part) as handle:
                        head = handle.read(4096).decode("utf-8", "replace")
                    match = re.search(r'<(?:\w+:)?dimension ref="([^"]+)"', head)
                    entry["dimension"] = match[1] if match else None
                    tables = []
                    for _, (kind2, target) in rels(archive, part).items():
                        if kind2 == "table" and target in info:
                            t = ET.fromstring(archive.read(target))
                            tables.append(
                                {
                                    "name": t.get("displayName") or t.get("name"),
                                    "ref": t.get("ref"),
                                    "columns": len(
                                        t.findall("m:tableColumns/m:tableColumn", NS)
                                    ),
                                    "calculatedColumns": sum(
                                        1
                                        for c in t.findall(
                                            "m:tableColumns/m:tableColumn", NS
                                        )
                                        if c.find("m:calculatedColumnFormula", NS)
                                        is not None
                                    ),
                                    "headerRowCount": t.get("headerRowCount", "1"),
                                    "totalsRowCount": t.get("totalsRowCount", "0"),
                                }
                            )
                    entry["tables"] = tables
                sheets.append(entry)
            styles = info.get("xl/styles.xml", (0, 0))[0]
            strings = info.get("xl/sharedStrings.xml", (0, 0))[0]
            report[file.name] = {
                "bytes": file.stat().st_size,
                "uncompressedBytes": sum(v[0] for v in info.values()),
                "sharedStringsXmlBytes": strings,
                "stylesXmlBytes": styles,
                "definedNames": len(
                    workbook.findall("m:definedNames/m:definedName", NS)
                ),
                "externalLinks": len(
                    [n for n in info if n.startswith("xl/externalLinks/externalLink")]
                ),
                "sheets": sheets,
            }
    out.write_text(json.dumps(report, indent=1, ensure_ascii=False), "utf-8")
    for name, r in report.items():
        mb = r["bytes"] / 1e6
        xml = r["uncompressedBytes"] / 1e6
        sst = r["sharedStringsXmlBytes"] / 1e6
        styles = r["stylesXmlBytes"] / 1e3
        print(
            f"== {name}: {mb:.1f} MB zip, {xml:.1f} MB xml, sst {sst:.1f} MB, "
            f"styles {styles:.0f} kB, {len(r['sheets'])} sheets, "
            f"{r['definedNames']} names, {r['externalLinks']} ext links"
        )
        big = sorted(r["sheets"], key=lambda s: -s.get("xmlBytes", 0))[:6]
        for s in big:
            tables = [
                (t["name"], t["ref"], t["calculatedColumns"])
                for t in s.get("tables", [])
            ]
            size = s.get("xmlBytes", 0) / 1e6
            print(f"   {s['name']!r}: {size:.1f} MB xml, dim {s.get('dimension')}")
            print(f"      tables {tables}")


if __name__ == "__main__":
    main()
