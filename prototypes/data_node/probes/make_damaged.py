"""Damaged copies of real corpus workbooks for failure-mode measurements.

Never touches the originals; writes to <out-dir>. Each case changes one part.
usage: LINEXCEL_LAB_WORKBOOKS=<linexcel-lab>/workbooks make_damaged.py <out-dir>
"""

import json
import os
import shutil
import struct
import sys
import zipfile
from pathlib import Path

# The public corpus folder (linexcel-lab/workbooks), never the repository.
LAB = Path(os.environ["LINEXCEL_LAB_WORKBOOKS"])
OUT = Path(sys.argv[1])
OUT.mkdir(parents=True, exist_ok=True)
ONS = LAB / "public" / "ons-mye23tablesuk.xlsx"
ATB = LAB / "public" / "nrel-atb-2024-v4.xlsx"
QRA = LAB / "public" / "qra-template.xlsm"


def sheet_part(path, name):
    import posixpath
    from xml.etree import ElementTree as ET

    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = {
            r.get("Id"): r.get("Target")
            for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        }
        for s in wb.iter():
            if s.tag.endswith("}sheet") and s.get("name") == name:
                rid = next(v for k, v in s.attrib.items() if k.endswith("}id"))
                t = rels[rid]
                return (
                    t.lstrip("/")
                    if t.startswith("/")
                    else posixpath.normpath("xl/" + t)
                )
    raise KeyError(name)


def rewrite(src, dst, edit):
    """Copy a ZIP member by member; edit(name, data) -> bytes | None (drop)."""
    with (
        zipfile.ZipFile(src) as zin,
        zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout,
    ):
        for info in zin.infolist():
            data = edit(info.filename, zin.read(info.filename))
            if data is not None:
                zout.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED)


def corrupt_member_bytes(path, member):
    """Flip compressed bytes in the middle of one member (CRC / inflate error)."""
    with zipfile.ZipFile(path) as z:
        info = z.getinfo(member)
    with open(path, "r+b") as f:
        f.seek(info.header_offset)
        header = f.read(30)
        name_len, extra_len = struct.unpack("<HH", header[26:30])
        start = info.header_offset + 30 + name_len + extra_len
        middle = start + info.compress_size // 2
        f.seek(middle)
        chunk = bytearray(f.read(64))
        for i in range(len(chunk)):
            chunk[i] ^= 0x5A
        f.seek(middle)
        f.write(chunk)


cases = {}


def case(name, base, note, **extra):
    cases[name] = {
        "file": str(OUT / f"{name}{base.suffix}"),
        "base": base.name,
        "note": note,
        **extra,
    }
    return OUT / f"{name}{base.suffix}"


persons = sheet_part(ONS, "MYE2 - Persons")
atb_sheet = sheet_part(ATB, "Land-Based Wind")
truncate = lambda part: lambda n, d: d[: len(d) // 2] if n == part else d  # noqa: E731

rewrite(
    ONS, case("ons_truncated_sheet", ONS, f"{persons} cut at 50%"), truncate(persons)
)
p = case("ons_corrupt_sheet_bytes", ONS, f"{persons} compressed bytes flipped")
shutil.copyfile(ONS, p)
corrupt_member_bytes(p, persons)
rewrite(
    ONS,
    case("ons_missing_sheet_part", ONS, f"{persons} removed, relationship kept"),
    lambda n, d: None if n == persons else d,
)
rewrite(
    ONS,
    case("ons_truncated_styles", ONS, "xl/styles.xml cut at 50%"),
    truncate("xl/styles.xml"),
)
rewrite(
    ONS,
    case("ons_truncated_shared_strings", ONS, "xl/sharedStrings.xml cut at 50%"),
    truncate("xl/sharedStrings.xml"),
)
rewrite(
    ONS,
    case("ons_missing_shared_strings", ONS, "xl/sharedStrings.xml removed"),
    lambda n, d: None if n == "xl/sharedStrings.xml" else d,
)
rewrite(
    ONS,
    case("ons_truncated_table", ONS, "xl/tables/table3.xml cut at 50%"),
    truncate("xl/tables/table3.xml"),
)
rewrite(
    ONS,
    case("ons_missing_workbook_rels", ONS, "xl/_rels/workbook.xml.rels removed"),
    lambda n, d: None if n == "xl/_rels/workbook.xml.rels" else d,
)
rewrite(
    ONS,
    case("ons_truncated_workbook_xml", ONS, "xl/workbook.xml cut at 50%"),
    truncate("xl/workbook.xml"),
)
data = ONS.read_bytes()
case(
    "ons_truncated_zip", ONS, "whole file cut at 60% (no central directory)"
).write_bytes(data[: int(len(data) * 0.6)])

rewrite(
    ATB,
    case("atb_truncated_sheet", ATB, f"{atb_sheet} cut at 50%", cell="Summary!E5"),
    truncate(atb_sheet),
)
p = case(
    "atb_corrupt_sheet_bytes",
    ATB,
    f"{atb_sheet} compressed bytes flipped",
    cell="Summary!E5",
)
shutil.copyfile(ATB, p)
corrupt_member_bytes(p, atb_sheet)
rewrite(
    ATB,
    case(
        "atb_truncated_shared_strings",
        ATB,
        "xl/sharedStrings.xml cut at 50%",
        cell="Summary!E5",
    ),
    truncate("xl/sharedStrings.xml"),
)
shutil.copyfile(
    QRA,
    case(
        "qra_external_link_absent",
        QRA,
        "real external link to an absent https workbook (unchanged copy)",
    ),
)

(OUT / "cases.json").write_text(json.dumps(cases, indent=1), "utf-8")
print(json.dumps(cases, indent=1))
