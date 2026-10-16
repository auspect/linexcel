"""Fetch a pinned public Excel corpus without launching spreadsheet applications.

Only the source inventory and this helper belong in Git. Downloads, receipts and
the generated manifest belong in ignored validation_screenshots/. Existing files
with a different digest are rejected rather than overwritten. ZIP/XML inspection
counts stored formula records; it does not calculate cells or execute VBA/XLM.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import posixpath
import re
import tempfile
from datetime import UTC, datetime
from pathlib import Path, PurePosixPath
from typing import Any
from urllib.parse import urlsplit
from urllib.request import HTTPRedirectHandler, Request, build_opener
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

ROOT = Path(__file__).resolve().parents[1]
SOURCES = Path(__file__).with_name("public_corpus_sources.json")
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
OLE_SIGNATURE = bytes.fromhex("d0cf11e0a1b11ae1")
MAX_DOWNLOAD = 10 * 1024 * 1024
MAX_UNCOMPRESSED = 64 * 1024 * 1024
CONTENT_TYPES = {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    ".xlam": "application/vnd.ms-excel.addin.macroEnabled.main+xml",
}


def validate_source(spec: dict[str, Any]) -> None:
    """Require a local basename, HTTPS URL, bounded size and an exact digest."""
    name = spec["filename"]
    if not name or name in {".", ".."} or any(c in name for c in "/\\:"):
        raise ValueError(f"Invalid corpus basename: {name!r}")
    url = urlsplit(spec["source_url"])
    if url.scheme != "https" or not url.hostname or url.username or url.password:
        raise ValueError("Corpus source must use HTTPS without URL credentials")
    if not re.fullmatch(r"[0-9a-f]{64}", spec["sha256"]):
        raise ValueError(f"Missing or invalid expected SHA-256: {name}")
    if (
        not isinstance(spec["max_bytes"], int)
        or not 0 < spec["max_bytes"] <= MAX_DOWNLOAD
    ):
        raise ValueError(f"Invalid download bound: {name}")


class CorpusRedirects(HTTPRedirectHandler):
    """Follow only HTTPS redirects on the explicitly allowed upstream hosts."""

    def __init__(self, spec: dict[str, Any]):
        self.host = urlsplit(spec["source_url"]).hostname
        self.suffixes = spec.get("redirect_host_suffixes", [])
        self.redirects: list[str] = []

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        parsed = urlsplit(newurl)
        host = parsed.hostname or ""
        allowed = host == self.host or any(
            host == suffix or host.endswith("." + suffix) for suffix in self.suffixes
        )
        if (
            parsed.scheme != "https"
            or parsed.username
            or parsed.password
            or not allowed
            or len(self.redirects) >= 5
        ):
            raise ValueError(f"Refusing corpus redirect to {parsed.netloc}")
        self.redirects.append(newurl)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def acquire(spec: dict[str, Any], output: Path, *, offline: bool) -> dict[str, Any]:
    """Return a receipt after validating existing or freshly downloaded bytes."""
    validate_source(spec)
    path = output / spec["filename"]
    receipt: dict[str, Any] = {"path": path.name, "source_url": spec["source_url"]}
    if path.exists():
        if (
            path.is_symlink()
            or not path.is_file()
            or path.stat().st_size > spec["max_bytes"]
        ):
            raise ValueError(f"Invalid existing corpus file: {path}")
        data = path.read_bytes()
        receipt["status"] = "reused_verified"
    else:
        if offline:
            raise FileNotFoundError(f"Corpus file missing in offline mode: {path}")
        redirects = CorpusRedirects(spec)
        opener = build_opener(redirects)
        request = Request(
            spec["source_url"], headers={"User-Agent": "linexcel-public-corpus/1"}
        )
        with opener.open(request, timeout=45) as response:
            length = response.headers.get("Content-Length")
            if length and int(length) > spec["max_bytes"]:
                raise ValueError(f"Corpus download exceeds size bound: {path.name}")
            data = response.read(spec["max_bytes"] + 1)
            receipt.update(
                status="downloaded_verified",
                final_url=response.url,
                redirects=redirects.redirects,
                content_type=response.headers.get("Content-Type"),
                last_modified=response.headers.get("Last-Modified"),
            )
    if len(data) > spec["max_bytes"]:
        raise ValueError(f"Corpus file exceeds size bound: {path.name}")
    digest = hashlib.sha256(data).hexdigest()
    if digest != spec["sha256"]:
        raise ValueError(
            f"SHA-256 mismatch for {path.name}; original file not replaced"
        )
    if not path.exists():
        temporary = None
        try:
            with tempfile.NamedTemporaryFile(
                dir=output, prefix=path.name + ".", delete=False
            ) as tmp:
                temporary = Path(tmp.name)
                tmp.write(data)
            # A hard link publishes the verified bytes atomically without
            # replacing a file another process may have created meanwhile.
            try:
                os.link(temporary, path)
            except FileExistsError:
                pass
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)
    if (
        path.is_symlink()
        or not path.is_file()
        or path.stat().st_size > spec["max_bytes"]
        or hashlib.sha256(path.read_bytes()).hexdigest() != digest
    ):
        raise ValueError(f"Corpus file changed during acquisition: {path.name}")
    receipt.update(sha256=digest, size_bytes=len(data))
    return receipt


def inspect_container(path: Path) -> dict[str, Any]:
    """Inspect container metadata without external links, recalculation or macros."""
    if path.suffix.lower() == ".xla":
        with path.open("rb") as file:
            if file.read(8) != OLE_SIGNATURE:
                raise ValueError(f"Expected an OLE legacy Excel add-in: {path.name}")
        return {
            "container": "ole_biff",
            "sheet_count": None,
            "formula_count": None,
            "macro_presence": "requires_static_ole_vba_inspection",
            "inspection_limit": "Signature check only; no BIFF formulas or VBA modules decoded by this helper.",
        }
    expected_type = CONTENT_TYPES.get(path.suffix.lower())
    if expected_type is None:
        raise ValueError(f"Unsupported corpus extension: {path.suffix}")
    try:
        with ZipFile(path) as archive:
            entries = archive.infolist()
            if (
                len(entries) > 10000
                or sum(e.file_size for e in entries) > MAX_UNCOMPRESSED
            ):
                raise ValueError("Corpus ZIP exceeds static inspection bound")
            for entry in entries:
                parts = PurePosixPath(entry.filename)
                if entry.flag_bits & 1 or parts.is_absolute() or ".." in parts.parts:
                    raise ValueError("Unsafe or encrypted corpus ZIP member")
            names = archive.namelist()
            if len(names) != len(set(names)):
                raise ValueError("Duplicate corpus ZIP member")
            types = ET.fromstring(archive.read("[Content_Types].xml"))
            if not any(
                e.attrib.get("PartName") == "/xl/workbook.xml"
                and e.attrib.get("ContentType") == expected_type
                for e in types
            ):
                raise ValueError(
                    f"Container does not match {path.suffix} workbook type"
                )
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            relationships = {
                e.attrib["Id"]: e.attrib["Target"]
                for e in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
                if e.attrib.get("TargetMode") != "External"
            }
            sheets = []
            for sheet in workbook.find(NS + "sheets"):
                target = relationships[sheet.attrib[REL + "id"]]
                member = (
                    target.lstrip("/")
                    if target.startswith("/")
                    else posixpath.normpath(posixpath.join("xl", target))
                )
                if not member.startswith("xl/") or ".." in PurePosixPath(member).parts:
                    raise ValueError("Sheet relationship escapes the workbook package")
                xml = ET.fromstring(archive.read(member))
                cells = xml.findall(".//" + NS + "c")
                formulas = [
                    c.find(NS + "f") for c in cells if c.find(NS + "f") is not None
                ]
                dimension = xml.find(NS + "dimension")
                sheets.append(
                    {
                        "name": sheet.attrib["name"],
                        "state": sheet.attrib.get("state", "visible"),
                        "dimension": dimension.attrib.get("ref")
                        if dimension is not None
                        else None,
                        "stored_cell_count": len(cells),
                        "stored_value_count": sum(
                            c.find(NS + "v") is not None
                            or c.find(NS + "is") is not None
                            for c in cells
                        ),
                        "formula_count": len(formulas),
                        "formula_text_count": sum(bool(f.text) for f in formulas),
                        "shared_formula_follower_count": sum(
                            f.attrib.get("t") == "shared" and not f.text
                            for f in formulas
                        ),
                    }
                )
            vba = [n for n in names if n.lower().endswith("vbaproject.bin")]
            return {
                "container": "ooxml",
                "sheet_count": len(sheets),
                "sheets": sheets,
                "formula_count": sum(s["formula_count"] for s in sheets),
                "formula_count_basis": "Stored worksheet <f> records, including shared-formula followers; no recalculation.",
                "defined_name_count": len(workbook.findall(".//" + NS + "definedName")),
                "external_link_count": sum(
                    bool(re.fullmatch(r"xl/externalLinks/externalLink\d+\.xml", n))
                    for n in names
                ),
                "chart_count": sum(
                    bool(re.fullmatch(r"xl/charts/chart\d+\.xml", n)) for n in names
                ),
                "vba_projects": [
                    {
                        "member": n,
                        "size_bytes": archive.getinfo(n).file_size,
                        "sha256": hashlib.sha256(archive.read(n)).hexdigest(),
                    }
                    for n in vba
                ],
                "macro_presence": "vba_project_present"
                if vba
                else "no_vba_project_member",
            }
    except (BadZipFile, KeyError, ET.ParseError) as error:
        raise ValueError(f"Invalid Excel package {path.name}: {error}") from error


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--sources", type=Path, default=SOURCES)
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=ROOT / "validation_screenshots/public_corpus_20261009",
    )
    parser.add_argument(
        "--offline",
        action="store_true",
        help="Verify already downloaded files without network access",
    )
    args = parser.parse_args()
    inventory_bytes = args.sources.read_bytes()
    inventory = json.loads(inventory_bytes.decode("utf-8-sig"))
    if inventory.get("version") != 1 or not inventory.get("files"):
        raise ValueError("Expected a non-empty version 1 source inventory")
    output = args.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=True)
    manifests = []
    receipts = []
    seen_names: set[str] = set()
    for spec in inventory["files"]:
        if spec.get("role") not in {"workbook", "addin"}:
            raise ValueError("Corpus file role must be workbook or addin")
        for item in [spec, *spec.get("support_files", [])]:
            if item["filename"].casefold() in seen_names:
                raise ValueError("Duplicate corpus output filename")
            seen_names.add(item["filename"].casefold())
            receipts.append(acquire(item, output, offline=args.offline))
        metadata = {
            k: v
            for k, v in spec.items()
            if k
            not in {"filename", "support_files", "max_bytes", "redirect_host_suffixes"}
        }
        metadata.update(
            path=spec["filename"],
            size_bytes=(output / spec["filename"]).stat().st_size,
            inspection=inspect_container(output / spec["filename"]),
            support_files=[
                {
                    "path": s["filename"],
                    "source_url": s["source_url"],
                    "sha256": s["sha256"],
                }
                for s in spec.get("support_files", [])
            ],
        )
        manifests.append(metadata)
        print(
            f"{spec['id']}: verified {metadata['size_bytes']} bytes; {metadata['inspection']['sheet_count']} sheets; {metadata['inspection']['formula_count']} stored formula records"
        )
    for receipt in receipts:
        if (
            hashlib.sha256((output / receipt["path"]).read_bytes()).hexdigest()
            != receipt["sha256"]
        ):
            raise ValueError(
                f"Corpus file changed during inspection: {receipt['path']}"
            )
    manifest = {
        "version": 1,
        "created_at": datetime.now(UTC).isoformat(),
        "source_inventory_sha256": hashlib.sha256(inventory_bytes).hexdigest(),
        "inspection_policy": "Static ZIP/XML or OLE signature only; no Excel application, link refresh, formula evaluation or macro execution.",
        "files": manifests,
    }
    (output / "manifest.json").write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )
    (output / "download_receipts.json").write_text(
        json.dumps(receipts, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )
    print(f"Manifest: {output / 'manifest.json'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
