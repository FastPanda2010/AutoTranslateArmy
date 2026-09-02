from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from translate_army import (
    build_filter_maps,
    infer_faction_id,
    profile_table_name_sources,
    sanitize_unit_image_name,
    should_add_unit_options_table,
    should_include_unit,
    unit_category_sort_key,
)


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}


def expected_table_names(json_path: Path) -> list[str]:
    data = json.loads(json_path.read_text(encoding="utf-8"))
    faction_id = infer_faction_id(json_path, data)
    maps = build_filter_maps(data)
    units = [
        (index, unit)
        for index, unit in enumerate(data.get("units", []))
        if should_include_unit(unit, faction_id)
    ]
    units.sort(key=lambda item: unit_category_sort_key(item[1], maps, item[0]))

    names: list[str] = []
    for _, unit in units:
        if should_add_unit_options_table(unit):
            names.append(unit.get("isc") or unit.get("name", ""))
        for pg in unit.get("profileGroups", []) or []:
            profiles = pg.get("profiles") or []
            if profiles:
                english_name, _ = profile_table_name_sources(unit, pg, profiles[0])
                names.append(english_name)
    return [name for name in names if name]


def relationship_targets(docx: zipfile.ZipFile) -> dict[str, str]:
    rels = ET.fromstring(docx.read("word/_rels/document.xml.rels"))
    return {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}


def zip_member_for_target(target: str) -> str:
    return posixpath.normpath(posixpath.join("word", target))


def cell_text(cell: ET.Element) -> str:
    return "".join(t.text or "" for t in cell.findall(".//w:t", NS))


def image_entries(cell: ET.Element, rels: dict[str, str], docx: zipfile.ZipFile) -> list[dict[str, Any]]:
    entries: list[dict[str, Any]] = []
    for container in cell.findall(".//wp:inline", NS) + cell.findall(".//wp:anchor", NS):
        blip = container.find(".//a:blip", NS)
        if blip is None:
            continue
        rid = blip.attrib.get(f"{{{NS['r']}}}embed")
        if not rid or rid not in rels:
            continue

        target = rels[rid]
        member = zip_member_for_target(target)
        data = docx.read(member)
        extent = container.find("wp:extent", NS)
        entries.append(
            {
                "target": target,
                "member": member,
                "data": data,
                "sha256": hashlib.sha256(data).hexdigest(),
                "width_emu": int(extent.attrib.get("cx", 0)) if extent is not None else 0,
                "height_emu": int(extent.attrib.get("cy", 0)) if extent is not None else 0,
                "extension": Path(member).suffix.lower() or ".png",
            }
        )
    return entries


def match_expected_name(title: str, expected: list[str], start_index: int) -> tuple[int, str] | None:
    for index in range(start_index, len(expected)):
        english_name = expected[index]
        if english_name and english_name in title:
            return index + 1, english_name
    return None


def save_manifest_table(
    output_dir: Path,
    counters: dict[str, int],
    english_name: str,
    images: list[dict[str, Any]],
) -> dict[str, Any]:
    stem = sanitize_unit_image_name(english_name)
    counters[stem] = counters.get(stem, 0) + 1
    occurrence = counters[stem]
    saved_images = []
    for image_index, image in enumerate(images, start=1):
        suffix_parts = []
        if occurrence > 1:
            suffix_parts.append(str(occurrence))
        if len(images) > 1:
            suffix_parts.append(str(image_index))
        suffix = f"-{'-'.join(suffix_parts)}" if suffix_parts else ""
        filename = f"{stem}{suffix}{image['extension']}"
        (output_dir / filename).write_bytes(image["data"])
        saved_images.append(
            {
                "file": filename,
                "width_emu": image["width_emu"],
                "height_emu": image["height_emu"],
                "sha256": image["sha256"],
            }
        )

    return {"english_name": english_name, "images": saved_images}


def extract_unit_images(docx_path: Path, json_path: Path, output_dir: Path) -> dict[str, Any]:
    expected = expected_table_names(json_path)
    manifest_tables: list[dict[str, Any]] = []
    expected_index = 0
    counters: dict[str, int] = {}

    output_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(docx_path) as docx:
        rels = relationship_targets(docx)
        root = ET.fromstring(docx.read("word/document.xml"))
        body = root.find("w:body", NS)
        if body is None:
            raise ValueError("document.xml has no body")

        for table in [el for el in list(body) if el.tag.endswith("}tbl")]:
            cells = table.findall(".//w:tc", NS)
            if not cells:
                continue
            for cell in cells:
                title = cell_text(cell)
                images = image_entries(cell, rels, docx)
                if not title or not images:
                    continue

                match = match_expected_name(title, expected, expected_index)
                if match is None:
                    continue
                expected_index, english_name = match
                manifest_tables.append(save_manifest_table(output_dir, counters, english_name, images))

    manifest = {
        "source_docx": str(docx_path),
        "source_json": str(json_path),
        "tables": manifest_tables,
    }
    (output_dir / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return manifest


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract manually inserted unit images from a generated DOCX.")
    parser.add_argument("docx", type=Path, help="DOCX containing manually inserted unit images.")
    parser.add_argument("json", type=Path, help="Matching Infinity Army JSON file.")
    parser.add_argument("--out-dir", type=Path, default=None, help="Output image directory.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output_dir = args.out_dir or Path("Asset") / "unit_images" / args.json.stem
    manifest = extract_unit_images(args.docx, args.json, output_dir)
    image_count = sum(len(table["images"]) for table in manifest["tables"])
    print(f"extracted {image_count} images for {len(manifest['tables'])} table headers into {output_dir}")


if __name__ == "__main__":
    main()
