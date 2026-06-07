from __future__ import annotations

import argparse
import csv
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


ATTR_LABELS = [
    ("move", "移动力\nMOV"),
    ("cc", "近战\nCC"),
    ("bs", "射击\nBS"),
    ("ph", "体格\nPH"),
    ("wip", "意志\nWIP"),
    ("arm", "护甲\nARM"),
    ("bts", "护盾\nBTS"),
    ("w", None),
    ("s", "轮廓值\nS"),
    ("ava", "可用数\nAVA"),
]

FILTER_CATEGORY = {
    "skills": "skill",
    "weapons": "weapon",
    "equip": "equip",
    "chars": "char",
    "type": "type",
    "category": "category",
    "extras": "extra",
}

NBSP = "\u00a0"


@dataclass
class Translator:
    entries: dict[tuple[str, str], str] = field(default_factory=dict)
    rules: list[tuple[re.Pattern[str], str]] = field(default_factory=list)
    missing: set[tuple[str, str]] = field(default_factory=set)

    @classmethod
    def from_path(cls, path: Path | None, rules_path: Path | None = None) -> "Translator":
        tr = cls()
        if path and path.exists():
            if path.suffix.lower() == ".json":
                raw = json.loads(path.read_text(encoding="utf-8-sig"))
                for item in raw:
                    tr.add(item.get("category", "*"), item["source"], item.get("target", ""))
            else:
                with path.open("r", encoding="utf-8-sig", newline="") as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        source = (row.get("source") or "").strip()
                        if source:
                            tr.add(row.get("category", "*"), source, row.get("target", ""))
        if rules_path and rules_path.exists():
            tr.load_rules(rules_path)
        return tr

    def add(self, category: str | None, source: str, target: str | None) -> None:
        self.entries[(category or "*", normalize(source))] = (target or "").strip()

    def translate(self, category: str, source: Any, *, record_missing: bool = True) -> str:
        if source is None:
            return ""
        text = str(source).strip()
        if not text:
            return ""
        key = normalize(text)
        for cat in (category, "*"):
            entry_key = (cat, key)
            if entry_key in self.entries:
                return self.entries[entry_key]
            entry_key_lower = (cat, key.lower())
            if entry_key_lower in self.entries:
                return self.entries[entry_key_lower]
        replaced = self.apply_rules(text)
        if replaced != text:
            return replaced
        guessed = self._rule_translate(category, text)
        if guessed != text:
            return guessed
        if record_missing:
            self.missing.add((category, text))
        return text

    def _rule_translate(self, category: str, text: str) -> str:
        if category == "skill":
            m = re.fullmatch(r"Martial Arts L(\d+)", text)
            if m:
                return f"武术{m.group(1)}级"
        return text

    def load_rules(self, path: Path) -> None:
        raw = json.loads(path.read_text(encoding="utf-8-sig"))
        rules = raw.get("rules", raw if isinstance(raw, list) else [])
        for item in rules:
            pattern = item.get("pattern")
            replacement = item.get("replacement", "")
            if not pattern:
                continue
            flags = re.IGNORECASE if "i" in item.get("flags", "") else 0
            try:
                self.rules.append((re.compile(pattern, flags), replacement))
            except re.error:
                continue

    def apply_rules(self, text: str) -> str:
        result = text
        for pattern, replacement in self.rules:
            result = pattern.sub(replacement, result)
        return result

    def write_missing(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["category", "source", "target"])
            for category, source in sorted(self.missing):
                writer.writerow([category, source, ""])


def normalize(text: Any) -> str:
    return str(text).replace(NBSP, " ").strip()


def index_by_id(items: list[dict[str, Any]]) -> dict[int, dict[str, Any]]:
    return {int(item["id"]): item for item in items}


def build_filter_maps(data: dict[str, Any]) -> dict[str, dict[int, dict[str, Any]]]:
    maps: dict[str, dict[int, dict[str, Any]]] = {}
    filters = data.get("filters", {})
    for key in FILTER_CATEGORY:
        if isinstance(filters.get(key), list):
            maps[key] = index_by_id(filters[key])
    return maps


def ref_name(ref: dict[str, Any] | int, filter_key: str, maps: dict[str, dict[int, dict[str, Any]]], tr: Translator) -> str:
    ref_id = ref if isinstance(ref, int) else ref.get("id")
    item = maps.get(filter_key, {}).get(int(ref_id), {"name": str(ref_id)})
    category = FILTER_CATEGORY[filter_key]
    name = tr.translate(category, item.get("name", ref_id))

    if isinstance(ref, dict):
        extras = ref.get("extra")
        if extras:
            extra_names = []
            for extra_id in extras:
                extra_item = maps.get("extras", {}).get(int(extra_id), {"name": str(extra_id)})
                extra_names.append(tr.translate("extra", extra_item.get("name", extra_id)))
            if extra_names:
                name = f"{name}（{'，'.join(extra_names)}）"
        q = ref.get("q")
        if q and q != 1:
            name = f"{name} x{q}"
    return name


def join_refs(refs: list[Any], filter_key: str, maps: dict[str, dict[int, dict[str, Any]]], tr: Translator) -> str:
    return "，".join(ref_name(ref, filter_key, maps, tr) for ref in sorted_refs(refs))


def sorted_refs(refs: list[Any]) -> list[Any]:
    return sorted(refs or [], key=lambda r: r.get("order", 999) if isinstance(r, dict) else 999)


def move_text(move: list[int] | tuple[int, int] | None) -> str:
    if not move:
        return ""
    if all(float(v) < 0 for v in move):
        return "-"
    return "-".join(str(cm_to_inch(v)) for v in move)


def cm_to_inch(value: int | float) -> int | float:
    if float(value) < 0:
        return "-"
    converted = float(value) / 2.5
    return int(converted) if converted.is_integer() else converted


def ava_text(value: Any) -> str:
    if value is None:
        return "-"
    try:
        n = int(value)
    except (TypeError, ValueError):
        return str(value)
    if n < 0:
        return "-"
    return "无限制" if n >= 99 else str(n)


def wound_label(profile: dict[str, Any]) -> str:
    return "结构值\nSTR" if profile.get("str") else "生命值\nVITA"


def profile_attr_values(profile: dict[str, Any]) -> list[str]:
    values = []
    for key, _ in ATTR_LABELS:
        if key == "move":
            values.append(move_text(profile.get("move")))
        elif key == "ava":
            values.append(ava_text(profile.get("ava")))
        else:
            values.append(stat_text(profile.get(key, "")))
    return values


def stat_text(value: Any) -> str:
    try:
        return "-" if float(value) < 0 else str(value)
    except (TypeError, ValueError):
        return str(value)


def profile_traits(profile: dict[str, Any], maps: dict[str, dict[int, dict[str, Any]]], tr: Translator) -> str:
    parts = []
    if profile.get("type"):
        parts.append(ref_name(profile["type"], "type", maps, tr))
    parts.extend(ref_name(c, "chars", maps, tr) for c in profile.get("chars", []))
    return "，".join(p for p in parts if p)


def category_name(category_id: Any, maps: dict[str, dict[int, dict[str, Any]]], tr: Translator) -> str:
    if category_id in (None, 0):
        return ""
    return ref_name(int(category_id), "category", maps, tr)


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_text(cell, text: str, *, bold: bool = False, size: int = 8, color: str | None = None) -> None:
    cell.text = ""
    lines = str(text).split("\n")
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for i, line in enumerate(lines):
        if i:
            paragraph.add_run().add_break()
        run = paragraph.add_run(line)
        run.bold = bold
        run.font.size = Pt(size)
        if color:
            run.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def style_table(table) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)


def add_unit_table(doc: Document, unit: dict[str, Any], pg: dict[str, Any], maps: dict[str, dict[int, dict[str, Any]]], tr: Translator) -> None:
    profiles = pg.get("profiles") or []
    if not profiles:
        return
    profile = profiles[0]
    options = pg.get("options") or []
    row_count = 7 + max(1, len(options))
    table = doc.add_table(rows=row_count, cols=10)
    style_table(table)

    title = tr.translate("unit", unit.get("name", ""))
    isc = tr.translate("unit", pg.get("isc") or unit.get("isc") or unit.get("name", ""))
    english = pg.get("isc") or unit.get("isc") or unit.get("name", "")
    cat = category_name(pg.get("category") or profile.get("category"), maps, tr)
    header = table.rows[0]
    left = header.cells[0].merge(header.cells[7])
    right = header.cells[8].merge(header.cells[9])
    set_cell_text(left, f"{isc}\n{english}", bold=True, size=10, color="FFFFFF")
    set_cell_text(right, cat, bold=True, size=9, color="FFFFFF")
    set_cell_shading(left, "2F5597")
    set_cell_shading(right, "2F5597")

    labels = [label if key != "w" else wound_label(profile) for key, label in ATTR_LABELS]
    for cell, label in zip(table.rows[1].cells, labels):
        set_cell_text(cell, label, bold=True, size=7)
        set_cell_shading(cell, "D9EAF7")
    for cell, value in zip(table.rows[2].cells, profile_attr_values(profile)):
        set_cell_text(cell, value, bold=True, size=8)

    traits = profile_traits(profile, maps, tr)
    row = table.rows[3]
    set_cell_text(row.cells[0].merge(row.cells[1]), traits, size=8)
    set_cell_text(row.cells[2].merge(row.cells[9]), "", size=8)

    equipment = join_refs(profile.get("equip", []), "equip", maps, tr)
    row = table.rows[4]
    set_cell_text(row.cells[0].merge(row.cells[1]), "装备", bold=True, size=8)
    set_cell_text(row.cells[2].merge(row.cells[9]), equipment, size=8)

    skills = join_refs(profile.get("skills", []), "skills", maps, tr)
    row = table.rows[5]
    set_cell_text(row.cells[0].merge(row.cells[1]), "特殊技能", bold=True, size=8)
    set_cell_text(row.cells[2].merge(row.cells[9]), skills, size=8)

    option_labels = ["名称", "射击武器", "", "", "近战武器", "", "", "SWC", "C"]
    row = table.rows[6]
    set_cell_text(row.cells[0].merge(row.cells[1]), option_labels[0], bold=True, size=8)
    set_cell_text(row.cells[2].merge(row.cells[4]), option_labels[1], bold=True, size=8)
    set_cell_text(row.cells[5].merge(row.cells[7]), option_labels[4], bold=True, size=8)
    set_cell_text(row.cells[8], option_labels[7], bold=True, size=8)
    set_cell_text(row.cells[9], option_labels[8], bold=True, size=8)
    for cell in row.cells:
        set_cell_shading(cell, "E7E6E6")

    if not options:
        options = [{"name": profile.get("name", ""), "weapons": profile.get("weapons", []), "swc": "-", "points": "-"}]

    for row_idx, option in enumerate(options, start=7):
        row = table.rows[row_idx]
        option_name = option_display_name(option, tr)
        bs_weapons, cc_weapons = split_weapons(option.get("weapons", []), maps)
        bs_text = join_refs(bs_weapons, "weapons", maps, tr)
        cc_text = join_refs(cc_weapons, "weapons", maps, tr)
        extra_skills = join_refs(option.get("skills", []), "skills", maps, tr)
        extra_equip = join_refs(option.get("equip", []), "equip", maps, tr)
        if extra_skills:
            option_name = f"{option_name}（{extra_skills}）"
        if extra_equip:
            bs_text = " | ".join(p for p in [bs_text, extra_equip] if p)
        set_cell_text(row.cells[0].merge(row.cells[1]), option_name, size=7)
        set_cell_text(row.cells[2].merge(row.cells[4]), bs_text, size=7)
        set_cell_text(row.cells[5].merge(row.cells[7]), cc_text, size=7)
        set_cell_text(row.cells[8], str(option.get("swc", "")), size=7)
        set_cell_text(row.cells[9], str(option.get("points", "")), size=7)

    doc.add_paragraph()


def option_display_name(option: dict[str, Any], tr: Translator) -> str:
    name = tr.translate("profile", option.get("name", ""))
    orders = [tr.translate("order", o.get("type", "")) for o in option.get("orders", []) if o.get("type") == "LIEUTENANT"]
    return f"{name}（{'，'.join(orders)}）" if orders else name


def split_weapons(weapons: list[dict[str, Any]], maps: dict[str, dict[int, dict[str, Any]]]) -> tuple[list[Any], list[Any]]:
    bs, cc = [], []
    for weapon in weapons or []:
        item = maps.get("weapons", {}).get(int(weapon.get("id", -1)), {})
        if item.get("type") == "CC":
            cc.append(weapon)
        else:
            bs.append(weapon)
    return bs, cc


def setup_document(doc: Document, title: str) -> None:
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.top_margin = Cm(1.2)
    section.bottom_margin = Cm(1.2)
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)

    styles = doc.styles
    styles["Normal"].font.name = "Microsoft YaHei"
    styles["Normal"].font.size = Pt(9)
    heading = doc.add_paragraph()
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = heading.add_run(title)
    run.bold = True
    run.font.size = Pt(20)
    run.font.name = "Microsoft YaHei"
    run.font.color.rgb = RGBColor(47, 85, 151)


def generate_docx(
    json_path: Path,
    output_path: Path,
    glossary_path: Path | None,
    missing_path: Path | None,
    rules_path: Path | None,
) -> None:
    data = json.loads(json_path.read_text(encoding="utf-8"))
    tr = Translator.from_path(glossary_path, rules_path if rules_path.exists() else None)
    maps = build_filter_maps(data)
    doc = Document()
    setup_document(doc, f"Infinity 中文军表 {data.get('version', '')}".strip())

    for unit in data.get("units", []):
        for pg in unit.get("profileGroups", []):
            add_unit_table(doc, unit, pg, maps, tr)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    if missing_path:
        tr.write_missing(missing_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Translate Infinity Army JSON into a Chinese DOCX army book.")
    parser.add_argument("json", type=Path, help="Official Infinity Army JSON file.")
    parser.add_argument("output", type=Path, help="Output .docx path.")
    parser.add_argument("--glossary", type=Path, default=Path("translations.csv"), help="CSV/JSON glossary path.")
    parser.add_argument("--rules", type=Path, default=Path("wordreplacer_rules.json"), help="Converted WordReplacer rules JSON.")
    parser.add_argument("--missing", type=Path, default=None, help="Write untranslated glossary entries to this CSV.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    generate_docx(args.json, args.output, args.glossary, args.missing, args.rules)


if __name__ == "__main__":
    main()
