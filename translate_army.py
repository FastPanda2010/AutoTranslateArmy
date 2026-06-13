from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
import re
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path
from typing import Any
from urllib.request import urlopen

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Pt, RGBColor


DOC_FONT = "Microsoft YaHei Light"
LOGO_WIDTH = Cm(2.0)
ASSET_DIR = Path("Asset")
LOGO_CACHE_DIR = Path(".cache") / "logos"
UNIT_IMAGE_DIR = ASSET_DIR / "unit_images"
ORDER_ICON_WIDTH = Pt(10)
ORDER_ICON_FILES = {
    "REGULAR": "regular.svg",
    "TACTICAL": "tactical.svg",
    "LIEUTENANT": "lieutenant.svg",
    "IRREGULAR": "irregular.svg",
    "IMPETUOUS": "impetuous.svg",
}
UNIT_TABLE_COLUMN_WIDTHS = [
    Cm(0.4), Cm(0.4),  # 命令
    Cm(0.75), Cm(0.75), Cm(0.75), Cm(0.75),  # 名称
    Cm(1.05), Cm(1.05), Cm(1.05), Cm(1.05), Cm(1.05), Cm(1.05),  # 射击武器
    Cm(0.95), Cm(0.95), Cm(0.95), Cm(0.95), Cm(0.95), Cm(0.95),  # 近战武器
    Cm(0.7),  # SWC
    Cm(0.7),  # C
]
FIRETEAM_TABLE_COLUMN_WIDTHS = [Cm(2.0), Cm(2.0), Cm(8.0)]


def apply_doc_font(run) -> None:
    """同时设置西文字体和中文 East Asia 字体，避免 Word 自动回退。"""
    run.font.name = DOC_FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), DOC_FONT)


def apply_style_font(style) -> None:
    """给 Word 样式设置中英文字体。"""
    style.font.name = DOC_FONT
    r_pr = style._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    r_fonts.set(qn("w:eastAsia"), DOC_FONT)


# 官方 JSON 里的属性字段顺序。标签中的换行会在 Word 单元格里显示成两行。
ATTR_LABELS = [
    ("move", "移动力\nMOV"),
    ("cc", "近战\nCC"),
    ("bs", "射击\nBS"),
    ("ph", "体格\nPH"),
    ("wip", "意志\nWIP"),
    ("arm", "护甲\nARM"),
    ("bts", "生化盾\nBTS"),
    ("w", None),
    ("s", "轮廓值\nS"),
    ("ava", "可用数\nAVA"),
]

# 官方 JSON 的 filters 里使用复数键；词汇表里使用更容易理解的单数类别。
FILTER_CATEGORY = {
    "skills": "skill",
    "weapons": "weapon",
    "equip": "equip",
    "chars": "char",
    "type": "type",
    "category": "category",
    "extras": "extra",
    "peripheral": "profile",
}

FIRETEAM_TYPE_LABELS = {
    "DUO": "搭档",
    "HARIS": "守护者",
    "CORE": "核心",
}

CATEGORY_SORT_ORDER = [
    "Garrison Troops",
    "Line Troops",
    "Spec. Trained Troops",
    "Veteran Troops",
    "Elite Troops",
    "Headquarters Troops",
    "Headquarters Troops / Mechanized Troops",
    "Mechanized Troops",
    "Support Troops",
    "Character",
    "Mercenary Troops",
]

# 官方数据中有些名称含不换行空格，先统一成普通空格再匹配词汇表。
NBSP = "\u00a0"


@dataclass
class Translator:
    """读取词汇表，并负责把英文术语翻译成中文。"""

    entries: dict[tuple[str, str], str] = field(default_factory=dict)
    missing: dict[tuple[str, str], tuple[str, str]] = field(default_factory=dict)

    @classmethod
    def from_path(cls, path: Path | None) -> "Translator":
        """从 CSV 或 JSON 词汇表创建翻译器。"""
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
        return tr

    def add(self, category: str | None, source: str, target: str | None) -> None:
        """加入一条翻译；同时保存原大小写和小写版本，方便大小写不敏感匹配。"""
        cat = category or "*"
        key = normalize(source)
        value = (target or "").strip()
        self.entries[(cat, key)] = value
        self.entries[(cat, key.lower())] = value

    def translate(self, category: str, source: Any, *, record_missing: bool = True) -> str:
        """按类别翻译文本；先查具体类别，再查全局 '*' 类别。"""
        if source is None:
            return ""
        text = str(source).strip()
        if not text:
            return ""
        key = normalize(text)
        found = self._lookup_entry(category, key)
        if found is not None:
            return found
        for candidate in inflection_candidates(key):
            found = self._lookup_entry(category, candidate)
            if found is not None:
                return found
        guessed = self._rule_translate(category, text)
        if guessed != text:
            return guessed
        if record_missing:
            # 没命中的词保留英文，并记录到 missing.csv 方便以后补词。
            self.record_missing(category, text)
        return text

    def _lookup_entry(self, category: str, key: str) -> str | None:
        for cat in (category, "*"):
            entry_key = (cat, key)
            if entry_key in self.entries:
                return self.entries[entry_key]
            entry_key_lower = (cat, key.lower())
            if entry_key_lower in self.entries:
                return self.entries[entry_key_lower]
        return None

    def record_missing(self, category: str, source: str) -> None:
        """大小写不敏感地记录缺词，避免同一术语输出大小写两份。"""
        key = (category, normalize(source).lower())
        current = self.missing.get(key)
        if current is None or missing_source_score(source) > missing_source_score(current[1]):
            self.missing[key] = (category, source)

    def _rule_translate(self, category: str, text: str) -> str:
        """少量可由格式稳定推断的翻译规则，避免词汇表重复写 L1-L5。"""
        if category == "skill":
            m = re.fullmatch(r"Martial Arts L(\d+)", text)
            if m:
                return f"武术{m.group(1)}级"
        return text

    def write_missing(self, path: Path) -> None:
        """把未翻译词导出成可以直接复制进词汇表的 CSV 格式。"""
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["category", "source", "target"])
            for category, source in sorted(self.missing.values(), key=lambda item: (item[0], normalize(item[1]).lower(), item[1])):
                writer.writerow([category, source, ""])


def normalize(text: Any) -> str:
    """统一文本格式，避免同一个词因为特殊空格导致匹配失败。"""
    return str(text).replace(NBSP, " ").strip()


def inflection_candidates(text: str) -> list[str]:
    words = text.split()
    if not words:
        return []

    last = words[-1]
    lower = last.lower()
    candidates: list[str] = []

    def add_last(replacement: str) -> None:
        if replacement and replacement.lower() != lower:
            candidates.append(" ".join([*words[:-1], match_case(last, replacement)]))

    if len(last) > 3 and lower.endswith("ies"):
        add_last(last[:-3] + "y")
    elif len(last) > 4 and re.search(r"(ches|shes|xes|zes|ses)$", lower):
        add_last(last[:-2])
    elif len(last) > 3 and lower.endswith("s") and not lower.endswith(("ss", "us")):
        add_last(last[:-1])
    elif len(last) > 1:
        if lower.endswith("y") and len(last) > 2 and lower[-2] not in "aeiou":
            add_last(last[:-1] + "ies")
        elif re.search(r"(ch|sh|x|z|s)$", lower):
            add_last(last + "es")
        else:
            add_last(last + "s")

    return list(dict.fromkeys(candidates))


def match_case(source: str, replacement: str) -> str:
    if source.isupper():
        return replacement.upper()
    if source[:1].isupper() and source[1:].islower():
        return replacement.capitalize()
    return replacement


def missing_source_score(source: str) -> tuple[int, int]:
    """缺词大小写重复时，优先保留更像自然显示名的写法。"""
    text = normalize(source)
    has_lower = any(ch.islower() for ch in text)
    has_upper = any(ch.isupper() for ch in text)
    return (int(has_lower and has_upper), -int(text.isupper()))


def normalize_category_for_sort(text: Any) -> str:
    """类别排序时忽略连续空白差异，例如官方的双空格类别名。"""
    return re.sub(r"\s+", " ", normalize(text))


CATEGORY_SORT_INDEX = {normalize_category_for_sort(name): index for index, name in enumerate(CATEGORY_SORT_ORDER)}
SPECIAL_CHAR_BADGES = {
    30: ("D68623", "000000"),  # Deepspace
    29: ("256D1B", "FFFFFF"),  # Surface
}


def index_by_id(items: list[dict[str, Any]]) -> dict[int, dict[str, Any]]:
    return {int(item["id"]): item for item in items}


def build_filter_maps(data: dict[str, Any]) -> dict[str, dict[int, dict[str, Any]]]:
    """把 filters 中的列表转成 id -> 项目字典，便于后续按 id 查名称。"""
    maps: dict[str, dict[int, dict[str, Any]]] = {}
    filters = data.get("filters", {})
    for key in FILTER_CATEGORY:
        if isinstance(filters.get(key), list):
            maps[key] = index_by_id(filters[key])
    return maps


def ref_name(ref: dict[str, Any] | int, filter_key: str, maps: dict[str, dict[int, dict[str, Any]]], tr: Translator, *, extra_brackets: str = "paren") -> str:
    """把 JSON 里的 id 引用转换成翻译后的名称，并附加 extra 修正。"""
    ref_id = ref if isinstance(ref, int) else ref.get("id")
    item = maps.get(filter_key, {}).get(int(ref_id), {"name": str(ref_id)})
    category = FILTER_CATEGORY[filter_key]
    name = tr.translate(category, item.get("name", ref_id))

    if isinstance(ref, dict):
        extras = ref.get("extra")
        if extras:
            # extra 通常是括号里的修正，例如 Mimetism (-3)、Dodge (+3)。
            extra_names = []
            for extra_id in extras:
                extra_item = maps.get("extras", {}).get(int(extra_id), {"name": str(extra_id)})
                extra_names.append(extra_name(extra_item, tr))
            if extra_names:
                joined_extras = "，".join(extra_names)
                if extra_brackets == "square":
                    name = f"{name}[{joined_extras}]"
                else:
                    name = f"{name}（{joined_extras}）"
        q = ref.get("q")
        if q and q != 1:
            name = f"{name} x{q}"
    return name


def extra_name(extra_item: dict[str, Any], tr: Translator) -> str:
    """格式化 extra 修正；DISTANCE 类型从厘米换算成英寸。"""
    raw_name = extra_item.get("name", "")
    if extra_item.get("type") == "DISTANCE":
        return f"{distance_extra_text(raw_name)}\""
    return tr.translate("extra", raw_name)


def distance_extra_text(value: Any) -> str:
    """官方 DISTANCE extra 使用厘米数值；军表显示时除以 2.5。"""
    text = str(value).strip()
    match = re.fullmatch(r"([+-]?)(\d+(?:\.\d+)?)", text)
    if not match:
        return text

    sign, number = match.groups()
    converted = float(number) / 2.5
    converted_text = str(int(converted)) if converted.is_integer() else f"{converted:g}"
    return f"{sign}{converted_text}"


def join_refs(refs: list[Any], filter_key: str, maps: dict[str, dict[int, dict[str, Any]]], tr: Translator, *, extra_brackets: str = "paren") -> str:
    return "，".join(ref_name(ref, filter_key, maps, tr, extra_brackets=extra_brackets) for ref in sorted_refs(refs))


def sorted_refs(refs: list[Any]) -> list[Any]:
    """官方 JSON 用 order 控制显示顺序；没有 order 的项目放到最后。"""
    return sorted(refs or [], key=lambda r: r.get("order", 999) if isinstance(r, dict) else 999)


def move_text(move: list[int] | tuple[int, int] | None) -> str:
    """官方移动力以厘米保存；Infinity N5 军表显示为英寸。"""
    if not move:
        return ""
    if all(float(v) < 0 for v in move):
        return "-"
    return "-".join(str(cm_to_inch(v)) for v in move)


def cm_to_inch(value: int | float) -> int | float:
    """把厘米换算成英寸；官方 JSON 中 -1 表示无此属性。"""
    if float(value) < 0:
        return "-"
    converted = float(value) / 2.5
    return int(converted) if converted.is_integer() else converted


def ava_text(value: Any) -> str:
    """格式化 AVA；99 这类大数按“无限制”显示。"""
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
    """机械单位显示 STR，普通单位显示 VITA。"""
    return "结构值\nSTR" if profile.get("str") else "生命值\nVITA"


def profile_attr_values(profile: dict[str, Any]) -> list[str]:
    """按 ATTR_LABELS 的顺序取出属性值，供 Word 表格第二行使用。"""
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
    """把 -1 等无效属性显示为 '-'。"""
    try:
        return "-" if float(value) < 0 else str(value)
    except (TypeError, ValueError):
        return str(value)


def profile_traits(profile: dict[str, Any], maps: dict[str, dict[int, dict[str, Any]]], tr: Translator) -> str:
    """组合兵种类型和特性，例如“轻步兵，正规军，可入侵”。"""
    parts = []
    if profile.get("type"):
        parts.append(ref_name(profile["type"], "type", maps, tr))
    parts.extend(ref_name(c, "chars", maps, tr) for c in profile.get("chars", []))
    return "，".join(p for p in parts if p)


def category_name(category_id: Any, maps: dict[str, dict[int, dict[str, Any]]], tr: Translator) -> str:
    """翻译部队类别；0 或空值表示没有类别标签。"""
    if category_id in (None, 0):
        return ""
    return ref_name(int(category_id), "category", maps, tr)


def unit_category_sort_key(unit: dict[str, Any], maps: dict[str, dict[int, dict[str, Any]]], original_index: int) -> tuple[int, int]:
    """按第一个可识别的 profileGroup/profile category 给单位排序。"""
    category_id = first_unit_category_id(unit)
    return category_sort_key(category_id, maps, original_index)


def category_sort_key(category_id: Any, maps: dict[str, dict[int, dict[str, Any]]], original_index: int) -> tuple[int, int]:
    """把 category id 转成稳定排序 key。"""
    category_item = maps.get("category", {}).get(int(category_id), {}) if category_id not in (None, 0) else {}
    category_sort = CATEGORY_SORT_INDEX.get(normalize_category_for_sort(category_item.get("name", "")), len(CATEGORY_SORT_ORDER))
    return category_sort, original_index


def first_unit_category_id(unit: dict[str, Any]) -> Any:
    """提取单位排序用类别，优先使用第一个 profileGroup 的类别。"""
    for pg in unit.get("profileGroups", []) or []:
        if pg.get("category") not in (None, 0):
            return pg.get("category")
        for profile in pg.get("profiles", []) or []:
            if profile.get("category") not in (None, 0):
                return profile.get("category")
    return unit.get("category")


def infer_faction_id(json_path: Path, data: dict[str, Any]) -> int | None:
    """优先从文件名推断军表编号；703.json 这样的文件名正好对应 faction id。"""
    if json_path.stem.isdigit():
        return int(json_path.stem)

    counts: dict[int, int] = {}
    for unit in data.get("units", []):
        for faction in unit.get("factions") or []:
            counts[int(faction)] = counts.get(int(faction), 0) + 1
    return max(counts, key=counts.get) if counts else None


def should_include_unit(unit: dict[str, Any], faction_id: int | None) -> bool:
    """只输出属于当前军表的单位；factions=[] 的佣兵池单位不输出。"""
    factions = unit.get("factions") or []
    if faction_id is None:
        return bool(factions)
    return faction_id in factions


def set_cell_shading(cell, fill: str) -> None:
    """python-docx 没有直接设置单元格底色的高级 API，这里写入底层 OOXML。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_run_shading(run, fill: str) -> None:
    r_pr = run._element.get_or_add_rPr()
    shd = r_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        r_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_text(cell, text: str, *, bold: bool = False, italic: bool = False, size: float = 8, color: str | None = None, alignment: int = WD_ALIGN_PARAGRAPH.CENTER) -> None:
    """统一写入单元格文字，保证居中、字号和换行表现一致。"""
    cell.text = ""
    lines = str(text).split("\n")
    paragraph = cell.paragraphs[0]
    paragraph.alignment = alignment
    for i, line in enumerate(lines):
        if i:
            paragraph.add_run().add_break()
        run = paragraph.add_run(line)
        run.bold = bold
        run.italic = italic
        apply_doc_font(run)
        run.font.size = Pt(size)
        if color:
            run.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


@lru_cache(maxsize=16)
def order_icon_png(order_type: str) -> bytes | None:
    """读取 Asset 中的命令 SVG，并转换成 Word 可插入的 PNG。"""
    icon_file = ORDER_ICON_FILES.get(order_type.upper())
    if not icon_file:
        return None
    path = ASSET_DIR / icon_file
    if not path.exists():
        return None
    try:
        import resvg_py

        return resvg_py.svg_to_bytes(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def set_order_icons_cell(cell, orders: list[dict[str, Any]], *, italic: bool = False) -> None:
    """在命令列插入 order 图标；无法插图时退回为文本缩写。"""
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)

    for order in orders or []:
        order_type = str(order.get("type", "")).upper()
        repeat = max(1, int(order.get("total") or 1))
        for _ in range(repeat):
            png = order_icon_png(order_type)
            run = paragraph.add_run()
            if png:
                run.add_picture(io.BytesIO(png), width=ORDER_ICON_WIDTH)
            else:
                run.text = order_type[:1]
                run.italic = italic
                apply_doc_font(run)
                run.font.size = Pt(8)
            spacer = paragraph.add_run(" ")
            apply_doc_font(spacer)
            spacer.font.size = Pt(2)

    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


@lru_cache(maxsize=256)
def fetch_logo_png(logo_url: str) -> bytes | None:
    """下载官方 SVG logo，并转换成 python-docx 可插入的 PNG。

    官方 JSON 里的 logo 通常是 SVG 链接。python-docx 对 SVG 支持不好，
    所以这里优先用 cairosvg 转成 PNG。网络失败或缺少 cairosvg 时返回 None。
    """
    if not logo_url:
        return None
    cache_path = LOGO_CACHE_DIR / f"{hashlib.sha256(logo_url.encode('utf-8')).hexdigest()}.png"
    if cache_path.exists():
        try:
            return cache_path.read_bytes()
        except Exception:
            pass

    try:
        svg_bytes = urlopen(logo_url, timeout=8).read()
    except Exception:
        return None

    try:
        import resvg_py

        png_bytes = resvg_py.svg_to_bytes(svg_bytes.decode("utf-8"))
    except Exception:
        return None

    try:
        LOGO_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        cache_path.write_bytes(png_bytes)
    except Exception:
        pass
    return png_bytes


def set_logo_cell(cell, logo_url: str | None) -> None:
    """写入单位 logo 单元格。"""
    cell.text = ""

    logo_png = fetch_logo_png(logo_url or "")
    if logo_png:
        image_paragraph = cell.paragraphs[0]
        image_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        image_paragraph.paragraph_format.space_after = Pt(0)
        image_run = image_paragraph.add_run()
        image_run.add_picture(io.BytesIO(logo_png), width=LOGO_WIDTH)

    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def sanitize_unit_image_name(name: str) -> str:
    """Return a stable Windows-safe filename stem for extracted unit images."""
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", normalize(name))
    stem = re.sub(r"\s+", " ", stem).strip(" .")
    return stem[:120] or "unit"


def _unit_image_entry(image_dir: Path, entry: dict[str, Any]) -> dict[str, Any] | None:
    path = image_dir / entry.get("file", "")
    if not path.exists():
        return None
    return {
        "path": path,
        "width_emu": int(entry.get("width_emu") or 0),
        "height_emu": int(entry.get("height_emu") or 0),
    }


def load_unit_images(image_set: str) -> dict[str, list[list[dict[str, Any]]]]:
    image_dir = UNIT_IMAGE_DIR / image_set
    manifest_path = image_dir / "manifest.json"
    if not manifest_path.exists():
        return {}

    raw = json.loads(manifest_path.read_text(encoding="utf-8"))
    result: dict[str, list[list[dict[str, Any]]]] = {}

    if raw.get("tables"):
        for table in raw.get("tables", []):
            english_name = table.get("english_name", "")
            valid_entries = [
                parsed
                for entry in table.get("images", [])
                if (parsed := _unit_image_entry(image_dir, entry)) is not None
            ]
            if english_name and valid_entries:
                result.setdefault(english_name, []).append(valid_entries)
        return result

    for english_name, entries in raw.get("units", {}).items():
        valid_entries = [
            parsed
            for entry in entries
            if (parsed := _unit_image_entry(image_dir, entry)) is not None
        ]
        if valid_entries:
            result[english_name] = [valid_entries]
    return result


def next_unit_images(unit_images: dict[str, list[list[dict[str, Any]]]] | None, english_name: str) -> list[dict[str, Any]]:
    groups = (unit_images or {}).get(english_name)
    if not groups:
        return []
    return groups.pop(0)


def set_unit_header_text(cell, chinese_name: str, english_name: str, unit_images: list[dict[str, Any]] | None = None) -> None:
    """写入单位表左侧表头：中文名左对齐大字，英文名右对齐小字。"""
    cell.text = ""

    chinese_paragraph = cell.paragraphs[0]
    chinese_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    chinese_paragraph.paragraph_format.space_after = Pt(0)
    chinese_run = chinese_paragraph.add_run(chinese_name)
    # chinese_run.bold = True
    chinese_run.font.size = Pt(14)
    chinese_run.font.color.rgb = RGBColor(0, 0, 0)

    for image in unit_images or []:
        image_paragraph = cell.add_paragraph()
        image_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        image_paragraph.paragraph_format.space_after = Pt(0)
        image_run = image_paragraph.add_run()
        width = image.get("width_emu") or None
        if width:
            image_run.add_picture(str(image["path"]), width=Emu(width))
        else:
            image_run.add_picture(str(image["path"]), width=Cm(4.0))

    english_paragraph = cell.add_paragraph()
    english_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    english_paragraph.paragraph_format.space_after = Pt(0)
    english_run = english_paragraph.add_run(english_name)
    # english_run.bold = True
    english_run.font.size = Pt(9)
    english_run.font.color.rgb = RGBColor(0, 0, 0)

    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def merge_row(row, start: int, end: int):
    """合并一行中的连续单元格，让表头可以横跨多列。"""
    return row.cells[start].merge(row.cells[end])


def style_table(table) -> None:
    """给每张单位表应用基础表格样式。"""
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)


def set_table_column_widths(table, widths: list[Any]) -> None:
    """设置表格列宽，Word 仍可能微调，但会优先参考这些宽度。"""
    table.autofit = False
    for row in table.rows:
        for cell, width in zip(row.cells, widths):
            cell.width = width


def add_note_paragraph(doc: Document, note: str | None, tr: Translator, *, label: str = "备注") -> None:
    """输出 note 段落；无内容时不输出。"""
    note_text = normalize(note or "")
    if not note_text:
        return
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(3)
    paragraph.paragraph_format.space_after = Pt(6)
    label_run = paragraph.add_run(f"{label}：")
    label_run.bold = True
    apply_doc_font(label_run)
    label_run.font.size = Pt(9)
    text_run = paragraph.add_run(tr.translate("note", note_text))
    apply_doc_font(text_run)
    text_run.font.size = Pt(9)


def add_unit_intro(doc: Document, unit: dict[str, Any], tr: Translator) -> None:
    """多子单位条目前的总标题和总 note，例如 POST-HUMANS。"""
    title = tr.translate("unit", unit.get("isc") or unit.get("name", ""))
    english = unit.get("isc") or unit.get("name", "")

    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(3)
    title_run = paragraph.add_run(title)
    title_run.bold = True
    apply_doc_font(title_run)
    title_run.font.size = Pt(16)
    if english and english != title:
        english_run = paragraph.add_run(f"\n{english}")
        apply_doc_font(english_run)
        english_run.font.size = Pt(9)
        english_run.italic = True

    add_note_paragraph(doc, unit.get("notes"), tr)


def should_add_unit_options_table(unit: dict[str, Any]) -> bool:
    """有些组队单位把可选项放在 unit.options，子 profileGroup 只保留 disabled 明细。"""
    unit_options = unit.get("options") or []
    profile_groups = unit.get("profileGroups") or []
    if not unit_options or len(profile_groups) <= 1:
        return False

    for pg in profile_groups:
        options = pg.get("options") or []
        if not options or not all(bool(option.get("disabled")) for option in options):
            return False
    return True


def add_unit_options_table(
    doc: Document,
    unit: dict[str, Any],
    maps: dict[str, dict[int, dict[str, Any]]],
    tr: Translator,
    unit_images: dict[str, list[list[dict[str, Any]]]] | None = None,
) -> None:
    """渲染只有配置行的父单位表，用于 Jazz & Billie 这类组队单位。"""
    options = unit.get("options") or []
    if not options:
        return

    row_count = 2 + len(options)
    table = doc.add_table(rows=row_count, cols=20)
    style_table(table)
    set_table_column_widths(table, UNIT_TABLE_COLUMN_WIDTHS)

    title = tr.translate("unit", unit.get("isc") or unit.get("name", ""))
    english = unit.get("isc") or unit.get("name", "")
    title_cell = merge_row(table.rows[0], 0, 19)
    set_unit_header_text(title_cell, title, english, next_unit_images(unit_images, english))
    set_cell_shading(title_cell, "FFFFFF")

    add_option_header_row(table.rows[1])
    for row_idx, option in enumerate(options, start=2):
        add_option_row(table.rows[row_idx], option, maps, tr)
        if row_idx % 2 == 1:
            for cell in table.rows[row_idx].cells:
                set_cell_shading(cell, "EFEEEE")

    doc.add_paragraph()


def fireteam_type_text(types: list[str]) -> str:
    """把官方火力组类型缩写转换成中文显示名。"""
    return "，".join(FIRETEAM_TYPE_LABELS.get(t, t) for t in types)


def fireteam_limit_text(spec: dict[str, Any]) -> str:
    """生成火力组数量限制说明，例如“最多1个核心火力组”。"""
    parts = []
    for key in ("CORE", "HARIS", "DUO"):
        value = spec.get(key)
        if value is None:
            continue
        label = FIRETEAM_TYPE_LABELS.get(key, key)
        try:
            count = int(value)
        except (TypeError, ValueError):
            parts.append(f"最多{value}个{label}火力组，")
            continue
        if count >= 99:
            parts.append(f"{label}火力组的数量没有限制。")
        else:
            parts.append(f"最多{count}个{label}火力组，")
    return "\n".join(parts)


def translate_comment_text(comment: str, tr: Translator) -> str:
    """翻译火力组备注；括号内多个逗号分隔的词会分别翻译。"""
    text = normalize(comment)
    if not text:
        return ""
    if text.startswith("(") and text.endswith(")"):
        inner = text[1:-1]
        pieces = [tr.translate("fireteam", piece.strip()) for piece in inner.split(",")]
        return f"（{'，'.join(piece for piece in pieces if piece)}）"
    if "(" in text and text.endswith(")"):
        prefix, inner = text.split("(", 1)
        prefix_text = tr.translate("fireteam", prefix.strip()) if prefix.strip() else ""
        inner_text = translate_comment_text(f"({inner}", tr)
        return f"{prefix_text}{inner_text}"
    return tr.translate("fireteam", text)


def fireteam_unit_text(unit: dict[str, Any], tr: Translator) -> str:
    """火力组成员名称，由单位名和备注组合而成。"""
    name = tr.translate("profile", unit.get("name", "").strip())
    comment = translate_comment_text(unit.get("comment", ""), tr)
    return f"{name}{comment}"


def fireteam_min_text(unit: dict[str, Any]) -> str:
    """required=true 且 min=0 时，样例军书用 '*' 表示必选骨干。"""
    minimum = unit.get("min", 0)
    if unit.get("required") and int(minimum or 0) == 0:
        return "*"
    return str(minimum)


def add_fireteam_chart(doc: Document, chart: dict[str, Any] | None, tr: Translator) -> None:
    """在单位表之前生成 fireteamChart 火力组表；reinforcements 不参与生成。"""
    if not chart or not chart.get("teams"):
        return

    heading = doc.add_paragraph()
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = heading.add_run("火力组")
    run.bold = True
    run.font.size = Pt(14)
    run.font.name = DOC_FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), DOC_FONT)
    run.font.color.rgb = RGBColor(47, 85, 151)

    rows = 1
    for team in chart.get("teams", []):
        rows += 2 + len(team.get("units", []))
    table = doc.add_table(rows=rows, cols=6)
    style_table(table)
    set_table_column_widths(table, FIRETEAM_TABLE_COLUMN_WIDTHS)

    row_idx = 0
    limit_cell = merge_row(table.rows[row_idx], 0, 5)
    set_cell_text(limit_cell, fireteam_limit_text(chart.get("spec", {})), bold=True, size=9)
    set_cell_shading(limit_cell, "D9EAF7")
    row_idx += 1

    for team in chart.get("teams", []):
        title = tr.translate("fireteam", team.get("name", ""))
        types = fireteam_type_text(team.get("type", []))
        if types and "（" not in title:
            title = f"{title}（{types}）"

        title_cell = merge_row(table.rows[row_idx], 0, 5)
        set_cell_text(title_cell, title, bold=True, size=9, color="FFFFFF")
        set_cell_shading(title_cell, "2F5597")
        row_idx += 1

        header = table.rows[row_idx]
        set_cell_text(header.cells[0], "最小", size=8)
        set_cell_shading(header.cells[0], "E7E6E6")
        set_cell_text(header.cells[1], "最大", size=8)
        set_cell_shading(header.cells[1], "E7E6E6")
        unitname_cell = merge_row(header, 2, 5)
        set_cell_text(unitname_cell, "", size=8)
        set_cell_shading(unitname_cell, "E7E6E6")
        row_idx += 1

        for idx ,unit in enumerate(team.get("units", [])):
            row = table.rows[row_idx]
            set_cell_text(row.cells[0], fireteam_min_text(unit), size=8)
            set_cell_text(row.cells[1], str(unit.get("max", "")), size=8)

            unit_cell = merge_row(table.rows[row_idx], 2, 5)
            set_cell_text(unit_cell, fireteam_unit_text(unit, tr), size=8, alignment=WD_ALIGN_PARAGRAPH.LEFT)

            row_idx += 1

    doc.add_paragraph()


def char_ref_id(ref: dict[str, Any] | int) -> int | None:
    try:
        return int(ref if isinstance(ref, int) else ref.get("id"))
    except (TypeError, ValueError):
        return None


def special_char_badges(profile: dict[str, Any], maps: dict[str, dict[int, dict[str, Any]]], tr: Translator) -> list[tuple[str, str, str]]:
    badges = []
    seen = set()
    for char_ref in profile.get("chars", []) or []:
        char_id = char_ref_id(char_ref)
        if char_id in SPECIAL_CHAR_BADGES and char_id not in seen:
            fill, color = SPECIAL_CHAR_BADGES[char_id]
            badges.append((ref_name(char_ref, "chars", maps, tr), fill, color))
            seen.add(char_id)
    return badges


def add_special_char_badges(doc: Document, profile: dict[str, Any], maps: dict[str, dict[int, dict[str, Any]]], tr: Translator) -> None:
    badges = special_char_badges(profile, maps, tr)
    if not badges:
        return

    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(3)
    paragraph.paragraph_format.space_after = Pt(1)
    for index, (text, fill, color) in enumerate(badges):
        if index:
            spacer = paragraph.add_run(" ")
            apply_doc_font(spacer)
            spacer.font.size = Pt(14)
        run = paragraph.add_run(text)
        apply_doc_font(run)
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor.from_string(color)
        set_run_shading(run, fill)


def add_unit_table(
    doc: Document,
    unit: dict[str, Any],
    pg: dict[str, Any],
    maps: dict[str, dict[int, dict[str, Any]]],
    tr: Translator,
    unit_images: dict[str, list[list[dict[str, Any]]]] | None = None,
) -> None:
    """把一个 profileGroup 渲染成 Word 中的一张单位表。

    Infinity 官方 JSON 的层级大致是：

    - unit：一个单位条目，例如 MARUTS。
    - profileGroups：同一个 unit 下的具体资料组；有些单位会把附属遥控单位放在另一个 group。
    - profiles：单位的基础属性、技能、装备、特性。
    - options：玩家建表时能选择的武器/技能/点数组合。

    这个函数只负责生成一张 20 列 Word 表：

    - 第 0 行：单位中文名/英文名 + 部队类别。
    - 第 1 行：属性栏标题，MOV/CC/BS/PH/WIP/ARM/BTS/VITA或STR/S/AVA。
    - 第 2 行：属性数值。
    - 第 3 行：兵种类型和特性，例如“轻步兵，正规军，可入侵”。
    - 第 4 行：装备。
    - 第 5 行：特殊技能。
    - 第 6 行：配置列表表头。
    - 第 7 行以后：每个 option 一行，显示名称、射击武器、近战武器、SWC、点数。
    """
    profiles = pg.get("profiles") or []
    if not profiles:
        # 没有 profile 就没有基础属性，无法生成单位表。
        return

    # 当前版本先采用每个 profileGroup 的第一个 profile 作为基础属性来源。
    # 对大多数 Infinity Army JSON 来说，差异主要体现在 options，而不是 profiles。
    profile = profiles[0]
    options = pg.get("options") or []
    add_special_char_badges(doc, profile, maps, tr)

    # 固定前 8 行为单位信息；后面每个 option 是一条配置/武器行。
    row_count = 8 + max(1, len(options))
    table = doc.add_table(rows=row_count, cols=20)
    style_table(table)
    set_table_column_widths(table, UNIT_TABLE_COLUMN_WIDTHS)

    # pg.isc 通常是资料组的英文显示名；unit.name 往往是全大写内部名。
    # 这里优先使用 pg.isc 翻译成中文，同时保留英文名作为第二行，方便对照官方。
    title = tr.translate("unit", unit.get("name", ""))
    isc = tr.translate("unit", pg.get("isc") or unit.get("isc") or unit.get("name", ""))
    english = pg.get("isc") or unit.get("isc") or unit.get("name", "")

    # category 可能出现在 profileGroup 或 profile 上；取到后通过 filters.category 翻译。
    cat = category_name(pg.get("category") or profile.get("category"), maps, tr)
    header_top = table.rows[0]
    header_bottom = table.rows[1]

    # 表头左 16 列放单位名，并跨 category/logo 两行；右 4 列上方是类别，下方是 logo。
    left = merge_row(header_top, 0, 15).merge(merge_row(header_bottom, 0, 15))
    category_cell = merge_row(header_top, 16, 19)
    logo_cell = merge_row(header_bottom, 16, 19)
    set_unit_header_text(left, isc, english, next_unit_images(unit_images, english))
    set_cell_text(category_cell, cat, bold=True, size=9)
    set_logo_cell(logo_cell, profile.get("logo") or unit.get("logo"))
    set_cell_shading(left, "FFFFFF")
    set_cell_shading(category_cell, "FFFFFF")
    set_cell_shading(logo_cell, "FFFFFF")

    # profile["str"] 为 true 时，官方资料用 STR；否则用 VITA。
    # 其他属性标题固定来自 ATTR_LABELS，保证所有单位表列顺序一致。
    labels = [label if key != "w" else wound_label(profile) for key, label in ATTR_LABELS]

    # 属性标题行和属性数值行。
    for index, label in enumerate(labels):
        cell = merge_row(table.rows[2], index * 2, index * 2 + 1)
        set_cell_text(cell, label, bold=True, size=7)
        set_cell_shading(cell, "D9EAF7")
    for index, value in enumerate(profile_attr_values(profile)):
        cell = merge_row(table.rows[3], index * 2, index * 2 + 1)
        set_cell_text(cell, value, bold=True, size=8)

    # profile_traits 会把 type id 和 chars id 翻译后拼起来。
    # 例如 type=1, chars=[3,5,21] 可能显示成“轻步兵，正规军，可入侵”。
    traits = profile_traits(profile, maps, tr)

    # 特性行、装备行、技能行都使用左侧标签 + 右侧内容的结构。
    row = table.rows[4]
    set_cell_text(merge_row(row, 0, 19), traits, size=8)
    #set_cell_text(merge_row(row, 2, 9), "", size=8)

    # 装备和技能在 JSON 中都是 id 引用；join_refs 会按 order 排序、查 filters 名称、
    # 套用 translations.csv，并把 extra 修正写成中文括号。
    equipment = join_refs(profile.get("equip", []), "equip", maps, tr)
    row = table.rows[5]
    set_cell_text(merge_row(row, 0, 3), "装备", bold=True, size=8, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(merge_row(row, 4, 19), equipment, size=8, alignment=WD_ALIGN_PARAGRAPH.LEFT)

    skills = join_refs(profile.get("skills", []), "skills", maps, tr)
    row = table.rows[6]
    set_cell_text(merge_row(row, 0, 3), "特殊技能", bold=True, size=8, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(merge_row(row, 4, 19), skills, size=8, alignment=WD_ALIGN_PARAGRAPH.LEFT)

    add_option_header_row(table.rows[7])

    if not options:
        # 极少数资料没有 options，此时用 profile 自身数据生成一行占位配置。
        options = [{"name": profile.get("name", ""), "weapons": profile.get("weapons", []), "swc": "-", "points": "-"}]

    for row_idx, option in enumerate(options, start=8):
        add_option_row(table.rows[row_idx], option, maps, tr)
        
        if row_idx % 2 == 1:
            for cell in table.rows[row_idx].cells:
                set_cell_shading(cell, "EFEEEE")

    add_note_paragraph(doc, pg.get("notes") or profile.get("notes"), tr)


def add_option_header_row(row) -> None:
    """写入配置列表表头。"""
    # option 行采用 2+4+6+6+1+1 的列宽分组：
    # 命令占 2 列，名称占 4 列，射击武器占 6 列，近战武器占 6 列，SWC 和 C 各占 1 列。
    option_labels = ["命令", "名称", "射击武器", "近战武器", "S", "C"]
    set_cell_text(merge_row(row, 0, 1), option_labels[0], bold=True, size=8)
    set_cell_text(merge_row(row, 2, 5), option_labels[1], bold=True, size=8)
    set_cell_text(merge_row(row, 6, 11), option_labels[2], bold=True, size=8)
    set_cell_text(merge_row(row, 12, 17), option_labels[3], bold=True, size=8)
    set_cell_text(row.cells[18], option_labels[4], bold=True, size=8)
    set_cell_text(row.cells[19], option_labels[5], bold=True, size=8)
    for cell in row.cells:
        set_cell_shading(cell, "E7E6E6")


def add_option_row(row, option: dict[str, Any], maps: dict[str, dict[int, dict[str, Any]]], tr: Translator) -> None:
    """写入一条 option 配置行。"""
    row_italic = bool(option.get("disabled"))

    # option.name 是这一行配置的名字；尉官等技能由 option.skills 自己输出。
    option_name = option_display_name(option, tr)

    # 官方把全部武器放在 option.weapons 中；这里按 filters.weapons[type]
    # 拆成射击武器 BS 和近战武器 CC，分别放到不同列。
    bs_weapons, cc_weapons = split_weapons(option.get("weapons", []), maps)
    bs_text = join_refs(bs_weapons, "weapons", maps, tr)
    cc_text = join_refs(cc_weapons, "weapons", maps, tr)

    # 少数配置会在 option 层级额外增加技能或装备，而不是写在 profile 层级。
    # PERSON 装备更像配置说明，放进名称括号；WEAPON 装备拼到射击武器栏。
    person_equip, weapon_equip = split_equip_by_type(option.get("equip", []), maps)
    extra_skills = join_refs(option.get("skills", []), "skills", maps, tr, extra_brackets="square")
    person_equip_text = join_refs(person_equip, "equip", maps, tr, extra_brackets="square")
    name_details = ". ".join(p for p in [extra_skills, person_equip_text] if p)
    if name_details:
        option_name = f"{option_name}（{name_details}）"

    weapon_equip_text = join_refs(weapon_equip, "equip", maps, tr)
    bs_text = " | ".join(p for p in [bs_text, weapon_equip_text] if p)
    peripheral_text = join_refs(option.get("peripheral", []), "peripheral", maps, tr)
    if peripheral_text:
        bs_text = " || ".join(p for p in [bs_text, peripheral_text] if p)

    set_order_icons_cell(merge_row(row, 0, 1), option.get("orders", []), italic=row_italic)
    set_cell_text(merge_row(row, 2, 5), option_name, italic=row_italic, size=8, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(merge_row(row, 6, 11), bs_text, italic=row_italic, size=8, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(merge_row(row, 12, 17), cc_text, italic=row_italic, size=8, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(row.cells[18], str(option.get("swc", "")), italic=row_italic, size=7.5, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(row.cells[19], str(option.get("points", "")), italic=row_italic, size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)


def split_equip_by_type(equip_refs: list[dict[str, Any]], maps: dict[str, dict[int, dict[str, Any]]]) -> tuple[list[Any], list[Any]]:
    """按 filters.equip.type 把 option 装备拆成名称栏装备和远程武器栏装备。"""
    person_equip, weapon_equip = [], []
    for equip_ref in equip_refs or []:
        item = maps.get("equip", {}).get(int(equip_ref.get("id", -1)), {})
        if item.get("type") == "WEAPON":
            weapon_equip.append(equip_ref)
        else:
            person_equip.append(equip_ref)
    return person_equip, weapon_equip


def option_display_name(option: dict[str, Any], tr: Translator) -> str:
    """配置名称。"""
    return tr.translate("profile", option.get("name", ""))


def split_weapons(weapons: list[dict[str, Any]], maps: dict[str, dict[int, dict[str, Any]]]) -> tuple[list[Any], list[Any]]:
    """按官方 weapon.type 把武器分为射击武器和近战武器。"""
    bs, cc = [], []
    for weapon in weapons or []:
        item = maps.get("weapons", {}).get(int(weapon.get("id", -1)), {})
        if item.get("type") == "CC":
            cc.append(weapon)
        else:
            bs.append(weapon)
    return bs, cc


def setup_document(doc: Document, title: str) -> None:
    """设置整份 Word 文档的页面方向、边距、默认字体和标题。"""
    section = doc.sections[0]
    # section.orientation = WD_ORIENT.LANDSCAPE
    # section.page_width, section.page_height = section.page_height, section.page_width
    section.top_margin = Cm(1.2) 
    section.bottom_margin = Cm(1.2)
    # section.left_margin = Cm(1.0)
    # section.right_margin = Cm(1.0)

    styles = doc.styles
    apply_style_font(styles["Normal"])
    styles["Normal"].font.size = Pt(10)
    heading = doc.add_paragraph()
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = heading.add_run(title)
    run.bold = True
    run.font.size = Pt(20)
    run.font.name = DOC_FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), DOC_FONT)
    # run.font.color.rgb = RGBColor(47, 85, 151)


def generate_docx(
    json_path: Path,
    output_path: Path,
    glossary_path: Path | None,
    missing_path: Path | None,
) -> None:
    """完整生成流程：读 JSON -> 读词汇表 -> 建索引 -> 写 Word -> 导出缺词表。"""
    data = json.loads(json_path.read_text(encoding="utf-8"))
    faction_id = infer_faction_id(json_path, data)
    tr = Translator.from_path(glossary_path)
    maps = build_filter_maps(data)
    unit_images = load_unit_images(json_path.stem)
    doc = Document()
    setup_document(doc, f"Infinity 中文军表 {data.get('version', '')}".strip())
    add_fireteam_chart(doc, data.get("fireteamChart"), tr)
    if data.get("fireteamChart", {}).get("teams"):
        doc.add_page_break()

    # reinforcements 是增援规则数据，不属于常规军书单位表，故意不读取。
    # factions=[] 的佣兵池单位也不输出，例如 Freelance Operator Samsa、Uhahu。
    included_units = [
        (index, unit)
        for index, unit in enumerate(data.get("units", []))
        if should_include_unit(unit, faction_id)
    ]
    included_units.sort(key=lambda item: unit_category_sort_key(item[1], maps, item[0]))

    for _, unit in included_units:
        # 一个 unit 下可能有多个 profileGroup，例如主单位和附属遥控单位。
        profile_groups = unit.get("profileGroups", [])
        if len(profile_groups) > 1 and normalize(unit.get("notes") or ""):
            add_unit_intro(doc, unit, tr)
        if should_add_unit_options_table(unit):
            add_unit_options_table(doc, unit, maps, tr, unit_images)
        for pg in unit.get("profileGroups", []):
            add_unit_table(doc, unit, pg, maps, tr, unit_images)
        doc.add_page_break()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    if missing_path:
        tr.write_missing(missing_path)


def parse_args() -> argparse.Namespace:
    """命令行参数定义。"""
    parser = argparse.ArgumentParser(description="Translate Infinity Army JSON into a Chinese DOCX army book.")
    parser.add_argument("json", type=Path, help="Official Infinity Army JSON file.")
    parser.add_argument("output", type=Path, help="Output .docx path.")
    parser.add_argument("--glossary", type=Path, default=Path("translations.csv"), help="CSV/JSON glossary path.")
    parser.add_argument("--missing", type=Path, default=None, help="Write untranslated glossary entries to this CSV.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    generate_docx(args.json, args.output, args.glossary, args.missing)


if __name__ == "__main__":
    main()
