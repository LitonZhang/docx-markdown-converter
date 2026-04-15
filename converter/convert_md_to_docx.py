#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
ET.register_namespace("w", W_NS)

TITLE_H1_RE = re.compile(r"^\s*#\s+(.+?)\s*$")
NUMBERED_TITLE_RE = re.compile(r"^\s*\d+(?:\.\d+)*(?:[.)]|\s)")
ZH_ABSTRACT_LABEL_RE = re.compile(r"^(?:摘要|中文摘要)\s*[:：]?\s*$", re.IGNORECASE)
EN_ABSTRACT_LABEL_RE = re.compile(r"^abstract\s*[:：]?\s*$", re.IGNORECASE)
FIGURE_CAPTION_RE = re.compile(r"^(?:图|figure)\s*\d+(?:[.:-]|\s+).+", re.IGNORECASE)
TABLE_CAPTION_RE = re.compile(r"^(?:表|table)\s*\d+(?:[.:-]|\s+).+", re.IGNORECASE)
HEADING_LIKE_RE = re.compile(
    r"^(?:\d+(?:\.\d+)*\.?\s+\S+|结论\b|参考文献\b|conclusion\b|references\b)",
    re.IGNORECASE,
)
NUMBERED_HEADING_RE = re.compile(r"^\s*(\d+(?:\.\d+)*)(?:[.)]?\s+)\S+")
KEYWORD_LINE_RE = re.compile(r"^(?:关键词|關鍵詞|keywords?)\s*[:：]?\s*", re.IGNORECASE)

ALIGNMENT_MAP = {
    "left": "left",
    "center": "center",
    "right": "right",
    "justify": "both",
}

MATH_GREEK_MAP = {
    "λ": r"\lambda",
    "Λ": r"\Lambda",
    "α": r"\alpha",
    "β": r"\beta",
    "γ": r"\gamma",
    "δ": r"\delta",
    "ε": r"\epsilon",
    "θ": r"\theta",
    "μ": r"\mu",
    "π": r"\pi",
    "σ": r"\sigma",
    "τ": r"\tau",
    "φ": r"\phi",
    "ω": r"\omega",
}

MATH_ACCENT_MAP = {
    "\u0302": "hat",
    "\u0303": "tilde",
    "\u0304": "bar",
    "\u0307": "dot",
    "\u0308": "ddot",
}

IDENT_BRACE_SUBSCRIPT_RE = re.compile(
    r"(?<![A-Za-z\\])(?P<base>[A-Za-z])\{(?P<sub>[A-Za-z][A-Za-z0-9]*)\}"
    r"(?=(?:\^|_|\\left|\\right|[=+\-*/,)\]}]|$|\s))"
)

PAPER_STYLE_MAP = {
    "title": "PaperTitle",
    "abstract_zh": "PaperAbstractZh",
    "abstract_en": "PaperAbstractEn",
    "heading1": "PaperHeading1",
    "heading2": "PaperHeading2",
    "heading3": "PaperHeading3",
    "figure_caption": "PaperFigureCaption",
    "table_caption": "PaperTableCaption",
    "body": "PaperBody",
}

PAPER_STYLE_NAMES = {
    "PaperTitle": "题目",
    "PaperAbstractZh": "中文摘要",
    "PaperAbstractEn": "英文摘要",
    "PaperHeading1": "一级标题",
    "PaperHeading2": "二级标题",
    "PaperHeading3": "三级标题",
    "PaperFigureCaption": "图标题",
    "PaperTableCaption": "表标题",
    "PaperBody": "正文",
}

PAPER_STYLE_ORDER = [
    "PaperTitle",
    "PaperHeading1",
    "PaperHeading2",
    "PaperHeading3",
    "PaperAbstractZh",
    "PaperAbstractEn",
    "PaperFigureCaption",
    "PaperTableCaption",
    "PaperBody",
]


@dataclass
class StyleSpec:
    zh_font: str
    en_font: str
    font_size_pt: float
    line_spacing_mode: str
    line_spacing_value: float
    align: str
    before: "SpacingSetting"
    after: "SpacingSetting"
    first_line_indent_chars: float
    bold: bool
    italic: bool


@dataclass
class AdvancedSettings:
    before: "SpacingSetting"
    after: "SpacingSetting"
    first_line_indent_chars: float
    bold: bool
    italic: bool


@dataclass
class SpacingSetting:
    mode: str
    value: float


@dataclass
class StyleConfig:
    title: StyleSpec
    abstract_zh: StyleSpec
    abstract_en: StyleSpec
    heading1: StyleSpec
    heading2: StyleSpec
    heading3: StyleSpec
    figure_caption: StyleSpec
    table_caption: StyleSpec
    body: StyleSpec
    table_settings: "TableSettings"
    advanced_defaults: AdvancedSettings


@dataclass
class TableSettings:
    table_preset: str
    header_bold: bool
    text_style: StyleSpec
    apply_text_style: bool


def w_tag(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def get_or_create(parent: ET.Element, tag: str) -> ET.Element:
    node = parent.find(tag)
    if node is None:
        node = ET.SubElement(parent, tag)
    return node


def set_bool_node(parent: ET.Element, tag: str, enabled: bool) -> None:
    node = parent.find(tag)
    if enabled:
        if node is None:
            node = ET.SubElement(parent, tag)
        node.set(w_tag("val"), "1")
    elif node is not None:
        parent.remove(node)


def set_fonts(rpr: ET.Element, zh_font: str, en_font: str) -> None:
    fonts = get_or_create(rpr, w_tag("rFonts"))
    fonts.set(w_tag("eastAsia"), zh_font)
    fonts.set(w_tag("ascii"), en_font)
    fonts.set(w_tag("hAnsi"), en_font)
    fonts.set(w_tag("cs"), en_font)


def set_font_size(rpr: ET.Element, font_size_pt: float) -> None:
    half_points = max(2, int(round(font_size_pt * 2)))
    for tag in ("sz", "szCs"):
        node = get_or_create(rpr, w_tag(tag))
        node.set(w_tag("val"), str(half_points))


def spacing_to_word(line_spacing_mode: str, line_spacing_value: float) -> tuple[str, str]:
    if line_spacing_mode == "fixed":
        twips = max(20, int(round(line_spacing_value * 20)))
        return str(twips), "exact"
    multiple = max(0.5, line_spacing_value)
    twips = max(120, int(round(multiple * 240)))
    return str(twips), "auto"


def pt_to_twips(value: float) -> str:
    return str(max(0, int(round(value * 20))))


def chars_to_hundredths(value: float) -> str:
    return str(max(0, int(round(value * 100))))


def set_paragraph_spacing(
    ppr: ET.Element,
    line_spacing_mode: str,
    line_spacing_value: float,
    before: SpacingSetting,
    after: SpacingSetting,
) -> None:
    line, rule = spacing_to_word(line_spacing_mode, line_spacing_value)
    spacing = get_or_create(ppr, w_tag("spacing"))
    spacing.set(w_tag("line"), line)
    spacing.set(w_tag("lineRule"), rule)

    spacing.attrib.pop(w_tag("before"), None)
    spacing.attrib.pop(w_tag("after"), None)
    spacing.attrib.pop(w_tag("beforeLines"), None)
    spacing.attrib.pop(w_tag("afterLines"), None)

    if before.mode == "lines":
        spacing.set(w_tag("beforeLines"), str(max(0, int(round(before.value * 100)))))
    else:
        spacing.set(w_tag("before"), pt_to_twips(before.value))

    if after.mode == "lines":
        spacing.set(w_tag("afterLines"), str(max(0, int(round(after.value * 100)))))
    else:
        spacing.set(w_tag("after"), pt_to_twips(after.value))

    spacing.set(w_tag("beforeAutospacing"), "0")
    spacing.set(w_tag("afterAutospacing"), "0")


def set_alignment(ppr: ET.Element, align: str) -> None:
    jc = get_or_create(ppr, w_tag("jc"))
    jc.set(w_tag("val"), ALIGNMENT_MAP.get(align, "left"))


def set_first_line_indent(ppr: ET.Element, first_line_indent_chars: float) -> None:
    ind = get_or_create(ppr, w_tag("ind"))
    ind.attrib.pop(w_tag("firstLine"), None)
    ind.attrib.pop(w_tag("hanging"), None)
    ind.attrib.pop(w_tag("hangingChars"), None)
    ind.set(w_tag("firstLineChars"), chars_to_hundredths(first_line_indent_chars))


def parse_non_negative_float(raw_value: object, label: str) -> float:
    value = float(raw_value)
    if value < 0:
        raise ValueError(f"{label} must be non-negative")
    return value


def parse_advanced_settings(
    raw: object,
    label: str,
    fallback: AdvancedSettings,
) -> AdvancedSettings:
    if raw is None:
        return AdvancedSettings(**fallback.__dict__)
    if not isinstance(raw, dict):
        raise ValueError(f"{label} must be an object")

    def pick(name: str, fallback_value: float | bool | dict[str, object]) -> object:
        value = raw.get(name, fallback_value)
        return fallback_value if value is None else value

    def parse_spacing(
        spacing_raw: object,
        spacing_label: str,
        spacing_fallback: SpacingSetting,
        legacy_raw: object | None,
    ) -> SpacingSetting:
        if isinstance(spacing_raw, dict):
            mode = str(spacing_raw.get("mode", spacing_fallback.mode)).strip().lower()
            if mode not in {"pt", "lines"}:
                mode = spacing_fallback.mode
            value = parse_non_negative_float(spacing_raw.get("value", spacing_fallback.value), f"{spacing_label}.value")
            return SpacingSetting(mode=mode, value=value)
        if legacy_raw is not None:
            return SpacingSetting(mode="pt", value=parse_non_negative_float(legacy_raw, spacing_label))
        return SpacingSetting(mode=spacing_fallback.mode, value=spacing_fallback.value)

    before = parse_spacing(raw.get("before"), f"{label}.before", fallback.before, raw.get("beforePt"))
    after = parse_spacing(raw.get("after"), f"{label}.after", fallback.after, raw.get("afterPt"))

    return AdvancedSettings(
        before=before,
        after=after,
        first_line_indent_chars=parse_non_negative_float(
            pick("firstLineIndentChars", fallback.first_line_indent_chars),
            f"{label}.firstLineIndentChars",
        ),
        bold=bool(pick("bold", fallback.bold)),
        italic=bool(pick("italic", fallback.italic)),
    )


def parse_style_spec(raw: dict[str, object], label: str, global_advanced: AdvancedSettings) -> StyleSpec:
    required_fields = [
        "zhFont",
        "enFont",
        "fontSizePt",
        "lineSpacingMode",
        "lineSpacingValue",
        "align",
    ]
    for field in required_fields:
        if field not in raw:
            raise ValueError(f"Missing field '{field}' in style '{label}'")

    line_spacing_mode = str(raw["lineSpacingMode"]).strip().lower()
    if line_spacing_mode not in {"multiple", "fixed"}:
        raise ValueError(f"Invalid lineSpacingMode in style '{label}': {line_spacing_mode}")

    align = str(raw["align"]).strip().lower()
    if align not in {"left", "center", "right", "justify"}:
        raise ValueError(f"Invalid align in style '{label}': {align}")

    font_size_pt = float(raw["fontSizePt"])
    line_spacing_value = float(raw["lineSpacingValue"])
    if font_size_pt <= 0 or line_spacing_value <= 0:
        raise ValueError(f"fontSizePt and lineSpacingValue must be positive in style '{label}'")

    override = parse_advanced_settings(raw.get("advancedOverride"), f"{label}.advancedOverride", global_advanced)

    # Legacy v1/v2 compatibility: per-style bold/italic lived at top level.
    if raw.get("advancedOverride") is None and ("bold" in raw or "italic" in raw):
        override.bold = bool(raw.get("bold", global_advanced.bold))
        override.italic = bool(raw.get("italic", global_advanced.italic))

    return StyleSpec(
        zh_font=str(raw["zhFont"]).strip(),
        en_font=str(raw["enFont"]).strip(),
        font_size_pt=font_size_pt,
        line_spacing_mode=line_spacing_mode,
        line_spacing_value=line_spacing_value,
        align=align,
        before=override.before,
        after=override.after,
        first_line_indent_chars=override.first_line_indent_chars,
        bold=override.bold,
        italic=override.italic,
    )


def build_default_table_text_style() -> StyleSpec:
    return StyleSpec(
        zh_font="宋体",
        en_font="Times New Roman",
        font_size_pt=12.0,
        line_spacing_mode="multiple",
        line_spacing_value=1.0,
        align="center",
        before=SpacingSetting(mode="pt", value=0.0),
        after=SpacingSetting(mode="pt", value=0.0),
        first_line_indent_chars=0.0,
        bold=False,
        italic=False,
    )


def clone_style_spec(spec: StyleSpec) -> StyleSpec:
    return StyleSpec(
        zh_font=spec.zh_font,
        en_font=spec.en_font,
        font_size_pt=spec.font_size_pt,
        line_spacing_mode=spec.line_spacing_mode,
        line_spacing_value=spec.line_spacing_value,
        align=spec.align,
        before=SpacingSetting(mode=spec.before.mode, value=spec.before.value),
        after=SpacingSetting(mode=spec.after.mode, value=spec.after.value),
        first_line_indent_chars=spec.first_line_indent_chars,
        bold=spec.bold,
        italic=spec.italic,
    )


def parse_table_settings(
    raw: object,
    global_advanced: AdvancedSettings,
    body_style: StyleSpec,
) -> TableSettings:
    if not isinstance(raw, dict):
        return TableSettings(
            table_preset="tableGrid",
            header_bold=False,
            text_style=clone_style_spec(body_style),
            apply_text_style=False,
        )

    table_preset = str(raw.get("tablePreset", "threeLine")).strip()
    if table_preset not in {"threeLine", "tableGrid", "table"}:
        table_preset = "threeLine"

    header_bold = bool(raw.get("headerBold", False))
    apply_text_style = bool(raw.get("applyTextStyle", True))

    text_style_raw = raw.get("textStyle")
    if isinstance(text_style_raw, dict):
        text_style = parse_style_spec(text_style_raw, "tableSettings.textStyle", global_advanced)
    else:
        text_style = build_default_table_text_style()

    return TableSettings(
        table_preset=table_preset,
        header_bold=header_bold,
        text_style=text_style,
        apply_text_style=apply_text_style,
    )


def parse_style_config(raw: dict[str, object]) -> StyleConfig:
    style_map = {
        "title": "title",
        "abstractZh": "abstract_zh",
        "abstractEn": "abstract_en",
        "heading1": "heading1",
        "heading2": "heading2",
        "heading3": "heading3",
        "figureCaption": "figure_caption",
        "tableCaption": "table_caption",
        "body": "body",
    }

    default_advanced = AdvancedSettings(
        before=SpacingSetting(mode="pt", value=0.0),
        after=SpacingSetting(mode="pt", value=0.0),
        first_line_indent_chars=0.0,
        bold=False,
        italic=False,
    )
    global_advanced = parse_advanced_settings(raw.get("advancedDefaults"), "advancedDefaults", default_advanced)

    parsed: dict[str, StyleSpec] = {}
    for json_key, attr_name in style_map.items():
        section_raw = raw.get(json_key)
        if json_key == "heading3" and not isinstance(section_raw, dict):
            section_raw = raw.get("heading2")
        if not isinstance(section_raw, dict):
            raise ValueError(f"Missing style section: {json_key}")
        parsed[attr_name] = parse_style_spec(section_raw, json_key, global_advanced)

    table_settings = parse_table_settings(
        raw.get("tableSettings"),
        global_advanced,
        parsed["body"],
    )

    return StyleConfig(
        **parsed,
        table_settings=table_settings,
        advanced_defaults=global_advanced,
    )


def load_style_config(style_json_path: Path) -> StyleConfig:
    try:
        raw = json.loads(style_json_path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"Invalid style JSON: {exc}") from exc

    if not isinstance(raw, dict):
        raise ValueError("Style JSON root must be an object")

    if "styleConfig" in raw and isinstance(raw["styleConfig"], dict):
        raw = raw["styleConfig"]

    return parse_style_config(raw)


def emit_stdout(message: str) -> None:
    sys.stdout.buffer.write((message + "\n").encode("utf-8", errors="backslashreplace"))
    sys.stdout.flush()


def emit_stderr(message: str) -> None:
    sys.stderr.buffer.write((message + "\n").encode("utf-8", errors="backslashreplace"))
    sys.stderr.flush()


def find_style(root: ET.Element, style_id: str) -> ET.Element | None:
    for style in root.findall("w:style", NS):
        if style.get(w_tag("styleId")) == style_id:
            return style
    return None


def remove_child(parent: ET.Element, tag: str) -> None:
    node = parent.find(tag)
    if node is not None:
        parent.remove(node)


def ensure_style(root: ET.Element, style_id: str, style_name: str) -> ET.Element:
    style = find_style(root, style_id)
    if style is not None:
        style.set(w_tag("type"), "paragraph")
        style.set(w_tag("customStyle"), "1")
        name_node = style.find(w_tag("name"))
        if name_node is None:
            name_node = ET.SubElement(style, w_tag("name"))
        name_node.set(w_tag("val"), style_name)
        return style

    style = ET.SubElement(
        root,
        w_tag("style"),
        {
            w_tag("type"): "paragraph",
            w_tag("styleId"): style_id,
            w_tag("customStyle"): "1",
        },
    )
    ET.SubElement(style, w_tag("name"), {w_tag("val"): style_name})
    return style


def set_based_on(style: ET.Element, based_on: str) -> None:
    based = style.find(w_tag("basedOn"))
    if based is None:
        based = ET.SubElement(style, w_tag("basedOn"))
    based.set(w_tag("val"), based_on)


def set_quick_style(style: ET.Element, enabled: bool) -> None:
    qformat = style.find(w_tag("qFormat"))
    if enabled:
        if qformat is None:
            ET.SubElement(style, w_tag("qFormat"))
    elif qformat is not None:
        style.remove(qformat)


def set_ui_priority(style: ET.Element, priority: int) -> None:
    node = style.find(w_tag("uiPriority"))
    if node is None:
        node = ET.SubElement(style, w_tag("uiPriority"))
    node.set(w_tag("val"), str(priority))


def hide_style_from_gallery(style: ET.Element) -> None:
    set_quick_style(style, False)
    set_bool_node(style, w_tag("semiHidden"), True)
    set_bool_node(style, w_tag("unhideWhenUsed"), True)
    set_ui_priority(style, 99)


def expose_style_in_gallery(style: ET.Element) -> None:
    set_quick_style(style, True)
    remove_child(style, w_tag("semiHidden"))
    remove_child(style, w_tag("unhideWhenUsed"))


def trim_latent_styles(styles_root: ET.Element) -> None:
    latent = styles_root.find("w:latentStyles", NS)
    if latent is None:
        return
    for node in list(latent.findall("w:lsdException", NS)):
        latent.remove(node)
    latent.set(w_tag("count"), "0")
    latent.set(w_tag("defQFormat"), "0")


def apply_style_spec(style: ET.Element, spec: StyleSpec) -> None:
    ppr = get_or_create(style, w_tag("pPr"))
    rpr = get_or_create(style, w_tag("rPr"))

    set_fonts(rpr, spec.zh_font, spec.en_font)
    set_font_size(rpr, spec.font_size_pt)
    set_bool_node(rpr, w_tag("b"), spec.bold)
    set_bool_node(rpr, w_tag("i"), spec.italic)

    set_paragraph_spacing(
        ppr,
        spec.line_spacing_mode,
        spec.line_spacing_value,
        spec.before,
        spec.after,
    )
    set_alignment(ppr, spec.align)
    set_first_line_indent(ppr, spec.first_line_indent_chars)


def apply_style_config_to_styles_xml(styles_root: ET.Element, config: StyleConfig) -> None:
    style_targets = [
        ("PaperTitle", PAPER_STYLE_NAMES["PaperTitle"], config.title, "Normal"),
        ("PaperHeading1", PAPER_STYLE_NAMES["PaperHeading1"], config.heading1, "Normal"),
        ("PaperHeading2", PAPER_STYLE_NAMES["PaperHeading2"], config.heading2, "Normal"),
        ("PaperHeading3", PAPER_STYLE_NAMES["PaperHeading3"], config.heading3, "Normal"),
        ("PaperAbstractZh", PAPER_STYLE_NAMES["PaperAbstractZh"], config.abstract_zh, "PaperBody"),
        ("PaperAbstractEn", PAPER_STYLE_NAMES["PaperAbstractEn"], config.abstract_en, "PaperBody"),
        (
            "PaperFigureCaption",
            PAPER_STYLE_NAMES["PaperFigureCaption"],
            config.figure_caption,
            "PaperBody",
        ),
        ("PaperTableCaption", PAPER_STYLE_NAMES["PaperTableCaption"], config.table_caption, "PaperBody"),
        ("PaperBody", PAPER_STYLE_NAMES["PaperBody"], config.body, "Normal"),
    ]

    for priority, (style_id, style_name, spec, based_on) in enumerate(style_targets, start=1):
        style = ensure_style(styles_root, style_id, style_name)
        set_based_on(style, based_on)
        expose_style_in_gallery(style)
        set_ui_priority(style, priority)
        apply_style_spec(style, spec)

    paper_style_ids = set(PAPER_STYLE_ORDER)
    for style in styles_root.findall("w:style", NS):
        style_id = style.get(w_tag("styleId"), "")
        if style_id not in paper_style_ids:
            hide_style_from_gallery(style)

    trim_latent_styles(styles_root)


def ensure_pandoc() -> str:
    pandoc = shutil.which("pandoc")
    if not pandoc:
        raise SystemExit("Pandoc not found in PATH. Please install Pandoc and retry.")
    return pandoc


def has_front_matter(lines: list[str]) -> tuple[bool, int]:
    if not lines or lines[0].strip() != "---":
        return False, -1
    for idx in range(1, len(lines)):
        if lines[idx].strip() == "---":
            return True, idx
    return False, -1


def extract_title_from_markdown(markdown_text: str) -> tuple[str, str | None]:
    lines = markdown_text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    has_fm, fm_end = has_front_matter(lines)

    title_idx = None
    title_value = None

    search_start = fm_end + 1 if has_fm else 0
    for idx in range(search_start, len(lines)):
        match = TITLE_H1_RE.match(lines[idx])
        if not match:
            continue

        candidate = match.group(1).strip()
        if "**" in candidate or "__" in candidate:
            continue
        if NUMBERED_TITLE_RE.match(candidate):
            continue

        title_idx = idx
        title_value = candidate
        break

    if title_idx is None or title_value is None:
        return markdown_text, None

    del lines[title_idx]

    title_line = f"title: {json.dumps(title_value, ensure_ascii=False)}"
    if has_fm:
        fm_lines = lines[: fm_end + 1]
        if any(re.match(r"^\s*title\s*:\s*", line, re.IGNORECASE) for line in fm_lines):
            return "\n".join(lines).strip() + "\n", title_value
        fm_lines.insert(fm_end, title_line)
        merged = fm_lines + lines[fm_end + 1 :]
        return "\n".join(merged).strip() + "\n", title_value

    body = "\n".join(lines).lstrip("\n")
    transformed = f"---\n{title_line}\n---\n\n{body}".rstrip() + "\n"
    return transformed, title_value


def normalize_tex_math_content(math_text: str) -> str:
    normalized = math_text
    normalized = normalized.replace(r"\left‖", r"\left\|")
    normalized = normalized.replace(r"\right‖", r"\right\|")
    normalized = normalized.replace(r"\left{", r"\left\{")
    normalized = normalized.replace(r"\right}", r"\right\}")
    normalized = normalized.replace(r"\left}", r"\left\}")
    normalized = normalized.replace(r"\right{", r"\right\{")
    normalized = normalized.replace("‖", r"\|")
    normalized = normalized.replace("∥", r"\|")

    parts: list[str] = []
    idx = 0
    while idx < len(normalized):
        ch = normalized[idx]
        if ch == "\\":
            if idx + 1 < len(normalized):
                parts.append(normalized[idx : idx + 2])
                idx += 2
            else:
                parts.append(ch)
                idx += 1
            continue

        if ch in MATH_GREEK_MAP:
            parts.append(MATH_GREEK_MAP[ch])
            idx += 1
            continue

        decomp = unicodedata.normalize("NFD", ch)
        if len(decomp) >= 2 and decomp[0].isalpha():
            combining_marks = [mark for mark in decomp[1:] if unicodedata.combining(mark)]
            if len(combining_marks) == 1 and combining_marks[0] in MATH_ACCENT_MAP:
                accent = MATH_ACCENT_MAP[combining_marks[0]]
                base = decomp[0]
                parts.append(rf"\{accent}{{{base}}}")
                idx += 1
                continue

        parts.append(ch)
        idx += 1

    normalized = "".join(parts)
    normalized = IDENT_BRACE_SUBSCRIPT_RE.sub(r"\g<base>_{\g<sub>}", normalized)
    return normalized


def normalize_markdown_math(markdown_text: str) -> str:
    out: list[str] = []
    i = 0
    n = len(markdown_text)

    while i < n:
        ch = markdown_text[i]

        if ch == "\\" and i + 1 < n:
            out.append(markdown_text[i : i + 2])
            i += 2
            continue

        if ch != "$":
            out.append(ch)
            i += 1
            continue

        delim_len = 2 if i + 1 < n and markdown_text[i + 1] == "$" else 1
        j = i + delim_len
        found = False

        while j < n:
            if markdown_text[j] == "\\" and j + 1 < n:
                j += 2
                continue

            if delim_len == 2:
                if j + 1 < n and markdown_text[j] == "$" and markdown_text[j + 1] == "$":
                    content = markdown_text[i + 2 : j]
                    normalized = normalize_tex_math_content(content)
                    out.append("$$" + normalized + "$$")
                    i = j + 2
                    found = True
                    break
            else:
                if markdown_text[j] == "$":
                    content = markdown_text[i + 1 : j]
                    normalized = normalize_tex_math_content(content)
                    out.append("$" + normalized + "$")
                    i = j + 1
                    found = True
                    break
            j += 1

        if not found:
            out.append(markdown_text[i : i + delim_len])
            i += delim_len

    return "".join(out)


def paragraph_text(paragraph: ET.Element) -> str:
    chunks: list[str] = []
    for node in paragraph.iter():
        if node.tag == w_tag("t") and node.text:
            chunks.append(node.text)
        elif node.tag == w_tag("tab"):
            chunks.append(" ")
        elif node.tag == w_tag("br"):
            chunks.append(" ")
    return "".join(chunks).strip()


def normalized_hash_text(text: str) -> str:
    normalized = re.sub(r"\s+", " ", text).strip()
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


def table_plain_text(table: ET.Element) -> str:
    values: list[str] = []
    for cell in table.findall(".//w:tc", NS):
        parts: list[str] = []
        for node in cell.iter():
            if node.tag == w_tag("t") and node.text:
                parts.append(node.text)
            elif node.tag == w_tag("tab"):
                parts.append(" ")
            elif node.tag == w_tag("br"):
                parts.append(" ")
        text = "".join(parts).strip()
        if text:
            values.append(text)
    return re.sub(r"\s+", " ", " | ".join(values)).strip()


def get_paragraph_style(paragraph: ET.Element) -> str:
    p_style = paragraph.find("./w:pPr/w:pStyle", NS)
    if p_style is None:
        return ""
    return p_style.get(w_tag("val"), "")


def set_paragraph_style(paragraph: ET.Element, style_id: str) -> None:
    ppr = paragraph.find(w_tag("pPr"))
    if ppr is None:
        ppr = ET.SubElement(paragraph, w_tag("pPr"))
    p_style = ppr.find(w_tag("pStyle"))
    if p_style is None:
        p_style = ET.SubElement(ppr, w_tag("pStyle"))
    p_style.set(w_tag("val"), style_id)


def clear_paragraph_direct_formatting(paragraph: ET.Element) -> None:
    ppr = paragraph.find(w_tag("pPr"))
    if ppr is None:
        return

    # Keep structure semantics (style/numbering/section breaks), but clear direct
    # paragraph overrides so visual output follows the configured Paper* styles.
    removable = [
        w_tag("jc"),
        w_tag("spacing"),
        w_tag("ind"),
        w_tag("mirrorInd"),
        w_tag("textAlignment"),
        w_tag("contextualSpacing"),
        w_tag("rPr"),
    ]
    for tag in removable:
        node = ppr.find(tag)
        if node is not None:
            ppr.remove(node)


def apply_style_and_normalize(paragraph: ET.Element, style_id: str) -> None:
    set_paragraph_style(paragraph, style_id)
    clear_paragraph_direct_formatting(paragraph)


def heading_level_from_style(style_id: str) -> int | None:
    lowered = style_id.lower().replace(" ", "")
    if lowered in {"papertitle", "title"}:
        return 1
    if lowered in {"paperheading1", "heading1"}:
        return 1
    if lowered in {"paperheading2", "heading2"}:
        return 2
    if lowered in {"paperheading3", "heading3"}:
        return 3
    if lowered.startswith("heading") and len(lowered) > 7 and lowered[7:].isdigit():
        return min(int(lowered[7:]), 3)
    return None


def infer_heading_level_from_text(text: str) -> int | None:
    normalized = text.strip()
    if not normalized:
        return None

    numbered = NUMBERED_HEADING_RE.match(normalized)
    if numbered:
        depth = numbered.group(1).count(".") + 1
        return min(max(depth, 1), 3)

    if HEADING_LIKE_RE.match(normalized):
        return 1
    return None


def looks_like_title_candidate(text: str) -> bool:
    normalized = text.strip()
    if not normalized:
        return False
    if ZH_ABSTRACT_LABEL_RE.match(normalized) or EN_ABSTRACT_LABEL_RE.match(normalized):
        return False
    if FIGURE_CAPTION_RE.match(normalized) or TABLE_CAPTION_RE.match(normalized):
        return False
    if KEYWORD_LINE_RE.match(normalized):
        return False
    if infer_heading_level_from_text(normalized) is not None:
        return False
    return True


def apply_semantic_styles(document_root: ET.Element) -> int | None:
    abstract_mode: str | None = None
    title_assigned = False
    first_heading1_index: int | None = None
    paragraphs = document_root.findall(".//w:body/w:p", NS)
    paragraph_texts = [paragraph_text(paragraph) for paragraph in paragraphs]

    def next_non_empty_text(current_index: int) -> str:
        for idx in range(current_index + 1, len(paragraph_texts)):
            candidate = paragraph_texts[idx].strip()
            if candidate:
                return candidate
        return ""

    for index, paragraph in enumerate(paragraphs):
        text = paragraph_text(paragraph)
        if not text:
            continue

        style_id = get_paragraph_style(paragraph).strip()
        style_key = style_id.lower().replace(" ", "")
        heading_level = heading_level_from_style(style_id)
        inferred_heading_level = infer_heading_level_from_text(text)
        effective_heading_level = heading_level if heading_level is not None else inferred_heading_level

        if not title_assigned and style_key in {"title", "papertitle"}:
            apply_style_and_normalize(paragraph, PAPER_STYLE_MAP["title"])
            title_assigned = True
            abstract_mode = None
            continue

        if not title_assigned and index == 0 and looks_like_title_candidate(text):
            next_text = next_non_empty_text(index)
            if next_text and (
                ZH_ABSTRACT_LABEL_RE.match(next_text)
                or EN_ABSTRACT_LABEL_RE.match(next_text)
                or infer_heading_level_from_text(next_text) is not None
            ):
                apply_style_and_normalize(paragraph, PAPER_STYLE_MAP["title"])
                title_assigned = True
                abstract_mode = None
                continue

        if ZH_ABSTRACT_LABEL_RE.match(text):
            apply_style_and_normalize(paragraph, PAPER_STYLE_MAP["abstract_zh"])
            abstract_mode = "zh"
            continue

        if EN_ABSTRACT_LABEL_RE.match(text):
            apply_style_and_normalize(paragraph, PAPER_STYLE_MAP["abstract_en"])
            abstract_mode = "en"
            continue

        if effective_heading_level == 1 and not title_assigned:
            # Only treat the first top-level heading at the very beginning as title.
            # This avoids misclassifying chapter headings (for example "1 引言") as title.
            if index == 0 and looks_like_title_candidate(text):
                apply_style_and_normalize(paragraph, PAPER_STYLE_MAP["title"])
                title_assigned = True
                abstract_mode = None
                continue

        if effective_heading_level == 1:
            apply_style_and_normalize(paragraph, PAPER_STYLE_MAP["heading1"])
            abstract_mode = None
            if first_heading1_index is None:
                first_heading1_index = index
            continue

        if effective_heading_level == 2:
            apply_style_and_normalize(paragraph, PAPER_STYLE_MAP["heading2"])
            abstract_mode = None
            continue

        if effective_heading_level is not None and effective_heading_level >= 3:
            apply_style_and_normalize(paragraph, PAPER_STYLE_MAP["heading3"])
            abstract_mode = None
            continue

        if FIGURE_CAPTION_RE.match(text):
            apply_style_and_normalize(paragraph, PAPER_STYLE_MAP["figure_caption"])
            continue

        if TABLE_CAPTION_RE.match(text):
            apply_style_and_normalize(paragraph, PAPER_STYLE_MAP["table_caption"])
            continue

        if abstract_mode == "zh":
            apply_style_and_normalize(paragraph, PAPER_STYLE_MAP["abstract_zh"])
        elif abstract_mode == "en":
            apply_style_and_normalize(paragraph, PAPER_STYLE_MAP["abstract_en"])
        else:
            apply_style_and_normalize(paragraph, PAPER_STYLE_MAP["body"])

    return first_heading1_index


def ensure_body_node(document_root: ET.Element) -> ET.Element:
    body = document_root.find("w:body", NS)
    if body is None:
        raise ValueError("DOCX document.xml has no body")
    return body


def ensure_body_section(document_root: ET.Element) -> ET.Element:
    body = ensure_body_node(document_root)
    sect_pr = body.find("w:sectPr", NS)
    if sect_pr is None:
        sect_pr = ET.SubElement(body, w_tag("sectPr"))
    return sect_pr


def set_section_columns(sect_pr: ET.Element, columns: str) -> None:
    cols = sect_pr.find(w_tag("cols"))
    if cols is None:
        cols = ET.SubElement(sect_pr, w_tag("cols"))
    if columns == "double":
        cols.set(w_tag("num"), "2")
        cols.set(w_tag("space"), "720")
    else:
        cols.set(w_tag("num"), "1")
        cols.attrib.pop(w_tag("space"), None)


def remove_paragraph_section_breaks(paragraphs: list[ET.Element]) -> None:
    for paragraph in paragraphs:
        ppr = paragraph.find(w_tag("pPr"))
        if ppr is None:
            continue
        sect_pr = ppr.find(w_tag("sectPr"))
        if sect_pr is not None:
            ppr.remove(sect_pr)


def apply_body_layout(document_root: ET.Element, columns: str, body_start_index: int | None) -> list[str]:
    warnings: list[str] = []
    paragraphs = document_root.findall(".//w:body/w:p", NS)
    remove_paragraph_section_breaks(paragraphs)

    body_sect_pr = ensure_body_section(document_root)
    if columns == "single":
        set_section_columns(body_sect_pr, "single")
        return warnings

    if body_start_index is None:
        set_section_columns(body_sect_pr, "single")
        warnings.append("Body start not detected; fallback to single-column layout.")
        return warnings

    set_section_columns(body_sect_pr, "double")

    if body_start_index > 0:
        section_break_paragraph = paragraphs[body_start_index - 1]
        ppr = section_break_paragraph.find(w_tag("pPr"))
        if ppr is None:
            ppr = ET.SubElement(section_break_paragraph, w_tag("pPr"))
        sect_pr = ppr.find(w_tag("sectPr"))
        if sect_pr is None:
            sect_pr = ET.SubElement(ppr, w_tag("sectPr"))
        set_section_columns(sect_pr, "single")
        section_type = sect_pr.find(w_tag("type"))
        if section_type is None:
            section_type = ET.SubElement(sect_pr, w_tag("type"))
        section_type.set(w_tag("val"), "continuous")

    return warnings


def set_border(
    borders: ET.Element,
    edge: str,
    val: str,
    size: int | None = None,
    color: str = "auto",
) -> None:
    node = get_or_create(borders, w_tag(edge))
    node.set(w_tag("val"), val)
    if val == "nil":
        node.attrib.pop(w_tag("sz"), None)
        node.attrib.pop(w_tag("space"), None)
        node.attrib.pop(w_tag("color"), None)
        return
    node.set(w_tag("sz"), str(size if size is not None else 4))
    node.set(w_tag("space"), "0")
    node.set(w_tag("color"), color)


def apply_three_line_table_borders(table: ET.Element) -> None:
    tbl_pr = get_or_create(table, w_tag("tblPr"))
    tbl_style = get_or_create(tbl_pr, w_tag("tblStyle"))
    tbl_style.set(w_tag("val"), "Table")

    tbl_borders = get_or_create(tbl_pr, w_tag("tblBorders"))
    set_border(tbl_borders, "top", "single", 12)
    set_border(tbl_borders, "bottom", "single", 12)
    set_border(tbl_borders, "left", "nil")
    set_border(tbl_borders, "right", "nil")
    set_border(tbl_borders, "insideH", "nil")
    set_border(tbl_borders, "insideV", "nil")

    rows = table.findall("./w:tr", NS)
    for row_index, row in enumerate(rows):
        for cell in row.findall("./w:tc", NS):
            tc_pr = get_or_create(cell, w_tag("tcPr"))
            existing = tc_pr.find(w_tag("tcBorders"))
            if existing is not None:
                tc_pr.remove(existing)
            if row_index == 0:
                tc_borders = ET.SubElement(tc_pr, w_tag("tcBorders"))
                set_border(tc_borders, "top", "nil")
                set_border(tc_borders, "left", "nil")
                set_border(tc_borders, "right", "nil")
                set_border(tc_borders, "bottom", "single", 8)


def apply_table_text_style_to_paragraph(
    paragraph: ET.Element,
    spec: StyleSpec,
    force_bold: bool | None = None,
) -> None:
    ppr = get_or_create(paragraph, w_tag("pPr"))
    set_paragraph_spacing(
        ppr,
        spec.line_spacing_mode,
        spec.line_spacing_value,
        spec.before,
        spec.after,
    )
    set_alignment(ppr, spec.align)
    set_first_line_indent(ppr, spec.first_line_indent_chars)

    for run in paragraph.findall(".//w:r", NS):
        rpr = get_or_create(run, w_tag("rPr"))
        set_fonts(rpr, spec.zh_font, spec.en_font)
        set_font_size(rpr, spec.font_size_pt)
        set_bool_node(rpr, w_tag("b"), spec.bold if force_bold is None else force_bold)
        set_bool_node(rpr, w_tag("i"), spec.italic)


def apply_table_header_bold_to_paragraph(paragraph: ET.Element) -> None:
    for run in paragraph.findall(".//w:r", NS):
        rpr = get_or_create(run, w_tag("rPr"))
        set_bool_node(rpr, w_tag("b"), True)


def apply_table_settings(document_root: ET.Element, settings: TableSettings) -> None:
    style_id_map = {
        "tableGrid": "TableGrid",
        "table": "Table",
    }

    for table in document_root.findall(".//w:body/w:tbl", NS):
        tbl_pr = get_or_create(table, w_tag("tblPr"))
        tbl_style = get_or_create(tbl_pr, w_tag("tblStyle"))

        if settings.table_preset == "threeLine":
            apply_three_line_table_borders(table)
        else:
            tbl_style.set(w_tag("val"), style_id_map.get(settings.table_preset, "TableGrid"))
            tbl_borders = tbl_pr.find(w_tag("tblBorders"))
            if tbl_borders is not None:
                tbl_pr.remove(tbl_borders)
            for cell in table.findall(".//w:tc", NS):
                tc_pr = get_or_create(cell, w_tag("tcPr"))
                tc_borders = tc_pr.find(w_tag("tcBorders"))
                if tc_borders is not None:
                    tc_pr.remove(tc_borders)

        rows = table.findall("./w:tr", NS)
        for row_index, row in enumerate(rows):
            is_header_row = row_index == 0
            for paragraph in row.findall(".//w:tc/w:p", NS):
                if settings.apply_text_style:
                    force_bold = True if (settings.header_bold and is_header_row) else None
                    apply_table_text_style_to_paragraph(
                        paragraph,
                        settings.text_style,
                        force_bold=force_bold,
                    )
                elif settings.header_bold and is_header_row:
                    apply_table_header_bold_to_paragraph(paragraph)


def apply_table_grid_style(document_root: ET.Element) -> None:
    apply_table_settings(
        document_root,
        TableSettings(
            table_preset="tableGrid",
            header_bold=False,
            text_style=build_default_table_text_style(),
            apply_text_style=False,
        ),
    )


def load_roundtrip_metadata(roundtrip_meta_path: Path | None) -> dict[str, object] | None:
    if roundtrip_meta_path is None or not roundtrip_meta_path.exists():
        return None
    try:
        payload = json.loads(roundtrip_meta_path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError:
        return None
    return payload if isinstance(payload, dict) else None


def collect_doc_blocks(document_root: ET.Element) -> list[dict[str, object]]:
    body = ensure_body_node(document_root)
    blocks: list[dict[str, object]] = []
    for child in list(body):
        local = child.tag.split("}", 1)[-1]
        if local == "p":
            plain = paragraph_text(child)
            blocks.append(
                {
                    "type": "paragraph",
                    "node": child,
                    "plainText": plain,
                    "textHash": normalized_hash_text(plain),
                }
            )
        elif local == "tbl":
            plain = table_plain_text(child)
            blocks.append(
                {
                    "type": "table",
                    "node": child,
                    "plainText": plain,
                    "textHash": normalized_hash_text(plain),
                }
            )
    return blocks


def match_roundtrip_blocks(
    meta_blocks: list[dict[str, object]],
    doc_blocks: list[dict[str, object]],
) -> list[dict[str, object]]:
    matches: list[dict[str, object]] = []
    used_meta: set[int] = set()

    # Pass 1: exact type + hash
    for doc_index, doc_block in enumerate(doc_blocks):
        doc_type = str(doc_block.get("type", ""))
        doc_hash = str(doc_block.get("textHash", ""))
        found_index: int | None = None
        for meta_index, meta_block in enumerate(meta_blocks):
            if meta_index in used_meta:
                continue
            if str(meta_block.get("type", "")) != doc_type:
                continue
            if str(meta_block.get("textHash", "")) != doc_hash:
                continue
            found_index = meta_index
            break
        if found_index is None:
            continue
        used_meta.add(found_index)
        matches.append(
            {
                "docIndex": doc_index,
                "metaIndex": found_index,
                "quality": "exact",
            }
        )

    # Pass 2: near-neighbor by type order
    for doc_index, doc_block in enumerate(doc_blocks):
        if any(item["docIndex"] == doc_index for item in matches):
            continue
        doc_type = str(doc_block.get("type", ""))
        found_index = None
        for meta_index, meta_block in enumerate(meta_blocks):
            if meta_index in used_meta:
                continue
            if str(meta_block.get("type", "")) != doc_type:
                continue
            found_index = meta_index
            break
        if found_index is None:
            continue
        used_meta.add(found_index)
        matches.append(
            {
                "docIndex": doc_index,
                "metaIndex": found_index,
                "quality": "near",
            }
        )

    return matches


def _ensure_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph.find(w_tag("pPr"))
    if ppr is None:
        ppr = ET.SubElement(paragraph, w_tag("pPr"))
    return ppr


def _apply_paragraph_snapshot(paragraph: ET.Element, snapshot: dict[str, object]) -> None:
    ppr = _ensure_ppr(paragraph)

    align = str(snapshot.get("align") or "").strip()
    if align:
        jc = ppr.find(w_tag("jc"))
        if jc is None:
            jc = ET.SubElement(ppr, w_tag("jc"))
        jc.set(w_tag("val"), align)

    spacing_snapshot = snapshot.get("spacing")
    if isinstance(spacing_snapshot, dict):
        spacing = ppr.find(w_tag("spacing"))
        if spacing is None:
            spacing = ET.SubElement(ppr, w_tag("spacing"))
        for key in ("before", "after", "beforeLines", "afterLines", "line", "lineRule"):
            value = spacing_snapshot.get(key)
            if value is None:
                continue
            spacing.set(w_tag(key), str(value))
        spacing.set(w_tag("beforeAutospacing"), "0")
        spacing.set(w_tag("afterAutospacing"), "0")

    indent_snapshot = snapshot.get("indent")
    if isinstance(indent_snapshot, dict):
        ind = ppr.find(w_tag("ind"))
        if ind is None:
            ind = ET.SubElement(ppr, w_tag("ind"))
        for key in ("left", "right", "firstLine", "firstLineChars"):
            value = indent_snapshot.get(key)
            if value is None:
                continue
            ind.set(w_tag(key), str(value))


def _apply_run_snapshot(paragraph: ET.Element, run_snapshots: list[dict[str, object]]) -> None:
    runs = paragraph.findall(".//w:r", NS)
    for index, run in enumerate(runs):
        if index >= len(run_snapshots):
            break
        snapshot = run_snapshots[index]
        if not isinstance(snapshot, dict):
            continue

        rpr = run.find(w_tag("rPr"))
        if rpr is None:
            rpr = ET.SubElement(run, w_tag("rPr"))

        def set_toggle(tag: str, enabled: object) -> None:
            node = rpr.find(w_tag(tag))
            if bool(enabled):
                if node is None:
                    node = ET.SubElement(rpr, w_tag(tag))
                node.set(w_tag("val"), "1")
            elif node is not None:
                rpr.remove(node)

        set_toggle("b", snapshot.get("bold"))
        set_toggle("i", snapshot.get("italic"))
        set_toggle("u", snapshot.get("underline"))

        zh_font = snapshot.get("zhFont")
        en_font = snapshot.get("enFont")
        if zh_font or en_font:
            fonts = rpr.find(w_tag("rFonts"))
            if fonts is None:
                fonts = ET.SubElement(rpr, w_tag("rFonts"))
            if zh_font:
                fonts.set(w_tag("eastAsia"), str(zh_font))
            if en_font:
                fonts.set(w_tag("ascii"), str(en_font))
                fonts.set(w_tag("hAnsi"), str(en_font))

        size = snapshot.get("fontSizeHalfPt")
        if size:
            sz = rpr.find(w_tag("sz"))
            if sz is None:
                sz = ET.SubElement(rpr, w_tag("sz"))
            sz.set(w_tag("val"), str(size))
            sz_cs = rpr.find(w_tag("szCs"))
            if sz_cs is None:
                sz_cs = ET.SubElement(rpr, w_tag("szCs"))
            sz_cs.set(w_tag("val"), str(size))


def apply_roundtrip_restore(
    document_root: ET.Element,
    metadata: dict[str, object],
) -> tuple[list[str], bool]:
    warnings: list[str] = []
    blocks_raw = metadata.get("blocks")
    if not isinstance(blocks_raw, list):
        return ["Roundtrip metadata missing blocks list; skipped restore."], False

    meta_blocks = [item for item in blocks_raw if isinstance(item, dict)]
    if not meta_blocks:
        return ["Roundtrip metadata blocks empty; skipped restore."], False

    doc_blocks = collect_doc_blocks(document_root)
    matches = match_roundtrip_blocks(meta_blocks, doc_blocks)
    if not matches:
        return ["No roundtrip block match found; kept pandoc output."], False

    body = ensure_body_node(document_root)
    for match in matches:
        doc_index = int(match["docIndex"])
        meta_index = int(match["metaIndex"])
        quality = str(match["quality"])
        doc_block = doc_blocks[doc_index]
        meta_block = meta_blocks[meta_index]
        block_type = str(meta_block.get("type", ""))

        if block_type == "paragraph" and isinstance(meta_block.get("paragraphStyle"), dict):
            _apply_paragraph_snapshot(
                doc_block["node"],  # type: ignore[arg-type]
                meta_block["paragraphStyle"],  # type: ignore[arg-type]
            )
            run_snapshots = meta_block.get("runs")
            if isinstance(run_snapshots, list):
                filtered_runs = [item for item in run_snapshots if isinstance(item, dict)]
                _apply_run_snapshot(doc_block["node"], filtered_runs)  # type: ignore[arg-type]

        if block_type == "table":
            table_snapshot = meta_block.get("tableSnapshot")
            if isinstance(table_snapshot, dict):
                raw_xml = table_snapshot.get("rawXml")
                if quality == "exact" and isinstance(raw_xml, str) and raw_xml.strip():
                    try:
                        replacement = ET.fromstring(raw_xml)
                        parent_children = list(body)
                        target = doc_block["node"]  # type: ignore[assignment]
                        if target in parent_children:
                            replace_index = parent_children.index(target)
                            body.remove(target)
                            body.insert(replace_index, replacement)
                    except ET.ParseError:
                        warnings.append("Failed to parse table rawXml; fallback to TableGrid style.")

    # Restore section columns if metadata contains layout.
    layout = metadata.get("documentLayout")
    if isinstance(layout, dict):
        default_columns = str(layout.get("defaultColumns", "1"))
        body_sect = ensure_body_section(document_root)
        set_section_columns(body_sect, "double" if default_columns == "2" else "single")

        sections = layout.get("sections")
        if isinstance(sections, list):
            paragraph_nodes = document_root.findall(".//w:body/w:p", NS)
            remove_paragraph_section_breaks(paragraph_nodes)
            order_to_doc_index: dict[int, int] = {}
            for match in matches:
                meta_index = int(match["metaIndex"])
                doc_index = int(match["docIndex"])
                order_value = meta_blocks[meta_index].get("order")
                if isinstance(order_value, int):
                    order_to_doc_index[order_value] = doc_index

            for section in sections:
                if not isinstance(section, dict):
                    continue
                end_order = section.get("endOrder")
                columns = section.get("columns")
                if not isinstance(end_order, int):
                    continue
                if int(columns or 1) == int(default_columns or 1):
                    continue
                doc_index = order_to_doc_index.get(end_order)
                if doc_index is None or doc_index >= len(doc_blocks):
                    continue
                target_node = doc_blocks[doc_index]["node"]
                if target_node.tag != w_tag("p"):
                    continue
                ppr = target_node.find(w_tag("pPr"))
                if ppr is None:
                    ppr = ET.SubElement(target_node, w_tag("pPr"))
                sect = ppr.find(w_tag("sectPr"))
                if sect is None:
                    sect = ET.SubElement(ppr, w_tag("sectPr"))
                set_section_columns(sect, "double" if int(columns or 1) == 2 else "single")
                sect_type = sect.find(w_tag("type"))
                if sect_type is None:
                    sect_type = ET.SubElement(sect, w_tag("type"))
                sect_type.set(w_tag("val"), "continuous")

    return warnings, True


def rewrite_docx_with_styles(
    docx_path: Path,
    config: StyleConfig,
    apply_semantics: bool,
) -> list[str]:
    temp_docx = docx_path.with_suffix(".styled.tmp.docx")
    warnings: list[str] = []

    with ZipFile(docx_path) as source:
        styles_xml = source.read("word/styles.xml")
        document_xml = source.read("word/document.xml")

        styles_root = ET.fromstring(styles_xml)
        document_root = ET.fromstring(document_xml)

        apply_style_config_to_styles_xml(styles_root, config)
        apply_table_settings(document_root, config.table_settings)
        if apply_semantics:
            apply_semantic_styles(document_root)

        updated_styles = ET.tostring(styles_root, encoding="utf-8", xml_declaration=True)
        updated_document = ET.tostring(document_root, encoding="utf-8", xml_declaration=True)

        with ZipFile(temp_docx, "w") as target:
            for info in source.infolist():
                if info.filename == "word/styles.xml":
                    target.writestr(info, updated_styles)
                elif info.filename == "word/document.xml":
                    target.writestr(info, updated_document)
                else:
                    target.writestr(info, source.read(info.filename))

    shutil.move(temp_docx, docx_path)
    return warnings


def create_reference_docx(pandoc: str, config: StyleConfig, output_path: Path) -> None:
    reference_bytes = subprocess.check_output([pandoc, "--print-default-data-file", "reference.docx"])
    output_path.write_bytes(reference_bytes)
    rewrite_docx_with_styles(output_path, config, apply_semantics=False)


def run_pandoc_export(pandoc: str, input_md: Path, output_docx: Path, reference_docx: Path) -> tuple[str, str, bool]:
    cmd_with_reference = [
        pandoc,
        str(input_md),
        "-f",
        "markdown",
        "-t",
        "docx",
        "-o",
        str(output_docx),
        "--reference-doc",
        str(reference_docx),
    ]

    first = subprocess.run(cmd_with_reference, capture_output=True, text=True, encoding="utf-8")
    if first.returncode == 0:
        return first.stdout.strip(), first.stderr.strip(), False

    cmd_plain = [
        pandoc,
        str(input_md),
        "-f",
        "markdown",
        "-t",
        "docx",
        "-o",
        str(output_docx),
    ]
    second = subprocess.run(cmd_plain, capture_output=True, text=True, encoding="utf-8")
    if second.returncode == 0:
        fallback_stderr = (first.stderr or "").strip()
        second_stderr = (second.stderr or "").strip()
        merged_stderr = "\n".join(part for part in [fallback_stderr, second_stderr] if part)
        return (second.stdout or "").strip(), merged_stderr, True

    combined_stdout = "\n".join(part for part in [(first.stdout or "").strip(), (second.stdout or "").strip()] if part)
    combined_stderr = "\n".join(part for part in [(first.stderr or "").strip(), (second.stderr or "").strip()] if part)
    raise RuntimeError(combined_stderr or combined_stdout or "Pandoc export failed")


def convert_markdown_to_docx(
    input_path: Path,
    output_path: Path,
    config: StyleConfig,
) -> dict[str, object]:
    if not input_path.exists():
        raise ValueError(f"Input markdown not found: {input_path}")
    if input_path.suffix.lower() != ".md":
        raise ValueError("Input file must be a .md file")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    pandoc = ensure_pandoc()
    original_text = input_path.read_text(encoding="utf-8-sig")
    normalized_markdown = original_text
    title_value: str | None = None
    normalized_markdown = normalize_markdown_math(normalized_markdown)

    with tempfile.TemporaryDirectory(prefix="md_docx_export_") as tmp_dir:
        temp_dir = Path(tmp_dir)
        temp_markdown = temp_dir / "prepared_input.md"
        temp_reference = temp_dir / "reference.docx"

        temp_markdown.write_text(normalized_markdown, encoding="utf-8")

        create_reference_docx(pandoc, config, temp_reference)
        stdout, stderr, fallback_used = run_pandoc_export(pandoc, temp_markdown, output_path, temp_reference)

    warnings = rewrite_docx_with_styles(
        output_path,
        config,
        apply_semantics=True,
    )

    return {
        "success": True,
        "outputPath": str(output_path),
        "titleDetected": bool(title_value),
        "fallbackUsed": fallback_used,
        "stdout": stdout,
        "stderr": stderr,
        "warnings": warnings,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Convert Markdown to DOCX with custom style config")
    parser.add_argument("--input", required=True, help="Input markdown (.md) path")
    parser.add_argument("--output", required=True, help="Output docx path")
    parser.add_argument("--style", required=True, help="Style JSON path")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        style_config = load_style_config(Path(args.style))
        report = convert_markdown_to_docx(
            Path(args.input),
            Path(args.output),
            style_config,
        )
        emit_stdout(json.dumps(report, ensure_ascii=False))
        return 0
    except ValueError as exc:
        emit_stderr(str(exc))
        return 1
    except subprocess.CalledProcessError as exc:
        message = (exc.stderr or exc.stdout or str(exc)).strip()
        emit_stderr(message)
        return 2
    except RuntimeError as exc:
        emit_stderr(str(exc))
        return 2
    except Exception as exc:  # noqa: BLE001
        emit_stderr(f"Unexpected error: {exc}")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
