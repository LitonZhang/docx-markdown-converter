
from __future__ import annotations

from dataclasses import asdict, dataclass, field
from html import escape as html_escape
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import re
import zipfile
from xml.etree import ElementTree as ET

from .omml_to_latex import convert_omml_to_latex

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

CAPTION_LINE_RE = re.compile(r"^(?:(?:图|表|figure|table)\s*\d+)", re.IGNORECASE)
HEADING_NUMBER_RE = re.compile(r"^(\d+(?:[\.．]\d+)*)(?:\s*[)）、.]|\s+)\s*\S+")
ABSTRACT_LABEL_RE = re.compile(r"^(?:摘要|中文摘要|abstract)\s*[:：]?\s*$", re.IGNORECASE)
KEYWORD_LABEL_RE = re.compile(r"^(?:关键词|關鍵詞|keywords?)\s*[:：]?\s*", re.IGNORECASE)
HEADING_STYLE_RE = re.compile(r"heading\s*([1-6])", re.IGNORECASE)
MARKDOWN_LINK_RE = re.compile(r"\[([^\]]+)\]\([^\)]+\)")
MARKDOWN_STRIP_RE = re.compile(r"[`*_>#]")


def _local_name(tag: str) -> str:
    return tag.split("}", 1)[-1]


def _qn(ns: str, name: str) -> str:
    return f"{{{ns}}}{name}"


def _normalize_ws(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _text_hash(text: str) -> str:
    normalized = _normalize_ws(text)
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


def _escape_markdown_text(text: str) -> str:
    escapes = {
        "\\": "\\\\",
        "*": "\\*",
        "_": "\\_",
        "`": "\\`",
        "[": "\\[",
        "]": "\\]",
    }
    for source, target in escapes.items():
        text = text.replace(source, target)
    return text


def _plain_text_from_markdown(markdown_text: str) -> str:
    simplified = MARKDOWN_LINK_RE.sub(r"\1", markdown_text)
    simplified = simplified.replace("<br>", " ")
    simplified = MARKDOWN_STRIP_RE.sub("", simplified)
    simplified = simplified.replace("|", " ")
    return _normalize_ws(simplified)


class RecoverableConversionError(Exception):
    pass


@dataclass
class ConversionStats:
    headings: int = 0
    tables: int = 0
    images: int = 0
    equations: int = 0


@dataclass
class ConversionResult:
    success: bool
    warnings: list[str]
    stats: ConversionStats
    duration_ms: int
    markdown: str
    assets_dir: str | None = None

    def to_report(self) -> dict[str, object]:
        payload: dict[str, object] = {
            "success": self.success,
            "warnings": self.warnings,
            "stats": asdict(self.stats),
            "duration_ms": self.duration_ms,
        }
        if self.assets_dir is not None:
            payload["assetsDir"] = self.assets_dir
        return payload


@dataclass
class Relationship:
    target: str
    rel_type: str
    target_mode: str | None = None


@dataclass
class ConversionContext:
    zip_ref: zipfile.ZipFile
    relationships: dict[str, Relationship]
    style_map: dict[str, str]
    output_path: Path
    image_dir: Path | None
    assets_dir: Path | None
    extract_images: bool
    stats: ConversionStats
    warnings: list[str] = field(default_factory=list)
    image_cache: dict[str, Path] = field(default_factory=dict)
    image_counter: int = 1
    skipped_images_noted: bool = False


class DocxToMarkdownConverter:
    def __init__(self, math: str = "latex", extract_images: bool = False) -> None:
        if math != "latex":
            raise RecoverableConversionError("Only --math latex is supported in v1")
        self.math = math
        self.extract_images = extract_images

    def convert_file(
        self,
        input_path: str | os.PathLike[str],
        output_path: str | os.PathLike[str],
        image_dir: str | os.PathLike[str] | None = None,
        assets_dir: str | os.PathLike[str] | None = None,
    ) -> ConversionResult:
        input_file = Path(input_path)
        output_file = Path(output_path)

        if not input_file.exists():
            raise RecoverableConversionError(f"Input file not found: {input_file}")
        if input_file.suffix.lower() != ".docx":
            raise RecoverableConversionError("Input file must be a .docx document")

        output_file.parent.mkdir(parents=True, exist_ok=True)

        assets_dir_path = Path(assets_dir) if assets_dir is not None else output_file.parent / f"{output_file.stem}_assets"
        image_dir_path = Path(image_dir) if image_dir is not None else assets_dir_path / "images"
        if self.extract_images:
            image_dir_path.mkdir(parents=True, exist_ok=True)

        stats = ConversionStats()
        warnings: list[str] = []

        try:
            with zipfile.ZipFile(input_file, "r") as zip_ref:
                document_xml = self._read_xml(zip_ref, "word/document.xml")
                relationships = self._read_relationships(zip_ref)
                style_map = self._read_style_map(zip_ref)
                context = ConversionContext(
                    zip_ref=zip_ref,
                    relationships=relationships,
                    style_map=style_map,
                    output_path=output_file,
                    image_dir=image_dir_path,
                    assets_dir=assets_dir_path,
                    extract_images=self.extract_images,
                    stats=stats,
                    warnings=warnings,
                )
                markdown = self._convert_document(document_xml, context)
        except zipfile.BadZipFile as exc:
            raise RecoverableConversionError("Invalid DOCX file (ZIP format broken)") from exc
        except KeyError as exc:
            raise RecoverableConversionError(f"Malformed DOCX package: missing {exc}") from exc
        except ET.ParseError as exc:
            raise RecoverableConversionError(f"Malformed DOCX XML: {exc}") from exc

        output_file.write_text(markdown, encoding="utf-8")

        return ConversionResult(
            success=True,
            warnings=list(dict.fromkeys(warnings)),
            stats=stats,
            duration_ms=0,
            markdown=markdown,
            assets_dir=str(assets_dir_path) if self.extract_images else None,
        )

    def _read_xml(self, zip_ref: zipfile.ZipFile, zip_path: str) -> ET.Element:
        xml_text = zip_ref.read(zip_path)
        return ET.fromstring(xml_text)

    def _read_relationships(self, zip_ref: zipfile.ZipFile) -> dict[str, Relationship]:
        rel_path = "word/_rels/document.xml.rels"
        if rel_path not in zip_ref.namelist():
            return {}
        root = self._read_xml(zip_ref, rel_path)
        relationships: dict[str, Relationship] = {}
        for rel in root.findall(_qn(REL_NS, "Relationship")):
            rel_id = rel.attrib.get("Id")
            target = rel.attrib.get("Target")
            rel_type = rel.attrib.get("Type", "")
            target_mode = rel.attrib.get("TargetMode")
            if rel_id and target:
                relationships[rel_id] = Relationship(target=target, rel_type=rel_type, target_mode=target_mode)
        return relationships

    def _read_style_map(self, zip_ref: zipfile.ZipFile) -> dict[str, str]:
        style_path = "word/styles.xml"
        if style_path not in zip_ref.namelist():
            return {}
        root = self._read_xml(zip_ref, style_path)
        mapping: dict[str, str] = {}
        for style in root.findall(_qn(W_NS, "style")):
            style_id = style.attrib.get(_qn(W_NS, "styleId"), "")
            name_node = style.find(_qn(W_NS, "name"))
            style_name = name_node.attrib.get(_qn(W_NS, "val"), "") if name_node is not None else ""
            if style_id:
                mapping[style_id] = style_name
        return mapping

    def _convert_document(self, document_root: ET.Element, context: ConversionContext) -> str:
        body = document_root.find(_qn(W_NS, "body"))
        if body is None:
            raise RecoverableConversionError("DOCX body is missing")

        blocks: list[str] = []
        first_paragraph_written = False
        for child in list(body):
            name = _local_name(child.tag)
            if name == "p":
                paragraph_md, _, _ = self._parse_paragraph(
                    child,
                    context,
                    is_first_paragraph=(not first_paragraph_written),
                )
                if paragraph_md:
                    blocks.append(paragraph_md)
                    first_paragraph_written = True
            elif name == "tbl":
                table_md, _ = self._parse_table(child, context)
                if table_md:
                    blocks.append(table_md)

        if not blocks:
            context.warnings.append("Empty document content")
            return ""

        return "\n\n".join(blocks).strip() + "\n"

    def _parse_paragraph(
        self,
        paragraph: ET.Element,
        context: ConversionContext,
        is_first_paragraph: bool = False,
    ) -> tuple[str, dict[str, object] | None, int | None]:
        style_id = self._get_paragraph_style_id(paragraph)
        style_name = context.style_map.get(style_id, "")
        list_level = self._list_level(paragraph)
        is_quote = "quote" in style_name.lower()
        is_code = "code" in style_name.lower() or "pre" in style_name.lower()

        segments: list[str] = []
        run_snapshots: list[dict[str, object]] = []
        equation_snapshots: list[dict[str, object]] = []
        image_snapshots: list[dict[str, object]] = []
        plain_parts: list[str] = []
        plain_cursor = 0

        for child in list(paragraph):
            name = _local_name(child.tag)
            if name == "r":
                run_parts, run_snapshot, equations, images, plain_text = self._parse_run(child, context)
                segments.extend(run_parts)
                equation_snapshots.extend(equations)
                image_snapshots.extend(images)
                if run_snapshot is not None and plain_text:
                    run_snapshot["start"] = plain_cursor
                    plain_cursor += len(plain_text)
                    run_snapshot["end"] = plain_cursor
                    run_snapshots.append(run_snapshot)
                    plain_parts.append(plain_text)
            elif name == "hyperlink":
                hyperlink_text = self._parse_hyperlink(child, context)
                segments.append(hyperlink_text)
                plain_link_text = _plain_text_from_markdown(hyperlink_text)
                if plain_link_text:
                    plain_parts.append(plain_link_text)
                    plain_cursor += len(plain_link_text)
            elif name in {"oMath", "oMathPara"}:
                equation_md = self._parse_equation(child, context, block=(name == "oMathPara"))
                if equation_md:
                    segments.append(equation_md)
                    plain_equation = _plain_text_from_markdown(equation_md)
                    if plain_equation:
                        equation_snapshots.append({"latex": plain_equation, "display": name == "oMathPara"})
                        plain_parts.append(plain_equation)
                        plain_cursor += len(plain_equation)

        content = "".join(segments).strip()
        if not content:
            return "", None, self._extract_section_columns_from_paragraph(paragraph)

        content_plain = _normalize_ws(" ".join(plain_parts)) or _plain_text_from_markdown(content)
        heading_level = self._heading_level(content, style_id, style_name)
        if (
            is_first_paragraph
            and heading_level == 0
            and list_level < 0
            and not is_quote
            and not is_code
            and self._looks_like_title_candidate(content)
        ):
            heading_level = 1

        paragraph_md = content
        block_type = "paragraph"
        if is_code:
            paragraph_md = f"```\n{content}\n```"
            block_type = "code"
        elif heading_level > 0:
            context.stats.headings += 1
            paragraph_md = f"{'#' * heading_level} {content}"
            block_type = "heading"
        elif list_level >= 0:
            indent = "  " * list_level
            paragraph_md = f"{indent}- {content}"
            block_type = "list"
        elif is_quote:
            paragraph_md = f"> {content}"
            block_type = "quote"

        paragraph_meta: dict[str, object] = {
            "type": block_type,
            "plainText": content_plain,
            "textHash": _text_hash(content_plain),
            "paragraphStyle": self._snapshot_paragraph_style(paragraph, style_id, style_name),
            "runs": run_snapshots,
            "equationSnapshot": equation_snapshots,
            "imageSnapshot": image_snapshots,
        }
        return paragraph_md, paragraph_meta, self._extract_section_columns_from_paragraph(paragraph)

    def _snapshot_paragraph_style(self, paragraph: ET.Element, style_id: str, style_name: str) -> dict[str, object]:
        ppr = paragraph.find(_qn(W_NS, "pPr"))
        snapshot: dict[str, object] = {
            "styleId": style_id,
            "styleName": style_name,
            "align": "left",
            "spacing": {},
            "indent": {},
            "list": {},
        }
        if ppr is None:
            return snapshot

        jc = ppr.find(_qn(W_NS, "jc"))
        if jc is not None:
            snapshot["align"] = jc.attrib.get(_qn(W_NS, "val"), "left")

        spacing = ppr.find(_qn(W_NS, "spacing"))
        if spacing is not None:
            snapshot["spacing"] = {
                "before": spacing.attrib.get(_qn(W_NS, "before")),
                "after": spacing.attrib.get(_qn(W_NS, "after")),
                "beforeLines": spacing.attrib.get(_qn(W_NS, "beforeLines")),
                "afterLines": spacing.attrib.get(_qn(W_NS, "afterLines")),
                "line": spacing.attrib.get(_qn(W_NS, "line")),
                "lineRule": spacing.attrib.get(_qn(W_NS, "lineRule")),
            }

        ind = ppr.find(_qn(W_NS, "ind"))
        if ind is not None:
            snapshot["indent"] = {
                "left": ind.attrib.get(_qn(W_NS, "left")),
                "right": ind.attrib.get(_qn(W_NS, "right")),
                "firstLine": ind.attrib.get(_qn(W_NS, "firstLine")),
                "firstLineChars": ind.attrib.get(_qn(W_NS, "firstLineChars")),
            }

        num_pr = ppr.find(_qn(W_NS, "numPr"))
        if num_pr is not None:
            ilvl = num_pr.find(_qn(W_NS, "ilvl"))
            num_id = num_pr.find(_qn(W_NS, "numId"))
            snapshot["list"] = {
                "level": ilvl.attrib.get(_qn(W_NS, "val")) if ilvl is not None else None,
                "numId": num_id.attrib.get(_qn(W_NS, "val")) if num_id is not None else None,
            }

        snapshot["rawPprXml"] = ET.tostring(ppr, encoding="unicode")
        return snapshot

    def _parse_run(
        self,
        run: ET.Element,
        context: ConversionContext,
    ) -> tuple[list[str], dict[str, object] | None, list[dict[str, object]], list[dict[str, object]], str]:
        bold = run.find("./w:rPr/w:b", {"w": W_NS}) is not None
        italic = run.find("./w:rPr/w:i", {"w": W_NS}) is not None
        underline = run.find("./w:rPr/w:u", {"w": W_NS}) is not None
        run_style = run.find("./w:rPr/w:rStyle", {"w": W_NS})
        inline_code = run_style is not None and "code" in run_style.attrib.get(_qn(W_NS, "val"), "").lower()

        r_fonts = run.find("./w:rPr/w:rFonts", {"w": W_NS})
        r_size = run.find("./w:rPr/w:sz", {"w": W_NS})
        run_snapshot: dict[str, object] = {
            "bold": bold,
            "italic": italic,
            "underline": underline,
            "zhFont": r_fonts.attrib.get(_qn(W_NS, "eastAsia")) if r_fonts is not None else None,
            "enFont": r_fonts.attrib.get(_qn(W_NS, "ascii")) if r_fonts is not None else None,
            "fontSizeHalfPt": r_size.attrib.get(_qn(W_NS, "val")) if r_size is not None else None,
        }

        parts: list[str] = []
        text_buffer: list[str] = []
        plain_chunks: list[str] = []
        equation_snapshots: list[dict[str, object]] = []
        image_snapshots: list[dict[str, object]] = []

        def flush_text() -> None:
            if not text_buffer:
                return
            text = "".join(text_buffer)
            text_buffer.clear()
            escaped = _escape_markdown_text(text)
            if inline_code and escaped.strip():
                escaped = f"`{escaped}`"
            parts.append(escaped)
            plain_chunks.append(_normalize_ws(text.replace("\n", " ")))

        for child in list(run):
            name = _local_name(child.tag)
            if name == "t":
                text_buffer.append(child.text or "")
            elif name == "tab":
                text_buffer.append("    ")
            elif name in {"br", "cr"}:
                text_buffer.append("  \n")
            elif name in {"drawing", "pict"}:
                flush_text()
                image_md, image_snapshot = self._extract_image(child, context)
                if image_md:
                    parts.append(image_md)
                if image_snapshot is not None:
                    image_snapshots.append(image_snapshot)
            elif name in {"oMath", "oMathPara"}:
                flush_text()
                equation_md = self._parse_equation(child, context, block=(name == "oMathPara"))
                if equation_md:
                    parts.append(equation_md)
                    equation_snapshots.append({"latex": _plain_text_from_markdown(equation_md), "display": name == "oMathPara"})

        flush_text()
        plain_text = _normalize_ws(" ".join(chunk for chunk in plain_chunks if chunk))
        if plain_text:
            run_snapshot["text"] = plain_text
            return parts, run_snapshot, equation_snapshots, image_snapshots, plain_text
        return parts, None, equation_snapshots, image_snapshots, ""

    def _parse_hyperlink(self, hyperlink: ET.Element, context: ConversionContext) -> str:
        rel_id = hyperlink.attrib.get(_qn(R_NS, "id"))
        text_parts: list[str] = []
        for child in list(hyperlink):
            if _local_name(child.tag) == "r":
                run_parts, _, _, _, _ = self._parse_run(child, context)
                text_parts.extend(run_parts)

        link_text = "".join(text_parts).strip() or "link"
        if not rel_id:
            return link_text

        relationship = context.relationships.get(rel_id)
        if not relationship:
            context.warnings.append(f"Hyperlink relationship missing for id={rel_id}")
            return link_text

        if relationship.target_mode == "External" or relationship.target.startswith("http"):
            return f"[{link_text}]({relationship.target})"
        return link_text

    def _parse_equation(self, equation_element: ET.Element, context: ConversionContext, block: bool) -> str:
        latex, warnings = convert_omml_to_latex(equation_element)
        context.stats.equations += 1
        for warning in warnings:
            context.warnings.append(f"Equation warning: {warning}")
        if not latex:
            context.warnings.append("Equation converted to empty output")
            return ""
        if block:
            return f"$$\n{latex}\n$$"
        return f"${latex}$"

    def _extract_image(self, drawing_element: ET.Element, context: ConversionContext) -> tuple[str, dict[str, object] | None]:
        if not context.extract_images:
            if not context.skipped_images_noted:
                context.warnings.append("Images detected but skipped (image export disabled)")
                context.skipped_images_noted = True
            return "", None

        relationship_id = None
        for node in drawing_element.iter():
            for key, value in node.attrib.items():
                if _local_name(key) in {"embed", "link"}:
                    relationship_id = value
                    break
            if relationship_id:
                break

        if not relationship_id:
            context.warnings.append("Image relationship id missing")
            return "", None

        relationship = context.relationships.get(relationship_id)
        if not relationship:
            context.warnings.append(f"Image relationship missing for id={relationship_id}")
            return "", None

        context.stats.images += 1
        if relationship_id in context.image_cache:
            saved_path = context.image_cache[relationship_id]
            image_ref = self._render_image_ref(saved_path, context.output_path.parent)
            return image_ref, {
                "relationshipId": relationship_id,
                "target": relationship.target,
                "relativePath": image_ref[8:-1] if image_ref.startswith("![image](") else image_ref,
            }

        zip_target = self._resolve_word_target(relationship.target)
        if zip_target not in context.zip_ref.namelist():
            context.warnings.append(f"Image target missing in package: {zip_target}")
            return "", None

        image_bytes = context.zip_ref.read(zip_target)
        filename = Path(relationship.target).name or f"image_{context.image_counter}.bin"
        safe_name = self._safe_filename(filename)
        if context.image_dir is None:
            context.warnings.append("Image export directory unavailable")
            return "", None
        while (context.image_dir / safe_name).exists():
            context.image_counter += 1
            stem = Path(filename).stem or "image"
            ext = Path(filename).suffix or ".bin"
            safe_name = self._safe_filename(f"{stem}_{context.image_counter}{ext}")

        saved_path = context.image_dir / safe_name
        saved_path.write_bytes(image_bytes)
        context.image_cache[relationship_id] = saved_path
        context.image_counter += 1

        image_ref = self._render_image_ref(saved_path, context.output_path.parent)
        return image_ref, {
            "relationshipId": relationship_id,
            "target": relationship.target,
            "relativePath": image_ref[8:-1] if image_ref.startswith("![image](") else image_ref,
        }

    def _render_image_ref(self, image_path: Path, output_dir: Path) -> str:
        relative = os.path.relpath(image_path, output_dir)
        return f"![image]({relative.replace('\\', '/')})"

    def _resolve_word_target(self, target: str) -> str:
        if target.startswith("/"):
            return target.lstrip("/")
        normalized = PurePosixPath("word") / PurePosixPath(target)
        return str(normalized)

    def _safe_filename(self, filename: str) -> str:
        filename = re.sub(r"[^A-Za-z0-9._-]", "_", filename)
        return filename or "image.bin"

    def _escape_table_cell_markdown(self, text: str) -> str:
        return text.replace("|", r"\|")

    def _parse_table(self, table: ET.Element, context: ConversionContext) -> tuple[str, dict[str, object] | None]:
        rows: list[list[str]] = []
        has_merge = False
        active_rowspans: dict[int, str] = {}

        for row in table.findall(_qn(W_NS, "tr")):
            cell_texts: list[str] = []
            col_index = 0
            for cell in row.findall(_qn(W_NS, "tc")):
                tc_pr = cell.find(_qn(W_NS, "tcPr"))
                colspan = 1
                vmerge_mode: str | None = None
                if tc_pr is not None:
                    grid_span = tc_pr.find(_qn(W_NS, "gridSpan"))
                    if grid_span is not None:
                        colspan_raw = grid_span.attrib.get(_qn(W_NS, "val"), "1")
                        colspan = int(colspan_raw) if colspan_raw.isdigit() else 1
                    vmerge_node = tc_pr.find(_qn(W_NS, "vMerge"))
                    if vmerge_node is not None:
                        vmerge_mode = vmerge_node.attrib.get(_qn(W_NS, "val"), "continue") or "continue"
                    if colspan > 1 or vmerge_mode is not None:
                        has_merge = True
                cell_text = self._extract_cell_text(cell, context)

                if vmerge_mode == "continue":
                    cell_text = active_rowspans.get(col_index, cell_text)

                if vmerge_mode == "restart":
                    for span in range(colspan):
                        active_rowspans[col_index + span] = cell_text
                elif vmerge_mode is None:
                    for span in range(colspan):
                        active_rowspans.pop(col_index + span, None)

                repeated = [cell_text] * max(colspan, 1)
                cell_texts.extend(repeated)
                col_index += max(colspan, 1)
            if cell_texts:
                rows.append(cell_texts)

        if not rows:
            return "", None

        context.stats.tables += 1
        width_set = {len(row) for row in rows}
        max_cols = max(width_set)
        normalized_rows = [row + [""] * (max_cols - len(row)) for row in rows]
        if has_merge or len(width_set) != 1:
            context.warnings.append("Complex table normalized to pipe table; merged layout simplified.")

        header = [self._escape_table_cell_markdown(value) for value in normalized_rows[0]]
        separator = ["---"] * max_cols
        markdown_rows = [
            "| " + " | ".join(header) + " |",
            "| " + " | ".join(separator) + " |",
        ]
        for row in normalized_rows[1:]:
            padded = [self._escape_table_cell_markdown(value) for value in row]
            markdown_rows.append("| " + " | ".join(padded) + " |")
        table_md = "\n".join(markdown_rows)

        return table_md, None

    def _extract_cell_text(self, cell: ET.Element, context: ConversionContext) -> str:
        paragraphs = cell.findall(_qn(W_NS, "p"))
        texts: list[str] = []
        for paragraph in paragraphs:
            text = self._extract_plain_paragraph_text(paragraph, context)
            if text:
                texts.append(text)
        return "<br>".join(texts) if texts else ""

    def _extract_plain_paragraph_text(self, paragraph: ET.Element, context: ConversionContext) -> str:
        chunks: list[str] = []
        for child in list(paragraph):
            name = _local_name(child.tag)
            if name == "r":
                chunks.extend(self._extract_plain_run_text(child, context))
            elif name in {"oMath", "oMathPara"}:
                chunks.append(self._parse_equation(child, context, block=False))
        return _normalize_ws("".join(chunks))

    def _extract_plain_run_text(self, run: ET.Element, context: ConversionContext) -> list[str]:
        values: list[str] = []
        for child in list(run):
            name = _local_name(child.tag)
            if name == "t":
                values.append(child.text or "")
            elif name == "tab":
                values.append(" ")
            elif name in {"br", "cr"}:
                values.append(" ")
            elif name in {"oMath", "oMathPara"}:
                values.append(self._parse_equation(child, context, block=False))
        return values

    def _get_paragraph_style_id(self, paragraph: ET.Element) -> str:
        p_style = paragraph.find("./w:pPr/w:pStyle", {"w": W_NS})
        if p_style is None:
            return ""
        return p_style.attrib.get(_qn(W_NS, "val"), "")

    def _heading_level_from_text(self, content: str) -> int:
        normalized = _normalize_ws(content)
        if not normalized:
            return 0
        if CAPTION_LINE_RE.match(normalized):
            return 0
        numbered = HEADING_NUMBER_RE.match(normalized)
        if numbered:
            depth = numbered.group(1).replace("．", ".").count(".") + 1
            return min(max(depth, 1), 3)
        return 0

    def _heading_level_from_style(self, style_id: str, style_name: str) -> int:
        style_tokens = f"{style_id} {style_name}".strip().lower()
        if not style_tokens:
            return 0
        if "papertitle" in style_tokens or re.search(r"\btitle\b", style_tokens):
            return 1
        if "paperheading1" in style_tokens:
            return 1
        if "paperheading2" in style_tokens:
            return 2
        match = HEADING_STYLE_RE.search(style_tokens)
        if match:
            return min(int(match.group(1)), 3)
        return 0

    def _heading_level(self, content: str, style_id: str, style_name: str) -> int:
        text_level = self._heading_level_from_text(content)
        style_level = self._heading_level_from_style(style_id, style_name)
        if CAPTION_LINE_RE.match(_normalize_ws(content)):
            return 0
        if text_level and style_level:
            return max(text_level, style_level)
        return text_level or style_level

    def _looks_like_title_candidate(self, content: str) -> bool:
        normalized = _normalize_ws(content)
        if not normalized:
            return False
        if CAPTION_LINE_RE.match(normalized):
            return False
        if HEADING_NUMBER_RE.match(normalized):
            return False
        if ABSTRACT_LABEL_RE.match(normalized):
            return False
        if KEYWORD_LABEL_RE.match(normalized):
            return False
        return True

    def _extract_section_columns_from_paragraph(self, paragraph: ET.Element) -> int | None:
        ppr = paragraph.find(_qn(W_NS, "pPr"))
        if ppr is None:
            return None
        sect_pr = ppr.find(_qn(W_NS, "sectPr"))
        if sect_pr is None:
            return None
        return self._extract_section_columns(sect_pr)

    def _extract_section_columns(self, sect_pr: ET.Element | None) -> int:
        if sect_pr is None:
            return 1
        cols = sect_pr.find(_qn(W_NS, "cols"))
        if cols is None:
            return 1
        raw = cols.attrib.get(_qn(W_NS, "num"), "1")
        return int(raw) if raw.isdigit() and int(raw) > 0 else 1

    def _list_level(self, paragraph: ET.Element) -> int:
        num_pr = paragraph.find("./w:pPr/w:numPr", {"w": W_NS})
        if num_pr is None:
            return -1
        ilvl = num_pr.find("./w:ilvl", {"w": W_NS})
        if ilvl is None:
            return 0
        raw = ilvl.attrib.get(_qn(W_NS, "val"), "0")
        return int(raw) if raw.isdigit() else 0
