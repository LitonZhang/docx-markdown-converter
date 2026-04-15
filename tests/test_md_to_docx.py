from __future__ import annotations

from pathlib import Path
import sys
import unittest
from xml.etree import ElementTree as ET

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT / "converter"))

from convert_md_to_docx import (  # noqa: E402
    apply_table_settings,
    apply_table_grid_style,
    apply_style_config_to_styles_xml,
    apply_semantic_styles,
    extract_title_from_markdown,
    get_paragraph_style,
    normalize_markdown_math,
    parse_style_config,
    spacing_to_word,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def build_document(paragraphs: list[str]) -> ET.Element:
    body = "".join(
        f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>" for text in paragraphs
    )
    xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<w:document xmlns:w="{W_NS}"><w:body>{body}</w:body></w:document>'
    )
    return ET.fromstring(xml)


class MdToDocxTests(unittest.TestCase):
    def test_style_names_are_short_chinese(self) -> None:
        styles_xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            f'<w:styles xmlns:w="{W_NS}">'
            '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
            "<w:latentStyles/>"
            "</w:styles>"
        )
        styles_root = ET.fromstring(styles_xml)
        apply_style_config_to_styles_xml(styles_root, parse_style_config(self._style_config_raw()))

        expected_names = {
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

        for style_id, expected in expected_names.items():
            style = styles_root.find(f".//{{{W_NS}}}style[@{{{W_NS}}}styleId='{style_id}']")
            self.assertIsNotNone(style, f"Missing style: {style_id}")
            name = style.find(f"{{{W_NS}}}name")
            self.assertIsNotNone(name)
            self.assertEqual(name.get(f"{{{W_NS}}}val"), expected)

    def test_style_definition_matches_config(self) -> None:
        styles_xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            f'<w:styles xmlns:w="{W_NS}">'
            '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
            "<w:latentStyles/>"
            "</w:styles>"
        )
        styles_root = ET.fromstring(styles_xml)
        raw = self._style_config_raw()
        raw["advancedDefaults"] = {
            "before": {"mode": "pt", "value": 0},
            "after": {"mode": "pt", "value": 0},
            "firstLineIndentChars": 0,
            "bold": False,
            "italic": False,
        }
        raw["body"] = {
            "zhFont": "黑体",
            "enFont": "Arial",
            "fontSizePt": 13,
            "lineSpacingMode": "fixed",
            "lineSpacingValue": 20,
            "align": "justify",
            "advancedOverride": {
                "before": {"mode": "lines", "value": 1.5},
                "after": {"mode": "pt", "value": 6},
                "firstLineIndentChars": 2,
                "bold": True,
                "italic": True,
            },
        }
        apply_style_config_to_styles_xml(styles_root, parse_style_config(raw))

        style = styles_root.find(f".//{{{W_NS}}}style[@{{{W_NS}}}styleId='PaperBody']")
        self.assertIsNotNone(style)

        rpr = style.find(f"{{{W_NS}}}rPr")
        self.assertIsNotNone(rpr)
        fonts = rpr.find(f"{{{W_NS}}}rFonts")
        self.assertIsNotNone(fonts)
        self.assertEqual(fonts.get(f"{{{W_NS}}}eastAsia"), "黑体")
        self.assertEqual(fonts.get(f"{{{W_NS}}}ascii"), "Arial")

        sz = rpr.find(f"{{{W_NS}}}sz")
        self.assertIsNotNone(sz)
        self.assertEqual(sz.get(f"{{{W_NS}}}val"), "26")
        self.assertIsNotNone(rpr.find(f"{{{W_NS}}}b"))
        self.assertIsNotNone(rpr.find(f"{{{W_NS}}}i"))

        ppr = style.find(f"{{{W_NS}}}pPr")
        self.assertIsNotNone(ppr)
        spacing = ppr.find(f"{{{W_NS}}}spacing")
        self.assertIsNotNone(spacing)
        self.assertEqual(spacing.get(f"{{{W_NS}}}line"), "400")
        self.assertEqual(spacing.get(f"{{{W_NS}}}lineRule"), "exact")
        self.assertEqual(spacing.get(f"{{{W_NS}}}beforeLines"), "150")
        self.assertEqual(spacing.get(f"{{{W_NS}}}after"), "120")
        jc = ppr.find(f"{{{W_NS}}}jc")
        self.assertIsNotNone(jc)
        self.assertEqual(jc.get(f"{{{W_NS}}}val"), "both")
        ind = ppr.find(f"{{{W_NS}}}ind")
        self.assertIsNotNone(ind)
        self.assertEqual(ind.get(f"{{{W_NS}}}firstLineChars"), "200")

    def test_parse_style_config_legacy_bold_italic_compatible(self) -> None:
        raw = self._style_config_raw()
        raw["title"]["bold"] = True
        raw["title"]["italic"] = False
        config = parse_style_config(raw)
        self.assertTrue(config.title.bold)
        self.assertFalse(config.title.italic)

    def test_parse_style_config_legacy_before_after_pt_compatible(self) -> None:
        raw = self._style_config_raw()
        raw["body"]["advancedOverride"] = {
            "beforePt": 8,
            "afterPt": 12,
            "firstLineIndentChars": 0,
            "bold": False,
            "italic": False,
        }
        config = parse_style_config(raw)
        self.assertEqual(config.body.before.mode, "pt")
        self.assertEqual(config.body.before.value, 8)
        self.assertEqual(config.body.after.mode, "pt")
        self.assertEqual(config.body.after.value, 12)

    def test_extract_title_from_plain_h1(self) -> None:
        text = "# 论文题目\n\n## 1 引言\n正文"
        transformed, title = extract_title_from_markdown(text)
        self.assertEqual(title, "论文题目")
        self.assertIn("title:", transformed)
        self.assertNotIn("# 论文题目", transformed)

    def test_extract_title_ignores_numbered_h1(self) -> None:
        text = "# 1 引言\n\n正文"
        transformed, title = extract_title_from_markdown(text)
        self.assertIsNone(title)
        self.assertEqual(transformed, text)

    def test_extract_title_ignores_bold_h1(self) -> None:
        text = "# **论文题目**\n\n正文"
        transformed, title = extract_title_from_markdown(text)
        self.assertIsNone(title)
        self.assertEqual(transformed, text)

    def test_spacing_conversion(self) -> None:
        line, rule = spacing_to_word("multiple", 1.5)
        self.assertEqual(rule, "auto")
        self.assertEqual(line, "360")

        line_fixed, rule_fixed = spacing_to_word("fixed", 18)
        self.assertEqual(rule_fixed, "exact")
        self.assertEqual(line_fixed, "360")

    def test_apply_semantic_styles(self) -> None:
        root = build_document(["摘要", "这是中文摘要正文", "图1 系统框图", "1 引言"])
        body_start = apply_semantic_styles(root)

        paragraphs = root.findall(f".//{{{W_NS}}}body/{{{W_NS}}}p")
        styles = [get_paragraph_style(p) for p in paragraphs]
        self.assertEqual(styles[0], "PaperAbstractZh")
        self.assertEqual(styles[1], "PaperAbstractZh")
        self.assertEqual(styles[2], "PaperFigureCaption")
        self.assertEqual(styles[3], "PaperHeading1")
        self.assertEqual(body_start, 3)

    def test_apply_semantic_styles_first_h1_becomes_title(self) -> None:
        xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            f'<w:document xmlns:w="{W_NS}"><w:body>'
            '<w:p><w:pPr><w:pStyle w:val="Heading1" /></w:pPr><w:r><w:t>论文题目</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>摘要</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>摘要正文</w:t></w:r></w:p>'
            '<w:p><w:pPr><w:pStyle w:val="Heading1" /></w:pPr><w:r><w:t>1 引言</w:t></w:r></w:p>'
            '<w:p><w:pPr><w:pStyle w:val="Heading2" /></w:pPr><w:r><w:t>1.1 方法</w:t></w:r></w:p>'
            '<w:p><w:pPr><w:pStyle w:val="Heading3" /></w:pPr><w:r><w:t>1.1.1 细节</w:t></w:r></w:p>'
            '</w:body></w:document>'
        )
        root = ET.fromstring(xml)
        body_start = apply_semantic_styles(root)

        paragraphs = root.findall(f".//{{{W_NS}}}body/{{{W_NS}}}p")
        styles = [get_paragraph_style(p) for p in paragraphs]
        self.assertEqual(styles[0], "PaperTitle")
        self.assertEqual(styles[1], "PaperAbstractZh")
        self.assertEqual(styles[2], "PaperAbstractZh")
        self.assertEqual(styles[3], "PaperHeading1")
        self.assertEqual(styles[4], "PaperHeading2")
        self.assertEqual(styles[5], "PaperHeading3")
        self.assertEqual(body_start, 3)

    def test_apply_semantic_styles_first_plain_paragraph_can_be_title(self) -> None:
        root = build_document(["论文题目", "摘要", "摘要正文", "1 引言"])
        body_start = apply_semantic_styles(root)

        paragraphs = root.findall(f".//{{{W_NS}}}body/{{{W_NS}}}p")
        styles = [get_paragraph_style(p) for p in paragraphs]
        self.assertEqual(styles[0], "PaperTitle")
        self.assertEqual(styles[1], "PaperAbstractZh")
        self.assertEqual(styles[2], "PaperAbstractZh")
        self.assertEqual(styles[3], "PaperHeading1")
        self.assertEqual(body_start, 3)

    def test_apply_semantic_styles_without_heading1_returns_none_start(self) -> None:
        root = build_document(["前置说明", "普通正文", "更多正文"])
        body_start = apply_semantic_styles(root)
        paragraphs = root.findall(f".//{{{W_NS}}}body/{{{W_NS}}}p")
        styles = [get_paragraph_style(p) for p in paragraphs]
        self.assertEqual(styles[0], "PaperBody")
        self.assertEqual(styles[1], "PaperBody")
        self.assertEqual(styles[2], "PaperBody")
        self.assertIsNone(body_start)

    def test_parse_style_config_heading3_falls_back_to_heading2(self) -> None:
        raw = {
            "title": self._basic_style(),
            "abstractZh": self._basic_style(),
            "abstractEn": self._basic_style(),
            "heading1": self._basic_style(),
            "heading2": self._basic_style(),
            "figureCaption": self._basic_style(),
            "tableCaption": self._basic_style(),
            "body": self._basic_style(),
        }
        config = parse_style_config(raw)
        self.assertEqual(config.heading3.font_size_pt, config.heading2.font_size_pt)
        self.assertEqual(config.heading3.align, config.heading2.align)

    def test_parse_style_config_accepts_explicit_heading3(self) -> None:
        raw = {
            "title": self._basic_style(),
            "abstractZh": self._basic_style(),
            "abstractEn": self._basic_style(),
            "heading1": self._basic_style(),
            "heading2": self._basic_style(),
            "heading3": {**self._basic_style(), "fontSizePt": 11},
            "figureCaption": self._basic_style(),
            "tableCaption": self._basic_style(),
            "body": self._basic_style(),
        }
        config = parse_style_config(raw)
        self.assertEqual(config.heading3.font_size_pt, 11)

    def test_apply_table_grid_style_sets_table_style(self) -> None:
        xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            f'<w:document xmlns:w="{W_NS}"><w:body>'
            '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
            '<w:tbl><w:tblPr><w:tblStyle w:val="Table"/></w:tblPr>'
            '<w:tr><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
            '</w:body></w:document>'
        )
        root = ET.fromstring(xml)
        apply_table_grid_style(root)

        tables = root.findall(f".//{{{W_NS}}}body/{{{W_NS}}}tbl")
        self.assertEqual(len(tables), 2)
        for table in tables:
            tbl_pr = table.find(f"{{{W_NS}}}tblPr")
            self.assertIsNotNone(tbl_pr)
            tbl_style = tbl_pr.find(f"{{{W_NS}}}tblStyle")
            self.assertIsNotNone(tbl_style)
            self.assertEqual(tbl_style.get(f"{{{W_NS}}}val"), "TableGrid")

    def test_apply_table_settings_table_preset_sets_tbl_style(self) -> None:
        xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            f'<w:document xmlns:w="{W_NS}"><w:body>'
            '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblBorders><w:top w:val="single"/></w:tblBorders></w:tblPr>'
            '<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
            '</w:body></w:document>'
        )
        root = ET.fromstring(xml)
        config_raw = self._style_config_raw()
        config_raw["tableSettings"] = {
            "tablePreset": "table",
            "headerBold": False,
            "applyTextStyle": False,
            "textStyle": self._basic_style(),
        }
        settings = parse_style_config(config_raw).table_settings
        apply_table_settings(root, settings)

        tbl_pr = root.find(f".//{{{W_NS}}}tbl/{{{W_NS}}}tblPr")
        self.assertIsNotNone(tbl_pr)
        tbl_style = tbl_pr.find(f"{{{W_NS}}}tblStyle")
        self.assertIsNotNone(tbl_style)
        self.assertEqual(tbl_style.get(f"{{{W_NS}}}val"), "Table")
        self.assertIsNone(tbl_pr.find(f"{{{W_NS}}}tblBorders"))

    def test_parse_style_config_missing_table_settings_keeps_legacy_behavior(self) -> None:
        config = parse_style_config(self._style_config_raw())
        self.assertEqual(config.table_settings.table_preset, "tableGrid")
        self.assertFalse(config.table_settings.apply_text_style)
        self.assertFalse(config.table_settings.header_bold)

    def test_apply_table_settings_three_line_and_text_style(self) -> None:
        xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            f'<w:document xmlns:w="{W_NS}"><w:body>'
            '<w:tbl>'
            '<w:tr><w:tc><w:p><w:r><w:t>H</w:t></w:r></w:p></w:tc></w:tr>'
            '<w:tr><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>'
            '</w:tbl>'
            '</w:body></w:document>'
        )
        root = ET.fromstring(xml)
        config_raw = self._style_config_raw()
        config_raw["tableSettings"] = {
            "tablePreset": "threeLine",
            "headerBold": False,
            "applyTextStyle": True,
            "textStyle": {
                "zhFont": "宋体",
                "enFont": "Times New Roman",
                "fontSizePt": 12,
                "lineSpacingMode": "multiple",
                "lineSpacingValue": 1.0,
                "align": "center",
                "advancedOverride": {
                    "before": {"mode": "pt", "value": 0},
                    "after": {"mode": "pt", "value": 0},
                    "firstLineIndentChars": 0,
                    "bold": False,
                    "italic": False,
                },
            },
        }
        settings = parse_style_config(config_raw).table_settings
        apply_table_settings(root, settings)

        table = root.find(f".//{{{W_NS}}}tbl")
        self.assertIsNotNone(table)
        tbl_pr = table.find(f"{{{W_NS}}}tblPr")
        self.assertIsNotNone(tbl_pr)
        tbl_borders = tbl_pr.find(f"{{{W_NS}}}tblBorders")
        self.assertIsNotNone(tbl_borders)
        self.assertEqual(
            tbl_borders.find(f"{{{W_NS}}}top").get(f"{{{W_NS}}}val"),
            "single",
        )
        self.assertEqual(
            tbl_borders.find(f"{{{W_NS}}}bottom").get(f"{{{W_NS}}}val"),
            "single",
        )
        self.assertEqual(
            tbl_borders.find(f"{{{W_NS}}}insideV").get(f"{{{W_NS}}}val"),
            "nil",
        )

        header_para = root.find(f".//{{{W_NS}}}tr[1]//{{{W_NS}}}p")
        self.assertIsNotNone(header_para)
        ppr = header_para.find(f"{{{W_NS}}}pPr")
        self.assertIsNotNone(ppr)
        jc = ppr.find(f"{{{W_NS}}}jc")
        self.assertIsNotNone(jc)
        self.assertEqual(jc.get(f"{{{W_NS}}}val"), "center")
        spacing = ppr.find(f"{{{W_NS}}}spacing")
        self.assertIsNotNone(spacing)
        self.assertEqual(spacing.get(f"{{{W_NS}}}line"), "240")
        ind = ppr.find(f"{{{W_NS}}}ind")
        self.assertIsNotNone(ind)
        self.assertEqual(ind.get(f"{{{W_NS}}}firstLineChars"), "0")

        header_run = header_para.find(f".//{{{W_NS}}}r")
        self.assertIsNotNone(header_run)
        rpr = header_run.find(f"{{{W_NS}}}rPr")
        self.assertIsNotNone(rpr)
        fonts = rpr.find(f"{{{W_NS}}}rFonts")
        self.assertIsNotNone(fonts)
        self.assertEqual(fonts.get(f"{{{W_NS}}}eastAsia"), "宋体")
        self.assertEqual(fonts.get(f"{{{W_NS}}}ascii"), "Times New Roman")
        sz = rpr.find(f"{{{W_NS}}}sz")
        self.assertIsNotNone(sz)
        self.assertEqual(sz.get(f"{{{W_NS}}}val"), "24")

    def test_apply_table_settings_header_bold_compat_mode(self) -> None:
        xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            f'<w:document xmlns:w="{W_NS}"><w:body>'
            '<w:tbl>'
            '<w:tr><w:tc><w:p><w:r><w:t>H</w:t></w:r></w:p></w:tc></w:tr>'
            '<w:tr><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>'
            '</w:tbl>'
            '</w:body></w:document>'
        )
        root = ET.fromstring(xml)
        config_raw = self._style_config_raw()
        config_raw["tableSettings"] = {
            "tablePreset": "tableGrid",
            "headerBold": True,
            "applyTextStyle": False,
            "textStyle": self._basic_style(),
        }
        settings = parse_style_config(config_raw).table_settings
        apply_table_settings(root, settings)

        header_run = root.find(f".//{{{W_NS}}}tr[1]//{{{W_NS}}}r")
        body_run = root.find(f".//{{{W_NS}}}tr[2]//{{{W_NS}}}r")
        self.assertIsNotNone(header_run)
        self.assertIsNotNone(body_run)
        self.assertIsNotNone(header_run.find(f".//{{{W_NS}}}b"))
        self.assertIsNone(body_run.find(f".//{{{W_NS}}}b"))

    def test_apply_semantic_styles_clears_conflicting_paragraph_properties(self) -> None:
        xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            f'<w:document xmlns:w="{W_NS}"><w:body>'
            '<w:p><w:r><w:t>1 引言</w:t></w:r></w:p>'
            '<w:p><w:pPr><w:numPr><w:ilvl w:val="0" /><w:numId w:val="1" /></w:numPr>'
            '<w:jc w:val="right" /><w:spacing w:line="720" w:lineRule="auto" />'
            '</w:pPr><w:r><w:t>列表项</w:t></w:r></w:p>'
            '</w:body></w:document>'
        )
        root = ET.fromstring(xml)
        apply_semantic_styles(root)

        paragraphs = root.findall(f".//{{{W_NS}}}body/{{{W_NS}}}p")
        self.assertEqual(get_paragraph_style(paragraphs[1]), "PaperBody")

        ppr = paragraphs[1].find(f"{{{W_NS}}}pPr")
        self.assertIsNotNone(ppr)
        self.assertIsNotNone(ppr.find(f"{{{W_NS}}}numPr"))
        self.assertIsNone(ppr.find(f"{{{W_NS}}}jc"))
        self.assertIsNone(ppr.find(f"{{{W_NS}}}spacing"))

    def test_normalize_markdown_math_for_pandoc(self) -> None:
        source = r"$L=\left‖y_{i}-ŷ_{i}\right‖+λ\left‖z_{time}-ẑ_{time}\right‖$"
        normalized = normalize_markdown_math(source)
        self.assertIn(r"\left\|", normalized)
        self.assertIn(r"\right\|", normalized)
        self.assertIn(r"\hat{y}", normalized)
        self.assertIn(r"\hat{z}", normalized)
        self.assertIn(r"\lambda", normalized)

    def test_normalize_markdown_math_repairs_brace_set_and_subscripts(self) -> None:
        source = (
            r"$O{DES}^{(k)} = f{s}\left( P^{(k)} \right) = "
            r"\left{ Y{due}^{(k)},Y{capacity}^{(k)},Y_{bottleneck}^{(k)},\ldots \right}$"
        )
        normalized = normalize_markdown_math(source)
        self.assertIn(r"O_{DES}^{(k)}", normalized)
        self.assertIn(r"f_{s}\left", normalized)
        self.assertIn(r"\left\{", normalized)
        self.assertIn(r"Y_{due}^{(k)}", normalized)
        self.assertIn(r"Y_{capacity}^{(k)}", normalized)
        self.assertIn(r"\right\}", normalized)

    def test_normalize_markdown_math_keeps_command_braces(self) -> None:
        source = r"$\text{abc} + \sin{x}$"
        normalized = normalize_markdown_math(source)
        self.assertIn(r"\text{abc}", normalized)
        self.assertIn(r"\sin{x}", normalized)

    @staticmethod
    def _basic_style() -> dict[str, object]:
        return {
            "zhFont": "宋体",
            "enFont": "Times New Roman",
            "fontSizePt": 12,
            "lineSpacingMode": "multiple",
            "lineSpacingValue": 1.5,
            "align": "left",
            "bold": False,
            "italic": False,
        }

    @classmethod
    def _style_config_raw(cls) -> dict[str, object]:
        base = cls._basic_style()
        return {
            "title": base.copy(),
            "abstractZh": base.copy(),
            "abstractEn": base.copy(),
            "heading1": base.copy(),
            "heading2": base.copy(),
            "heading3": base.copy(),
            "figureCaption": base.copy(),
            "tableCaption": base.copy(),
            "body": base.copy(),
        }


if __name__ == "__main__":
    unittest.main()
