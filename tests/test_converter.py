from __future__ import annotations

from pathlib import Path
import sys
import tempfile
import unittest
import zipfile

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT / "converter"))

from docx_to_md.converter import DocxToMarkdownConverter


CONTENT_TYPES = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
  <Default Extension=\"xml\" ContentType=\"application/xml\"/>
  <Default Extension=\"png\" ContentType=\"image/png\"/>
  <Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>
</Types>
"""

ROOT_RELS = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>
</Relationships>
"""

EMPTY_DOC_RELS = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>
"""


def build_docx(doc_xml: str, rels_xml: str = EMPTY_DOC_RELS, media: dict[str, bytes] | None = None) -> bytes:
    media = media or {}
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as fp:
        path = Path(fp.name)

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", ROOT_RELS)
        zf.writestr("word/document.xml", doc_xml)
        zf.writestr("word/_rels/document.xml.rels", rels_xml)
        for item_path, data in media.items():
            zf.writestr(item_path, data)

    return path.read_bytes()


class ConverterTests(unittest.TestCase):
    def setUp(self) -> None:
        self.converter = DocxToMarkdownConverter()

    def _convert_with_result(self, docx_bytes: bytes) -> tuple[str, object]:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            input_path = tmp_path / "input.docx"
            output_path = tmp_path / "output.md"
            image_dir = tmp_path / "assets"
            input_path.write_bytes(docx_bytes)

            result = self.converter.convert_file(input_path, output_path, image_dir)
            self.assertTrue(result.success)
            self.assertTrue(output_path.exists())
            return output_path.read_text(encoding="utf-8"), result

    def _convert(self, docx_bytes: bytes) -> str:
        markdown, _ = self._convert_with_result(docx_bytes)
        return markdown

    def test_heading_and_bold(self) -> None:
        doc_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>
      <w:r><w:t>Title</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""
        md = self._convert(build_docx(doc_xml))
        self.assertIn("Title", md)
        self.assertIn("# Title", md)
        self.assertIn("Bold", md)
        self.assertNotIn("**Bold**", md)

    def test_title_style_becomes_first_h1(self) -> None:
        doc_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val=\"Title\"/></w:pPr>
      <w:r><w:t>论文题目</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""
        md = self._convert(build_docx(doc_xml))
        self.assertIn("# 论文题目", md)

    def test_first_plain_paragraph_title_becomes_h1(self) -> None:
        doc_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
  <w:body>
    <w:p><w:r><w:t>论文题目</w:t></w:r></w:p>
    <w:p><w:r><w:t>摘要</w:t></w:r></w:p>
    <w:p><w:r><w:t>这是摘要</w:t></w:r></w:p>
  </w:body>
</w:document>
"""
        md = self._convert(build_docx(doc_xml))
        self.assertIn("# 论文题目", md)

    def test_paper_heading_style_detected(self) -> None:
        doc_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val=\"PaperHeading1\"/></w:pPr>
      <w:r><w:t>1 引言</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""
        md = self._convert(build_docx(doc_xml))
        self.assertIn("# 1 引言", md)

    def test_caption_line_not_promoted_to_heading(self) -> None:
        doc_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>
      <w:r><w:t>表1 参数设置</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""
        md = self._convert(build_docx(doc_xml))
        self.assertIn("表1 参数设置", md)
        self.assertNotIn("# 表1 参数设置", md)

    def test_equation_to_latex(self) -> None:
        doc_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"
            xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">
  <w:body>
    <w:p>
      <m:oMath>
        <m:sSup>
          <m:e><m:r><m:t>x</m:t></m:r></m:e>
          <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
        </m:sSup>
      </m:oMath>
    </w:p>
  </w:body>
</w:document>
"""
        md = self._convert(build_docx(doc_xml))
        self.assertIn("$x^{2}$", md)

    def test_equation_unicode_delimiter_and_symbols(self) -> None:
        doc_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"
            xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">
  <w:body>
    <w:p>
      <m:oMath>
        <m:d>
          <m:dPr>
            <m:begChr m:val=\"‖\"/>
            <m:endChr m:val=\"‖\"/>
          </m:dPr>
          <m:e>
            <m:r><m:t>ŷ</m:t></m:r>
            <m:r><m:t>+λ</m:t></m:r>
          </m:e>
        </m:d>
      </m:oMath>
    </w:p>
  </w:body>
</w:document>
"""
        md = self._convert(build_docx(doc_xml))
        self.assertIn(r"$\left\|", md)
        self.assertIn(r"\hat{y}", md)
        self.assertIn(r"\lambda", md)

    def test_table_markdown(self) -> None:
        doc_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>H1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>H2</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>
"""
        md = self._convert(build_docx(doc_xml))
        self.assertIn("| H1 | H2 |", md)
        self.assertIn("| A | B |", md)

    def test_table_cell_pipe_is_escaped(self) -> None:
        doc_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>键</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>值</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A|B</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>
"""
        md = self._convert(build_docx(doc_xml))
        self.assertIn("| 键 | 值 |", md)
        self.assertIn("| A\\|B | C |", md)

    def test_complex_table_normalizes_to_pipe_table(self) -> None:
        doc_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:tcPr><w:gridSpan w:val=\"2\"/></w:tcPr>
          <w:p><w:r><w:t>跨列</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>
"""
        md, result = self._convert_with_result(build_docx(doc_xml))
        self.assertIn("| 跨列 | 跨列 |", md)
        self.assertIn("| A | B |", md)
        self.assertTrue(
            any("normalized to pipe table" in warning for warning in result.warnings),
            "missing complex table normalization warning",
        )

    def test_extract_image(self) -> None:
        doc_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"
            xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"
            xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"
            xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"
            xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <a:graphic>
              <a:graphicData>
                <pic:pic>
                  <pic:blipFill>
                    <a:blip r:embed=\"rIdImg1\"/>
                  </pic:blipFill>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>
"""
        rels_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rIdImg1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.png\"/>
</Relationships>
"""
        png_bytes = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR"

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            input_path = tmp_path / "input.docx"
            output_path = tmp_path / "output.md"
            image_dir = tmp_path / "assets"
            image_converter = DocxToMarkdownConverter(extract_images=True)

            input_path.write_bytes(build_docx(doc_xml, rels_xml=rels_xml, media={"word/media/image1.png": png_bytes}))
            result = image_converter.convert_file(input_path, output_path, image_dir)

            self.assertTrue(result.success)
            md = output_path.read_text(encoding="utf-8")
            self.assertIn("![image](assets/image1.png)", md)
            self.assertTrue((image_dir / "image1.png").exists())

    def test_no_image_export_by_default(self) -> None:
        doc_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <a:graphic>
              <a:graphicData>
                <pic:pic>
                  <pic:blipFill>
                    <a:blip r:embed="rIdImg1"/>
                  </pic:blipFill>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>
"""
        rels_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>
"""
        png_bytes = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR"

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            input_path = tmp_path / "input.docx"
            output_path = tmp_path / "output.md"
            image_dir = tmp_path / "assets"

            input_path.write_bytes(build_docx(doc_xml, rels_xml=rels_xml, media={"word/media/image1.png": png_bytes}))
            result = self.converter.convert_file(input_path, output_path, image_dir)

            self.assertTrue(result.success)
            md = output_path.read_text(encoding="utf-8")
            self.assertNotIn("![image](", md)
            self.assertFalse((image_dir / "image1.png").exists())

    def test_md_only_mode_does_not_write_roundtrip_metadata(self) -> None:
        doc_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
  <w:body>
    <w:p><w:r><w:t>1 引言</w:t></w:r></w:p>
    <w:p><w:r><w:t>正文内容</w:t></w:r></w:p>
  </w:body>
</w:document>
"""
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            input_path = tmp_path / "input.docx"
            output_path = tmp_path / "output.md"
            input_path.write_bytes(build_docx(doc_xml))

            result = self.converter.convert_file(input_path, output_path)
            metadata_path = tmp_path / "output_assets" / "roundtrip.json"
            self.assertFalse(metadata_path.exists())
            self.assertIsNone(result.assets_dir)


if __name__ == "__main__":
    unittest.main()
