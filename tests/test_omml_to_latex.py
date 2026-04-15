from __future__ import annotations

from pathlib import Path
import sys
import unittest
from xml.etree import ElementTree as ET

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT / "converter"))

from docx_to_md.omml_to_latex import convert_omml_to_latex


class OmmlToLatexTests(unittest.TestCase):
    def test_piecewise_eqarr_uses_array_and_one_sided_brace(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:d>
    <m:dPr>
      <m:begChr m:val="{"/>
      <m:endChr m:val=""/>
    </m:dPr>
    <m:e>
      <m:eqArr>
        <m:e>
          <m:f>
            <m:num>
              <m:r><m:t>l-l'</m:t></m:r>
            </m:num>
            <m:den>
              <m:r><m:t>l+l'</m:t></m:r>
            </m:den>
          </m:f>
          <m:r><m:t>,(</m:t></m:r>
          <m:sSub>
            <m:e><m:r><m:t>a</m:t></m:r></m:e>
            <m:sub><m:r><m:t>i</m:t></m:r></m:sub>
          </m:sSub>
          <m:r><m:t>≠0且满足</m:t></m:r>
          <m:d>
            <m:dPr>
              <m:begChr m:val="("/>
              <m:endChr m:val=")"/>
            </m:dPr>
            <m:e><m:r><m:t>20</m:t></m:r></m:e>
          </m:d>
          <m:r><m:t>)</m:t></m:r>
        </m:e>
        <m:e>
          <m:r><m:t>0,(</m:t></m:r>
          <m:sSub>
            <m:e><m:r><m:t>a</m:t></m:r></m:e>
            <m:sub><m:r><m:t>i</m:t></m:r></m:sub>
          </m:sSub>
          <m:r><m:t>=0)</m:t></m:r>
        </m:e>
      </m:eqArr>
    </m:e>
  </m:d>
</m:oMath>
"""
        latex, warnings = convert_omml_to_latex(ET.fromstring(xml))
        self.assertEqual(warnings, [])
        self.assertEqual(
            latex,
            r"\left\{\begin{array}{l}\frac{l-l'}{l+l'},(a_{i}≠0且满足\left(20\right)) \\ 0,(a_{i}=0)\end{array}\right.",
        )

    def test_matrix_rows_render_as_aligned_array(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:m>
    <m:mr>
      <m:e><m:r><m:t>a</m:t></m:r></m:e>
      <m:e><m:r><m:t>b</m:t></m:r></m:e>
    </m:mr>
    <m:mr>
      <m:e><m:r><m:t>c</m:t></m:r></m:e>
      <m:e><m:r><m:t>d</m:t></m:r></m:e>
    </m:mr>
  </m:m>
</m:oMath>
"""
        latex, warnings = convert_omml_to_latex(ET.fromstring(xml))
        self.assertEqual(warnings, [])
        self.assertEqual(latex, r"\begin{array}{ll}a & b \\ c & d\end{array}")


if __name__ == "__main__":
    unittest.main()
