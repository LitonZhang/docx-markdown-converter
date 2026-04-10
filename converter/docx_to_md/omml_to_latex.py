from __future__ import annotations

from dataclasses import dataclass, field
import unicodedata
from xml.etree import ElementTree as ET

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _local_name(tag: str) -> str:
    return tag.split("}", 1)[-1]


def _escape_latex_text(text: str) -> str:
    replacements = {
        "\\": r"\\",
        "{": r"\{",
        "}": r"\}",
        "#": r"\#",
        "%": r"\%",
        "&": r"\&",
        "_": r"\_",
    }
    greek_map = {
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
    accent_map = {
        "\u0302": "hat",
        "\u0303": "tilde",
        "\u0304": "bar",
        "\u0307": "dot",
        "\u0308": "ddot",
    }

    out: list[str] = []
    for ch in text:
        if ch in greek_map:
            out.append(greek_map[ch])
            continue

        decomp = unicodedata.normalize("NFD", ch)
        if len(decomp) >= 2 and decomp[0].isalpha():
            combining = [mark for mark in decomp[1:] if unicodedata.combining(mark)]
            if len(combining) == 1 and combining[0] in accent_map:
                out.append(rf"\{accent_map[combining[0]]}{{{decomp[0]}}}")
                continue

        out.append(replacements.get(ch, ch))

    return "".join(out)


def _latex_delimiter(symbol: str, is_left: bool) -> str:
    mapping = {
        "(": "(",
        ")": ")",
        "[": "[",
        "]": "]",
        "{": r"\{",
        "}": r"\}",
        "|": r"\|",
        "‖": r"\|",
        "∥": r"\|",
        "⟨": r"\langle" if is_left else r"\rangle",
        "⟩": r"\langle" if is_left else r"\rangle",
    }
    return mapping.get(symbol, symbol)


@dataclass
class OmmlConversionState:
    warnings: list[str] = field(default_factory=list)
    warned_tags: set[str] = field(default_factory=set)

    def warn_once(self, tag: str) -> None:
        if tag in self.warned_tags:
            return
        self.warned_tags.add(tag)
        self.warnings.append(f"Unsupported OMML element: {tag}")


class OmmlToLatexConverter:
    def __init__(self) -> None:
        self.state = OmmlConversionState()

    def convert(self, element: ET.Element) -> tuple[str, list[str]]:
        latex = self._convert_node(element).strip()
        return latex, list(self.state.warnings)

    def _convert_children(self, element: ET.Element) -> str:
        return "".join(self._convert_node(child) for child in list(element))

    def _find(self, element: ET.Element, name: str) -> ET.Element | None:
        return element.find(f"{{{M_NS}}}{name}")

    def _node_text(self, element: ET.Element | None) -> str:
        if element is None:
            return ""
        return self._convert_node(element)

    def _convert_node(self, element: ET.Element) -> str:
        name = _local_name(element.tag)

        if name in {"oMath", "oMathPara"}:
            return self._convert_children(element)

        if name in {"r", "num", "den", "e", "sub", "sup", "deg", "fName", "lim", "dPr"}:
            return self._convert_children(element)

        if name == "t":
            return _escape_latex_text(element.text or "")

        if name == "f":
            numerator = self._node_text(self._find(element, "num"))
            denominator = self._node_text(self._find(element, "den"))
            return rf"\frac{{{numerator}}}{{{denominator}}}"

        if name == "sSup":
            base = self._node_text(self._find(element, "e"))
            sup = self._node_text(self._find(element, "sup"))
            return rf"{base}^{{{sup}}}"

        if name == "sSub":
            base = self._node_text(self._find(element, "e"))
            sub = self._node_text(self._find(element, "sub"))
            return rf"{base}_{{{sub}}}"

        if name == "sSubSup":
            base = self._node_text(self._find(element, "e"))
            sub = self._node_text(self._find(element, "sub"))
            sup = self._node_text(self._find(element, "sup"))
            return rf"{base}_{{{sub}}}^{{{sup}}}"

        if name == "rad":
            degree = self._node_text(self._find(element, "deg"))
            expr = self._node_text(self._find(element, "e"))
            if degree:
                return rf"\sqrt[{degree}]{{{expr}}}"
            return rf"\sqrt{{{expr}}}"

        if name == "d":
            expr = self._node_text(self._find(element, "e"))
            begin = "("
            end = ")"
            dpr = self._find(element, "dPr")
            if dpr is not None:
                beg_chr = self._find(dpr, "begChr")
                end_chr = self._find(dpr, "endChr")
                begin = beg_chr.attrib.get(f"{{{M_NS}}}val", begin) if beg_chr is not None else begin
                end = end_chr.attrib.get(f"{{{M_NS}}}val", end) if end_chr is not None else end
            left_delim = _latex_delimiter(begin, is_left=True)
            right_delim = _latex_delimiter(end, is_left=False)
            return rf"\left{left_delim}{expr}\right{right_delim}"

        if name == "nary":
            nary_pr = self._find(element, "naryPr")
            char = "\u2211"
            if nary_pr is not None:
                chr_node = self._find(nary_pr, "chr")
                if chr_node is not None:
                    char = chr_node.attrib.get(f"{{{M_NS}}}val", char)
            symbol_map = {
                "\u2211": "\\sum",
                "\u222b": "\\int",
                "\u220f": "\\prod",
                "\u22c2": "\\bigcap",
                "\u22c3": "\\bigcup",
            }
            operator = symbol_map.get(char, "\\sum")
            sub = self._node_text(self._find(element, "sub"))
            sup = self._node_text(self._find(element, "sup"))
            expr = self._node_text(self._find(element, "e"))
            if sub and sup:
                return rf"{operator}_{{{sub}}}^{{{sup}}}{expr}"
            if sub:
                return rf"{operator}_{{{sub}}}{expr}"
            if sup:
                return rf"{operator}^{{{sup}}}{expr}"
            return rf"{operator}{expr}"

        if name == "func":
            fname = self._node_text(self._find(element, "fName"))
            expr = self._node_text(self._find(element, "e"))
            return rf"{fname}\left({expr}\right)"

        if name == "limLow":
            base = self._node_text(self._find(element, "e"))
            lower = self._node_text(self._find(element, "lim"))
            return rf"{base}_{{{lower}}}"

        if name == "limUpp":
            base = self._node_text(self._find(element, "e"))
            upper = self._node_text(self._find(element, "lim"))
            return rf"{base}^{{{upper}}}"

        if name == "acc":
            acc_pr = self._find(element, "accPr")
            accent = "hat"
            if acc_pr is not None:
                chr_node = self._find(acc_pr, "chr")
                if chr_node is not None:
                    val = chr_node.attrib.get(f"{{{M_NS}}}val", "")
                    accent_map = {
                        "\u0302": "hat",
                        "\u0303": "tilde",
                        "\u0304": "bar",
                        "\u0307": "dot",
                    }
                    accent = accent_map.get(val, "hat")
            expr = self._node_text(self._find(element, "e"))
            return rf"\{accent}{{{expr}}}"

        if name in {"accPr", "naryPr", "begChr", "endChr", "chr", "ctrlPr", "rPr"}:
            return ""

        if list(element):
            self.state.warn_once(name)
            return self._convert_children(element)

        if element.text:
            return _escape_latex_text(element.text)

        self.state.warn_once(name)
        return ""


def convert_omml_to_latex(element: ET.Element) -> tuple[str, list[str]]:
    converter = OmmlToLatexConverter()
    return converter.convert(element)
