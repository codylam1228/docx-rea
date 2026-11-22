"""
Core OOXML inspection, anchor-verified patching, deep XML node cloning,
unified JSON plan execution, and native Word OMML (<m:oMath>) math builder.
Zero external non-stdlib/lxml/python-docx dependencies — runs identically on desktop Python & Pyodide WASM.
"""

from __future__ import annotations

import copy
import re
import shutil
import subprocess
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# Matches either protected placeholders [MATH:0]/[IMG:0] OR inline LaTeX math $...$
SEGMENT_RE = re.compile(r"(\[(?:MATH|IMG):\d+\]|\$\$[^$\n]+?\$\$|\$[^$\n]+?\$)")


@dataclass
class EditInstruction:
    target: str
    anchor_text: str
    new_text: str


@dataclass
class DiffEntry:
    target: str
    before: str
    after: str


@dataclass
class OperationResult:
    output_path: str
    applied_count: int
    diffs: list[DiffEntry] = field(default_factory=list)

    def to_compact_report(self) -> str:
        lines = [
            f"OK | Output: {self.output_path} | Changes applied: {self.applied_count}",
            "--- Compact Diff ---",
        ]
        for d in self.diffs:
            before_esc = d.before.replace("\n", "\\n")
            after_esc = d.after.replace("\n", "\\n")
            lines.append(f"[{d.target}]")
            lines.append(f"  - {before_esc}")
            lines.append(f"  + {after_esc}")
        return "\n".join(lines)


# ==============================================================================
# Native Word OMML (<m:oMath>) Builder from LaTeX
# ==============================================================================

_LATEX_SYMBOLS = {
    r"\times": "×",
    r"\div": "÷",
    r"\pm": "±",
    r"\mp": "∓",
    r"\neq": "≠",
    r"\ne": "≠",
    r"\leq": "≤",
    r"\le": "≤",
    r"\geq": "≥",
    r"\ge": "≥",
    r"\approx": "≈",
    r"\equiv": "≡",
    r"\cdot": "·",
    r"\infty": "∞",
    r"\pi": "π",
    r"\theta": "θ",
    r"\alpha": "α",
    r"\beta": "β",
    r"\gamma": "γ",
    r"\delta": "δ",
    r"\Delta": "Δ",
    r"\lambda": "λ",
    r"\mu": "μ",
    r"\sigma": "σ",
    r"\phi": "φ",
    r"\omega": "ω",
    r"\degree": "°",
    r"\circ": "°",
    r"\rightarrow": "→",
    r"\to": "→",
    r"\Rightarrow": "⇒",
    r"\therefore": "∴",
    r"\because": "∵",
    r"\sin": "sin",
    r"\cos": "cos",
    r"\tan": "tan",
    r"\log": "log",
    r"\ln": "ln",
}


def _m_el(tag: str) -> etree._Element:
    return etree.Element(f"{{{M_NS}}}{tag}", nsmap={"m": M_NS, "w": W_NS})


def _m_run(text: str, italic: bool | None = None) -> etree._Element:
    mr = _m_el("r")
    if italic is False:
        rpr = _m_el("rPr")
        sty = _m_el("sty")
        sty.set(f"{{{M_NS}}}val", "p")
        rpr.append(sty)
        mr.append(rpr)
    mt = _m_el("t")
    mt.set(XML_SPACE, "preserve")
    mt.text = text
    mr.append(mt)
    return mr


def _extract_brace_group(s: str, pos: int) -> tuple[str, int]:
    """Extract {...} starting at pos; if not '{', extract single char or \command."""
    while pos < len(s) and s[pos].isspace():
        pos += 1
    if pos >= len(s):
        return "", pos
    if s[pos] == "{":
        depth = 0
        start = pos + 1
        i = pos
        while i < len(s):
            if s[i] == "{":
                depth += 1
            elif s[i] == "}":
                depth -= 1
                if depth == 0:
                    return s[start:i], i + 1
            i += 1
        return s[start:], len(s)
    if s[pos] == "\\":
        m = re.match(r"\\[a-zA-Z]+", s[pos:])
        if m:
            return m.group(0), pos + len(m.group(0))
    return s[pos], pos + 1


def _append_latex_tokens(container: etree._Element, tex: str) -> None:
    """Recursively parse LaTeX math into OMML child elements inside `container`."""
    i = 0
    n = len(tex)
    while i < n:
        ch = tex[i]
        if ch.isspace():
            i += 1
            continue

        # Handle \frac{num}{den}
        if tex.startswith("\frac", i) or tex.startswith("\dfrac", i) or tex.startswith("\tfrac", i):
            cmd_len = 6 if (tex.startswith("\dfrac", i) or tex.startswith("\tfrac", i)) else 5
            num_str, next_i = _extract_brace_group(tex, i + cmd_len)
            den_str, next_i = _extract_brace_group(tex, next_i)
            mf = _m_el("f")
            num_el = _m_el("num")
            den_el = _m_el("den")
            _append_latex_tokens(num_el, num_str)
            _append_latex_tokens(den_el, den_str)
            mf.append(num_el)
            mf.append(den_el)
            container.append(mf)
            i = next_i
            continue

        # Handle \sqrt[deg]{rad} or \sqrt{rad}
        if tex.startswith("\\sqrt", i):
            pos = i + 5
            while pos < n and tex[pos].isspace():
                pos += 1
            deg_str = None
            if pos < n and tex[pos] == "[":
                end_b = tex.find("]", pos)
                if end_b != -1:
                    deg_str = tex[pos + 1 : end_b]
                    pos = end_b + 1
            rad_str, next_i = _extract_brace_group(tex, pos)
            mrad = _m_el("rad")
            if deg_str is None:
                rad_pr = _m_el("radPr")
                deg_hide = _m_el("degHide")
                deg_hide.set(f"{{{M_NS}}}val", "1")
                rad_pr.append(deg_hide)
                mrad.append(rad_pr)
                mrad.append(_m_el("deg"))
            else:
                deg_el = _m_el("deg")
                _append_latex_tokens(deg_el, deg_str)
                mrad.append(deg_el)
            e_el = _m_el("e")
            _append_latex_tokens(e_el, rad_str)
            mrad.append(e_el)
            container.append(mrad)
            i = next_i
            continue

        # Handle \left and \right sizing commands cleanly
        if tex.startswith("\\left", i):
            i += 5
            continue
        if tex.startswith("\\right", i):
            i += 6
            continue

        # Parse base atom (grouped {...}, \cmd, or single char)
        if ch == "{":
            grp_str, next_i = _extract_brace_group(tex, i)
            base_nodes: list[etree._Element] = []
            tmp = _m_el("e")
            _append_latex_tokens(tmp, grp_str)
            base_nodes.extend(list(tmp))
            i = next_i
        elif ch == "\\":
            m = re.match(r"\\[a-zA-Z]+", tex[i:])
            if m:
                cmd = m.group(0)
                sym = _LATEX_SYMBOLS.get(cmd, cmd.lstrip("\\"))
                is_func = cmd in (r"\sin", r"\cos", r"\tan", r"\log", r"\ln")
                base_nodes = [_m_run(sym, italic=False if is_func else None)]
                i += len(cmd)
            else:
                base_nodes = [_m_run(tex[i + 1 : i + 2])]
                i += 2
        else:
            base_nodes = [_m_run(ch)]
            i += 1

        # Check for trailing superscript ^ or subscript _
        sup_str: str | None = None
        sub_str: str | None = None
        while i < n:
            while i < n and tex[i].isspace():
                i += 1
            if i < n and tex[i] == "^" and sup_str is None:
                sup_str, i = _extract_brace_group(tex, i + 1)
            elif i < n and tex[i] == "_" and sub_str is None:
                sub_str, i = _extract_brace_group(tex, i + 1)
            else:
                break

        if sup_str is not None and sub_str is not None:
            ss = _m_el("sSubSup")
            e_el = _m_el("e")
            for bn in base_nodes:
                e_el.append(bn)
            sub_el = _m_el("sub")
            _append_latex_tokens(sub_el, sub_str)
            sup_el = _m_el("sup")
            _append_latex_tokens(sup_el, sup_str)
            ss.extend([e_el, sub_el, sup_el])
            container.append(ss)
        elif sup_str is not None:
            ss = _m_el("sSup")
            e_el = _m_el("e")
            for bn in base_nodes:
                e_el.append(bn)
            sup_el = _m_el("sup")
            _append_latex_tokens(sup_el, sup_str)
            ss.extend([e_el, sup_el])
            container.append(ss)
        elif sub_str is not None:
            ss = _m_el("sSub")
            e_el = _m_el("e")
            for bn in base_nodes:
                e_el.append(bn)
            sub_el = _m_el("sub")
            _append_latex_tokens(sub_el, sub_str)
            ss.extend([e_el, sub_el])
            container.append(ss)
        else:
            for bn in base_nodes:
                container.append(bn)


def latex_to_omml(latex_expr: str) -> etree._Element:
    """Convert a LaTeX math expression (without outer $) into a native Word <m:oMath> element."""
    omath = _m_el("oMath")
    _append_latex_tokens(omath, latex_expr.strip())
    return omath


# ==============================================================================
# OOXML Inspection & Format-Preserving Patching
# ==============================================================================


def _local_name(el: etree._Element) -> str:
    return etree.QName(el).localname


def _is_protected_node(el: etree._Element) -> str | None:
    """Return 'MATH' or 'IMG' if element is or contains protected non-text XML."""
    lname = _local_name(el)
    if lname in ("oMath", "oMathPara"):
        return "MATH"
    if lname == "r":
        for child in el:
            clname = _local_name(child)
            if clname in ("drawing", "pict", "object"):
                return "IMG"
    return None


def _extract_run_text(r_el: etree._Element) -> str:
    parts: list[str] = []
    for child in r_el:
        clname = _local_name(child)
        if clname == "t":
            parts.append(child.text or "")
        elif clname == "tab":
            parts.append("\t")
        elif clname in ("br", "cr"):
            parts.append("\n")
    return "".join(parts)


def _extract_omath_plain_text(omath_el: etree._Element) -> str:
    """Extract readable characters inside an <m:oMath> element for context hints."""
    chars: list[str] = []
    for sub in omath_el.iter():
        if _local_name(sub) == "t" and sub.text:
            chars.append(sub.text)
    return "".join(chars)


def _extract_paragraph_view(p_el: etree._Element) -> tuple[str, list[etree._Element]]:
    """
    Build a human/AI-readable representation of a <w:p> element while indexing
    embedded protected XML elements ([MATH:0], [IMG:0], ...).
    """
    parts: list[str] = []
    protected: list[etree._Element] = []
    math_idx = 0
    img_idx = 0

    for child in p_el:
        kind = _is_protected_node(child)
        if kind == "MATH":
            parts.append(f"[MATH:{math_idx}]")
            protected.append(child)
            math_idx += 1
        elif kind == "IMG":
            parts.append(f"[IMG:{img_idx}]")
            protected.append(child)
            img_idx += 1
        elif _local_name(child) == "r":
            parts.append(_extract_run_text(child))
        elif _local_name(child) in ("hyperlink", "smartTag", "sdt"):
            for sub in child.iter():
                if _local_name(sub) == "t":
                    parts.append(sub.text or "")

    return "".join(parts), protected


def _summarize_rpr(rpr: etree._Element | None) -> str:
    if rpr is None:
        return ""
    attrs: list[str] = []
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is not None:
        font_name = (
            rfonts.get(qn("w:eastAsia"))
            or rfonts.get(qn("w:ascii"))
            or rfonts.get(qn("w:hAnsi"))
        )
        if font_name:
            attrs.append(font_name)
    sz = rpr.find(qn("w:sz"))
    if sz is not None and sz.get(qn("w:val")):
        try:
            pt = int(sz.get(qn("w:val"))) / 2
            attrs.append(f"{pt:g}pt")
        except ValueError:
            pass
    if rpr.find(qn("w:b")) is not None:
        attrs.append("Bold")
    if rpr.find(qn("w:i")) is not None:
        attrs.append("Italic")
    color = rpr.find(qn("w:color"))
    if color is not None and color.get(qn("w:val")) not in (None, "auto"):
        attrs.append(f"#{color.get(qn('w:val'))}")
    return " ".join(attrs)


def _summarize_paragraph_style(p_el: etree._Element) -> str:
    meta: list[str] = []
    ppr = p_el.find(qn("w:pPr"))
    if ppr is not None:
        pstyle = ppr.find(qn("w:pStyle"))
        if pstyle is not None and pstyle.get(qn("w:val")):
            meta.append(f"Style:{pstyle.get(qn('w:val'))}")
        if ppr.find(qn("w:numPr")) is not None:
            meta.append("List")

    runs = [c for c in p_el if _local_name(c) == "r" and _is_protected_node(c) is None]
    if runs:
        first_rpr = runs[0].find(qn("w:rPr"))
        font_desc = _summarize_rpr(first_rpr)
        if font_desc:
            meta.append(f"Font:{font_desc}")
    if len(runs) > 1:
        sigs = {
            etree.tostring(r.find(qn("w:rPr"))) if r.find(qn("w:rPr")) is not None else b""
            for r in runs
        }
        if len(sigs) > 1:
            meta.append(f"runs:{len(runs)}")
    return f" [{' | '.join(meta)}]" if meta else ""


def _build_run_with_rpr(text: str, template_rpr: etree._Element | None) -> etree._Element:
    r_el = OxmlElement("w:r")
    if template_rpr is not None:
        r_el.append(copy.deepcopy(template_rpr))
    lines = text.split("\n")
    for idx, line in enumerate(lines):
        if idx > 0:
            r_el.append(OxmlElement("w:br"))
        t_el = OxmlElement("w:t")
        t_el.set(XML_SPACE, "preserve")
        t_el.text = line
        r_el.append(t_el)
    return r_el


def _set_run_text_preserve_format(r_el: etree._Element, new_text: str) -> None:
    if _is_protected_node(r_el) is not None:
        raise ValueError("Cannot overwrite a protected image/drawing run directly.")
    for child in list(r_el):
        if _local_name(child) != "rPr":
            r_el.remove(child)
    lines = new_text.split("\n")
    for idx, line in enumerate(lines):
        if idx > 0:
            r_el.append(OxmlElement("w:br"))
        t_el = OxmlElement("w:t")
        t_el.set(XML_SPACE, "preserve")
        t_el.text = line
        r_el.append(t_el)


def _set_paragraph_text_preserve_format(p_el: etree._Element, new_text: str) -> None:
    """
    Replace text inside <w:p> while preserving <w:pPr>, the first <w:rPr>,
    embedded protected nodes ([MATH:k], [IMG:k]), AND converting inline $LaTeX$
    into native Word <m:oMath> equations.
    """
    _, protected_nodes = _extract_paragraph_view(p_el)

    template_rpr: etree._Element | None = None
    for child in p_el:
        if _local_name(child) == "r" and _is_protected_node(child) is None:
            rpr = child.find(qn("w:rPr"))
            if rpr is not None:
                template_rpr = copy.deepcopy(rpr)
            break

    math_nodes = [n for n in protected_nodes if _is_protected_node(n) == "MATH"]
    img_nodes = [n for n in protected_nodes if _is_protected_node(n) == "IMG"]

    removable = {"r", "oMath", "oMathPara", "hyperlink", "smartTag", "sdt"}
    for child in list(p_el):
        if _local_name(child) in removable:
            p_el.remove(child)

    parts = SEGMENT_RE.split(new_text)
    used_protected: set[int] = set()
    has_explicit_math_tokens = bool(re.search(r"\[(?:MATH|IMG):\d+\]", new_text))
    has_latex_math = bool(re.search(r"\$[^$\n]+?\$", new_text))

    for seg in parts:
        if not seg:
            continue
        if seg.startswith("[MATH:") and seg.endswith("]"):
            idx_num = int(seg[6:-1])
            if 0 <= idx_num < len(math_nodes):
                node = math_nodes[idx_num]
                used_protected.add(id(node))
                p_el.append(node)
        elif seg.startswith("[IMG:") and seg.endswith("]"):
            idx_num = int(seg[5:-1])
            if 0 <= idx_num < len(img_nodes):
                node = img_nodes[idx_num]
                used_protected.add(id(node))
                p_el.append(node)
        elif seg.startswith("$$") and seg.endswith("$$") and len(seg) >= 5:
            p_el.append(latex_to_omml(seg[2:-2]))
        elif seg.startswith("$") and seg.endswith("$") and len(seg) >= 3:
            p_el.append(latex_to_omml(seg[1:-1]))
        else:
            p_el.append(_build_run_with_rpr(seg, template_rpr))

    # Preserve existing images always; preserve old math nodes unless replaced by new $LaTeX$
    for node in protected_nodes:
        if id(node) in used_protected:
            continue
        kind = _is_protected_node(node)
        if kind == "IMG":
            p_el.append(node)
        elif kind == "MATH" and not has_latex_math and has_explicit_math_tokens:
            p_el.append(node)


# --- Coordinate Resolution ---


def _get_container_paragraphs(container_el: etree._Element) -> list[etree._Element]:
    return [c for c in container_el if _local_name(c) == "p"]


def _get_container_tables(container_el: etree._Element) -> list[etree._Element]:
    return [c for c in container_el if _local_name(c) == "tbl"]


def _resolve_hf_container(doc: Document, sec_idx: int, part_name: str) -> etree._Element:
    sections = doc.sections
    if sec_idx < 0 or sec_idx >= len(sections):
        raise ValueError(f"Section index hf[{sec_idx}] out of range (total: {len(sections)}).")
    sec = sections[sec_idx]
    mapping = {
        "header": sec.header,
        "footer": sec.footer,
        "first_header": sec.first_page_header,
        "first_footer": sec.first_page_footer,
        "even_header": sec.even_page_header,
        "even_footer": sec.even_page_footer,
    }
    if part_name not in mapping:
        raise ValueError(f"Unknown header/footer part '{part_name}'. Expected one of {list(mapping)}.")
    return mapping[part_name]._element


def _resolve_target(doc: Document, target: str) -> tuple[str, etree._Element]:
    expr = target.strip()
    container: etree._Element = doc.element.body

    hf_match = re.match(r"^hf\[(\d+)\]\.([a-z_]+)\.(.+)$", expr)
    if hf_match:
        sec_idx = int(hf_match.group(1))
        part_name = hf_match.group(2)
        expr = hf_match.group(3)
        container = _resolve_hf_container(doc, sec_idx, part_name)

    p_match = re.fullmatch(r"p\[(\d+)\](?:\.r\[(\d+)\])?", expr)
    if p_match:
        p_idx = int(p_match.group(1))
        paras = _get_container_paragraphs(container)
        if p_idx >= len(paras):
            raise ValueError(f"Target '{target}': paragraph index {p_idx} out of range (count={len(paras)}).")
        p_el = paras[p_idx]
        if p_match.group(2) is None:
            return "p", p_el
        r_idx = int(p_match.group(2))
        runs = [c for c in p_el if _local_name(c) == "r" and _is_protected_node(c) is None]
        if r_idx >= len(runs):
            raise ValueError(f"Target '{target}': run index {r_idx} out of range (count={len(runs)}).")
        return "r", runs[r_idx]

    t_match = re.fullmatch(
        r"t\[(\d+)\]\.r\[(\d+)\](?:\.c\[(\d+)\](?:\.p\[(\d+)\])?(?:\.r\[(\d+)\])?)?",
        expr,
    )
    if t_match:
        t_idx = int(t_match.group(1))
        r_idx = int(t_match.group(2))
        tables = _get_container_tables(container)
        if t_idx >= len(tables):
            raise ValueError(f"Target '{target}': table index {t_idx} out of range (count={len(tables)}).")
        tbl_el = tables[t_idx]
        rows = [c for c in tbl_el if _local_name(c) == "tr"]
        if r_idx >= len(rows):
            raise ValueError(f"Target '{target}': row index {r_idx} out of range (count={len(rows)}).")
        tr_el = rows[r_idx]
        if t_match.group(3) is None:
            return "tr", tr_el
        c_idx = int(t_match.group(3))
        cells = [c for c in tr_el if _local_name(c) == "tc"]
        if c_idx >= len(cells):
            raise ValueError(f"Target '{target}': cell index {c_idx} out of range (count={len(cells)}).")
        tc_el = cells[c_idx]
        cell_paras = _get_container_paragraphs(tc_el)
        p_idx = int(t_match.group(4)) if t_match.group(4) is not None else 0
        if p_idx >= len(cell_paras):
            raise ValueError(f"Target '{target}': cell paragraph index {p_idx} out of range.")
        p_el = cell_paras[p_idx]
        if t_match.group(5) is None:
            return "p", p_el
        run_idx = int(t_match.group(5))
        runs = [c for c in p_el if _local_name(c) == "r" and _is_protected_node(c) is None]
        if run_idx >= len(runs):
            raise ValueError(f"Target '{target}': run index {run_idx} out of range.")
        return "r", runs[run_idx]

    raise ValueError(f"Invalid target coordinate syntax: '{target}'")


def _get_node_text(kind: str, el: etree._Element) -> str:
    if kind == "p":
        txt, _ = _extract_paragraph_view(el)
        return txt
    if kind == "r":
        return _extract_run_text(el)
    if kind == "tr":
        cells = [c for c in el if _local_name(c) == "tc"]
        cell_texts: list[str] = []
        for tc in cells:
            paras = _get_container_paragraphs(tc)
            joined = " / ".join(_extract_paragraph_view(p)[0] for p in paras)
            cell_texts.append(joined)
        return " | ".join(cell_texts)
    raise ValueError(f"Unsupported node kind: {kind}")


# --- Public Engine Functions ---


def inspect_docx(
    path: str | Path,
    include_empty: bool = False,
    include_runs: bool = False,
    target_prefix: str | None = None,
    max_items: int = 300,
) -> str:
    """Return a compact, token-efficient structural map of a .docx file."""
    doc_path = Path(path)
    if not doc_path.exists():
        raise FileNotFoundError(f"File not found: {doc_path}")

    doc = Document(str(doc_path))
    lines: list[str] = [f"FILE: {doc_path.name} (sections={len(doc.sections)})"]

    def emit_paragraph(coord: str, p_el: etree._Element) -> None:
        text, protected = _extract_paragraph_view(p_el)
        if not include_empty and not text.strip():
            return
        if target_prefix and not coord.startswith(target_prefix):
            return
        style_tag = _summarize_paragraph_style(p_el)
        math_hints = [
            f"MATH:{idx}={_extract_omath_plain_text(node)}"
            for idx, node in enumerate([n for n in protected if _is_protected_node(n) == "MATH"])
        ]
        hint_str = f" ({', '.join(math_hints)})" if math_hints else ""
        lines.append(f'{coord}{style_tag} | "{text}"{hint_str}')
        if include_runs:
            runs = [c for c in p_el if _local_name(c) == "r" and _is_protected_node(c) is None]
            for r_i, r_el in enumerate(runs):
                r_txt = _extract_run_text(r_el)
                r_meta = _summarize_rpr(r_el.find(qn("w:rPr")))
                r_tag = f" [{r_meta}]" if r_meta else ""
                lines.append(f'  {coord}.r[{r_i}]{r_tag} | "{r_txt}"')

    def emit_table(base_coord: str, tbl_el: etree._Element) -> None:
        rows = [c for c in tbl_el if _local_name(c) == "tr"]
        if not target_prefix or base_coord.startswith(target_prefix) or target_prefix.startswith(base_coord):
            lines.append(f"{base_coord} (rows={len(rows)})")
        for r_i, tr_el in enumerate(rows):
            cells = [c for c in tr_el if _local_name(c) == "tc"]
            for c_i, tc_el in enumerate(cells):
                cell_paras = _get_container_paragraphs(tc_el)
                for p_i, p_el in enumerate(cell_paras):
                    coord = (
                        f"{base_coord}.r[{r_i}].c[{c_i}]"
                        if p_i == 0
                        else f"{base_coord}.r[{r_i}].c[{c_i}].p[{p_i}]"
                    )
                    emit_paragraph(coord, p_el)

    for s_i, sec in enumerate(doc.sections):
        parts = [("header", sec.header), ("footer", sec.footer)]
        if sec.different_first_page_header_footer:
            parts.extend([("first_header", sec.first_page_header), ("first_footer", sec.first_page_footer)])
        for part_name, hf_obj in parts:
            if hf_obj.is_linked_to_previous and s_i > 0:
                continue
            hf_el = hf_obj._element
            for p_i, p_el in enumerate(_get_container_paragraphs(hf_el)):
                emit_paragraph(f"hf[{s_i}].{part_name}.p[{p_i}]", p_el)
            for t_i, tbl_el in enumerate(_get_container_tables(hf_el)):
                emit_table(f"hf[{s_i}].{part_name}.t[{t_i}]", tbl_el)

    p_counter = 0
    t_counter = 0
    for child in doc.element.body:
        lname = _local_name(child)
        if lname == "p":
            emit_paragraph(f"p[{p_counter}]", child)
            p_counter += 1
        elif lname == "tbl":
            emit_table(f"t[{t_counter}]", child)
            t_counter += 1
        if len(lines) >= max_items:
            lines.append(f"... [Truncated at {max_items} lines; use target_prefix to narrow scope]")
            break

    return "\n".join(lines)


def _resolve_output_path(input_path: Path, output_path: str | Path | None) -> Path:
    if output_path is not None:
        return Path(output_path)
    return input_path.with_name(f"{input_path.stem}_edited{input_path.suffix}")


def batch_edit_docx(
    path: str | Path,
    edits: list[dict[str, str] | EditInstruction],
    output_path: str | Path | None = None,
) -> OperationResult:
    """Apply multiple anchor-verified edits to a .docx file in one atomic pass."""
    in_path = Path(path)
    if not in_path.exists():
        raise FileNotFoundError(f"File not found: {in_path}")

    parsed_edits = [e if isinstance(e, EditInstruction) else EditInstruction(**e) for e in edits]
    if not parsed_edits:
        raise ValueError("No edits provided.")

    doc = Document(str(in_path))

    resolved_targets: list[tuple[EditInstruction, str, etree._Element, str]] = []
    for item in parsed_edits:
        kind, el = _resolve_target(doc, item.target)
        if kind == "tr":
            raise ValueError(
                f"Target '{item.target}' is a table row. Specify a cell (e.g. '{item.target}.c[0]') or use clone_node."
            )
        current_text = _get_node_text(kind, el)
        if item.anchor_text not in current_text:
            raise ValueError(
                f"Anchor mismatch at '{item.target}': expected substring {item.anchor_text!r}, "
                f"but found {current_text!r}. Re-run inspect to verify current indices."
            )
        resolved_targets.append((item, kind, el, current_text))

    diffs: list[DiffEntry] = []
    for item, kind, el, before_text in resolved_targets:
        if kind == "p":
            _set_paragraph_text_preserve_format(el, item.new_text)
        elif kind == "r":
            _set_run_text_preserve_format(el, item.new_text)
        after_text = _get_node_text(kind, el)
        diffs.append(DiffEntry(target=item.target, before=before_text, after=after_text))

    out_path = _resolve_output_path(in_path, output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return OperationResult(output_path=str(out_path), applied_count=len(diffs), diffs=diffs)


def clone_node_docx(
    path: str | Path,
    target: str,
    anchor_text: str,
    items: list[Any],
    mode: str = "after",
    output_path: str | Path | None = None,
) -> OperationResult:
    """
    Deep-clone a table row (e.g. 't[0].r[2]') or styled paragraph (e.g. 'p[8]')
    and populate each clone while preserving 100% of XML geometry & styles.
    """
    if mode not in ("after", "replace"):
        raise ValueError("mode must be 'after' or 'replace'.")
    if not items:
        raise ValueError("items list cannot be empty.")

    in_path = Path(path)
    if not in_path.exists():
        raise FileNotFoundError(f"File not found: {in_path}")

    doc = Document(str(in_path))
    kind, el = _resolve_target(doc, target)
    if kind not in ("tr", "p"):
        raise ValueError(f"Target '{target}' cannot be cloned; must be a table row or paragraph.")

    before_text = _get_node_text(kind, el)
    if anchor_text not in before_text:
        raise ValueError(
            f"Anchor mismatch at '{target}': expected substring {anchor_text!r}, but found {before_text!r}."
        )

    parent = el.getparent()
    if parent is None:
        raise RuntimeError(f"Target '{target}' has no XML parent.")

    insert_idx = parent.index(el)
    if mode == "after":
        insert_idx += 1

    diffs: list[DiffEntry] = []
    for offset, row_or_text in enumerate(items):
        cloned = copy.deepcopy(el)
        if kind == "tr":
            tr_pr = cloned.find(qn("w:trPr"))
            if tr_pr is not None:
                tbl_hdr = tr_pr.find(qn("w:tblHeader"))
                if tbl_hdr is not None:
                    tr_pr.remove(tbl_hdr)
            if not isinstance(row_or_text, (list, tuple)):
                raise ValueError("When cloning a table row, each item in 'items' must be a list of cell strings.")
            cells = [c for c in cloned if _local_name(c) == "tc"]
            for c_i, val in enumerate(row_or_text):
                if c_i >= len(cells):
                    break
                tc_el = cells[c_i]
                cell_paras = _get_container_paragraphs(tc_el)
                if not cell_paras:
                    p_new = OxmlElement("w:p")
                    tc_el.append(p_new)
                    cell_paras = [p_new]
                _set_paragraph_text_preserve_format(cell_paras[0], str(val))
                for extra_p in cell_paras[1:]:
                    tc_el.remove(extra_p)
        else:
            if not isinstance(row_or_text, str):
                raise ValueError("When cloning a paragraph, each item in 'items' must be a string.")
            _set_paragraph_text_preserve_format(cloned, row_or_text)

        parent.insert(insert_idx + offset, cloned)
        diffs.append(
            DiffEntry(
                target=f"{target} -> clone[{offset}] ({mode})",
                before=before_text,
                after=_get_node_text(kind, cloned),
            )
        )

    if mode == "replace":
        parent.remove(el)

    out_path = _resolve_output_path(in_path, output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return OperationResult(output_path=str(out_path), applied_count=len(diffs), diffs=diffs)


def delete_nodes_docx(
    path: str | Path,
    deletions: list[dict[str, str]],
    output_path: str | Path | None = None,
) -> OperationResult:
    """
    Safely delete table rows (e.g. 't[0].r[15]') or paragraphs (e.g. 'p[9]')
    with anchor verification. Resolves all targets before removing so index shifts never corrupt.
    """
    in_path = Path(path)
    if not in_path.exists():
        raise FileNotFoundError(f"File not found: {in_path}")
    if not deletions:
        raise ValueError("No deletions provided.")

    doc = Document(str(in_path))
    resolved: list[tuple[str, etree._Element, str]] = []
    for d in deletions:
        target = d["target"]
        anchor_text = d.get("anchor_text", "")
        kind, el = _resolve_target(doc, target)
        if kind not in ("tr", "p"):
            raise ValueError(f"Target '{target}' cannot be deleted; must be a table row or paragraph.")
        before_text = _get_node_text(kind, el)
        if anchor_text and anchor_text not in before_text:
            raise ValueError(
                f"Anchor mismatch at '{target}': expected substring {anchor_text!r}, but found {before_text!r}."
            )
        resolved.append((target, el, before_text))

    diffs: list[DiffEntry] = []
    for target, el, before_text in resolved:
        parent = el.getparent()
        if parent is not None and el in parent:
            parent.remove(el)
            diffs.append(DiffEntry(target=f"{target} (deleted)", before=before_text, after="<DELETED>"))

    out_path = _resolve_output_path(in_path, output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return OperationResult(output_path=str(out_path), applied_count=len(diffs), diffs=diffs)


def apply_plan_docx(
    path: str | Path,
    plan: dict[str, Any],
    output_path: str | Path | None = None,
) -> OperationResult:
    """
    Execute a unified declarative JSON plan generated by an external LLM (Gemini / Grok web):
    {
      "delete_targets": [{"target": "t[0].r[15]", "anchor_text": "..."}],
      "clones": [{"target": "t[0].r[0]", "anchor_text": "...", "mode": "replace", "items": [...]}],
      "edits":  [{"target": "hf[0].header.p[0]", "anchor_text": "...", "new_text": "..."}]
    }
    """
    in_path = Path(path)
    if not in_path.exists():
        raise FileNotFoundError(f"File not found: {in_path}")

    deletions = plan.get("delete_targets") or []
    clones = plan.get("clones") or []
    edits = plan.get("edits") or []
    if not deletions and not clones and not edits:
        raise ValueError("JSON plan must contain at least one item in 'delete_targets', 'clones', or 'edits'.")

    out_path = _resolve_output_path(in_path, output_path)
    working_path = in_path
    all_diffs: list[DiffEntry] = []

    if deletions:
        res = delete_nodes_docx(
            path=working_path,
            deletions=deletions,
            output_path=out_path,
        )
        all_diffs.extend(res.diffs)
        working_path = out_path

    for c_spec in clones:
        res = clone_node_docx(
            path=working_path,
            target=c_spec["target"],
            anchor_text=c_spec["anchor_text"],
            items=c_spec["items"],
            mode=c_spec.get("mode", "after"),
            output_path=out_path,
        )
        all_diffs.extend(res.diffs)
        working_path = out_path

    if edits:
        res = batch_edit_docx(
            path=working_path,
            edits=edits,
            output_path=out_path,
        )
        all_diffs.extend(res.diffs)

    return OperationResult(output_path=str(out_path), applied_count=len(all_diffs), diffs=all_diffs)


def render_preview_images(path: str | Path, out_dir: str | Path | None = None) -> list[str]:
    """Convert .docx to PDF via LibreOffice headless and render PNG page images if pdftoppm is installed."""
    doc_path = Path(path).resolve()
    if not doc_path.exists():
        raise FileNotFoundError(f"File not found: {doc_path}")

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("LibreOffice ('soffice') not found in PATH.")

    target_dir = Path(out_dir).resolve() if out_dir else doc_path.parent / f".{doc_path.stem}_preview"
    target_dir.mkdir(parents=True, exist_ok=True)

    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(target_dir), str(doc_path)],
        check=True,
        capture_output=True,
        text=True,
    )
    pdf_path = target_dir / f"{doc_path.stem}.pdf"
    if not pdf_path.exists():
        raise RuntimeError(f"LibreOffice failed to generate PDF at {pdf_path}")

    pdftoppm = shutil.which("pdftoppm")
    if pdftoppm:
        prefix = target_dir / "page"
        subprocess.run(
            [pdftoppm, "-png", "-r", "150", str(pdf_path), str(prefix)],
            check=True,
            capture_output=True,
            text=True,
        )
        pngs = sorted(str(p) for p in target_dir.glob("page-*.png"))
        if pngs:
            return pngs

    return [str(pdf_path)]
