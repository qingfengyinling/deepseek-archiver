#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepSeek 对话归档 → Word (.docx)
================================

两种用法：

  1) 直接抓取（需要登录后的 token，见 README 里的取法）
       python deepseek_archive.py fetch --token <TOKEN>
       python deepseek_archive.py fetch --token <TOKEN> --merge --think --limit 50

  2) 把已导出的原始数据（.json 或整包 .zip）重新生成 Word
       python deepseek_archive.py convert 原始数据.json
       python deepseek_archive.py convert DeepSeek归档_20261022_0855.zip --merge

输出结构：
  输出目录/
    会话/001_标题.docx
    全部会话合集.docx      （加 --merge 时）
    原始数据.json
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
import zipfile
from datetime import datetime
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor, Twips

VERSION = "1.0.0"
BASE_URL = "https://chat.deepseek.com"
API = "/api/v0"

BODY_FONT_CN = "微软雅黑"
BODY_FONT_EN = "Calibri"
MONO_FONT = "Consolas"
MONO_FONT_CN = "等线"

TEXT = "1F2328"
GRAY = "6E7781"
LIGHT_GRAY = "9AA4AF"
BLUE = "0B5CAD"
USER_BLUE = "1A56B0"
AMBER = "9A6700"


# ===========================================================================
# 一、通用工具
# ===========================================================================

def fmt_time(sec) -> str:
    if sec in (None, "", 0, "0"):
        return ""
    try:
        v = float(sec)
    except (TypeError, ValueError):
        return ""
    if v <= 0:
        return ""
    if v < 1e11:
        v *= 1000
    try:
        return datetime.fromtimestamp(v / 1000).strftime("%Y-%m-%d %H:%M")
    except (OverflowError, OSError, ValueError):
        return ""


def fmt_day(sec) -> str:
    t = fmt_time(sec)
    return t[:10] if t else "未知日期"


def safe_name(s: str, limit: int = 60) -> str:
    t = re.sub(r"[\x00-\x1f]", " ", str(s or ""))
    t = re.sub(r'[\\/:*?"<>|]', "_", t)
    t = re.sub(r"\s+", " ", t).strip().strip(".")
    if len(t) > limit:
        t = t[:limit].strip()
    if not t:
        t = "未命名会话"
    if re.fullmatch(r"(CON|PRN|AUX|NUL|COM\d|LPT\d)", t, re.I):
        t = "_" + t
    return t


CJK_RE = re.compile(r"[\u2e80-\u9fff\u3000-\u303f\uff00-\uffef]")


def join_soft_lines(lines) -> str:
    out = ""
    for raw in lines:
        ln = raw.rstrip()
        if not out:
            out = ln
            continue
        a, b = out[-1:], ln[:1]
        out += ln if (CJK_RE.match(a) or CJK_RE.match(b)) else " " + ln
    return out


# ===========================================================================
# 二、Markdown → 块结构
# ===========================================================================

RE_HEADING = re.compile(r"^(#{1,6})\s+(.+?)\s*#*\s*$")
RE_FENCE = re.compile(r"^\s{0,3}(```|~~~)\s*([\w+#.\-]*)\s*$")
RE_HR = re.compile(r"^\s{0,3}([-*_])(?:\s*\1){2,}\s*$")
RE_LIST = re.compile(r"^(\s*)([-*+]|\d{1,3}[.)])\s+(.*)$")
RE_QUOTE = re.compile(r"^\s{0,3}>\s?")
RE_TABLE_SEP = re.compile(r"^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$")


def split_table_row(line: str):
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    cells, cur, i = [], "", 0
    while i < len(s):
        if s[i] == "\\" and i + 1 < len(s) and s[i + 1] == "|":
            cur += "|"
            i += 2
            continue
        if s[i] == "|":
            cells.append(cur.strip())
            cur = ""
            i += 1
            continue
        cur += s[i]
        i += 1
    cells.append(cur.strip())
    return cells


def parse_blocks(md: str):
    lines = str(md or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    blocks, buf, i = [], [], 0
    indents: list[int] = []

    def flush():
        if buf:
            blocks.append({"type": "p", "text": join_soft_lines(buf)})
            buf.clear()

    while i < len(lines):
        line = lines[i]

        m = RE_FENCE.match(line)
        if m:
            flush()
            fence = m.group(1)
            lang = m.group(2) or ""
            code, i = [], i + 1
            close = re.compile(r"^\s{0,3}" + re.escape(fence))
            while i < len(lines) and not close.match(lines[i]):
                code.append(lines[i])
                i += 1
            i += 1
            blocks.append({"type": "code", "lang": lang, "text": "\n".join(code)})
            continue

        if not line.strip():
            flush()
            indents.clear()
            i += 1
            continue

        m = RE_HEADING.match(line)
        if m:
            flush()
            indents.clear()
            blocks.append({"type": "h", "level": len(m.group(1)), "text": m.group(2).strip()})
            i += 1
            continue

        if RE_HR.match(line):
            flush()
            indents.clear()
            blocks.append({"type": "hr"})
            i += 1
            continue

        if "|" in line and i + 1 < len(lines) and "-" in lines[i + 1] and RE_TABLE_SEP.match(lines[i + 1]):
            flush()
            indents.clear()
            header = split_table_row(line)
            i += 2
            rows = []
            while i < len(lines) and "|" in lines[i] and lines[i].strip():
                rows.append(split_table_row(lines[i]))
                i += 1
            blocks.append({"type": "table", "header": header, "rows": rows})
            continue

        if RE_QUOTE.match(line):
            flush()
            indents.clear()
            q = []
            while i < len(lines) and (RE_QUOTE.match(lines[i]) or (
                    lines[i].strip() and q and not RE_HEADING.match(lines[i]) and not RE_FENCE.match(lines[i]))):
                q.append(RE_QUOTE.sub("", lines[i]))
                i += 1
            blocks.append({"type": "quote", "text": join_soft_lines(q).strip()})
            continue

        m = RE_LIST.match(line)
        if m:
            flush()
            indent = len(m.group(1).replace("\t", "  "))
            while indents and indent < indents[-1]:
                indents.pop()
            if not indents or indent > indents[-1]:
                indents.append(indent)
            level = min(4, max(0, len(indents) - 1))
            marker, text = m.group(2), m.group(3)
            ordered = bool(re.search(r"\d", marker))
            task = re.match(r"^\[([ xX])\]\s+", text)
            if task:
                text = re.sub(r"^\[[ xX]\]\s+", "", text)
            i += 1
            cont = []
            while i < len(lines) and lines[i].strip() and re.match(r"^\s{2,}\S", lines[i]) and not RE_LIST.match(lines[i]):
                cont.append(lines[i].strip())
                i += 1
            blocks.append({
                "type": "li", "ordered": ordered, "level": level, "marker": marker,
                "task": (task.group(1).lower() == "x") if task else None,
                "text": join_soft_lines([text] + cont),
            })
            continue

        if re.match(r"^\s*\$\$", line):
            flush()
            math, closed, first = [], False, re.sub(r"^\s*\$\$", "", line)
            if re.search(r"\$\$\s*$", first):
                math.append(re.sub(r"\$\$\s*$", "", first))
                i += 1
                closed = True
            if not closed:
                if first.strip():
                    math.append(first)
                i += 1
                while i < len(lines):
                    if re.search(r"\$\$\s*$", lines[i]):
                        math.append(re.sub(r"\$\$\s*$", "", lines[i]))
                        i += 1
                        closed = True
                        break
                    math.append(lines[i])
                    i += 1
            blocks.append({"type": "math", "text": "\n".join(math).strip()})
            continue

        buf.append(re.sub(r"^\s+", "", line))
        i += 1

    flush()
    return blocks


TEX_SYMBOLS = {
    "times": "×", "div": "÷", "cdot": "·", "pm": "±", "mp": "∓",
    "leq": "≤", "le": "≤", "geq": "≥", "ge": "≥", "neq": "≠", "ne": "≠",
    "approx": "≈", "equiv": "≡", "infty": "∞", "to": "→", "rightarrow": "→",
    "Rightarrow": "⇒", "leftarrow": "←", "sum": "Σ", "prod": "∏", "int": "∫",
    "partial": "∂", "nabla": "∇", "in": "∈", "notin": "∉", "subset": "⊂",
    "cup": "∪", "cap": "∩", "forall": "∀", "exists": "∃", "emptyset": "∅",
    "sqrt": "√", "angle": "∠", "degree": "°",
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "theta": "θ", "lambda": "λ", "mu": "μ", "pi": "π", "rho": "ρ", "sigma": "σ",
    "tau": "τ", "phi": "φ", "omega": "ω", "Delta": "Δ", "Sigma": "Σ",
    "Omega": "Ω", "Gamma": "Γ", "Theta": "Θ", "Lambda": "Λ",
    "log": "log", "ln": "ln", "exp": "exp", "sin": "sin", "cos": "cos",
    "tan": "tan", "max": "max", "min": "min", "lim": "lim",
}
TEX_FUNCS = {"log", "ln", "exp", "sin", "cos", "tan", "max", "min", "lim"}


def de_tex(s: str) -> str:
    t = str(s)
    t = re.sub(r"\\frac\s*\{([^{}]*)\}\s*\{([^{}]*)\}", r"(\1)/(\2)", t)
    t = re.sub(r"\\sqrt\s*\{([^{}]*)\}", r"√(\1)", t)
    t = re.sub(r"\\left\s*|\\right\s*", "", t)
    t = re.sub(r"\\[,;:!]", " ", t)
    t = re.sub(r"\\quad|\\qquad", "  ", t)

    def repl(m):
        name = m.group(1)
        sym = TEX_SYMBOLS.get(name, name)
        if name in TEX_FUNCS and m.start() > 0 and re.match(r"[A-Za-z0-9)\]]", t[m.start() - 1]):
            return " " + sym
        return sym

    t = re.sub(r"\\([A-Za-z]+)", repl, t)
    t = re.sub(r"\\([{}%_])", r"\1", t)
    t = re.sub(r"[{}]", "", t)
    return re.sub(r"\s+", " ", t).strip()


def parse_inline(text: str, st=None):
    """返回 [{text, b, i, strike, code, link, math}]"""
    st = dict(st or {})
    out, buf, i = [], "", 0

    def flush():
        nonlocal buf
        if buf:
            d = dict(st)
            d["text"] = buf
            out.append(d)
            buf = ""

    def push(t, extra=None):
        d = dict(st)
        if extra:
            d.update(extra)
        d["text"] = t
        out.append(d)

    while i < len(text):
        rest = text[i:]
        ch = text[i]

        if ch == "\\" and i + 1 < len(text) and text[i + 1] in "\\`*_{}[]()#+-.!~>|$":
            buf += text[i + 1]
            i += 2
            continue

        if ch == "$":
            mm = re.match(r"^\$\$?([^$]+?)\$\$?", rest)
            if mm and mm.group(1).strip():
                flush()
                push(mm.group(1), {"code": True, "math": True})
                i += len(mm.group(0))
                continue

        lm = re.match(r"^\[([^\]\n]*)\]\(\s*<?([^)>\s]+)>?(?:\s+[\"'][^\"']*[\"'])?\s*\)", rest)
        if lm:
            flush()
            inner = parse_inline(lm.group(1), {**st, "link": lm.group(2)})
            out.extend(inner)
            i += len(lm.group(0))
            continue

        am = re.match(r"^https?://[^\s<>\"'）】\]]+", rest)
        if am and not st.get("link"):
            url = am.group(0).rstrip(".,;:")
            flush()
            push(url, {"link": url})
            i += len(url)
            continue

        if rest.startswith("**") or rest.startswith("__"):
            mk = rest[:2]
            end = text.find(mk, i + 2)
            if end > i + 2:
                flush()
                out.extend(parse_inline(text[i + 2:end], {**st, "b": True}))
                i = end + 2
                continue

        if rest.startswith("~~"):
            end = text.find("~~", i + 2)
            if end > i + 2:
                flush()
                out.extend(parse_inline(text[i + 2:end], {**st, "strike": True}))
                i = end + 2
                continue

        if ch == "`":
            mm = re.match(r"^(`+)([\s\S]*?[^`])\1(?!`)", rest)
            if mm:
                flush()
                push((mm.group(2).strip() or mm.group(2)), {"code": True})
                i += len(mm.group(0))
                continue

        if ch in "*_":
            if ch == "*" and i + 1 < len(text) and text[i + 1] == "*":
                buf += ch
                i += 1
                continue
            end = text.find(ch, i + 1)
            if end > i + 1 and end - i < 400:
                inner = text[i + 1:end]
                if inner.strip() and not inner[:1].isspace() and not inner[-1:].isspace():
                    flush()
                    out.extend(parse_inline(inner, {**st, "i": True}))
                    i = end + 1
                    continue

        buf += ch
        i += 1

    flush()
    return out


# ===========================================================================
# 三、python-docx 样式与底层工具
# ===========================================================================

STYLE_IDS = {
    "meta": "Meta", "caption": "Caption", "quote": "DS Quote", "think": "DS Think",
    "code": "DS Code Block", "math": "DS Math Block", "table": "DS Table Text",
    "toc": "DS TOC Entry", "role_user": "DS Role User", "role_ai": "DS Role AI",
}


def _set_east_asia(style_or_run, name: str):
    rpr = style_or_run.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    # 主题字体优先于显式字体名，必须先去掉主题属性
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        if rfonts.get(qn("w:" + attr)) is not None:
            del rfonts.attrib[qn("w:" + attr)]
    rfonts.set(qn("w:eastAsia"), name)
    return rfonts


PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "numPr", "pBdr", "shd", "tabs",
    "spacing", "ind", "jc", "outlineLvl", "rPr", "sectPr",
]


def _pPr_insert(pPr, el, tag: str):
    """按 OOXML schema 顺序把元素插入 pPr，避免 Word 报“内容有问题”。"""
    idx = PPR_ORDER.index(tag)
    for child in list(pPr):
        name = child.tag.split("}")[-1]
        if name in PPR_ORDER and PPR_ORDER.index(name) > idx:
            child.addprevious(el)
            return
    pPr.append(el)


def _add_bottom_border(paragraph, color="D8DEE4", sz=6, space=6):
    pPr = paragraph._p.get_or_add_pPr()
    _pPr_insert(pPr, parse_xml(
        f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="{sz}" w:space="{space}" w:color="{color}"/></w:pBdr>'),
        "pBdr")


def _add_style(doc, name, base="Normal", ppr_xml=None, rpr_xml=None, char=False):
    styles = doc.styles
    try:
        st = styles[name]
    except KeyError:
        st = styles.add_style(name, WD_STYLE_TYPE.CHARACTER if char else WD_STYLE_TYPE.PARAGRAPH)
    if base and not char:
        try:
            st.base_style = styles[base]
        except KeyError:
            pass
    if ppr_xml:
        el = st.element
        old = el.find(qn("w:pPr"))
        if old is not None:
            el.remove(old)
        el.append(parse_xml(f"<w:pPr {nsdecls('w')}>{ppr_xml}</w:pPr>"))
    if rpr_xml:
        el = st.element
        old_rpr = el.find(qn("w:rPr"))
        if old_rpr is not None:
            el.remove(old_rpr)
        el.append(parse_xml(f"<w:rPr {nsdecls('w')}>{rpr_xml}</w:rPr>"))
    # rPr 必须是 style 的最后一个子元素
    tail = st.element.find(qn("w:rPr"))
    if tail is not None:
        st.element.remove(tail)
        st.element.append(tail)
    return st


def setup_styles(doc):
    normal = doc.styles["Normal"]
    normal.font.name = BODY_FONT_EN
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor.from_string(TEXT)
    _set_east_asia(normal, BODY_FONT_CN)
    pf = normal.paragraph_format
    pf.line_spacing = 1.35
    pf.space_before = Pt(0)
    pf.space_after = Pt(6)

    def heading(name, size, before, after, outline, color=TEXT):
        st = doc.styles[name]
        st.font.name = BODY_FONT_EN
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = RGBColor.from_string(color)
        _set_east_asia(st, BODY_FONT_CN)
        st.paragraph_format.space_before = Pt(before)
        st.paragraph_format.space_after = Pt(after)
        st.paragraph_format.line_spacing = 1.25
        st.paragraph_format.keep_with_next = True

    try:
        heading("Title", 20, 12, 6, 0)
    except KeyError:
        doc.styles.add_style("Title", WD_STYLE_TYPE.PARAGRAPH)
    # 默认模板的 Title 带一条蓝色下边框，去掉它
    _add_style(doc, "Title", base=None,
               ppr_xml='<w:spacing w:before="0" w:after="6" w:line="280" w:lineRule="auto"/><w:outlineLvl w:val="0"/>')
    heading("Heading 1", 16, 16, 7, 0)
    heading("Heading 2", 13.5, 14, 5, 1)
    heading("Heading 3", 12, 11, 4, 2)
    heading("Heading 4", 11, 9, 3, 3, "30363D")

    _add_style(doc, STYLE_IDS["meta"], rpr_xml=f'<w:color w:val="{GRAY}"/><w:sz w:val="17"/><w:szCs w:val="17"/>',
               ppr_xml='<w:spacing w:before="0" w:after="2" w:line="260" w:lineRule="auto"/>')
    _add_style(doc, STYLE_IDS["caption"], rpr_xml=f'<w:color w:val="{LIGHT_GRAY}"/><w:sz w:val="16"/><w:szCs w:val="16"/>',
               ppr_xml='<w:spacing w:before="0" w:after="3" w:line="240" w:lineRule="auto"/>')

    _add_style(doc, STYLE_IDS["quote"],
               ppr_xml=('<w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="D0D7DE"/></w:pBdr>'
                        '<w:shd w:val="clear" w:color="auto" w:fill="F7F8FA"/>'
                        '<w:spacing w:before="4" w:after="7" w:line="300" w:lineRule="auto"/>'
                        '<w:ind w:left="360" w:right="180"/>'),
               rpr_xml='<w:color w:val="4B5563"/><w:sz w:val="20"/>')

    _add_style(doc, STYLE_IDS["think"],
               ppr_xml=('<w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="E3B341"/></w:pBdr>'
                        '<w:shd w:val="clear" w:color="auto" w:fill="FCFAF4"/>'
                        '<w:spacing w:before="5" w:after="8" w:line="280" w:lineRule="auto"/>'
                        '<w:ind w:left="300" w:right="180"/>'),
               rpr_xml=f'<w:color w:val="{GRAY}"/><w:sz w:val="19"/>')

    _add_style(doc, STYLE_IDS["code"],
               ppr_xml=('<w:pBdr>'
                        '<w:top w:val="single" w:sz="4" w:space="6" w:color="E2E6EA"/>'
                        '<w:left w:val="single" w:sz="4" w:space="6" w:color="E2E6EA"/>'
                        '<w:bottom w:val="single" w:sz="4" w:space="6" w:color="E2E6EA"/>'
                        '<w:right w:val="single" w:sz="4" w:space="6" w:color="E2E6EA"/>'
                        '</w:pBdr>'
                        '<w:shd w:val="clear" w:color="auto" w:fill="F6F8FA"/>'
                        '<w:spacing w:before="4" w:after="9" w:line="250" w:lineRule="auto"/>'
                        '<w:ind w:left="120" w:right="120"/>'),
               rpr_xml=(f'<w:rFonts w:ascii="{MONO_FONT}" w:hAnsi="{MONO_FONT}" '
                        f'w:eastAsia="{MONO_FONT_CN}" w:cs="{MONO_FONT}"/>'
                        '<w:color w:val="24292F"/><w:sz w:val="18"/><w:szCs w:val="18"/>'))

    _add_style(doc, STYLE_IDS["math"],
               ppr_xml=('<w:shd w:val="clear" w:color="auto" w:fill="F8F8FA"/>'
                        '<w:spacing w:before="4" w:after="8" w:line="280" w:lineRule="auto"/>'
                        '<w:ind w:left="240"/>'),
               rpr_xml=f'<w:i/><w:color w:val="{TEXT}"/><w:sz w:val="19"/>')

    _add_style(doc, STYLE_IDS["table"], ppr_xml='<w:spacing w:before="0" w:after="0" w:line="260" w:lineRule="auto"/>',
               rpr_xml='<w:sz w:val="19"/><w:szCs w:val="19"/>')

    _add_style(doc, STYLE_IDS["toc"],
               ppr_xml=('<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9026"/></w:tabs>'
                        '<w:spacing w:before="0" w:after="4" w:line="280" w:lineRule="auto"/>'),
               rpr_xml='<w:sz w:val="20"/>')

    _add_style(doc, STYLE_IDS["role_user"],
               ppr_xml=('<w:keepNext/>'
                        '<w:shd w:val="clear" w:color="auto" w:fill="EAF2FE"/>'
                        '<w:spacing w:before="15" w:after="7" w:line="280" w:lineRule="auto"/>'),
               rpr_xml='')
    _add_style(doc, STYLE_IDS["role_ai"],
               ppr_xml=('<w:keepNext/>'
                        '<w:shd w:val="clear" w:color="auto" w:fill="F1F2F4"/>'
                        '<w:spacing w:before="15" w:after="7" w:line="280" w:lineRule="auto"/>'),
               rpr_xml='')


def setup_page(doc):
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.3)
    sec.bottom_margin = Cm(2.2)
    sec.left_margin = Cm(2.2)
    sec.right_margin = Cm(2.2)
    sec.header_distance = Cm(1.5)
    sec.footer_distance = Cm(1.4)


def _shade_run(run, fill: str):
    rpr = run._element.get_or_add_rPr()
    rpr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill}"/>'))


def _shade_cell(cell, fill: str):
    tcpr = cell._tc.get_or_add_tcPr()
    tcpr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill}"/>'))


def add_hyperlink(paragraph, url: str, text: str, size=None, color=BLUE):
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                          is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    hyperlink.set(qn("w:history"), "1")
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    c = OxmlElement("w:color")
    c.set(qn("w:val"), color)
    rpr.append(c)
    if size:
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(size * 2)))
        rpr.append(sz)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rpr.append(u)
    run.append(rpr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    run.append(t)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)
    return hyperlink


def add_anchor_link(paragraph, anchor: str, text: str, size=None, color=BLUE):
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), anchor)
    hyperlink.set(qn("w:history"), "1")
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    c = OxmlElement("w:color")
    c.set(qn("w:val"), color)
    rpr.append(c)
    if size:
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(size * 2)))
        rpr.append(sz)
    run.append(rpr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    run.append(t)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def add_bookmark(paragraph, name: str, bm_id: int):
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bm_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bm_id))
    p = paragraph._p
    p.insert(0 if p.pPr is None else 1, start)
    p.append(end)


def add_field(paragraph, instr: str, placeholder: str = "1", size=8, color=LIGHT_GRAY):
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    c = OxmlElement("w:color")
    c.set(qn("w:val"), color)
    rpr.append(c)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))
    rpr.append(sz)
    run.append(rpr)
    t = OxmlElement("w:t")
    t.text = placeholder
    run.append(t)
    fld.append(run)
    paragraph._p.append(fld)


def setup_header_footer(doc, left_text: str, right_text: str):
    sec = doc.sections[0]

    sec.header.is_linked_to_previous = False
    hp = sec.header.paragraphs[0]
    hp.style = doc.styles[STYLE_IDS["caption"]]
    hp.paragraph_format.tab_stops.add_tab_stop(Twips(9026), WD_TAB_ALIGNMENT.RIGHT)
    hp.paragraph_format.space_after = Pt(0)
    hp.paragraph_format.line_spacing = 1.0
    r = hp.add_run(left_text)
    r.font.size = Pt(8)
    r.font.color.rgb = RGBColor.from_string(LIGHT_GRAY)
    r.add_tab()
    r2 = hp.add_run(right_text)
    r2.font.size = Pt(8)
    r2.font.color.rgb = RGBColor.from_string(LIGHT_GRAY)

    pPr = hp._p.get_or_add_pPr()
    _pPr_insert(pPr, parse_xml(
        f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="4" w:space="4" w:color="D8DEE4"/></w:pBdr>'),
        "pBdr")

    sec.footer.is_linked_to_previous = False
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.space_before = Pt(0)
    fp.paragraph_format.space_after = Pt(0)

    def small(text):
        rr = fp.add_run(text)
        rr.font.size = Pt(8)
        rr.font.color.rgb = RGBColor.from_string(LIGHT_GRAY)

    small("第 ")
    add_field(fp, " PAGE ")
    small(" 页 / 共 ")
    add_field(fp, " NUMPAGES ")
    small(" 页")


# ===========================================================================
# 四、渲染：块 → 文档
# ===========================================================================

class Ctx:
    def __init__(self):
        self.bm_id = 1000

    def next_bm(self) -> int:
        self.bm_id += 1
        return self.bm_id


def _run(paragraph, text: str, *, bold=False, italic=False, strike=False, code=False,
         color=None, size=None, shade=None):
    if not text:
        return None
    r = paragraph.add_run(text.replace("\t", "    "))
    if bold:
        r.bold = True
    if italic:
        r.italic = True
    if strike:
        r.font.strike = True
    if code:
        r.font.name = MONO_FONT
        r._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), MONO_FONT_CN)
        r.font.size = Pt(size or 9.5)
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    if size and not code:
        r.font.size = Pt(size)
    if shade:
        _shade_run(r, shade)
    return r


def render_inline(paragraph, text, *, base_size=None, bold=False):
    tokens = parse_inline(str(text or ""))
    for tk in tokens:
        txt = de_tex(tk["text"]) if tk.get("math") else tk["text"]
        if tk.get("link"):
            add_hyperlink(paragraph, tk["link"], txt or tk["link"], size=base_size)
            continue
        _run(paragraph,
             txt,
             bold=bold or bool(tk.get("b")),
             italic=bool(tk.get("i")),
             strike=bool(tk.get("strike")),
             code=bool(tk.get("code")),
             color=None if not tk.get("code") else "24292F",
             size=base_size,
             shade="EFF2F5" if tk.get("code") else None)


def add_code_block(doc, text: str):
    lines = str(text).replace("\r\n", "\n").split("\n")
    chunk = 200
    for start in range(0, len(lines), chunk):
        part = lines[start:start + chunk]
        p = doc.add_paragraph(style=STYLE_IDS["code"])
        pf = p.paragraph_format
        pf.space_before = Pt(0 if start else 4)
        pf.space_after = Pt(9 if start + chunk >= len(lines) else 0)
        for idx, ln in enumerate(part):
            if idx:
                p.add_run().add_break()
            r = p.add_run(ln.replace("\t", "    ") if ln else " ")
            r.font.name = MONO_FONT
            r._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), MONO_FONT_CN)
            r.font.size = Pt(9)


def add_table(doc, header, rows, ctx: Ctx):
    n = max([len(header)] + [len(r) for r in rows] + [1])
    table = doc.add_table(rows=1, cols=n)
    try:
        table.style = doc.styles["Table Grid"]
    except KeyError:
        pass
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = True
    table._tbl.tblPr.insert_element_before(
        parse_xml(f'<w:tblCellMar {nsdecls("w")}>'
                  '<w:top w:w="50" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>'
                  '<w:bottom w:w="50" w:type="dxa"/><w:right w:w="100" w:type="dxa"/>'
                  '</w:tblCellMar>'),
        "w:tblLook", "w:tblCaption", "w:tblDescription")

    for idx in range(n):
        cell = table.rows[0].cells[idx]
        cell.text = ""
        p = cell.paragraphs[0]
        p.style = doc.styles[STYLE_IDS["table"]]
        txt = header[idx] if idx < len(header) else ""
        render_inline(p, txt, bold=True)
        _shade_cell(cell, "F2F4F7")

    tr_pr = table.rows[0]._tr.get_or_add_trPr()
    tr_pr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

    for row in rows:
        cells = table.add_row().cells
        for idx in range(n):
            cell = cells[idx]
            cell.text = ""
            p = cell.paragraphs[0]
            p.style = doc.styles[STYLE_IDS["table"]]
            render_inline(p, row[idx] if idx < len(row) else "")

    doc.add_paragraph().paragraph_format.space_after = Pt(2)


HEADING_STYLE = {1: "Heading 1", 2: "Heading 2", 3: "Heading 3", 4: "Heading 4", 5: "Heading 4", 6: "Heading 4"}


def render_blocks(doc, md: str, ctx: Ctx):
    for b in parse_blocks(md):
        t = b["type"]
        if t == "h":
            lvl = min(6, max(1, b["level"]))
            p = doc.add_paragraph(style=HEADING_STYLE[lvl])
            bm = ctx.next_bm()
            add_bookmark(p, f"_Toc{bm}", bm)
            render_inline(p, b["text"], bold=True)
        elif t == "p":
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            render_inline(p, b["text"])
        elif t == "hr":
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(7)
            _add_bottom_border(p, "E4E8ED")
        elif t == "quote":
            p = doc.add_paragraph(style=STYLE_IDS["quote"])
            render_inline(p, b["text"])
        elif t == "code":
            add_code_block(doc, b["text"])
        elif t == "math":
            p = doc.add_paragraph(style=STYLE_IDS["math"])
            r = p.add_run(de_tex(b["text"].replace("\n", " ")))
            r.font.name = MONO_FONT
            r.font.size = Pt(9.5)
            r.italic = True
        elif t == "table":
            add_table(doc, b["header"], b["rows"], ctx)
        elif t == "li":
            lvl = b["level"]
            if b["task"] is not None:
                prefix = "☑ " if b["task"] else "☐ "
            elif b["ordered"]:
                prefix = re.sub(r"[.)]$", "", b["marker"]) + ". "
            else:
                prefix = ["•", "–", "▪"][lvl % 3] + " "
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Twips(360 + lvl * 360)
            p.paragraph_format.first_line_indent = Twips(-300)
            p.paragraph_format.space_after = Pt(3)
            _run(p, prefix, bold=b["ordered"], color="30363D" if b["ordered"] else "57606A")
            render_inline(p, b["text"])


def role_band(doc, msg):
    is_user = msg.get("role") == "user"
    p = doc.add_paragraph(style=STYLE_IDS["role_user"] if is_user else STYLE_IDS["role_ai"])
    _run(p, "你" if is_user else "DeepSeek", bold=True,
         color=USER_BLUE if is_user else "24292F", size=10.5)
    _run(p, "    " + fmt_time(msg.get("time")), color="8B949E", size=8.5)


def render_message(doc, msg, ctx: Ctx, opts):
    role_band(doc, msg)

    if opts.get("think") and msg.get("think"):
        p = doc.add_paragraph(style=STYLE_IDS["think"])
        _run(p, "思考过程", bold=True, size=8.5, color=AMBER)
        p2 = doc.add_paragraph(style=STYLE_IDS["think"])
        _run(p2, msg["think"], size=9.5)

    body = (msg.get("content") or "").strip()
    if body:
        render_blocks(doc, body, ctx)
    else:
        p = doc.add_paragraph()
        _run(p, "（此条内容为空）", italic=True, size=9.5, color=LIGHT_GRAY)

    files = msg.get("files") or []
    if files:
        p = doc.add_paragraph(style=STYLE_IDS["meta"])
        p.paragraph_format.space_before = Pt(5)
        p.paragraph_format.space_after = Pt(11)
        _run(p, "附件：" + "、".join(files), size=8.5, color=GRAY)


def session_meta(doc, sess):
    parts = []
    if sess.get("created_at"):
        parts.append("创建 " + fmt_time(sess.get("created_at")))
    if sess.get("updated_at"):
        parts.append("更新 " + fmt_time(sess.get("updated_at")))
    parts.append(f"共 {len(sess['messages'])} 条消息")
    if sess.get("model_type"):
        parts.append("深度思考模式" if sess["model_type"] == "expert" else "快速模式")
    p = doc.add_paragraph(style=STYLE_IDS["meta"])
    _run(p, "　·　".join(parts), size=8.5, color=GRAY)
    p2 = doc.add_paragraph(style=STYLE_IDS["caption"])
    _run(p2, "会话 ID：" + str(sess.get("id", "")), size=7.5, color=LIGHT_GRAY)
    rule = doc.add_paragraph()
    rule.paragraph_format.space_before = Pt(0)
    rule.paragraph_format.space_after = Pt(7)
    _add_bottom_border(rule, "D8DEE4")


def build_docx(sess, opts) -> Document:
    doc = Document()
    setup_page(doc)
    setup_styles(doc)
    header_left = f"DeepSeek 对话归档 · {sess['title']}"
    setup_header_footer(doc, header_left, "导出于 " + datetime.now().strftime("%Y-%m-%d %H:%M"))

    p = doc.add_paragraph(style="Title")
    _run(p, sess["title"], bold=True)
    session_meta(doc, sess)

    ctx = Ctx()
    for msg in sess["messages"]:
        render_message(doc, msg, ctx, opts)
    return doc


def build_merged(sessions, opts) -> Document:
    doc = Document()
    setup_page(doc)
    setup_styles(doc)
    setup_header_footer(doc, "DeepSeek 对话归档 · 全部会话",
                        "导出于 " + datetime.now().strftime("%Y-%m-%d %H:%M"))
    ctx = Ctx()
    now = datetime.now()
    total = sum(len(s["messages"]) for s in sessions)

    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(0)
    spacer.paragraph_format.space_after = Pt(0)
    spacer.paragraph_format.line_spacing = Pt(130)
    p = doc.add_paragraph(style="Title")
    _run(p, "DeepSeek 对话归档", bold=True, size=20)
    p = doc.add_paragraph(style=STYLE_IDS["meta"])
    _run(p, "C O N V E R S A T I O N   A R C H I V E", size=8, color="8B949E")

    rule = doc.add_paragraph()
    rule.paragraph_format.space_before = Pt(6)
    rule.paragraph_format.space_after = Pt(16)
    _add_bottom_border(rule, "D8DEE4")

    for label, value in (("导出时间", now.strftime("%Y-%m-%d %H:%M")),
                         ("会话数量", f"{len(sessions)} 个"),
                         ("消息总数", f"{total} 条")):
        lp = doc.add_paragraph(style=STYLE_IDS["meta"])
        _run(lp, label, size=8.5, color=GRAY)
        vp = doc.add_paragraph(style=STYLE_IDS["meta"])
        vp.paragraph_format.space_after = Pt(14)
        _run(vp, value, bold=True, size=15)

    doc.add_page_break()
    p = doc.add_paragraph(style="Heading 1")
    _run(p, "目录", bold=True)

    for idx, s in enumerate(sessions):
        anchor = f"_TocS{idx + 1}"
        p = doc.add_paragraph(style=STYLE_IDS["toc"])
        add_anchor_link(p, anchor, f"{idx + 1:02d}　{s['title']}")
        p.add_run().add_tab()
        _run(p, f"{fmt_day(s.get('updated_at') or s.get('created_at'))}　·　{len(s['messages'])} 条",
             size=8.5, color="8B949E")

    for idx, s in enumerate(sessions):
        anchor = f"_TocS{idx + 1}"
        p = doc.add_paragraph(style="Heading 1")
        p.paragraph_format.page_break_before = True
        bm = ctx.next_bm()
        add_bookmark(p, anchor, bm)
        _run(p, f"{idx + 1:02d}　{s['title']}", bold=True)
        session_meta(doc, s)
        for msg in s["messages"]:
            render_message(doc, msg, ctx, opts)

    return doc


# ===========================================================================
# 五、DeepSeek 数据抓取
# ===========================================================================

def headers(token: str):
    h = {
        "x-client-locale": "zh_CN",
        "x-client-platform": "web",
        "x-app-version": "2.0.0",
        "x-client-version": "2.0.0",
        "Referer": BASE_URL + "/",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    }
    if token:
        h["Authorization"] = token if token.lower().startswith("bearer ") else "Bearer " + token
    return h


def api_get(path: str, params=None, token=""):
    url = BASE_URL + API + path
    if params:
        url += "?" + urlencode({k: v for k, v in params.items() if v not in (None, "")})
    req = Request(url, headers=headers(token))
    try:
        with urlopen(req, timeout=30) as resp:
            raw = resp.read().decode("utf-8", "replace")
    except HTTPError as e:
        body = ""
        try:
            body = e.read().decode("utf-8", "replace")[:200]
        except Exception:
            pass
        if e.code in (401, 403):
            raise SystemExit(f"登录状态无效（HTTP {e.code}）。请重新复制 token。{body}")
        raise SystemExit(f"接口返回 HTTP {e.code}：{body}")
    except URLError as e:
        raise SystemExit(f"网络错误：{e.reason}")

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        raise SystemExit("接口返回的不是 JSON，可能被风控拦截。请稍后重试或换成浏览器脚本方案。")

    code = data.get("code", (data.get("data") or {}).get("biz_code", 0))
    if code not in (0, None):
        msg = data.get("msg") or (data.get("data") or {}).get("biz_msg") or f"code {code}"
        raise SystemExit(f"接口错误：{msg}")
    return data


def list_sessions(token: str):
    out, seen = [], set()
    for pinned in (True, False):
        cursor, idle = None, 0
        for _ in range(400):
            params = {"lte_cursor.pinned": "true" if pinned else "false"}
            if cursor is not None:
                params["lte_cursor.updated_at"] = cursor
            data = api_get("/chat_session/fetch_page", params, token)
            items = ((data.get("data") or {}).get("biz_data") or {}).get("chat_sessions") or []
            if not items:
                break
            added = 0
            for it in items:
                sid = it.get("id")
                if sid and sid not in seen:
                    seen.add(sid)
                    out.append(it)
                    added += 1
            if not added:
                idle += 1
                if idle >= 2:
                    break
            else:
                idle = 0
            nxt = items[-1].get("updated_at")
            if nxt is None or nxt == cursor:
                break
            cursor = nxt
    out.sort(key=lambda s: s.get("updated_at") or s.get("inserted_at") or 0, reverse=True)
    return out


def fetch_messages(sid: str, token: str):
    data = api_get("/chat/history_messages", {"chat_session_id": sid}, token)
    return (data.get("data") or {}).get("biz_data") or None


def normalize_message(m):
    if not m:
        return None
    role_raw = str(m.get("role") or m.get("author") or "").upper()
    role = "user" if role_raw == "USER" else ("assistant" if role_raw in ("ASSISTANT", "BOT") else "")
    if not role:
        return None

    content_parts, thinks, files = [], [], []
    for f in m.get("fragments") or []:
        if not f:
            continue
        t = str(f.get("type") or "").upper()
        c = f.get("content")
        if isinstance(c, dict):
            c = c.get("content") or json.dumps(c, ensure_ascii=False)
        c = c if isinstance(c, str) else ""
        if t in ("REQUEST", "RESPONSE", "TEXT", "CONTENT"):
            content_parts.append(c)
        elif t in ("THINK", "REASONING"):
            thinks.append(c)
        elif t in ("FILE", "DOCUMENT"):
            name = ""
            try:
                j = json.loads(c)
                name = j.get("name") or j.get("file_name") or j.get("filename") or (j.get("file") or {}).get("name") or ""
            except (json.JSONDecodeError, AttributeError):
                pass
            if name:
                files.append(name)

    content = "\n\n".join(x for x in content_parts if x and x.strip())
    if not content and isinstance(m.get("content"), str):
        content = m["content"]
    if not content and isinstance(m.get("content"), dict):
        content = str(m["content"].get("content") or "")
    if not content and isinstance(m.get("text"), str):
        content = m["text"]

    return {
        "role": role,
        "time": m.get("inserted_at") or m.get("created_at") or m.get("updated_at"),
        "content": content or "",
        "think": "\n\n".join(x for x in thinks if x and x.strip()),
        "files": files,
        "status": m.get("status") or "",
    }


def normalize_session(info, biz):
    merged = dict(info or {})
    merged.update((biz or {}).get("chat_session") or {})
    messages = []
    for m in (biz or {}).get("chat_messages") or []:
        nm = normalize_message(m)
        if not nm:
            continue
        if nm["status"] == "WIP" and not nm["content"] and not nm["think"]:
            continue
        if not nm["content"] and not nm["think"] and not nm["files"]:
            continue
        messages.append(nm)
    return {
        "id": merged.get("id") or (info or {}).get("id") or "",
        "title": str(merged.get("title") or "").strip() or "未命名会话",
        "created_at": merged.get("inserted_at") or (info or {}).get("inserted_at"),
        "updated_at": merged.get("updated_at") or (info or {}).get("updated_at"),
        "model_type": merged.get("model_type") or "",
        "pinned": bool((info or {}).get("pinned")),
        "messages": messages,
    }


# ===========================================================================
# 六、输出
# ===========================================================================

README_TXT = """DeepSeek 对话归档
==================

导出时间：{when}
会话数量：{count}
消息条数：{msgs}
工具版本：v{ver}

目录说明：
  会话/              每个会话一份 Word 文档（.docx），按时间从早到晚编号
  全部会话合集.docx   所有会话合并成的一份文档（含目录，可点击跳转）
  原始数据.json      结构化原始数据，可用于二次处理或重新生成

说明：Word 文档可直接用 Microsoft Word / WPS / LibreOffice 打开。
      代码块、表格、标题层级、超链接均已尽量保留。
      数学公式按 LaTeX 源码形式保留（未做公式排版）。
"""


def write_outputs(sessions, out_dir: Path, opts):
    sess_dir = out_dir / "会话"
    sess_dir.mkdir(parents=True, exist_ok=True)
    sessions = sorted(sessions, key=lambda s: s.get("created_at") or 0)

    for i, s in enumerate(sessions, 1):
        doc = build_docx(s, opts)
        name = f"{i:03d}_{safe_name(s['title'])}.docx"
        doc.save(sess_dir / name)
        print(f"  [{'%d' % i}/{len(sessions)}] {name}")

    if opts.get("merge") and len(sessions) > 1:
        doc = build_merged(sessions, opts)
        doc.save(out_dir / "全部会话合集.docx")
        print("  [合并] 全部会话合集.docx")

    payload = {
        "tool": "deepseek-archiver",
        "version": VERSION,
        "source": BASE_URL,
        "exported_at": datetime.now().astimezone().isoformat(),
        "count": len(sessions),
        "message_count": sum(len(s["messages"]) for s in sessions),
        "sessions": sessions,
    }
    (out_dir / "原始数据.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    (out_dir / "归档说明.txt").write_text(
        README_TXT.format(when=datetime.now().strftime("%Y-%m-%d %H:%M"), count=len(sessions),
                          msgs=payload["message_count"], ver=VERSION),
        encoding="utf-8-sig")


def load_payload(path: Path):
    if path.suffix.lower() == ".zip":
        with zipfile.ZipFile(path) as z:
            candidates = [n for n in z.namelist() if n.lower().endswith(".json") and "原始数据" in n]
            if not candidates:
                candidates = [n for n in z.namelist() if n.lower().endswith(".json")]
            if not candidates:
                raise SystemExit("压缩包里没找到 .json 数据文件")
            raw = z.read(candidates[0]).decode("utf-8-sig")
    else:
        raw = path.read_text(encoding="utf-8-sig")
    data = json.loads(raw)
    if isinstance(data, dict):
        if "sessions" in data:
            return data["sessions"]
        if "data" in data and isinstance(data["data"], dict) and "sessions" in data["data"]:
            return data["data"]["sessions"]
    if isinstance(data, list):
        return data
    raise SystemExit("无法识别的数据格式")


def main(argv=None):
    ap = argparse.ArgumentParser(description="把 DeepSeek 网页端会话归档为 Word (.docx)")
    sub = ap.add_subparsers(dest="cmd")

    f = sub.add_parser("fetch", help="直接调用接口抓取（需要 token）")
    f.add_argument("--token", default=os.environ.get("DEEPSEEK_TOKEN", ""), help="Bearer token，或设置环境变量 DEEPSEEK_TOKEN")
    f.add_argument("-o", "--out", default="DeepSeek归档", help="输出目录")
    f.add_argument("--merge", action="store_true", help="额外生成合并总文档")
    f.add_argument("--think", action="store_true", help="保留思考过程")
    f.add_argument("--limit", type=int, default=0, help="只导出最近 N 个会话")
    f.add_argument("--id", default="", help="只导出指定会话 ID")

    c = sub.add_parser("convert", help="把已导出的 json/zip 转成 Word")
    c.add_argument("input", help="原始数据 .json 或下载的 .zip")
    c.add_argument("-o", "--out", default="", help="输出目录（默认与输入同名）")
    c.add_argument("--merge", action="store_true", help="额外生成合并总文档")
    c.add_argument("--think", action="store_true", help="保留思考过程")

    args = ap.parse_args(argv)
    if not args.cmd:
        ap.print_help()
        return 1

    opts = {"think": args.think, "merge": args.merge}

    if args.cmd == "convert":
        src = Path(args.input)
        if not src.exists():
            raise SystemExit(f"文件不存在：{src}")
        sessions = load_payload(src)
        out = Path(args.out) if args.out else src.with_name(src.stem + "_Word")
        out.mkdir(parents=True, exist_ok=True)
        print(f"读取 {len(sessions)} 个会话 → {out}")
        write_outputs(sessions, out, opts)
        print("完成")
        return 0

    token = args.token.strip()
    if not token:
        raise SystemExit("缺少 token。用法：--token <TOKEN>，或先设置环境变量 DEEPSEEK_TOKEN。"
                         "\n取法：登录 chat.deepseek.com → F12 → Application → Local Storage → userToken → value")
    if not token.lower().startswith("bearer "):
        token = "Bearer " + token

    if args.id:
        info = {"id": args.id}
        sessions = [normalize_session(info, fetch_messages(args.id, token))]
    else:
        print("读取会话列表…")
        listing = list_sessions(token)
        if not listing:
            raise SystemExit("没有拿到任何会话，请确认账号里有对话记录")
        print(f"共 {len(listing)} 个会话")
        if args.limit > 0:
            listing = listing[:args.limit]
        sessions = []
        for i, info in enumerate(listing, 1):
            try:
                sess = normalize_session(info, fetch_messages(info["id"], token))
                if sess["messages"]:
                    sessions.append(sess)
                sys.stdout.write(f"\r  抓取 {i}/{len(listing)}  {info.get('title', '')[:30]:<32}")
                sys.stdout.flush()
            except SystemExit as e:
                print(f"\n  跳过 {info.get('id')}：{e}")
            time.sleep(0.15)
        print()

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    write_outputs(sessions, out, opts)
    print(f"完成，输出目录：{out.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
