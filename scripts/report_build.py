#!/usr/bin/env python
"""把 Markdown 内容 + 图片清单装配成图文结合的 Word 调研报告。

用法:
    python report_build.py <content.md> [--manifest assets/manifest.json]
                           [--out report.docx] [--assets assets]

内容文件语法（见 skill 文档）:
    ---
    title: ...            # 报告标题
    subtitle: ...         # 副标题（可选）
    author: ...           # 编制（可选）
    date: 2026-10-09      # 日期
    version: v1.0         # 版本（可选）
    ---

    [[TOC]]                        目录占位
    # 1 概述                        一级标题
    ## 1.1 背景                     二级标题
    普通段落，支持 **加粗**、`等宽`、[链接](url)。
    - 列表项
    1. 编号项
    > 摘要/提示段落（左侧缩进的引用样式）

    | 表头 | 表头 |                  表格（表题写在表格上方）
    | --- | --- |
    | 值 | 值 |

    [[表:tab-01]]                   正文中的表号引用 → “表 1”
    [[图:fig-01]]                   正文中的图号引用 → “图 3”

    :::figure fig-01                插图块；图号按出现顺序自动生成
    caption: 重组蛋白疫苗典型下游纯化工艺流程
    source: 依据 Cytiva 公开工艺资料整理
    note: 本图为技术示意图，非真实照片
    :::

    :::table tab-01                 插表块（表题在表上方）
    caption: 主要层析技术对比
    :::
    | 技术 | 分离原理 | 适用阶段 |
    | --- | --- | --- |
    | 亲和层析 | 生物特异性识别 | 捕获 |
    :::end

图片信息（文件路径、类型、来源、许可）从 manifest.json 读取，并在构建时做红线校验。
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image

# ---- 排版常量 ---------------------------------------------------------------
FONT_BODY_EA = "宋体"
FONT_BODY_ASCII = "Times New Roman"
FONT_HEAD_EA = "微软雅黑"
FONT_HEAD_ASCII = "Segoe UI"
SIZE_BODY = Pt(10.5)
SIZE_H1 = Pt(16)
SIZE_H2 = Pt(14)
SIZE_H3 = Pt(12)
SIZE_CAPTION = Pt(9)
SIZE_TABLE = Pt(9.5)
CONTENT_WIDTH_CM = 14.6          # A4 宽 21cm - 左右页边距 3.2cm*2
MAX_IMAGE_HEIGHT_CM = 13.0       # 单张图在页面上的最大高度，避免整页独占、留白过大
TABLE_HEADER_FILL = "DEEAF6"


# ---- 基础工具 ---------------------------------------------------------------
def set_run_font(run, ea=FONT_BODY_EA, ascii_font=FONT_BODY_ASCII, size=SIZE_BODY,
                 bold=None, italic=None, color=None):
    run.font.size = size
    run.font.name = ascii_font
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), ascii_font)
    rfonts.set(qn("w:hAnsi"), ascii_font)
    rfonts.set(qn("w:eastAsia"), ea)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)


def set_style_font(style, ea, ascii_font, size, bold=None, color="000000"):
    style.font.size = size
    style.font.name = ascii_font
    if bold is not None:
        style.font.bold = bold
    if color is not None:
        style.font.color.rgb = RGBColor.from_string(color)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), ascii_font)
    rfonts.set(qn("w:hAnsi"), ascii_font)
    rfonts.set(qn("w:eastAsia"), ea)


def add_hyperlink(paragraph, url: str, text: str, size=SIZE_BODY):
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    rfonts = OxmlElement("w:rFonts")
    rfonts.set(qn("w:ascii"), FONT_BODY_ASCII)
    rfonts.set(qn("w:hAnsi"), FONT_BODY_ASCII)
    rfonts.set(qn("w:eastAsia"), FONT_BODY_EA)
    rpr.append(rfonts)
    for tag, val in (("w:color", "0563C1"),):
        el = OxmlElement(tag)
        el.set(qn("w:val"), val)
        rpr.append(el)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rpr.append(u)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size.pt * 2)))
    rpr.append(sz)
    new_run.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def shade_cell(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def mark_row_as_header(row):
    """让表头行在跨页时自动重复。"""
    tr_pr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    tr_pr.append(el)


def mark_row_no_split(row):
    """禁止表格行被分页拆成两半。"""
    tr_pr = row._tr.get_or_add_trPr()
    tr_pr.append(OxmlElement("w:cantSplit"))


def add_toc_field(paragraph):
    """插入 Word 目录域（打开/导出时更新）。"""
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = r' TOC \o "1-3" \h \z \u '
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = "打开文档时按 F9 或导出时将自动生成目录"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)
    run._r.append(placeholder)
    run._r.append(fld_end)
    set_run_font(run, size=SIZE_BODY)


def set_update_fields_on_open(document):
    settings = document.settings.element
    for el in settings.findall(qn("w:updateFields")):
        settings.remove(el)
    uf = OxmlElement("w:updateFields")
    uf.set(qn("w:val"), "true")
    settings.append(uf)


# ---- 文档级设置 -------------------------------------------------------------
def setup_document(doc: Document):
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.2)
    section.right_margin = Cm(3.2)

    normal = doc.styles["Normal"]
    set_style_font(normal, FONT_BODY_EA, FONT_BODY_ASCII, SIZE_BODY, color="000000")
    pf = normal.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.5
    pf.space_after = Pt(6)
    pf.first_line_indent = Pt(21)  # 首行缩进 2 字符

    set_style_font(doc.styles["Heading 1"], FONT_HEAD_EA, FONT_HEAD_ASCII, SIZE_H1, bold=True)
    set_style_font(doc.styles["Heading 2"], FONT_HEAD_EA, FONT_HEAD_ASCII, SIZE_H2, bold=True)
    set_style_font(doc.styles["Heading 3"], FONT_HEAD_EA, FONT_HEAD_ASCII, SIZE_H3, bold=True)
    for name in ("Heading 1", "Heading 2", "Heading 3"):
        p = doc.styles[name].paragraph_format
        p.space_before = Pt(14 if name == "Heading 1" else 10)
        p.space_after = Pt(8 if name == "Heading 1" else 6)
        p.first_line_indent = Pt(0)
        p.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        p.line_spacing = 1.25
        p.keep_with_next = True


# ---- 行内解析 ---------------------------------------------------------------
INLINE_RE = re.compile(
    r"(\*\*.+?\*\*|`[^`]+`|\[\[图:[^\]]+\]\]|\[\[表:[^\]]+\]\]|\[[^\]]+\]\([^)]+\))"
)


class Numbering:
    """图号 / 表号登记表。"""

    def __init__(self):
        self.figures: list[str] = []
        self.tables: list[str] = []

    def fig_no(self, fid: str) -> int:
        return self.figures.index(fid) + 1

    def tab_no(self, tid: str) -> int:
        return self.tables.index(tid) + 1


def render_inline(paragraph, text: str, numbering: Numbering, size=SIZE_BODY, ea=FONT_BODY_EA):
    for token in INLINE_RE.split(text):
        if not token:
            continue
        if token.startswith("**") and token.endswith("**") and len(token) > 4:
            run = paragraph.add_run(token[2:-2])
            set_run_font(run, ea=ea, size=size, bold=True)
        elif token.startswith("`") and token.endswith("`") and len(token) > 2:
            run = paragraph.add_run(token[1:-1])
            set_run_font(run, ea="Consolas", ascii_font="Consolas", size=size)
        elif token.startswith("[[图:") and token.endswith("]]"):
            fid = token[4:-2].strip()
            no = numbering.fig_no(fid) if fid in numbering.figures else 0
            run = paragraph.add_run(f"图 {no}" if no else "图 ?")
            set_run_font(run, ea=ea, size=size)
        elif token.startswith("[[表:") and token.endswith("]]"):
            tid = token[4:-2].strip()
            no = numbering.tab_no(tid) if tid in numbering.tables else 0
            run = paragraph.add_run(f"表 {no}" if no else "表 ?")
            set_run_font(run, ea=ea, size=size)
        elif token.startswith("[") and "](" in token and token.endswith(")"):
            label, url = token[1:-1].split("](", 1)
            add_hyperlink(paragraph, url, label, size=size)
        else:
            run = paragraph.add_run(token)
            set_run_font(run, ea=ea, size=size)


# ---- 内容解析 ---------------------------------------------------------------
def parse_front_matter(text: str) -> tuple[dict, str]:
    meta: dict = {}
    if text.lstrip().startswith("---"):
        stripped = text.lstrip()
        end = stripped.find("\n---", 3)
        if end != -1:
            block = stripped[3:end].strip()
            rest = stripped[end + 4 :].lstrip("\n")
            for line in block.splitlines():
                if ":" in line:
                    k, v = line.split(":", 1)
                    meta[k.strip().lower()] = v.strip().strip('"').strip("'")
            return meta, rest
    return meta, text


def parse_pipe_table(lines: list[str]) -> tuple[list[str], list[list[str]]]:
    """从若干行中解析 Markdown 管道表格，返回 (表头, 数据行)。"""
    rows = [ln.strip() for ln in lines if ln.strip().startswith("|")]
    if not rows:
        return [], []
    header = [c.strip() for c in rows[0].strip("|").split("|")]
    data: list[list[str]] = []
    for ln in rows[1:]:
        if re.match(r"^\|[\s:\-|]+\|$", ln):
            continue
        data.append([c.strip() for c in ln.strip("|").split("|")])
    return header, data


def parse_blocks(body: str) -> list[dict]:
    lines = body.splitlines()
    blocks: list[dict] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if not stripped:
            i += 1
            continue

        if stripped.startswith(":::"):
            header = stripped[3:].strip()
            kind, _, ident = header.partition(" ")
            ident = ident.strip()
            if ident in ("end", "") and kind == "end":
                i += 1
                continue
            meta: dict = {}
            i += 1
            inner: list[str] = []
            while i < len(lines):
                cur = lines[i].strip()
                if cur == ":::" or cur == ":::end":
                    i += 1
                    break
                if re.match(r"^(caption|source|note|title|license)\s*:", cur):
                    k, v = cur.split(":", 1)
                    meta[k.strip().lower()] = v.strip()
                    i += 1
                    continue
                inner.append(lines[i])
                i += 1
            block = {"type": kind, "id": ident, "meta": meta, "body": "\n".join(inner)}
            if kind == "table":
                header, rows = parse_pipe_table(inner)
                block["header"] = header
                block["rows"] = rows
            blocks.append(block)
            continue

        if stripped == "[[TOC]]":
            blocks.append({"type": "toc"})
            i += 1
            continue

        m = re.match(r"^(#{1,4})\s+(.*)$", stripped)
        if m:
            blocks.append({"type": "heading", "level": len(m.group(1)), "text": m.group(2).strip()})
            i += 1
            continue

        if stripped.startswith("|") and i + 1 < len(lines) and re.match(
            r"^\|[\s:\-|]+\|$", lines[i + 1].strip()
        ):
            rows = []
            header = [c.strip() for c in stripped.strip("|").split("|")]
            i += 2
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
                i += 1
            blocks.append({"type": "table", "header": header, "rows": rows})
            continue

        if stripped.startswith("> "):
            buf = []
            while i < len(lines) and lines[i].strip().startswith("> "):
                buf.append(lines[i].strip()[2:].strip())
                i += 1
            blocks.append({"type": "quote", "text": " ".join(buf)})
            continue

        if re.match(r"^[-*]\s+", stripped):
            items = []
            while i < len(lines) and re.match(r"^[-*]\s+", lines[i].strip()):
                items.append(re.sub(r"^[-*]\s+", "", lines[i].strip()))
                i += 1
            blocks.append({"type": "ul", "items": items})
            continue

        if re.match(r"^\d+[.)]\s+", stripped):
            items = []
            while i < len(lines) and re.match(r"^\d+[.)]\s+", lines[i].strip()):
                items.append(re.sub(r"^\d+[.)]\s+", "", lines[i].strip()))
                i += 1
            blocks.append({"type": "ol", "items": items})
            continue

        buf = [stripped]
        i += 1
        while i < len(lines):
            nxt = lines[i].strip()
            if (not nxt or nxt.startswith(":::") or nxt.startswith("#")
                    or nxt.startswith("|") or nxt.startswith("> ")
                    or re.match(r"^[-*]\s+", nxt) or re.match(r"^\d+[.)]\s+", nxt)
                    or nxt == "[[TOC]]"):
                break
            buf.append(nxt)
            i += 1
        blocks.append({"type": "para", "text": " ".join(buf)})
    return blocks


def scan_numbering(blocks: list[dict]) -> Numbering:
    n = Numbering()
    for b in blocks:
        if b["type"] == "figure" and b["id"] and b["id"] not in n.figures:
            n.figures.append(b["id"])
        if b["type"] == "table" and b.get("id") and b["id"] not in n.tables:
            n.tables.append(b["id"])
    return n


# ---- 图片清单与红线校验 -----------------------------------------------------
REQUIRED = {
    "photo": ("url", "source"),
    "schematic": ("generator",),
    "chart": ("generator",),
    "ai": (),
}


def load_manifest(path: Path) -> dict:
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    out = {}
    for item in data.get("figures", []):
        out[item["id"]] = item
    return out


def validate_figure(entry: dict, root: Path) -> list[str]:
    """红线校验，返回问题列表。"""
    problems = []
    fid = entry.get("id", "?")
    ftype = entry.get("type", "")
    if ftype not in REQUIRED:
        problems.append(f"[{fid}] 图片类型缺失或非法: {ftype!r}（应为 photo/schematic/chart/ai）")
        return problems
    for field in REQUIRED[ftype]:
        if not entry.get(field):
            problems.append(f"[{fid}] 类型 {ftype} 缺少必需字段: {field}")
    if ftype == "photo":
        if not str(entry.get("url", "")).startswith("http"):
            problems.append(f"[{fid}] 真实图片必须记录可核实的原始链接（url）")
        if not entry.get("license"):
            problems.append(f"[{fid}] 真实图片必须记录许可信息（license，未知时写“未标注/需确认”）")
    if ftype == "ai":
        title = entry.get("title", "")
        if "AI" not in title.upper() and "ai" not in title:
            problems.append(f"[{fid}] AI 生成图必须在校验标题中标注“AI 生成”，避免被误认为真实资料")
    if ftype in ("schematic", "chart") and not entry.get("source"):
        problems.append(f"[{fid}] 示意图/数据图必须记录依据来源（source）")
    file_rel = entry.get("file")
    if not file_rel:
        problems.append(f"[{fid}] 缺少图片文件路径 file")
    elif not (root / file_rel).exists():
        problems.append(f"[{fid}] 图片文件不存在: {file_rel}")
    return problems


def fit_image(path: Path, max_w_cm: float, max_h_cm: float) -> tuple[float, float]:
    with Image.open(path) as im:
        w_px, h_px = im.size
        dpi = im.info.get("dpi") or (96, 96)
    w_cm = w_px / (dpi[0] or 96) * 2.54
    h_cm = h_px / (dpi[1] or 96) * 2.54
    # 统一按“能放多大放多大”的规则：宽高都限制在版心内，不影响清晰度（像素不变）
    scale = min(max_w_cm / w_cm, max_h_cm / h_cm)
    return w_cm * scale, h_cm * scale


# ---- 渲染 -------------------------------------------------------------------
def caption_paragraph(doc, text: str, kind: str = "figure"):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.first_line_indent = Pt(0)
    pf.space_before = Pt(4 if kind == "figure" else 2)
    pf.space_after = Pt(10 if kind == "figure" else 6)
    pf.line_spacing = 1.2
    pf.keep_with_next = kind == "table"
    return p


def add_source_line(doc, text: str):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.first_line_indent = Pt(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(12)
    pf.line_spacing = 1.15
    run = p.add_run(text)
    set_run_font(run, ea=FONT_BODY_EA, size=Pt(8.5), color="595959")
    return p


def render_table_block(doc, blk: dict, numbering: Numbering, caption_text: str | None = None):
    header = blk.get("header") or []
    rows = blk.get("rows") or []
    if not header:
        return
    if caption_text:
        cap = caption_paragraph(doc, caption_text, kind="table")
        run = cap.add_run(caption_text)
        set_run_font(run, ea=FONT_HEAD_EA, size=SIZE_CAPTION, bold=True)
        cap.paragraph_format.space_before = Pt(10)

    table = doc.add_table(rows=1, cols=len(header))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    width_cm = CONTENT_WIDTH_CM / max(len(header), 1)
    mark_row_as_header(table.rows[0])
    mark_row_no_split(table.rows[0])

    for j, text in enumerate(header):
        cell = table.rows[0].cells[j]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.line_spacing = 1.1
        run = p.add_run(text)
        set_run_font(run, ea=FONT_HEAD_EA, size=SIZE_TABLE, bold=True)
        shade_cell(cell, TABLE_HEADER_FILL)
        cell.width = Cm(width_cm)

    for r in rows:
        row_obj = table.add_row()
        mark_row_no_split(row_obj)
        cells = row_obj.cells
        for j, text in enumerate(r[: len(header)]):
            cells[j].text = ""
            p = cells[j].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.line_spacing = 1.1
            render_inline(p, text, numbering, size=SIZE_TABLE)
            cells[j].width = Cm(width_cm)

    tail = doc.add_paragraph()
    tail.paragraph_format.space_after = Pt(4)
    tail.paragraph_format.first_line_indent = Pt(0)


def build(content_path: Path, manifest_path: Path, out_path: Path, assets_root: Path) -> dict:
    text = content_path.read_text(encoding="utf-8")
    meta, body = parse_front_matter(text)
    blocks = parse_blocks(body)
    numbering = scan_numbering(blocks)
    manifest = load_manifest(manifest_path)

    problems: list[str] = []
    used_figures: list[str] = []
    for b in blocks:
        if b["type"] == "figure":
            fid = b["id"]
            entry = manifest.get(fid)
            if not entry:
                problems.append(f"[{fid}] 图片清单中没有该条目")
                continue
            used_figures.append(fid)
            problems.extend(validate_figure(entry, assets_root.parent))

    if problems:
        raise SystemExit("图片红线校验未通过：\n  - " + "\n  - ".join(problems))

    doc = Document()
    setup_document(doc)
    set_update_fields_on_open(doc)

    # 封面标题
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.first_line_indent = Pt(0)
    title_p.paragraph_format.space_before = Pt(24)
    title_p.paragraph_format.space_after = Pt(6)
    run = title_p.add_run(meta.get("title", content_path.stem))
    set_run_font(run, ea=FONT_HEAD_EA, ascii_font=FONT_HEAD_ASCII, size=Pt(22), bold=True)

    if meta.get("subtitle"):
        sp = doc.add_paragraph()
        sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sp.paragraph_format.first_line_indent = Pt(0)
        sp.paragraph_format.space_after = Pt(14)
        run = sp.add_run(meta["subtitle"])
        set_run_font(run, ea=FONT_HEAD_EA, size=Pt(13), color="404040")

    info_bits = []
    if meta.get("author"):
        info_bits.append(f"编制：{meta['author']}")
    if meta.get("date"):
        info_bits.append(f"日期：{meta['date']}")
    if meta.get("version"):
        info_bits.append(f"版本：{meta['version']}")
    if info_bits:
        ip = doc.add_paragraph()
        ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
        ip.paragraph_format.first_line_indent = Pt(0)
        ip.paragraph_format.space_after = Pt(18)
        run = ip.add_run("　|　".join(info_bits))
        set_run_font(run, ea=FONT_BODY_EA, size=Pt(10), color="595959")

    # 封面单独成页
    doc.add_page_break()

    fig_counter = 0
    pending_table_caption = None

    for b in blocks:
        kind = b["type"]

        if kind == "heading":
            level = min(b["level"], 3)
            h = doc.add_heading(b["text"], level=level)
            for r in h.runs:
                set_run_font(
                    r,
                    ea=FONT_HEAD_EA,
                    ascii_font=FONT_HEAD_ASCII,
                    size={1: SIZE_H1, 2: SIZE_H2, 3: SIZE_H3}[level],
                    bold=True,
                    color="000000",
                )

        elif kind == "toc":
            # 目录标题故意不用 Heading 1 样式，避免它自己被收进目录
            h = doc.add_paragraph()
            h.paragraph_format.first_line_indent = Pt(0)
            h.paragraph_format.space_before = Pt(6)
            h.paragraph_format.space_after = Pt(12)
            run = h.add_run("目录")
            set_run_font(run, ea=FONT_HEAD_EA, ascii_font=FONT_HEAD_ASCII,
                         size=SIZE_H1, bold=True, color="000000")
            tp = doc.add_paragraph()
            tp.paragraph_format.first_line_indent = Pt(0)
            add_toc_field(tp)
            doc.add_page_break()

        elif kind == "para":
            p = doc.add_paragraph()
            render_inline(p, b["text"], numbering)

        elif kind == "quote":
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.left_indent = Cm(0.75)
            pf.first_line_indent = Pt(0)
            pf.space_before = Pt(6)
            pf.space_after = Pt(10)
            render_inline(p, b["text"], numbering, size=Pt(10))
            for r in p.runs:
                r.font.color.rgb = RGBColor.from_string("404040")
            pPr = p._p.get_or_add_pPr()
            pbdr = OxmlElement("w:pBdr")
            left = OxmlElement("w:left")
            left.set(qn("w:val"), "single")
            left.set(qn("w:sz"), "18")
            left.set(qn("w:space"), "8")
            left.set(qn("w:color"), "2E75B6")
            pbdr.append(left)
            pPr.append(pbdr)

        elif kind == "ul":
            for item in b["items"]:
                p = doc.add_paragraph()
                pf = p.paragraph_format
                pf.left_indent = Cm(0.75)
                pf.first_line_indent = Cm(-0.35)
                pf.space_after = Pt(3)
                run = p.add_run("•  ")
                set_run_font(run)
                render_inline(p, item, numbering)

        elif kind == "ol":
            for idx, item in enumerate(b["items"], 1):
                p = doc.add_paragraph()
                pf = p.paragraph_format
                pf.left_indent = Cm(0.85)
                pf.first_line_indent = Cm(-0.45)
                pf.space_after = Pt(3)
                run = p.add_run(f"{idx}.  ")
                set_run_font(run)
                render_inline(p, item, numbering)

        elif kind == "table":
            caption = b.get("meta", {}).get("caption")
            if caption:
                no = numbering.tab_no(b["id"]) if b.get("id") else None
                caption = f"表 {no}　{caption}" if no else caption
            render_table_block(doc, b, numbering, caption)

        elif kind == "figure":
            entry = manifest[b["id"]]
            fig_counter += 1
            no = numbering.fig_no(b["id"])
            file_path = assets_root.parent / entry["file"]
            w_cm, h_cm = fit_image(file_path, CONTENT_WIDTH_CM, MAX_IMAGE_HEIGHT_CM)

            pic_p = doc.add_paragraph()
            pic_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf = pic_p.paragraph_format
            pf.first_line_indent = Pt(0)
            pf.space_before = Pt(10)
            pf.space_after = Pt(2)
            pf.keep_with_next = True
            pf.keep_together = True
            pic_p.add_run().add_picture(str(file_path), width=Cm(w_cm), height=Cm(h_cm))

            cap_text = b["meta"].get("caption") or entry.get("title", "")
            cap = caption_paragraph(doc, "", kind="figure")
            run = cap.add_run(f"图 {no}　{cap_text}")
            set_run_font(run, ea=FONT_HEAD_EA, size=SIZE_CAPTION, bold=True)
            cap.paragraph_format.keep_with_next = True

            src_bits = []
            src = b["meta"].get("source") or entry.get("source")
            if src:
                src_bits.append(f"来源：{src}")
            if entry.get("type") == "photo" and entry.get("url"):
                # 原始链接已保存在 manifest.json 与参考文献中，图注里只保留可读的署名信息
                lic = str(entry.get("license", ""))
                if entry.get("author") and entry["author"] not in lic:
                    src_bits.append(f"作者：{entry['author']}")
            if entry.get("license"):
                src_bits.append(f"许可：{entry['license']}")
            if b["meta"].get("note"):
                src_bits.append(b["meta"]["note"])
            if src_bits:
                add_source_line(doc, "　".join(src_bits))

            pending_table_caption = None

        elif kind == "table_end":
            continue

    # 若正文未写参考文献章节，这里不自动补；交由内容文件负责
    doc.save(out_path)

    return {
        "out": str(out_path),
        "figures": len(numbering.figures),
        "tables": len(numbering.tables),
        "used_figures": used_figures,
        "manifest_entries": len(manifest),
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("content", help="Markdown 内容文件")
    ap.add_argument("--manifest", default=None, help="图片清单 manifest.json")
    ap.add_argument("--out", default=None, help="输出 docx 路径")
    ap.add_argument("--assets", default=None, help="assets 目录（默认与内容同级的 assets）")
    args = ap.parse_args()

    content_path = Path(args.content).resolve()
    base = content_path.parent
    assets_root = Path(args.assets).resolve() if args.assets else base / "assets"
    manifest_path = Path(args.manifest).resolve() if args.manifest else assets_root / "manifest.json"
    out_path = Path(args.out).resolve() if args.out else base / "report.docx"

    info = build(content_path, manifest_path, out_path, assets_root)
    print(f"DOCX: {info['out']}")
    print(f"  图 {info['figures']} 张 / 表 {info['tables']} 个 / 图片清单 {info['manifest_entries']} 条")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
