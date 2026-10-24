#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把预览稿 HTML 转成「可编辑的 Word 工作稿」。

为什么需要它：PDF 是给客户"看"的，改不了字。客户想在上面改文案时，
需要一份 Word —— 表格是真表格（能直接在单元格里改），标题是真标题
（能在导航窗格里跳），待确认项是能打勾的清单。

零依赖：只用标准库 zipfile / xml.etree / html.parser，不需要 pip install
任何东西。通用 HTML→DOCX 转换器（pandoc、html-for-docx 等）处理不了
预览稿里的 flex 布局，会把表格整块丢掉，所以这里专门按预览稿的结构写。

用法：
    python3 html_to_docx.py input.html [-o output.docx]

不传 -o 时默认与 input 同目录、同名 .docx。

视觉元素怎么在 Word 里表达（没有色块，改用文字标记）：
    .placeholder  → 【权益·待客户提供】
    .stock-note   → ［素材剪辑］…
    .improv-note  → ［拍摄提示］…
    .blank        → ＿＿＿［现场发挥区］＿＿＿
    .dlg-label    → ▸ 预演对话（可自由发挥）
    .tag-shoot    → 〔实拍〕
    .tag-stock    → 〔素材剪辑〕
"""

import argparse
import html
import os
import sys
import unicodedata
import zipfile
from html.parser import HTMLParser
from xml.sax.saxutils import escape

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

FONT_ASCII = "Microsoft YaHei"
FONT_EA = "微软雅黑"
COLOR_MAIN = "8B4513"      # 标题主色（与预览稿一致）
COLOR_HEAD_FILL = "F0F0F0"  # 表头底纹
COLOR_LABEL_FILL = "FAFAFA"  # 首列底纹
COLOR_BORDER = "CCCCCC"
BODY_WIDTH_DXA = 9026       # A4 纵向 - 1 英寸左右边距

VOID_TAGS = {"br", "img", "hr", "meta", "link", "input", "col"}

# 带这些 class 的内联元素，单独成段并加文字标记
MARK_CLASS = {
    "placeholder": ("【", "】"),
    "stock-note": ("［素材剪辑］", ""),
    "improv-note": ("［拍摄提示］", ""),
    "blank": ("＿＿＿［", "］＿＿＿"),
    "dlg-label": ("▸ ", ""),
    "tag-shoot": ("〔实拍〕", ""),
    "tag-stock": ("〔素材剪辑〕", ""),
    "ref-title": ("▸ ", ""),
}

H_TAGS = {"h1": "Heading1", "h2": "Heading2", "h3": "Heading3",
          "h4": "Heading3", "h5": "Heading3", "h6": "Heading3"}

# 项目卡片（模块 1）在预览稿 HTML 里是 hero 区，没有 .module 容器，
# 转换时补一个标题，让 Word 里的模块编号 1-9 连续。不需要就传 --no-module1
MODULE1_TITLE = "1. 项目卡片"


class El:
    """极简 DOM 节点。children 里既有 El 也有 str（直接文本）。"""
    __slots__ = ("tag", "classes", "children")

    def __init__(self, tag, attrs=None):
        self.tag = tag
        self.classes = set((attrs or {}).get("class", "").split())
        self.children = []

    def has(self, *names):
        return any(n in self.classes for n in names)

    def text(self):
        """递归取出全部纯文本（去重空白）。"""
        parts = []
        for c in self.children:
            if isinstance(c, str):
                parts.append(c)
            else:
                parts.append(c.text())
        return "".join(parts)

    def walk(self):
        for c in self.children:
            if isinstance(c, El):
                yield c
                yield from c.walk()


class MiniHTML(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = El("root")
        self.stack = [self.root]

    def handle_starttag(self, tag, attrs):
        el = El(tag, dict(attrs))
        self.stack[-1].children.append(el)
        if tag not in VOID_TAGS:
            self.stack.append(el)

    def handle_startendtag(self, tag, attrs):
        self.stack[-1].children.append(El(tag, dict(attrs)))

    def handle_endtag(self, tag):
        for i in range(len(self.stack) - 1, 0, -1):
            if self.stack[i].tag == tag:
                del self.stack[i:]
                return

    def handle_data(self, data):
        if self.stack[-1].children and isinstance(self.stack[-1].children[-1], str):
            self.stack[-1].children[-1] += data
        else:
            self.stack[-1].children.append(data)


def norm(s):
    """压缩空白，保留首尾语义。"""
    return " ".join(html.unescape(s).split())


# ---------------------------------------------------------------- 内联收集

def is_punct(ch):
    """Unicode 标点类（P*）：中英文标点都算，emoji 不算（属 So）。"""
    return unicodedata.category(ch).startswith("P")


def need_space(prev, nxt):
    """相邻元素之间要不要补空格：PDF 靠 flex 间距，Word 里得显式加。"""
    if not prev or not nxt:
        return False
    if prev[-1] == " " or nxt[0] == " ":
        return False
    if is_punct(prev[-1]) or is_punct(nxt[0]):
        return False
    return True


def collect_paragraphs(el, bold=False):
    """
    把一个元素的内容收成若干段落，每段是 [(text, bold), ...]。
    遇到 <br> 或带 MARK_CLASS 的元素就另起一段。
    """
    paras = [[]]

    def emit(text, is_bold):
        text = norm(text)
        if text:
            if paras[-1]:
                prev, pb = paras[-1][-1]
                if need_space(prev, text):
                    paras[-1][-1] = (prev + " ", pb)
            paras[-1].append((text, is_bold))

    def rec(node, is_bold):
        for c in node.children:
            if isinstance(c, str):
                emit(c, is_bold)
                continue
            tag = c.tag
            if tag == "br":
                paras.append([])
                continue
            marks = [k for k in MARK_CLASS if k in c.classes]
            if marks:
                pre, suf = MARK_CLASS[marks[0]]
                inner = norm(c.text())
                if inner:
                    paras.append([])
                    paras[-1].append((pre + inner + suf, True))
                    paras.append([])
                continue
            if tag in ("strong", "b"):
                rec(c, True)
            elif tag in ("em", "i"):
                rec(c, is_bold)
            elif tag in ("span", "small", "code", "a", "font", "sub", "sup"):
                rec(c, is_bold)
            else:
                # div/p 等块级出现在内联位置，也当作分段
                inner = collect_paragraphs(c, is_bold)
                if len(inner) > 1 or (inner and inner[0]):
                    paras.append([])
                    for p in inner:
                        if p:
                            paras.append(p)
                            paras.append([])
        return

    rec(el, bold)

    # 收尾：去掉空段，以及只有符号的段（.ref-thumb 这类纯 emoji 缩略图占位，
    # 在 PDF 里是个方块，搬到 Word 里没有意义）
    def has_content(p):
        return any(unicodedata.category(ch)[0] in ("L", "N")
                   for text, _ in p for ch in text)

    return [p for p in paras if p and has_content(p)]


# ---------------------------------------------------------------- 块级转换

def build_blocks(el, blocks):
    """按预览稿结构输出 block。block 是 dict。"""
    for c in el.children:
        if isinstance(c, str):
            continue

        # --- 模块：编号 + 标题 → Heading2 ---
        if c.has("module"):
            for sub in c.children:
                if isinstance(sub, El) and sub.has("module-header"):
                    num = ""
                    title = ""
                    for s2 in sub.children:
                        if not isinstance(s2, El):
                            continue
                        if s2.has("module-num"):
                            num = norm(s2.text())
                        elif s2.has("module-title"):
                            title = norm(s2.text())
                    head = f"{num}. {title}" if num and num != "·" else title
                    if head:
                        blocks.append({"type": "h2", "runs": [(head, True)]})
                    break
            build_blocks(c, blocks)
            continue

        if c.has("module-header", "module-num", "module-title", "module-sub"):
            continue  # 标题已由父级输出

        # --- 表格 → Word 真表格 ---
        if c.tag == "table":
            blocks.append({"type": "table", "rows": table_rows(c)})
            continue

        # --- 流程图 → 一行箭头文字 ---
        if c.has("flow"):
            steps = []
            for s in c.children:
                if isinstance(s, El) and (s.has("flow-step") or s.tag == "div"):
                    t = norm(s.text())
                    if t:
                        steps.append(t)
            if steps:
                blocks.append({"type": "p", "runs": [(" → ".join(steps), False)]})
            continue

        # --- 待确认项 → 可勾选清单 ---
        if c.has("check-item"):
            box = "☐ "
            text_parts = []
            tag_text = ""
            for s in c.children:
                if not isinstance(s, El):
                    continue
                if s.has("check-box"):
                    continue
                if s.has("check-tag"):
                    tag_text = norm(s.text())
                else:
                    t = norm(s.text())
                    if t:
                        text_parts.append(t)
            line = box + "".join(text_parts)
            if tag_text:
                line += f"　［{tag_text}］"
            if line.strip():
                blocks.append({"type": "check", "runs": [(line, False)]})
            continue

        # --- 角色卡 → Heading3 + 说明 ---
        if c.has("role-card"):
            name = ""
            rest = []
            for s in c.children:
                if isinstance(s, El) and s.has("role-name"):
                    name = norm(s.text())
                else:
                    for p in collect_paragraphs(s) if isinstance(s, El) else []:
                        if p:
                            rest.append(p)
            if name:
                blocks.append({"type": "h3", "runs": [(name, True)]})
            for p in rest:
                blocks.append({"type": "p", "runs": p})
            continue

        # --- 标题标签 ---
        if c.tag in H_TAGS:
            t = norm(c.text())
            if t:
                # type 用标签名（h1/h2…），渲染层再映射到 Word 样式
                blocks.append({"type": c.tag, "runs": [(t, True)]})
                # 项目卡片（模块 1）在 HTML 里是 hero 区，没有 .module 容器，
                # 这里在大标题后补一个标题，保证 Word 里 1-9 的模块编号连续
                if c.tag == "h1" and MODULE1_TITLE and not any(
                        b.get("runs") and b["runs"][0][0] == MODULE1_TITLE for b in blocks):
                    blocks.append({"type": "h2", "runs": [(MODULE1_TITLE, True)]})
            continue

        # --- 普通段落容器 ---
        if c.tag in ("p", "div", "section", "li", "ul", "ol"):
            paras = collect_paragraphs(c)
            # 容器本身可能还有子块（比如 div 里套 table）
            has_block_child = any(
                isinstance(s, El) and (
                    s.tag == "table" or s.tag in H_TAGS
                    or s.has("module", "flow", "check-item", "role-card", "benefit-card")
                )
                for s in c.children
            )
            if paras and not has_block_child:
                for p in paras:
                    blocks.append({"type": "p", "runs": p})
            elif has_block_child:
                build_blocks(c, blocks)
            if c.tag == "li" and paras:
                continue
            continue

        # 其它标签：递归下去
        build_blocks(c, blocks)


def table_rows(table):
    rows = []
    ncols = max((len([x for x in tr.children if isinstance(x, El) and x.tag in ("td", "th")])
                 for tr in table.children if isinstance(tr, El) and tr.tag == "tr"), default=1)
    col_w = max(int(BODY_WIDTH_DXA / max(ncols, 1)), 800)

    for tr in table.children:
        if not isinstance(tr, El) or tr.tag != "tr":
            continue
        is_head = False
        cells = []
        for tc in tr.children:
            if not isinstance(tc, El) or tc.tag not in ("td", "th"):
                continue
            if tc.tag == "th":
                is_head = True
            paras = collect_paragraphs(tc)
            cells.append(paras if paras else [[("", False)]])
        if cells:
            rows.append({"cells": cells, "head": is_head, "col_w": col_w})
    return rows


# ---------------------------------------------------------------- XML 生成

def p_xml(runs, style=None, indent=False):
    parts = ['<w:p>']
    pr = []
    if style:
        pr.append(f'<w:pStyle w:val="{style}"/>')
    if indent:
        pr.append('<w:ind w:left="360"/>')
    if pr:
        parts.append("<w:pPr>" + "".join(pr) + "</w:pPr>")
    for text, bold in runs:
        rpr = "<w:rPr>"
        rpr += (f'<w:rFonts w:ascii="{FONT_ASCII}" w:hAnsi="{FONT_ASCII}" w:eastAsia="{FONT_EA}"/>')
        if bold:
            rpr += "<w:b/>"
        rpr += "</w:rPr>"
        parts.append(f'<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>')
    parts.append("</w:p>")
    return "".join(parts)


def table_xml(row_specs):
    borders = "".join(
        f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="{COLOR_BORDER}"/>'
        for side in ("top", "left", "bottom", "right", "insideH", "insideV")
    )
    ncols = max((len(r["cells"]) for r in row_specs), default=1)
    col_w = row_specs[0]["col_w"] if row_specs else BODY_WIDTH_DXA
    grid = "".join(f'<w:gridCol w:w="{col_w}"/>' for _ in range(ncols))

    out = ["<w:tbl>",
           "<w:tblPr>",
           f'<w:tblW w:w="{BODY_WIDTH_DXA}" w:type="dxa"/>',
           f"<w:tblBorders>{borders}</w:tblBorders>",
           '<w:tblLayout w:type="fixed"/>',
           "</w:tblPr>",
           f"<w:tblGrid>{grid}</w:tblGrid>"]

    for ri, row in enumerate(row_specs):
        out.append("<w:tr>")
        for ci, paras in enumerate(row["cells"]):
            fill = ""
            if row["head"]:
                fill = COLOR_HEAD_FILL
            elif ci == 0:
                fill = COLOR_LABEL_FILL
            shd = (f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else "")
            out.append(
                "<w:tc>"
                f'<w:tcPr><w:tcW w:w="{row["col_w"]}" w:type="dxa"/>{shd}'
                '<w:vAlign w:val="center"/></w:tcPr>'
            )
            body = paras if paras else [[("", False)]]
            for p in body:
                out.append(p_xml(p, indent=False))
            out.append("</w:tc>")
        out.append("</w:tr>")

    out.append("</w:tbl>")
    out.append(p_xml([("", False)]))  # 表格后必须跟一个段落
    return "".join(out)


def document_xml(blocks, note):
    body = []
    if note:
        body.append(p_xml([(note, False)], style=None))
    for b in blocks:
        t = b["type"]
        if t == "table":
            body.append(table_xml(b["rows"]))
        elif t == "check":
            body.append(p_xml(b["runs"], indent=True))
        elif t in ("h1", "h2", "h3"):
            style = {"h1": "Heading1", "h2": "Heading2", "h3": "Heading3"}[t]
            body.append(p_xml(b["runs"], style=style))
        else:
            body.append(p_xml(b["runs"]))

    sect = ('<w:sectPr>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
            ' w:header="851" w:footer="992" w:gutter="0"/>'
            '</w:sectPr>')

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}"><w:body>'
        + "".join(body)
        + sect
        + "</w:body></w:document>"
    )


def styles_xml():
    def style(sid, name, size, bold, color, outline=None):
        rpr = f'<w:rFonts w:ascii="{FONT_ASCII}" w:hAnsi="{FONT_ASCII}" w:eastAsia="{FONT_EA}"/>'
        if bold:
            rpr += "<w:b/>"
        if color:
            rpr += f'<w:color w:val="{color}"/>'
        rpr += f'<w:sz w:val="{size}"/>'
        ppr = f'<w:outlineLvl w:val="{outline}"/>' if outline is not None else ""
        return (f'<w:style w:type="paragraph" w:styleId="{sid}">'
                f'<w:name w:val="{name}"/><w:basedOn w:val="Normal"/>'
                f'<w:pPr>{ppr}</w:pPr><w:rPr>{rpr}</w:rPr></w:style>')

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{W}">'
        "<w:docDefaults>"
        "<w:rPrDefault><w:rPr>"
        f'<w:rFonts w:ascii="{FONT_ASCII}" w:hAnsi="{FONT_ASCII}" w:eastAsia="{FONT_EA}"/>'
        '<w:sz w:val="21"/>'
        "</w:rPr></w:rPrDefault>"
        "<w:pPrDefault/>"
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/></w:style>'
        + style("Heading1", "heading 1", 36, True, COLOR_MAIN, 0)
        + style("Heading2", "heading 2", 28, True, COLOR_MAIN, 1)
        + style("Heading3", "heading 3", 24, True, "3E3E3E", 2)
        + "</w:styles>"
    )


CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    "</Types>"
)

RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    "</Relationships>"
)

DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    "</Relationships>"
)


def convert(input_path, output_path, note=None, module1=True):
    global MODULE1_TITLE
    if not module1:
        MODULE1_TITLE = ""

    with open(input_path, "r", encoding="utf-8") as f:
        raw = f.read()

    parser = MiniHTML()
    parser.feed(raw)

    blocks = []
    build_blocks(parser.root, blocks)
    if not blocks:
        raise SystemExit("没解析出任何内容，检查输入是不是预览稿 HTML")

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", styles_xml())
        z.writestr("word/document.xml", document_xml(blocks, note))

    return len(blocks)


def main():
    ap = argparse.ArgumentParser(description="预览稿 HTML → 可编辑 Word 工作稿（零依赖）")
    ap.add_argument("input", help="预览稿 .html")
    ap.add_argument("-o", "--output", help="输出 .docx，默认与输入同目录同名")
    ap.add_argument("--no-note", action="store_true", help="不在开头加说明行")
    ap.add_argument("--no-module1", action="store_true", help="不补「1. 项目卡片」标题")
    args = ap.parse_args()

    if not os.path.isfile(args.input):
        raise SystemExit(f"找不到输入文件: {args.input}")

    out = args.output or os.path.splitext(args.input)[0] + ".docx"
    note = None if args.no_note else "【可编辑工作稿】文字可直接修改，改完发我，我出 PDF 视觉版与制作版脚本。"
    n = convert(args.input, out, note, module1=not args.no_module1)
    print(f"✅ {out}（{n} 个内容块）")


if __name__ == "__main__":
    main()
