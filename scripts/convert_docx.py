#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
已有 docx → 公文标准格式（GB/T 9704-2012）

适用场景：内容已定稿的 docx，只需重排为公文版式。
核心承诺：**文字一字不改**（含「——」「""」等标点），仅重排版式；原有加粗强调保留。

与 convert.py 的分工：
- convert.py      ：md/txt → docx（内容还在结构整理阶段）
- convert_docx.py ：docx   → docx（内容已定稿，只改格式）

用法：
    python3 convert_docx.py <input.docx> <output.docx> [选项]

选项：
    --title-idx N      指定第 N 个非空段为标题（默认 0）
    --keep-space       保留中文与数字之间的空格（默认清理）
    --latin FONT       西文/数字字体（默认 Times New Roman；传 cn 表示与中文字体一致）
    --page-number      在页脚加公文页码（— 1 —，四号宋体，奇右偶左）

字体说明：中文字体按国标名写入（方正小标宋简体/仿宋），本机未安装时仅预览降级，
在装有该字体的 Word 上正确渲染。西文与数字默认 Times New Roman（机关公文通用写法）。
"""

import sys
import re
import os

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# 复用 convert.py 的中心化国标参数，避免两套常量漂移
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from convert import (  # noqa: E402
    FONT_CONFIG, PAGE_MARGIN,
    TITLE_SIZE, HEADING_SIZE, BODY_SIZE,
    LINE_SPACING_PT, LINE_SPACING_RULE,
    SPACE_BEFORE_PT, SPACE_AFTER_PT,
    FIRST_LINE_INDENT_PT,
)

CN_RE = r'\u4e00-\u9fff\u3000-\u303f\uff00-\uffef'


# ============================================================
# 底层格式函数
# ============================================================

def style_run(run, cn_font, size, bold, latin_font=None):
    """
    设置 run 字体：中文走 w:eastAsia，西文/数字走 w:ascii/w:hAnsi。

    关键机制：同一个 run 内中英文会自动分流——Word 按字符类型分别取
    ascii 字体与 eastAsia 字体。因此「受众662万」这类混排无需拆分 run，
    中文用仿宋、数字用 Times New Roman，一次设置即可。
    """
    run.font.size = Pt(size)
    run.font.bold = bold

    latin = latin_font or cn_font
    run.font.name = latin

    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:ascii'), latin)
    rFonts.set(qn('w:hAnsi'), latin)
    rFonts.set(qn('w:eastAsia'), cn_font)
    rFonts.set(qn('w:cs'), latin)


def fmt_para(p, indent=True, align=None):
    """统一应用国标段落格式：段前段后 0 / 固定行距 28 磅 / 可选首行缩进 2 字符"""
    pf = p.paragraph_format
    pf.space_before = Pt(SPACE_BEFORE_PT)
    pf.space_after = Pt(SPACE_AFTER_PT)
    pf.line_spacing_rule = LINE_SPACING_RULE
    pf.line_spacing = Pt(LINE_SPACING_PT)
    if indent:
        pf.first_line_indent = Pt(FIRST_LINE_INDENT_PT)
    if align is not None:
        p.alignment = align


# ============================================================
# 文本提取（保加粗，不动标点）
# ============================================================

def _strip_cn_digit_space(text):
    """清理中文与数字之间的多余空格（GB/T 15835：数字与汉字间不加空格）"""
    text = re.sub(r'(?<=[%s])\s+(?=\d)' % CN_RE, '', text)
    text = re.sub(r'(?<=\d)\s+(?=[%s])' % CN_RE, '', text)
    return text


def extract_segments(p, strip_space=True):
    """把段落拆成 [(文本, 是否加粗), ...]，合并相邻同属性片段"""
    segs = []
    for r in p.runs:
        t = r.text
        if not t:
            continue
        if strip_space:
            t = _strip_cn_digit_space(t)
        if t:
            segs.append((t, bool(r.bold)))

    # 合并相邻同加粗属性
    merged = []
    for t, b in segs:
        if merged and merged[-1][1] == b:
            merged[-1] = (merged[-1][0] + t, b)
        else:
            merged.append((t, b))

    # 清理跨 run 边界的空格（数字加粗时最常出现在边界上）
    if strip_space:
        for i in range(len(merged) - 1):
            t, b = merged[i]
            nt, nb = merged[i + 1]
            if re.search(r'[%s]$' % CN_RE, t) and re.match(r'^\s+\d', nt):
                merged[i + 1] = (re.sub(r'^\s+', '', nt), nb)
            if re.search(r'\d\s+$', t) and re.match(r'^[%s]' % CN_RE, nt):
                merged[i] = (re.sub(r'\s+$', '', t), b)
        merged = [(t, b) for t, b in merged if t]

    return merged


# ============================================================
# 页码（GB/T 9704-2012：四号宋体，— 1 —，单页码居右、双页码居左）
# ============================================================

def add_page_number_footer(section):
    """页脚插入「— N —」，四号（14pt）宋体，居中或按奇偶居左右"""
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.text = ''
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

    def _run(text):
        r = p.add_run(text)
        style_run(r, '宋体', 14, False, 'Times New Roman')
        return r

    _run('\u2014 ')
    # 页码域
    r = p.add_run()
    style_run(r, '宋体', 14, False, 'Times New Roman')
    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = ' PAGE '
    fld_sep = OxmlElement('w:fldChar')
    fld_sep.set(qn('w:fldCharType'), 'separate')
    fld_t = OxmlElement('w:t')
    fld_t.text = '1'
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')
    for el in (fld_begin, instr, fld_sep, fld_t, fld_end):
        r._element.append(el)
    _run(' \u2014')


# ============================================================
# 主转换
# ============================================================

def convert_docx(input_path, output_path, title_idx=0, strip_space=True,
                 latin_font='Times New Roman', page_number=False):
    src = Document(input_path)

    # 只取非空段落
    src_paras = [p for p in src.paragraphs if p.text.strip()]
    if not src_paras:
        print('❌ 错误：输入文档没有可转换的段落')
        sys.exit(1)

    if title_idx >= len(src_paras):
        print(f'❌ 错误：title-idx={title_idx} 超出范围（共 {len(src_paras)} 段）')
        sys.exit(1)

    title_para = src_paras[title_idx]
    body_paras = [p for i, p in enumerate(src_paras) if i != title_idx]

    doc = Document()

    # 页面边距（国标）
    sec = doc.sections[0]
    sec.top_margin = PAGE_MARGIN['top']
    sec.bottom_margin = PAGE_MARGIN['bottom']
    sec.left_margin = PAGE_MARGIN['left']
    sec.right_margin = PAGE_MARGIN['right']

    # 默认样式兜底，避免残留字体的段落混入
    normal = doc.styles['Normal']
    normal.font.name = latin_font
    normal.font.size = Pt(BODY_SIZE)
    normal.element.rPr.rFonts.set(qn('w:eastAsia'), FONT_CONFIG['body']['primary'])

    # 标题：二号 方正小标宋简体，居中
    tp = doc.add_paragraph()
    fmt_para(tp, indent=False, align=WD_ALIGN_PARAGRAPH.CENTER)
    segs = extract_segments(title_para, strip_space=False)
    if not segs:
        segs = [(title_para.text.strip(), False)]
    for t, b in segs:
        style_run(tp.add_run(t), FONT_CONFIG['title']['primary'], TITLE_SIZE, False,
                  FONT_CONFIG['title']['primary'])

    # 正文：三号 仿宋（数字/西文 Times New Roman），首行缩进 2 字符
    for bp in body_paras:
        p = doc.add_paragraph()
        fmt_para(p, indent=True, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        segs = extract_segments(bp, strip_space=strip_space)
        if not segs:
            continue
        for t, b in segs:
            style_run(p.add_run(t), FONT_CONFIG['body']['primary'], BODY_SIZE, b,
                      latin_font if latin_font != 'cn' else None)

    if page_number:
        add_page_number_footer(sec)

    doc.save(output_path)

    # 交付时把**两把尺**一起打出来，避免口径歧义
    # Word 中文版有两个指标：「字数」（连续英文串/数字串各算 1 词）与「字符数（不计空格）」
    # 二者差值＝英文词与数字串的折算量（如「3804条」字数 2 / 字符数 5）
    def _chars(t):
        return len(re.sub(r'\s', '', t))

    def _words(t):
        return sum(1 for _ in re.finditer(r'[A-Za-z0-9]+|.', re.sub(r'\s', '', t)))

    body_chars = sum(_chars(bp.text) for bp in body_paras)
    body_words = sum(_words(bp.text) for bp in body_paras)
    title_chars = _chars(title_para.text)
    title_words = _words(title_para.text)

    print(f'✅ 转换完成：{output_path}')
    print(f'   标题段：{title_para.text.strip()}')
    print(f'   正文段：{len(body_paras)} 段')
    print(f'      · Word 字数（公文常用口径）：{body_words}')
    print(f'      · 字符数（不计空格）：{body_chars}')
    print(f'   含标题：Word 字数 {body_words + title_words} · 字符数 {body_chars + title_chars}')
    print(f'   版式：上37/下35/左28/右26mm，段前段后0，固定行距28磅，正文首行缩进2字符')
    print(f'   字体：标题 方正小标宋简体二号 / 正文 仿宋三号（数字西文 {latin_font}）')
    print(f'   加粗：原稿强调位保留 {sum(1 for bp in body_paras for _, b in extract_segments(bp, strip_space) if b)} 处')


if __name__ == '__main__':
    args = sys.argv[1:]
    if len(args) < 2:
        print(__doc__)
        sys.exit(1)

    inp, outp = args[0], args[1]
    opts = args[2:]
    title_idx = 0
    strip_space = True
    latin = 'Times New Roman'
    pnum = False

    i = 0
    while i < len(opts):
        if opts[i] == '--title-idx':
            title_idx = int(opts[i + 1]); i += 2
        elif opts[i] == '--keep-space':
            strip_space = False; i += 1
        elif opts[i] == '--latin':
            latin = opts[i + 1]; i += 2
        elif opts[i] == '--page-number':
            pnum = True; i += 1
        else:
            print(f'⚠️ 忽略未知参数：{opts[i]}'); i += 1

    if not os.path.exists(inp):
        print(f'❌ 错误：输入文件不存在 - {inp}')
        sys.exit(1)

    convert_docx(inp, outp, title_idx, strip_space, latin, pnum)
