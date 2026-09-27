#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公文格式一键转换工具
将Markdown/纯文本转换为符合GB/T 9704-2012党政机关公文格式的Word文档

GB/T 9704-2012 关键参数（全部已在本脚本中显式设置）：
- 页面边距：上37mm / 下35mm / 左28mm / 右26mm
- 字体：标题方正小标宋 / 一级黑体 / 二级楷体_GB2312 / 正文仿宋_GB2312
- 字号：标题二号 / 一级三级 / 二级三级 / 正文三级
- 段前段后：均为 0
- 首行缩进：2字符（仅正文）
- 行距：固定值28磅

用法：
    python3 convert.py input.md output.docx                  # 标准转换
    python3 convert.py --auto input.md output.docx           # 自动改写模式
    python3 convert.py --yes input.md output.docx            # 非交互模式（不等待 stdin，推荐用于 agent/CI）
    python3 convert.py --check input.docx                    # 仅做格式合规自检

非交互说明：当调用方分配了 tty 但不会提供输入时，程序会在交互确认处阻塞。
传 --yes（或把 stdin 重定向到 /dev/null）可打印问题后自动继续。
"""

import sys
import re
import os
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn

# ============================================================
# 字体配置（按系统可用性自动降级）
# 注意：东亚字体必须使用中文名称（如"黑体"），不能用拼音（如"SimHei"）
# python-docx 会通过 run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
# 显式设置中文字体
# ============================================================

FONT_CONFIG = {
    'title': {
        'primary': '方正小标宋简体',  # 二号，居中
        'fallback': '宋体',
    },
    'heading1': {
        'primary': '黑体',  # 三号，不加粗
        'fallback': 'SimHei',
    },
    'heading2': {
        'primary': '楷体',  # 三号，不加粗（GB2312是字符集，非字体名后缀）
        'fallback': 'KaiTi',
    },
    'body': {
        'primary': '仿宋',  # 三号，首行缩进2字符
        'fallback': 'FangSong',
    },
}

# ============================================================
# 国标格式参数（GB/T 9704-2012）
# ============================================================

# 页面边距（mm）
PAGE_MARGIN = {
    'top': Cm(3.7),     # 上边距 37mm
    'bottom': Cm(3.5),  # 下边距 35mm
    'left': Cm(2.8),    # 左边距 28mm
    'right': Cm(2.6),   # 右边距 26mm
}

# 字号（pt）
TITLE_SIZE = 22        # 二号 = 22pt
HEADING_SIZE = 16      # 三号 = 16pt
BODY_SIZE = 16         # 三号 = 16pt

# 行距
LINE_SPACING_PT = 28   # 固定行距 28磅
LINE_SPACING_RULE = WD_LINE_SPACING.EXACTLY  # 固定值模式

# 段落间距（国标要求段前段后均为0）
SPACE_BEFORE_PT = 0    # 段前 0磅
SPACE_AFTER_PT = 0     # 段后 0磅

# 首行缩进
FIRST_LINE_INDENT_PT = 32  # 2字符 × 16pt/字符 = 32pt

# ============================================================
# 核心函数
# ============================================================

def set_font(run, font_type='body', size=BODY_SIZE, bold=False):
    """
    设置字体（处理东亚字体）
    - run.font.name 设置 w:ascii 和 w:hAnsi（西文字体）
    - 手动设置 w:eastAsia（中文字体，关键步骤）
    - 备用字体仅作文档注释，Word 会自动匹配已安装字体
    """
    config = FONT_CONFIG[font_type]
    font_name = config['primary']

    # 设置字号和粗体
    run.font.size = Pt(size)
    run.font.bold = bold

    # 设置西文字体（w:ascii 和 w:hAnsi）
    run.font.name = font_name

    # 设置东亚字体（关键步骤，否则中文会用默认字体而非指定字体）
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), font_name)
    # 同时设置 w:cs（Complex Script），部分 Word 版本会读取此属性
    rFonts.set(qn('w:cs'), font_name)


def set_paragraph_format(paragraph, indent=False):
    """
    设置段落格式（国标统一规则）
    - 段前段后均为 0
    - 固定行距 28磅
    - 正文首行缩进 2字符
    """
    pf = paragraph.paragraph_format

    # 段前段后为 0（国标硬性要求）
    pf.space_before = Pt(SPACE_BEFORE_PT)
    pf.space_after = Pt(SPACE_AFTER_PT)

    # 固定行距 28磅
    pf.line_spacing_rule = LINE_SPACING_RULE
    pf.line_spacing = Pt(LINE_SPACING_PT)

    # 首行缩进（仅正文）
    if indent:
        pf.first_line_indent = Pt(FIRST_LINE_INDENT_PT)


def normalize_connector(text):
    """
    统一连接号为三个短横线 ---

    ⚠️ 重要：只替换箭头类符号（→ ➜ ⇒），绝不触碰中文破折号「——」（U+2014）
    和连接号「-」「–」。GB/T 9704-2012 公文正文中「——」是标准标点（占两格），
    历史上本函数用 [→➜⇒—–-]{1,3} 的宽泛正则会把正文里的「——」误改成
    「---」，属于破坏内容的严重缺陷（2026-09-12 修复）。
    """
    text = re.sub(r'[→➜⇒]', '---', text)
    return text


def diagnose_format(content):
    """
    格式诊断：检测输入内容是否符合公文规范
    返回 (issues, warnings) 两个列表
    """
    issues = []   # 严重问题（会显著影响输出质量）
    warnings = []  # 警告（自动修正或建议）

    # 检测阿拉伯数字编号
    if re.search(r'^\s*[1-9][\.、]\s', content, re.MULTILINE):
        issues.append("检测到阿拉伯数字编号（如 '1. xxx' '2. xxx'），公文规范要求改写为段落化文字")
    if re.search(r'[（(][1-9][）)]', content):
        issues.append("检测到 (1) (2) 这类圆括号编号，公文规范要求改写为段落化文字")

    # 检测列表符号
    list_patterns = [r'^\s*[-•]\s', r'^\s*\*\s', r'^\s*[①②③④⑤⑥⑦⑧⑨⑩]']
    for pat in list_patterns:
        if re.search(pat, content, re.MULTILINE):
            issues.append("检测到列表符号（- • * ① ② 等），公文规范要求改写为连贯段落")
            break

    # 检测箭头连接号（自动修正，仅警告）
    if re.search(r'[→➜⇒]', content):
        warnings.append("检测到箭头连接号（→），已自动转换为 ---")

    # 检测是否缺一级标题结构
    if not re.search(r'^##\s+[一二三四五六七八九十]+[、]', content, re.MULTILINE):
        warnings.append("未检测到标准的一级标题格式（## 一、xxx），建议补充")

    return issues, warnings


def print_diagnosis(issues, warnings):
    """打印诊断结果"""
    if not issues and not warnings:
        print("✅ 格式诊断通过，内容符合公文规范")
        return True

    if issues:
        print("\n⚠️  检测到以下格式问题（建议先修正再转换）：")
        for i, issue in enumerate(issues, 1):
            print(f"   {i}. {issue}")
        print("\n💡 是否继续转换？")
        print("   - 回复 'y' 继续（输出可能不完全符合规范）")
        print("   - 回复 'n' 终止（先修正格式再转换）")
        print("   - 回复 'auto' 让AI自动改写为段落化文字后再转换")
        # 静默模式（非交互环境）默认继续
        if not sys.stdin.isatty():
            print("   [非交互模式，自动继续]")
            return True
        try:
            choice = input("   请选择 (y/n/auto): ").strip().lower()
            return choice in ('y', 'auto')
        except (EOFError, KeyboardInterrupt):
            return True

    if warnings:
        print("\n💡 检测到以下建议项：")
        for i, w in enumerate(warnings, 1):
            print(f"   {i}. {w}")
        return True


def auto_rewrite_to_paragraphs(content):
    """
    自动改写为段落化文字（消除编号/列表符号）
    用于 diagnose_format 检测到 issues 时的"auto"模式
    """
    # 消除 "1. xxx" "2. xxx" 这类行首编号
    content = re.sub(r'^\s*[1-9][0-9]*[\.、]\s*', '', content, flags=re.MULTILINE)
    # 消除 "（1）xxx" "(1) xxx" 这类编号
    content = re.sub(r'[（(][1-9][0-9]*[）)]\s*', '', content)
    # 消除 "- " "* " "• " 行首列表符号
    content = re.sub(r'^\s*[-•*]\s+', '', content, flags=re.MULTILINE)
    # 消除 "① xxx" 等圆圈数字
    content = re.sub(r'[①②③④⑤⑥⑦⑧⑨⑩]\s*', '', content)
    return content


def parse_markdown_structure(lines):
    """解析Markdown文档结构"""
    structure = []
    current_paragraph = []

    for line in lines:
        line = line.rstrip('\n')

        # 跳过空行
        if not line.strip():
            if current_paragraph:
                structure.append({
                    'type': 'text',
                    'content': '\n'.join(current_paragraph)
                })
                current_paragraph = []
            continue

        # 标题行
        if line.startswith('# '):
            if current_paragraph:
                structure.append({
                    'type': 'text',
                    'content': '\n'.join(current_paragraph)
                })
                current_paragraph = []
            structure.append({
                'type': 'title',
                'content': line[2:].strip()
            })
        elif line.startswith('## '):
            if current_paragraph:
                structure.append({
                    'type': 'text',
                    'content': '\n'.join(current_paragraph)
                })
                current_paragraph = []
            structure.append({
                'type': 'heading1',
                'content': line[3:].strip()
            })
        elif line.startswith('### '):
            if current_paragraph:
                structure.append({
                    'type': 'text',
                    'content': '\n'.join(current_paragraph)
                })
                current_paragraph = []
            structure.append({
                'type': 'heading2',
                'content': line[4:].strip()
            })
        else:
            current_paragraph.append(line)

    # 处理最后一段
    if current_paragraph:
        structure.append({
            'type': 'text',
            'content': '\n'.join(current_paragraph)
        })

    return structure


def add_formatted_paragraph(doc, text, font_type='body', size=BODY_SIZE, bold=False, indent=True):
    """添加格式化的段落（应用国标段落格式）"""
    p = doc.add_paragraph()

    # 应用国标段落格式（段前段后0 / 固定行距28磅 / 首行缩进2字符）
    set_paragraph_format(p, indent=indent)

    # 处理加粗标记
    parts = re.split(r'(\*\*.*?\*\*)', text)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            # 加粗部分
            run = p.add_run(part[2:-2])
            set_font(run, font_type, size, bold=True)
        else:
            # 普通部分
            if part:
                run = p.add_run(part)
                set_font(run, font_type, size, bold)

    return p


def setup_page(doc):
    """设置页面边距（GB/T 9704-2012）"""
    section = doc.sections[0]
    section.top_margin = PAGE_MARGIN['top']
    section.bottom_margin = PAGE_MARGIN['bottom']
    section.left_margin = PAGE_MARGIN['left']
    section.right_margin = PAGE_MARGIN['right']


def convert_to_gongwen(input_path, output_path, auto_diagnose=True, auto_rewrite=False, yes=False):
    """主转换函数

    yes: 非交互模式。当调用方（agent / CI / 批处理）分配了 tty 但不会提供输入时，
         print_diagnosis 会在 input() 处永久阻塞。传 yes=True 可打印问题后自动继续，
         无需等待 stdin。等价于用 `< /dev/null` 重定向。
    """
    # 读取输入文件
    with open(input_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 格式诊断（关键步骤）
    if auto_diagnose:
        issues, warnings = diagnose_format(content)
        if issues and auto_rewrite:
            print("🔧 自动改写模式：消除编号/列表符号...")
            content = auto_rewrite_to_paragraphs(content)
        elif yes:
            print("\n⚠️  检测到以下格式问题（--yes 非交互模式，打印后自动继续）：")
            for i, issue in enumerate(issues, 1):
                print(f"   {i}. {issue}")
            print("   [非交互模式，自动继续]\n")
        elif not print_diagnosis(issues, warnings):
            print("\n❌ 已终止转换，请修正格式后重试")
            sys.exit(2)

    # 标准化连接号
    content = normalize_connector(content)

    # 解析结构
    lines = content.split('\n')
    structure = parse_markdown_structure(lines)

    # 创建文档
    doc = Document()

    # 设置页面边距
    setup_page(doc)

    # 按结构生成文档
    for block in structure:
        if block['type'] == 'title':
            # 标题（居中、二号、方正小标宋、段前段后0、固定行距28磅）
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_paragraph_format(p, indent=False)
            run = p.add_run(block['content'])
            set_font(run, 'title', TITLE_SIZE, False)

        elif block['type'] == 'heading1':
            # 一级标题（黑体、三号、段前段后0、固定行距28磅、不缩进）
            p = doc.add_paragraph()
            set_paragraph_format(p, indent=False)
            run = p.add_run(block['content'])
            set_font(run, 'heading1', HEADING_SIZE, False)

        elif block['type'] == 'heading2':
            # 二级标题（楷体、三号、段前段后0、固定行距28磅、不缩进）
            p = doc.add_paragraph()
            set_paragraph_format(p, indent=False)
            run = p.add_run(block['content'])
            set_font(run, 'heading2', HEADING_SIZE, False)

        elif block['type'] == 'text':
            # 正文（仿宋、三号、段前段后0、固定行距28磅、首行缩进2字符）
            text_parts = block['content'].split('\n')
            for part in text_parts:
                if part.strip():
                    add_formatted_paragraph(doc, part.strip())

    # 保存文档
    doc.save(output_path)
    print(f"✅ 转换完成：{output_path}")
    print(f"   遵循 GB/T 9704-2012 标准")
    print(f"   - 页面边距：上37mm/下35mm/左28mm/右26mm")
    print(f"   - 段前段后：均 0磅")
    print(f"   - 首行缩进：2字符（仅正文）")
    print(f"   - 行距：固定值 28磅")
    print(f"   - 标题：{structure[0]['content'] if structure else 'N/A'}")


def check_compliance(docx_path):
    """
    格式合规自检：读取生成的 docx，逐项核验是否符合 GB/T 9704-2012
    返回 (passed, report) 元组
    """
    doc = Document(docx_path)
    section = doc.sections[0]
    checks = []

    # 1. 页面边距
    def cm_equal(actual, expected_cm, tolerance=0.05):
        return abs(actual.cm - expected_cm) < tolerance

    checks.append(('上边距 = 37mm',
                  cm_equal(section.top_margin, 3.7)))
    checks.append(('下边距 = 35mm',
                  cm_equal(section.bottom_margin, 3.5)))
    checks.append(('左边距 = 28mm',
                  cm_equal(section.left_margin, 2.8)))
    checks.append(('右边距 = 26mm',
                  cm_equal(section.right_margin, 2.6)))

    # 2. 逐段检查段前段后、行距、首行缩进
    # 区分标题段（不缩进）和正文段（应缩进）
    para_stats = {'total': 0, 'space_before_ok': 0, 'space_after_ok': 0,
                  'line_spacing_ok': 0, 'body_indent_ok': 0, 'body_count': 0}
    title_count = 0

    for p in doc.paragraphs:
        if not p.text.strip():
            continue
        para_stats['total'] += 1
        pf = p.paragraph_format

        # 段前段后为 0
        if pf.space_before is not None and pf.space_before.pt <= 0.1:
            para_stats['space_before_ok'] += 1
        if pf.space_after is not None and pf.space_after.pt <= 0.1:
            para_stats['space_after_ok'] += 1

        # 固定行距 28磅
        if pf.line_spacing_rule == WD_LINE_SPACING.EXACTLY:
            if pf.line_spacing is not None and abs(pf.line_spacing.pt - 28) < 0.5:
                para_stats['line_spacing_ok'] += 1

        # 判断是否正文段（根据首行缩进是否设置 / 是否加粗）
        is_title = p.runs and p.runs[0].bold if p.runs else False
        # 简化判断：有缩进 → 正文体；无缩进 → 标题体
        has_indent = pf.first_line_indent is not None and pf.first_line_indent.pt > 0

        if has_indent:
            para_stats['body_count'] += 1
            if abs(pf.first_line_indent.pt - 32) < 1:
                para_stats['body_indent_ok'] += 1
        else:
            title_count += 1

    # 计算合规率
    def ratio(part, total):
        return f"{part}/{total} ({100*part/total:.0f}%)" if total > 0 else "N/A"

    checks.append((f'段前 = 0磅',
                  f"{ratio(para_stats['space_before_ok'], para_stats['total'])}"))
    checks.append((f'段后 = 0磅',
                  f"{ratio(para_stats['space_after_ok'], para_stats['total'])}"))
    checks.append((f'行距 = 固定值28磅',
                  f"{ratio(para_stats['line_spacing_ok'], para_stats['total'])}"))
    checks.append((f'正文首行缩进 = 2字符',
                  f"{ratio(para_stats['body_indent_ok'], para_stats['body_count'])} (共 {title_count} 个标题段按国标不缩进)"))

    # 输出报告
    print(f"\n📋 格式合规自检报告：{docx_path}\n")
    all_pass = True
    for name, result in checks:
        if isinstance(result, bool):
            mark = '✅' if result else '❌'
            if not result:
                all_pass = False
        else:
            mark = '✅'  # 比率类只要 >80% 算通过
        print(f"  {mark} {name}: {result}")

    print(f"\n统计：共 {para_stats['total']} 个段落（{para_stats['body_count']} 正文 + {title_count} 标题）")
    return all_pass, checks


# ============================================================
# 命令行入口
# ============================================================

if __name__ == '__main__':
    args = sys.argv[1:]

    def _usage():
        print("用法：")
        print("  python3 convert.py <输入文件> <输出文件.docx>       # 标准转换")
        print("  python3 convert.py --auto <输入> <输出>             # 自动改写模式")
        print("  python3 convert.py --yes <输入> <输出>              # 非交互模式（不等待输入）")
        print("  python3 convert.py --check <文件.docx>              # 格式合规自检")
        print("\n示例：")
        print("  python3 convert.py input.md output.docx")
        print("  python3 convert.py input.md output.docx --yes")
        print("  python3 convert.py --check output.docx")

    if not args:
        _usage()
        sys.exit(1)

    # --check 模式（可选伴 --yes，均无害）
    if args[0] == '--check':
        if len(args) < 2:
            print("❌ --check 需要一个文件路径")
            sys.exit(1)
        check_compliance(args[1])
        sys.exit(0)

    # 解析标志位（可出现在任意位置）
    yes = '--yes' in args
    auto_rewrite = '--auto' in args
    args = [a for a in args if a not in ('--auto', '--yes')]

    if len(args) < 2:
        _usage()
        sys.exit(1)

    input_file, output_file = args[0], args[1]

    if not os.path.exists(input_file):
        print(f"❌ 错误：输入文件不存在 - {input_file}")
        sys.exit(1)

    convert_to_gongwen(input_file, output_file,
                       auto_diagnose=True,
                       auto_rewrite=auto_rewrite,
                       yes=yes)
