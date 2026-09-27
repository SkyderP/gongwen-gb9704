---
name: gongwen-quick-convert
description: |
  将 Markdown、纯文本或已有 docx 转换为符合 GB/T 9704-2012《党政机关公文格式》的 Word 文档。
  自动处理：页面边距（上37/下35/左28/右26mm）、字体字号（标题方正小标宋简体二号居中、一级标题黑体三号、二级标题楷体三号、正文仿宋三号）、段前段后 0磅、固定行距 28磅、正文首行缩进 2字符。
  适用：项目申报材料、政策文件、汇报材料、请示报告、方案总结等公文的最终排版。
  触发词：公文格式、公文排版、转换成公文格式、按国标格式排版、GB/T 9704、党政机关公文格式、红头文件、转红头文件、国标docx、公文转Word、排版成公文、排版成红头文件。
agent_created: true
---

# 公文格式一键转换技能

将 Markdown / 纯文本 / 已定稿 docx，快速转换为符合 GB/T 9704-2012 标准的 Word 文档。

## 安装说明（跨 agent 移植用）

本技能为**自包含包**：

```
gongwen-quick-convert/
├── SKILL.md                      # 本文件
└── scripts/
    ├── convert.py                # md/txt → docx（含格式诊断与自动改写）
    └── convert_docx.py           # docx → docx（定稿直转，文字一字不改）
```

安装方式：

| 目标 agent | 安装位置 | 说明 |
|---|---|---|
| Claude Code / WorkBuddy | `~/.claude/skills/` 或 `~/.workbuddy/skills/` | 用户级；或放项目 `.claude/skills/` |
| 其他支持 SKILL.md 的 agent | 其约定的 skills 根目录 | 保持上述两级目录结构即可 |

**硬性前置条件**：

1. `python-docx` 已安装：`pip install python-docx`（或 `python -m pip install python-docx`）
2. 脚本必须放在 `scripts/` 下并保持 `convert_docx.py` 与 `convert.py` **同目录**——后者靠 `sys.path` 动态导入前者的国标常量，不可拆散
3. 示例命令中的 `<skill_dir>` 需替换为本技能的实际安装根目录

## 何时使用

当你需要把已写好的内容快速排版为**正式公文格式**时使用——典型场景：

- 项目申报材料、政策文件、汇报材料的最终排版
- 已有 Markdown / 纯文本内容，需要出符合国标的 docx
- 领导要求按红头文件格式提交

⚠️ **不适用场景**：

- 内容还在创作阶段（先写完再排版）
- 内容格式严重不符合公文规范（先用"格式诊断"功能修正）
- 需要其他格式（如 HTML / PDF）→ 用其他工具

📌 **两种输入形态，走两个脚本**：

- 内容在 md / txt 里 → **方法 1–5**（`scripts/convert.py`，带格式诊断与自动改写）
- 内容已是**定稿 docx** → **方法 6**（`scripts/convert_docx.py`，docx 直转；一字不改、加粗位与标点全保留）

## 文档结构识别规则

输入文本按以下规则解析：

```
# 标题         → 文档总标题（方正小标宋简体，二号，居中）
## 一、xxx      → 一级标题（黑体，三号）
### （一）xxx   → 二级标题（楷体，三号）
普通段落        → 正文（仿宋，三号，首行缩进2字符）
```

## 国标格式参数（GB/T 9704-2012，全部显式设置）

| 参数类别 | 参数项 | 标准值 | 代码常量 |
|---------|--------|--------|----------|
| **页面边距** | 上边距 | 37mm | `Cm(3.7)` |
| | 下边距 | 35mm | `Cm(3.5)` |
| | 左边距 | 28mm | `Cm(2.8)` |
| | 右边距 | 26mm | `Cm(2.6)` |
| **字号** | 标题（二号） | 22pt | `TITLE_SIZE = 22` |
| | 一级/二级/正文（三号） | 16pt | `HEADING_SIZE / BODY_SIZE = 16` |
| **行距** | 固定值 | 28磅 | `LINE_SPACING_PT = 28` |
| | 行距规则 | EXACTLY（固定值） | `WD_LINE_SPACING.EXACTLY` |
| **段前段后** | 段前 | 0磅 | `SPACE_BEFORE_PT = 0` |
| | 段后 | 0磅 | `SPACE_AFTER_PT = 0` |
| **首行缩进** | 正文首行 | 2字符（32pt） | `FIRST_LINE_INDENT_PT = 32` |
| | 标题/一级/二级 | 不缩进 | `indent=False` |
| **字体** | 标题 | 方正小标宋简体（二号） | `'方正小标宋简体'` |
| | 一级标题 | 黑体（三号） | `'黑体'` |
| | 二级标题 | 楷体（三号） | `'楷体'` |
| | 正文 | 仿宋（三号） | `'仿宋'` |

> ⚠️ **所有参数都通过 `set_paragraph_format()` 集中应用**，避免逐段重复设置导致漏项。

## 完整工作流

### 步骤 1：接收输入

支持两种输入方式：

- **文件路径**：传入 `.md` 或 `.txt` 文件
- **直接粘贴文本**：用户对话中粘贴的内容

### 步骤 2：格式诊断（关键！先诊断再转换）

在转换之前先扫描内容，识别以下问题：

| 问题类型 | 检测规则 | 严重程度 |
|---------|---------|---------|
| 阿拉伯数字编号 | 检测 `1. ` `2. ` `（1）` `（2）` 等 | ⚠️ 中等 |
| 列表符号 | 检测行首 `-` `*` `•` `①` `②` | ⚠️ 中等 |
| 箭头连接号 | 检测 `→` `➜` `⇒` | 🔧 轻微（自动修正） |
| 缺一级标题 | 无 `## 一、` `## 二、` 等结构 | ⚠️ 中等 |
| 加粗缺失 | 关键目标句未使用 `**加粗**` 标记 | 💡 建议（不强制） |

**诊断输出**：

- 若发现问题：**先展示给用户确认**，给出「自动修正建议」
- 若无问题：直接进入步骤 3

### 步骤 3：自动修正

执行以下修正：

- 箭头连接号统一为 `---`（三个短横线）
- 移除标题前后多余空行
- 中文引号保留（不处理）

⚠️ **不自动修正**（需用户决策）：

- 阿拉伯数字编号 → 需改写为段落化文字
- 列表符号 → 需改写为段落
- 这些修改影响内容结构，由用户决定

### 步骤 4：转换生成

```bash
python3 <skill_dir>/scripts/convert.py <input> <output.docx>
```

⚠️ **务必加 `--yes`**（推荐写法）：

```bash
python3 <skill_dir>/scripts/convert.py <input> <output.docx> --yes
```

原因：当输入含阿拉伯数字编号或列表符号时，脚本默认会**交互询问是否继续**。若在分配了 tty 但不会提供输入的 agent / CI / 批处理环境中执行，会在 `input()` 处**永久阻塞**。`--yes` 会打印诊断问题后自动继续，不等待 stdin。等价写法：`< input.txt /dev/null`。

### 步骤 5：交付前检查

- [ ] 文档总标题存在且居中
- [ ] 一级标题为"一、二、三、xxx"格式
- [ ] 二级标题为"（一）（二）xxx"格式
- [ ] 正文首行缩进 2 字符
- [ ] 箭头连接号已统一为 `---`
- [ ] 中文破折号「——」**未被误改**
- [ ] 文件能正常打开（无损坏）

## 使用方法

### 方法 1：转换 Markdown 文件

```bash
python3 <skill_dir>/scripts/convert.py input.md output.docx --yes
```

### 方法 2：转换纯文本

```bash
python3 <skill_dir>/scripts/convert.py input.txt output.docx --yes
```

### 方法 3：自动改写模式（消除编号/列表符号后转换）

```bash
python3 <skill_dir>/scripts/convert.py --auto input.md output.docx
```

### 方法 3.5：非交互模式（推荐用于 agent / CI / 批处理）

```bash
python3 <skill_dir>/scripts/convert.py --yes input.md output.docx
```

`--yes` 跳过交互确认，打印诊断问题后直接转换。不加则可能在 tty 环境下阻塞。

### 方法 4：格式合规自检（转换后核验）

```bash
python3 <skill_dir>/scripts/convert.py --check output.docx
```

输出示例：

```
📋 格式合规自检报告：output.docx

  ✅ 上边距 = 37mm: True
  ✅ 下边距 = 35mm: True
  ✅ 左边距 = 28mm: True
  ✅ 右边距 = 26mm: True
  ✅ 段前 = 0磅: 12/12 (100%)
  ✅ 段后 = 0磅: 12/12 (100%)
  ✅ 行距 = 固定值28磅: 12/12 (100%)
  ✅ 首行缩进 = 2字符: 8/12 (67%)

统计：共 12 个段落
```

### 方法 5：直接传入文本内容

在对话中直接说"按公文格式排版这段内容"+ 粘贴文本，agent 会调用此技能。

### 方法 6：已有 docx → 公文格式（内容已定稿，只改版式）

```bash
python3 <skill_dir>/scripts/convert_docx.py input.docx output.docx
```

适用：内容已定稿的 docx，只需重排为公文版式。
**承诺：文字一字不改**（含「——」「""」等标点），原有加粗强调全部保留。

```bash
# 常用选项
python3 <skill_dir>/scripts/convert_docx.py in.docx out.docx --title-idx 0   # 指定标题段（默认第 1 个非空段）
python3 <skill_dir>/scripts/convert_docx.py in.docx out.docx --keep-space    # 保留中文与数字之间的空格（默认清理）
python3 <skill_dir>/scripts/convert_docx.py in.docx out.docx --latin cn      # 数字/西文也用中文字体（默认 Times New Roman）
python3 <skill_dir>/scripts/convert_docx.py in.docx out.docx --page-number   # 页脚加公文页码 — 1 —
```

转换后**必须**做零改动核验（逐字 + 加粗位比对）：

```python
import re, docx

def dump(p):
    return [(re.sub(r'\s','',x.text),
             [(re.sub(r'\s','',r.text), bool(r.bold)) for r in x.runs if r.text.strip()])
            for x in docx.Document(p).paragraphs if x.text.strip()]

a, b = dump('in.docx'), dump('out.docx')
assert len(a) == len(b) and all(x[0]==y[0] for x,y in zip(a,b)), '文字被改动！'
assert all(''.join(t for t,bo in x[1] if bo)==''.join(t for t,bo in y[1] if bo) for x,y in zip(a,b)), '加粗位被改动！'
print('零改动核验通过')
```

> **为什么不能走 convert.py 的 md 中转**：定稿正文里的中文破折号「——」和引号在 md 往返中易受连接号归一化影响。docx 直转可保证标点零损失。

## 注意事项

1. **字体依赖**：系统需安装方正小标宋简体、黑体、楷体、仿宋等字体，未安装时 Word 自动降级，不影响结构仅影响观感
2. **标题格式**：确保一级标题使用"一、二、三、"格式，二级标题使用"（一）（二）（三）"格式
3. **箭头连接号**：技能会自动将 `→` `➜` `⇒` 转换为 `---`；**中文破折号「——」绝不触碰**
4. **加粗保留**：Markdown 的 `**加粗**` 会保留为 Word 加粗格式

## 常见问题

**Q：转换后字体显示异常？**

A：中文字体名称必须与系统安装的字体名完全匹配。常见情况：

- **方正小标宋简体**：非 Windows 自带字体，需单独安装。未安装时 Word 自动降级为宋体
- **黑体/楷体/仿宋**：Windows 自带，通常无需额外安装；macOS 可能需要手动安装
- ⚠️ 字体名使用**中文名称**（如"黑体"），不要用拼音或英文名（如"SimHei"），否则 `w:eastAsia` 属性可能无法正确匹配

**Q：标题识别错误？**

A：严格按照 `## 一、xxx` 格式书写一级标题，避免使用 `## 一.xxx` 或 `## 1. xxx` 等变体。

**Q：内容中有表格怎么办？**

A：当前版本不支持表格转换。请先把表格改为段落化文字（按公文写作惯例，正式公文正文不推荐用表格）。

**Q：输出文件打不开？**

A：检查输入文件编码（必须 UTF-8）和内容是否包含特殊字符。

**Q：报 `ModuleNotFoundError: No module named 'docx'`？**

A：执行 `pip install python-docx` 安装依赖后重试。

## 踩坑记录（跨环境复用时必须知道）

- ✅ 已处理：python-docx 的东亚字体设置（使用 `_element.rPr.rFonts.set(qn('w:eastAsia'), font_name)`）
- ❗ **已修复（v1.5）**：`normalize_connector()` 原用 `[→➜⇒—–-]{1,3}` 宽泛正则，会把正文里的**中文破折号「——」误改成「---」**——这是破坏内容的严重缺陷（公文正文「——」极常见）。现改为**只替换箭头类 `[→➜⇒]`**，保留 U+2014 破折号与连接号
- ✅ 已处理：连接号统一为三个短横线 `---`（**仅对箭头符号**，不对中文破折号）
- ✅ **v1.5 新增**：`convert_docx.py` 支持「docx → 公文格式」，绕过 md 中转以保标点零损失
- ⚠️ **腾讯文档 / WPS 等编辑器保存 docx 会漂移格式**：标题二号 22pt → **18pt**、西文字体名被写成**无效值**「Times New Roman Regular」（`Regular` 是样式名不是字族名）、`eastAsia` 被改写为「仿宋_GB2312」、中文与数字间空格被**回填**。凡「用户已在编辑器里改过、要继续改」的 docx，**交付前必须重刷格式**，不能假定上一版格式还在
- ⚠️ **python-docx 的 `Paragraph` 不是同一个对象**：`doc.paragraphs` 每次访问都新建包装对象，`p is saved_para` **永远为 False**。判标题段必须比底层元素 `p._p is saved_para._p`。踩过一次：标题被误刷成正文格式（16pt + 缩进 32pt + 两端对齐）
- 💡 经验：**在同一个 docx 上做增量修改的标准动作**——① 先 `cp` 备份到 `/tmp` ② 读全文看用户改了什么 ③ 插入/修改 ④ 重刷全文格式 ⑤ 与备份做「忽略空格」的逐字 + 加粗位比对，确认原文未被碰 ⑥ 跑 `--check` 自检
- ⚠️ **tty 环境会卡死（v1.6 修复）**：诊断环节检测到编号/列表符号时会 `input()` 等待用户输入。若调用方分配了 tty 却不提供输入（agent Bash 工具、CI、批处理），`sys.stdin.isatty()` 为 **True**，脚本**永久阻塞被超时杀掉**（实测 exit 137）。修复：新增 `--yes` 标志，打印问题后自动继续；也可 `< /dev/null` 重定向。**跨 agent 移植后跑不通，第一查这个**
- 💡 经验：**一个 run 内中英文会自动分流**——同时设 `w:ascii`（西文）与 `w:eastAsia`（中文），「受众662万」这类混排无需拆分 run，数字自动走 Times New Roman
- 💡 经验：跨 run 边界的空格（如加粗数字前的空格）在 run 内部正则里抓不到，**必须在片段序列层面清理边界空格**

## 版本历史

- v1.0 (2026-06-27)：初始版本，基于普陀区融媒项目文档排版经验创建
- v1.1 (2026-06-27)：增加格式诊断步骤、明确何时使用、扩展触发词
- v1.2 (2026-06-27)：脚本层实现诊断函数 + 交互式确认
- v1.3 (2026-06-27)：**国标参数完整化**——段前段后显式设 0 磅；抽出 `set_paragraph_format()`；增加 `--check` 自检；中心化国标参数
- v1.4 (2026-06-27)：**字体名称修正**——全部改用中文名称（方正小标宋简体/黑体/楷体/仿宋）；`set_font()` 增强
- v1.5 (2026-09-12)：**docx 直转 + 破坏性缺陷修复**——修复「——」误转；新增 `convert_docx.py`；中英文同 run 分流；跨 run 边界空格清理
- v1.6 (2026-09-27)：**可移植版**——标准化 frontmatter（触发词内嵌 description）；去除本机绝对路径（改用 `<skill_dir>` 占位）；脚本归入 `scripts/`；中性化表述；新增「安装说明」与依赖前置检查
