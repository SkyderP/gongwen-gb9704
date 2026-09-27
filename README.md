# gongwen-quick-convert

将 Markdown / 纯文本 / 已定稿 docx 一键排版为符合 **GB/T 9704-2012《党政机关公文格式》** 的 Word 文档。

这是一个 Agent Skill（SKILL.md 规范），可以把 Install 一行命令装进任意支持 Skill 的 Agent（Claude Code、WorkBuddy 等），之后说一句「按公文格式排版这段内容」即可出符合国标的 docx。

> A portable Agent Skill that converts Markdown, plain text, or finished .docx files into Word documents conforming to China's GB/T 9704-2012 state paper format standard.

---

## 为什么需要它

公文排版有几个"看不见的坑"：

- 页面边距必须是 上 37mm / 下 35mm / 左 28mm / 右 26mm，差 1mm 就不合规
- 行距是**固定值 28 磅**（不是 1.5 倍），段前段后必须 **0 磅**
- 标题用方正小标宋简体二号，一级标题黑体三号，二级标题楷体三号，正文仿宋三号
- 中文字体要通过 `w:eastAsia` 设置，只设 `ascii` 在 Word 里不生效
- 正文里的中文破折号「——」极易被连接号归一化规则**误删误改**，这是破坏内容的严重缺陷

本技能把这些参数全部集中管理、显式设置，并把踩过的坑写进文档，跨环境复用不会重蹈覆辙。

---

## 快速开始

### 1. 安装

```bash
pip install python-docx
```

把整个目录放到 Agent 的 skills 目录下：

| Agent | 安装位置 |
|---|---|
| Claude Code | `~/.claude/skills/` 或项目 `.claude/skills/` |
| WorkBuddy | `~/.workbuddy/skills/` |

保持以下结构即可：

```
gongwen-quick-convert/
├── SKILL.md
├── README.md
└── scripts/
    ├── convert.py         # md/txt → docx
    └── convert_docx.py    # docx → docx（定稿直转）
```

⚠️ `convert_docx.py` 靠 `sys.path` 动态 import `convert.py` 里的国标常量，**两个脚本必须同目录，不可拆散**。

### 2. 转换

```bash
# Markdown → 公文格式 docx
python3 scripts/convert.py input.md output.docx --yes

# 已有定稿 docx，只改版式（文字一字不改）
python3 scripts/convert_docx.py in.docx out.docx
```

`--yes` 会跳过交互确认。不加的话，当内容含阿拉伯数字编号或列表符号时，脚本会 `input()` 等待输入——在 Agent、CI、批处理这类"分配了 tty 却不提供输入"的环境里会**永久阻塞**。

### 3. 交付前自检

```bash
python3 scripts/convert.py --check output.docx
```

输出示例：

```
📋 格式合规自检报告：output.docx

  ✅ 上边距 = 37mm: True
  ✅ 下边距 = 35mm: True
  ✅ 左边距 = 28mm: True
  ✅ 右边距 = 26mm: True
  ✅ 段前 = 0磅: 15/15 (100%)
  ✅ 段后 = 0磅: 15/15 (100%)
  ✅ 行距 = 固定值28磅: 15/15 (100%)
  ✅ 正文首行缩进 = 2字符: 7/7 (100%) (共 8 个标题段按国标不缩进)

统计：共 15 个段落（7 正文 + 8 标题）
```

标题段按国标不缩进，所以分母是正文段数而非总段数。

---

## 文档结构识别规则

```markdown
# 标题          → 文档总标题（方正小标宋简体，二号，居中）
## 一、xxx      → 一级标题（黑体，三号）
### （一）xxx   → 二级标题（楷体，三号）
普通段落        → 正文（仿宋，三号，首行缩进 2 字符）
```

## 国标参数对照表

| 参数 | 标准值 |
|---|---|
| 上 / 下 / 左 / 右边距 | 37 / 35 / 28 / 26 mm |
| 标题字号（二号） | 22pt |
| 一级 / 二级 / 正文字号（三号） | 16pt |
| 行距 | 固定值 28 磅（EXACTLY） |
| 段前 / 段后 | 0 磅 |
| 正文首行缩进 | 2 字符（32pt） |
| 标题字体 | 方正小标宋简体 |
| 一级标题 | 黑体 |
| 二级标题 | 楷体 |
| 正文 | 仿宋 |

## 两种输入形态，两个脚本

| 输入 | 用哪个 | 特点 |
|---|---|---|
| 内容在 md / txt | `convert.py` | 带格式诊断 + 自动改写 |
| 内容已是定稿 docx | `convert_docx.py` | 文字一字不改，加粗位与标点全保留 |

**定稿 docx 为什么不能走 md 中转**：中文破折号「——」和引号在 Markdown 往返中易受连接号归一化影响。docx 直转可保证标点零损失。

---

## 命令行参数

### convert.py

| 参数 | 作用 |
|---|---|
| `--yes` | 非交互模式，跳过确认（推荐，agent/CI 必加） |
| `--auto` | 自动改写：消除阿拉伯数字编号与列表符号后再转换 |
| `--check` | 只做格式合规自检，不生成文件 |

### convert_docx.py

| 参数 | 作用 |
|---|---|
| `--title-idx N` | 指定标题段（默认第 1 个非空段） |
| `--keep-space` | 保留中文与数字之间的空格（默认清理） |
| `--latin cn` | 数字/西文也用中文字体（默认 Times New Roman） |
| `--page-number` | 页脚插入公文页码 — 1 — |

---

## 已知限制

- 不支持表格转换。正式公文正文本就不推荐用表格，请先改为段落化文字。
- **字体依赖**：方正小标宋简体非 Windows 自带，需单独安装；未安装时 Word 自动降级，不影响结构只影响观感。
- **腾讯文档 / WPS 等编辑器保存 docx 会漂移格式**：标题二号 22pt 常被改成 18pt，西文字体名被写成无效值「Times New Roman Regular」。凡"用户已在编辑器里改过、要继续改"的 docx，交付前必须重刷格式。
- 输入文件编码需为 UTF-8。

## 踩坑记录（跨环境复用前必读）

1. **中文破折号「——」保护**（v1.5 修复）：旧版 `normalize_connector()` 用 `[→➜⇒—–-]{1,3}` 宽泛正则，会把正文「——」误改成「---」。现只替换箭头类 `[→➜⇒]`。
2. **tty 环境阻塞**（v1.6 修复）：诊断环节检测到编号/列表符号时 `input()` 会等待；在分配 tty 但不提供输入的调用方（Agent Bash 工具、CI）中 `isatty()` 为 True，脚本被超时杀掉（实测 exit 137）。用 `--yes` 或 `< /dev/null`。**移植后跑不通，第一查这个。**
3. **`python-docx` 的 `Paragraph` 不是同一个对象**：`doc.paragraphs` 每次访问都新建包装对象，`p is saved_para` 永远为 False。判标题段必须比底层元素 `p._p is saved_para._p`。
4. **中英文同 run 分流**：同时设 `w:ascii` 与 `w:eastAsia`，混排无需拆分 run。
5. **跨 run 边界空格**：片段序列层面清理，run 内部正则抓不到。

---

## 适用场景

- 项目申报材料、政策文件的最终排版
- 领导要求按红头文件格式提交
- 已有 Markdown / 纯文本内容，需要出符合国标的 docx

## 不适用场景

- 内容还在创作阶段（先写完再排版）
- 格式严重不符合公文规范（先用格式诊断功能修正）
- 需要 HTML / PDF 等其他格式

---

## License

MIT License — 详见 [LICENSE](LICENSE)。

公文中可能含涉密或内部信息，**请仅在确认内容可公开的前提下使用本技能处理文档**。
