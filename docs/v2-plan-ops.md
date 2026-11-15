# AI PPT 设计系统 · v2.0 阶段操作手册（P0–P4）

| 项 | 内容 |
|---|---|
| 版本 | 操作书 v1.0（配套 [v2-plan.md](v2-plan.md) 使用） |
| 日期 | 2026-11-15 |
| 定位 | 规划书回答"做什么、为什么"；本文回答"**具体怎么做、做到什么算完、做完拿什么证明**" |
| 判据 | 第三方（或三个月后的我）只看本文，不问你，就能把每个阶段复现出来 |
| 状态 | 等你看完点头后，按 P0 开工 |

---

## 0. 这份文档怎么用

### 0.1 与规划书的分工

| 文件 | 回答什么 | 章节 |
|---|---|---|
| [v2-plan.md](v2-plan.md) | 范围、边界、四种形态、数据模型概要、版本策略 | 0–8、10–12 |
| 本文 | **每个阶段逐条操作**：文件清单、创建顺序、命令、期望输出、判据、排错 | 2–6 |
| 共用 | 术语、硬边界、命令序列 | 规划书附录 A–D |

### 0.2 详细度自检结论（你问的那个问题）

我对规划书第 9 章做了一遍"能不能照着做"的自查，结论是**不够**。具体缺口与在本文的补位：

| 编号 | 规划书里只写了 | 差在哪 | 本文补在哪 |
|---|---|---|---|
| G1 | "6 页 A4，含封面两章一表一图目录" | 没写这 6 页逐页放什么、每块多少个字 | §2.6 |
| G2 | "`build_docx.py`（仅六种块）" | 没写函数划分、样式映射、每步产出 | §2.3、§2.4 |
| G3 | "`doc.schema.json`：12 种块" | 只给了 5 个字段示例，没有逐块完整字段与默认值 | §2.5 |
| G4 | "验收：docx >30 KB、页数 5–7" | 没写用什么命令量、量不出来怎么办 | §2.7 |
| G5 | "30 页长文档" | 没给大纲，也没说 30 页从哪来 | §3.2 |
| G6 | "12 组 tokens 变体" | 没说 12 组是哪 12 组 | §3.3 |
| G7 | "表格跨页表头重复" | 没给 OOXML 片段与验证方法 | §3.4 |
| G8 | "`--preset` 贯穿 build 与 QA" | **与"不改 `build_pptx.py` 行为"的红线冲突**，必须给出解法 | §4.2 |
| G9 | "SVG 文字仍是文字" | 没写 SVG 不支持自动换行，需预排版 | §4.4 |
| G10 | "`layers.json` 记录坐标与层级" | 没给结构 | §4.5 |
| G11 | "`rebuild.json` 每页 keep/mask/overlay/recolor" | 没给字段与算法，色彩替换的判定也没定义 | §5.3、§5.4 |
| G12 | "案例：技术报告 / 研究摘要 / 一页纸" | 没给大纲与"真实感"要求 | §6.2 |

另外发现**两处自相矛盾**，本文直接改正（不是留给你发现）：

1. 规划书附录 C2 写了 `build_pptx.py --preset print-A4`，但该脚本现在没有 `--preset` 参数，而 4.3 又规定"不改 `build_pptx.py` 行为"。解法见 §4.2：**加可选参数，不带该参数时输出与现在逐字节一致**。
2. 规划书 A2 写"docx 与 pdf 页数差 ≤1"，但 `docx` 的页数在没打开 Word 前拿不到（本机无 Word）。解法见 §3.6：**以 PDF 页数为准，校验"目录最后一条页码 = PDF 实际页数"，再人工抽查 3 处**。

### 0.3 每个阶段的统一结构

```text
阶段目标与完成定义 → 前置检查（3 条命令）→ 文件清单与创建顺序 →
步骤 S1..Sn（动作 / 命令 / 期望输出 / 失败处理）→ 验收勾选表（含证据文件）→ 排错表
```

---

## 1. 通用约定

### 1.1 路径与命令前缀

```bash
# 仓库根
cd C:\Users\h2965\AppData\Roaming\dsh-desktop\harness\projects\ai-ppt-design-system

# Python（本机唯一可用解释器，不许 pip install）
set PY=%LOCALAPPDATA%\..\..\Software\DSH_Desktop\DSH Desktop\resources\office-runtime\primary-runtime\dependencies\python\python.exe
# 实际使用下面的绝对路径
set PY=D:\Software\DSH_Desktop\DSH Desktop\resources\office-runtime\primary-runtime\dependencies\python\python.exe
set NODE=C:\Users\h2965\AppData\Roaming\dsh-desktop\harness\.desktop-bin\node.cmd
set CLI=D:\Software\DSH_Desktop\DSH Desktop\resources\office-cli.mjs
```

### 1.2 产物落盘位置（所有阶段统一）

| 目录 | 放什么 | 是否进 git |
|---|---|---|
| `records/` | 需求确认书、简报、提案、大纲、逐页/逐块规格、QA 报告 | 是 |
| `source/` | `deck.json` / `doc.json` / `rebuild.json` + `assets/` | 是 |
| `build/` | `.pptx`、`.docx`、`.pdf`、`preview/`、`qa.md`、`qa.json`、`*.manifest.json` | 否（`.gitignore` 已排除 `build/`） |
| `deliver/` | 带版本号的交付副本 | 是 |

### 1.3 preset 生效规则（全阶段一致）

1. 命令行 `--preset <name>` 显式指定 → 用它。
2. 未指定但 `meta.pageSize` 存在 → 由尺寸推导（`13.333x7.5`→`slide-16x9`，`8.268x11.693`→`print-A4`，`11.693x16.535`→`print-A3`，`8.5x11`→`print-letter`）。
3. 都没有 → `slide-16x9`（保持 v1.1 行为）。
4. 任何脚本的报告头部**必须打印**：`preset: <name> (source: cli|meta|default)`。

### 1.4 回归基线（每个阶段收尾都跑）

| 基线 | 命令 | 合格线 |
|---|---|---|
| 全库自检 | `python scripts/selfcheck.py --strict` | `P0 0 / P1 0 / P2 0` |
| 演示稿 | `python scripts/qa_check.py --deck demo/source/deck.json --out build/qa-demo.md --json build/qa-demo.json` | `0/0/0/0` |
| smoke | `qa_check.py --deck demo/smoke/deck.json …` | `0/0/0/0` |
| 三个 slide 案例 | 各自 `qa_check.py` | `0/0/0/0` |

### 1.5 退出码约定（新增脚本沿用现有语义）

| 码 | 含义 |
|---|---|
| 0 | 成功（可能有 P1/P2 警告） |
| 1 | 成功但触发阻断级问题（有 P0），或产物降级（缺图/缺字体） |
| 2 | 失败（参数错、I/O 错、schema 不合法、后端不可用） |

---

## 2. P0 · 文档骨架（目标：先证明"能出文档"）

### 2.1 阶段目标与完成定义

- 目标：`doc.json` → `.docx` → `.pdf` 全链路打通，6 页样张可交付。
- **完成定义**：下面 5 条同时满足，缺一条不算过。
  1. `build_docx.py` 能出 `.docx`，且用 python-docx 复读能确认 6 种块都落成了对应样式。
  2. `export_pdf.py` 能出 `.pdf`，`missingFonts` 为空。
  3. `qa_doc.py` 七项检查全部实现，对样张报 `P0=0`。
  4. 样张逐页 PNG 目视通过（我出 6 张截图）。
  5. 回归基线四份 QA 仍 `0/0/0/0`。

### 2.2 前置检查（动手前先跑，任何一条不过就先修环境）

```bash
python -c "import docx, PIL, lxml; print(docx.__version__, PIL.__version__)"
python scripts/probe.py --json
node office-cli.mjs capabilities --json
```

期望：`1.2.0 12.3.0`；probe 报 Python 3.12.14 / python-pptx 1.0.2 / 中文字体 38 族；capabilities 的 `conversions` 里含 `doc/docx → pdf`。

### 2.3 文件清单与创建顺序（6 个文件）

| 顺序 | 文件 | 规模预估 | 关键函数签名 |
|---|---|---|---|
| 1 | `schema/doc.schema.json` | 300–380 行 | —（纯声明） |
| 2 | `scripts/lib/docload.py` | 180–220 行 | `load_doc(path) -> dict`；`validate(doc) -> list[Issue]`；`iter_blocks(doc) -> Iterator[(section, block, index)]` |
| 3 | `scripts/build_docx.py` | 550–700 行 | `main(argv) -> int`；`build(doc, out_path, preset) -> BuildReport`；`_apply_styles(docx_doc, tokens, preset)`；`_add_heading/_add_paragraph/_add_list/_add_table/_add_figure/_add_pagebreak` |
| 4 | `scripts/export_pdf.py` | 220–280 行 | `main(argv) -> int`；`convert(input_path, out_path) -> ConvertReport`；`verify(pdf_path) -> list[Issue]` |
| 5 | `scripts/qa_doc.py` | 600–750 行 | `main(argv) -> int`；`build_context(doc, thresholds, preset) -> Ctx`；`run_checks(ctx) -> list[Issue]`；`render_md/issues_json` |
| 6 | `examples/docs/report-min/{brief.md, source/doc.json, README.md}` | 简报 60 行 / 源 300 行 / 说明 90 行 | — |

**为什么先写 `docload.py`**：schema 校验与遍历是所有下游脚本的地基；先写它，`build_docx` 与 `qa_doc` 就能共享同一套"块遍历 + 错误定位"，避免两处各写一遍（v1.1 的 `deckload.py` 就是这套做法，30 个 check 才能复用）。

### 2.4 步骤 S1–S12

| 步 | 动作 | 命令 / 检查 | 期望输出 | 失败处理 |
|---|---|---|---|---|
| S1 | 写 `doc.schema.json` | — | 字段按 §2.5 全表覆盖 | 缺字段先补全再往下 |
| S2 | 写 `docload.py` | `python -c "from scripts.lib.docload import load_doc; d=load_doc(r'examples/docs/report-min/source/doc.json'); print(len(d['sections']))"` | 打印 5（样例 5 个 section） | 报错先修 loader |
| S3 | 造样张源 `doc.json` | — | 6 页内容按 §2.6 | 内容不够就补块，别改结构 |
| S4 | 写 `build_docx.py` 骨架（样式注入 + heading/paragraph） | `python scripts/build_docx.py --doc examples/docs/report-min/source/doc.json --out examples/docs/report-min/build/report.docx` | 退出 0，文件 >25 KB | 见 §2.8 排错表 |
| S5 | 补 list / table / figure / pageBreak 四种块 | 同上 | 退出 0，python-docx 复读能数到 1 张表 1 张图 | 表格宽度报警先记下不阻断 |
| S6 | 写 `export_pdf.py` | `python scripts/export_pdf.py --input …/report.docx --out …/report.pdf` | 退出 0，`pageCount` 5–7，`missingFonts: []` | 缺字体按 §2.8 替换 |
| S7 | 写 `qa_doc.py` 的 DOC-01/02/05/06/08/12/13 | `python scripts/qa_doc.py --doc … --out …/qa.md --json …/qa.json` | 有 `SUMMARY` 行；样张 `P0=0` | 先让七项能跑，判据可先严后松 |
| S8 | 造两份 fixture（故意犯错） | 见 §2.7 的 fixture 清单 | 两份各触发 5–8 条 | 触发不到说明判据没接上 |
| S9 | 渲 PDF 成 PNG | `node office-cli.mjs render --input …/report.pdf --output-dir …/build/preview --dpi 150` | 6 张 1241×1754 PNG | 页数不符回到 S4 查分页 |
| S10 | 目视复核 | 逐页看：标题层级、表格边框、图注、页码 | 6 张截图存 `records/` | 问题写进 `qa.md` 备注 |
| S11 | 跑回归基线 | §1.4 四条命令 | 四份 `0/0/0/0` | 任何一份不为 0 就不算完成 |
| S12 | 写 `README.md`（案例） | — | 命令、真实数字、已知限制、检查清单 | 数字必须来自实测 |

### 2.5 `doc.schema.json` 逐块字段（P0 只做前六种）

**通用块字段**：`type`（必填）、`id`（可选，`^b[0-9]{2,3}$`）、`note`（可选，作者备注，不渲染）。

| 块 | 必填 | 可选 | 约束 |
|---|---|---|---|
| `heading` | `level`(1–3)、`text` | `id` | 文本 ≤24 字；层级不得跳级 |
| `paragraph` | `runs[]` 或 `text` | `align`(`left/justify`)、`indentFirstLine`(bool，默认 `true`)、`spaceAfterPt`(默认 6)、`style`(`body`/`small`) | 每段 ≤300 字 |
| `list` | `items[]`（每项 `runs[]` 或字符串） | `ordered`(默认 `false`)、`level`(1–2) | 每项 ≤80 字；只允许两级 |
| `table` | `header[]`、`rows[][]` | `caption`、`source`、`colWidthsIn[]`、`align[]`、`zebra`(默认 `true`) | 列 ≤6；`caption`/`source` 为必填业务字段（缺失判 P0） |
| `figure` | `src` | `widthIn`、`caption`、`source`、`alt` | `caption`/`source` 必填业务字段；`alt` 缺失是 P2 |
| `pageBreak` | — | — | 连续两个 `pageBreak` 判 P2（blank page 风险） |

**P1 才做的六种**：`quote`、`callout`、`code`、`toc`、`bibliography`、`footnoteGroup`（字段见 §3.5）。

**JSON Schema 的关键约束写法（照这个来，别自创）**

```json
{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "type": "object",
  "required": ["meta", "tokens", "sections"],
  "properties": {
    "meta": {
      "type": "object",
      "required": ["title", "pageSize", "orientation", "marginsIn"],
      "properties": {
        "title": { "type": "string", "minLength": 1, "maxLength": 60 },
        "pageSize": { "type": "string", "pattern": "^(A4|A3|Letter|[0-9]+(\\.[0-9]+)?x[0-9]+(\\.[0-9]+)?)$" },
        "orientation": { "enum": ["portrait", "landscape"] },
        "marginsIn": {
          "type": "object",
          "required": ["top", "bottom", "left", "right"],
          "additionalProperties": false,
          "properties": {
            "top": { "type": "number", "minimum": 0.5, "maximum": 1.6 },
            "bottom": { "type": "number", "minimum": 0.5, "maximum": 1.6 },
            "left": { "type": "number", "minimum": 0.5, "maximum": 1.6 },
            "right": { "type": "number", "minimum": 0.5, "maximum": 1.6 }
          }
        }
      }
    },
    "sections": {
      "type": "array", "minItems": 1,
      "items": {
        "type": "object",
        "required": ["id", "blocks"],
        "properties": {
          "id": { "type": "string", "pattern": "^s[0-9]{1,3}$" },
          "blocks": { "type": "array", "minItems": 1, "items": { "$ref": "#/$defs/block" } }
        }
      }
    }
  }
}
```

### 2.6 样张 `report-min` 的 6 页内容（逐页定稿，避免"随便填"）

| 页 | 内容 | 块构成 | 验收点 |
|---|---|---|---|
| 1 | 封面：标题 + 副标题 + 作者/单位/日期 + 一条分隔线 | `heading`(1) + `paragraph`×3 + `pageBreak` | 标题字号 20 pt、居中、日期右对齐 |
| 2 | 目录占位 | `toc` + `heading`(1)"一、研究问题" + `paragraph`×2 | 目录域已插入（F9 可更新） |
| 3 | 研究问题（续） | `paragraph`×3 + `list`(无序，4 项) | 列表缩进对齐、项目符号一致 |
| 4 | 观测方法 | `table`(4 列 × 5 行，含表头) + `paragraph`×1 | 表头跨页控制、列宽和 = 正文宽 |
| 5 | 结果 | `figure`(示意图) + `paragraph`×2 + `table`(3 列 × 4 行) | 图注"图 1 · …"、来源行存在 |
| 6 | 结论与局限 | `heading`(2) + `paragraph`×3 + `pageBreak` | 二级标题进目录、末页无空白页 |

字数控制：全篇 900–1200 字（A4 6 页、10.5 pt、1.5 行距的合理密度）。

### 2.7 验收勾选表（每条都要留证据）

| 编号 | 判据 | 证据 |
|---|---|---|
| P0-1 | `build_docx.py` 退出 0，`.docx` ≥25 KB | 命令输出 + 文件大小 |
| P0-2 | 复读样式：`python -c "from docx import Document; d=Document(p); print(sorted({r.style.name for r in d.paragraphs}))"` 含 `Heading 1`/`Heading 2`/`Normal` | 打印结果 |
| P0-3 | `export_pdf.py` 退出 0，`pageCount` 在 5–7，`missingFonts` 为空 | manifest JSON 片段 |
| P0-4 | `qa_doc.py` 对样张 `P0=0`；两份 fixture 各触发 ≥5 条 | 三份 qa.md |
| P0-5 | 6 张 PNG 目视通过 | 6 张图存 `records/` |
| P0-6 | 回归基线四份 `0/0/0/0` | 四份 qa.md |

**fixture 清单（故意做错，用来证明检查项真的在工作）**

| fixture | 故意犯的错 | 应触发 |
|---|---|---|
| `qa-doc-bad.json` | h1 后直接 h3；表格无 source；图无 caption；字号 8 pt；正文行距 1.2；页边 0.4 in | `DOC-01` `DOC-05` `DOC-06` `DOC-09` `DOC-10` `DOC-11`（≥5 条） |
| `qa-doc-edge.json` | 单段 380 字；6 列以上表格；目录关闭但标题 9 个；图片 120×68 px 铺 6 in | `DOC-04` `DOC-14` `DOC-07` `DOC-12`（≥4 条） |

### 2.8 排错表（P0 阶段真会遇到的）

| 现象 | 原因 | 动作 |
|---|---|---|
| docx 打开后中文是宋体 | 只设了 `font.name`（西文），没设 `w:eastAsia` | 双写：`run.font.name` + `run._element.rPr.rFonts.set(qn('w:eastAsia'), family)` |
| 生成的 PDF 页数与目录页码不符 | TOC 是域，未更新；页码是按"未分页"的估算 | 交付说明写"打开后 F9"；验收改为"目录最后一条页码 = PDF 实际页数" |
| `missingFonts:["Courier"]` | Word 默认样式里残留 Courier | 覆盖 `Normal`/`Heading`/`Caption` 三处字体，复跑转换确认告警消失 |
| 表格右边缘超出正文宽 | `colWidthsIn` 之和 ≠ 正文宽 | 生成器按比例归一到正文宽，并在 stderr 提示原值 |
| 图片超出页面高度 | `widthIn` 用了图的原生宽度 | 按"高度 ≤ 页高 80%"反算宽度上限 |
| 末页空白 | 末尾多了 `pageBreak` | 允许末尾 `pageBreak` 但生成器修剪最后一个 |

---

## 3. P1 · 文档完整能力（目标：能交付 30 页长文档）

### 3.1 阶段目标与完成定义

1. 12 种块全部实现（补 `quote`/`callout`/`code`/`toc`/`bibliography`/`footnoteGroup`）。
2. 表格跨页表头重复、行不断裂。
3. 页眉页脚分节（首页不同、正文有页码）。
4. 图注与表注自动编号（图 1…图 N、表 1…表 N）。
5. 30 页长文档产物与目录页码自洽。
6. `qa_doc.py` 15 项全部被真实触发过（两份以上 fixture）。
7. 两份规范文档落盘，`SKILL.md` 路由与判定表接上。

### 3.2 30 页长文档的骨架（`examples/docs/technical-report`）

| 章 | 页数 | 内容要求 | 需要的块 |
|---|---|---|---|
| 封面 | 1 | 标题、副标题、编制单位、版本、日期 | heading/paragraph/pageBreak |
| 目录 | 1 | 自动目录（3 级） | toc |
| 1 背景与目标 | 3 | 问题陈述、范围、读者对象 | heading/paragraph/list |
| 2 方法 | 5 | 观测设计、参数表、流程图（图） | table/figure/list |
| 3 结果 | 7 | 3 组数据（每组件 1 表 + 1 图 + 3 段解读） | table/figure/paragraph/callout |
| 4 讨论 | 6 | 三条结论、每条带证据与边界 | heading/paragraph/quote |
| 5 局限与后续 | 3 | 局限清单、下一步计划 | list/callout |
| 6 附录 | 4 | 完整数据表（跨页）、参数定义、代码片段 | table(跨页)/code |
| 参考 | 1 | 参考文献 8 条 | bibliography |

**"30 页"怎么来**：A4、正文 10.5 pt、行距 1.5、每页约 700–900 字；上表按"内容字数 ÷ 每页字数"分配，允许 ±2 页浮动。**不通过调大字号凑页数**。

### 3.3 12 组 tokens 变体（回归"换主题不塌"）

| 组 | 字体（display / body） | 主色 | 预期风险点 |
|---|---|---|---|
| 1 | 思源宋体 / 微软雅黑 | `#8C3A1E` 赭 | 基线 |
| 2 | 微软雅黑 / 微软雅黑 | `#1F3B33` 墨绿 | 标题与正文同族，层级靠字号 |
| 3 | Noto Sans SC / Noto Sans SC | `#0B3D91` 深蓝 | 无衬线长文可读性 |
| 4 | 思源宋体 / 思源宋体 | `#3A3A3A` 中性灰 | 全衬线，行距需加大 |
| 5–8 | 同 1–4 | 各换一种深浅对比较小的主色 | 对比度可能 <4.5:1 |
| 9–12 | 同 1–4 | 换暖底 `#F7F3EA` + 深字 | 表头底色与斑马纹需重算 |

每组都要跑：`build_docx.py` → `export_pdf.py` → `qa_doc.py`，记录 `P0/P1` 与页数；**验收线**：`P0=0` 且页数变化 ≤1。

### 3.4 表格跨页的三处 OOXML（照抄即可）

```python
# 1) 表头行跨页重复
tr = table.rows[0]._tr
trPr = tr.get_or_add_trPr()
th = OxmlElement('w:tblHeader'); th.set(qn('w:val'), 'true'); trPr.append(th)

# 2) 行内不断裂（所有行都加，含表头）
for row in table.rows:
    trPr = row._tr.get_or_add_trPr()
    cs = OxmlElement('w:cantSplit'); cs.set(qn('w:val'), 'true'); trPr.append(cs)

# 3) 表头行高固定，防止压缩
table.rows[0].height = Pt(22)
```

验证方法：造一张"25 行、正好跨页"的表，渲成 PNG，确认第 2 页顶部仍有表头且没有半行被切断（截图留证）。

### 3.5 P1 新增块的字段

| 块 | 必填 | 可选 | 渲染 |
|---|---|---|---|
| `quote` | `text` | `cite` | 左缩进 0.3 in + 左侧 2 pt 竖线 + 字号 -0.5 pt |
| `callout` | `kind`(`note`/`warn`/`key`)、`text` | `title` | 单格表格：浅底 + 0.75 pt 边 + 内边距 0.1 in |
| `code` | `text` | `lang` | 等宽 9 pt、灰底、左右各缩进 0.1 in |
| `toc` | — | `depth`(默认 3) | 插入 TOC 域 + 一行提示 |
| `bibliography` | `entries[]` | `title` | 编号 `[1]`–`[N]`，悬挂缩进 0.3 in |
| `footnoteGroup` | `id`、`items[]` | — | **降级**：渲成章末"注释"小节（脚注 API 不支持，见规划书 R2） |

### 3.6 页数一致性的校验方法（替代规划书 A2）

```bash
# 1) PDF 实际页数（以 pdfium 渲染计数为准）
node office-cli.mjs render --input build/report.pdf --output-dir build/preview --pages all --dpi 96
# 输出 JSON 的 pageCount 就是实际页数

# 2) 目录最后一条的页码（PDF 里读第 2 页）
# 人工看：目录最后一行右对齐的数字必须等于上面 pageCount

# 3) 抽查 3 处：目录第 1/4/7 条 → 翻到对应页确认标题一致
```

判据：`pageCount == 目录最后一条页码`，且 3 处抽查全对。**允许 ±1**（LibreOffice 与 Word 分页引擎差异），超出即 P1。

### 3.7 验收勾选表

| 编号 | 判据 | 证据 |
|---|---|---|
| P1-1 | 12 种块全实现，样张 30 页含全部块 | `tech-report` 的 docx/pdf |
| P1-2 | 表格跨页：25 行表跨页后第 2 页有表头、无半行 | 2 张截图 |
| P1-3 | 页眉页脚：首页无页眉、正文页脚有页码 | 2 张截图 |
| P1-4 | 图注表注编号连续（图 1–6、表 1–5） | 抽查 3 处 |
| P1-5 | 30 页文档页码自洽（§3.6） | pageCount + 3 处抽查 |
| P1-6 | 12 组 tokens 变体均 `P0=0`，页数变化 ≤1 | 12 行记录表 |
| P1-7 | `qa_doc.py` 15 项各被触发 ≥1 次 | fixture 的 qa.md |
| P1-8 | `references/document-*.md`、`workflows/05b`、两个模板落盘；`SKILL.md` 路由接上 | 文件 + selfcheck |

---

## 4. P2 · 设计型导出与打印 preset（目标：给设计方向交付）

### 4.1 阶段目标与完成定义

1. 四档 preset 落地，`--preset` 贯穿 `build_pptx` / `qa_check` / `export_svg`。
2. A4 海报一份（PDF + PNG），`qa_check.py` 含 `PRT-01/02` 均 `P0=0`。
3. 任一页 `deck.json` 导出 SVG，文字可搜、位置与 PNG 差 ≤2 px。
4. 分层 PNG 包五层齐全，PS 能逐层导入。

### 4.2 `--preset` 如何不破坏红线（G8 的解法）

**规则：加可选参数，不带该参数时输出与 v1.1.2 逐字节一致**
- **触发条件**：需要打印尺寸（A4/A3）的版面。
- **具体动作**：`build_pptx.py` 增加 `--preset`（默认 `None`）。非空时做三件事：① 校验 `meta.size` 与 preset 画布一致，不一致退出 2；② 图片默认缩放与最小字号改用 preset 阈值；③ stderr 打印 `preset: print-A4 (source: cli)`。**坐标、元素顺序、XML 结构一律不动**。
- **常见错误**：把 preset 当成"换一套 tokens"，导致同一份 deck 在不同 preset 下版面漂移。

验收这条红线的方法：
```bash
# 不带 --preset 重建 demo，与仓库里的 pptx 逐成员比对，必须 0 差异
python scripts/build_pptx.py --deck demo/source/deck.json --out build/deck-v112.pptx
python -c "import zipfile,sys; a=zipfile.ZipFile('build/deck-v112.pptx'); b=zipfile.ZipFile('demo/build/deck.pptx'); print(sum(1 for n in a.namelist() if a.read(n)!=b.read(n)))"
```
期望：打印 `0`。

### 4.3 preset 的写入位置

| 文件 | 加什么 |
|---|---|
| `schema/thresholds.json` | 顶层加 `print` 段（`A4`/`A3`/`letter` 的字号区间、行距、页边、密度），`version` → `2.0.0` |
| `schema/presets.json` | 只做引用：`{"print-A4": {"canvas": "print.A4", "type": "type.print"}}` |
| 脚本 | 统一从 `presets.json` 取 key，再去 `thresholds.json` 取值（不复制数字） |

### 4.4 SVG 导出：文字必须预排版（G9）

SVG 1.1 的 `<text>` 不自动换行，行宽超了会直接溢出画布。做法：

```python
# 复用现有度量：scripts/lib/metrics.py 的 measure_text / wrap_lines
from scripts.lib.metrics import wrap_lines
lines = wrap_lines(text, font_family, size_pt, box_width_in)   # 已处理中英文混排与避头尾
for i, line in enumerate(lines):
    y = base_y + i * (size_pt * line_height / 72.0)
    emit(f'<tspan x="{x:.2f}" y="{y:.2f}">{escape(line)}</tspan>')
```

**不许做的事**：用 `textLength` 拉伸字形去"凑宽度"（会变形），或把中文转成 `<path>`（不可搜、不可改）。`EXP-03` 专门抓这条。

### 4.5 `layers.json` 结构

```json
{
  "schemaVersion": 1,
  "slide": 1,
  "canvasPx": [1280, 720],
  "preset": "print-A4",
  "layers": [
    { "order": 0, "name": "01-bg",     "file": "layers/01-bg.png",     "opacity": 1.0, "elements": ["bg"] },
    { "order": 1, "name": "02-decor",  "file": "layers/02-decor.png",  "opacity": 1.0, "elements": ["e12", "e13"] },
    { "order": 2, "name": "03-images", "file": "layers/03-images.png", "opacity": 1.0, "elements": ["e21"] },
    { "order": 3, "name": "04-charts", "file": "layers/04-charts.png", "opacity": 1.0, "elements": [] },
    { "order": 4, "name": "05-text",   "file": "layers/05-text.png",   "opacity": 1.0, "elements": ["e31", "e32", "e33"] }
  ]
}
```

分层实现：按 `zRole` 与元素类型分组，逐层用 Pillow 画（`ImageDraw` + 已渲好的文字位图裁切），每层输出同等画布尺寸的透明 PNG（**不是裁到内容包围盒**），这样在 PS 里导入后逐层对齐无需手动移动。

### 4.6 A4 海报 fixture（`examples/docs/poster-a4`）

| 要求 | 数值 |
|---|---|
| 画布 | `meta.size: "8.268x11.693"`、`preset: print-A4` |
| 主标题 | 44–56 pt、占页宽 ≥60% |
| 正文 | 10–11 pt、行距 1.45 |
| 元素数 | ≤14、文本块 ≤4 |
| 图 | 1 张满版图（`zRole:"bgDecor"`）+ 1 张数据图 |
| 边距 | 上 0.9 / 下 0.9 / 左右 1.0 in |
| 验收 | PDF 1 页、PNG 1241×1754 @150 dpi、`P0=0`（含 `PRT-01/02`） |

### 4.7 验收勾选表

| 编号 | 判据 | 证据 |
|---|---|---|
| P2-1 | 四档 preset 生效，报告头部打印 `source` | 4 次命令输出 |
| P2-2 | 不带 `--preset` 重建 demo 与仓库 pptx **0 差异成员** | 比对脚本输出 `0` |
| P2-3 | A4 海报 PDF 1 页 + PNG 1241×1754，`P0=0` | 产物 + qa.md |
| P2-4 | SVG：浏览器 `Ctrl+F` 搜到标题；与 PNG 叠图差 ≤2 px | 截图 + 叠图 |
| P2-5 | 分层包五层 + `layers.json`，PS 导入成功 | 截图 + json |
| P2-6 | `EXP-01/02/03` 各被触发 ≥1 次 | fixture |

---

## 5. P3 · 既有 PDF 改版（目标：能接"改设计"的活，且边界写死）

### 5.1 阶段目标与完成定义

1. 8 页 PDF 进、8 页 PDF 出，页序页数一致。
2. 底图 150 dpi、无叠影、无重影。
3. 支持三类操作：遮挡（`mask`）、覆盖（`overlay`）、色彩重映射（`recolor`）。
4. `changes.md` 逐页记录改了什么，可与对照图逐条核对。
5. `--help` 第一段写清能力边界（文字不可编辑）。

### 5.2 流水线五步

| 步 | 动作 | 产物 |
|---|---|---|
| 1 | 原 PDF 渲成 PNG（150 dpi） | `build/origin/p1..pN.png` |
| 2 | 读 `rebuild.json`，逐页处理底图（遮挡 + 色彩重映射） | `build/base/p1..pN.png` |
| 3 | 用 `deck.json` 语法写覆盖层，底图作为 `zRole:"bgDecor"` 的 `image` | `build/overlay/deck.json` |
| 4 | `build_pptx.py` → `export_pdf.py` | `build/rebuild.pdf` |
| 5 | 生成对照图与改动清单 | `build/compare/pN-before.png|after.png`、`build/changes.md` |

### 5.3 `rebuild.json` 完整结构

```json
{
  "meta": { "dpi": 150, "pageSize": "A4", "orientation": "portrait", "fallback": "keep-original" },
  "pages": [
    {
      "page": 1,
      "keep": "full",
      "mask": [
        { "x": 0.55, "y": 6.70, "w": 5.80, "h": 0.42, "fill": "#FFFFFF" }
      ],
      "recolor": [
        { "from": "#C8501E", "to": "#1F3B33", "tolerance": 24, "maxPixelRatio": 0.35 }
      ],
      "overlay": [
        { "kind": "text", "role": "h1", "x": 0.62, "y": 0.95, "w": 6.90, "h": 0.90,
          "paragraphs": [ { "text": "把观测变成可比的数据" } ] }
      ]
    }
  ]
}
```

| 字段 | 约束 |
|---|---|
| `keep` | `full`（整页保留，只叠加） / `none`（丢弃底图，纯重建） |
| `mask` | 坐标单位英寸、原点左上；`fill` 缺省 `#FFFFFF` |
| `recolor.tolerance` | 0–60（RGB 欧氏距离）；>40 容易误伤邻近色 |
| `recolor.maxPixelRatio` | 命中像素占比上限，超限**中止该页替换**并写警告（防整页变色） |
| `overlay` | 与 `deck.json` 元素语法完全相同（`kind`/`role`/`x`/`y`/`w`/`h`/`paragraphs`） |

### 5.4 色彩替换算法（两步，避免锯齿光晕）

```text
1) 硬替换：distance(rgb, from) <= tolerance  → rgb = to
2) 边缘过渡：distance <= tolerance * 1.35 且未被替换  → 50% 混合(from_mix, to)
   判定：仅在"距已替换像素 1 px 内"的边缘像素上做，避免污染大片区域
3) 统计：replaced / 目标区域像素数 → 写入 changes.md
```

判据：`替换率 ≥95%`（目标区域内），且 2 倍放大目视**无彩色描边**（光晕即失败）。

### 5.5 验收勾选表

| 编号 | 判据 | 证据 |
|---|---|---|
| P3-1 | 8 页进 8 页出，页序一致 | 两份 PDF 的 pageCount |
| P3-2 | 底图 150 dpi（1241×1754 量级） | PNG 尺寸 |
| P3-3 | 无叠影：覆盖元素与底图错位 ≤2 px | 2 倍放大截图 |
| P3-4 | 色彩替换率 ≥95% 且无光晕 | changes.md + 截图 |
| P3-5 | `changes.md` 每页有"改了什么/为什么"两行 | 文件 |
| P3-6 | `--help` 首段含能力边界声明 | 命令输出 |

---

## 6. P4 · 案例与验收（目标：能给别人看）

### 6.1 阶段目标

三个文档案例 + 一份 v2 验证报告 + 发布包。

### 6.2 三个案例的大纲与"真实感"要求

| 案例 | 页数 | 结构 | 真实感要求 |
|---|---|---|---|
| `technical-report` | 24–30 | 见 §3.2 | 每章有数字与单位；两条结论各带反例边界 |
| `research-brief` | 6 | 问题 / 方法（3 条）/ 3 个发现（每组图或表 + 2 段）/ 局限 / 建议 | 发现必须是"判断句"，不能是"整理了数据" |
| `one-pager` | 2 | 一页纸方案：背景 1 段 / 方案 3 条 / 成本与周期表 / 风险 2 条 / 下一步 | 所有数字虚构但要标注"演示样本，不对应任何真实项目" |

### 6.3 案例 README 模板（四段固定）

1. **怎么复现**：4 条命令（build_docx / export_pdf / qa_doc / render）。
2. **真实数字表**：doc.json 字节、docx 字节、pdf 字节、页数、字数、块数、字体族、QA 摘要。
3. **诚实限制**：3 处已知弱点（如"目录需 F9 更新""脚注降级为章末注释"）。
4. **检查清单**：5–8 条可判定项。

### 6.4 验证报告 V9–V14

| 编号 | 验什么 | 判据 |
|---|---|---|
| V9 | 文档最小闭环 | P0 阶段 6 条全过 |
| V10 | 30 页长文档 | §3.6 页码自洽 + 表格跨页 |
| V11 | A4 海报 | PDF 1 页 + PNG + `P0=0` |
| V12 | SVG / 分层包 | 文字可搜 + 五层 + PS 导入 |
| V13 | PDF 改版 | 8→8 页 + 替换率 ≥95% |
| V14 | 12 组 tokens | 全 `P0=0`、页数变化 ≤1 |

### 6.5 打包与解包复验（每次发布都做）

```bash
# 打包（排除缓存与发布目录）
python -c "import os,zipfile; R='.'; S={'__pycache__','.git','dist','.venv'}; \
z=zipfile.ZipFile('dist/ai-ppt-design-system-v2.0.0.zip','w',zipfile.ZIP_DEFLATED,compresslevel=9); \
[z.write(os.path.join(d,f),'ai-ppt-design-system/'+os.path.relpath(os.path.join(d,f),R).replace('\\\\','/')) \
 for d,ds,fs in os.walk(R) if not ds.__setitem__(slice(None),[x for x in ds if x not in S]) \
 for f in fs if not f.endswith(('.pyc','.pyo'))]; z.close()"

# 解包复验（不能只信打包前）
# 1) 解到临时目录 2) selfcheck --strict 3) build 一次 pdf 4) qa_doc 一次
```

---

## 7. 跨阶段：回归与发布节奏

| 时点 | 动作 | 不做会怎样 |
|---|---|---|
| 每阶段开始前 | 跑 §1.4 回归基线，记下当前数字 | 出问题时分不清是不是本阶段引入的 |
| 每阶段收尾 | 再跑一次基线，数字必须与开始前一致 | 悄悄破坏 slide 分支 |
| 版本号变更时 | 更新 `CHANGELOG.md` + README 当前版本 + 重新打包 | 交付物与文档不符 |
| 每次打包后 | 解包复验（selfcheck + 一次 build + 一次 QA） | 包里可能少了文件（v1.1 曾漏过 further-reading 类新文件） |

---

## 8. 风险触发后的具体动作（把规划书第 11 章变成命令）

| 风险 | 触发信号 | 立刻执行 |
|---|---|---|
| R2 无脚注 API | 需要真页脚注 | 改用 `footnoteGroup` 渲成章末"注释"，并在 README 写明；不硬改 OOXML |
| R3 分页不一致 | Word 里比 PDF 多 2 页 | 以 PDF 为交付；README 写"docx 页码以 PDF 为准"；`DOC-13` 只校验 PAGE 域存在 |
| R4 缺字体 | `missingFonts` 非空 | 按映射表替换（display→思源宋体/雅黑，body→雅黑，mono→Cascadia Mono），复跑确认空 |
| R5 大文档慢 | build >60 s | 按 section 切分构建 + 缓存已渲图片；不承诺 200 页 |
| R6 要 .psd | 用户坚持 | 交付 SVG + 分层 PNG + 说明差异（文字层是位图） |
| R7 表格断了 | 抽页看到半行 | 补 `tblHeader` + `cantSplit`；必要处前置 `pageBreak` |
| R8 叠影 | 覆盖层错位 >2 px | 用底图 PNG 的像素坐标做基准（`--align-to 150`），错位即返工 |

---

## 9. 详细度自检（这份手册够不够，逐条打勾）

| 检查 | 判据 | 结果 |
|---|---|---|
| 文件级 | 每个要新建的文件都有路径、规模、关键函数签名 | 是（§2.3、§4.5） |
| 步骤级 | 每个阶段有 S1..Sn，每步有命令与期望输出 | 是（§2.4 等） |
| 判据级 | 每条验收都能用一条命令或一张截图证明 | 是（各验收表） |
| 失败级 | 每个阶段有排错表，覆盖最可能踩的 5–7 个坑 | 是（§2.8 等） |
| 冲突级 | 与既有红线的冲突已发现并给出解法 | 是（§0.2 的 G8、A2） |
| 数据级 | 所有阈值有出处（`thresholds.json`），不散落在脚本里 | 是（§1.3、§4.3） |
| 边界级 | 做不到的事写在最前面，并给替代 | 是（规划书 1.2、本文 §5.1） |

---

## 检查清单

- [ ] P0 的 6 页样张内容（§2.6）符合你的真实文档习惯，而不是我编的结构
- [ ] P1 的 30 页长文档大纲（§3.2）覆盖你要写的文档类型
- [ ] P1 的 12 组 tokens 变体（§3.3）够不够证明"换主题不塌"
- [ ] P2 的 A4 海报规格（§4.6）与你想要的"设计方向产出"一致
- [ ] P3 的改版能力边界（§5.1，文字不可编辑）你能接受
- [ ] 三个案例选型（§6.2）与你的实际场景对得上
- [ ] §0.2 里我主动改正的两处矛盾（preset 参数、页数判据）你认可
- [ ] 确认这份手册的详细度足够，我可以按 P0 开工
