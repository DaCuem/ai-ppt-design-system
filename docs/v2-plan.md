# AI PPT 设计系统 · v2.0 补全规划与操作书

| 项 | 内容 |
|---|---|
| 版本 | 规划稿 v1.0（待确认） |
| 日期 | 2026-11-15 |
| 依据 | 你给的 v1.1 实施框架（1108 行 / 11 章）+ 本次需求：PDF 从零做、改既有 PDF、兼顾 Word、偏设计方向的 PNG/PSD |
| 现状基线 | 本仓库 v1.1.2（159 文件 / 16.87 MB / 30 个 check id），PPT 分支已端到端验证 |
| 读者 | 你（决策与验收）、我（执行）、将来接手的人（复现） |
| 状态 | **等你看完点头才开工**；开工前不改任何脚本行为 |

---

## 0. 这份文档怎么用

### 0.1 它是什么

不是需求清单，是**可执行的操作书**：每一阶段写清「交付什么文件 → 用什么命令产出 → 拿什么判据验收 → 不达标就停」。你看完只需要做三件事：确认范围（第 1 章）、确认顺序（第 9 章）、确认边界（附录 B）。

### 0.2 与 v1.1 规范的关系

| v1.1 已有的 | v2.0 怎么处理 |
|---|---|
| 8 阶段工作流（0 需求探索 → 7 迭代交付） | **不动**。加一条分支：阶段 5 产出"文档"时走 `05b`，其余阶段共用 |
| 设计知识库（风格/版式/字体/色彩/图像/图表/反模式） | **不动**，直接复用。文档形态只新增字号与版面的映射规则 |
| spec-first 主路径（JSON → 生成器 → 产物 → 渲染 → 复核） | **沿用同一条思路**：`doc.json → docx → pdf → 逐页 PNG → 目视复核` |
| 30 个 check id 的机器质检 | **不动**。文档形态另起 `qa_doc.py`，共享分级与报告格式 |
| 三条红线（不虚构/不设默认风格/不覆盖原件） | **不动**，并新增 3 条（附录 B） |

### 0.3 三种读法

- **只想知道结论**：读第 1 章 + 第 9 章的两张表 + 附录 B。
- **要验收**：读第 10 章验收总表，每条都给了命令与判据。
- **要复现或接手**：从第 4 章目录树、第 5–8 章模型与契约读起，附录 C 是复制即用的命令序列。

### 0.4 术语（沿用 + 新增）

沿用：版式、版式签名、文本块、留白（必须带数值）、阈值唯一真值、spec-first。

新增四个：

| 术语 | 定义 |
|---|---|
| **产物形态** | 一次任务的最终交付类型：演示稿（slide）/ 文档（doc）/ 固定版面设计（page）/ 既有文件改版（rebuild） |
| **文档块** | `doc.json` 里最小的内容单位，共 12 种：标题、段落、列表、表格、图、图注、引文、提示框、代码、分页、目录、脚注组 |
| **preset** | 一套绑定的阈值组合（画布尺寸 + 字号档 + 行距 + 边距 + 密度），任务开始时选定，中途不换 |
| **同源双出** | 一份 `doc.json` 同时产出可编辑 DOCX 与 PDF，两者页码一致（差异超出 ±1 页即判 P1） |

---

## 1. 目标与边界

### 1.1 要解决的三件事

| 编号 | 需求（你的原话） | v2.0 的解法 |
|---|---|---|
| N1 | 「PDF 一般是完整的文档，可不可以兼顾 Word 制作」 | 新增**文档分支**：`doc.json` → `build_docx.py` → `.docx`（可编辑）→ `export_pdf.py` → `.pdf`。一份源两种交付 |
| N2 | 「从零做一个或者是修改 pdf」 | 从零：走 N1，或走"固定版面设计"（A4/A3 单页、海报、画册跨页）。修改：`pdf_rebuild.py` 走"逐页底图 + 覆盖层重排"（能力边界见 1.2） |
| N3 | 「也不排除要更偏向设计的方向（可能也要求 png 或者 psd）」 | PNG 已有（`render.py` 任意画布）。新增 `export_svg.py`（真矢量、文字可编辑）与 `export_layers.py`（分层 PNG 包 + 坐标清单）。**PSD 二进制不做**，理由见 1.2 |

### 1.2 明确不做（这一节比"做什么"更重要）

| 不做 | 原因（实测，不是推测） |
|---|---|
| **PDF 文字级编辑**（改原 PDF 里的字） | 本机 Python 无任何 PDF 库（`fitz/pypdf/pdfplumber` 全部 ModuleNotFoundError），环境禁止 `pip install`；office-cli 只能把 PDF 渲成 PNG，不能回写内容流。自写 PDF 解析器只对极简单文件有效，误伤风险高于收益 |
| **PSD 二进制写入** | 无 `psd_tools`/`pytoshop`；8BPS 分层格式手写需处理通道、图层组、混合模式、智能对象，出错时 Photoshop 直接判文件损坏。替代品是 SVG + 分层 PNG，能进 PS 继续编辑，代价是"不是一个 .psd 文件" |
| **网页 / 长图 / H5** | 现有模型是固定画布 + 绝对坐标，没有流式布局与断点概念。硬做会得到"不能响应式的网页"，不如不做 |
| **提升 AI 生图质量** | 这属于图像生成能力，不是版式技能；v1.1.1 只解决了"生成图必须有结构核对与像素密度门槛" |
| **Word 里做复杂域**（交叉引用自动编号、题注域、Zotero 域） | python-docx 1.2.0 无对应 API，需手写 OOXML 且 LibreOffice 转换行为不稳。降级方案见 6.2.6（编号由生成器算好写成纯文本） |

### 1.3 一份知识库，四种形态

```text
              ┌─────────────── 共用知识层（不改） ───────────────┐
              │ visual-styles / design-principles / layout-recipes │
              │ typography / color / imagery / charts / rhythm     │
              │ anti-patterns / accessibility / further-reading    │
              └──────────────────────┬───────────────────────────┘
                                     │
        ┌──────────────┬─────────────┴────────────┬──────────────────┐
   形态 A 演示稿    形态 B 文档              形态 C 固定版面      形态 D 既有文件改版
   deck.json        doc.json                 deck.json(A4/A3)     pdf + rebuild.json
   → pptx           → docx → pdf             → pptx → pdf/png     → 底图 + 覆盖层 → pdf
   → png            → png(逐页)              → svg / 分层 png     → png(对照)
   （现状，不动）    （新增，v2.0 主线）        （新增，v1.2）        （新增，v2.1，受限）
```

### 1.4 成功判据（六条，都可测）

1. 一份 8 页 A4 文档，`doc.json` 一份源，产出 `.docx` 与 `.pdf`，页码一致（差 ≤1 页），`qa_doc.py` 报 `P0=0`。
2. 同一份 `doc.json` 只改 `tokens` 的字体与主色（≤6 个值），重建后版面不塌、QA 仍 `P0=0`（对应 v1.1 的 V2 验证思路）。
3. A4 海报：`deck.json` + `preset=print-A4` → PDF 与 PNG，`qa_check.py`（含新增 `PRT-01/02`）`P0=0`。
4. 任意一页 `deck.json` 可导出 SVG，浏览器打开后文字仍是文字（可用 Ctrl+F 搜到），形状无错位（与 PNG 叠图比对，位置差 ≤2 px @150dpi）。
5. 既有 PDF 改版：输入 8 页 PDF，输出 8 页 PDF，底图 ≥150 dpi，无叠影/重影，色彩重映射后目视一致（差异区域由人工确认清单记录）。
6. 三个文档案例（技术报告 / 研究摘要 / 一页纸方案）落盘，各自带 `brief.md`、`source/doc.json`、`build/`、`README.md`，QA 全绿。

---

## 2. 现状盘点（v1.1.2 实测）

### 2.1 仓库数字

| 目录 | 文件数 | 体积 | 说明 |
|---|---|---|---|
| 根目录 | 7 | 41.6 KB | SKILL / README / CONTRIBUTING / install / CHANGELOG / LICENSE / .gitignore |
| `discovery/` | 5 | 62.6 KB | 5 份探索规范 + 问题库 + 对照表 |
| `references/` | 11 | 169.5 KB | 10 种风格、10 条原则、L01–L15 版式、AP-01–16 反模式、6 份专项、延伸阅读 |
| `workflows/` | 8 | 74.7 KB | 00 需求探索 → 07 迭代交付 |
| `templates/` | 6 | 15.3 KB | 需求确认书 / 项目简报 / 设计提案 / 大纲 / 逐页规格 / QA 报告 |
| `examples/` | 71 | 14,363.5 KB | 3 份分析示例 + 3 个完整案例（含 pptx 与预览图） |
| `scripts/` | 15 | 493.0 KB | 6 条命令 + 3 个库 + selfcheck + README |
| `schema/` | 2 | 17.2 KB | `deck.schema.json`、`thresholds.json`（v1.1.0） |
| `docs/` | 3 | 35.8 KB | 环境报告、验证报告、发现测试 |
| 合计 | **159** | **16.87 MB** | zip 15.23 MB / sha256 `6eb8c747ce0bc224` |

### 2.2 已验证能力（都跑过，数字是产物体积）

| 通路 | 命令 | 实测产物 |
|---|---|---|
| PPTX 生成 | `build_pptx.py --deck demo/source/deck.json --out demo/build/deck.pptx` | 8 页 / 156,719 B / 字体 2 族 |
| PPTX → PNG | `render.py` 或 `office-cli render --dpi 144` | 1920×1081×8，147 KB/页量级 |
| PPTX → PDF | `office-cli convert --input deck.pptx --output deck.pdf` | 8 页 / 404,760 B / `%PDF-1.7` |
| A4 竖版版面 | `deck.json` 的 `meta.size = "8.268x11.693"` | 1 页 / 28,634 B → 渲染 1241×1754 px @150 dpi |
| DOCX 生成 → PDF | python-docx 生成 → `convert` | 37,568 B → 30,485 B / 1 页（告警 `missingFonts:["Courier"]`） |
| PDF → PNG | `office-cli render --input x.pdf --pages 1,2 --dpi 150` | pdfium 后端 / 128,460 B + 81,885 B |
| 机器质检 | `qa_check.py --deck ... --json` | 30 个 check id，demo/三案例均 `P0=0 P1=0 P2=0 P3=0` |

### 2.3 本机依赖与硬限制

- Python 3.12.14 可用包：`python-docx 1.2.0`、`python-pptx 1.0.2`、`openpyxl 3.1.5`、`lxml 6.1.3`、`Pillow 12.3.0`、`numpy 2.3.5`、`pandas 3.0.1`。
- **缺**：任何 PDF 库、`reportlab`、`psd_tools`、`cairosvg/svglib`、`matplotlib`、`fontTools`、`yaml`、`jinja2`。
- 渲染/转换后端：DSH 自带 `office-cli`（LibreOffice 内核 + pdfium），能力 = 转换（docx/xlsx/pptx ↔ pdf）+ 逐页栅格化。**不带 PDF 内容编辑**。
- 无 MS Office / WPS / COM；不许 `pip install`。
- 字体 240 个（微软雅黑、思源宋体 Heavy、Noto Sans SC、Lato、Cascadia Mono 等），转换时缺字体告警可用于校验。

### 2.4 缺口表

| 需要 | 现在有 | 差什么 |
|---|---|---|
| 文档级排版（标题层级/页眉页脚/目录/表格跨页） | 无 | `doc.schema.json` + `build_docx.py` |
| 文档级质检 | 无（`qa_check.py` 只认 slide 画布） | `qa_doc.py` + 文档阈值 |
| 打印尺寸的音量档位 | 只有 `slide-16x9` 的字号标准（正文 ≥16 pt） | `presets`：`print-A4` / `print-A3` |
| PDF 交付 | 能转换，但脚本里没封装 | `export_pdf.py`（校验页数/缺字体） |
| 矢量导出 | 无 | `export_svg.py` |
| 分层交付（给设计师继续做） | 无 | `export_layers.py` |
| 既有 PDF 改版 | 能渲成图 | `pdf_rebuild.py`（底图 + 覆盖层 + 色彩重映射） |

---

## 3. 四种产出形态的判定

### 3.1 判定顺序（每次都先走这一张表）

| 用户手上有什么 / 要什么 | 走哪种形态 | 首条命令 |
|---|---|---|
| 要讲给人听，一页一个观点，投屏 | A 演示稿 | `build_pptx.py` |
| 要发出去读，有分段、章节、目录、页码 | B 文档 | `build_docx.py` |
| 单页/多页的版面作品（海报、一页纸、画册跨页、信息图） | C 固定版面 | `build_pptx.py` + `preset=print-*` |
| 已经有一份 PDF，要改设计/改版式 | D 改版 | `pdf_rebuild.py` |
| 只改 PDF 里几行字 | **不接**，改为：要源文件（docx/keynote/figma），或按 D 重新排 | — |
| 要 .psd | 改交付 **SVG + 分层 PNG**，并说明差异 | `export_svg.py` / `export_layers.py` |

### 3.2 四种形态的产物矩阵

| 形态 | 主产物 | 附带产物 | 质检脚本 |
|---|---|---|---|
| A 演示稿 | `.pptx` | `.pdf`、逐页 `.png`、`contact-sheet.png`、`qa.md/json` | `qa_check.py` |
| B 文档 | `.docx` + `.pdf` | 逐页 `.png`、`qa.md/json` | `qa_doc.py` |
| C 固定版面 | `.pdf` + `.png` | `.pptx`（可编辑源）、`.svg`、分层 png 包 | `qa_check.py`（+PRT 项） |
| D 改版 | `.pdf` | 对照图（before/after 并排）、改动清单 `.md` | `qa_check.py` + 人工清单 |

---

## 4. 目录与文件变更

### 4.1 新增（14 个文件 / 2 个目录）

| 路径 | 职责（一句话） |
|---|---|
| `schema/doc.schema.json` | 文档模型：`meta` + `tokens` + `sections[].blocks[]`，12 种块 |
| `schema/presets.json` | 四档 preset 的阈值绑定（slide / A4 / A3 / letter），阈值仍只有这一处真值 |
| `scripts/build_docx.py` | `doc.json` → `.docx`（样式注入、TOC 域、页眉页脚、表格跨页、图片缩放） |
| `scripts/export_pdf.py` | `.docx`/`.pptx` → `.pdf`，校验页数、缺字体、文件可打开 |
| `scripts/qa_doc.py` | 文档级质检 15 项 + 复用图片像素密度检查 |
| `scripts/export_svg.py` | 一页 `deck.json` → 单文件 `.svg`（文字仍是文字，图片内联 base64） |
| `scripts/export_layers.py` | 一页/一套 → 分层 PNG 包（`layers/*.png` + `layers.json`，可被 PS 逐层导入） |
| `scripts/pdf_rebuild.py` | 既有 PDF → 逐页底图 → 新 `deck.json` → 新 PDF（含色彩重映射与改动清单） |
| `references/document-typography.md` | 文档排版规范（字号/行距/行长/页边/页眉页脚基线/图表编号） |
| `references/document-structure.md` | 文档结构规范（章节深度、目录、表/图/式编号、参考文献、附录） |
| `templates/doc-brief.md` | 文档任务简报模板 |
| `templates/doc-outline.md` | 文档大纲模板（章节 → 每节一句话任务 → 预估页数） |
| `workflows/05b_document-production.md` | 阶段 5 的文档分支（块级制作顺序、每页放什么、表格与图注怎么落） |
| `examples/docs/<case>/` | 3 个文档案例目录（结构同现有 slide 案例） |

### 4.2 修改（7 个文件，都只加段不改旧行为）

| 文件 | 改哪一段 | 加什么 |
|---|---|---|
| `SKILL.md` | 第 2 节「任务形态」 | 加"产物形态判定表"（第 3.1 节的表），并指向 `05b` |
| `SKILL.md` | 第 6 节「硬指标」 | 加一行"文档形态的硬指标见 `references/document-typography.md`" |
| `SKILL.md` | 路由表 | 加 4 行：要 docx / 要 pdf / 要 svg / 要改 PDF |
| `schema/thresholds.json` | 顶层加 `presets` 指针 + `print` 段 | 打印字号与边距阈值（第 7 章），顶层结构保持兼容 |
| `scripts/README.md` | 命令表 + check id 表 | 加 5 条命令、18 个 check id |
| `README.md` | 第 1 节 + 结构树 | 一句话说明"三种形态"，结构树补新文件 |
| `CHANGELOG.md` | 版本段 | `[1.2.0]` / `[2.0.0]` / `[2.1.0]` 三段 |

### 4.3 不动（红线）

- `deck.schema.json` 的字段与 `build_pptx.py` 的输出行为：**一字节都不改**。唯一例外：允许新增**可选**参数 `--preset`，不带该参数时输出必须与 v1.1.2 逐字节一致（比对方法见 [v2-plan-ops.md §4.2](v2-plan-ops.md)）。
- 现有 30 个 check id 的编号、判据与分级：新增项只能加在 `PRT-*` / `DOC-*` / `PDF-*` / `EXP-*` 命名空间。
- demo 与三个 slide 案例的产物：v2.0 完成后重新回归一次，QA 必须仍是 `0/0/0/0`。

### 4.4 v2.0 完成后的目录树（只标新增）

```text
ai-ppt-design-system/
├── SKILL.md                        # 改：加产物形态判定表 + 4 条路由
├── README.md  install.md  CHANGELOG.md  LICENSE  CONTRIBUTING.md
├── discovery/                      # 不动（5）
├── references/                     # 11 + 2（document-typography / document-structure）
├── workflows/                      # 8 + 1（05b_document-production）
├── templates/                      # 6 + 2（doc-brief / doc-outline）
├── schema/
│   ├── deck.schema.json            # 不动
│   ├── doc.schema.json             # 新增
│   ├── thresholds.json             # 改：加 print 段
│   └── presets.json                # 新增
├── scripts/                        # 15 + 6 → 21（build_docx / export_pdf / qa_doc / export_svg / export_layers / pdf_rebuild）
├── examples/
│   ├── cases/                      # 不动（3 个 slide 案例）
│   └── docs/                       # 新增 3 个文档案例
├── docs/
│   ├── environment-report.md  validation-report.md  discovery-tests.md   # 不动
│   ├── v2-plan.md                  # 本文件
│   └── v2-validation-report.md     # 新增（V9–V14 验证矩阵）
└── dist/                           # 发布包
```

---

## 5. 文档数据模型（`doc.schema.json`）

### 5.1 骨架

```json
{
  "meta": {
    "title": "某园区冬季遮阴实测报告",
    "subtitle": "6 周观测 · 42 个测点",
    "author": "林晓远",
    "org": "某高校建筑学院",
    "date": "2026-11",
    "lang": "zh-CN",
    "version": "v1.0",
    "pageSize": "A4",
    "orientation": "portrait",
    "marginsIn": { "top": 0.9, "bottom": 0.9, "left": 1.0, "right": 1.0 },
    "toc": { "enabled": true, "depth": 3 },
    "numbering": { "figures": true, "tables": true, "headings": true },
    "header": { "text": "某园区冬季遮阴实测报告", "showOnFirstPage": false },
    "footer": { "pageNumber": true, "text": "" }
  },
  "tokens": {
    "fonts": { "display": "Source Han Serif SC", "body": "Microsoft YaHei", "mono": "Cascadia Mono" },
    "colors": { "fg": "#1A1A1A", "bg": "#FFFFFF", "fgMuted": "#5F5F5F", "accent": "#8C3A1E", "rule": "#D8D3C8" },
    "type": { "h1": 20, "h2": 15, "h3": 12.5, "body": 10.5, "caption": 9, "table": 9.5, "code": 9, "footnote": 8 }
  },
  "sections": [
    { "id": "s1", "blocks": [ { "type": "heading", "level": 1, "text": "一、研究问题" } ] }
  ]
}
```

### 5.2 meta 字段约束

| 字段 | 类型 | 必填 | 约束 |
|---|---|---|---|
| `title` | string | 是 | ≤30 字（超长触发 `DOC-03`） |
| `pageSize` | string | 是 | `A4` / `A3` / `Letter` / 显式 `"WxH"`（英寸，如 `"8.268x11.693"`） |
| `orientation` | enum | 是 | `portrait` / `landscape` |
| `marginsIn` | object | 是 | 每个边 0.5–1.6 in；与 preset 不一致时告警 |
| `toc.enabled` | bool | 否 | true 时生成器插入 TOC 域，并在交付说明里写"打开后按 F9 更新目录" |
| `numbering.figures/tables` | bool | 否 | true 时图注写成"图 1 · 冠层剖面示意"这种格式，编号由生成器按出现顺序算 |
| `header` / `footer` | object | 否 | 页码用 PAGE 域；首页是否显示单独控制 |

### 5.3 12 种文档块

| `type` | 关键字段 | 渲染为 |
|---|---|---|
| `heading` | `level`(1–3)、`text` | Word 标题 1–3 样式，进目录 |
| `paragraph` | `runs[]`、`align`、`indentFirstLine`、`spaceAfterPt` | 正文段落，中文首行缩进 2 字符（可关） |
| `list` | `ordered`(bool)、`items[][]`（每项是 runs）、`level` | 项目符号 / 编号列表，编号连续由生成器维护 |
| `table` | `caption`、`header[]`、`rows[][]`、`colWidthsIn[]`、`align[]`、`zebra` | Word 表格，表头跨页重复（`tblHeader`），行不跨页断开（`cantSplit`） |
| `figure` | `src`、`widthIn`、`caption`、`source`、`alt` | 图片 + 图注 + 来源行；宽度缺省按正文栏宽等比缩放 |
| `quote` | `text`、`cite` | 引文块（缩进 + 左侧竖线） |
| `callout` | `kind`(`note`/`warn`/`key`)、`title`、`text` | 单格表格做的提示框（浅底 + 细边） |
| `code` | `lang`、`text` | 等宽字体块（10 pt，灰底） |
| `pageBreak` | — | 硬分页 |
| `toc` | — | 目录域占位（位置由生成器插入） |
| `bibliography` | `entries[]` | 参考文献列表（编号 `[1]`，与正文引用号人工对齐） |
| `footnoteGroup` | `id` | **P2 阶段才做**；先降级为章末"注释" |

### 5.4 与 `deck.schema.json` 的本质差异

| 维度 | deck（演示稿） | doc（文档） |
|---|---|---|
| 定位方式 | 绝对坐标（`x/y/w/h` 英寸） | 流式（块顺序 + 段落/表格/图片） |
| 分页 | 由人决定（每页元素固定） | 由渲染器决定（会流动、会断页） |
| 质检重心 | 溢出、越界、重叠、密度 | 层级、跨页、目录、页码、字号档位 |
| 失败模式 | 框太小装不下 | 断页断在表格中间、图注与图分离、目录页码不符 |
| 谁负责美观 | 我们（坐标就是设计） | 样式表 + 上图下图注的间距（生成器只保证规则） |

### 5.5 三条模型级约束（写成规则的）

**规则：块顺序即阅读顺序，不许用坐标把内容挪到别处**
- **触发条件**：需要在文档里做"左右分栏""图上文下"之类的特殊摆放。
- **具体动作**：用 `table`（两列各放一块内容）或 `figure` 的 `widthIn` 控制，不要引入绝对坐标。
- **常见错误**：在 doc 模型里塞 `x/y`，结果是 DOCX 里出现浮动文本框，Word 打开后错位、也无法回流。

**规则：图、表必须有编号、题注与来源**
- **触发条件**：插入任何 `figure` 或 `table`。
- **具体动作**：`caption` 写"图 N + 一句话结论式题注"，`source` 写数据出处与口径；缺任一项 `qa_doc.py` 判 `DOC-05`/`DOC-06`（P0）。
- **常见错误**：把来源塞在正文里、或写"如图所示"却不给编号，读者无法回指。

**规则：标题层级只能加深一级**
- **触发条件**：写第 N 级标题。
- **具体动作**：从 h1 到 h2 到 h3 依次加深，跳过即 `DOC-01`（P0）。
- **常见错误**：为了字号好看直接 h1 → h3，导致目录塌成两层、Word 导航窗格断链。

---

## 6. 脚本契约（七条命令）

### 6.1 总表

| 命令 | 输入 | 输出 | 退出码 |
|---|---|---|---|
| `build_docx.py --doc D --out O [--check-only]` | `doc.json` | `.docx` | 0 成功 / 1 有 P0 级模型错误 / 2 参数或 I/O 错 |
| `export_pdf.py --input F --out O [--pages all|"1-3"]` | `.docx`/`.pptx` | `.pdf` + `pdf-manifest.json` | 0 / 1 转换成功但缺字体 / 2 失败 |
| `qa_doc.py --doc D --out O [--json J]` | `doc.json` | `qa.md` + `qa.json` | 0 无 P0 / 1 有 P0 |
| `export_svg.py --deck D --slide N --out O` | `deck.json` | `.svg` | 0 / 1 有元素类型不支持（跳过并记录） / 2 失败 |
| `export_layers.py --deck D --slide N --out-dir O` | `deck.json` | `layers/*.png` + `layers.json` | 0 / 1 部分层为空 / 2 失败 |
| `pdf_rebuild.py --pdf P --spec S --out-dir O` | 原 PDF + `rebuild.json` | 新 PDF + 对照图 + `changes.md` | 0 / 1 有页面未处理 / 2 失败 |
| `render.py`（现状） | `.pptx`/`.pdf` | 逐页 `.png` | 0 / 2 无渲染器 |

### 6.2 `build_docx.py` 关键实现约定

1. **样式注入而非手工排版**：只改 `Normal` / `Heading 1–3` / 自定义 `FigCaption`、`TblCaption`、`Source`、`Quote`、`Callout`、`Code` 七个样式；段落只引用样式，不逐段设字体。这样用户在 Word 里改样式能整套生效。
2. **中文字体必须双写**：`run.font.name`（西文）+ `rPr.rFonts` 的 `w:eastAsia`（中文），否则 Word 里中文回落到宋体（实测过，这是最常见的"字体没生效"）。
3. **行距与段后**：正文行距 1.5 倍，段后 6 pt，首行缩进 2 字符（`paragraph_format.first_line_indent = Pt(21)` 对 10.5 pt 正文）。
4. **目录**：插 `w:fldSimple` + `instr='TOC \o "1-3" \h \z \u'`，页面上留一行"目录（打开后按 F9 更新）"；**不假装目录页码已经正确**。
5. **页眉页脚**：页脚居中 `PAGE` 域 + 可选 `NUMPAGES`；页眉用 `different_first_page_header_footer` 控制首页不显示。
6. **表头跨页**：给首行 `<w:tblHeader/>`，并给所有行 `<w:cantSplit/>`，防止表格在页底断成两半（这是文档形态最常见的低级错误）。
7. **图片**：`widthIn` 缺省 = 正文栏宽（A4 竖版 6.27 in），按原图比例算高；超过一页高度时按页高 80% 反算宽度，并在 stderr 提示"图片过大已缩放"。
8. **编号**：图/表编号由生成器按出现顺序算成纯文本（`图 1`、`表 2`），**不用 SEQ 域**——理由见 1.2。

### 6.3 `export_pdf.py` 校验项

| 校验 | 判据 | 失败时 |
|---|---|---|
| 输出存在且 >1 KB | 文件大小 | 退出码 2，打印后端原始报错 |
| 页数 >0 | manifest 的 `pageCount` | 退出码 2 |
| 缺字体列表 | `missingFonts == []` | 退出码 1，列出哪些字体 + 建议替换（微软雅黑/思源宋体/Noto Sans SC/Lato 可用） |
| 可再次栅格化 | 用自己的 `render` 渲第 1 页成功 | 退出码 1（PDF 可能损坏，但能打开） |

### 6.4 `qa_doc.py` 的 15 项（编号即判据）

| ID | 检查 | 级别 | 判据（数字来自 `presets.json`） |
|---|---|---|---|
| `DOC-01` | 标题层级跳级 | P0 | 出现 h1→h3 直接判 |
| `DOC-02` | 无 h1 或标题数 <2 | P1 | 文档型必须有章节结构 |
| `DOC-03` | 标题字数超限 | P2 | title >30 字 / heading >24 字 |
| `DOC-04` | 段落过长 | P2 | 单段 >300 字 |
| `DOC-05` | 图表缺题注 | P0 | `figure`/`table` 的 `caption` 为空 |
| `DOC-06` | 图表缺来源 | P0 | `source` 为空 |
| `DOC-07` | 图片有效像素密度 | P1/P2 | <80 ppi 报 P1、<110 报 P2（复用 `IMG-01` 逻辑） |
| `DOC-08` | 字体族超限 | P1 | >3 族 |
| `DOC-09` | 正文字号不在档 | P1 | 不在 preset 的 body 区间（A4 打印 9–11 pt） |
| `DOC-10` | 行距不在档 | P2 | 正文 1.4–1.6 之外 |
| `DOC-11` | 页边距不在档 | P2 | 与 preset 差 >0.1 in |
| `DOC-12` | 目录与标题数不符 | P1 | `toc.enabled` 时，目录条目数 ≠ h1–h3 总数 |
| `DOC-13` | 页码缺失 | P1 | `footer.pageNumber=true` 但生成文件里找不到 PAGE 域 |
| `DOC-14` | 表格列宽和≠正文宽 | P2 | 差 >0.15 in |
| `DOC-15` | 中英文之间缺空格 | P3 | 启发式，仅提示 |

`PRT-01`/`PRT-02`（给固定版面形态用）：`PRT-01` 打印边距不足（A4 <0.6 in、A3 <0.8 in，P1）；`PRT-02` 出血元素缺失（满版图未设 `zRole:"bgDecor"`，P2）。

`PDF-01` 转换页数与源不符（P1）、`PDF-02` 缺字体（P1）、`PDF-03` 页面尺寸不符 preset（P1）。
`EXP-01` SVG 元素越界（P1）、`EXP-02` 分层包缺 `layers.json` 或层为空（P1）、`EXP-03` SVG 中文字被转成路径（P2，允许但需在 README 说明）。

### 6.5 `export_svg.py` 映射表

| `deck.json` 元素 | SVG |
|---|---|
| `rect`（含 `corners`） | `<rect>` + `rx` |
| `line` | `<line>`（`dash` → `stroke-dasharray`） |
| `ellipse` | `<ellipse>` |
| `polygon` | `<polygon points>` |
| `text`（runs/paragraphs） | `<text>` + `<tspan>`，`font-family`/`font-size`/`letter-spacing`/`line-height` 逐项落；`textLength` 不用（会拉伸字形） |
| `image` | `<image href="data:image/png;base64,...">`，内联保证单文件可移植 |
| `chart`/`table` | **不支持**（v2.0 明确跳过并在 stderr 列出），需要时先 `render` 成图再当 `image` |

坐标换算：`px = inch × 96`（SVG 用户单位），`viewBox="0 0 1280 720"`（16:9）或按画布算。

### 6.6 `export_layers.py` 分组规则

按 `zRole` 与元素类型分五层，顺序即绘制顺序：`01-bg`（背景色块/满版图）、`02-decor`（装饰线/色块）、`03-images`（图片）、`04-charts`（图表，渲染成位图）、`05-text`（文字，栅格化）。`layers.json` 记录每层的 `x/y/w/h`、不透明度、来源元素 id。

**诚实说明**：分层 PNG 的文字层是位图，进 PS 后不能再改字；要改字请用 SVG。

### 6.7 `pdf_rebuild.py` 流水线（形态 D）

```text
原 PDF ──render(150dpi)──→ pages/p1..pN.png
        │
        ├─ 读 rebuild.json：每页 { keep: full|none, mask:[{x,y,w,h,fill}], overlay:[元素],
        │                          recolor:[{from:"#RRGGBB", to:"#RRGGBB", tolerance:0..40}] }
        │
        ├─ 底图：PIL 做色相/色值替换（tolerance 用欧氏距离）+ 遮挡矩形填充
        ├─ 覆盖层：元素按 deck.json 语法写，走同一套 build_pptx → convert
        └─ 输出：build/rebuild.pdf + compare/p1-before.png|p1-after.png + changes.md
```

**能力边界（写进脚本 `--help` 与 `README`）**：文字被栅格化，改完是"重新设计的版本"；原文字不可选、不可搜；要改文字只能要源文件。

### 6.8 错误处理表

| 情况 | 处理 | 退出码 |
|---|---|---|
| `doc.json` 不合 schema | 打印首个错误路径（`$.sections[2].blocks[4].type`）+ 修法提示 | 2 |
| 字体在本机不存在 | 换成同角色可用族（display→思源宋体/雅黑，body→雅黑，mono→Cascadia Mono），写进 manifest | 0（带 warning） |
| 图片缺失 | 保留占位框 + 题注，写进 `missing` 列表 | 1 |
| TOC 域未更新 | 交付说明里固定一句"打开后按 F9 更新目录" | 0 |
| 转换缺字体 | 列出字体名 + 建议替换 | 1 |
| PDF 页面尺寸与 preset 差 >0.05 in | 报 `PDF-03` | 0（P1 不阻断） |

---

## 7. 阈值与 preset

### 7.1 四档 preset

| preset | 画布 | 正文 | 行距 | 页边 | 行长 | 每页字数 |
|---|---|---|---|---|---|---|
| `slide-16x9`（现状） | 13.333×7.5 in | 16–24 pt | 1.25–1.55 | 0.6/0.45 in | ≤34 汉字 | ≤220 |
| `print-A4` | 8.268×11.693 in | 9–11 pt | 1.4–1.6 | 0.8–1.1 in | ≤48 汉字 | 450–900 |
| `print-A3` | 11.693×16.535 in | 10–13 pt | 1.4–1.6 | 1.0–1.4 in | ≤60 汉字 | 500–1200 |
| `print-letter` | 8.5×11 in | 9–11 pt | 1.4–1.6 | 0.8–1.1 in | ≤48 汉字 | 450–900 |

### 7.2 文档专用阈值（新增段）

| 键 | 值 | 含义 |
|---|---|---|
| `doc.titleMaxChars` | 30 | 文档标题上限 |
| `doc.headingMaxChars` | 24 | 任一级标题上限 |
| `doc.paragraphMaxChars` | 300 | 单段上限（超过应拆段） |
| `doc.bodySizeRange` | `{A4:[9,11], A3:[10,13]}` | 正文字号区间 |
| `doc.lineSpacingRange` | `[1.4, 1.6]` | 正文行距区间 |
| `doc.maxFontFamilies` | 3 | 字体族上限 |
| `doc.figureCaptionMaxChars` | 60 | 题注上限 |
| `doc.tableMaxCols` | 6 | 表格列数上限（A4 竖版） |

### 7.3 组织方式（重要：不破坏"唯一真值"）

- 所有数字仍只写在 `schema/thresholds.json`；`schema/presets.json` **只做引用**（`{"print-A4": {"canvas": "canvas.print.A4", "type": "type.print", ...}}`），不再复制一份数字。
- `thresholds.json` 新增 `presets` 与 `print` 两段，`version` → `2.0.0`。
- 读取顺序：**显式命令行 `--preset` > `meta.pageSize` 推导 > 默认 `slide-16x9`**，且脚本必须在报告头部打印实际生效的 preset（`preset: print-A4 (source: meta.pageSize)`）。

---

## 8. QA 标准（文档形态）

### 8.1 分级与阻断

沿用现有规则：`P0` 阻断交付（`blockOnP0=true`），`P1` 必须修或书面接受，`P2` 建议修，`P3` 仅记录。文档形态额外一条：**P0 里包含"事实类"问题**（图表缺来源、数据无单位、引用号对不上），这类不做"接受"选项。

### 8.2 人工判定项（脚本不实现，写进 `qa.md` 固定段落）

| 项 | 谁判 | 判据 |
|---|---|---|
| 章节顺序是否符合读者的提问顺序 | 人 | 每章标题能回答上一章留下的问题 |
| 表格数据是否与正文数字一致 | 人 | 抽查 3 处，逐一对照 |
| 目录页码是否与正文一致 | 人 | 打开 PDF 抽查 3 处 |
| 图注是否描述"结论"而非"图里有什么" | 人 | 题注应含判断，不只是"示意图" |

### 8.3 报告格式

复用现有的 `qa.md` + `qa.json` 结构与 9 列表格（编号 / 页码 / 类别 / 严重程度 / 检查项 / 问题描述 / 修复建议 / 状态 / 复查），头部多一行 `preset:`；`qa.json` 的 `summary` 增加 `preset` 与 `pageCount` 字段。

---

## 9. 分阶段实施

> 本章只写「做什么、什么算完」。**每阶段逐条怎么执行**（文件清单、创建顺序、S1..Sn 步骤、命令、期望输出、排错表、验收证据）见 [v2-plan-ops.md](v2-plan-ops.md)：那是操作书，本文是规划书。

### 9.1 总览

| 阶段 | 内容 | 规模 | 依赖 | 落到哪个版本 |
|---|---|---|---|---|
| **P0** | 文档骨架：schema + build_docx 最小闭环 + export_pdf + qa_doc 前 6 项 | 中（新增约 900–1200 行） | 无 | v2.0.0 |
| **P1** | 文档完整能力：表格跨页、目录、页眉页脚、图注编号、分节、callout/code | 中（+800 行） | P0 | v2.0.0 |
| **P2** | 设计型导出：`export_svg.py`、`export_layers.py`、`print-*` preset、PRT 项 | 小-中（+600 行） | P0 的 preset 段 | v1.2.0 |
| **P3** | 既有 PDF 改版：`pdf_rebuild.py`（底图 + 遮挡 + 覆盖 + 色彩重映射） | 中-大（+700 行） | P2（要覆盖层能力） | v2.1.0 |
| **P4** | 案例与验收：3 个文档案例 + `docs/v2-validation-report.md`（V9–V14） | 中（产物多） | P0–P3 | 随各自版本 |

### 9.2 P0 · 文档骨架（先证明"能出文档"）

**交付**：`schema/doc.schema.json`、`scripts/build_docx.py`（仅 heading/paragraph/list/table/figure/pageBreak 六种块）、`scripts/export_pdf.py`、`scripts/qa_doc.py`（DOC-01/02/05/06/08/12/13 七项）、`examples/docs/report-min/`（6 页 A4，含封面、两章、一张表、一张图、目录）。

**验收（四条全中才算过）**
1. `python scripts/build_docx.py --doc examples/docs/report-min/source/doc.json --out build/report.docx` → 0，`.docx` >30 KB，用 python-docx 复读能拿到 6 个段落样式名。
2. `python scripts/export_pdf.py --input build/report.docx --out build/report.pdf` → 0，页数 5–7，`missingFonts` 为空。
3. `python scripts/qa_doc.py --doc … --out build/qa.md --json build/qa.json` → `P0=0`。
4. 逐页 PNG 目视：标题层级、表格、图注、页码都对（我给截图）。

**规模诚实说**：这是一个新脚本 + 一个新 schema + 一个案例，不是"改几行"。

### 9.3 P1 · 文档完整能力

**交付**：剩余 5 种块（quote/callout/code/toc/bibliography）、表格跨页表头、图片缩放规则、页眉页脚分节、图注自动编号、`references/document-typography.md`、`references/document-structure.md`、`workflows/05b_document-production.md`、`templates/doc-brief.md`、`templates/doc-outline.md`、`SKILL.md` 路由与判定表。

**验收**
1. 同一份 `doc.json` 改 `tokens`（字体 3 种 × 主色 4 种，共 12 组）重建 12 次，`qa_doc.py` 均 `P0=0`，无版面塌陷（逐组抽 1 页目视）。
2. 一份 30 页长文档：目录条目数 = 标题数，页码连续，表格跨页处表头重复（抽 3 页目视）。
3. `qa_doc.py` 15 项全部被真实触发过至少一次（用两份 fixture，同 slide 分支的做法）。

### 9.4 P2 · 设计型导出与打印 preset

**交付**：`schema/presets.json`、`thresholds.json` 的 `print` 段、`export_svg.py`、`export_layers.py`、`PRT-01/02`、A4 海报案例一份（`examples/docs/poster-a4/`）。

**验收**
1. A4 海报 → `build.pdf` 1 页 + `build.png` 1241×1754 @150 dpi，`qa_check.py` `P0=0`。
2. SVG：浏览器打开，`Ctrl+F` 能搜到中文标题（证明文字未被转路径），与 PNG 叠图位置差 ≤2 px。
3. 分层包：`layers.json` 五层齐全，PS 手动导入一次（我截图），文字层位图、说明已写入 README。

### 9.5 P3 · 既有 PDF 改版（能力受限，先把边界写死）

**交付**：`scripts/pdf_rebuild.py`、`examples/docs/pdf-rebuild-min/`（拿我们自己的一份 PDF 当输入，改成另一套 tokens）。

**验收**
1. 输入 8 页 PDF → 输出 8 页 PDF，页序与页数一致。
2. 底图 150 dpi（`1241×1754` 量级），无叠影：`compare/pN-before.png` 与 `after.png` 并排，人工确认改动区域清单与 `changes.md` 一致。
3. 色彩重映射：`recolor` 把主色从 `#C8501E` 换到 `#1F3B33`，容差 24，目标像素替换率 >95%（脚本统计并写进 `changes.md`）。
4. `--help` 里第一段就是能力边界声明（文字不可编辑）。

### 9.6 P4 · 案例与验证报告

**交付**：三个文档案例（`technical-report` 12 页 / `research-brief` 6 页 / `one-pager` 2 页），各带 `brief.md`、`source/doc.json`、`build/`、`README.md`（节奏表 + 真实数字 + 诚实限制 + 检查清单）；`docs/v2-validation-report.md` 记 V9–V14：V9 文档最小闭环、V10 长文档 30 页、V11 A4 海报、V12 SVG/分层、V13 PDF 改版、V14 十二组 tokens 变体。

**验收**：三案例 `qa_doc.py` 均 `P0=0`；验证报告每条都有命令 + 产物字节 + QA 摘要；`selfcheck.py --strict` 归零；重新打包 v2.0.0 zip 并解包复验。

### 9.7 每阶段的"不达标就停"规则

**规则：一阶段验收不过，不进下一阶段**
- **触发条件**：阶段验收任一条不满足。
- **具体动作**：停下来补，不平行推进；把失败写进 `docs/v2-validation-report.md` 的"未通过项"。
- **常见错误**：为了"看起来有进度"先做下一阶段的文件，导致后一阶段建立在错的模型上（v1.1 阶段 6 曾因此返工一次）。

---

## 10. 验收总表（我看这张表决定是否签收）

| 编号 | 验收项 | 判据 | 证据形式 | 状态 |
|---|---|---|---|---|
| A1 | 文档最小闭环 | 6 页 doc → docx + pdf，`P0=0` | 命令输出 + 产物字节 + 3 张截图 | 待做 |
| A2 | 同源双出页码一致 | `PDF 实际页数 = 目录最后一条页码`（允许 ±1），并抽查 3 处标题 | pageCount + 3 处抽查 | 待做 |
| A3 | 长文档 30 页 | 目录数=标题数、页码连续、表格跨页表头重复 | 3 页截图 + qa.md | 待做 |
| A4 | 15 项文档检查全覆盖 | 每项至少被真实触发一次 | 两份 fixture 的 qa.md | 待做 |
| A5 | A4 海报 | 1 页 PDF + 1241×1754 PNG，`P0=0` | 产物 + 截图 | 待做 |
| A6 | SVG 文字可搜 | `Ctrl+F` 命中标题 | 截图 + 叠图比对 | 待做 |
| A7 | 分层包可用 | 五层齐全、PS 导入成功 | 截图 + layers.json | 待做 |
| A8 | PDF 改版 | 页数一致、无叠影、改动清单对得上 | before/after 对照 + changes.md | 待做 |
| A9 | 三案例 | 各自 `P0=0` | 三份 qa.md | 待做 |
| A10 | 全库自检 | `selfcheck.py --strict` 全 0 | 命令行输出 | 待做 |
| A11 | 发布包 | 159+N 文件、解包复验通过、sha256 | 打包日志 | 待做 |
| A12 | 现有 PPT 未被破坏 | demo + 3 个 slide 案例 QA 仍 `0/0/0/0` | 四份 qa.md | 待做 |

---

## 11. 风险与回退

| 编号 | 风险 | 触发信号 | 回退动作 |
|---|---|---|---|
| R1 | 无 PDF 库导致"改 PDF 文字"期望落空 | 用户要"把第 3 页那句改掉" | 明确回绝 + 给两条替代（要源文件 / 按 D 重排），不假装能做到 |
| R2 | python-docx 无脚注 API | 需要页脚注 | 降级为章末"注释"块（P2 再做）；不硬改 OOXML |
| R3 | LibreOffice 分页与 Word 不一致 | docx 在 Word 里页数差 2 页以上 | 交付以 PDF 为准；README 写明"docx 打开可能 ±1 页，页码以 PDF 为准" |
| R4 | 转换缺字体 | `missingFonts` 非空 | 替换为可用族（雅黑/思源宋体/Noto Sans SC/Lato）+ 记录替换表 |
| R5 | 大文档性能（>60 页） | build 超过 60 秒 | 分节构建 + 只重建改动段；不做"一次生成 200 页"的承诺 |
| R6 | PSD 需求无法满足 | 用户坚持要 `.psd` | 交付 SVG + 分层 PNG + PS 导入 JSX，并在 README 说明差异 |
| R7 | 表格跨页断行难看 | 抽页发现表格被截断 | `tblHeader` + `cantSplit` + 必要时手动 `pageBreak` 前置 |
| R8 | 改 PDF 后叠影/重影 | 底图与新元素位置错位 | 以底图 PNG 为基准做像素级对齐（`--align-to 150`），错位 >2 px 判失败 |

---

## 12. 版本与发布

### 12.1 版本序列

| 版本 | 内容 | 语义 |
|---|---|---|
| `v1.2.0` | P2（SVG / 分层 PNG / print preset） | 加能力，不改既有行为 |
| `v2.0.0` | P0 + P1（文档分支） | 技能范围从"演示稿"扩到"文档"，属破坏性扩范围 |
| `v2.1.0` | P3（PDF 改版） | 新增受限形态 |
| `v2.1.x` | 修 bug、补检查项 | 补丁 |

### 12.2 发布步骤（每次发布都走这七步）

1. 全库 `selfcheck.py --strict` → `P0 0 / P1 0 / P2 0`。
2. `demo` + 3 个 slide 案例回归：`qa_check.py` 四份均 `0/0/0/0`。
3. 新形态回归：对应阶段验收全部通过（第 10 章）。
4. `CHANGELOG.md` 写版本段（做什么 / 为什么 / 怎么验）。
5. 打包 `dist/ai-ppt-design-system-v<X>.zip`，排除 `__pycache__`、`.git`、`dist`。
6. **解包复验**：解到临时目录 → 跑 selfcheck → 跑一次 build + 一次 QA（不能只信打包前）。
7. 更新 README 的"当前版本"与文件数，删除旧 zip。

### 12.3 CHANGELOG 规则

- 每条写"改了什么 / 为什么 / 怎么验证"，不写"优化了体验"这类无法核对的句子。
- 破坏性变化（如 v2.0.0 扩范围）必须在标题下第一段说明影响面：现有 slide 用户是否需要动任何文件（答案：不需要）。

---

## 附录 A · 新增术语

| 术语 | 定义 |
|---|---|
| 同源双出 | 一份 `doc.json` 产出 docx + pdf，页码一致 |
| 块级制作 | 文档形态下按"块"而非按"页"推进（页面由渲染器决定） |
| 底图重排 | 既有 PDF 渲成 PNG 当底，新的元素画在其上，产出新 PDF |
| 分层包 | 一组透明 PNG + `layers.json`，可在 PS 里逐层导入 |

## 附录 B · 六条硬边界（v2.0 期间不可越过）

1. **不改 PDF 里已有的文字**，只重排/遮挡/覆盖。
2. **不写 .psd 二进制**，改交付 SVG + 分层 PNG。
3. **不改既有 slide 的行为**：`deck.schema.json` 与 `build_pptx.py` 输出保持一字节不差（新增字段只能可选、默认关闭）。
4. **不假装目录页码已经正确**：TOC 是域，交付说明必须写"按 F9 更新"。
5. **不为文档形态虚构"自动排版美学"**：生成器只保证规则（层级、间距、编号、跨页），美观仍由知识库与人工复核负责。
6. **没有实测证据的能力不写进 README**：每条能力都要有命令、产物字节、QA 摘要。

## 附录 C · 操作序列（复制即用）

**C1 出一份文档（docx + pdf）**
```bash
python scripts/build_docx.py --doc source/doc.json --out build/report.docx
python scripts/export_pdf.py --input build/report.docx --out build/report.pdf
python scripts/qa_doc.py --doc source/doc.json --out build/qa.md --json build/qa.json
office-cli render --input build/report.pdf --output-dir build/preview --dpi 150
```

**C2 出一页 A4 海报（固定版面）**
```bash
python scripts/build_pptx.py --deck source/deck.json --out build/poster.pptx --preset print-A4   # 新增可选参数；不带时输出与 v1.1.2 一致
python scripts/export_pdf.py --input build/poster.pptx --out build/poster.pdf
python scripts/qa_check.py --deck source/deck.json --out build/qa.md --json build/qa.json --preset print-A4
```

**C3 出 SVG / 分层包（给设计师）**
```bash
python scripts/export_svg.py --deck source/deck.json --slide 1 --out build/p1.svg
python scripts/export_layers.py --deck source/deck.json --slide 1 --out-dir build/layers-p1
```

**C4 改既有 PDF（受限形态）**
```bash
python scripts/pdf_rebuild.py --pdf input/original.pdf --spec source/rebuild.json --out-dir build
# → build/rebuild.pdf、build/compare/pN-before.png|after.png、build/changes.md
```

**C5 发布**
```bash
python scripts/selfcheck.py --strict
python -c "..."   # 打包（见 12.2 第 5 步）
```

## 附录 D · 我建议你先看的三处决策点

1. **第 1.2 节"不做"**：尤其"不改 PDF 文字"与"不写 .psd"，这两条决定你以后能不能拿它接这类活。若其中一条是硬需求，请现在推翻我，我换方案（例如：接受"装库"或"接外部工具"）。
2. **第 9.1 节的顺序**：P0→P1→P2→P3→P4。也可以先 P2（设计型 + print preset，两周内能见到 SVG 与 A4 海报）再回头做文档。**我推荐按 9.1 原序**，因为文档是你这次的主要诉求。
3. **第 7.1 节的打印字号档**（A4 正文 9–11 pt、行距 1.4–1.6）：这是排版惯例值，你可以按你的行业习惯调（比如正文统一 10.5 pt）。

---

## 检查清单

- [ ] 第 1.2 节"明确不做"的五项，与你实际需求没有冲突（尤其 PDF 文字编辑、PSD 二进制）
- [ ] 第 9.1 节的阶段顺序你认可（或指定调整）
- [ ] 第 7.1 节的四档 preset 数值你认可（A4 正文 9–11 pt、页边 0.8–1.1 in）
- [ ] 第 10.1 节验收总表 12 条，没有你想删或想加的项
- [ ] 第 4.2 节列出的 7 个"要改的现有文件"，改动范围你能接受（只加段、不改旧行为）
- [ ] 文档案例选型（技术报告 / 研究摘要 / 一页纸）覆盖你的真实使用场景
- [ ] 版本号策略（v1.2.0 / v2.0.0 / v2.1.0）与你的发布节奏一致
- [ ] 你确认本规划可以开工，且知道 P0 完成会先给你看一次 6 页样张再继续
