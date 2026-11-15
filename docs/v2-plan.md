# AI PPT 设计系统 · v2 补全规划书（R2）

| 项 | 内容 |
|---|---|
| 文档 | 规划书（决策层）。执行层是 [v2-plan-ops.md](v2-plan-ops.md)，两份按同一套编号交叉引用 |
| 版本 | R2（2026-11-15）。R1 经外部审阅后修订，逐条处理见 [v2-revision-notes.md](v2-revision-notes.md) |
| 依据 | 你给的 v1.1 实施框架 + 本次需求（PDF 从零做 / 改既有 PDF / 兼顾 Word / 偏设计方向的 PNG·PSD）+ 外部审阅意见 |
| 现状基线 | 本仓库 **v1.4.0**：177 文件 / 17.13 MB / 42 个 check id（`qa_check` 30 + `copy_check` 12），PPT 与文案分支已端到端验证 |
| 读者 | 你（决策与验收）、我（执行）、将来接手的人（复现） |
| 状态 | **等你看完点头才开工**；开工前不改任何脚本行为 |

---

## 0. 这份文档怎么用

### 0.1 它是什么

不是需求清单，是**可执行的规划**：每一阶段写清「交付什么文件 → 用什么命令产出 → 拿什么判据验收 → 不达标就停」。你看完只需做三件事：确认范围（第 1 章）、确认顺序（第 9 章）、确认边界（附录 B）。

### 0.2 与既有能力的关系

| 已有的 | v2 怎么处理 |
|---|---|
| 8 阶段工作流（0 需求探索 → 7 迭代交付，含 1.5 文案） | **不动**。阶段 5 产出"文档/固定版面/改版"时走分支，其余阶段共用 |
| 设计知识库（10 风格 / L01–L15 / AP-01–16 / 字体·色彩·图像·图表） | **不动**，直接复用。文档形态只新增"块模型 + 打印字号档" |
| spec-first 主路径（JSON → 生成器 → 产物 → 渲染 → 复核） | **沿用同一条思路**：`doc.json → docx → pdf → 逐页 PNG → 三层 QA → 目视复核` |
| 42 个 check id 的机器质检 | **不动**。文档形态另起 `qa_doc.py`（三层输入），共享 P0–P3 分级与报告格式 |
| 偏好记忆（v1.4.0） | **不动**。文档任务的偏好同样走 `pref_memory.py`，`kind` 记 `doc` |

### 0.3 三种读法

- **只想知道结论**：读第 1 章 + 第 9 章两张表 + 附录 B。
- **要验收**：读第 10 章验收总表，每条都有命令、输入、预期、证据落点、是否阻断。
- **要复现或接手**：从第 4 章目录树、第 5–8 章模型与契约读起，附录 C 是复制即用的命令序列。

### 0.4 术语

沿用：版式、版式签名、文本块、留白（必须带数值）、阈值唯一真值、spec-first。

新增：

| 术语 | 定义 |
|---|---|
| **产物形态** | 一次任务的最终交付类型：演示稿（slide）/ 文档（doc）/ 固定版面（page）/ 既有文件改版（rebuild） |
| **文档块** | `doc.json` 里最小的内容单位，共 12 种；**P0 只实现 6 种**（见 5.3） |
| **preset** | 一套绑定的阈值组合（画布尺寸 + 字号档 + 行距 + 边距 + 密度），任务开始时选定，中途不换；权威定义只在第 7 章 |
| **静态目录占位** | P0 用普通标题 + 段落写出的目录页，条目是死的文本，不代表自动目录功能（自动目录是 P1 的 `toc` 块） |
| **三层 QA** | 按输入对象分的三档检查：Schema QA（`doc.json`）/ DOCX QA（`.docx`）/ PDF·Render QA（`.pdf` 或 PNG），见第 8 章 |

---

## 1. 目标与边界

### 1.1 要解决的三件事

| 编号 | 需求（你的原话） | v2 的解法 |
|---|---|---|
| N1 | 「PDF 一般是完整的文档，可不可以兼顾 Word 制作」 | 新增**文档分支**：`doc.json` → `build_docx.py` → `.docx`（可编辑）→ `export_pdf.py` → `.pdf`。一份源两种交付 |
| N2 | 「从零做一个或者是修改 pdf」 | 从零：走 N1，或走"固定版面设计"（A4/A3 单页、海报、画册跨页）。修改：`pdf_rebuild.py` 走"逐页底图 + 覆盖层重排"（能力边界见 1.2.1） |
| N3 | 「也不排除要更偏向设计的方向（可能也要求 png 或者 psd）」 | PNG 已有（`render.py` 任意画布）。新增 `export_svg.py`（受支持元素的矢量导出）与 `export_layers.py`（分层 PNG 包 + 坐标清单）。**PSD 二进制不做**，理由见 1.2 |

### 1.2 明确不做

| 不做 | 原因（实测，不是推测） |
|---|---|
| **PDF 文字级编辑**（改原 PDF 里的字） | 本机 Python 无任何 PDF 库（`fitz/pypdf/pdfplumber` 全部 ModuleNotFoundError），环境禁止 `pip install`；office-cli 只能把 PDF 渲成 PNG，不能回写内容流。自写 PDF 解析器只对极简单文件有效，误伤风险高于收益 |
| **PSD 二进制写入** | 无 `psd_tools`/`pytoshop`；8BPS 分层格式手写需处理通道、图层组、混合模式、智能对象，出错时 Photoshop 直接判文件损坏。替代品是 SVG + 分层 PNG，能进 PS 继续编辑，代价是"不是一个 .psd 文件" |
| **网页 / 长图 / H5** | 现有模型是固定画布 + 绝对坐标，没有流式布局与断点概念。硬做会得到"不能响应式的网页"，不如不做 |
| **提升 AI 生图质量** | 属于图像生成能力，不是版式技能；v1.1.1 只解决"生成图必须有结构核对与像素密度门槛" |
| **Word 里做复杂域**（交叉引用自动编号、题注域、Zotero 域） | python-docx 1.2.0 无对应 API，需手写 OOXML 且 LibreOffice 转换行为不稳。降级方案：编号由生成器算好写成纯文本 |
| **复杂 PDF 背景自动修复** | 位图遮挡无法可靠恢复渐变、纹理、照片背景。列为未来方向，不进验收（见 1.2.1） |

### 1.2.1 既有 PDF 改版：承诺什么、不承诺什么

**v2 承诺（P3 验收范围）**：

1. 原页面整体作为底图保留，底图 ≥150 dpi；
2. 对**纯色或简单背景**区域做遮挡（`mask`）并填充指定色；
3. 在指定位置叠加新的文字、形状、图片（覆盖层用 `deck.json` 元素语法）；
4. 对明确指定的局部颜色做**像素级重映射**（`recolor`，容差与统计口径见 1.2.2）；
5. 输出改版前后对照图（`before/after` 并排 PNG）与 `changes.md`。

**v2 明确不承诺**：

- 自动修复复杂纹理 / 渐变 / 照片背景（遮不住就是遮不住，脚本会报告"该区域背景复杂度高，建议人工重排"）；
- 自动恢复被遮挡的原始内容；
- 编辑 PDF 内部的原生文字（导出后文字不可选中、不可搜索）；
- 对任意 PDF 做到"无损改版"；
- 页面重排（把原页面的段落重新流式排版）——那是"从零做"，走 N1。

### 1.2.2 `recolor` 的统计口径（替代原来的"95%"）

| 项 | 定义 |
|---|---|
| 统计区域 | 由 `recolor.region` 指定（`x,y,w,h` 英寸，转像素后取整）；未指定时取整页排除 `recolor.excludePixels` 命中的像素 |
| 原色判定 | 在 sRGB 下计算 RGB 欧氏距离：`d = sqrt((r-r0)² + (g-g0)² + (b-b0)²)`，`d ≤ tolerance` 记为"目标像素"（`tolerance` 默认 30，范围 0–60） |
| 抗锯齿边缘 | 与背景主色差落在 `1.0–1.35` 倍容差区间的像素视为边缘，**只替换 50% 混合色**，且**不计入**统计分母 |
| 分母 | 符合条件的"目标像素"总数（不含边缘、不含 `excludePixels`） |
| 替换率 | `已替换目标像素 / 目标像素总数`；**只作为 P3 的辅助异常检测**（<90% 提示人工复核），**不作为验收判据** |
| 验收看什么 | 对照图人工复核 + `changes.md` 逐条列"改了哪页/哪个区域/怎么改/为什么" |

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
   （现状，不动）    （新增，P0+P1 阶段）      （新增，P2 阶段）      （新增，P3 阶段，受限）
```

### 1.4 成功判据（六条，都能被第 10 章的命令证明）

1. 一份文档两种交付：任一份 `doc.json` 产出 `.docx` 与 `.pdf`，三层 QA 全跑且各自输入对象正确（第 8 章）。
2. 换 tokens 不塌：只改 ≤6 个字体/颜色值，重建后版面不塌、三层 QA 仍全绿。
3. 打印形态：A4 海报经 `preset=print-A4` 出 PDF 与 PNG，`PRT-01/02` 无 P0/P1。
4. 矢量导出：受支持元素清单内的页面导出 SVG，文字仍是文字；**不支持的 `chart`/`table` 渲染为位图并写进跳过报告**（不静默丢弃）。
5. 改版受限但可交付：8 页 PDF → 8 页 PDF，底图 ≥150 dpi，遮挡区域无残留（对照图人工确认），`changes.md` 条目与改动一致。
6. 五个案例（技术报告 / 研究摘要 / 一页纸 / A4 海报 / PDF 改版）各自带 `brief.md`、`source/`、`build/`、`README.md`、三层 QA 记录。

---

## 2. 现状盘点（v1.4.0 实测）

### 2.1 仓库数字

| 目录 | 文件数 | 体积 | 说明 |
|---|---|---|---|
| 根目录 | 7 | 41.6 KB | SKILL / README / CONTRIBUTING / install / CHANGELOG / LICENSE / .gitignore |
| `discovery/` | 5 | 67.8 KB | 探索规范 + 问题库（Q1–Q38）+ 对照表 |
| `references/` | 14 | 195.1 KB | 10 种风格、10 条原则、L01–L15、AP-01–16、6 份专项、延伸阅读、文案规格、去 AI 味、偏好记忆 |
| `workflows/` | 9 | 87.6 KB | 00 → 07 + 01b 文案 |
| `templates/` | 9 | 26.1 KB | 需求确认书 / 简报 / 文案稿 / 换模型交接词 / 提案 / 大纲 / 单页规格 / QA 报告 / 偏好条目 |
| `examples/` | 71 | 14,363.5 KB | 3 份分析示例 + 3 个完整案例（含 pptx 与预览图） |
| `scripts/` | 17 | 502.7 KB | probe / build_pptx / render / qa_check / contact_sheet / copy_check / pref_memory / selfcheck + 3 库 + README |
| `schema/` | 2 | 19.7 KB | `deck.schema.json`、`thresholds.json`（v1.3.0） |
| `docs/` | 6 | 125.8 KB | 环境 / 验证 / 发现测试 / 阶段总结 / 本规划书 / 操作书 |
| 合计 | **177** | **17.13 MB** | zip 15.36 MB / sha256 `6208ab9d…` |

### 2.2 已验证能力（都跑过，数字是产物体积）

| 通路 | 命令 | 实测产物 |
|---|---|---|
| PPTX 生成 | `build_pptx.py --deck demo/source/deck.json --out demo/build/deck.pptx` | 8 页 / 156,716 B / 字体 2 族 |
| PPTX → PNG | `render.py` 或 `office-cli render --dpi 144` | 1920×1081×8 |
| PPTX → PDF | `office-cli convert --input deck.pptx --output deck.pdf` | 8 页 / 404,760 B / `%PDF-1.7` |
| A4 竖版版面 | `deck.json` 的 `meta.size = "8.268x11.693"` | 1 页 / 28,634 B → 渲染 1241×1754 px @150 dpi |
| DOCX 生成 → PDF | python-docx 生成 → `convert` | 37,568 B → 30,485 B / 1 页（告警 `missingFonts:["Courier"]`） |
| PDF → PNG | `office-cli render --input x.pdf --pages 1,2 --dpi 150` | pdfium 后端 / 128,460 B + 81,885 B |
| 版面质检 | `qa_check.py` | 30 个 check id，demo/三案例均 `0/0/0/0` |
| 文案质检 | `copy_check.py` | 12 个 check id，正例 P1=0，负向 fixture 命中 11 项 |

### 2.3 本机依赖与硬限制

- Python 3.12.14 可用包：`python-docx 1.2.0`、`python-pptx 1.0.2`、`openpyxl 3.1.5`、`lxml 6.1.3`、`Pillow 12.3.0`、`numpy 2.3.5`、`pandas 3.0.1`。
- **缺**：任何 PDF 库、`reportlab`、`psd_tools`、`cairosvg/svglib`、`matplotlib`、`fontTools`、`yaml`、`jinja2`。
- 渲染/转换后端：DSH 自带 `office-cli`（LibreOffice 内核 + pdfium），能力 = 转换（docx/xlsx/pptx ↔ pdf）+ 逐页栅格化。**不带 PDF 内容编辑**。
- 无 MS Office / WPS / COM；无 Inkscape / Illustrator；不许 `pip install`。
- 字体 240 个（微软雅黑、思源宋体 Heavy、Noto Sans SC、Lato、Cascadia Mono 等），转换时缺字体告警可用于校验。

### 2.4 缺口表

| 需要 | 现在有 | 差什么 |
|---|---|---|
| 文档级排版（标题层级/页眉页脚/目录/表格跨页） | 无 | `doc.schema.json` + `build_docx.py` |
| 文档级质检（能看 docx 与 pdf 的真实属性） | 无（`qa_check.py` 只认 slide 画布） | `qa_doc.py`（三层输入） |
| 打印尺寸的字号档位 | 只有 `slide-16x9`（正文 ≥16 pt） | `presets`：`print-A4` / `print-A3` / `print-letter` |
| PDF 交付 | 能转换，但脚本里没封装 | `export_pdf.py`（校验页数/缺字体/可栅格化） |
| 矢量导出 | 无 | `export_svg.py`（受支持元素） |
| 分层交付 | 无 | `export_layers.py` |
| 既有 PDF 改版 | 能渲成图 | `pdf_rebuild.py`（底图 + 覆盖层 + 颜色重映射） |

### 2.5 PPT 回归基线（开工前先冻结）

冻结对象（2026-11-15，v1.4.0）：

| 冻结项 | 内容 | 比对方式 |
|---|---|---|
| 代码 | `scripts/*.py`、`scripts/lib/*.py` | 文件 SHA-256 清单 `docs/ppt-baseline.json` |
| Schema | `schema/deck.schema.json`、`schema/thresholds.json` | 同上 |
| 产物 | `demo/build/deck.pptx`、三案例 `build/deck.pptx` | **zip 成员 CRC 清单**（PPTX 内含时间戳，文件级 sha 会变，不能只靠它） |
| 质检结果 | demo / smoke / 三案例的 `qa.json` 与 `copy.json` 摘要 | 逐项对比 `SUMMARY` 行 |
| 视觉 | `demo/build/contact-sheet.png`、三案例 contact sheet | 逐页 PNG 像素差 ≤0.5% 视为一致（LibreOffice 版本不变时） |

基线不达标即停：新增功能导致的任何 baseline 差异，必须先解释清楚再继续。

---

## 3. 四种产出形态的判定

### 3.1 判定顺序（每次都先走这张表）

| 用户手上有什么 / 要什么 | 走哪种形态 | 首条命令 |
|---|---|---|
| 要讲给人听，一页一个观点，投屏 | A 演示稿 | `build_pptx.py` |
| 要发出去读，有分段、章节、目录、页码 | B 文档 | `build_docx.py` |
| 单页/多页的版面作品（海报、一页纸、画册跨页、信息图） | C 固定版面 | `build_pptx.py` + `preset=print-*` |
| 已经有一份 PDF，要改设计/改版式 | D 改版 | `pdf_rebuild.py` |
| 只改 PDF 里几行字 | **不接**：要源文件，或按 D 重排 | — |
| 要 .psd | 改交付 **SVG + 分层 PNG**，并说明差异 | `export_svg.py` / `export_layers.py` |

### 3.2 四种形态的产物矩阵

| 形态 | 主产物 | 附带产物 | 质检 |
|---|---|---|---|
| A 演示稿 | `.pptx` | `.pdf`、逐页 `.png`、`contact-sheet.png`、`qa.md/json` | `qa_check.py` |
| B 文档 | `.docx` + `.pdf` | 逐页 `.png`、`qa.md/json`（三层） | `qa_doc.py` |
| C 固定版面 | `.pdf` + `.png` | `.pptx`（可编辑源）、`.svg`、分层 png 包 | `qa_check.py`（+PRT 项） |
| D 改版 | `.pdf` | 对照图、`changes.md` | `qa_check.py` + 人工清单 |

---

## 4. 目录与文件变更

### 4.1 新增（14 个文件 / 2 个目录）

| 路径 | 职责（一句话） | 首实现 |
|---|---|---|
| `schema/doc.schema.json` | 文档模型：`meta` + `tokens` + `sections[].blocks[]`，12 种块 | P0（6 种）/ P1（其余） |
| `schema/presets.json` | 四档 preset 的阈值绑定（slide / A4 / A3 / letter） | P0 |
| `scripts/build_docx.py` | `doc.json` → `.docx`（样式注入、页眉页脚、表格跨页、图片缩放） | P0（TOC 域 P1） |
| `scripts/export_pdf.py` | `.docx`/`.pptx` → `.pdf`，校验页数、缺字体、可栅格化 | P0 |
| `scripts/qa_doc.py` | 三层文档质检（`--doc` / `--docx` / `--pdf` / `--preview-dir`） | P0（L1+L3）/ P1（L2） |
| `scripts/export_svg.py` | 一页 `deck.json` → 单文件 `.svg`（受支持元素；`chart`/`table` 位图化 + 跳过报告） | P2 |
| `scripts/export_layers.py` | 一页/一套 → 分层 PNG 包（`layers/*.png` + `layers.json`） | P2 |
| `scripts/pdf_rebuild.py` | 既有 PDF → 逐页底图 → 覆盖层 → 新 PDF（含颜色重映射与 `changes.md`） | P3 |
| `references/document-typography.md` | 文档排版规范（字号/行距/行长/页边/页眉页脚基线/图表编号） | P0 |
| `references/document-structure.md` | 文档结构规范（章节深度、目录、表图编号、参考文献、附录） | P1 |
| `templates/doc-brief.md` | 文档任务简报模板 | P0 |
| `templates/doc-outline.md` | 文档大纲模板（章节 → 每节一句话任务 → 预估页数） | P1 |
| `workflows/05b_document-production.md` | 阶段 5 的文档分支（块级制作顺序、每页放什么、表图怎么落） | P0 |
| `examples/docs/<case>/` | 3 个文档案例目录（结构同现有 slide 案例） | P4 |

### 4.2 修改（7 个文件，都只加段不改旧行为）

| 文件 | 改哪一段 | 加什么 |
|---|---|---|
| `SKILL.md` | 第 2 节「任务形态」 | 加产物形态判定表（3.1），指向 `05b` |
| `SKILL.md` | 第 5 节「硬指标」 | 加一行"文档形态的硬指标见 `references/document-typography.md`" |
| `SKILL.md` | 路由表 | 加 4 行：要 docx / 要 pdf / 要 svg / 要改 PDF |
| `schema/thresholds.json` | 顶层加 `presets` 指针 + `print` 段 | 打印字号与边距阈值（第 7 章），顶层结构保持兼容 |
| `scripts/README.md` | 命令表 + check id 表 | 加 6 条命令、23 个 check id（第 8 章）、三层 QA 说明 |
| `README.md` | 第 1 节 + 结构树 | 一句话说明"四种形态"，结构树补新文件 |
| `CHANGELOG.md` | 版本段 | `[2.0.0]` / `[2.2.0]` / `[2.3.0]` / `[2.4.0]` 四段 |

### 4.3 不动（红线）

- `deck.schema.json` 的字段与 `build_pptx.py` 的**默认输出行为**：一字节都不改。唯一例外：允许新增**可选**参数 `--preset`，不带该参数时输出必须与 v1.4.0 逐字节一致（比对方法：解 zip 逐成员 CRC，期望"差异 0"）。
- 现有 42 个 check id 的编号、判据与分级：新增项只能加在 `PRT-*` / `DOC-*` / `PDF-*` / `EXP-*` 命名空间。
- demo 与三个 slide 案例的产物：每个阶段收尾都回归一次（第 2.5 节基线），QA 必须仍是 `0/0/0/0`。

### 4.4 全部阶段完成后的目录树（只标新增）

```text
ai-ppt-design-system/
├── SKILL.md                        # 改：加产物形态判定表 + 4 条路由
├── README.md  install.md  CHANGELOG.md  LICENSE  CONTRIBUTING.md
├── discovery/                      # 不动（5）
├── references/                     # 14 + 2（document-typography / document-structure）
├── workflows/                      # 9 + 1（05b_document-production）
├── templates/                      # 9 + 2（doc-brief / doc-outline）
├── schema/
│   ├── deck.schema.json            # 不动
│   ├── doc.schema.json             # 新增
│   ├── thresholds.json             # 改：加 print 段
│   └── presets.json                # 新增
├── scripts/                        # 17 + 6 → 23
├── examples/
│   ├── cases/                      # 不动（3 个 slide 案例）
│   └── docs/                       # 新增 3 个文档案例 + 1 个海报 + 1 个改版案例
├── docs/
│   ├── environment-report.md  validation-report.md  discovery-tests.md  project-summary.md   # 不动
│   ├── v2-plan.md                  # 本文件
│   ├── v2-plan-ops.md              # 操作书
│   ├── v2-revision-notes.md        # 修订说明（R1 → R2 逐条处理）
│   ├── ppt-baseline.json           # PPT 回归基线
│   └── v2-validation-report.md     # 新增（V9–V20 验证矩阵）
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
    "toc": { "enabled": false, "depth": 3 },
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
| `marginsIn` | object | 是 | 每个边 0.5–1.6 in；与 preset 不一致时告警（`DOC-11`） |
| `toc.enabled` | bool | 否 | **P1 起生效**：true 时生成器插 TOC 域，交付说明写"打开后按 F9 更新目录"。P0 阶段该字段必须为 false（置 true 直接退出码 1） |
| `numbering.figures/tables` | bool | 否 | true 时图注写成"图 1 · 冠层剖面示意"，编号由生成器按出现顺序算 |
| `header` / `footer` | object | 否 | 页码用 PAGE 域；首页是否显示单独控制 |

### 5.3 12 种文档块（含首实现阶段）

| `type` | 首实现 | 关键字段 | 渲染为 |
|---|---|---|---|
| `heading` | **P0** | `level`(1–3)、`text` | Word 标题 1–3 样式 |
| `paragraph` | **P0** | `runs[]`、`align`、`indentFirstLine`、`spaceAfterPt` | 正文段落，中文首行缩进 2 字符（可关） |
| `list` | **P0** | `ordered`、`items[][]`、`level` | 项目符号 / 编号列表，编号由生成器维护 |
| `table` | **P0** | `caption`、`header[]`、`rows[][]`、`colWidthsIn[]`、`align[]`、`zebra` | Word 表格，表头跨页重复、行不跨页断开 |
| `figure` | **P0** | `src`、`widthIn`、`caption`、`source`、`alt` | 图片 + 图注 + 来源行 |
| `pageBreak` | **P0** | — | 硬分页 |
| `quote` | P1 | `text`、`cite` | 引文块（缩进 + 左侧竖线） |
| `callout` | P1 | `kind`(`note`/`warn`/`key`)、`title`、`text` | 单格表格做的提示框 |
| `code` | P1 | `lang`、`text` | 等宽字体块（灰底） |
| `toc` | P1 | — | 目录域（插 `TOC` 域，生成后需 F9 更新） |
| `bibliography` | P1 | `entries[]` | 参考文献列表（编号 `[1]`，与正文引用号人工对齐） |
| `footnoteGroup` | P2 | `id` | 脚注组（先降级为章末"注释"） |

**P0 的硬约束**：P0 的 `doc.schema.json` 只接受前 6 种；出现其余类型 → 退出码 1，报 `DOC-16 块类型未在该阶段实现`。

### 5.4 与 `deck.schema.json` 的本质差异

| 维度 | deck（演示稿） | doc（文档） |
|---|---|---|
| 定位方式 | 绝对坐标（`x/y/w/h` 英寸） | 流式（块顺序 + 段落/表格/图片） |
| 分页 | 由人决定（每页元素固定） | 由渲染器决定（会流动、会断页） |
| 质检重心 | 溢出、越界、重叠、密度 | 层级、跨页、目录、页码、字号档位 |
| 失败模式 | 框太小装不下 | 断页断在表格中间、图注与图分离、目录页码不符 |
| 谁负责美观 | 我们（坐标就是设计） | 样式表 + 上图下图注的间距（生成器只保证规则） |

### 5.5 三条模型级约束

**规则：块顺序即阅读顺序，不许用坐标把内容挪到别处**
- **触发条件**：需要在文档里做"左右分栏""图上文下"之类的特殊摆放。
- **具体动作**：用 `table`（两列各放一块内容）或 `figure` 的 `widthIn` 控制，不要引入绝对坐标。
- **常见错误**：在 doc 模型里塞 `x/y`，结果是 DOCX 里出现浮动文本框，Word 打开后错位、也无法回流。

**规则：图、表必须有编号、题注与来源**
- **触发条件**：插入任何 `figure` 或 `table`。
- **具体动作**：`caption` 写"图 N + 一句话结论式题注"，`source` 写数据出处与口径；缺任一项判 `DOC-05`/`DOC-06`（P0）。
- **常见错误**：把来源塞在正文里、或写"如图所示"却不给编号，读者无法回指。

**规则：标题层级只能加深一级**
- **触发条件**：写第 N 级标题。
- **具体动作**：h1 → h2 → h3 依次加深，跳过即 `DOC-01`（P0）。
- **常见错误**：为了字号好看直接 h1 → h3，导致目录塌成两层、Word 导航窗格断链。

---

## 6. 脚本契约（七条命令）

### 6.1 总表

| 命令 | 输入 | 输出 | 退出码 |
|---|---|---|---|
| `build_docx.py --doc D --out O [--check-only]` | `doc.json` | `.docx` | 0 成功 / 1 有 P0 级模型错误 / 2 参数或 I/O 错 |
| `export_pdf.py --input F --out O [--pages all\|"1-3"]` | `.docx`/`.pptx` | `.pdf` + `pdf-manifest.json` | 0 / 1 转换成功但缺字体 / 2 失败 |
| `qa_doc.py --doc D [--docx F] [--pdf F] [--preview-dir DIR] --out O [--json J]` | 三类输入任意组合 | `qa.md` + `qa.json` | 0 无 P0 / 1 有 P0 / 2 参数错误 |
| `export_svg.py --deck D --slide N --out O [--allow-raster]` | `deck.json` | `.svg` + 跳过报告 | 0 / 1 有元素不支持（已位图化） / 2 失败 |
| `export_layers.py --deck D --slide N --out-dir O` | `deck.json` | `layers/*.png` + `layers.json` | 0 / 1 部分层为空 / 2 失败 |
| `pdf_rebuild.py --pdf P --spec S --out-dir O` | 原 PDF + `rebuild.json` | 新 PDF + 对照图 + `changes.md` | 0 / 1 有页面未处理 / 2 失败 |
| `render.py`（现状） | `.pptx`/`.pdf` | 逐页 `.png` | 0 / 2 无渲染器 |

### 6.2 `build_docx.py` 关键实现约定

1. **样式注入而非手工排版**：只改 `Normal` / `Heading 1–3` / 自定义 `FigCaption`、`TblCaption`、`Source`、`Quote`、`Callout`、`Code` 七个样式；段落只引用样式，不逐段设字体。用户在 Word 里改样式能整套生效。
2. **中文字体必须双写**：`run.font.name`（西文）+ `rPr.rFonts` 的 `w:eastAsia`（中文），否则 Word 里中文回落到宋体（实测过，这是最常见的"字体没生效"）。
3. **行距与段后**：正文行距 1.5 倍，段后 6 pt，首行缩进 2 字符（`first_line_indent = Pt(21)` 对 10.5 pt 正文）。
4. **目录**：P1 起插 `w:fldSimple` + `instr='TOC \o "1-3" \h \z \u'`，页面上留一行"目录（打开后按 F9 更新）"；**不假装目录页码已经正确**。P0 不插域，目录页是静态文本。
5. **页眉页脚**：页脚居中 `PAGE` 域 + 可选 `NUMPAGES`；页眉用 `different_first_page_header_footer` 控制首页不显示。
6. **表头跨页**：给首行 `<w:tblHeader/>`，所有行 `<w:cantSplit/>`，防止表格在页底断成两半。
7. **图片**：`widthIn` 缺省 = 正文栏宽（A4 竖版 6.27 in），按比例算高；超过一页高度时按页高 80% 反算宽度，stderr 提示"图片过大已缩放"。
8. **编号**：图/表编号由生成器按出现顺序算成纯文本（`图 1`、`表 2`），不用 SEQ 域。

### 6.3 `export_pdf.py` 校验项

| 校验 | 判据 | 失败时 |
|---|---|---|
| 输出存在且 >1 KB | 文件大小 | 退出码 2，打印后端原始报错 |
| 页数 >0 | manifest 的 `pageCount` | 退出码 2 |
| 缺字体列表 | `missingFonts == []` | 退出码 1，列出字体 + 建议替换 |
| 可再次栅格化 | 用自己的 `render` 渲第 1 页成功 | 退出码 1（PDF 可能损坏但能打开） |

### 6.4 `qa_doc.py` 的三层结构与 16 项检查

**三层输入（缺哪层，该层的检查就记 `SKIP` 并计入"未验证"，不许算通过）**：

```bash
# L1 只查模型
$PY scripts/qa_doc.py --doc source/doc.json --out build/qa-schema.md
# L2 查生成的 docx（样式、域、表格属性、图片嵌入）
$PY scripts/qa_doc.py --doc source/doc.json --docx build/report.docx --out build/qa-docx.md
# L3 查 pdf / 渲染图（页数、尺寸、字体、越界、视觉问题）
$PY scripts/qa_doc.py --doc source/doc.json --docx build/report.docx \
    --pdf build/report.pdf --preview-dir build/preview --out build/qa-full.md --json build/qa.json
```

报告头必须打印 `inputs: doc=… docx=… pdf=… preview=…`，并统计 `verified=… skipped=…`。

| ID | 层 | 检查 | 所需输入 | 阶段 | 级别 |
|---|---|---|---|---|---|
| `DOC-01` | L1 | 标题层级跳级 | doc | P0 | P0 |
| `DOC-02` | L1 | 无 h1 或标题数 <2 | doc | P0 | P1 |
| `DOC-03` | L1 | 标题字数超限（title >30 / heading >24） | doc | P0 | P2 |
| `DOC-04` | L1 | 段落过长（>300 字） | doc | P0 | P2 |
| `DOC-05` | L1 | 图/表缺题注 | doc | P0 | P0 |
| `DOC-06` | L1 | 图/表缺来源 | doc | P0 | P0 |
| `DOC-07` | L1/L3 | 图片有效像素密度（<80 P1、<110 P2） | doc + 图片文件 | P0 | P1/P2 |
| `DOC-08` | L1 | 字体族 >3 | doc | P0 | P1 |
| `DOC-09` | L1 | 正文字号不在 preset 档（A4 9–11 pt） | doc | P0 | P1 |
| `DOC-10` | L1 | 行距不在档（1.4–1.6） | doc | P0 | P2 |
| `DOC-11` | L1 | 页边距与 preset 差 >0.1 in | doc | P0 | P2 |
| `DOC-12` | L2 | 目录与标题数不符（`toc.enabled` 时） | docx | P1 | P1 |
| `DOC-13` | L2 | 应有 PAGE 域但找不到 | docx | P0 | P1 |
| `DOC-14` | L1 | 表格列宽和 ≠ 正文宽（差 >0.15 in） | doc | P0 | P2 |
| `DOC-15` | L1 | 中英文之间缺空格 | doc | P0 | P3 |
| `DOC-16` | L1 | 块类型未在该阶段实现（P0 只允许 6 种） | doc | P0 | P0 |

`PRT-01` 打印边距不足（A4 <0.6 in、A3 <0.8 in，P1）；`PRT-02` 出血元素缺 `zRole:"bgDecor"`（P2）。
`PDF-01` 页数与源不符（P1）；`PDF-02` 缺字体（P1）；`PDF-03` 页面尺寸不符 preset（P1）。
`EXP-01` SVG 元素越界（P1）；`EXP-02` 分层包缺 `layers.json` 或层为空（P1）；`EXP-03` SVG 中文字被转成路径（P2，允许但需在 README 说明）。

### 6.5 `export_svg.py`：受支持元素与跳过报告

**支持**（就地转 SVG 元素）：`rect`（含 `corners`）、`line`（`dash`→`stroke-dasharray`）、`ellipse`、`polygon`、`text`（runs/paragraphs → `<text>`+`<tspan>`，逐项落 `font-family`/`font-size`/`letter-spacing`/`line-height`；不用 `textLength`，会拉伸字形）、`image`（内联 base64）。

**不支持**：`chart`、`table`、`icon`（图标字体）。

不支持的处置（必须显式，不许静默丢）：

1. 默认 `--allow-raster` 关闭时：跳过并写进 `skip-report.md`，退出码 1；
2. 开启时：该元素先渲染为 PNG（同一画布网格、同坐标），以 `<image>` 嵌入，退出码 1，报告里写"已位图化"，并在 SVG 的 `<desc>` 标注；
3. 报告必须列：元素序号、类型、位置、处置（跳过/位图化）、原因。

**已写明的限制**：SVG 的 `<text>` 保留为文字对象，但**不同软件/不同机器字体与换行会变**；不能因为浏览器能 Ctrl+F 搜到标题，就声称"在设计软件里编辑体验良好"。本机无 Inkscape / Illustrator，**"在设计软件中可正常选择编辑"列为待验证项（V18）**，验证前不得写进 README 的能力声明。

### 6.6 `export_layers.py`：分层 PNG 的真实边界

- 五层：`01-bg` / `02-shapes` / `03-images` / `04-decor` / `05-text`，整画布同尺寸透明 PNG，逐层对齐。
- `layers.json` 每层记录：`name`、`order`、`bounds{x,y,w,h}`、`opacity`、`sourceElementIds[]`、`rasterized`(bool)、`note`。
- **文字层是位图**，不是可编辑文本：`rasterized: true`，`note` 固定写"文字层为栅格图，改字请用 SVG 或回源 `deck.json`"。
- README 与交付说明都必须写明："分层 PNG 不是 PSD 的等价替代品；要改文字用 SVG 或源文件。"

---

## 7. 四档 preset（唯一权威定义）

> 其它章节只引用本章，不重复数值。`schema/presets.json` 与 `schema/thresholds.json` 的 `print` 段必须与本章逐项一致；`scripts/selfcheck.py` 之外再加一项一致性检查（P0 阶段实现，见操作书 S3）。

| preset | 画布（in） | 边距（in，四边） | 正文字号 | 行距 | 行长上限 | 页面字数预算 | 用途 |
|---|---|---|---|---|---|---|---|
| `slide-16x9` | 13.333 × 7.5 | 0.6 / 0.45 | 16–24 pt | 1.25–1.55 | 34 汉字 | 220 字/页 | 现有演示稿（不动） |
| `print-A4` | 8.268 × 11.693 | 1.0 / 0.9 | 9–11 pt | 1.4–1.6 | 48 汉字 | 600–900 字/页 | 报告、研究摘要、一页纸 |
| `print-A3` | 11.693 × 16.535 | 1.2 / 1.0 | 10–12 pt | 1.4–1.6 | 56 汉字 | 700–1000 字/页 | 展板、大图谱 |
| `print-letter` | 8.5 × 11 | 1.0 / 0.9 | 9–11 pt | 1.4–1.6 | 48 汉字 | 600–900 字/页 | 英文交付 |

**生效顺序（唯一规则）**：`命令行 --preset` > `meta.pageSize` 推导 > 默认 `slide-16x9`。报告头必须打印 `preset: <name> (source: cli|meta|default)`。

**TOC 与 preset 无关**：`toc.enabled` 是块能力，`print-*` 不自动开目录。

---

## 8. QA 检查项映射表（规划书与操作书共用）

> 编号、所属层、所需输入、阶段、级别、fixture：**两份文档与 `scripts/README.md` 必须逐行一致**。

| ID | 层 | 输入对象 | 阶段 | 级别 | 对应 fixture |
|---|---|---|---|---|---|
| `DOC-01` | L1 | doc.json | P0 | P0 | `qa-doc-bad.json`（h1→h3） |
| `DOC-02` | L1 | doc.json | P0 | P1 | `qa-doc-bad.json`（仅 1 个 h1） |
| `DOC-03` | L1 | doc.json | P0 | P2 | `qa-doc-bad.json`（title 34 字） |
| `DOC-04` | L1 | doc.json | P0 | P2 | `qa-doc-edge.json`（paragraph 340 字） |
| `DOC-05` | L1 | doc.json | P0 | P0 | `qa-doc-bad.json`（figure 无 caption） |
| `DOC-06` | L1 | doc.json | P0 | P0 | `qa-doc-bad.json`（table 无 source） |
| `DOC-07` | L1+L3 | doc.json + 图片 | P0 | P1/P2 | `qa-doc-edge.json`（300×170 铺 6 in） |
| `DOC-08` | L1 | doc.json | P0 | P1 | `qa-doc-edge.json`（4 字族） |
| `DOC-09` | L1 | doc.json | P0 | P1 | `qa-doc-bad.json`（body 8.5 pt） |
| `DOC-10` | L1 | doc.json | P0 | P2 | `qa-doc-bad.json`（行距 1.8） |
| `DOC-11` | L1 | doc.json | P0 | P2 | `qa-doc-bad.json`（左 1.5 in 与 preset 差 0.5） |
| `DOC-12` | L2 | docx | P1 | P1 | `qa-doc-toc-bad.docx` |
| `DOC-13` | L2 | docx | P0 | P1 | `qa-doc-nopage.docx` |
| `DOC-14` | L1 | doc.json | P0 | P2 | `qa-doc-edge.json`（列宽和差 0.4 in） |
| `DOC-15` | L1 | doc.json | P0 | P3 | `qa-doc-edge.json`（"用Word打开"） |
| `DOC-16` | L1 | doc.json | P0 | P0 | `qa-doc-p1block.json`（P0 用 quote） |
| `PRT-01` | L1 | deck.json | P2 | P1 | `qa-prt-bad.json`（A4 边距 0.4 in） |
| `PRT-02` | L1 | deck.json | P2 | P2 | `qa-prt-bad.json`（满版图无 bgDecor） |
| `PDF-01` | L3 | pdf + manifest | P0 | P1 | `qa-pdf-pages.pdf`（8 页 vs 4） |
| `PDF-02` | L3 | pdf + manifest | P0 | P1 | `qa-pdf-fonts.pdf`（缺 Courier） |
| `PDF-03` | L3 | pdf | P0 | P1 | `qa-pdf-a3.pdf`（A3 尺寸 vs A4 preset） |
| `EXP-01` | L1 | deck.json | P2 | P1 | `qa-exp-bad.json`（元素越界） |
| `EXP-02` | L2 | layers 目录 | P2 | P1 | `layers-missing/`（缺 layers.json） |
| `EXP-03` | L2 | svg 文件 | P2 | P2 | `svg-outlined.svg`（文字转路径） |

合计：新增 **23 项**（DOC 16 + PRT 2 + PDF 3 + EXP 3，其中 `DOC-16` 是本轮为"阶段范围"新增），加上现有 42 项 → **65 项**。

---

## 9. 五阶段与版本

### 9.1 阶段总览

| 阶段 | 目标 | 交付物 | 发布版本 |
|---|---|---|---|
| **P0** | 文档最小闭环：4 页样张，只用 6 种块 | `doc.schema.json`、`qa_doc.py` L1+L3、`build_docx.py`（无 TOC）、`export_pdf.py`、`examples/docs/report-min/` | **1.5.0** |
| **P1** | 文档完整能力：10–15 页真实文档 + 30 页压力测试 | 其余 6 种块、TOC 域、L2 QA、`examples/docs/research-brief/`（10–15 页）、`examples/docs/technical-report/`（30 页压力） | **1.6.0** |
| **P2** | 固定版面与设计型导出 | `presets` 生效、`export_svg.py`、`export_layers.py`、`PRT/EXP` 项、A4 海报案例 | **1.7.0** |
| **P3** | 既有 PDF 受限改版 | `pdf_rebuild.py`、`rebuild.json`、对照图与 `changes.md`、8 页改版案例 | **1.8.0** |
| **P4** | 综合验收与发布（**不是功能阶段**） | 五案例最终版、V9–V20 验证报告、发布包与解包复验 | **1.9.0**（只加案例与文档） |

### 9.2 每阶段只依赖已实现的能力

| 阶段 | 只许调用 | 不许调用（尚未实现） |
|---|---|---|
| P0 | `build_docx`（6 块）、`export_pdf`、`qa_doc`（L1/L3）、`render` | TOC 域、L2 QA、SVG、分层、改版 |
| P1 | 上述 + L2 QA + 其余 6 块 + TOC | SVG、分层、改版 |
| P2 | 上述 + `--preset` + SVG + 分层 | 改版 |
| P3 | 上述 + 改版 | — |
| P4 | 全部（只做案例、验证、打包） | 任何功能改动 |

### 9.3 P0 完成标准（两份文档共用这一份，编号 A0-*）

| 编号 | 判据 | 证据 |
|---|---|---|
| A0-1 | `build_docx.py --check-only` 退出码 0，无 schema 错误 | 命令输出 |
| A0-2 | `.docx` 生成成功且 `office-cli convert` 能打开（不报错） | 命令输出 + 文件 |
| A0-3 | `.pdf` 页数 == 4（P0 样张固定 4 页） | `pdf-manifest.json` 的 `pageCount` |
| A0-4 | 逐页 PNG 4 张，150 dpi，A4 尺寸（1241×1754 ±2 px） | `build/preview/` + 尺寸打印 |
| A0-5 | 三层 QA 都跑过：L1 `P0=0`、L3 `P0=0`、L2 记 `SKIP` 并计入未验证（P0 不实现 L2） | `qa.json` 的 `verified/skipped` |
| A0-6 | 4 页逐页目视复核通过（6 项清单：层级、断页、题注、来源、页码、留白） | 目视记录写进 `build/qa.md` |
| A0-7 | PPT 回归：demo + 三案例 QA 仍 `0/0/0/0`；pptx zip 成员 CRC 与 `docs/ppt-baseline.json` 一致 | 回归脚本输出 |

### 9.4 P1 完成标准（A1-*）

| 编号 | 判据 |
|---|---|
| A1-1 | 12 种块全部可用，`DOC-16` 不再触发 |
| A1-2 | `research-brief`（10–15 页）三层 QA 全绿（`verified == total`，`skipped == 0`） |
| A1-3 | TOC 域存在（L2 `DOC-12` 通过），目录条目覆盖全部 h1–h3 |
| A1-4 | 30 页 `technical-report` 压力测试：无空白页、无表格跨页断裂、无图注分离 |
| A1-5 | 换 tokens（≤6 值）重建后 QA 仍全绿 |

### 9.5 P2 / P3 / P4 完成标准

| 阶段 | 判据 |
|---|---|
| P2 | `preset=print-A4` 生效且报告头打印来源；A4 海报 `PRT-01/02` 无 P0/P1；SVG 跳过报告与实际元素一致；`layers.json` 字段齐全且 `05-text.rasterized == true` |
| P3 | 8 页 PDF → 8 页 PDF；底图 ≥150 dpi；对照图逐页可比；`changes.md` 条目数与实际改动一致；`recolor` 替换率仅记录不进验收 |
| P4 | 五案例齐全；V9–V20 全跑；解包复验（selfcheck 0 / demo QA 0 / 新增命令 `--help` 全通）；`docs/v2-validation-report.md` 落盘 |

### 9.6 一阶段不过不进下一阶段

每个阶段收尾必须同时满足：① 该阶段全部 A* 判据通过；② PPT 回归通过；③ `selfcheck.py --strict` 全 0；④ 操作书对应章节的"失败处理"没有被触发过的遗留项。任何一条不满足 → 停在该阶段，不往下开。

---

## 10. 验收总表（三类证据，不只靠文件大小）

> 每条都回答五个问题：用什么命令查 / 查哪个输入 / 预期是什么 / 证据存哪 / 失败是否阻断。

### 10.1 P0

| 编号 | 类型 | 命令 | 输入 | 预期 | 证据 | 阻断 |
|---|---|---|---|---|---|---|
| A0-1 | 功能 | `$PY scripts/build_docx.py --doc source/doc.json --out build/report-min.docx --check-only` | doc.json | rc=0 | 终端输出存 `build/logs/` | 是 |
| A0-2 | 功能 | `$PY scripts/build_docx.py --doc … --out build/report-min.docx` | doc.json | rc=0，文件存在 | 文件 + 大小 | 是 |
| A0-3 | 功能 | `$NODE $CLI convert --input build/report-min.docx --output build/report-min.pdf` | docx | 200 状态、pdf >10 KB | `pdf-manifest.json` | 是 |
| A0-4 | 结构 | `$PY scripts/qa_doc.py --doc … --pdf … --preview-dir build/preview --out build/qa.md --json build/qa.json` | doc+pdf+png | `PDF-01..03` 通过 | `qa.json` | 是 |
| A0-5 | 结构 | 同上 | doc+docx | `DOC-13` 通过（PAGE 域存在） | `qa.json` | 是 |
| A0-6 | 视觉 | 人工看 4 张 PNG | preview | 6 项清单全过 | 写进 `build/qa.md` 末节 | 是 |
| A0-7 | 回归 | `$PY scripts/regress_ppt.py`（P0 新增，只读） | 基线 + 现仓 | CRC 差异 0、QA 全 0 | `build/regress.json` | 是 |

### 10.2 P1–P4（摘要，细则见操作书）

| 阶段 | 关键验收 | 辅助指标（不作为通过依据） |
|---|---|---|
| P1 | 12 块全可用；`research-brief` 三层 QA `verified==total`；TOC 条目覆盖；30 页压力测试无异常 | pdf 页数在 10–15；字数区间 |
| P2 | `preset` 生效且来源可打印；SVG 跳过报告一致；`layers.json` 字段齐全 | svg 文件 >2 KB；层数 =5 |
| P3 | 页数一致；底图 dpi；对照图；`changes.md` 一致 | 替换率 ≥90%（仅提示） |
| P4 | 五案例齐全；V9–V20；解包复验 | zip 体积、文件数 |

**不用作验收依据的**：文件大小、页数范围、PNG 数量、字数区间——它们只用来发现"明显异常"，单独出现不能证明功能通过。

---

## 11. 风险

| 编号 | 风险 | 触发信号 | 处理 |
|---|---|---|---|
| R1 | LibreOffice 对 docx 的分页与 Word 不同 | 同一 docx 转换页数与预期差 >2 | 以 PDF 为准，Word 分页差异写进交付说明；不承诺"Word 里页数完全一致" |
| R2 | 缺字体导致度量偏差 | `missingFonts` 非空 | 换字体或记录替换；`PDF-02` P1 |
| R3 | 表格跨页表头失效 | 长表格翻页后无表头 | L2 `DOC-13` 扩展检查 `tblHeader` 属性 |
| R4 | SVG 在不同软件表现不一 | Inkscape/AI 打开错位 | 待验证项 V18；验证前不承诺编辑体验 |
| R5 | 改版遮挡留痕 | 对照图可见边缘 | 提示"背景复杂度高"，改人工重排；不做自动修复 |
| R6 | 长文档生成慢 | 30 页 >10 s | 只做提示，不优化到牺牲正确性 |
| R7 | 阶段越界（P0 用了 P1 能力） | `DOC-16` 触发 | 停，回到 P0 范围 |
| R8 | 假验收（用文档属性冒充产物检查） | L2/L3 输入缺失却报通过 | 报告必须打印 `verified/skipped`；`skipped>0` 不许签收 |
| R9 | 回归基线漂移（PPTX 时间戳） | 文件 sha 变了但内容没变 | 用 zip 成员 CRC 比对，不用文件 sha |

---

## 12. 版本与发布

**发布版本号正常迭代**，不做特殊编号：

| 版本 | 内容 | 发布条件 |
|---|---|---|
| `1.4.1` | 本轮文档修订（R2） | 本文与操作书定稿 |
| `1.5.0` | 文档分支 P0：4 页最小闭环 | A0-* 全过 |
| `1.6.0` | 文档分支 P1：10–15 页 + 30 页压力 | A1-* 全过 |
| `1.7.0` | 固定版面 + SVG + 分层 PNG（P2） | P2 判据全过 + 回归 |
| `1.8.0` | 既有 PDF 受限改版（P3） | P3 判据全过 + 回归 |
| `1.9.0` | 案例、验证报告、发布包（P4） | P4 判据全过 |

**关于"版本号"的那条外部建议，本意是功能开发顺序，不是发布编号**：它要求的其实是"先做编号小的功能、功能顺序不能颠倒"。本方案的阶段顺序 P0 → P1 → P2 → P3 本来就单调，因此这条要求天然满足；发布号只是跟着发版一次次递增（1.5.0 → 1.6.0 → …）。

**本文标题里的 `v2` 是功能线名称**（"文档与固定版面分支"这条线），不是版本号——不要把 `v2.0.0` 之类的编号写进 CHANGELOG 或交付文件名。

**阈值版本约定**：`schema/thresholds.json` 的 `version` 只随**阈值数字**变化；文档与规则的改动不改它（沿用 v1.3 起的老规矩）。

---

## 附录 A · 术语速查

| 术语 | 一句话 |
|---|---|
| spec-first | 先写 JSON 契约，再由脚本生成产物，产物可被机器复查 |
| 三层 QA | Schema / DOCX / PDF·Render，按输入对象分，缺输入记 SKIP |
| 静态目录占位 | P0 的目录页是普通文本，不是自动目录 |
| 视觉验收 | 人看渲染图，检查层级、断页、题注、留白——机器不替代 |
| 回归基线 | 冻结的代码/schema/产物 CRC/QA 摘要，每阶段收尾比对 |
| 假验收 | 拿不到真实输入却报"通过"；本方案用 `verified/skipped` 拦住它 |

## 附录 B · 八条硬边界（六条沿用 + 两条新增）

1. 不虚构事实、素材、来源、数字。
2. 不设默认风格：每次从内容与受众推，不给"通用模板"。
3. 不覆盖用户原件；所有改动写进新文件并带版本号。
4. 不把"理论上可以"写成"已经实现"。
5. 不用文件大小、页数、字数当作功能通过的证据。
6. 不在未验证前声明设计软件兼容性（SVG/分层 PNG）。
7. **新增**：不改既有 PPT 分支的默认行为；`--preset` 必须可选，不带参时输出与 v1.4.0 一致。
8. **新增**：每个阶段只调用已实现的能力；越界即停（`DOC-16`）。

## 附录 C · 命令序列（复制即用，PowerShell）

```powershell
$PY   = "<python.exe>"
$NODE = "<HOST_ROOT>/.desktop-bin/node.cmd"
$CLI  = "<office-cli.mjs>"
$REPO = "<仓库根目录>"
Set-Location $REPO\examples\docs\report-min
```

```powershell
# P0 全链路
& $PY ..\..\..\scripts\build_docx.py --doc source\doc.json --out build\report-min.docx --check-only
& $PY ..\..\..\scripts\build_docx.py --doc source\doc.json --out build\report-min.docx
& $NODE $CLI convert --input build\report-min.docx --output build\report-min.pdf
& $PY ..\..\..\scripts\render.py --pptx build\report-min.pdf --out-dir build\preview --dpi 150
& $PY ..\..\..\scripts\qa_doc.py --doc source\doc.json --docx build\report-min.docx --pdf build\report-min.pdf --preview-dir build\preview --out build\qa.md --json build\qa.json
```

## 附录 D · 三个决策点（需要你拍板）

| 编号 | 决策 | 选项 | 我的建议 |
|---|---|---|---|
| D1 | P0 样张页数 | 4 页（推荐）/ 6 页（需补内容到每页 600–900 字） | 4 页：先证明链路，不证明长文档 |
| D2 | 版本号策略 | 发布号正常迭代（1.5.0 → 1.9.0，推荐）/ 按功能线重新编号 | 正常迭代：“v2”只是功能线名称，不写进版本号 |
| D3 | SVG 位图化默认开关 | 默认关闭（推荐，跳过并报错）/ 默认开启（图能出但文字少了） | 默认关闭：宁可少一个 chart，也不静默降级 |

## 检查清单

- [ ] 四种形态的边界与 1.2 的"不做"清单我已确认
- [ ] 1.2.1 的"承诺/不承诺"与 1.2.2 的 recolor 口径可接受
- [ ] P0 只做 6 种块、4 页样张、无自动目录——范围我认可
- [ ] P1 的 10–15 页样张与 30 页压力测试定位清楚
- [ ] 第 7 章 preset 是唯一权威，别处不再重复数值
- [ ] 第 8 章 QA 映射表覆盖 DOC/PRT/PDF/EXP 全部 23 项
- [ ] 第 9 章阶段依赖没有"用未来能力做当前阶段"的情况
- [ ] 第 10 章验收不再以文件大小/页数为主证据
- [ ] 第 12 章发布号为正常迭代，且阶段顺序 P0 → P1 → P2 → P3 未颠倒
- [ ] 附录 B 八条硬边界、附录 D 三个决策点我已表态
