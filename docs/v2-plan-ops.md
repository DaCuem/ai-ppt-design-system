# AI PPT 设计系统 · v2 补全操作书（R2）

| 项 | 内容 |
|---|---|
| 文档 | 操作书（执行层）。决策边界看 [v2-plan.md](v2-plan.md)；方向基准看 [lighthouse-plan.md](lighthouse-plan.md) |
| 版本 | R3（2026-11-15）。R2→R3 与灯塔对齐的逐条改动见 [v2-revision-notes.md](v2-revision-notes.md) |
| 范围 | 覆盖阶段 1A–2C（旧称 P0–P4）。**阶段 0（方向证据）另有专册** [stage0-direction-evidence.md](stage0-direction-evidence.md) |
| 读者 | 我（执行）。你验收时只需看 §2.6 / §3.3 / §4.2 与附录 |
| 环境 | Windows 11 + PowerShell 7。**本文所有命令在 PowerShell 下验证，CMD 语法不保证** |
| 原则 | 每步都有命令、预期输出、证据落点、失败处理；没有证据的步骤不算完成 |

---

## 0. 怎么用这份文档

### 0.1 与规划书的分工

| 问题 | 看哪份 |
|---|---|
| 为什么做、不做什么、范围到哪 | 规划书 1、9、12 章 |
| 具体敲什么命令、文件写到哪、失败怎么办 | 本文 |
| 某条阈值是多少 | 规划书第 7 章（唯一权威）与 `schema/presets.json` |
| 某项 QA 判什么、需要什么输入 | 规划书第 8 章（映射表）+ 本文各阶段 |
| 阶段 0（方向证据 / A·B 消融）怎么做 | [stage0-direction-evidence.md](stage0-direction-evidence.md)（**本文不重复**） |

### 0.2 R1 的十二处"只写了做什么、没写怎么做"，本轮逐条补位

| 编号 | R1 的缺口 | 本轮补在哪 |
|---|---|---|
| G1 | 只说"P0 样张 6 页"，没给逐页内容 | §2.6（4 页逐页定稿） |
| G2 | `doc.schema.json` 只给骨架，没给块的必填字段 | §2.5（六块逐字段表 + 片段） |
| G3 | 命令里 `python`/`node` 混用，没有统一变量 | §1.1（`$PY`/`$NODE`/`$CLI`，全文统一） |
| G4 | `qa_doc.py` 只收 `--doc`，却声称检查 docx/pdf 属性 | §1.5（三层输入）+ 规划书 6.4 |
| G5 | 目录域在 P0 与 P1 之间越界 | §2.4（P0 明确禁用 `toc`） |
| G6 | 表格跨页只说"要 tblHeader"，没给 XML | §3.4（三段 OOXML） |
| G7 | 页数自洽的判据含糊（"目录末条 = 页数"） | §3.6（四段判据） |
| G8 | SVG 只列支持元素，没说 `chart/table` 怎么办 | §4.3（跳过/位图化二选一 + 报告格式） |
| G9 | `layers.json` 只有 order/name/file，缺 bounds/opacity/source/rasterized | §4.4（完整字段） |
| G10 | 改版 `recolor` 的 95% 无口径，伪代码有未定义的 `from_mix` | §5.3（口径定义 + 伪代码删 `from_mix`） |
| G11 | 每阶段"收尾回归"没给命令与判据 | §7（一份回归脚本 + 期望输出） |
| G12 | 验收证据散落，不知存哪 | §1.6（证据落点表） |

---

### 0.3 阶段编号映射（本文与灯塔、规划书怎么对应）

| 旧编号（本章名） | 灯塔阶段 | 规划书阶段 | 发布版本 |
|---|---|---|---|
| —（另一册） | 阶段 0 方向证据 | 阶段 0 | 1.5.0 |
| P0 | 阶段 1（文档）·骨架 | 阶段 1A | 1.6.0 |
| P1 | 阶段 1（文档）·完整 | 阶段 1B | 1.7.0 |
| P2 | 阶段 2（固定版面/PDF）·设计导出 | 阶段 2A | 1.8.0 |
| P3 | 阶段 2·PDF 改版 | 阶段 2B | 1.9.0 |
| P4 | 阶段 2·验收发布 | 阶段 2C | 1.10.0 |

**本文章名继续用 P0–P4**，因为它们已写进 S1–S12 步骤号与验收表；看到 P0 就当作“阶段 1A”读。保留旧编号是为了不重写三百多行步骤，不代表两套制度。

## 1. 通用约定（所有阶段共用）

### 1.1 环境变量（PowerShell，全文只用这四个）

```powershell
$PY   = "<python.exe>"
$NODE = "<HOST_ROOT>/.desktop-bin/node.cmd"
$CLI  = "<office-cli.mjs>"
$REPO = "<仓库根目录>"
```

规则：

1. 本文出现的每条命令都用 `& $PY` / `& $NODE` / `& $CLI`，**不许**直接写 `python` 或 `node`。
2. `office-cli.mjs` 只通过绝对路径 `$CLI` 调用，不假设当前目录有它。
3. 每段命令前先确认位置：`Set-Location $REPO\examples\docs\<case>`（或项目目录）。
4. 拿不到退出码就到不了验收：命令后跟 `"exit=$LASTEXITCODE"`。

### 1.2 preset 生效顺序（引用规划书第 7 章，不重复数值）

`--preset` > `meta.pageSize` 推导 > 默认 `slide-16x9`。任何报告头必须打印：

```text
preset: print-A4 (source: cli)
```

`source ∈ {cli, meta, default}`。缺这一行视为未验证。

### 1.3 退出码约定

| 码 | 含义 | 该做什么 |
|---|---|---|
| 0 | 成功（允许 P2/P3 存在） | 继续 |
| 1 | 成功但有问题：有 P0/P1、缺字体、有不支持元素 | 修问题或写进遗留项，不得当"全绿" |
| 2 | 失败：参数错、I/O 错、schema 校验失败 | 停，读报错原文 |

### 1.4 项目目录约定（沿用 v1.4.0）

```text
<project>/
├── records/    requirement-summary / copy-deck / approved / qa / changes
├── source/     doc.json + assets/
├── build/      .docx / .pdf / preview/ / qa.json / logs/
└── deliver/    带版本号的交付副本
```

### 1.5 三层 QA 的输入约定

| 层 | 命令参数 | 缺了会怎样 |
|---|---|---|
| L1 Schema | `--doc source/doc.json` | 必填，缺了直接退出码 2 |
| L2 DOCX | `--docx build/x.docx` | 缺了 → 该层检查记 `SKIP`，报告里 `skipped+=1` |
| L3 PDF/Render | `--pdf build/x.pdf --preview-dir build/preview` | 缺了 → 同上 |

**硬规则**：`skipped > 0` 的报告不许签收；签收条件写死为 `verified == total`（P0 例外：L2 只实现 `DOC-13` 的 docx 结构核对，其余 L2 项记 SKIP 并在报告写明"P1 实现"）。

### 1.6 证据落点（每阶段收尾必须齐）

| 证据 | 存哪 | 谁生成 |
|---|---|---|
| 命令与输出 | `build/logs/<step>.log` | 执行时 `2>&1 \| Tee-Object` |
| QA 报告 | `build/qa.md` + `build/qa.json` | `qa_doc.py` |
| 结构与页数 | `build/pdf-manifest.json` | `export_pdf.py` |
| 页面图 | `build/preview/page-0001.png` … | `render.py` |
| 人工复核 | `build/qa.md` 末节「目视复核」 | 手写 |
| 回归结果 | `build/regress.json` | `regress_ppt.py`（P0 新增，只读） |
| 最终产物 | `deliver/<name>_<version>.<ext>` | 手工复制 + 记录 sha |

---

### 1.7 知识 ID 与决策记录（阶段 0 起强制）

每个关键设计决策都要能回答“读了哪几条规则”，否则知识库就是资料库。两件事：

1. **记录**：每个产物目录下写 `knowledge_refs.json`（格式见 [stage0-direction-evidence.md](stage0-direction-evidence.md) 第 4 节）。只记关键决策（方向、版式结构、色彩角色、字号层级），不记微调。
2. **编号**：被实际引用过的高价值规则才分配 `KP-<文件短码>-<两位序号>`（如 `KP-LAY-05`）。**本轮（R3）只定规则不回填**，第一轮实验后再给被用过的规则编号。

不记的后果：无法证明知识参与了决策，阶段 0 的结论只能停在“已接入工作流”一级。

---

## 2. P0（阶段 1A）：文档最小闭环（4 页 / 6 种块）

### 2.1 完成定义（与规划书 9.5 同一份，A0-1…A0-7）

A0-1 schema 通过 → A0-2 docx 生成 → A0-3 pdf 4 页 → A0-4 4 张 PNG → A0-5 三层 QA（L2 仅 docx 结构项）→ A0-6 目视 6 项 → A0-7 PPT 回归。**七条全过才算 P0 完成。**

### 2.2 前置检查（三条命令，任一失败先修环境）

```powershell
& $PY -c "import docx, PIL, lxml; print('docx', docx.__version__)"
& $NODE $CLI capabilities --json | Select-String 'conversions'
& $PY $REPO\scripts\probe.py --json | Select-String 'python|pptx|docx|Pillow'
```

预期：`docx 1.2.0`；capabilities 里有 `doc/docx/odt` 转换；probe 显示 `python-docx 1.2.0`、`Pillow 12.3.0`。

### 2.3 文件清单与创建顺序（六步，前一步不通不写下一步）

| 顺序 | 文件 | 规模 | 先写它的理由 |
|---|---|---|---|
| S1 | `schema/doc.schema.json` | 300–380 行 | 契约先行；后续两个脚本共用同一份校验规则 |
| S2 | `scripts/lib/docload.py` | 180–220 行 | 提供 `load_doc/validate/iter_blocks`，让 build 与 qa 共用块遍历与错误定位 |
| S3 | `scripts/build_docx.py` | 550–700 行 | 生成侧；只依赖 S1+S2 |
| S4 | `scripts/export_pdf.py` | 220–280 行 | 转换 + manifest；不依赖 docx 内部结构 |
| S5 | `scripts/qa_doc.py` | 600–750 行 | 最后写：它要同时读 doc/docx/pdf 三类输入 |
| S6 | `examples/docs/report-min/` | 4 页样张 | 有了前五个才谈得上跑通 |

### 2.4 P0 明确禁用（越界即停）

- 禁用块类型：`quote` / `callout` / `code` / `toc` / `bibliography` / `footnoteGroup` → `DOC-16`（P0）。
- 禁用参数：`build_docx.py --toc`（P1 才有）；`meta.toc.enabled = true` 直接退出码 1。
- 禁用命令：`export_svg` / `export_layers` / `pdf_rebuild`（P2、P3）。

### 2.5 六块逐字段（`schema/doc.schema.json` 要点）

| 块 | 必填 | 可选 | 约束 |
|---|---|---|---|
| `heading` | `type`,`level`,`text` | — | `level ∈ {1,2,3}`；同级不能跳（`DOC-01`） |
| `paragraph` | `type`,`runs[]` 或 `text` | `align`,`indentFirstLine`,`spaceAfterPt` | `runs[].text` 必填；单段 ≤300 字（`DOC-04` P2） |
| `list` | `type`,`items[][]` | `ordered`,`level` | 每项是 runs 数组；`ordered=true` 时编号由生成器维护 |
| `table` | `type`,`header[]`,`rows[][]` | `caption`,`source`,`colWidthsIn[]`,`align[]`,`zebra` | `caption`/`source` 缺失 → `DOC-05`/`DOC-06`（P0） |
| `figure` | `type`,`src` | `widthIn`,`caption`,`source`,`alt` | `src` 相对 `doc.json` 所在目录解析（与 deck 同一约定） |
| `pageBreak` | `type` | — | 连续两个分页 → 另记 P3 提示 |

```json
{
  "$schema": "doc.schema.json",
  "meta": {
    "title": "标题", "pageSize": "A4", "orientation": "portrait",
    "marginsIn": {"top": 0.9, "bottom": 0.9, "left": 1.0, "right": 1.0},
    "toc": {"enabled": false, "depth": 3}
  },
  "tokens": {"fonts": {"body": "Microsoft YaHei", "display": "Source Han Serif SC"},
             "colors": {"fg": "#1A1A1A", "bg": "#FFFFFF", "accent": "#8C3A1E"},
             "type": {"h1": 20, "h2": 15, "h3": 12.5, "body": 10.5, "caption": 9, "table": 9.5}},
  "sections": [
    {"id": "s1", "blocks": [
      {"type": "heading", "level": 1, "text": "一、研究问题"},
      {"type": "paragraph", "runs": [{"text": "正文第一段。"}]},
      {"type": "table", "caption": "表 1 · 三个测点", "source": "演示数据（虚构）",
       "header": ["测点", "覆盖率", "温度 ℃"], "rows": [["A", "31 %", "34.1"]]},
      {"type": "figure", "src": "assets/fig-01.png", "widthIn": 6.27,
       "caption": "图 1 · 冠层剖面示意", "source": "作者绘制"},
      {"type": "pageBreak"}
    ]}
  ]
}
```

### 2.6 样张 `report-min`：4 页逐页定稿（G1 补位）

| 页 | 内容 | 块序列 | 字数预算 |
|---|---|---|---|
| 1 | 封面：标题 + 副标题 + 作者/机构/日期 | `paragraph`×3（居中对齐，不新建样式） | 60–90 |
| 2 | 一、研究问题：3 段（背景 / 问题 / 本文做什么） | `heading`1 + `paragraph`×3 | 260–330 |
| 3 | 二、观测方法：1 段方法 + 1 张 5 行表 | `heading`1 + `paragraph`1 + `table`1 | 220–280 |
| 4 | 三、结论：3 条要点 + 1 张图 + 图注 | `heading`1 + `list`1（3 项）+ `figure`1 | 160–220 |
| — | 合计 | **无目录页**（P0 不实现 `toc`） | **700–900 字** |

规则：**不许为凑页数加空白、连续分页或无意义段落**。页数由内容决定；4 页是这四段内容在 A4 / 10.5 pt / 1.5 行距下的自然页数。若实测只有 3 页 → 补内容，不加 `pageBreak`。

### 2.7 十二步（S1–S12，每步带预期与失败处理）

| 步 | 动作 | 命令 / 判据 | 失败处理 |
|---|---|---|---|
| S1 | 写 `doc.schema.json` | `$PY -c "import json;json.load(open('schema/doc.schema.json',encoding='utf-8'))"` 先过语法 | 语法错：读报错行号 |
| S2 | 写 `docload.py` | 单测：`load_doc` 对 `qa-doc-bad.json` 报出 5 类错误 | 定位不清 → 错误消息带 `path` |
| S3 | 写 `build_docx.py` | `--check-only` 正例 rc=0、bad fixture rc=1 | 中文回宋体 → 查 `w:eastAsia` |
| S4 | 写 `export_pdf.py` | `convert` 出 pdf，manifest `pageCount == 4` | 页数不符 → 先看分节与字号 |
| S5 | 写 `qa_doc.py`（L1+L3） | 正例 `P0=0`；bad fixture 命中指定 ID | 假通过 → 检查是否真读了输入 |
| S6 | 建 `report-min` 目录 | 四件齐：`brief.md`、`source/doc.json`、`source/assets/fig-01.png`、`README.md` | 缺图 → PIL 画示意图并标"示意图" |
| S7 | 构建 docx | `& $PY $REPO\scripts\build_docx.py --doc source\doc.json --out build\report-min.docx` | rc=1 → 按报错改 doc.json |
| S8 | 转 pdf | `& $NODE $CLI convert --input build\report-min.docx --output build\report-min.pdf` | 缺字体告警 → 换字体或记录 |
| S9 | 渲 4 张 PNG | `& $PY $REPO\scripts\render.py --pptx build\report-min.pdf --out-dir build\preview --dpi 150` | 无渲染器 → 退出码 2，停 |
| S10 | 三层 QA | §2.8 三条命令 | `skipped>0` → 补输入再跑 |
| S11 | 目视复核 | 6 项清单写进 `build/qa.md` 末节 | 不过 → 回 S7 改版式参数 |
| S12 | PPT 回归 | `regress_ppt.py` 输出 `crc_diff=0`、四份 QA 全 0 | 有差异 → 停，查是否动了旧代码 |

### 2.8 P0 的三条 QA 命令（照抄）

```powershell
Set-Location $REPO\examples\docs\report-min
& $PY $REPO\scripts\qa_doc.py --doc source\doc.json --out build\qa-schema.md 2>&1 | Tee-Object build\logs\qa-schema.log
& $PY $REPO\scripts\qa_doc.py --doc source\doc.json --docx build\report-min.docx `
      --pdf build\report-min.pdf --preview-dir build\preview `
      --out build\qa.md --json build\qa.json 2>&1 | Tee-Object build\logs\qa-full.log
& $PY $REPO\scripts\regress_ppt.py --repo $REPO --json build\regress.json
```

预期（P0 阶段）：

```text
inputs: doc=ok docx=ok pdf=ok preview=ok
verified=18 skipped=2   # L2 中 DOC-12 及表格跨页项，P1 实现
SUMMARY P0=0 P1=0 P2=… P3=… exit=0
```

`skipped` 的项必须在 `build/qa.md` 写明"P1 实现"，否则视为假验收。

### 2.9 两份 fixture（故意做错，验证检查器真的会报）

| fixture | 应触发 |
|---|---|
| `demo/docs/qa-doc-bad.json` | `DOC-01`、`DOC-03`、`DOC-05`、`DOC-06`、`DOC-09`、`DOC-10`、`DOC-11` |
| `demo/docs/qa-doc-edge.json` | `DOC-04`、`DOC-07`、`DOC-08`、`DOC-14`、`DOC-15` |

要求：命中项与上表**逐项对上**；多报少报都改检查器，不许改 fixture 迁就。

### 2.10 P0 失败处理表

| 现象 | 原因（实测经验） | 处理 |
|---|---|---|
| Word 里中文是宋体 | 只设了 `run.font.name`，没设 `w:eastAsia` | 双写 `rFonts`；重跑 S7 |
| PDF 页数与预期差 >1 | 用了 slide 档字号/页边距 | 检查 `preset:` 行的来源 |
| 表格在页底断成两半 | 缺 `cantSplit` | 所有行加 `<w:cantSplit/>` |
| 图片超出页高 | 未按页高反算宽度 | 走 80% 页高规则并提示"已缩放" |
| 末页空白 | 末尾多一个 `pageBreak` | 删掉；页数由内容决定 |
| 出现目录页但功能没有 | P0 误用 `toc` | 删块；目录只能作为正文段落 |
| `missingFonts:["Courier"]` | Word 默认样式带入 | 显式把 `Normal` 设为微软雅黑后重跑 |

---

## 3. P1（阶段 1B）：文档完整能力（10–15 页正文 + 30 页压力测试）

### 3.1 新增能力

`quote` / `callout` / `code` / `toc` / `bibliography` / `footnoteGroup`（footnoteGroup 先降级为章末注释）+ L2 QA 扩展（`DOC-12`、`DOC-13` 的 `tblHeader` 检查）+ TOC 域。

### 3.2 三个不同用途的文档（不再混用页数）

| 用途 | 目录 | 页数 | 说明 |
|---|---|---|---|
| 开发样张 | `examples/docs/research-brief/` | **10–15 页** | 真实内容，验证复杂排版与 TOC |
| 压力测试 | `examples/docs/technical-report/` | **28–32 页** | 只测长文档：分页、表格跨页、图注分离、生成耗时 |
| 交付案例（P4 定稿） | `examples/docs/technical-report/` 案例版 | **24–28 页** | 从压力版裁剪，不带测试标记 |

规则：**30 页不是最早的样例**。先 10–15 页跑通，再上 30 页。

### 3.3 `research-brief` 逐章页数（10–15 页）

| 章 | 页 | 内容要点 |
|---|---|---|
| 封面 | 1 | 标题/副标题/作者/日期 |
| 目录 | 1 | `toc` 域（Word 里 F9 更新） |
| 一、问题 | 1–2 | 3 段 + 1 表 |
| 二、方法 | 2–3 | 流程列表 + 1 图 + 1 提示框 |
| 三、结果 | 3–4 | 2 图 + 1 表 + 引文 1 处 |
| 四、讨论 | 1–2 | 3 段 |
| 五、结论 | 1 | 3 项要点 |
| 参考 | 1 | `bibliography` 6–10 条 |

### 3.4 表格跨页：三段 OOXML（G6 补位）

```python
# 1) 表头行重复
trPr = row._tr.get_or_add_trPr()
trPr.append(parse_xml(
    r'<w:tblHeader xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
# 2) 行不拆页
trPr.append(parse_xml(
    r'<w:cantSplit xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
# 3) 表头行高 22 pt（避免窄行跨页时被压扁）
row.height = Pt(22)
row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
```

验证：L2 检查（`DOC-13` 扩展项）读 `build/x.docx`，`w:tblHeader` 出现次数 == 表格数。

### 3.5 tokens 变体：先做 3 组，不铺 12 组

| 组 | 字体（display / body） | 强调色 | 用途 |
|---|---|---|---|
| T1 学术 | 思源宋体 / 微软雅黑 | `#8C3A1E` | 报告 |
| T2 技术 | 微软雅黑 / 微软雅黑 | `#1F6FEB` | 技术文档 |
| T3 简报 | Lato / 微软雅黑 | `#C2410C` | 对外简报 |

验收：换组后重建，三层 QA 仍全绿（允许 P2 消失，不许新增 P0/P1）。

### 3.6 页数自洽的四段判据（G7 补位，替代"目录末条 = 页数"）

1. `pdf-manifest.json` 的 `pageCount` 落在目标区间（research-brief 10–15；压力测试 28–32）；
2. 目录覆盖：每个 `heading` 的 `level ∈ {1,2,3}` 在目录里出现一次（L2）；
3. 抽查 3 条目录条目，页码与 PDF 中该标题实际页码一致（人工，写进 `build/qa.md`）；
4. 无空白页、无重复页码、无分页异常（L3 + 人工看）。

四条都过才算自洽。**不许**用"目录最后一条页码 = 总页数"代替。

### 3.7 P1 失败处理（摘要）

| 现象 | 处理 |
|---|---|
| TOC 打开是空的 | LibreOffice 转换不触发域更新；交付说明写"Word 里 F9 更新"，PDF 用 L2 检查域存在 |
| 30 页生成 >30 s | 记录耗时作辅助指标，不为提速牺牲正确性 |
| 图注与图被分到两页 | 图与图注放进同一 `keep_with_next` 容器 |

---

## 4. P2（阶段 2A）：固定版面与设计型导出

### 4.1 preset 生效（规划书第 7 章是唯一权威）

```powershell
& $PY $REPO\scripts\build_pptx.py --deck source\poster.json --out build\poster.pptx --preset print-A4
```

预期：报告头 `preset: print-A4 (source: cli)`；不带 `--preset` 时输出与 v1.4.0 逐字节一致（CRC 比对见 §7）。

### 4.2 A4 海报规格（案例验收用）

| 项 | 值 |
|---|---|
| 画布 | 8.268 × 11.693 in |
| 主标题 | 44–56 pt |
| 正文 | 10–11 pt，行距 1.4–1.6 |
| 元素 | ≤14，文本块 ≤4 |
| 满版图 | `zRole:"bgDecor"`，边距 0（出血） |

### 4.3 SVG：支持元素与两种处置（G8 补位）

支持：`rect` / `line` / `ellipse` / `polygon` / `text` / `image`。
不支持：`chart` / `table` / `icon`。

```powershell
# 默认：不支持就跳过并 rc=1（推荐，不静默降级）
& $PY $REPO\scripts\export_svg.py --deck source\poster.json --slide 1 --out build\poster.svg
# 允许位图化：chart/table 先渲 PNG 再嵌 <image>
& $PY $REPO\scripts\export_svg.py --deck source\poster.json --slide 1 --out build\poster.svg --allow-raster
```

`skip-report.md` 四列：元素序号 / 类型 / 位置 / 处置（跳过·位图化）+ 原因。
SVG 的 `<desc>` 里写明"含位图化元素 N 个"。

### 4.4 `layers.json` 完整字段（G9 补位）

```json
{
  "canvas": {"wIn": 8.268, "hIn": 11.693, "px": 1241, "pxH": 1754, "dpi": 150},
  "layers": [
    {"name": "01-bg", "order": 1, "file": "layers/01-bg.png",
     "bounds": {"x": 0, "y": 0, "w": 8.268, "h": 11.693},
     "opacity": 1.0, "sourceElementIds": [], "rasterized": true, "note": "底色层"},
    {"name": "05-text", "order": 5, "file": "layers/05-text.png",
     "bounds": {"x": 0.6, "y": 0.6, "w": 7.0, "h": 10.5},
     "opacity": 1.0, "sourceElementIds": ["text-title", "text-body"], "rasterized": true,
     "note": "文字层是位图，改字请用 SVG 或回源 deck.json"}
  ]
}
```

验证：`EXP-02` 检查层数 =5、`rasterized` 齐全、`05-text.rasterized == true`；人工用 PS 逐层导入一次确认对齐（PNG 必须保留透明）。

### 4.5 待验证项（不得写成已完成）

| 编号 | 内容 | 现状 |
|---|---|---|
| V18-a | SVG 在 Inkscape / Illustrator 中可选中编辑文字与形状 | **本机无这两个软件，未验证**；验证前 README 不得声明"可在设计软件中编辑" |
| V18-b | 分层 PNG 在 Photoshop 中逐层导入位置准确 | 同上 |

---

## 5. P3（阶段 2B）：既有 PDF 的受限改版

### 5.1 五步流水线

1. 原 PDF → 逐页 PNG（150 dpi 底图）；2. 底图处理（`mask` 遮挡、`recolor` 颜色重映射）；3. 写 `rebuild.json`（覆盖层用 deck 元素语法）；4. `build` → `export` 新 PDF；5. 对照图 + `changes.md`。

```powershell
& $NODE $CLI render --input source\original.pdf --out-dir build\base --dpi 150
& $PY $REPO\scripts\pdf_rebuild.py --pdf source\original.pdf --spec source\rebuild.json --out-dir build\rebuild
```

### 5.2 `rebuild.json` 结构

```json
{
  "meta": {"version": "v1", "baseDpi": 150},
  "pages": [
    {"index": 1, "keep": "full",
     "mask": [{"x": 6.2, "y": 1.4, "w": 2.0, "h": 0.6, "fill": "#FFFFFF"}],
     "recolor": [{"from": "#B23A2E", "to": "#1F6FEB", "tolerance": 30,
                  "region": {"x": 0, "y": 0, "w": 8.268, "h": 11.693},
                  "exclude": ["photo"]}],
     "overlay": [{"kind": "text", "text": "2026 版", "x": 6.4, "y": 1.5, "w": 1.6, "h": 0.4,
                  "size": 14, "color": "#1A1A1A"}]}
  ]
}
```

`keep` 只两种：`full`（整页保留作底图）/ `none`（重排整页，等于从零做）。

### 5.3 `recolor` 口径与伪代码（G10 补位，删掉未定义的 `from_mix`）

口径见规划书 §1.2.2：`region` 内、sRGB 欧氏距离 ≤ `tolerance`（默认 30，范围 0–60）的像素硬替换；距离在 1.0–1.35 倍 `tolerance` 之间的边缘像素只做 50% 混合，**不计入分母**；分母 = `region` 内被判定为目标色的像素数（`hits`）；替换率 = `replaced / hits`，**仅作辅助提示（<90% 时提示）不作验收**。

```python
def recolor(img, rgb_from=(0xB2, 0x3A, 0x2E), rgb_to=(0x1F, 0x6F, 0xEB),
            tolerance=30, region=None, exclude_mask=None):
    px = img.load()
    x0, y0, x1, y1 = region or (0, 0, img.width, img.height)
    hits = replaced = 0
    for y in range(y0, y1):
        for x in range(x0, x1):
            if exclude_mask and exclude_mask.getpixel((x, y)):
                continue
            r, g, b = px[x, y][:3]
            d = math.dist((r, g, b), rgb_from)
            if d <= tolerance:
                px[x, y] = (*rgb_to, 255); hits += 1; replaced += 1
            elif d <= tolerance * 1.35:
                px[x, y] = (*mix(rgb_to, (r, g, b), 0.5), 255)
    return {"hits": hits, "replaced": replaced,
            "ratio": (replaced / hits) if hits else 1.0}

def mix(a, b, t):          # 逐通道 a*(1-t) + b*t，与 qa_check.py 的 blend_alpha 同语义
    return tuple(round(a[i] * (1 - t) + b[i] * t) for i in range(3))
```

### 5.4 验收（P3）

| 判据 | 怎么查 |
|---|---|
| 页数一致 | `pdf-manifest.json` 的 `pageCount` == 原 PDF 页数 |
| 底图 dpi ≥150 | `render.py --dpi 150` + 输出 PNG 尺寸 |
| 遮挡无残留 | 对照图逐页看（人工），写进 `changes.md` |
| 颜色一致 | 对照图同区域目视一致；替换率只记录 |
| `changes.md` 与实际一致 | 条目数 == `rebuild.json` 的 mask + recolor + overlay 条数 |

---

## 6. P4（阶段 2C）：综合验收与发布（不写新功能）

1. 五案例定稿：`technical-report`（24–28 页）、`research-brief`（6 页版）、`one-pager`（2 页）、`a4-poster`（1 页）、`pdf-rebuild`（8 页）。
2. 跑 V9–V20（定义见 `docs/v2-validation-report.md`）。
3. 每案例必须有：`brief.md`、`source/`、`build/`（含三层 QA）、`deliver/`（带版本号）、`README.md`（四段：做了什么 / 怎么复现 / 真实数字 / 已知限制）。
4. 打包与解包复验：

```powershell
& $PY $REPO\scripts\pack.py --out dist\ai-ppt-design-system-v1.10.0.zip
Expand-Archive dist\ai-ppt-design-system-v1.10.0.zip -DestinationPath $env:TEMP\v110 -Force
& $PY "$env:TEMP\v110\ai-ppt-design-system\scripts\selfcheck.py" --strict
```

判据：解包副本 `selfcheck --strict` 全 0；`demo` QA 仍 `0/0/0/0`；新增六命令 `--help` 全部 rc=0。

---

## 7. 跨阶段回归（G11 补位）

每个阶段收尾跑同一段：

```powershell
& $PY $REPO\scripts\regress_ppt.py --repo $REPO --json build\regress.json
```

`regress_ppt.py`（P0 新增，只读）做四件事：

1. 读 `docs/ppt-baseline.json`，逐文件比 SHA-256（代码与 schema）；
2. 解四份 pptx 的 zip 成员 CRC，比基线差异数（PPTX 有时间戳，**不比文件 sha**）；
3. 跑 `qa_check.py` 于 demo + 三案例，比对 `SUMMARY` 行；
4. 输出 `{"sha_diff":0,"crc_diff":0,"qa_diff":0,"detail":[…]}`。

判据：三个 diff 全 0。任一非 0 → 停，先解释差异来源（改动意图、是否预期），不许"带差异继续"。

---

## 8. 风险触发的具体命令

| 触发信号 | 先跑这条 | 再看什么 |
|---|---|---|
| 文档中文字体回落 | `& $PY -c "from docx import Document;d=Document('build/x.docx');print(d.styles['Normal'].font.name)"` | 是 `None` 就说明没设样式 |
| PDF 缺字体 | `& $NODE $CLI capabilities --json` | `missingFonts` 列表 → 换可用字体 |
| 表格跨页断 | 解 docx 的 `word/document.xml`，搜 `tblHeader` | 没有就回 §3.4 |
| SVG 打开错位 | `Select-String -Path build\poster.svg -Pattern '<text|<image'` | 数位图化元素个数 |
| 改版颜色有残影 | 看 `changes.md` 的 `ratio` 与对照图 | 调小 `tolerance` 或排除边缘 |
| 回归有差异 | `Get-Content build\regress.json` | 差异清单逐条解释 |

---

## 9. 详细度自检（每阶段开工前读一遍）

1. 这一步的**输入文件**在哪？不存在就先建。
2. 命令**用 `$PY/$NODE/$CLI`** 了吗？
3. 预期输出写清了吗（rc、页数、文件、报告字段）？
4. 证据落在 §1.6 的哪一格？
5. 失败时先看哪一行报错、回到哪一步？
6. 有没有用到**还没实现**的能力（越界）？
7. 有没有把"文件大小/页数"当唯一证据？

## 检查清单

- [ ] 所有命令都用 `& $PY` / `& $NODE` / `& $CLI`，没有裸 `python` / `node`
- [ ] P0 只用了 6 种块、4 页样张、无 `toc`
- [ ] 三层 QA 的输入对象正确，`skipped>0` 不算通过
- [ ] 页数判据是四段（区间 / 覆盖 / 抽查页码 / 无异常），不是"目录末条"
- [ ] SVG 的 `chart/table` 有明确处置与报告，没静默丢弃
- [ ] `layers.json` 含 bounds / opacity / sourceElementIds / rasterized
- [ ] `recolor` 有口径、有伪代码、没有未定义变量
- [ ] 每阶段收尾跑了 `regress_ppt.py`，三个 diff 全 0
- [ ] 未验证项（V18-a/b）没有写成"已完成"
