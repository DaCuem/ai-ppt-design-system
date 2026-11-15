# 05 单页设计与制作

把 04 规划表的一行变成可构建的 `deck.json` 元素。本步是唯一写元素坐标的地方；页面做对的标准是「渲染图上焦点唯一、文字不出框、每页可编辑」。

## 输入

| 输入项 | 来源 | 必需 |
|---|---|---|
| 规划表一行 | `<project>/records/deck-outline.md` | 是 |
| 视觉令牌 | `<project>/source/deck.json` 的 `tokens` | 是 |
| 版式规则 | `../references/layout-recipes.md` 的对应 L 编号 | 是 |
| 元素契约 | `../schema/deck.schema.json` 的 `element` 定义 | 是 |
| 素材文件 | `<project>/source/assets/` 内已导入的图片 | 否 |

## 前置条件

- 04 已通过：每页有 1 个 L 编号与 1 个焦点位置，节奏检查全过；未过退回 `./04_deck-architecture.md` 步骤 8。
- `source/deck.json` 已有 `meta` + `tokens`，并能通过 `python scripts/build_pptx.py --check-only`；不通过先修令牌。
- `python scripts/probe.py` 的「渲染后端」栏至少 1 个可用；不可用先解环境，再进渲染。
- 已确认本页 `layout` 英文短名与 `gridTemplate`、`focalPosition` 三者齐备（三者是版式查重的签名）。

## 操作步骤

### 步骤 1–2 单页 7 问 → slide-spec

逐页回答 7 问：① 只讲哪个核心信息？② 信息间什么关系（并列/递进/因果/对比）？③ 用文字、图片、图表还是结构图？④ 视觉中心是什么？⑤ 阅读顺序是 Z 型、F 型还是单点？⑥ 哪些放大、哪些弱化（字号比 ≥ 1.6:1）？⑦ 与相邻页如何形成节奏变化（密→疏或换版式）？答案落进 `<project>/records/slide-spec.md`（12 栏，逐页一节，见 `../templates/slide-spec.md`）。判据：7 问答案在 12 栏各有落点。

### 步骤 3 坐标换算：12 栏 → 英寸

基准画布 13.333 × 7.5 in，边距 0.6 in（左）与 0.45 in（上），栏距 0.2 in，12 栏。单栏宽 `COL = (13.333 − 1.2 − 2.2) ÷ 12 = 0.8278 in`。

```
第 k 栏左边距      x(k) = 0.6 + (k − 1) × 1.0278        (k = 1…12)
跨 n 栏的宽度      w(n) = n × 0.8278 + (n − 1) × 0.2
纵向对齐基线       y(b) = 0.45 + b × 0.08               (b = 0,1,2…)
文本块高度         h = 字号pt × 行距 ÷ 72 × 行数 + 0.08   (上下 padding 各 0.04 in)
```

常用值：x(1)=0.60、x(2)=1.63、x(5)=4.71、x(7)=6.77、x(8)=7.79 in；w(3)=2.88、w(4)=3.91、w(6)=5.97、w(12)=12.13 in。48 pt 标题 1 行、行距 1.1 → h = 0.81 in；18 pt 正文 3 行、行距 1.45 → h = 1.17 in。

**规则：落坐标**
**触发条件**：写任一 `text` 元素。
**具体动作**：先定占几栏（用 `w(n)`），再按 `x(k)` 定左边，`y` 取最近基线，`h` 用上式四舍五入到 0.01 in，且 `y + h ≤ 7.05`（7.5 − 0.45）。
**常见错误**：凭感觉写 1.5/2.0 这类圆整值，导致左右不对齐、文本贴边；h 写小 0.1 in 后 06 阶段必报 OVF-01。

### 步骤 4 文案规则

- 标题用结论句：含动词或数值的判断，≤ 25 字；不写「关于 X 的介绍」。
- 每页总字数 ≤ 220 字；文本块 ≤ 4 个；bullet ≤ 6 条；单行正文 ≤ 34 字/行。
- 一个元素只承载一个语义；`role` 必须填（h1/h2/h3/lead/body/caption/source/pageNumber/kpi/label/quote），否则引擎回退默认字号。
- 图表必须给 `unit` 与 `source`；图片必须给 `alt`。

### 步骤 5 写成 elements（先骨架后细节）

顺序固定，每步都先过 `--check-only`：

1. 背景与页码：`bg` 写在 slide 上；页码 1 个 `text`（role=pageNumber，9–11 pt，右下）。
2. 焦点元素：标题或主图 1 个，面积 ≥ 页面 20%。
3. 证据块：≤ 4 个文本块，按栏对齐。
4. 图表/图片：`kind: chart|image`，带 `alt`、`unit`、`source`。
5. 装饰：默认为 0；出现必须承载分组边界或数据刻度，且有信息理由。
6. `notes`：写讲解要点与「本页不要出现」的内容。

**规则：不许整页栅格化**
**触发条件**：想用 1 张 PPT 截图/导出图当整页背景。
**具体动作**：拒绝。文字、表格、图表、页码必须是可编辑 `element`；位图只允许 `kind: image` 且覆盖面积 ≤ 60% 页面积，单页位图元素 ≤ 2 个，同时必须存在 ≥ 1 个 `text` 元素承载标题。
**常见错误**：把渲染图贴回去省事，交付后用户无法改字，且 06 阶段文字检查全部失效。

### 步骤 6 构建、渲染、看图

在项目根 `<project>/` 执行：

```bash
python scripts/build_pptx.py --deck source/deck.json --out build/deck.pptx --check-only
python scripts/build_pptx.py --deck source/deck.json --out build/deck.pptx
python scripts/render.py --pptx build/deck.pptx --out-dir build/preview --dpi 144
python scripts/render.py --pptx build/deck.pptx --out-dir build/preview --pages "1,3-5"
```

参数：`--deck` 输入 deck.json（图片 `src` 相对 deck.json 目录解析）；`--out` 输出 pptx，父目录自动创建；`--check-only` 只校验 schema、图片存在性与令牌可解析，不写文件；`--quiet` 只打印末行。退出码 0 成功 / 1 校验或构建失败（缺图、非法字段）/ 2 参数错误、输入不存在或依赖缺失。`--pptx` / `--out-dir` / `--pages`（省略即全部，可写 `1,2` 或 `1-3`）/ `--dpi`（默认 144）/ `--json`（stdout 只出一行 manifest）/ `--backend auto|dsh|env|soffice|com`。渲染产物为 `build/preview/slide-01.png`（0 填充宽度取页数位宽，> 99 页为 `slide-001.png`）与 `build/preview/render-manifest.json`（含 `pageCount`、`images[]`、`missingFonts[]`、`backend`）。退出码 2 = 所有后端不可用，按 stderr 的中文提示装渲染器或用 `--backend` 指定。脚本参数以 `scripts/README.md` 或各脚本 `--help` 为最终真值。

**规则：渲染后再改**
**触发条件**：任一页元素改完。
**具体动作**：只重渲受影响的页（`--pages`），看 3 处：焦点是否唯一、文字是否出框、`missingFonts` 是否为空。字体缺失不阻断构建，但在 `render-manifest.json` 里告警，必须在 06 之前处理。
**常见错误**：改完不看渲染图就宣称完成；字体回退导致中文变成宋体，全套字形走样。

### 步骤 6b 图像进版面（生成图 / 示意图专用，四关不过不放行）

| 关 | 做什么 | 判据 |
|---|---|---|
| 0 授权关 | 回读 `records/requirement-summary.md` 的「素材授权」区块（`references/imagery-guide.md` §5.0）：能否搜集 / 能否生成 / 指定工具 / 禁用来源 | 有答复且本次动作在允许范围内；**未答复 = 不许**，此时只画几何图形与留白，不得截图、搜图、跑生成 |
| 1 结构关 | 按 `references/imagery-guide.md` §5.2 写结构清单（零件 / 顺序 / 连接 / 朝向），再写提示词 | 四行都有值，提示词逐条对应 |
| 2 放大关 | 出图后按清单打勾，并 2 倍放大看连接处与边缘 | 无悬空件、无重复件、无卷边碎屑、无幻影垫块 |
| 3 密度关 | 算 版面宽度（英寸）× 110 = 所需像素 | 实际像素 ≥ 所需像素（IMG-01） |

四关中任一项不成立：先重画一次；再不成，按 `references/imagery-guide.md` §5 降档（真实素材 > 结构图 > 留白 ≥ 50% > AI 生成），并在交付说明写明「封面图从生成图降为结构图，原因是……」。**不要在同一条生成路径上无限重试**（§5.1：3 次即降档）。

### 步骤 7 逐页收口

每页按 06 的 12 个观察点自检，问题记入 `records/qa-report.md`；改完回步骤 6 重渲。

## 判断与分支

| 条件 | 动作 |
|---|---|
| 元素数 > 14 或文本块 > 4 | 合并信息块或删装饰；仍超则拆页（回 04） |
| 文字需 5 行才放下 | 先压字数到 ≤ 34 字/行 × 3 行；仍超则拆页 |
| 字号降到 16 pt 以下才放得下 | 不接受；删内容或拆页 |
| 图片比例与 3:2 / 16:9 / 1:1 不符 | 用 `fit` + `crop` 处理，不拉伸变形 |
| 图表数据未定 | 该页改文字页（L03），不画空图表 |
| 页与规划表不一致 | 以规划表为准；确需改则回 04 改表，再回本步 |
| `--check-only` 报缺图 | 先导入素材到 `source/assets/`，不改路径字符串硬指外部盘 |

## 输出文件

- `<project>/records/slide-spec.md`：每页一节，12 栏填满。
- `<project>/source/deck.json`：`slides[]` 逐页写 `id`（s01…）、`layout`、`gridTemplate`、`focalPosition`、`notes`、`elements[]`。
- `<project>/build/deck.pptx` 与 `<project>/build/preview/*.png`：构建与渲染产物。

## 完成标准

- 每页元素 ≤ 14 个、文本块 ≤ 4 个、bullet ≤ 6 条、字数 ≤ 220 字、标题 ≤ 25 字。
- 全部元素坐标满足 `x ≥ 0.6` 且 `x + w ≤ 12.73`，`y ≥ 0.45` 且 `y + h ≤ 7.05`（容差 0.02 in）。
- `build_pptx.py` 退出码 0；`render.py` 退出码 0 且 `missingFonts` 为空数组。
- 预览图张数 = 页数，命名连续 `slide-01.png` 起。
- 每页存在 ≥ 1 个 `text` 元素承载标题，位图覆盖 ≤ 60% 页面积。
- 每张图有 `alt`，每个图表有 `unit` 与 `source`。

## 失败处理

| 类别 | 具体处理 |
|---|---|
| 缺素材 | 用结构图（rect/line 组合）或数据表替代，`alt` 写清「示意」；不留空白图框 |
| 缺数据 | 该页改 L03 单点陈述，标注「数据待补」；图表元素整块删除，不留占位 |
| 工具不可用 | 渲染退出码 2：先跑 `probe.py` 看后端，按提示用 `--backend` 或 `PPT_RENDER_CMD` 指定；仍失败则只交 `--check-only` 通过的 pptx，并在 qa-report 记「未渲染验证」 |
| 内容冲突 | 规划表与文案冲突 → 以规划表结论为准，把文案差异记进需确认事项；不擅自改结论 |

## 检查清单

- [ ] 每页 7 问都有答案，已落进 slide-spec 的 12 栏
- [ ] 元素坐标用 x(k)/w(n)/y(b) 公式算出，无凭感觉的圆整值
- [ ] 每页 ≤ 14 个元素、≤ 4 个文本块、≤ 6 条 bullet、≤ 220 字
- [ ] 标题为结论句，≤ 25 字
- [ ] 全部元素在安全边距内，容差 0.02 in
- [ ] `build_pptx.py --check-only` 退出码 0
- [ ] `render.py` 退出码 0，`missingFonts` 为空
- [ ] 预览图数量 = 页数，命名从 slide-01.png 连续
- [ ] 无整页栅格化：位图覆盖 ≤ 60%，每页有可编辑标题文本
- [ ] 图片 100% 有 alt，图表 100% 有 unit 与 source
