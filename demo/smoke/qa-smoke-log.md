# 质检引擎 smoke 实测记录

本文件是 `scripts/qa_check.py` 与 `scripts/contact_sheet.py` 的实测证据：用了哪些输入、跑出什么结果、每类问题由哪份输入触发、哪些是误报并已修正、还有哪些已知限制。所有数字都来自本机真实运行，可用文末命令复现。

## 环境

| 项 | 值 |
| --- | --- |
| Python | 3.12.14（仓库指定运行时） |
| Pillow | 12.3.0 |
| python-pptx | 1.0.2 |
| 操作系统 | Windows，无 Office / WPS / COM |
| 渲染后端 | `dsh-native`（LibreOffice 光栅化，1920 × 1081 px，144 dpi，缺失字体：无） |
| 阈值文件 | `schema/thresholds.json`（版本见该文件的 `version` 字段） |

## 三份输入与结果

| 输入 | 页数 | P0 | P1 | P2 | P3 | 合计 | 退出码 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| `demo/smoke/qa-fixture.json`（故意埋问题的主样例） | 6 | 2 | 9 | 5 | 3 | 19 | 1 |
| `demo/smoke/qa-fixture-edge.json`（补充样例，覆盖主样例触发不到的检查） | 3 | 1 | 0 | 6 | 3 | 10 | 1 |
| `demo/smoke/deck.json`（真实 deck，`smoke.pptx` 的来源） | 3 | 0 | 2 | 0 | 1 | 3 | 0 |

主样例命中（按页序）：

```
P2 s02 REP-02 版式签名 签名 standard / G2 / ml 占 3/6 页 = 50% > 上限 35%
P1 s03 CON-01 body-lowcontrast 对比度 2.42:1 < 3.0:1（文字 #9AA0A6 / 背景 #F7F5F0，24pt，大字）
P1 s03 CON-01 caption-lowcontrast 对比度 2.42:1 < 4.5:1（文字 #9AA0A6 / 背景 #F7F5F0，12pt）
P1 s03 OVF-01 body-overflow 文本溢出 0.17in（needed 1.67in > box 1.50in；3 行）
P1 s03 REP-01 版式签名 与上一页第 2 页版式签名相同（standard / G2 / ml）
P0 s04 OVF-02 title-too-long 文本严重溢出 1.57in（needed 2.57in > box 1.00in，超出 157% > 上限 30%）
P1 s04 REP-01 版式签名 与上一页第 3 页版式签名相同（standard / G2 / ml）
P2 s04 CNT-04 title-too-long 标题 26 字 > 上限 25 字
P2 s04 COL-01 body-brandcolor color 引用的颜色 'brand-blue' 既不在 tokens.colors 也不是合法 #RRGGBB
P3 s04 REP-03 focalPosition 焦点位置 ml 连续 3 页相同（第 2–4 页）
P1 s05 CNT-01 proof-shot image 缺 alt（src=qa-shot.png）
P1 s05 CNT-02 trend chart 缺 unit 与 source（chartType=column）
P2 s05 DEN-02 页面 文本块 5 个 > 上限 4 个（已排除 role=caption/source/pageNumber）
P0 s07 GEO-01 decor-blob 越出画布（右 0.77in；画布 13.333×7.500 in）
P1 s07 STR-01 slide.id slide id 不连续：第 6 页应为 s06，实际为 s07
P1 s07 TYP-01 kpi-small 字号 17.0pt < h2 的 minPt 24pt
P2 s07 GEO-03 eyebrow 越出安全边距：上 0.25in（应 ≥ 0.45in）
P3 s07 CNT-03 页面 缺页码（layout=data）
P3 s07 COL-02 kpi-small 字号 17.0pt 不在 tokens.type 阶梯（10、12、22、24、26、30、44、48、64 pt）
SUMMARY P0=2 P1=9 P2=5 P3=3 issues=19 pages=6 exit=1
```

补充样例命中：

```
P2 s01 COL-03 tokens.fonts 全 deck 出现 4 个字族（Cascadia Mono、Georgia、Microsoft YaHei、SimHei），上限 3 个
P2 s01 DEN-01 页面 元素 16 个 > 上限 14 个
P2 s01 DEN-02 页面 文本块 8 个 > 上限 4 个（已排除 role=caption/source/pageNumber）
P2 s01 DEN-03 页面 bullet 8 条 > 上限 6 条
P2 s01 DEN-04 页面 字数 235 字 > 上限 220 字
P2 s01 OVF-03 body-longline 每行 37 字 > 上限 34 字/行（role=body）
P3 s01 OVF-04 caption-4lines 段落 3 行 > 上限 2 行（role=caption）
P3 s01 OVL-01 overlap-a/overlap-b 文本块重叠 2.40×0.40in（容差 0.05in）
P3 s02 DEN-05 页面 留白比例 19% < 下限 25%（内容覆盖 81%）
P0 s03 GEO-02 zero-size 尺寸非正：w=0.000in h=1.000in（必须 > 0）
SUMMARY P0=1 P1=0 P2=6 P3=3 issues=10 pages=3 exit=1
```

真实 deck 命中（3 页，深色版式，无 P0 可交付）：

```
P1 s01 OVF-01 主标题 文本溢出 0.04in（needed 2.29in > box 2.25in；2 行）
P1 s03 OVF-01 引导句 文本溢出 0.02in（needed 0.52in > box 0.50in；1 行）
P3 s03 OVL-01 页标题/引导句 文本块重叠 6.40×0.13in（容差 0.05in）
SUMMARY P0=0 P1=2 P2=0 P3=1 issues=3 pages=3 exit=0
```

两条 OVF-01 的框高恰好等于「1 行 × 行距 + 上内边距」，属零余量紧框：渲染图里没有裁切，但换渲染器或字体就会溢出，所以保留 P1 作为提示，不是错报。

**后续修正（终态）**：上面这 3 条已按框高算术修掉，`demo/smoke/deck.json` 现在跑出 `P0=0 P1=0 P2=0 P3=0 issues=0 exit=0`（见 `demo/smoke/qa-smoke.md`）。修法 = 主标题 `h 2.25→2.35`、引导句 `y 1.70 / h 0.55`、第三页标题与引导句间距拉到 0.08 in 以上。这份日志保留中间态，是为了让读者看到"零余量紧框"长什么样。

## 检查覆盖矩阵（30 项全覆盖，IMG-01 由 1.1.0 新增）

| ID | 由哪份输入触发 | 实测证据 |
| --- | --- | --- |
| GEO-01 | 主样例 s07 | 装饰圆右边界越画布 0.77 in，P0 |
| GEO-02 | 补充样例 s03 | `w=0.000in`，P0 |
| GEO-03 | 主样例 s07 | 眉标上边距 0.25 in < 0.45 in |
| OVL-01 | 补充样例 s01 / 真实 deck s03 | 重叠 2.40×0.40 in / 6.40×0.13 in |
| OVF-01 | 主样例 s03 / 真实 deck s01、s03 | 溢出 0.17 in / 0.04 in / 0.02 in |
| OVF-02 | 主样例 s04 | 溢出 157% > 30%，P0 |
| OVF-03 | 补充样例 s01 | 每行 37 字 > 34 字 |
| OVF-04 | 补充样例 s01 | caption 3 行 > 2 行 |
| TYP-01 | 主样例 s07 | 17 pt < h2 下限 24 pt |
| TYP-02 | 补充样例 | 标题 60 pt / 正文 20 pt = 3.00:1 |
| CON-01 | 主样例 s03 | 2.42:1（大字口径 3:1、正文口径 4.5:1 各报一条） |
| DEN-01 | 补充样例 s01 | 16 个元素 > 14 |
| DEN-02 | 主样例 s05 / 补充样例 s01 | 5 个、8 个文本块 > 4（已排除 caption/source/pageNumber） |
| DEN-03 | 补充样例 s01 | 8 条 bullet > 6 |
| DEN-04 | 补充样例 s01 | 235 字 > 220 |
| DEN-05 | 补充样例 s02 | 留白 19% < 25% |
| REP-01 | 主样例 s03、s04 | 与上一页签名相同 |
| REP-02 | 主样例 s02 | 同签名占 50% > 35% |
| REP-03 | 主样例 s04 | 焦点 ml 连续 3 页 |
| COL-01 | 主样例 s04 | 引用 `brand-blue` 未登记 |
| COL-02 | 主样例 s07 | 17 pt 不在 tokens.type 阶梯 |
| COL-03 | 补充样例 s01 | 4 个字族 > 3 |
| CNT-01 | 主样例 s05 | 图片缺 alt |
| CNT-02 | 主样例 s05 | 图表缺 unit 与 source |
| CNT-03 | 主样例 s07 | data 版式无页码 |
| CNT-04 | 主样例 s04 | 标题 26 字 > 25 |
| STR-01 | 主样例 s07 | 第 6 页 id 为 s07 |
| STR-02 | 负向用例 | `slides: []` 时载入阶段拦截，退出码 2，报告不生成问题行 |
| THR-01 | 负向用例 | 阈值副本缺 2 个键，每键 1 条 P1 |

## 负向用例（错误处理）

| 用例 | 命令要点 | 实测输出 | 退出码 |
| --- | --- | --- | --- |
| 页面数为 0 | deck 里 `slides: []` | `ERROR deck 至少有 1 页（当前 0 页），无法质检` + `SUMMARY ... exit=2` | 2 |
| JSON 语法错误 | 第 2 行少一个逗号 | `ERROR deck JSON 解析失败：第 2 行第 1 列：Expecting value` | 2 |
| 阈值键缺失 | `--thresholds` 指向删掉 `density.maxElementsPerSlide` 与 `density.textBlockExcludeRoles` 的副本 | 两个键各 1 条 `THR-01`（P1），其余检查按回退值继续 | 1 |
| 联系表：空目录 | `--dir` 指向没有图片的目录 | `ERROR 目录内没有 PNG/JPEG/WebP 预览图：…` | 2 |
| 联系表：目录不存在 | `--dir` 指向不存在的路径 | `ERROR 找不到预览目录：…` | 2 |

## 联系表实测

| 输入 | 输出 | 结果 |
| --- | --- | --- |
| `demo/smoke/previews/`（`smoke.pptx` 渲染的 3 页，1920×1081） | `demo/smoke/contact-sheet.png` | 3 列 × 1 行，单格 560×315，整图 1728×389，自动深底 |
| `demo/smoke/preview-smoke/`（重渲染的 3 页） | `demo/smoke/contact-sheet-smoke.png` | 3 列 × 1 行，整图 1728×389，自动深底 |
| `demo/smoke/contact-test/`（PIL 现造 6 张，明暗混合） | `demo/smoke/contact-test-sheet.png` | 3 列 × 2 行，整图 1728×716，自动浅底 |

三张图都做了肉眼检查：序号（01–06）与页名清晰可读，深浅格内都保持对比，格间距 12 px、格子等宽、无拉伸。深色页混排时，浅底会让页面边框与深色格形成明显反差，便于看节奏；反之深底适合全深色 deck。

## 误报修正记录（已落到代码）

1. **页码被判越出安全边距**：页码定位在页脚边距内属于版式约定，早期 5 页的页码全部误报。现在 `GEO-03` 对 `role: pageNumber` 直接豁免，背景与装饰层（`zRole: bg|bgDecor`）允许出血，不参与边距与留白计算。
2. **上下双边距重复计算导致 6 条 OVF-01**：`fit_report` 早期按「行高 + 上下内边距」估所需高度，把真实 deck 里 6 个正常文本块判成溢出 0.03–0.08 in。现改为只加上内边距（底线内边距只影响自动缩放与换行），改后只剩 2 条真实紧框溢出。
3. **阈值缺失重复报**：`THR-01` 早期同一键会按使用次数重复报，现按「键 → 回退说明」去重，每个缺失键只报 1 条。
4. **deckload 降级提示污染 stdout**：早期每次跑都打印「deckload 未能通过…」，现在只有加 `--verbose` 才打印，stdout 始终是「问题行 + SUMMARY」的固定结构。

## 已知限制

- **溢出估算量的是行盒高度，不是字形墨迹**：公式为「上内边距 + 行数 × 字号 × lineHeight × 1.204 / 72」英寸；1.204 是 `thresholds.type.lineBoxRatio`，由 10 种中英文字体渲染实测得出（偏差 < 1%）。单行误差约 ±0.03 in，框高等于行盒高时按零余量处理并报 P1。
- **富文本按段内最大字号估算**：同一段混排不同字号时结果偏保守，字号差越大误差越大。
- **字族由 `tokens.fonts` 与元素 `font` 声明决定**：无法感知渲染时的字体替换（缺字体时渲染器回退），字体缺失需看 `render.py` 的 `missingFonts`。
- **不检查图片文件是否真实存在**：`CNT-01` 只查 `alt` 文本；缺图由 `build_pptx.py` 在构建期报出并返回 1。
- **REP-02 少于 4 页不统计**：3 页里出现 2 次同版式即 67%，统计意义不足，只保留 REP-01 相邻页判断。
- **阈值口径以文件为准**：当前 `schema/thresholds.json` 里 `type.titleBodyRatioMin/Max` 为 1.3 / 3.0，脚本按文件执行；任务书早期给的区间是 1.4–2.2，如要恢复需改阈值文件（改脚本无效）。
- **`qa.json` 是给 CI 用的**：断言 `summary.P0 == 0` 即可判定是否阻断，不要依赖问题条数。
- **机器质检不等于画面验收**：构图、图表可读性、文字观感仍需看 `render.py` 的预览图与联系表。

## 复现命令

```powershell
$py = "<python.exe>"

& $py scripts/qa_check.py --deck demo/smoke/qa-fixture.json --out demo/smoke/qa-fixture-report.md --json demo/smoke/qa-fixture-report.json
& $py scripts/qa_check.py --deck demo/smoke/qa-fixture-edge.json --out demo/smoke/qa-fixture-edge-report.md --json demo/smoke/qa-fixture-edge-report.json
& $py scripts/qa_check.py --deck demo/smoke/deck.json --out demo/smoke/deck-qa-report.md --json demo/smoke/deck-qa-report.json

& $py scripts/render.py --pptx demo/smoke/smoke.pptx --out-dir demo/smoke/preview-smoke --pages all
& $py scripts/contact_sheet.py --dir demo/smoke/preview-smoke --out demo/smoke/contact-sheet-smoke.png --cols 3 --title "smoke.pptx 渲染总览"
& $py scripts/contact_sheet.py --dir demo/smoke/contact-test --out demo/smoke/contact-test-sheet.png --cols 3 --title "6 张自测图（明暗混合）"
```
