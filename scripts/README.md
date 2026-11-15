# scripts/ 目录说明与用法总览

本目录是「AI PPT 设计系统」的可执行部分：把 `deck.json` 构建成 `.pptx`、渲染成预览图、按阈值做机器质检、再把预览拼成一张总览图。
所有脚本都遵守三条硬约束：

- 只用标准库 + `python-pptx` + `Pillow`（`render.py` 额外调用本机 node 渲染器，见第 4 节），不联网、不安装依赖；
- 错误信息为中文，Windows 控制台强制 UTF-8（`sys.stdout.reconfigure(encoding="utf-8")`）；
- 量化阈值只有一个真值来源：`schema/thresholds.json`（当前 v1.1.0）。脚本内不写死判据数字。
- `thresholds.version` 只随**阈值内容**变化；1.1.1 / 1.1.2 是规则与文档修订、没动任何数字，所以阈值仍是 1.1.0，qa.md 头部写 1.1.0 是正确的。

解释器统一用仓库指定的运行时（3.12.14，自带 python-pptx 1.0.2 与 Pillow 12.3.0）：

```powershell
$py = "<python.exe>"
```

## 目录内的脚本

| 脚本 | 职责 | 输入 | 主要输出 |
| --- | --- | --- | --- |
| `probe.py` | 探测本机能力（依赖、渲染后端、可用字体） | 无 | markdown / JSON 报告 |
| `build_pptx.py` | 由 `deck.json` 构建 `.pptx` | `deck.json` | `.pptx` + 构建摘要 |
| `render.py` | 把 `.pptx` 渲染成逐页 PNG | `.pptx` | `slide-NN.png` + `render-manifest.json` |
| `qa_check.py` | 按阈值质检 `deck.json`，出分级问题表 | `deck.json` | `qa-report.md` + `qa.json` + 控制台摘要 |
| `contact_sheet.py` | 把逐页 PNG 拼成一张总览图 | PNG 目录 | `contact-sheet.png` |
| `copy_check.py` | 文案机检：槽位字数、空话、AI 句式、句长、排比、数字出处、审校残留 | `deck.json`（可选 `copy-deck.md`） | `copy-report.md` + `copy.json` |
| `selfcheck.py` | 仓库自身文档与契约自检（不检查用户 deck） | 仓库文件 | 问题列表 |

退出码约定（四个可执行脚本一致）：

| 退出码 | 含义 |
| --- | --- |
| `0` | 成功；`qa_check.py` 表示没有 P0（可交付）；`copy_check.py` 表示没有 P0 且没有 P1 |
| `1` | 有阻断问题：`qa_check.py` 出现 P0，`copy_check.py` 出现 P1（可用 `--lenient` 降为只拦 P0），`build_pptx.py` 校验失败或缺图，`render.py` 输入/参数错误 |
| `2` | 输入不可用：文件读不到、JSON 非法、页面数为 0、目录里没有预览图、所有渲染后端都不可用 |

## 1. 一条命令跑完整链路

```powershell
& $py scripts/probe.py                                                     # 0. 先确认本机能力
& $py scripts/build_pptx.py --deck deck.json --out out/deck.pptx           # 1. 构建
& $py scripts/render.py --pptx out/deck.pptx --out-dir out/preview          # 2. 渲染预览
& $py scripts/qa_check.py --deck deck.json --out out/qa-report.md --json out/qa.json   # 3. 机器质检
& $py scripts/contact_sheet.py --dir out/preview --out out/contact-sheet.png --cols 3  # 4. 整体节奏检查
```

第 3 步不依赖第 2 步：`qa_check.py` 是纯结构 + 字体度量的离线性检查，渲染器不可用时也能跑完；渲染图用于人工看图复核构图与文字观感，两者互补。

## 2. probe.py —— 本机能力探测

```powershell
& $py scripts/probe.py                 # markdown
& $py scripts/probe.py --json          # JSON，便于脚本消费
& $py scripts/probe.py --no-fonts      # 跳过字体扫描（更快）
```

输出内容：Python 版本与解释器路径、`python-pptx`/`Pillow`/`lxml`/`jsonschema` 是否可用及版本、渲染后端探测结果（`PPT_RENDER_CMD` / `dsh-native` / `soffice` / `powerpoint-com`）、已安装字体的中文与西文分组。
退出码 `0` = 可构建（python-pptx 可用）；`1` = 缺少构建必需依赖。

换机器时先跑这个脚本，不要照抄另一台机器上「本机可用」的结论。

## 3. build_pptx.py —— deck.json 转 PPTX

```powershell
& $py scripts/build_pptx.py --deck deck.json --out out/deck.pptx
& $py scripts/build_pptx.py --deck deck.json --out out/deck.pptx --quiet   # 只打一行结果
& $py scripts/build_pptx.py --deck deck.json --check-only                  # 内存构建一遍再丢弃
```

参数：`--deck`（必填）输入 deck.json；`--out` 输出 .pptx；`--check-only` 只校验不写文件；`--quiet` 压缩摘要。

行为约定：16:9 = 13.333 in × 7.5 in = 1280 × 720 px @96 dpi（阈值文件 `canvas` 段为准）；元素越界、坐标非法、字号缺失等在构建期报错；缺图不中断构建，但记入摘要并把退出码置为 `1`，保证你仍能拿到一份可检查的 pptx。
退出码 `0` = 成功；`1` = deck 校验失败或资源缺失；`2` = 参数/文件/依赖问题。

## 4. render.py —— PPTX 渲染成 PNG

```powershell
& $py scripts/render.py --pptx out/deck.pptx --out-dir out/preview
& $py scripts/render.py --pptx out/deck.pptx --out-dir out/preview --pages "1,3" --dpi 96
& $py scripts/render.py --pptx out/deck.pptx --out-dir out/preview --pages all --json
```

参数：`--pptx`（必填）、`--out-dir`（默认 `preview/`，自动创建）、`--pages`（默认 `all`，可写 `"1,3"` 或区间）、`--dpi`（默认 96）、`--json`（把 manifest 打到 stdout）、`--timeout`（单次渲染超时秒数）。

后端按优先级探测，第一个可用的胜出：`PPT_RENDER_CMD` 自定义命令模板 → `dsh-native`（本机 DSH 自带 node + office-cli.mjs，LibreOffice 光栅化）→ `soffice`（转 PDF 再转 PNG）→ `powerpoint-com`（Windows PowerPoint COM，需要 pywin32 与已安装 PowerPoint）。
输出：`<out-dir>/slide-01.png`（编号 = 源页面号，左补零）与 `<out-dir>/render-manifest.json`（页数、尺寸、DPI、缺失字体、实际后端）。
退出码 `0` = 成功；`1` = 输入或参数有问题；`2` = 所有渲染后端都不可用或全部失败。

本机实测：`dsh-native` 可用，输出 1920 × 1081 px（约 144 px/in），返回 `backend=native`、`rasterEngine=libreoffice`。

## 5. qa_check.py —— 质检引擎

```powershell
& $py scripts/qa_check.py --deck deck.json --out qa-report.md --json qa.json
& $py scripts/qa_check.py --deck deck.json --thresholds schema/thresholds.json   # 显式指定阈值
& $py scripts/qa_check.py --deck deck.json --verbose                             # 打印降级诊断
```

参数：`--deck`（必填）、`--out`（Markdown 报告路径，省略则不写）、`--json`（机器可读 JSON，省略则不写）、`--thresholds`（默认相对脚本定位到 `../schema/thresholds.json`）、`--verbose`（打印 deckload 降级等诊断）。

三份输出：

1. **控制台**：每条问题一行 `P<级别> <slideId> <检查ID> <元素名> <中文描述（含实测值、阈值、单位）>`，末行固定 `SUMMARY P0=n P1=n P2=n P3=n issues=n pages=n exit=0|1|2`。
2. **Markdown**：头部写项目、生成时间、阈值版本；问题表固定 9 列 `编号 | 页码 | 类别 | 严重程度 | 检查项 | 问题描述 | 修复建议 | 状态 | 复查`；表后是 `**分级统计：…**` 与 `**结论：…**`，再加判定口径表。
3. **JSON**：`schemaVersion`、`deck`、`generatedAt`、`thresholdsVersion`、`canvas`、`summary`、`issues[]`；字段名英文、内容中文，便于 CI 断言 `summary.P0 == 0`。

退出码 `0` = 无 P0；`1` = 有 P0（按 `thresholds.qa.blockOnP0=true` 阻断）；`2` = 输入不可读（文件缺失、JSON 非法、`slides` 少于 1 页）。

### 检查项与级别（30 项）

| ID | 检查项 | 级别 | 判定口径（阈值键） |
| --- | --- | --- | --- |
| GEO-01 | 越出画布 | P0 | 元素盒越过画布边界 `canvas.16:9` |
| GEO-02 | 宽或高非正 | P0 | `w ≤ 0` 或 `h ≤ 0`（`kind: line` 豁免） |
| GEO-03 | 越出安全边距 | P2 | 超出 `canvas.safeMarginIn`（x 0.6 in / y 0.45 in，容差 0.02 in） |
| OVL-01 | 文本块重叠 | P3 | 两个文本块相交面积超过 `qa.overlapToleranceIn`（0.05 in） |
| OVF-01 | 文本溢出 | P1 | 估算所需高度超出框高 0.02 in 以上 |
| OVF-02 | 文本严重溢出 | P0 | 溢出比例超过 30%（键 `qa.overflowRatioBlockP0`，当前走回退值） |
| OVF-03 | 每行字数超限 | P2 | 最长行折算中文字数超过 `type.maxCjkCharsPerLine.*` |
| OVF-04 | 段行数超限 | P3 | 段行数超过 `type.maxLinesPerTextBlock.*` |
| TYP-01 | 字号低于下限 | P1 | 小于 `type.minPt.*`（按 role 映射） |
| TYP-02 | 标题正文比失衡 | P2 | 不在 `type.titleBodyRatioMin`–`type.titleBodyRatioMax` 区间（当前阈值文件为 1.3–3.0） |
| CON-01 | 对比度不足 | P1 | 小于 `contrast.bodyText`（4.5:1）；大字按 `contrast.largeText`（3:1） |
| DEN-01 | 元素数超限 | P2 | 超过 `density.maxElementsPerSlide`（14 个） |
| DEN-02 | 文本块超限 | P2 | 超过 `density.maxTextBlocksPerSlide`（4 个），已排除 `density.textBlockExcludeRoles` |
| DEN-03 | bullet 超限 | P2 | 超过 `density.maxBulletItemsPerSlide`（6 条） |
| DEN-04 | 字数超限 | P2 | 超过 `density.maxCharsPerSlide`（220 字） |
| DEN-05 | 留白比例不足 | P3 | 低于 `density.minWhitespaceRatio`（25%），用 0.1 in 栅格估并集覆盖率 |
| REP-01 | 相邻页版式重复 | P1 | 相邻两页 `repetition.signatureFields` 签名相同 |
| REP-02 | 版式占比过高 | P2 | 同一签名占比超过 `repetition.maxSameLayoutShare`（35%），少于 4 页不统计 |
| REP-03 | 焦点位置连续重复 | P3 | `focalPosition` 连续 3 页相同（`none`/`full` 不计） |
| COL-01 | 颜色非法或未登记 | P2 | 颜色既不在 `tokens.colors`，也不是合法 `#RRGGBB` |
| COL-02 | 字号不在阶梯 | P3 | 字号不在 `tokens.type.*.size` 阶梯内 |
| COL-03 | 字体族过多 | P2 | 全 deck 字族超过 3 个（键 `color.maxFontFamilies`，当前走回退值） |
| CNT-01 | 图片缺替代文本 | P1 | `kind: image` 缺少非空 `alt` |
| CNT-02 | 图表缺单位或来源 | P1 | `kind: chart` 缺 `unit` 或 `source`（受 `charts.requireUnitAndSource` 开关控制） |
| CNT-03 | 缺页码 | P3 | 页面没有 `role: pageNumber` 元素 |
| CNT-04 | 标题过长 | P2 | 标题超过 25 字（键 `type.maxTitleChars`，当前走回退值） |
| IMG-01 | 图片有效像素密度不足 | P2（<80 ppi 升级 P1） | 图片像素宽 ÷ 版面 w（in）小于 `qa.minImagePpi`（110）；低于 `qa.minImagePpiBlock`（80）直接 P1。全宽 12.13 in 需 ≥1335 px |
| STR-01 | slide id 不连续 | P1 | id 重复或与页序不对应（`^s[0-9]{2,3}$`） |
| STR-02 | 页面数为零 | P0 | 由载入阶段拦截并以退出码 2 结束，正常报告里不会出现该行 |
| THR-01 | 阈值键缺失 | P1 | `thresholds.json` 缺键；报告点名缺失的键并按回退值继续检查 |

标注「当前走回退值」的 3 个键（`qa.overflowRatioBlockP0` 30%、`type.maxTitleChars` 25 字、`color.maxFontFamilies` 3 个）在阈值文件 v1.0.0 中尚未定义，脚本按任务书给定值继续检查，并在问题描述里写明回退值；`thresholds.json` 补上同名键后自动生效，不需要改代码。

## 6. contact_sheet.py —— 预览总览图

```powershell
& $py scripts/contact_sheet.py --dir out/preview --out out/contact-sheet.png
& $py scripts/contact_sheet.py --dir out/preview --out out/contact-sheet.png --cols 3 --width 560
& $py scripts/contact_sheet.py --dir out/preview --out out/sheet.png --cols 4 --title "季度汇报 v3"
```

参数：`--dir`（必填，PNG/JPEG/WebP）、`--out`（默认 `contact-sheet.png`）、`--cols`（默认 3）、`--width`（单格宽度像素，默认 560）、`--gap`（格间距像素，默认 12）、`--title`（顶部标题）、`--recursive`（递归子目录）、`--dark` / `--light`（强制底色，默认按预览图平均亮度自动判断）、`--no-labels`（不画序号与页名）。

每个格子：统一缩放到同一宽度、等比缩放、格间距 12 px，左上角覆盖「序号 + 页名」标签块；整张图的底色按全部预览图的平均亮度自动选深底或浅底（深色 deck 用深底，避免大面积白边）。

用途：一眼看完整套页面的节奏——连续三页同一种版式、章节页有没有视觉断点、页脚与页码是否每页都在。三张页面以上的 deck 都建议出一张。

退出码 `0` = 成功；`2` = 目录不存在、目录里没有预览图、参数非法。

## 7. selfcheck.py —— 仓库自检

```powershell
& $py scripts/selfcheck.py            # 人类可读
& $py scripts/selfcheck.py --json     # 机器可读
& $py scripts/selfcheck.py --strict   # 存在 P1 也返回 1
```

检查对象是本仓库的 `.md` 文档与 `.json` 契约：内部链接是否失效、有没有违规禁用词、有没有未替换的模板标记、体积是否超标、`## 人工判定项（脚本不实现）

以下三条依赖语义判断，`qa_check.py` 不做检查，必须由人对着渲染图判，并把结论写进 `records/qa-report.md`：

| 项 | 阈值来源 | 为什么机器判不了 | 人工怎么判 |
|---|---|---|---|
| 节奏密度 | `thresholds.rhythm.maxConsecutiveDenseSlides` = 2 | "密"要先算出可见字数与留白率，再判断是否构成压迫感，且同一份 deck 的"密"标准随场合变化 | 逐页算可见字数与留白率，标记 `密 >= 180 字且留白 < 50%`，数连续段 |
| 换气间隔 | `thresholds.rhythm.minVisualBreakEvery` = 4 | "换气页"是叙事概念，不是版式签名 | 标出每页承担的功能，检查每 4 页内是否有一次低密度页 |
| 艺术质量 | 无阈值（刻意不设） | 构图、文字张力、叙事说服力没有可计算判据 | 逐页看渲染图，按 `references/design-principles.md` 的 10 条原则口述理由 |

其余 30 项都是可计算的，`qa_check.py` 全实现（见上表）。**0 issue ≠ 艺术验收通过**：机器过的只是结构与度量。

## 检查清单` 是否收尾等。它与 `qa_check.py` 分工不同：前者管「仓库自身写得对不对」，后者管「用户的 deck 能不能交付」。退出码 `0` = 无问题；`1` = 存在 P0（或 `--strict` 下的 P1）；`2` = 脚本自身无法运行。

## 8. 已知限制

- **溢出估算的是行盒高度，不是字形墨迹**：`needed = 上内边距 + 行数 × 字号 × lineHeight × lineBoxRatio / 72`，`lineBoxRatio` 读 `thresholds.type.lineBoxRatio`（1.204，10 种中英文字体渲染实测，偏差 < 1%）；改这个键，质检读数跟着变。单行误差约 ±0.03 in；框高恰好等于行盒高时（零余量设计）会被判 P1，这是提示「渲染器差异 ±1–2% 就可能溢出」，不是错报。
- **富文本按段落最大字号估算**：同一段内混排不同字号（`runs` 内 size 不一）时按最大字号算所需高度，会比实际略高，误差随字号差增大。
- **字体缺失时按回退链度量**：`Microsoft YaHei → Noto Sans SC → DengXian → SimHei → Georgia → Arial`，中文字符始终用中文字体度量（部分西文字体对中文返回 0.53 em 的无效字形）。
- **不检查图片文件是否存在**：`CNT-01` 只查 `alt` 文本；缺图由 `build_pptx.py` 在构建期报出并以退出码 1 结束。
- **REP-02 少于 4 页不统计**：3 页里出现 2 次同版式就是 67%，统计意义不足，只保留相邻页判断（REP-01）。
- **页码豁免安全边距**：`role: pageNumber` 元素允许落在页脚边距内，属版式约定；背景/装饰层（`zRole: bg|bgDecor`）允许出血，不参与 GEO-03 与留白计算。
- **渲染与质检相互独立**：`qa_check.py` 不调用渲染器，所以「机器质检通过」不等于「画面一定对」；文字观感、图表可读性仍需看 `render.py` 的 PNG。
- **阈值口径以文件为准**：文档与脚本出现分歧时，`schema/thresholds.json` 是唯一真值来源，改阈值必须先改文件。

## 检查清单

- [ ] `python scripts/probe.py` 退出码为 0，输出里 `python-pptx` 为可用。
- [ ] `python scripts/build_pptx.py --deck deck.json --out out/deck.pptx` 退出码为 0，`out/deck.pptx` 存在且能被 `python-pptx` 打开。
- [ ] `python scripts/render.py --pptx out/deck.pptx --out-dir out/preview` 退出码为 0，PNG 数量等于页数。
- [ ] `python scripts/qa_check.py --deck deck.json --out out/qa-report.md --json out/qa.json` 退出码为 0（有 P0 时为 1，需先修再交付）。
- [ ] `out/qa.json` 的 `summary.P0` 为 0，`summary.pages` 与 deck 页数一致。
- [ ] `python scripts/contact_sheet.py --dir out/preview --out out/contact-sheet.png` 退出码为 0，总览图里每格序号与页名可读。
- [ ] 报告里的每条 P0/P1 都能对应到具体 `slideId` 与 `check ID`，没有「页面」这种无法定位的记录。
- [ ] 阈值改动只发生在 `schema/thresholds.json`，脚本与文档里没有另写一份数字。
