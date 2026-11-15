# 验证报告：这套 skill 真的跑通了吗

写这份报告的目的只有一个：把"能跑通"和"我以为能跑通"分开。下面每一条都是本机（Windows 11 + DSH 桌面端，Python 3.12.14 / python-pptx 1.0.2 / LibreOffice 渲染后端）在 2026-11-14 实际执行得到的结果，附可复现命令。没跑过的东西单独列在最后一节，不混进结论里。

## 一、验证矩阵

| # | 任务形态 | 输入 | 关键结果 | 结论 |
| --- | --- | --- | --- | --- |
| V1 | 从零构建 8 页 deck | `demo/source/deck.json` | pptx 156,716 B / 8 页 / QA `0 0 0 0` | 通过 |
| V2 | 只换 tokens 的改版 | V1 的 deck.json 换成深色 tokens | build rc=0，QA `0 0 0 0`，渲染 2 页正确 | 通过（有 1 个资产路径坑，见 3.2） |
| V3 | 诊断故意做坏的 deck | `demo/smoke/qa-fixture.json`（6 页） | `P0=2 P1=9 P2=5 P3=3 issues=19`，rc=1 | 通过（缺陷全部命中） |
| V4 | 诊断边界用例 | `demo/smoke/qa-fixture-edge.json`（3 页） | `P0=1 P2=6 P3=3 issues=10`，rc=1 | 通过 |
| V5 | 元素覆盖冒烟 | `demo/smoke/deck.json`（3 页，11 种元素） | pptx 53,009 B，坐标反算误差 ≤ 0.02 in | 通过 |
| V6 | 契约负向测试 | 手写的非法 deck.json | 5 条 schema 错误全部报出 | 通过 |
| V7 | 三个完整案例（学术 / 工程 / 创意） | `examples/cases/` 三份 deck.json | 13 页 / 11 页 / 9 页，QA 均 `0 0 0 0`，三种完全不同的视觉体系 | 通过（泛化能力见第六节） |
| V8 | 真实会话复盘 → 规则补丁 | 一次真实 deck 改造会话（7 次用户介入） | 6 处缺口全部写成可判定规则；新增 IMG-01 并用两份 fixture 验证（33 ppi → P1、95 ppi → P2） | 通过（P2 级：只验证规则可执行，未重跑同题） |
| V15 | 文案机检（阶段 1.5） | demo / demo smoke / 三个案例 + 负向 fixture | 正例 5 份 `P1=0`；负向 fixture 命中 11 项 `P0=0 P1=5 P2=5 P3=2 exit=1` | 通过 |
| V16 | 本地偏好记忆（跨形态） | `scripts/pref_memory.py` + 临时目录 | 同意闸门拦住 2 次（enable/add 无 `--consented` 均 exit 1）；仓库内路径被拒（`.git` 检查）；禁止字段 `rawText` 被拒；add/list/export-md/forget 全链路 rc=0 | 通过 |
## 二、V1：从零构建（旗舰示例）

命令（四步，全部退出码 0）：

```bash
python3 scripts/build_pptx.py --deck demo/source/deck.json --out demo/build/deck.pptx
python3 scripts/render.py     --pptx demo/build/deck.pptx --out-dir demo/build/preview --dpi 144
python3 scripts/qa_check.py   --deck demo/source/deck.json --out demo/build/qa.md --json demo/build/qa.json
python3 scripts/contact_sheet.py --dir demo/build/preview --out demo/build/contact-sheet.png --cols 3
```

结果：`demo/build/qa.md` 的统计行为 `P0=0 条｜P1=0 条｜P2=0 条｜P3=0 条｜合计 0 条｜阻断=否`；8 张 1920×1081 PNG；字体族 2；缺失图片 0。

**关键在"从 20 条到 0 条"这 5 轮迭代**，它证明的是质检能被真正驱动，而不是一开始就恰好干净：

| 轮次 | QA 结果 | 触发问题 | 改法 |
| --- | --- | --- | --- |
| 1 | `P0=0 P1=7 P2=4 P3=9` | 标题正文比 2.59:1 越界、source 10 pt 低于 11 pt 下限、两页正文溢出、3 页字数超 220、16 pt 不在字号阶梯、与页码重叠 | 补 `bodySmall` 16 pt 阶梯、source 提到 11 pt、缩文案、把 8 行列表拆成两列 |
| 2 | `P0=0 P1=5 P2=1 P3=3` | 三处 source 需 0.28 in 但给了 0.25 in；第 5 页正文需 2.12 in 但给了 2.05 in；出处行越出下边距 0.06 in | 按 `行数 × 字号 × lineHeight × 1.204 / 72` 重算框高，逐个补足 |
| 3 | `P0=0 P1=0 P2=0 P3=1` | 页码与页脚说明重叠 0.10 × 0.23 in | 页脚文本框宽度从 11.6 in 收到 11.3 in |
| 4 | `P0=0 P1=0 P2=0 P3=0` | — | — |

这一轮里有一条教训值得写进文档：**框高不是审美问题，是算术问题**。第 2 轮的 5 条 P1 全部是"我目测留够了，但 `1.204 × lineHeight` 算下来不够"。`schema/thresholds.json` 把 1.204 定成唯一真值之后，这类错误从"渲染完看图才发现"变成"跑一遍 qa_check 就报出来"。

## 三、V2：只换 tokens 的改版

复制 V1 的 deck.json，只改 `tokens.colors` 六个值（`bg` `#0E1420`、`bgAlt` `#1B2536`、`fg` `#F2F4F8`、`fgMuted` `#9AA6B8`、`primary` `#8AB4F8`、`accent` `#F2B544`），版式、坐标、字号、文案一个字都没动：

- `build_pptx.py` rc=0，8 页，字体族仍为 2。
- `qa_check.py` 仍 `0 0 0 0`（对比度按新配色重新计算并通过，说明对比度检查是真的在算，不是在报缓存）。
- 渲染第 1、4 页，深底浅字、橙色强调换成琥珀色，版面结构完全一致。

### 3.2 这个测试暴露的一个坑（已写进 `SKILL.md` 失败处理表）

改了 deck.json 的**位置**后第一次构建报：

```
警告（2）：
  - $.slides[5].elements[2].src: 图片不存在：<新目录>\assets\page-04.png
```

原因：`src` 是相对路径，解析基准是 **deck.json 所在目录**，不是当前工作目录。把 deck.json 搬到别处必须连 `assets/` 一起搬。这是设计选择（deck 自带素材、可整体移动），不是 bug，但对"复制一份改改看"的使用者是真实的坑。

## 四、V3/V4：诊断能力（故意做坏的 deck）

一份只会说"没问题"的质检脚本等于没有质检。所以仓库里留了两份**故意做坏**的 fixture，它们的价值是"每次改动 qa_check.py 后跑一遍，看缺陷有没有漏报"：

- `demo/smoke/qa-fixture.json`（6 页，19 条问题）：命中 P0 元素出画布、P1 溢出 0.34 in、P1 相邻页版式签名重复、P1 图片缺 alt、P1 图表缺 unit/source、P1 slide id 不连续、P1 字号低于下限、P2 标题 26 字超限、P2 颜色未登记、P2 文本块 5 个超限、P3 焦点连续 3 页相同、P3 缺页码等。
- `demo/smoke/qa-fixture-edge.json`（3 页，10 条问题）：命中 bullet 8 条、字数 235、每行 37 字、caption 4 行、两框重叠 2.4 × 0.4 in、留白 19%、`w=0` 的零尺寸元素。
- 两者退出码均为 1（有 P0）。**退出码是给 CI 用的契约**：0 通过、1 有 P0、2 参数或依赖错误。

## 五、V5/V6：元素覆盖与契约拦截

- V5 `demo/smoke/deck.json`：3 页覆盖 text / runs / paragraphs / rect / ellipse / line / polygon / image（fit=cover）/ chart / table / notes 共 11 类。把渲染 PNG 反算回英寸，坐标误差 ≤ 0.02 in（色条 0→0.167 in、polygon 8,5→8,5、表头 8,2.05→8,2.056、icon 12.2→12.208）；`fit=cover` 后栅格横纵周期均为 21 px（理论 21.12），无拉伸。
- V6 手写非法 deck.json，被 `deckload` 拦下 5 条：`$.meta 缺少必填字段 size`、`$.meta.canvas 不允许多余字段`、`$.slides[0].id "c1"` 不符合 `^s[0-9]{2,3}$`、`$.slides[0].elements[i].id` 不允许多余字段（元素没有 id）等。**拦住我自己的那一次，比通过十次更能说明校验有效。**

## 六、V7：三个完整案例（学术答辩 / 工程评审 / 创意提案）

同一套 schema 换三种完全不同的视觉体系，验证 spec-first 不是只能做一种风格：

| 案例 | 页数 | 体系 | pptx | 字族 | QA | 关键手法 |
|---|---|---|---|---|---|---|
| `examples/cases/academic-defense/` | 13 | 深墨底 `#0E1620` + 衬线标题 + 单一金色 `#C8A24A` | 147,928 B | 3（思源宋体 / Noto Sans SC / Cascadia Mono） | 0/0/0/0 | 结论前置；两张 PIL 自绘示意图；出处走 caption |
| `examples/cases/engineering-review/` | 11 | 深蓝底 `#0B1420` + 等宽承担数据 | 78,273 B | 2（微软雅黑 / Cascadia Mono） | 0/0/0/0 | 2 个原生 chart（折线、横向条形，均带 unit+source）+ 2 个原生 table；11 页签名全不同 |
| `examples/cases/creative-portfolio/` | 9 | 纯黑 `#080808` + 单一朱红 `#FF3D2E` | 1,851,356 B | 3（思源宋体 / 微软雅黑 / Cascadia Mono） | 0/0/0/0 | 封面标题 96 pt 占页宽 79%；1 页满版图片主导（`zRole: bgDecor`）+ 1 页三图并置 |

三份都用同一条命令链跑出来：`build_pptx.py` → `render.py`（dsh-native，144 dpi）→ `qa_check.py` → `contact_sheet.py`，产物全部入库（`build/deck.pptx`、`build/preview/page-XXXX.png`、`build/qa.md`、`build/contact-sheet.png`），可直接对照 README 里的节奏表核数字。

透明的补充说明：这三个案例的 `source/deck.json` 在 2026-11-14 都经历过一次重建。`academic-defense` 用当时的确定性生成器重跑（重建后 pptx 与原 pptx 的 zip 成员 93 vs 93、差异 0 个，仅时间戳不同，判为等价）；`engineering-review` 与 `creative-portfolio` 的源文件在清理临时文件时被误删，改用「`build/deck.pptx` → 逆向导入 → 按 README 节奏表逐页修正版式签名与角色」的方式重建，重建后两者 QA 仍为 `0 0 0 0`，逐页目视复核与原始渲染一致（图表系列配色、数据标签、表格表头底色、卡片边框均已还原）。代价是这两份 deck.json 不保证与最初手写版本逐字节相同，元素顺序与内部命名可能不同——对使用者无影响，因为契约是 schema 而不是文件字节。

这一步暴露的真问题（都已修或已记录）：

- **深底上图表必须显式给 `color`**：默认 primary 深蓝在深底上几乎不可见（工程案例的折线图第一次渲染出来是空白的）。
- **KPI 数字与单位要拆 run**：`71.3` 与 `℃` 放同一段会被按最大字号（72 pt）度量，直接报 P1 溢出。
- **字族别名会触发 COL-03**：元素上的 `font` 必须写真实族名，写别名会被算成第 4、第 5 个字族。
- **原生 chart 的条宽与轴标题由脚本硬编码**（`build_pptx.py:847-848` 的 `gap_width=60` / `overlap=-8`），deck.json 改不了；需要更细控制时改用 rect 手绘（学术案例的第 7 页就是这么做的，代价是该页失去 CNT-02 的 unit/source 检查）。

## 七、V8：把一次真实使用里的"用户被迫引导"变成规则

2026-11-14，用本 skill 重做了一份真实 deck（4 轮迭代）。用户的 7 次介入没有一次是"skill 的规则被违反"，全部是"这类情况 skill 里没写规则"：

| 用户的介入 | 暴露的缺口 | 补到哪 |
|---|---|---|
| 封面图看不出主题；第 5 页多余 | 没有页面必要性判据 | `workflows/01_intake-and-content-audit.md` 页面必要性判据 |
| "第 7 页那张图改回去" | 没有已批准清单，改动外溢 | `workflows/07` §7.0、`discovery/requirement-summary.md` §7、`SKILL.md` 红线 13 |
| 生成图结构不成立（磁路在最前、无盆架） | 生成前无结构清单，出图后无放大核对 | `references/imagery-guide.md` §5.2、`workflows/05` 步骤 6b、`workflows/06` §6.5b |
| 折腾约 50 分钟才改用自备照片 | 素材获取没有失败闸门 | `imagery-guide.md` §5.1、`SKILL.md` 红线 11 与失败处理表 |
| 概念图"质量不够，再试试" | 质量不足没有判据，只会重复尝试 | 同上（降档判据含"质量不达标"） |
| "这个零件是不是畸形了"（2 倍放大才看得见） | 缩略图与机器质检都放过了畸形图 | `workflows/06` §6.5b 第 2 项 |
| 低分辨率图进了版面 | 没有像素密度检查 | `imagery-guide.md` §5.3、`thresholds.qa.minImagePpi`、`qa_check.py` 的 `IMG-01` |

其中只有最后一条能自动化，所以只有它进了 `qa_check.py`：`IMG-01` 用图片真实像素宽 ÷ 版面宽度算有效 ppi，低于 110 报 P2、低于 80 升 P1。两份临时 fixture 验证了方向：

| fixture | 图 | 版面 | 算得 | 结果 |
|---|---|---|---|---|
| 200×113 px | 33 ppi | 铺 6.00 in | 33 ppi | P1 `IMG-01`（并同时触发 CNT-03 缺页码） |
| 570×323 px | 95 ppi | 铺 6.00 in | 95 ppi | P2 `IMG-01` |

回归：`demo/`、`demo/smoke/`、三个案例在加 `IMG-01` 后仍全部 `0 0 0 0`（它们的图片都在 110 ppi 以上，没有误报）。

### 7.1 一次自我更正：闸门的位置放错了

1.1.0 的第一版把素材闸门写成「同一路径失败 3 次 → 降档」，也就是**事后补救**。用户当场指出重点不是"卡住之后怎么办"，而是"**动手之前该不该开始**"——正确顺序是先问「可不可以搜集 / 生成」，再问「有没有指定工具」，最后才动手。1.1.1 据此把它改成前置授权：

| 位置 | 1.1.0 的写法（事后） | 1.1.1 的写法（事前） |
|---|---|---|
| `imagery-guide.md` | §5.1 同路径失败 3 次才降档 | §5.0 第 0 关：未答复 = 不许；§5.1 退为第 2 关 |
| `workflows/00` | 无 | S0b 素材授权预判 + 完成标准 C6 + 门禁 G1 |
| `workflows/05` | 步骤 6b 三关（结构/放大/密度） | 四关，新增「0 授权关」 |
| `requirement-summary` | 无素材授权字段 | 「素材授权」5 字段 + 未许可默认路径 |
| `conversation-guide` | F1–F8 | 新增 F9「未获许可就动手找/生成素材」 |

第二次更正（同日）：1.1.1 里还写了「等许可时不许停工」，这同样是错的——**对话是回合制的**，用户必须等你结束回合才能回话，"边等边做"不存在。1.1.2 改为「**问完就交回合**」：一条回复内按「已做 / 要你答 / 不回则默认」三行交代清楚后结束本轮，默认取最保守档。这条比前一条更根本：它不是设计规则，而是**媒介约束**——任何"等你回复的同时我继续干"的指令在聊天界面里都不可执行。

教训本身值得记下来：**规则写在哪里，比写得多细更重要**——同一条规则放在"失败处理"里只能救火，放在"阶段 0 门禁"里才不点火。

**这条验证是 P2 级**：它证明规则能被流程和脚本执行，但不等于重跑同一个题目就一定不用用户引导。要升到 P1 需要拿同一份原始需求重做一次对照实验，且记录用户介入次数——本轮没做。

## 八、V15：文案机检（阶段 1.5）

- 正例五份（`demo/source/deck.json`、`demo/smoke/deck.json`、三个 `examples/cases/*/source/deck.json`）跑 `copy_check.py`，全部 `P1=0`；demo 有 1 条 P2（s05 的「至少 3 个方向」被归到无出处数字，实际出处是下一页），smoke 有 1 条 P2 + 1 条 P3，均不阻断。
- 负向 fixture `demo/smoke/copy-fixture.json`（3 页，故意写满空话与 AI 句式）：命中 CP-01 / 02×2 / 03×2 / 04 / 05 / 06 / 07 / 08 / 09 / 10 共 11 项，`P0=0 P1=5 P2=5 P3=2 issues=12 exit=1`。同一个文件跑 `qa_check.py` 报的是 9 条版面问题（`P0=3`）——说明两个检查器互不干扰，各查各的。
- 校准记录（避免以后调参把正例做成负例）：CP-01 对 `label/kpi/h2/h3` 只报 P2（这些角色常被借用为卡片标题）；CP-05 跳过 `source/pageNumber`、纯西文与命令行文本；CP-06 跳过列表型与编号清单；CP-07 只查「数字+单位」的量词、跳过序号、并降为 P2（出处常在相邻页）。
- 一个真 bug 与它的教训：`element_text()` 最初把 `paragraphs`、`runs`、`text` 三者拼在一起，导致同一段被数两遍、字数翻倍、CP-01 全页误报。现在按优先级只取一级。规则写对了、取文本的方式错了，结论会一样错。

## 八点五、V16：本地偏好记忆

目标：用户认可某一版之后，把"这个人喜欢什么"记在本机，下次从记录起步；硬约束是**先问、只存本机、不记事实**。

| 测什么 | 命令 | 结果 |
|---|---|---|
| 默认关闭 | `pref_memory.py --status` | `启用：否`，条目 0 |
| 未同意就启用 | `--enable`（不带 `--consented`） | 拒：提示先问用户，exit 1 |
| 同意后启用 | `--enable --consented` | rc=0，写入 `config.json`（含同意时间） |
| 未同意就写入 | `--add entry.json` | 拒，exit 1 |
| 同意后写入 | `--add entry.json --consented` | rc=0，生成 `entries.jsonl` + `preferences.md` |
| 条目里夹带原文 | 条目含 `rawText` | 拒：「只记偏好，不记原文与机密」 |
| 写进仓库 | `--dir <repo>/_memtest --enable --consented` | 拒：「记忆目录位于 git 仓库内」，目录未创建 |
| 查看 / 导出 / 删除 | `--list` / `--export-md` / `--forget <id>` | 全部 rc=0，删除后 `entries.jsonl` 与 `preferences.md` 同步 |

**没测的部分**：真实用户长期使用后的信噪比（记录会不会变噪音）、跨形态互用时的冲突（deck 偏好套到 A4 文档上是否合适）——这两项需要真实项目积累，现在无法造。

## 九、没验证的东西（不要当成已验证）

1. **阶段 0 的真实对话没跑过真人**。`docs/discovery-tests.md` 里 3 段对话是设计推演，不是真人录音转写。真实用户会在第 2 轮反问"你问我这个干嘛"，文档里的分支表覆盖了这种情况，但没有实测样本。
2. **只有本机一条渲染链路被真跑**。`render.py` 写了四级降级（自定义命令 / soffice / PowerPoint COM / DSH 后端），本机只跑通 DSH 后端；soffice 分支属未实测代码。
3. **只验证 16:9 + Microsoft YaHei**。4:3、16:10、3:2 的坐标要重算安全边距与栏宽；Noto Sans SC 的行盒比仍是 1.204（实测过），但中文断行位置不同，每行字数判定会变。
4. **长 deck 未验证**。最长只跑到 13 页（`examples/cases/academic-defense/`）、8 页（demo）、11 页与 9 页（另两个案例）。20 页以上 deck 的节奏检查（REP-02 占比、REP-03 焦点）只在 fixture 里做过单元级验证。
5. **图片密集 / 全图 deck 只验证了单页满版**。`examples/cases/creative-portfolio/` 有 1 页满版图片主导（`zRole: bgDecor`）与 1 页三图并置，但没有每页都是大图的 deck；满版图的裁切焦点是否落在主体上，仍只能人眼看。
6. **PDF 导出未纳入流水线**。`build/` 里没有 `deck.pdf`；需要时按 `workflows/07` 用外部工具转，本仓库不承诺。
7. **文案审核的人眼部分没验证**。脚本只能查可计算项（字数、空话、句式、重复、出处），「读起来像不像一个真人写的」没有判据，`copywriting-antislop.md` 只给了 12 个特征与改写动作。换模型审校那条路径（路径 C）也没有实测过：本机没有第二个模型的调用权，交接包格式只做了纸面设计。
8. **没做 A/B 审美评分**。机器质检只能证明"没有硬伤"，证明不了"好看"。`demo/README.md` 的"已知弱点"是我自己看出来的两处节奏问题，属于主观判断，不构成评分。

## 检查清单

- [ ] 每份 fixture 的退出码是 1，且问题条数与本文表格一致（改脚本后条数变化要同步更新本文）。
- [ ] `demo/build/qa.md` 统计行仍为 `P0=0 P1=0 P2=0 P3=0`。
- [ ] V2 的深色 tokens 只需改 6 个色值就能复现，若需要改版式说明 tokens 没有覆盖到位。
- [ ] 本文引用的每个数字都能在对应 `qa.md` / `render-manifest.json` / 构建摘要里找到。
- [ ] 新增验证项时，同步更新"没验证的东西"一节，删掉已覆盖项。
- [ ] V15 的负向 fixture 条数（11 项 / `P1=5 P2=5 P3=2`）与 `demo/smoke/copy-fixture-report.md` 一致；改 `copy_check.py` 后同步。
