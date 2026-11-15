---
name: ai-ppt-design-system
description: 通用 PPT / 演示文稿设计系统。当用户要做、改版或诊断一份 PPT（课程汇报、学术报告、项目或竞赛答辩、产品介绍、商业提案、求职作品集、工作汇报）时使用。先渐进式探索需求并出具需求确认书，再按内容选视觉方向、建立视觉系统、规划整份页面节奏，用 deck.json 声明式版面生成可编辑 PPTX，渲染成图后逐页视觉质检并迭代交付。不用于只写文案、只做图片或只处理表格的请求。
license: MIT
version: 1.4.5
---

# AI PPT 设计系统（主控制器）

本文件是**执行控制器**，不是设计教材。它规定：什么情况下启动、按什么顺序做、每步必须产出什么、读到哪个阈值、什么情况下**必须停下来**。

知识与方法在 `references/`，步骤细节在 `workflows/`，记录格式在 `templates/`，机器契约在 `schema/`，执行工具在 `scripts/`。

---

## 1. 启动条件

**启动**：用户要求制作、改版、扩写、诊断任意演示文稿；或提到汇报 / 答辩 / 路演 / 提案 / deck / slides / 讲稿页。
**不启动**：只要一段文案；只要一张图；只处理 Excel/Word 内容；只要模板外观描述。这些直接完成，不要套用本流程。

三种任务形态，流程不同：

| 形态 | 起点 | 跳过 |
|---|---|---|
| A 从零制作 | 阶段 0 | 无 |
| B 改版（有旧 PPT） | 阶段 1（把旧稿当素材审计）+ 阶段 0 的差异确认 | 完整需求探索，只确认"改什么、不改什么" |
| C 视觉诊断（不产出新稿） | 阶段 6 | 0–5 中除"读取现状"外的全部 |

---

## 2. 冲突裁决顺序（自上而下，前者覆盖后者）

1. **事实准确**：不虚构数据、经历、项目状态、图片来源、研究结论。
2. **用户明确要求**：用户指定的风格、颜色、页数、必须保留的内容。
3. **信息可读**：字号、对比度、行长、留白不得低于 `schema/thresholds.json` 的下限。
4. **叙事效率**：一页一个核心信息；不为凑版式牺牲逻辑。
5. **视觉表现**：风格、装饰、动效排最后。

用户要求与 3 冲突时，**指出冲突 + 给替代方案**，不要静默压到 3 以下，也不要静默改用户要求。

---

## 3. 强制工作流

八阶段，顺序不可跳。每阶段的输出物落成文件，后续阶段引用文件而不是记忆。

| 阶段 | 做什么 | 产出文件 | 通过门槛（不满足不许进下一步） |
|---|---|---|---|
| 0 需求探索 | 渐进式提问、视觉偏好探索、记录三类决策 | `requirement-summary.md` | 场景/受众/目的明确或已标注假设；素材与限制已知；用户确认或授权 |
| 1 内容审计 | 盘点素材、拆分论点证据、标注缺口 | `project-brief.md` | 每个数字/姓名/时间可追溯到输入；计划与已完成分开 |
| 1.5 文案生产 | 拆槽位、出逐页文案稿、**交用户审核**、按用户意思去 AI 味 | `copy-deck.md` | 每页一句话主张；槽位字数达标；空话与 AI 句式清零；**用户已确认**（未确认的记 `[假设]` 并在交付说明点名） |
| 2 方向选择 | ≥3 个实质不同的方向 + 六维评分 | `design-proposal.md` | 每个方向有构图/色彩/字体/图像的具体规则，而非形容词 |
| 3 视觉系统 | 色彩、字体、网格、图像四套规则落成 tokens | `deck.json` 的 `tokens` + 1 页样式样张 | 每个颜色有角色，每个字号有用途；对比度达标 |
| 4 页面架构 | 逐页核心信息、信息结构、版式编号、节奏 | `deck-outline.md` | 每页目的可一句话说清；版式签名无相邻重复；能量曲线有起伏 |
| 5 逐页制作 | 写 slide 元素 → build → render | `deck.json` + `deck.pptx` | 全部页可渲染；无 P0；元素均为可编辑对象 |
| 6 质量检查 | 结构自动检查 + 渲染图逐页视觉检查 | `qa-report.md` | P0 = 0，P1 有修复记录；修复后重新渲染 |
| 7 修订交付 | 版本化、打包、遗留问题记录 | `deliverables/` | 交付清单逐项核对通过 |

阶段细节见 `workflows/00_requirement-discovery.md` … `07_revision-and-handoff.md`；阶段 1.5 单独一份：`workflows/01b_copy-production.md`。

---

## 4. 文件调用规则（按需读取，不要全量加载）

| 当前任务 | 必读 | 按需读 |
|---|---|---|
| 开始任何任务 | 本文件 | `workflows/00` |
| 写文案 / 用户没给文案 / 嫌文案像 AI | `workflows/01b_copy-production.md`、`references/copywriting-patterns.md` | 去 AI 味读 `references/copywriting-antislop.md`（**只做判定与验收，修复走宿主 skill / 开源工具 / 换模型**）；载体是 `templates/copy-deck.md` |
| 需求还没问清 | `discovery/conversation-guide.md`、`discovery/requirement-question-bank.md` | `discovery/visual-preference-discovery.md` |
| 用户给了参考图/参考站 | `discovery/reference-analysis.md` | `references/visual-styles.md` 对应条目 |
| 选视觉方向 | `references/visual-styles.md`、`references/design-principles.md` | `references/color-guide.md` |
| 要找外部参考文献 | `references/further-reading.md`（先看 F 编号，再决定读哪条） | 断网时跳过本文件，规则全在本地 |
| 建视觉系统 | `references/color-guide.md`、`references/typography-guide.md` | `references/imagery-guide.md` |
| 规划整份页面 | `references/presentation-rhythm.md` | `references/layout-recipes.md` |
| 做单页 | `references/layout-recipes.md`、`references/typography-guide.md` | `references/charts-and-diagrams.md`、`references/imagery-guide.md` |
| 有图表/数据 | `references/charts-and-diagrams.md` | `references/color-guide.md`（图表色序） |
| 需要图片 / 生成图 | `references/imagery-guide.md` §5.0 先问授权与工具（第 0 关）、§1 角色、§5 降级、§5.2 生成四道关、§5.3 像素密度 | `assets/licenses.md`（许可记录） |
| 质检 / 诊断 | `workflows/06_quality-assurance.md`、`references/anti-patterns.md`、`references/accessibility-and-legibility.md` | — |
| 用户说满意 / 要不要记住偏好 | `references/preference-memory.md`（**先问再记，只存本机、不上传**） | `templates/preference-entry.md`、`scripts/pref_memory.py` |

**禁止**一次性读取整个 `references/`。每次只读当前阶段需要的最多 2 份，读完立刻产出，再读下一份。

---

## 5. 硬指标（数值真值见 `schema/thresholds.json`）

16:9 画布 13.333×7.5 in、观众 3–8 m 视距为基准：

| 项 | 规则 |
|---|---|
| 安全边距 | 左右 ≥ 0.6 in，上下 ≥ 0.45 in；任何元素不得越界（容差 0.02 in） |
| 字号下限 | h1 ≥ 36、h2 ≥ 24、h3 ≥ 20、正文 ≥ 16、注释 ≥ 11、页码 ≥ 9（pt） |
| 字号倍率 | 标题 : 正文 ∈ [1.4, 2.2] |
| 行长 | 中文正文 ≤ 34 字/行；标题 ≤ 16 字/行 |
| 行距 | 正文 1.25–1.55；标题 0.95–1.15 |
| 对比度 | 正文 ≥ 4.5:1；大字号（≥24 pt 或 ≥18.66 pt 粗体）≥ 3:1 |
| 密度 | 每页元素 ≤ 14、文本块 ≤ 4、bullet ≤ 6、总字数 ≤ 220、留白 ≥ 25% |
| 色彩 | 整个 deck 色相 ≤ 3；主/辅/强调面积 ≈ 60/30/10；强调色面积 3%–15% |
| 重复 | 版式签名不得相邻相同；同签名占比 ≤ 35% |
| 节奏 | 连续高密度页 ≤ 2，之后必须 1 页低密度 |

以上任何一条不达标都必须修，或者**在 `qa-report.md` 中写明为什么这条指标不适用于本稿**。不许沉默跳过。

---

## 6. 产出契约：deck.json + 六条命令

页面必须以 `deck.json`（契约见 `schema/deck.schema.json`）声明，而不是直接手写 PPTX 或整页贴图。原因：可计算 → 可检查 → 可复现。

```bash
# 0) 能力探测（首次在新机器上运行时）
python scripts/probe.py

# 1) 校验 + 构建可编辑 PPTX
python scripts/build_pptx.py --deck deck.json --out deck.pptx

# 2) 真实渲染为逐页 PNG（自动探测渲染后端）
python scripts/render.py --pptx deck.pptx --out-dir preview/ --dpi 144

# 3) 结构与版面自动检查（读 schema/thresholds.json 判级 P0–P3）
python scripts/qa_check.py --deck deck.json --out qa-report.md --json qa.json

# 4) 整份联系表，用于检查节奏与版式重复
python scripts/contact_sheet.py --dir preview/ --out contact-sheet.png

# 5) 文案机检（阶段 1.5 用；进制作前必须跑）
python scripts/copy_check.py --deck deck.json --out copy-report.md --json copy.json

# 6) 本地偏好记忆（**只在用户同意后**；默认目录 ~/.design-memory，永不上传）
python scripts/pref_memory.py --status          # 看是否启用、几条
python scripts/pref_memory.py --add entry.json --consented   # 写入一条（须先问用户）
python scripts/pref_memory.py --list            # 逐条查看
python scripts/pref_memory.py --forget mem-…    # 删单条
```

规则：
- 第 2 步失败（无渲染器）→ 明确说"视觉检查未完成"，只交结构检查结论，**不得**声称已做视觉质检。
- 第 3 步有 P0 → 必须修复并重跑 1–3 步，不得交付。
- 第 5 步（文案机检）有 P1 → 不许进阶段 2；脚本只查可计算项，**0 条不等于文案达标**，仍需按 `references/copywriting-antislop.md` 第 2 节人工过一遍。
- 所有文本、形状、图表、表格必须是可编辑对象；整页栅格化仅在用户明确要求"只要图片"时允许。
- 每轮修改都要递增 `meta.version`，并保留上一版 deck.json。
- 偏好记录（第 6 步）只能在用户明确同意后写；不写项目事实与客户信息；文件不得放进仓库或任何同步盘。

---

## 7. 红线（无条件禁止）

1. 跳过内容分析直接选模板/选配色。
2. 只给风格名称，不给构图、字号、色彩、图像的可执行规则。
3. 只规划单页，不检查整套叙事顺序与节奏。
4. 相邻页使用同一版式签名而无信息结构理由；用"统一"当作每页长一样的借口。
5. 把构成主义简化为红黑白+斜线+三角形，把瑞士风格简化为网格线+Helvetica（见 `references/anti-patterns.md` 与 `references/visual-styles.md` 的"容易被误解的表面符号"）。
6. 使用 `references/anti-patterns.md` 中 AP-01…AP-16 任一条：圆角卡片墙、四宫格、左右 5:5、图标堆叠、渐变玻璃、无意义装饰线、居中三件套、每页三段 bullet、3D 饼图、假数据、emoji 当图标等。
7. 虚构用户经历、数据、实验结果、项目状态、图片来源、用户评价。
8. 用"高级感""简洁大气""更有设计感"等无法执行的词下结论（见 `CONTRIBUTING.md` 禁用词表）。
9. 用户没要求就默认生成插图/配图，或把 AI 生成图当真实照片。
10. 覆盖或删除用户原始文件；交付物无版本号。
11. 同一条素材获取路径连续失败 3 次仍不降档、不告知（`references/imagery-guide.md` §5.1）。
12. 未做结构清单核对（零件 / 顺序 / 连接 / 朝向）就把 AI 生成图放进成片；生成图不做 2 倍放大复查。
13. 改动用户已认可的页 / 图 / 句（`records/approved.md`），或把一条局部意见扩大成全册改动。
14. 未拿到素材授权就动手：在用户答复「可否搜集 / 可否生成 / 用哪个工具」之前，就去联网搜图、截图、下载、装模型、跑生成（`references/imagery-guide.md` §5.0）。
15. 用户没给文案就直接开做页面，或文案未定稿（仍有 `[待确认]`）就进阶段 2；在排版时二次改写定稿文案（`workflows/01b_copy-production.md`）。
16. 交付物里有空话短语、AI 模板句式、无出处的数字；或为了让句子通顺改动用户的数字与结论方向（`references/copywriting-antislop.md` §2、§5）。
17. 把两件本该由用户拍板的事替用户拍了：**判定「像不像 AI」**与**宣布「改好了/已通过」**；或把「我自己润色了一遍」当成去 AI 味；或未获许可就联网检索、下载、安装外部工具（`copywriting-antislop.md` §1、§5，`discovery/conversation-guide.md` F11）。
18. 未经用户同意就写偏好记录，或把记忆文件放到仓库 / 网盘 / 任何联网位置；写入前没把条目原文给用户过目（`references/preference-memory.md` §1–§2）。

---

## 8. 项目目录约定

每个项目一个目录，文件按用途分开，不要把中间产物和交付物混在一起：

```
<project>/
├── records/                  # 阶段产出，按阶段编号
│   ├── requirement-summary.md
│   ├── project-brief.md
│   ├── copy-deck.md                 # 阶段 1.5：逐页文案稿（含审核状态）
│   ├── design-proposal.md
│   ├── deck-outline.md
│   ├── slide-spec.md
│   └── qa-report.md
├── source/                   # 唯一真值：版面源文件
│   ├── deck.json
│   └── assets/               # 图片等外部素材
├── build/                    # 可重跑的构建产物
│   ├── deck.pptx
│   ├── deck.pdf
│   ├── preview/              # 逐页 PNG，覆盖全部页
│   ├── contact-sheet.png     # 整份一览
│   ├── qa.json
│   ├── copy.json                   # 文案机检结果
│   └── render-manifest.json
└── deliver/                  # 交付副本，带版本号
    ├── <项目名>_v1.0.pptx
    ├── <项目名>_v1.0.pdf
    └── preview/ contact-sheet.png qa-report.md assets/asset-licenses.md
```

`build/` 随时可删可重建，`source/` 不可丢。交付前 9 项核对见 `workflows/07_revision-and-handoff.md`；**任何一项没做到就写在交付说明里**，不要省略。

---

## 9. 失败处理

| 情况 | 处理 |
|---|---|
| 读不到用户文件 | 报文件路径与错误；请用户换格式/重传；不要用"我猜内容"继续 |
| 没有 PPTX 生成能力 | 明确告知；给可执行替代（交 deck.json + HTML 预览 + 页面规格说明） |
| 没有渲染器 | 跑 `scripts/probe.py` 报告缺什么；只做结构检查并注明"视觉检查未完成" |
| 缺字体 | 走 `tokens.fonts.fallback`；不依赖特殊字体表达核心信息；记录替换情况 |
| 缺图片/数据 | 用结构图/示意图/留白替代并**标注为示意**；请用户补充；绝不伪造 |
| 内容冲突 | 列出冲突项与影响，请用户裁决；不自行选一个继续 |
| 工具报错 | 附完整错误；不要静默重试超过 2 次；不要改用未验证的替代方案 |
| 用户未答复素材授权 | 本回合内先做完不依赖素材的部分，然后按「已做 / 要你答 / 不回则默认」三行交回合结束；默认取最保守档（不搜集、不生成、几何图形 + 留白） |
| 素材生成 / 检索反复失败 | 同一路径 3 次即降档（真实素材 > 结构图 > 留白 ≥ 50% > AI 生成），并说明降档原因与影响页 |
| 用户说「改回去」 | 按最小改动原则只回退被指范围；把该元素写进 `records/approved.md`，后续不再动 |
| 用户嫌文案像 AI 写的 | **判定与验收都交回用户**：先按 §2 摆可数事实与位置，再按 §3 走 P1 宿主 skill → P2 联网装开源工具（须先问 Q37/Q38）→ P3 换模型审校；改完按 §5 问三问、等用户签字；不许自己判定"已经不像了" |
| 文案装不进版面 | 先删依据（保留结论 + 最有力的一条），不缩字号；删掉的记入 `copy-deck.md` 的待确认区 |

---

## 11. 本地偏好记忆（跨形态通用）

**触发条件**：用户明确表示满意（"就这样 / 可以了 / 不用再改"），或用户主动问"能不能记住我的偏好"。
**具体动作**：按 `references/preference-memory.md` §2 用三行交回合问一次 → 用户同意 → 把条目原文给他看 → `pref_memory.py --add entry.json --consented` 写入 `~/.design-memory/`。
**常见错误**：把"机检全 0"当用户满意；不问就记；记了项目事实或客户名；把记忆文件放进仓库跟着 commit。

| 规则 | 内容 |
|---|---|
| 范围 | 不限 PPT：deck / doc / pdf / image / sheet 用同一套条目格式（字段 `kind`） |
| 存哪 | 默认 `~/.design-memory/`；脚本拒绝任何 `.git` 仓库内路径 |
| 记什么 | 七类偏好：文案语气 / 信息密度 / 配色倾向 / 字体偏好 / 版式节奏 / 表达禁忌 / 交付方式 |
| 不记什么 | 项目事实与数据、客户名、凭据、他人素材原文、"他大概喜欢…"这类推断 |
| 下次怎么用 | 阶段 0 读一次，作为**默认选项并说明来源**；用户否决两次即降级或删除 |
| 上传 | **永不上传**。脚本无网络调用、无上传参数；放进同步盘同样算违规 |

## 10. 收尾自检（每次交付前）

- [ ] `requirement-summary.md` 存在，且"已确认/AI 建议/暂定假设"三类分开记录
- [ ] `deck-outline.md` 每页只有一个核心信息，且版式签名无相邻重复
- [ ] `records/copy-deck.md` 状态列全 `[已确认]`（或已记录默认档），且 `copy_check.py` 无 P1
- [ ] `qa_check.py` 输出的 P0 = 0
- [ ] 渲染图逐页看过，文字无溢出、无遮挡、无越界
- [ ] 每页能说出一个「观众问句」，且问句不重复（`workflows/01_intake-and-content-audit.md`）
- [ ] 生成图过了结构清单 + 2 倍放大两关，且标注「示意图（AI 生成）」
- [ ] `records/approved.md` 内元素本轮未被改动；改动只落在被指范围，其余页 diff 为空
- [ ] 联系表显示节奏有起伏，不是 20 页同一构图
- [ ] 所有数字、姓名、来源可追溯到输入
- [ ] 交付目录齐全，版本号已递增，用户原件未被覆盖
- [ ] 遗留问题已写入 `qa-report.md`
