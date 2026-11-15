# 阶段 0 基准任务（固定输入，供复现与消融实验）

> 依据：`docs/lighthouse-plan.md` 阶段 0「固定一个可复现的 PPT 基准任务，保留完整需求、内容、约束和输入素材」。
> 本文件定义**三份不许随手改动的基准任务**。跑一次实验就换一次输入，等于没有基准——这是本文件存在的理由。

## 1. 三份基准任务总表

| 编号 | 类型 | 案例目录 | 页数 | 受众 | 场合 | 目标时长 | 图像来源 |
|---|---|---|---|---|---|---|---|
| **T1** | 学术 | `examples/cases/academic-defense/` | 13 | 答辩委员会（含非本方向评委） | 学位答辩 | 12 分钟 | 2 张示意图（PIL 生成，仓库内） |
| **T2** | 工程 | `examples/cases/engineering-review/` | 11 | 评审组（结构 + 声学 + 电源） | 设计评审 | 10 分钟 | 无外部图；图表为 PPT 原生 chart/table |
| **T3** | 创意 | `examples/cases/creative-portfolio/` | 9 | 客户方品牌负责人 | 方案提案 | 8 分钟 | 4 张自生成 JPG（仓库内） |

选这三份的理由：三种任务的信息结构、密度上限、视觉语言完全不同（学术偏证据与引用、工程偏参数与取舍、创意偏符号与冲击），能检出"只换色不改结构"的退化。

## 2. 每份任务的固定输入

### 2.1 任务原文（冻结，不得改写）

| 编号 | 任务原文 | 附加约束 |
|---|---|---|
| T1 | `examples/cases/academic-defense/brief.md` | 必须出现 3 条街道 × 42 测点 × 6 周的观测结构；结论必须带适用范围 |
| T2 | `examples/cases/engineering-review/brief.md` | 必须出现 3 个读数（结温 71.3 ℃ / 余量 6.8 K / 噪声 40.3 dB(A)）；必须给出三项取舍结论 |
| T3 | `examples/cases/creative-portfolio/brief.md` | 必须以"像一张能从墙上撕下来的海报"为视觉主张；不得使用发布会式大屏视觉 |

**冻结规则**：任务原文只允许两种改动——① 修正事实错误（须在 ADR 里记录改了什么、为什么）；② 新增一条独立任务 T4（不动原三份）。已发布过的实验结果不重算、不覆盖。

### 2.2 素材与工具

| 项 | 固定值 |
|---|---|
| 素材 | 全部在仓库内（`source/assets/`），不依赖网络；缺图时必须显式标注"未使用真实图片" |
| 模型 | 实验报告里必须写清模型名与版本号；不同模型的结果分表统计，不混算 |
| 工具 | `scripts/build_pptx.py`、`render.py`、`qa_check.py`、`copy_check.py`、`contact_sheet.py`，版本号取仓库 tag |
| 阈值 | `schema/thresholds.json`（`version` 字段必须与报告一致） |
| 时间预算 | 每份任务记录实际耗时（分钟），用于比较"修改成本"维度 |

## 3. 复现命令（PowerShell，四步）

```powershell
$PY   = "<python.exe>"                 # 例：<宿主>\office-runtime\...\python.exe
$NODE = "<HOST_ROOT>/.desktop-bin/node.cmd"
$CLI  = "<office-cli.mjs>"
$REPO = "<仓库根目录>"
Set-Location $REPO

# 以 T1 为例，T2/T3 换目录名即可
$CASE = "examples/cases/academic-defense"
& $PY scripts\build_pptx.py --deck "$CASE\source\deck.json" --out "$CASE\build\deck.pptx"
& $NODE $CLI render --input "$CASE\build\deck.pptx" --output-dir "$CASE\build\preview"
& $PY scripts\qa_check.py --deck "$CASE\source\deck.json" --out "$CASE\build\qa.md" --json "$CASE\build\qa.json"
& $PY scripts\copy_check.py --deck "$CASE\source\deck.json" --out "$CASE\build\copy.md" --json "$CASE\build\copy.json"
& $PY scripts\contact_sheet.py --dir "$CASE\build\preview" --out "$CASE\build\contact-sheet.png" --cols 4
```

**复现判据**：`qa.json` 的 `summary` 四项全 0；`copy.json` 的 P1 = 0；`contact-sheet.png` 生成成功；`deck.pptx` 体积与基线记录差异 <1%（deflate 波动所致，允许）。

## 4. 每次运行必须记录的字段

不记录这些，就无法区分"知识库起了作用"和"模型升级了 / 我手动改多了"。

| 字段 | 说明 |
|---|---|
| `run_id` | `<日期>-<任务号>-R<仓库版本>`，例 `2026-11-20-T1-R1.5.0` |
| `repo_version` / `git_commit` | 仓库 tag 与 commit |
| `model` | 模型名 + 版本号 + 温度等参数 |
| `condition` | `A`（不带知识库）或 `B`（按工作流调用知识库） |
| `prompt_hash` | 任务原文与初始提示的 sha256 前 12 位 |
| `input_files` | 使用的 brief / 素材清单 |
| `duration_min` | 从开工到初稿的耗时 |
| `manual_edits` | 人工修改轮数 + 改动条目数（粗略计数即可） |
| `qa_summary` | P0/P1/P2/P3 计数 |
| `copy_summary` | 文案机检 P0/P1 计数 |
| `artifacts` | deck.json / pptx / preview / qa.json / copy.json / contact-sheet 路径 |
| `notes` | 中途断点、模型异常、临时降级等 |

## 5. 阶段 0 内部通过标准

1. 三份任务都能用第 3 节命令从源文件重建，产物齐全。
2. 三份的 `qa.json` 全 0、`copy.json` P1 = 0。
3. 两份负向 fixture（`demo/smoke/qa-fixture.json`、`qa-fixture-edge.json`、`copy-fixture.json`）仍能报出预期问题——证明检查器没有被改坏。
4. 至少一份任务保留失败记录（哪一版不行、为什么不行、怎么改的），不许只留最终稿。
5. 三份结果放在 `docs/evaluation/results/` 下，命名遵守该目录的 README。

## 6. 变更控制

- 改任务原文、页数、约束、素材 → 必须写 ADR（`templates/adr.md`），并说明"为什么现在改、不改会怎样、已有结果如何迁移"。
- 基准任务的页数与受众变化会改变字号与密度判据，因此**旧结果不与新结果直接比较**；跨版本比较只在同一 R 编号内进行。
- 有人提交"更好的基准任务"时，按 T4 新增，不替换 T1–T3。

## 检查清单

- [ ] 三份任务的原文与约束都在仓库内，且没有网络依赖。
- [ ] 每份任务的复现命令都能四步跑通，`qa.json` 全 0。
- [ ] 每次运行都填满了第 4 节的 12 个字段。
- [ ] 报告里既有成功页，也有失败记录。
- [ ] 没有为了好看而改动任务原文或页数。
- [ ] 结果文件命名符合 `docs/evaluation/results/README.md`。
