# 阶段 0 操作手册：方向证据

> 对应 `docs/lighthouse-plan.md` 的「阶段 0｜巩固 PPT 基线与验证设计知识的作用」。
> 与另两份文档的关系：`docs/v2-plan.md` 说做什么，`docs/v2-plan-ops.md` 说阶段 1–3 怎么做，**本文件只负责阶段 0**。
> 一句话目标：把"设计知识库有用"从信念变成可复核的证据；做不到就诚实地说"尚未证明"。

## 0. 完成定义（先看这一段）

阶段 0 完成 = 下面四件事同时成立：

1. 三份基准任务冻结（`docs/evaluation/baseline-tasks.md`），能用命令重建。
2. 至少一轮 A/B 消融实验跑完并落盘（`docs/evaluation/results/`），含盲评表。
3. 结论按协议第 6 节的分级写出，**没有越级**。
4. `docs/capability-status.md` 里"知识库有效性"这一行从"未开始"更新为实际结论等级。

不满足第 3 条，就算跑完了实验也不算完成——**跑出不利结果本身是合法结论**。

## 1. 交付物清单

| # | 文件 | 状态 |
|---|---|---|
| 1 | `docs/evaluation/baseline-tasks.md`（三份基准任务） | 已就位 |
| 2 | `docs/evaluation/ablation-protocol.md`（协议） | 已就位 |
| 3 | `docs/evaluation/results/<run>/`（每轮结果） | 待产生 |
| 4 | 更新后的 `docs/validation-report.md`（新增 V17：消融实验） | 待产生 |
| 5 | `docs/capability-status.md` 的结论等级 | 待更新 |

## 2. 前置检查（三条命令，任一失败先修环境）

```powershell
$PY   = "<python.exe>"                 # 例：<宿主>\office-runtime\...\python.exe
$NODE = "<HOST_ROOT>/.desktop-bin/node.cmd"
$CLI  = "<office-cli.mjs>"
$REPO = "<仓库根目录>"
Set-Location $REPO

& $PY scripts\selfcheck.py --strict      # 期望 P0 0 / P1 0 / P2 0
& $PY scripts\stats.py --check           # 期望 [OK] 数字一致
& $PY scripts\qa_check.py --deck examples\cases\academic-defense\source\deck.json --out $env:TEMP\t.md --json $env:TEMP\t.json
```

第三条期望 `P0=0 P1=0 P2=0 P3=0`。任一条不过，说明基线本身不干净，先修基线再做实验。

## 3. S1–S9：一轮完整实验

### S1 冻结基线

```powershell
git -C $REPO rev-parse --short HEAD      # 记 commit
git -C $REPO tag --points-at HEAD        # 记 tag（应为 v1.5.0 或更新）
& $PY scripts\stats.py --json            # 记文件数/体积，写进 run.json
```

产出：`run.json` 的 `repo_version` / `git_commit` / `stats`。

### S2 建结果目录

```powershell
$RUN = "docs\evaluation\results\2026-11-20-T1-R1.5.0"
New-Item -ItemType Directory -Force "$RUN\A","$RUN\B","$RUN\anon","$RUN\scores" | Out-Null
```

命名规则见 `docs/evaluation/results/README.md`；**同一任务同一版本只建一次**，重跑加 `-rerun` 后缀。

### S3 跑 A 组（对照：不提供知识库）

这是全套实验里最容易做错的一步。规程：

1. **开干净会话**：新会话，不加载本仓库的 `SKILL.md`、不读 `references/`。
2. 只给三样东西：① 基准任务的 `brief.md` 原文；② 可用的工具说明（`build_pptx.py` 的用法与 `deck.schema.json` 的字段表）；③ 输出要求（生成 `deck.json` 与 pptx）。
3. 禁止：不允许在提示里复述任何设计规则（字号下限、留白比例、版式清单等），否则 A 组被污染。
4. 记录：模型名与版本、会话 id、实际耗时、人工修改轮数。
5. 落盘：`A/deck.json`、`A/deck.pptx`、`A/preview/`。

**做不到干净会话时**：在 `run.json` 的 `notes` 里写明"无法隔离，A 组可能被知识污染"，并在结论里降级。**不许隐瞒。**

### S4 跑 B 组（实验：按工作流调用知识库）

1. 正常走 `SKILL.md` 的流程（阶段 0→1→1.5→2→3→4→5）。
2. 每一步记录用到的知识条目，写 `B/knowledge_refs.json`（格式见第 4 节）。
3. 落盘：`B/deck.json`、`B/deck.pptx`、`B/preview/`。

### S5 两组机器验收

```powershell
foreach ($g in @("A","B")) {
  & $PY scripts\qa_check.py  --deck "$RUN\$g\deck.json" --out "$RUN\$g\qa.md"  --json "$RUN\$g\qa.json"
  & $PY scripts\copy_check.py --deck "$RUN\$g\deck.json" --out "$RUN\$g\copy.md" --json "$RUN\$g\copy.json"
  & $NODE $CLI render --input "$RUN\$g\deck.pptx" --output-dir "$RUN\$g\preview"
}
```

记录 `qa_summary` 与 `copy_summary`。**这些数字只进"过程数据"，不进"质量结论"。**

### S6 生成匿名对照

1. 把两组的 contact sheet 复制成 `anon/P01.png`、`anon/P02.png`，顺序随机。
2. 把真实映射写进 `anon/mapping.json`：

```json
{ "P01": "B", "P02": "A", "seed": "2026-11-20T10:00", "tasks": ["T1"] }
```

3. 评分期间只看 `anon/`，不看 `A/`、`B/` 目录。

### S7 盲评

- 评审人数：≥3 人；达不到时用 1 人 + 两轮间隔 24 小时，并在报告披露。
- 每人一份 `scores/reviewer-N.md`，用协议第 5 节的 8 个维度 1–5 分，**每格必须写一句具体证据**。
- 评审者不知道 P01/P02 哪个是 A 组。
- 先看缩略图整体，再逐页看；顺序随机。

### S8 统计与结论

1. 每维度算均值与极差；极差 ≥3 的维度，结论里写明"评审分歧大"。
2. 按协议第 6 节确定能说到哪一级，写成一句话（例：*"在 T1 学术答辩任务、本机模型条件下，B 组在信息层级与反模式控制上均值高 0.7 分，其余维度差异不显著；单次运行，不足以说明稳定性。"*）
3. 写 `report.md`（协议第 7 节的十项，缺一不可）。

### S9 回写

| 回写到哪 | 写什么 |
|---|---|
| `references/anti-patterns.md` | 本轮暴露的失败模式（作为候选条目，标注来源 run_id） |
| `docs/capability-status.md` | 知识库有效性那一行的结论等级 |
| `docs/validation-report.md` | 新增 V17 行：消融实验（结论 + 目录路径） |
| `docs/project-summary.md` | 阶段 0 的进度与"最重要的未知数"是否已收敛 |

## 4. `knowledge_refs.json` 的最小格式

不需要日志系统，只要能让别人看出"这次决策读了哪几条规则"：

```json
{
  "run_id": "2026-11-20-T1-R1.5.0",
  "decisions": [
    {
      "page": "s03",
      "decision": "结论页用三列对照而不是单栏叙述",
      "knowledge_refs": ["references/layout-recipes.md#L05", "references/design-principles.md#contrast"],
      "reason": "三个方案要被横向比较，同位置对齐降低比较成本",
      "verification": "qa_check 无溢出；缩略图复核三列高度差 ≤0.2 in"
    }
  ]
}
```

规则：只记**关键决策**（方向、版式结构、色彩角色、字号层级），不记每一次微调，避免记录本身成为负担。

### 4.1 知识条目 ID 方案（本轮只定规则，不回填）

长期要给高价值规则稳定 ID（灯塔 §15 第三优先级）。格式：

```
KP-<文件短码>-<两位序号>     例：KP-LAY-05（layout-recipes 的 L05）、KP-TYPO-03
```

- 短码表：`DP` design-principles / `VS` visual-styles / `TYPO` typography / `COLOR` / `IMG` imagery / `CHART` / `RHYTHM` / `LAYOUT` / `AP` anti-patterns / `COPY` copywriting-patterns / `SLOP` copywriting-antislop / `A11Y` accessibility。
- ID 一经分配不改；规则内容大改时新建 ID 并在旧条目注明"被 KP-XXX 取代"。
- 回填时机：阶段 0 第一轮实验之后，只给"被实际引用过"的规则打 ID（先证明它被用了，再给它编号）。

## 5. 阶段 0 验收标准（A-0 系列）

| 编号 | 判据 | 证据 |
|---|---|---|
| A-0-1 | 三份基准任务冻结且能重建 | `docs/evaluation/baseline-tasks.md` + 三份 `build/qa.json` 全 0 |
| A-0-2 | 至少一轮 A/B 实验落盘 | `docs/evaluation/results/<run>/` 齐全（第 8 节目录结构） |
| A-0-3 | A 组隔离方式被记录（含做不到的披露） | `run.json` 的 `notes` |
| A-0-4 | 盲评表每格有证据 | `scores/reviewer-*.md` |
| A-0-5 | 结论等级未越级 | `report.md` + `capability-status.md` 一致 |
| A-0-6 | 失败案例被保留 | `report.md` 的失败章节 + `anti-patterns.md` 候选 |

## 6. 门禁 G0：什么时候可以开第二个媒介

同时满足才进阶段 1：

1. A-0-1…A-0-6 全过。
2. 结论至少达到协议第 6 节的第 3 级（"在指定任务与模型条件下，知识库带来可观察改善"）**或**明确写出"未观察到稳定改善 + 已定位问题环节"。
3. `docs/capability-status.md` 与 README 的定位表述一致（不说"已验证有效"除非到第 3 级）。

**G0 未过怎么办**：可以继续完善 PPT 分支（修知识、补案例、改工具），但**不得**把新媒介宣传为正式能力，也不得在 README 里写"支持多种媒介"。

## 7. 失败处理

| 情况 | 处理 |
|---|---|
| 无法开干净会话（A 组必被污染） | 记录限制，结论降为"无法隔离的观察"；先补"离线任务包"（把任务原文与工具说明导出成独立文件，在另一台机器/另一账号跑） |
| 凑不够评审者 | 降为单评 + 作者自查，标注"证据等级不足" |
| 两组差异只在措辞，不在结构 | 结论写"知识库改变了措辞，未改变结构"，这是重要发现，不是失败 |
| 机器指标 B 组更差 | 照实写；检查是否因为 B 组承担了更多信息，导致密度超限 |
| 实验跑歪（任务被改动） | 该轮作废并保留记录，不许改数据重算 |

## 检查清单

- [ ] 三份基准任务的原文与约束未改动，或改动有 ADR。
- [ ] 前置三条命令全绿后才开工。
- [ ] A 组确实未读 `references/` 与 `SKILL.md`（有记录，不是口头保证）。
- [ ] 两组的模型、参数、时间预算逐项一致，差异只允许在知识库。
- [ ] 每份任务至少 2 次运行，或明确标注"单次运行"。
- [ ] 盲评匿名，映射表与评分表分开存。
- [ ] 结论用词未超过协议第 6 节的等级。
- [ ] `capability-status.md`、README、`report.md` 三处口径一致。
- [ ] 失败案例与"尚不能证明的内容"都写进了报告。
