# 环境能力报告（第 0 步产出）

- 探测时间：2026-11-14
- 探测目标机：Windows 11 + DeepSeek Harness（DSH）桌面端
- 探测方式：逐项实跑，非推测。凡标注"已实测"的条目都有可复现命令与输出。

---

## 1. 结论摘要

| 能力 | 状态 | 对本系统的意义 |
|---|---|---|
| 文件读写 / 目录创建 / glob / grep | ✅ 已实测 | 知识库与 deck.json 可落盘 |
| Python 3.12.14 + python-pptx 1.0.2 | ✅ 已实测 | 可生成**原生可编辑** PPTX（形状、文本、图表、表格均非图片） |
| Pillow 12.3.0 | ✅ 已实测 | 字体度量（溢出估算）、图片裁切、联系表拼接 |
| PPTX → PNG 渲染 | ✅ 已实测 | 可以"生成→渲染→看图→修"的闭环，这是质量的地基 |
| Node v24.17.0 / pnpm | ✅ 已实测 | 可跑 `.mjs` 工具与脚本化流水线 |
| 中文图像识别（读渲染图） | ✅ 已实测 | 视觉质检可做，不只是结构检查 |
| Microsoft Office / WPS / PowerPoint COM | ❌ 不存在 | 不用 COM 路径；导出依赖上面第 4 项 |
| 独立安装的 LibreOffice（soffice 在 PATH） | ❌ 不存在 | 渲染走宿主提供的 office-cli，而不是系统 soffice |
| 平台专用 PPT 工件（Univer Slide / PPTD） | ✅ 可用 | 作为可选第二条产出路径，但不作为本 Skill 的主路径 |
| 联网检索 / 抓取 | ✅ 可用 | 仅用于外部资料核验，核心知识仍必须落本地 Markdown |

**关键判断**：本机具备"声明式版面 → 可编辑 PPTX → 真实渲染图 → 视觉复核"的完整闭环，因此本系统采用 **spec-first（先写 deck.json 再构建）** 而不是"让模型直接手写 PPTX"。这不是绕路，而是让版面可计算、可检查、可复现。

---

## 2. 逐项探测证据

### 2.1 Python 与库

```
> <bundled-python> -c "import sys, pptx; print(sys.version.split()[0], pptx.__version__)"
3.12.14 1.0.2
```

| 包 | 版本 | 用途 |
|---|---|---|
| python-pptx | 1.0.2 | 生成 PPTX（shape / text / chart / table / notes） |
| Pillow | 12.3.0 | 字体度量、图片裁切、联系表 |
| lxml | 6.1.3 | 直接写 PPTX XML（东亚字体 `a:ea`、字距 `spc`、透明度 `a:alpha`） |
| numpy | 2.3.5 | 图表数据预处理 |
| pandas | 3.0.1 | 数据表读取与统计 |
| python-docx | 1.2.0 | 用户提供 Word 素材时提取内容 |

Python 绝对路径：`<python.exe>`

### 2.2 PPTX 生成（已实测）

生成 2 页、含中文标题与 54 pt 粗体的 PPTX，重新打开校验：

```
saved probe.pptx
slides: 2 size: 12191695 6858000      # = 13.333 in × 7.5 in
```

结论：`slide_width/height` 可精确设为目标画布；中文经 `a:ea` 指定字体后可正确落字（见 2.3 渲染图）。

### 2.3 PPTX → PNG 渲染（已实测，本系统关键能力）

```
> & <node> "…\resources\office-cli.mjs" render --input probe.pptx --output-dir render1 --pages "1"
{"schemaVersion":1,"backend":"native","rasterEngine":"libreoffice",
 "pageCount":2,
 "images":[{"page":1,"path":"…\\page-0001.png","width":1920,"height":1081}],
 "missingFonts":[]}
```

- 输出 1920×1081（144 dpi，16:9 精确）。
- `missingFonts` 为空数组 → 本机字体可覆盖中文渲染。
- 渲染图目视检查：中文标题、粗体、位置、留白与 deck 定义一致。

**限制（必须记录）**：该渲染器由 DSH 宿主提供，等价于内置 LibreOffice 后端；在其他机器上可能不存在，`scripts/render.py` 必须做降级探测（系统 soffice → PowerPoint COM → 明确报错），不得假设可用。

### 2.4 备用预览路径（已实测）

```
> & msedge --headless --disable-gpu --hide-scrollbars --window-size=1280,720 \
    --screenshot=probe.png file:///…/probe.html
28161 bytes written to file probe.png
```

headless Edge 可把 HTML 渲成 1280×720 PNG，中文排版正常。**用途限制**：仅作为"没有 PPTX 渲染器时的近似预览"，其字体度量与 PowerPoint 不一致，不能替代 2.3 的真实渲染，报告中必须注明是近似预览。

### 2.5 字体清点（240 个字体文件）

| 类别 | 可用字体（本机实测存在） |
|---|---|
| 中文黑体 | Microsoft YaHei（msyh / msyhbd / msyhl）、Deng / Dengb / Dengl（等线）、SimHei、Noto Sans SC（VF + 三个静态字重） |
| 中文宋体/衬线 | SimSun、Source Han Serif SC Heavy、Noto Serif SC（VF）、BIZ UDMincho |
| 中文艺术 | SimKai（楷体）、SimFang（仿宋）、华文行楷 |
| 西文无衬线 | Lato（Thin→Black 全字重）、Bahnschrift（含可变字重）、Segoe UI 系列、Arial、Calibri、Verdana |
| 西文衬线 | Georgia、Cambria、Times、Palatino、Constantia |
| 等宽 | Cascadia Mono、Consolas、Cascadia Code |

设计含义：
- 可支撑"标题用重字重、正文用中字重"的层级做法（Lato Black、msyhbd、Bahnschrift 均可）。
- 可支撑衬线/无衬线对置的编辑式排版（Georgia + msyh）。
- **不能**假设小众开源字体（思源系列仅部分字重）在用户机器上存在 → tokens.fonts.fallback 必须写，且核心信息不得依赖特殊字体。

---

## 3. 平台专用能力（可用，但不作为主路径）

- `univer_*`：可创建 .univer Slide 单元、编译 SVG 页面、截图、导出 pptx。适合在 DSH 内做高层排版，但产物依赖本平台，违反"可移植"目标，仅列为可选路径。
- `pptd_*` / `ppt_*`：DSH 模板化 PPT 工作流，带模板资产与渲染。同上，仅作可选路径。
- `office-cli.mjs`：见 2.3。

**决策**：主路径固定为 `deck.json → python-pptx → PPTX → 渲染器 → PNG → 视觉复核`，因为它只依赖 Python 生态，Linux/macOS/Windows 都能复现；平台能力作为加速器记录在 `install.md` 的"平台差异"一节。

---

## 4. 能力缺口与对应处理（诚实记录）

| 缺口 | 影响 | 处理 |
|---|---|---|
| 无 PowerPoint/WPS COM | 无法用"另存为图片"做保真渲染 | 用 2.3 渲染器；探测失败则如实报告"视觉检查未完成"，只交结构检查结论 |
| 无系统 LibreOffice | 换一台无 DSH 的机器时渲染可能失败 | `scripts/render.py` 四级降级 + 明确安装指引；`qa_check.py` 的几何/溢出/对比度检查不依赖渲染器，仍可跑 |
| HTML 预览与 PPTX 字体度量不一致 | 近似预览可能"看着没溢出，实际溢出" | 只把 HTML 预览当兜底；布局结论以 `qa_check.py` 的字形度量与真实渲染为准 |
| 未验证 `jsonschema` 是否安装 | deck.json 校验可能退化 | `deckload.py` 做手写最小校验，schema 仅作文档契约 |
| 未做 pip 安装验证 | 不应引入新依赖 | 脚本只允许标准库 + python-pptx + Pillow |
| 无真实用户素材 | 案例与测试不得伪造 | 演示 deck 只用可公开的事实与自造示意数据，并标注"示意" |

---

## 5. 对实施计划的影响（阶段 0 结论）

1. **可以做**：全部文档层、deck.json 契约层、构建/渲染/质检脚本、真实渲染复核、3 类案例的真实产出。
2. **不能做、也不假装做**：PowerPoint 内的动画/母版级编辑、字体嵌入授权、素材版权取得、在线模板下载。
3. **必做加固**：渲染器探测与降级、字体回退链、无渲染器时的诚实结论输出。
4. **可选项**：Univer/PPTD 路径写进 `install.md` 的平台差异说明，不进主流程。

下一步按 §8 第 1 步进入需求探索模块（见 `docs/` 与 `discovery/`）。
