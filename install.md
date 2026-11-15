# 安装与平台差异

本仓库是一个 **Agent Skill**：仓库根目录就是 skill 根目录，`SKILL.md` 是入口。安装 = 把仓库放到宿主的 skills 目录。

---

## 方式一：让 AI 自己装（推荐）

把下面这段发给你的 AI（仓库地址：https://github.com/DaCuem/ai-ppt-design-system）：

```
把这个仓库安装为我本地的 skill：https://github.com/DaCuem/ai-ppt-design-system

1) 用 git clone 或下载解压，放到你的 skills 目录；目录名保持 ai-ppt-design-system，根目录必须有 SKILL.md。
2) 如果你不确定 skills 目录在哪，按这个顺序找：$DSH_HOME/skills、~/.claude/skills、~/.codex/skills、~/.config/<agent>/skills；都找不到就问我。
3) 安装后读 SKILL.md 第 1 节，判断什么请求该启动它。
4) 运行 python scripts/probe.py，把能力矩阵贴给我；缺 Python 库或渲染器要说清楚，不要静默跳过。
5) 最后告诉我：装在哪、依赖是否齐全、能否生成并渲染一页测试稿。
```

AI 需要具备：文件读写、执行 shell。没有 shell 也能用，但第 2 步（渲染）与第 3 步（自动质检）会退化为纯文档流程。

---

## 方式二：手动安装

```bash
git clone <仓库链接> <技能目录>/ai-ppt-design-system
```

常见宿主的技能目录：

| 宿主 | 技能目录 | 备注 |
|---|---|---|
| DeepSeek Harness (DSH) | `%APPDATA%\dsh-desktop\harness\skills\`（Windows） | 每个 skill 一个子目录，含 `SKILL.md` |
| Claude Code | `~/.claude/skills/` | 可在项目内用 `.claude/skills/` |
| Codex CLI | `~/.codex/skills/` | 以本地文件为准 |
| Cursor / Windsurf | 项目内 `.cursor/skills/` 或按宿主文档 | 以"能被 agent 读到"为准 |
| 其他 agent | 任意目录 | 在会话里让 agent 打开 `SKILL.md` 即可 |

**目录名不影响功能**，但建议保持 `ai-ppt-design-system`，便于多设备一致。

---

## 方式三：不安装，直接引用

```text
请读 <仓库绝对路径>/SKILL.md，严格按其中的 8 阶段流程和我一起做这份 PPT。
需要知识库时按 SKILL.md 第 4 节的"文件调用规则"去读对应文件，不要一次性读全部。
```

适合临时使用、或在没有 skill 机制的模型里使用。

---

## 依赖

| 依赖 | 必需性 | 用途 | 安装 |
|---|---|---|---|
| Python ≥ 3.10 | 主路径必需 | 全部脚本 | 系统安装 |
| `python-pptx` ≥ 0.6.23 | 主路径必需 | 生成可编辑 PPTX | `pip install python-pptx` |
| `Pillow` | 主路径必需 | 字体度量（溢出估算）、图片裁切、联系表 | `pip install pillow` |
| 渲染器（任一） | 视觉检查必需 | PPTX → PNG | 见下表 |
| `jsonschema` | 可选 | 更严格的 deck.json 校验 | `pip install jsonschema` |

没有 Python 环境时：仍可完整执行需求探索、内容审计、方向选择、视觉系统、页面架构，并输出 `deck.json` 与逐页规格说明；生成 PPTX 与视觉检查需另行完成 —— **要在交付说明里写明这一点**。

---

## 渲染器

`scripts/render.py` 按下列顺序自动探测，找到即用：

| 优先级 | 渲染器 | 探测方式 | 说明 |
|---|---|---|---|
| 1 | `PPT_RENDER_CMD` 环境变量 | 模板串 `{input} {outdir} {pages}` | 自定义，最高优先 |
| 2 | 宿主自带渲染器 | 探测 `office-cli.mjs` 等已知路径 | 例如 DSH 桌面端自带 LibreOffice 后端 |
| 3 | LibreOffice | `soffice`/`soffice.exe` 在 PATH，或常见安装路径 | `--headless --convert-to pdf`，再转 PNG |
| 4 | PowerPoint COM | Windows + `pywin32` | 仅 Windows，且需已安装 PowerPoint |

全都探测不到时：脚本以退出码 2 结束，并打印可操作的安装指引。此时**只能做结构检查**，交付说明必须写"视觉检查未完成"。

### 各平台安装 LibreOffice

- Windows：`winget install TheDocumentFoundation.LibreOffice`
- macOS：`brew install --cask libreoffice`（路径 `/Applications/LibreOffice.app/Contents/MacOS/soffice`）
- Ubuntu/Debian：`sudo apt install libreoffice`

---

## 平台差异

| 平台 | 差异 | 处理 |
|---|---|---|
| Windows | 字体路径 `C:\Windows\Fonts`；无 `soffice` 时优先用宿主渲染器 | `metrics.py` 内置常见字体路径表 |
| macOS | 字体在 `/System/Library/Fonts`、`/Library/Fonts`；Office 常装 | 渲染器优先探测 `/Applications/LibreOffice.app` |
| Linux | 常有 `soffice` 但无中文字体 | 缺字体时 `missingFonts` 会显示；要用 `fonts-noto-cjk` |
| 无 shell 的 agent | 脚本不可执行 | 退化为纯文档流程，交付 `deck.json` + 规格说明 |

---

## 验证安装

```bash
python scripts/probe.py          # 应输出能力矩阵：Python/库版本、渲染器、字体
python scripts/build_pptx.py --deck demo/smoke/deck.json --out /tmp/smoke.pptx
python scripts/render.py --pptx /tmp/smoke.pptx --out-dir /tmp/preview
```

三步都成功 = 安装可用。第二步失败但第一步成功 = 只能产出 PPTX，无法视觉复核。

---

## 更新与卸载

```bash
cd <技能目录>/ai-ppt-design-system && git pull     # 更新
rm -rf <技能目录>/ai-ppt-design-system             # 卸载
```

卸载不影响你已产出的 deck：交付目录里的 `deck.json` + `scripts/` 拷贝即可继续重建。
