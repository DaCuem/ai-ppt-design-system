#!/usr/bin/env python3
"""仓库自检：文档规范、内部链接、占位符、JSON 契约。

检查对象是本仓库自身的知识库（.md）与契约（.json），不是用户的 deck.json
——后者由 qa_check.py 负责。

用法：
    python scripts/selfcheck.py                # 人类可读报告
    python scripts/selfcheck.py --json         # 机器可读
    python scripts/selfcheck.py --strict       # 存在 P1 也返回 1

退出码：0 = 无问题；1 = 存在 P0（或 --strict 下的 P1）；2 = 脚本自身无法运行。
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# 禁用词：出现即需改写（上下文行豁免见 ALLOW_MARKERS）
BANNED = [
    "高级感", "更有质感", "简洁大气", "呼吸感", "突出重点",
    "层次分明", "视觉冲击力", "科技感", "避免 AI 味", "避免AI味",
]
ALLOW_MARKERS = ("禁用", "必须改写为", "反例", "不合格", "不要写成", "避免使用")

TRIPLE = ("**触发条件**", "**具体动作**", "**常见错误**")

CHECKLIST_HEADING = re.compile(r"^##\s*检查清单\s*$", re.M)
LINK_RE = re.compile(r"(?<!!)\[([^\]]*)\]\(([^)\s]+)\)")
PLACEHOLDERS = ["TODO", "TBD", "FIXME", "待补充", "待写", "Lorem ipsum", "xxx"]

# 检查清单为必需的文件范围（仓库级说明文件除外）
CHECKLIST_REQUIRED_DIRS = ("discovery", "references", "workflows", "templates", "examples")

DOC_MAX_BYTES = 60 * 1024  # 单文件超此体积提示拆分


def split_code(text: str):
    """返回 (正文行列表, 代码块行列表)，用于跳过代码块内的误报。"""
    prose, code, in_fence = [], [], False
    for line in text.splitlines():
        if line.lstrip().startswith("```"):
            in_fence = not in_fence
            code.append(line)
            continue
        (code if in_fence else prose).append(line)
    return prose, code


def iter_docs():
    for path in sorted(ROOT.rglob("*.md")):
        rel = path.relative_to(ROOT)
        # demo/ 与各案例的 build/ 是构建产物（qa.md 由 qa_check.py 生成，每次重跑都会覆盖，
        # 不能要求它带 `## 检查清单`），dist/ 是打包目录，一律不参与写作规范检查。
        if any(part in {".git", "node_modules", "demo", "build", "dist"} for part in rel.parts):
            continue
        yield path, rel


def check_docs():
    issues = []

    def add(level, rel, line, msg):
        issues.append({"level": level, "file": str(rel), "line": line, "msg": msg})

    for path, rel in iter_docs():
        text = path.read_text(encoding="utf-8")
        prose, code = split_code(text)
        ignore_banned = "selfcheck-ignore-banned" in text

        # 1. 禁用词
        if not ignore_banned:
            for i, line in enumerate(text.splitlines(), 1):
                if any(m in line for m in ALLOW_MARKERS):
                    continue
                for word in BANNED:
                    if word in line:
                        add("P0", rel, i, f"禁用词「{word}」")

        # 2. 代码围栏配对
        if text.count("```") % 2:
            add("P0", rel, 0, "代码围栏 ``` 数量为奇数，围栏未闭合")

        # 3. 占位符（只看正文，代码块里的 xxx 可能是示例）
        #    行内豁免：文档要教人「搜 TBD」时，写 `<!-- allow-placeholder -->` 即可
        allow_placeholder = "allow-placeholder" in text
        for i, line in enumerate(prose, 1):
            if allow_placeholder and "allow-placeholder" in line:
                continue
            for ph in PLACEHOLDERS:
                if ph in line:
                    add("P1", rel, i, f"占位符「{ph}」未清理")

        # 4. 内部链接
        for i, line in enumerate(text.splitlines(), 1):
            for label, target in LINK_RE.findall(line):
                if re.match(r"^(https?:|mailto:|#)", target):
                    continue
                clean = target.split("#")[0]
                if not clean:
                    continue
                resolved = (path.parent / clean).resolve()
                if not resolved.exists():
                    add("P0", rel, i, f"内部链接失效：{label} → {target}")

        # 5. 三段式规则（参考与流程文件建议使用；条目化结构可豁免，降为提醒）
        if rel.parts and rel.parts[0] in {"references", "workflows"} and rel.name != "README.md":
            missing = [t for t in TRIPLE if t not in text]
            if missing:
                add("P2", rel, 0, "未使用三段式字段：" + "、".join(m.strip("*") for m in missing)
                    + "（条目化结构可忽略，但必须含可执行的转化动作）")

        # 6. 检查清单
        if rel.parts and rel.parts[0] in CHECKLIST_REQUIRED_DIRS:
            if not CHECKLIST_HEADING.search(text):
                add("P1", rel, 0, "缺少 `## 检查清单` 收尾")
            elif not CHECKLIST_HEADING.search("\n".join(text.splitlines()[-60:])):
                add("P2", rel, 0, "`## 检查清单` 不在文件末尾 60 行内")

        # 7. 体积
        size = path.stat().st_size
        if size > DOC_MAX_BYTES:
            add("P2", rel, 0, f"单文件 {size // 1024} KB，超过 {DOC_MAX_BYTES // 1024} KB 建议拆分")

        # 8. 字面 \n 残留（从旧文档复制常带来的问题）
        for i, line in enumerate(prose, 1):
            if "\\n" in line:
                add("P2", rel, i, "出现字面 `\\n`，检查是否为复制残留")

    return issues


def check_json():
    issues = []

    def add(level, rel, msg):
        issues.append({"level": level, "file": str(rel), "line": 0, "msg": msg})

    thresholds = ROOT / "schema" / "thresholds.json"
    schema = ROOT / "schema" / "deck.schema.json"

    for path, rel in ((thresholds, "schema/thresholds.json"), (schema, "schema/deck.schema.json")):
        if not path.exists():
            add("P0", rel, "文件不存在")
            continue
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            add("P0", rel, f"JSON 解析失败：{exc}")
            continue
        if rel.endswith("thresholds.json"):
            for key in ("canvas", "type", "density", "contrast",
                        "color", "repetition", "qa"):
                if key not in data:
                    add("P0", rel, f"缺少顶层键 `{key}`")
            canvas = data.get("canvas", {})
            for sub in ("safeMarginIn", "grid"):
                if sub not in canvas:
                    add("P1", rel, f"canvas 缺少 `{sub}`")
            density = data.get("density", {})
            if "textBlockExcludeRoles" not in density:
                add("P1", rel, "density 缺少 `textBlockExcludeRoles`")
        else:
            for key in ("$schema", "type", "properties"):
                if key not in data:
                    add("P1", rel, f"缺少顶层键 `{key}`")

    return issues


def structure_summary():
    rows = []
    for dirname in ("discovery", "references", "workflows", "templates", "examples", "scripts", "schema", "assets", "docs"):
        d = ROOT / dirname
        if not d.is_dir():
            rows.append((dirname, 0, 0))
            continue
        files = [
            p for p in d.rglob("*")
            if p.is_file() and "__pycache__" not in p.parts and p.suffix.lower() not in {".pyc", ".pyo"}
        ]
        rows.append((dirname, len(files), sum(p.stat().st_size for p in files)))
    return rows


def main() -> int:
    ap = argparse.ArgumentParser(description="仓库自检")
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    ap.add_argument("--strict", action="store_true", help="存在 P1 也返回非零")
    args = ap.parse_args()

    issues = check_docs() + check_json()
    order = {"P0": 0, "P1": 1, "P2": 2}
    issues.sort(key=lambda x: (order[x["level"]], x["file"], x["line"]))
    counts = {k: sum(1 for i in issues if i["level"] == k) for k in order}

    if args.json:
        print(json.dumps({"root": str(ROOT), "counts": counts,
                          "issues": issues, "structure": structure_summary()},
                         ensure_ascii=False, indent=2))
    else:
        print(f"仓库根：{ROOT}")
        print("\n结构：")
        for name, files, size in structure_summary():
            print(f"  {name:<12} {files:>3} 个文件  {size / 1024:>7.1f} KB")
        print(f"\n问题：P0 {counts['P0']}  P1 {counts['P1']}  P2 {counts['P2']}")
        if not issues:
            print("  未发现问题。")
        for issue in issues:
            loc = f"{issue['file']}:{issue['line']}" if issue["line"] else issue["file"]
            print(f"  [{issue['level']}] {loc}  {issue['msg']}")

    if counts["P0"]:
        return 1
    if args.strict and counts["P1"]:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
