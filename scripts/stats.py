#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""stats.py — 仓库规模统计与「文档数字漂移」检查。

为什么有这个脚本：`selfcheck.py` 只查禁用词、占位符、链接、三段式、检查清单收尾、
JSON/Schema 合法性，**不校验文档里写的文件数与体积是否等于真实值**。每次发版都会有
一两处数字没跟上（v1.4.0 → v1.4.4 就漏了 `docs/project-summary.md` 的头部）。

用法：
    python scripts/stats.py                # 打印当前真实规模
    python scripts/stats.py --check        # 与 docs/project-summary.md 头部声明比对
    python scripts/stats.py --json         # 机器可读输出
    python scripts/stats.py --check --repo <仓库根>

退出码：0 = 一致 / 无声明；1 = 发现漂移；2 = 参数错误或读不到文件。
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

SKIP_DIRS = {"__pycache__", ".git", "dist", ".pytest_cache", ".venv", ".idea", ".vscode"}
SKIP_EXT = {".pyc", ".pyo", ".tmp", ".bak", ".DS_Store"}

# 声明示例：> 截至 **v1.4.5**（2026-11-15）：180 个文件 / 17.14 MB / 73 篇 Markdown / 14 个脚本 / 42 项机器检查
CLAIM_RE = re.compile(
    r"(?P<files>\d+)\s*个文件\s*/\s*(?P<mb>\d+(?:\.\d+)?)\s*MB\s*/\s*"
    r"(?P<md>\d+)\s*篇\s*Markdown\s*/\s*(?P<py>\d+)\s*个[^/\n]*脚本"
)


def collect(repo: Path) -> dict:
    files: list[tuple[str, int]] = []
    for path in repo.rglob("*"):
        if not path.is_file():
            continue
        parts = set(path.relative_to(repo).parts)
        if parts & SKIP_DIRS:
            continue
        if path.suffix.lower() in SKIP_EXT:
            continue
        files.append((str(path.relative_to(repo)), path.stat().st_size))

    total = sum(size for _, size in files)
    by_dir: dict[str, list[int]] = {}
    for rel, size in files:
        top = rel.split("\\")[0].split("/")[0]
        key = top if ("/" in rel or "\\" in rel) else "(root)"
        bucket = by_dir.setdefault(key, [0, 0])
        bucket[0] += 1
        bucket[1] += size

    return {
        "files": len(files),
        "bytes": total,
        "mb": round(total / 1048576, 2),
        "markdown": sum(1 for rel, _ in files if rel.endswith(".md")),
        "python": sum(1 for rel, _ in files if rel.endswith(".py")),
        "png": sum(1 for rel, _ in files if rel.endswith(".png")),
        "pptx": sum(1 for rel, _ in files if rel.endswith(".pptx")),
        "by_dir": {k: {"files": v[0], "kb": round(v[1] / 1024, 1)} for k, v in sorted(by_dir.items())},
    }


def read_claim(repo: Path) -> dict | None:
    doc = repo / "docs" / "project-summary.md"
    if not doc.is_file():
        return None
    head = "\n".join(doc.read_text(encoding="utf-8").splitlines()[:8])
    match = CLAIM_RE.search(head)
    if not match:
        return None
    return {
        "files": int(match.group("files")),
        "mb": float(match.group("mb")),
        "markdown": int(match.group("md")),
        "python": int(match.group("py")),
        "line": head.splitlines()[0] if head else "",
    }


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="仓库规模统计与文档数字漂移检查")
    parser.add_argument("--repo", default=str(Path(__file__).resolve().parents[1]))
    parser.add_argument("--check", action="store_true", help="与 docs/project-summary.md 头部声明比对")
    parser.add_argument("--json", action="store_true", help="输出机器可读 JSON")
    args = parser.parse_args(argv)

    repo = Path(args.repo).resolve()
    if not (repo / "SKILL.md").is_file():
        print(f"仓库根不正确（缺 SKILL.md）：{repo}", file=sys.stderr)
        return 2

    stats = collect(repo)
    claim = read_claim(repo) if args.check else None

    if args.json:
        print(json.dumps({"stats": stats, "claim": claim}, ensure_ascii=False, indent=2))
    else:
        print(f"文件总数：{stats['files']}")
        print(f"总体积：{stats['mb']} MB（{stats['bytes']} B）")
        print(f"Markdown：{stats['markdown']} 篇 / Python：{stats['python']} 个 / PNG：{stats['png']} 张 / PPTX：{stats['pptx']} 个")
        print("按目录：")
        for name, info in stats["by_dir"].items():
            print(f"  {name:<11s} {info['files']:>4d} 个文件 {info['kb']:>9.1f} KB")

    if not args.check:
        return 0

    if claim is None:
        print("\n未在 docs/project-summary.md 头部找到规模声明（跳过比对）。")
        return 0

    problems = []
    for key, label in (("files", "文件数"), ("mb", "体积 MB"), ("markdown", "Markdown 篇数"), ("python", "Python 脚本数")):
        if claim[key] != stats[key]:
            problems.append(f"{label}：文档写 {claim[key]}，实际 {stats[key]}")

    if problems:
        print("\n[FAIL] 数字漂移（改 docs/project-summary.md 头部）：")
        for item in problems:
            print(f"  - {item}")
        return 1

    print("\n[OK] 数字一致")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
