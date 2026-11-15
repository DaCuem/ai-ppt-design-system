#!/usr/bin/env python3
"""本地偏好记忆（Local Preference Memory）——只写本机，永不联网。

设计前提（与 references/preference-memory.md 一致）：
1. 默认关闭。启用、写入、导出，三步都要用户明确同意（--consented）。
2. 存储位置在本机用户目录，**不允许落在任何 git 仓库内**（脚本会向上查找 .git 并拒绝）。
3. 不记录项目事实内容、不记录机密、不记录他人素材原文；只记偏好与可复用判断。
4. 本脚本不 import 任何网络库，也不提供任何上传参数。

用法：
  python pref_memory.py --status
  python pref_memory.py --enable --consented
  python pref_memory.py --disable
  python pref_memory.py --add entry.json --consented
  python pref_memory.py --list [--json]
  python pref_memory.py --export-md [--out PREFERENCES.md]
  python pref_memory.py --forget <id>
可选：--dir <记忆目录>（默认 ~/.design-memory）

退出码：0 成功 / 1 被拒绝（未获同意、路径不安全、条目不合法）/ 2 参数错误
"""

import argparse
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

DEFAULT_DIR = Path.home() / ".design-memory"
CONFIG_NAME = "config.json"
ENTRIES_NAME = "entries.jsonl"
INDEX_NAME = "preferences.md"

REQUIRED_FIELDS = ("project", "kind", "satisfiedVersion", "preferences")
ALLOWED_KINDS = ("deck", "doc", "pdf", "image", "sheet", "other")
ALLOWED_SOURCE = ("user-said", "observed")
FORBIDDEN_KEYS = {
    "rawText", "fullText", "copiedContent", "credentials", "token", "password",
    "apiKey", "secret", "customerName", "personalData",
}


def now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def fail(msg: str, code: int = 1) -> "NoReturn":
    print(f"[拒绝] {msg}")
    sys.exit(code)


def resolve_dir(raw: str | None) -> Path:
    d = Path(raw).expanduser() if raw else DEFAULT_DIR
    return d.resolve()


def assert_not_in_repo(target: Path) -> None:
    """从目标目录向上找 .git；命中则拒绝（防止偏好被误提交到仓库）。"""
    for parent in [target, *target.parents]:
        if (parent / ".git").exists():
            fail(f"记忆目录位于 git 仓库内（{parent}），请换到仓库之外，例如 ~/.design-memory")


def load_config(d: Path) -> dict:
    f = d / CONFIG_NAME
    if not f.is_file():
        return {"enabled": False, "consentedAt": None, "defaultDir": str(d), "note": "默认关闭，启用需用户同意"}
    try:
        return json.loads(f.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        fail(f"配置文件损坏：{f}")


def save_config(d: Path, cfg: dict) -> None:
    d.mkdir(parents=True, exist_ok=True)
    (d / CONFIG_NAME).write_text(json.dumps(cfg, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def require_consent(args, action: str) -> None:
    if not args.consented:
        fail(
            f"{action}需要用户明确同意。先把要记录的内容原样给用户看，问一句"
            "「要不要把这次的偏好记到本机？只存本机、可以随时删」，得到同意后加 --consented 重跑。",
            code=1,
        )


def load_entries(d: Path) -> list:
    f = d / ENTRIES_NAME
    if not f.is_file():
        return []
    out = []
    for line in f.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            out.append(json.loads(line))
        except json.JSONDecodeError:
            fail(f"条目损坏，请手工检查 {f}")
    return out


def write_entries(d: Path, entries: list) -> None:
    d.mkdir(parents=True, exist_ok=True)
    with (d / ENTRIES_NAME).open("w", encoding="utf-8", newline="\n") as fh:
        for e in entries:
            fh.write(json.dumps(e, ensure_ascii=False) + "\n")


def validate_entry(e: dict) -> dict:
    for k in REQUIRED_FIELDS:
        if k not in e:
            fail(f"条目缺少必需字段：{k}")
    if e["kind"] not in ALLOWED_KINDS:
        fail(f"kind 只能是 {ALLOWED_KINDS}")
    if not isinstance(e["preferences"], list) or not e["preferences"]:
        fail("preferences 必须是非空数组")
    for p in e["preferences"]:
        if not isinstance(p, dict) or "area" not in p or "statement" not in p:
            fail("每条 preference 需含 area 与 statement")
        if p.get("source", "user-said") not in ALLOWED_SOURCE:
            fail(f"source 只能是 {ALLOWED_SOURCE}")
        extra = FORBIDDEN_KEYS & set(p.keys())
        if extra:
            fail(f"preference 含禁止字段：{sorted(extra)}（只记偏好，不记原文与机密）")
    leaks = FORBIDDEN_KEYS & set(e.keys())
    if leaks:
        fail(f"条目含禁止字段：{sorted(leaks)}")
    e.setdefault("created", now_iso())
    e.setdefault("id", "mem-" + datetime.now().strftime("%Y%m%d%H%M%S"))
    e.setdefault("recordedBy", "ai-ppt-design-system")
    return e


def render_index(entries: list) -> str:
    lines = [
        "# 本地偏好记录",
        "",
        "> 本文件由 `scripts/pref_memory.py` 生成，只存在本机，请勿提交到任何仓库或网盘。",
        f"> 共 {len(entries)} 条；最新在前。删除单条：`pref_memory.py --forget <id>`。",
        "",
    ]
    for e in sorted(entries, key=lambda x: x.get("created", ""), reverse=True):
        lines.append(f"## {e.get('created', '')[:10]} · {e.get('project', '未命名')}（{e.get('kind', '')}）")
        lines.append("")
        lines.append(f"- 满意版本：`{e.get('satisfiedVersion', '未记')}`　id：`{e.get('id', '')}`")
        for p in e.get("preferences", []):
            star = "★" if p.get("source") == "user-said" else "·"
            lines.append(f"- {star} **{p.get('area', '')}**：{p.get('statement', '')}")
        avoid = e.get("avoid") or []
        if avoid:
            lines.append("- 明确不要：")
            for a in avoid:
                lines.append(f"  - {a}")
        ev = e.get("evidence")
        if ev:
            lines.append(f"- 依据：{ev}")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def cmd_status(d: Path) -> int:
    cfg = load_config(d)
    entries = load_entries(d)
    print(f"记忆目录：{d}")
    print(f"启用：{'是' if cfg.get('enabled') else '否'}")
    if cfg.get("consentedAt"):
        print(f"启用同意时间：{cfg['consentedAt']}")
    print(f"条目数：{len(entries)}")
    print("网络：本脚本不含任何网络调用")
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="本地偏好记忆（只写本机，永不联网）")
    ap.add_argument("--dir", default=None, help="记忆目录，默认 ~/.design-memory")
    ap.add_argument("--status", action="store_true")
    ap.add_argument("--enable", action="store_true")
    ap.add_argument("--disable", action="store_true")
    ap.add_argument("--add", metavar="ENTRY_JSON")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--export-md", action="store_true")
    ap.add_argument("--out", default=None)
    ap.add_argument("--forget", metavar="ID")
    ap.add_argument("--consented", action="store_true", help="用户已明确同意本次操作")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args(argv)

    d = resolve_dir(args.dir)

    if args.status:
        return cmd_status(d)

    if args.enable:
        require_consent(args, "启用偏好记忆")
        assert_not_in_repo(d)
        cfg = load_config(d)
        cfg.update({"enabled": True, "consentedAt": now_iso(), "defaultDir": str(d)})
        save_config(d, cfg)
        print(f"已启用：{d}（只存本机；随时可 --disable 或 --forget 删除）")
        return 0

    if args.disable:
        cfg = load_config(d)
        cfg["enabled"] = False
        save_config(d, cfg)
        print("已停用。已有条目保留，可用 --forget 逐条删除。")
        return 0

    if args.add:
        require_consent(args, "写入偏好")
        assert_not_in_repo(d)
        cfg = load_config(d)
        if not cfg.get("enabled"):
            fail("偏好记忆未启用。先用 --enable --consented 启用（并先问过用户）。")
        src = Path(args.add)
        if not src.is_file():
            fail(f"找不到条目文件：{src}", code=2)
        try:
            e = json.loads(src.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            fail(f"条目不是合法 JSON：{exc}", code=2)
        e = validate_entry(e)
        entries = load_entries(d)
        entries.append(e)
        write_entries(d, entries)
        (d / INDEX_NAME).write_text(render_index(entries), encoding="utf-8", newline="\n")
        print(f"已写入 1 条（共 {len(entries)} 条）：{e['id']} · {e['project']}")
        print(f"人类可读版：{d / INDEX_NAME}")
        return 0

    if args.list:
        entries = load_entries(d)
        if args.json:
            print(json.dumps(entries, ensure_ascii=False, indent=2))
            return 0
        if not entries:
            print("还没有条目。")
            return 0
        for e in sorted(entries, key=lambda x: x.get("created", ""), reverse=True):
            print(f"{e.get('created','')[:10]}  {e.get('id',''):22}  {e.get('project','')}  ({e.get('kind','')})  {len(e.get('preferences',[]))} 条偏好")
        return 0

    if args.export_md:
        entries = load_entries(d)
        if not entries:
            print("还没有条目，无需导出。")
            return 0
        text = render_index(entries)
        out = Path(args.out).expanduser() if args.out else (d / "PREFERENCES-export.md")
        out.write_text(text, encoding="utf-8", newline="\n")
        print(f"已导出：{out}（记得别放进任何仓库）")
        return 0

    if args.forget:
        entries = load_entries(d)
        keep = [e for e in entries if e.get("id") != args.forget]
        if len(keep) == len(entries):
            fail(f"没有这个 id：{args.forget}")
        write_entries(d, keep)
        (d / INDEX_NAME).write_text(render_index(keep), encoding="utf-8", newline="\n")
        print(f"已删除 1 条，剩 {len(keep)} 条。")
        return 0

    ap.print_help()
    return 2


if __name__ == "__main__":
    sys.exit(main())
