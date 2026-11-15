#!/usr/bin/env python3
"""环境能力探测：把「这台机器能不能构建、能不能出预览、有哪些字体」打印成 markdown。

用途：SKILL.md 第 0 步、docs/environment-report.md 都以此为唯一来源，避免文档里
写死「本机可用」而换台机器就翻车。默认输出 markdown，--json 输出同构数据。

检查项：
  - Python 版本与解释器路径
  - python-pptx / Pillow / lxml / jsonschema / pywin32 是否可用及版本
  - 渲染后端（ppt-render-cmd / dsh-native / soffice / powerpoint-com）的探测结果
  - 已安装字体，分成中文字体与西文字体，并给出可用字重

退出码：0 环境可构建（python-pptx 可用）/ 1 缺少构建必需的依赖（如 python-pptx）。

用法：
    python scripts/probe.py
    python scripts/probe.py --json > env.json
"""

from __future__ import annotations

import argparse
import json
import os
import platform
import re
import struct
import sys
from datetime import datetime
from typing import Dict, List, Optional, Sequence, Tuple

# 这些字体名出现时优先判定为「可能含中文」，只是加速启发式，最终以字形探测为准
_CJK_HINTS = (
    "cjk", "han", "hei", "song", "ming", "kai", "fang", "yahei", "yuan", "sans sc",
    "serif sc", "noto sans sc", "noto serif sc", "source han", "pingfang", "hiragino",
    "wenquanyi", "dengxian", "simsun", "simhei", "msjh", "msyh", "mingliu", "fangsong",
    "yuanti", "lisu", "kaiti",
)


def _configure_stdio() -> None:
    """Windows 控制台默认不是 UTF-8，中文会乱码。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[union-attr]
        except (AttributeError, ValueError):
            pass


# ---------------------------------------------------------------------------
# 依赖探测
# ---------------------------------------------------------------------------


def _module_version(name: str) -> Optional[str]:
    """import 一个模块并尽量取版本号；失败返回 None。"""
    try:
        module = __import__(name)
    except Exception:
        return None
    for attribute in ("__version__", "VERSION", "version"):
        value = getattr(module, attribute, None)
        if isinstance(value, str) and value:
            return value
        if isinstance(value, tuple):
            return ".".join(str(part) for part in value)
    return "已安装（版本未知）"


def probe_dependencies() -> List[dict]:
    """逐个探测依赖，返回 [{name, id, available, version, note}]。"""
    rows: List[dict] = []
    checks = [
        ("python-pptx", "pptx", "构建 PPTX 必需"),
        ("Pillow", "PIL", "图片尺寸与裁切必需"),
        ("lxml", "lxml", "python-pptx 的 XML 层依赖"),
        ("jsonschema", "jsonschema", "可选：有则交叉校验 deck.json"),
        ("pywin32", "win32com.client", "可选：PowerPoint COM 渲染后端"),
    ]
    for label, module, note in checks:
        version = _module_version(module)
        rows.append({
            "id": module,
            "name": label,
            "available": version is not None,
            "version": version or "",
            "note": note,
        })
    return rows


def probe_renderers() -> List[dict]:
    """复用 render.py 的后端探测逻辑（唯一真值，避免两处判断不一致）。"""
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    try:
        import render as render_mod  # noqa: WPS433 (运行时导入，避免循环依赖)
    except Exception as exc:  # pragma: no cover
        return [{
            "id": "unknown",
            "label": "渲染后端",
            "priority": 0,
            "available": False,
            "detail": f"无法导入 scripts/render.py：{exc}",
            "hint": "确认 render.py 与 probe.py 在同一目录",
        }]
    rows = []
    for index, backend in enumerate(render_mod.probe_backends(), start=1):
        rows.append({
            "id": backend.id,
            "label": backend.label,
            "priority": index,
            "available": backend.available,
            "detail": backend.detail,
            "hint": backend.hint,
        })
    return rows


# ---------------------------------------------------------------------------
# 字体探测
# ---------------------------------------------------------------------------


def font_directories() -> List[str]:
    """列出各平台常见字体目录（不存在的会被过滤掉）。"""
    dirs: List[str] = []
    if os.name == "nt":
        windir = os.environ.get("WINDIR", r"C:\Windows")
        dirs.append(os.path.join(windir, "Fonts"))
        local = os.environ.get("LOCALAPPDATA")
        if local:
            dirs.append(os.path.join(local, "Microsoft", "Windows", "Fonts"))
    elif sys.platform == "darwin":
        dirs.extend([
            "/System/Library/Fonts",
            "/System/Library/Fonts/Supplemental",
            "/Library/Fonts",
            os.path.expanduser("~/Library/Fonts"),
        ])
    else:
        dirs.extend([
            "/usr/share/fonts",
            "/usr/local/share/fonts",
            os.path.expanduser("~/.fonts"),
            os.path.expanduser("~/.local/share/fonts"),
        ])
    return [d for d in dirs if os.path.isdir(d)]


def _decode_name_record(data: bytes, platform_id: int, encoding_id: int) -> Optional[str]:
    try:
        if platform_id in (0, 3):
            return data.decode("utf-16-be", errors="replace").strip("\x00 ").strip()
        if platform_id == 1:
            return data.decode("mac-roman", errors="replace").strip()
    except Exception:
        return None
    return None


def read_font_names(path: str) -> Tuple[Optional[str], Optional[str]]:
    """从字体文件里读 (family, subfamily)。

    自己解析 sfnt 的 name 表：只依赖标准库，且拿到的是真实字体族名
    （"Microsoft YaHei"），而不是文件名（"msyh"）——AI 写 deck.json 时要用前者。
    支持 .ttf/.otf 与 .ttc（取第一个字体）；失败返回 (None, None)。
    """
    try:
        with open(path, "rb") as handle:
            header = handle.read(12)
            if len(header) < 12:
                return None, None
            tag = header[:4]
            if tag == b"ttcf":
                handle.seek(12)
                first = struct.unpack(">I", handle.read(4))[0]
                handle.seek(first)
                header = handle.read(12)
                if len(header) < 12:
                    return None, None
                tag = header[:4]
            if tag not in (b"\x00\x01\x00\x00", b"OTTO", b"true", b"typ1"):
                return None, None
            num_tables = struct.unpack(">H", header[4:6])[0]
            if not 0 < num_tables <= 512:
                return None, None
            table_dir = handle.read(num_tables * 16)
            table_offset = table_length = None
            for index in range(num_tables):
                record = table_dir[index * 16:(index + 1) * 16]
                if len(record) < 16:
                    break
                if record[:4] == b"name":
                    table_offset, table_length = struct.unpack(">II", record[8:16])
                    break
            if table_offset is None or table_length is None or table_length > 8 * 1024 * 1024:
                return None, None
            handle.seek(table_offset)
            table = handle.read(table_length)
    except (OSError, struct.error):
        return None, None
    try:
        _format, count, string_offset = struct.unpack(">HHH", table[:6])
    except struct.error:
        return None, None
    # name 表里同一项有多种语言，必须挑英文（en-US）：否则会拿到 "Negreta" 这种
    # 加泰罗尼亚语字重名，AI 抄进 deck.json 就错了
    best: Dict[int, Tuple[int, str]] = {}
    for index in range(count):
        base = 6 + index * 12
        chunk = table[base:base + 12]
        if len(chunk) < 12:
            break
        platform_id, encoding_id, language_id, name_id, length, offset = struct.unpack(">HHHHHH", chunk)
        if name_id not in (1, 2):
            continue
        raw = table[string_offset + offset:string_offset + offset + length]
        value = _decode_name_record(raw, platform_id, encoding_id)
        if not value:
            continue
        if platform_id == 3 and language_id == 0x0409:
            score = 0
        elif platform_id == 3:
            score = 1
        elif platform_id == 0:
            score = 2
        else:
            score = 3
        current = best.get(name_id)
        if current is None or score < current[0]:
            best[name_id] = (score, value)
    family = best[1][1] if 1 in best else None
    subfamily = best[2][1] if 2 in best else None
    return family, subfamily


def _font_tables(handle) -> Optional[Dict[bytes, Tuple[int, int]]]:
    """读 sfnt 表目录，返回 {表名: (偏移, 长度)}；.ttc 取第一个字体。"""
    handle.seek(0)
    header = handle.read(12)
    if len(header) < 12:
        return None
    if header[:4] == b"ttcf":
        handle.seek(12)
        raw = handle.read(4)
        if len(raw) < 4:
            return None
        handle.seek(struct.unpack(">I", raw)[0])
        header = handle.read(12)
        if len(header) < 12:
            return None
    if header[:4] not in (b"\x00\x01\x00\x00", b"OTTO", b"true", b"typ1"):
        return None
    count = struct.unpack(">H", header[4:6])[0]
    if not 0 < count <= 1024:
        return None
    directory = handle.read(count * 16)
    tables: Dict[bytes, Tuple[int, int]] = {}
    for index in range(count):
        record = directory[index * 16:(index + 1) * 16]
        if len(record) < 16:
            break
        tables[record[:4]] = struct.unpack(">II", record[8:16])
    return tables


def _cmap_has_codepoint(data: bytes, codepoint: int) -> bool:
    """在 cmap 表里查一个码位是否有映射（只解析 format 0/4/6/12，够用且快）。"""
    if len(data) < 4:
        return False
    try:
        _version, count = struct.unpack(">HH", data[:4])
    except struct.error:
        return False
    subtables: List[Tuple[int, int, int]] = []
    for index in range(min(count, 64)):
        record = data[4 + index * 8:12 + index * 8]
        if len(record) < 8:
            break
        platform_id, encoding_id, offset = struct.unpack(">HHI", record)
        subtables.append((platform_id, encoding_id, offset))
    # 平台 3（Windows）与平台 0（Unicode）的表最权威，先查它们
    subtables.sort(key=lambda item: 0 if item[0] == 3 else (1 if item[0] == 0 else 2))
    for _platform, _encoding, offset in subtables:
        if offset + 2 > len(data):
            continue
        try:
            fmt = struct.unpack(">H", data[offset:offset + 2])[0]
            if fmt == 4 and codepoint <= 0xFFFF:
                seg_count = struct.unpack(">H", data[offset + 6:offset + 8])[0] // 2
                if not 0 < seg_count <= 20000:
                    continue
                cursor = offset + 14
                ends = struct.unpack(">%dH" % seg_count, data[cursor:cursor + 2 * seg_count])
                cursor += 2 * seg_count + 2  # 跳过 reservedPad
                starts = struct.unpack(">%dH" % seg_count, data[cursor:cursor + 2 * seg_count])
                for start, end in zip(starts, ends):
                    if start <= codepoint <= end:
                        return True
            elif fmt == 12:
                groups = struct.unpack(">I", data[offset + 12:offset + 16])[0]
                if groups > 200000:
                    continue
                cursor = offset + 16
                for index in range(groups):
                    chunk = data[cursor + index * 12:cursor + index * 12 + 12]
                    if len(chunk) < 12:
                        break
                    start, end, _glyph = struct.unpack(">III", chunk)
                    if start <= codepoint <= end:
                        return True
            elif fmt == 6:
                first, char_count = struct.unpack(">HH", data[offset + 6:offset + 10])
                if first <= codepoint < first + char_count:
                    return True
            elif fmt == 0:
                if codepoint < 256 and data[offset + 6 + codepoint] != 0:
                    return True
        except (struct.error, IndexError):
            continue
    return False


def _supports_cjk(path: str) -> bool:
    """判断字体能否显示中文：直接查 cmap 里 U+4E2D（中）有没有映射。

    比 Pillow 渲染探测更可靠——Pillow 的 getmask 在缺字时也会返回 .notdef 方框，
    容易被误判成「有字形」或「无字形」；cmap 是渲染器真正用的那份数据。
    """
    try:
        with open(path, "rb") as handle:
            tables = _font_tables(handle)
            if not tables or b"cmap" not in tables:
                return False
            offset, length = tables[b"cmap"]
            if length > 8 * 1024 * 1024:
                return False
            handle.seek(offset)
            data = handle.read(length)
    except (OSError, struct.error):
        name = os.path.basename(path).lower()
        return any(hint in name for hint in _CJK_HINTS)
    return _cmap_has_codepoint(data, 0x4E2D) or _cmap_has_codepoint(data, 0x6C49)


def scan_fonts(limit_files: int = 4000) -> Dict[str, List[dict]]:
    """扫描字体目录，返回 {"cjk": [...], "latin": [...]}。"""
    extensions = (".ttf", ".otf", ".ttc", ".otc")
    families: Dict[str, dict] = {}
    scanned = 0
    for directory in font_directories():
        for root, _dirs, files in os.walk(directory):
            for name in files:
                if not name.lower().endswith(extensions):
                    continue
                scanned += 1
                if scanned > limit_files:
                    break
                path = os.path.join(root, name)
                family, subfamily = read_font_names(path)
                family = (family or os.path.splitext(name)[0]).strip()
                key = family.lower()
                entry = families.setdefault(key, {
                    "family": family,
                    "styles": [],
                    "files": [],
                    "cjk": False,
                })
                style = (subfamily or "Regular").strip()
                if style not in entry["styles"]:
                    entry["styles"].append(style)
                entry["files"].append(name)
                if not entry["cjk"]:
                    entry["cjk"] = _supports_cjk(path)
            if scanned > limit_files:
                break
    cjk = sorted((f for f in families.values() if f["cjk"]), key=lambda f: f["family"].lower())
    latin = sorted((f for f in families.values() if not f["cjk"]), key=lambda f: f["family"].lower())
    for group in (cjk, latin):
        for entry in group:
            entry["styleSummary"] = "、".join(sorted(entry["styles"], key=str.lower))
            entry["fileCount"] = len(entry["files"])
    return {"cjk": cjk, "latin": latin}


# ---------------------------------------------------------------------------
# 报告
# ---------------------------------------------------------------------------


def collect() -> dict:
    """采集全部能力数据（--json 与 markdown 共用同一份结构）。"""
    dependencies = probe_dependencies()
    renderers = probe_renderers()
    fonts = scan_fonts()
    versions = {row["id"]: row["available"] for row in dependencies}
    usable_renderers = [row["id"] for row in renderers if row["available"]]
    can_build = bool(versions.get("pptx"))
    return {
        "generatedAt": datetime.now().astimezone().isoformat(timespec="seconds"),
        "platform": {
            "system": platform.system(),
            "release": platform.release(),
            "machine": platform.machine(),
            "pythonVersion": platform.python_version(),
            "pythonExecutable": sys.executable,
            "cwd": os.getcwd(),
        },
        "dependencies": dependencies,
        "renderers": renderers,
        "fonts": {
            "directories": font_directories(),
            "cjkCount": len(fonts["cjk"]),
            "latinCount": len(fonts["latin"]),
            "cjk": fonts["cjk"],
            "latin": fonts["latin"],
        },
        "verdict": {
            "canBuild": can_build,
            "canRender": bool(usable_renderers),
            "renderersUsable": usable_renderers,
            "notes": _verdict_notes(dependencies, renderers, fonts),
        },
    }


def _verdict_notes(dependencies: List[dict], renderers: List[dict], fonts: Dict[str, List[dict]]) -> List[str]:
    notes: List[str] = []
    by_id = {row["id"]: row for row in dependencies}
    if not by_id["pptx"]["available"]:
        notes.append("缺少 python-pptx，构建 PPTX 不可用（pip install python-pptx）")
    if not by_id["PIL"]["available"]:
        notes.append("缺少 Pillow，image 的 fit/crop 无法按真实尺寸计算（pip install Pillow）")
    if not by_id["jsonschema"]["available"]:
        notes.append("没有 jsonschema：build_pptx.py 会退到内置校验器（对 schema/deck.schema.json 做等价检查）")
    usable = [row for row in renderers if row["available"]]
    if usable:
        notes.append(f"预览渲染优先使用 {usable[0]['id']}（{usable[0]['detail']}）")
    else:
        notes.append("没有任何可用渲染后端：只能构建 pptx，无法生成 PNG 预览，视觉检查不可做")
    if fonts["cjk"]:
        notes.append(f"检测到 {len(fonts['cjk'])} 个可显示中文的字体族，deck 的 tokens.fonts 必须取其中之一")
    else:
        notes.append("没有检测到中文字体：中文会回退成方框或宋体，先装字体再构建")
    return notes


def to_markdown(report: dict) -> str:
    """渲染成 markdown（供 SKILL.md / docs 直接粘贴）。"""
    lines: List[str] = []
    platform_info = report["platform"]
    lines.append("# 环境能力报告")
    lines.append("")
    lines.append(f"- 生成时间：{report['generatedAt']}")
    lines.append(f"- 平台：{platform_info['system']} {platform_info['release']} ({platform_info['machine']})")
    lines.append(f"- Python：{platform_info['pythonVersion']}（`{platform_info['pythonExecutable']}`）")
    lines.append("")

    lines.append("## 1. 运行时依赖")
    lines.append("")
    lines.append("| 依赖 | 状态 | 版本 | 说明 |")
    lines.append("| --- | --- | --- | --- |")
    for row in report["dependencies"]:
        status = "可用" if row["available"] else "缺失"
        mark = "✅" if row["available"] else "❌"
        lines.append(f"| {row['name']} | {mark} {status} | {row['version'] or '—'} | {row['note']} |")
    lines.append("")

    lines.append("## 2. 渲染后端")
    lines.append("")
    lines.append("| 优先级 | 后端 | 状态 | 探测结果 |")
    lines.append("| --- | --- | --- | --- |")
    for row in report["renderers"]:
        mark = "✅" if row["available"] else "❌"
        status = "就绪" if row["available"] else "不可用"
        detail = row["detail"].replace("|", "\\|")
        lines.append(f"| {row['priority']} | {row['label']}（`{row['id']}`） | {mark} {status} | {detail} |")
    lines.append("")
    for row in report["renderers"]:
        if not row["available"] and row.get("hint"):
            lines.append(f"- `{row['id']}` 缺失：{row['hint']}")
    if any(not row["available"] and row.get("hint") for row in report["renderers"]):
        lines.append("")

    fonts = report["fonts"]
    lines.append(f"## 3. 可用字体（中文字体 {fonts['cjkCount']} 个 / 西文字体 {fonts['latinCount']} 个）")
    lines.append("")
    lines.append("扫描目录：" + "、".join(f"`{d}`" for d in fonts["directories"]) if fonts["directories"] else "扫描目录：未找到标准字体目录")
    lines.append("")
    lines.append("### 中文字体")
    lines.append("")
    if fonts["cjk"]:
        lines.append("| 字体族 | 可用字重 | 文件数 |")
        lines.append("| --- | --- | --- |")
        for entry in fonts["cjk"]:
            lines.append(f"| {entry['family']} | {entry['styleSummary']} | {entry['fileCount']} |")
    else:
        lines.append("（未检测到中文字体，中文会回退，先安装字体）")
    lines.append("")
    lines.append("### 西文字体")
    lines.append("")
    if fonts["latin"]:
        lines.append("| 字体族 | 可用字重 | 文件数 |")
        lines.append("| --- | --- | --- |")
        for entry in fonts["latin"]:
            lines.append(f"| {entry['family']} | {entry['styleSummary']} | {entry['fileCount']} |")
    else:
        lines.append("（未检测到西文字体）")
    lines.append("")

    lines.append("## 4. 结论")
    lines.append("")
    verdict = report["verdict"]
    lines.append(f"- 能否构建 PPTX：{'能' if verdict['canBuild'] else '不能'}")
    lines.append(f"- 能否渲染 PNG 预览：{'能' if verdict['canRender'] else '不能'}")
    if verdict["renderersUsable"]:
        lines.append(f"- 可用渲染后端：{'、'.join('`' + r + '`' for r in verdict['renderersUsable'])}")
    for note in verdict["notes"]:
        lines.append(f"- {note}")
    lines.append("")
    return "\n".join(lines)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="probe.py",
        description="打印本机能力矩阵：依赖、渲染后端、可用字体（markdown 或 JSON）。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python scripts/probe.py\n"
            "  python scripts/probe.py --json > env.json\n"
            "\n"
            "退出码：0 可构建 / 1 缺少 python-pptx 等必需依赖"
        ),
    )
    parser.add_argument("--json", action="store_true", help="输出 JSON 而不是 markdown")
    parser.add_argument("--no-fonts", action="store_true", help="跳过字体扫描（更快，但字体一栏为空）")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    _configure_stdio()
    args = build_parser().parse_args(argv)
    report = collect()
    if args.no_fonts:
        report["fonts"] = {
            "directories": font_directories(),
            "cjkCount": 0,
            "latinCount": 0,
            "cjk": [],
            "latin": [],
        }
        report["verdict"]["notes"] = [n for n in report["verdict"]["notes"] if "字体" not in n]
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(to_markdown(report))
    if not report["verdict"]["canBuild"]:
        sys.stderr.write("错误：缺少 python-pptx，无法构建 PPTX（pip install python-pptx）\n")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
