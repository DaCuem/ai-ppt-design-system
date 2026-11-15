#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""qa_check.py —— deck.json 自动质检（几何 / 溢出 / 字号 / 对比度 / 密度 / 重复 / 一致性 / 内容 / 结构）。

用法
    python scripts/qa_check.py --deck demo/smoke/qa-fixture.json \
        --out demo/smoke/qa-fixture-report.md --json demo/smoke/qa-fixture-report.json

退出码
    0  无 P0（P1/P2/P3 不阻断）
    1  存在 P0
    2  输入不可读：文件不存在 / JSON 非法 / 缺 meta|tokens|slides / pages < 1 / 阈值文件缺失

输出
    stdout 每行一条：`P1 s03 OVF-01 title 文本溢出 0.34in（needed 1.74in > box 1.40in）`
    末行：`SUMMARY P0=n P1=n P2=n P3=n issues=n pages=n exit=0|1|2`
    --out 写 Markdown 报告；--json 写机器可读 JSON。

只依赖标准库 + Pillow（经 lib/metrics.py）。不联网。
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import math
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:  # Windows 控制台默认 GBK，中文报告会乱码
    sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    sys.stderr.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
except (AttributeError, OSError):
    pass

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent
DEFAULT_THRESHOLDS = REPO_ROOT / "schema" / "thresholds.json"
SCHEMA_PATH = REPO_ROOT / "schema" / "deck.schema.json"

QA_SCHEMA_VERSION = "1.0.0"
SEVERITIES = ("P0", "P1", "P2", "P3")
SEVERITY_RANK = {level: index for index, level in enumerate(SEVERITIES)}


# ---------------------------------------------------------------------------
# 动态加载同目录 lib（不要求 lib 是包，不依赖他人文件）
# ---------------------------------------------------------------------------


def _load_lib(module_name: str, filename: str):
    """按文件路径加载 scripts/lib 下的模块，模块名加前缀避免冲突。"""
    path = SCRIPT_DIR / "lib" / filename
    if not path.is_file():
        raise SystemExit("错误：缺少依赖文件 %s（应与 qa_check.py 同仓库）" % path)
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise SystemExit("错误：无法加载 %s" % path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


try:
    colorlib = _load_lib("qa_color", "color.py")
    metricslib = _load_lib("qa_metrics", "metrics.py")
except SystemExit:
    raise
except Exception as exc:  # 依赖损坏时给出中文错误而不是堆栈
    raise SystemExit("错误：加载 scripts/lib 失败：%s" % exc)


# ---------------------------------------------------------------------------
# 检查项元数据：ID / 名称 / 级别 / 修复建议
# ---------------------------------------------------------------------------

CHECK_META: Dict[str, Tuple[str, str, str]] = {
    # id: (名称, 默认级别, 修复建议)
    "GEO-01": ("越出画布", "P0", "把 x/y/w/h 收回到画布内：0 ≤ x ≤ 13.333-w in、0 ≤ y ≤ 7.5-h in；满版装饰改用 zRole=bg 并允许出血"),
    "GEO-02": ("宽或高非正", "P0", "w/h 必须 > 0；line 用 x2/y2 表达端点，不要用 0 厚度"),
    "GEO-03": ("越出安全边距", "P2", "内容元素 x ≥ 0.6 in、y ≥ 0.45 in，右/下边界各留 0.6 in / 0.45 in"),
    "OVL-01": ("文本块重叠", "P3", "两个文本块间距拉到 0.05 in 以上；重叠会遮挡文字"),
    "OVF-01": ("文本溢出", "P1", "缩短文本、按阶梯降 1–2 pt，或把 h 加大到 needed_h_in；不要依赖 PowerPoint 自动缩放"),
    "OVF-02": ("文本严重溢出", "P0", "溢出 > 框高 30%：删减字数或拆页；正文每行 ≤ 34 字"),
    "OVF-03": ("每行字数超限", "P2", "每行中文字数压到 body 34 / h1 16 / caption 48 字以内；加宽文本框或分栏"),
    "OVF-04": ("段行数超限", "P3", "正文段落 ≤ 6 行、caption ≤ 2 行；超出改为要点或拆页"),
    "TYP-01": ("字号低于下限", "P1", "字号提到 role 的 minPt 之上：h1 36 / h2 24 / h3 20 / body 16 / caption 11 / pageNumber 9 pt"),
    "TYP-02": ("标题正文比失衡", "P2", "把 标题:正文 字号比调到 1.4–2.2:1（如 h1 48 pt / body 24 pt = 2.0:1）"),
    "CON-01": ("对比度不足", "P1", "文字与背景对比度 ≥ 4.5:1（大字 ≥ 3:1）；半透明元素按叠加后的实色计算"),
    "DEN-01": ("元素数超限", "P2", "每页元素 ≤ 14 个；合并纯装饰元素"),
    "DEN-02": ("文本块超限", "P2", "每页文本块 ≤ 4 个；同组信息合并进一个多段落文本框"),
    "DEN-03": ("bullet 超限", "P2", "每页 bullet ≤ 6 条；其余下沉到下一页或改为一句话结论"),
    "DEN-04": ("字数超限", "P2", "每页字数 ≤ 220 字；删掉形容词与重复表述"),
    "DEN-05": ("留白比例不足", "P3", "元素并集面积占比降到 ≤ 75%，为页面留出 ≥ 25% 的空白区"),
    "REP-01": ("相邻页版式重复", "P1", "相邻两页换版式：layout / gridTemplate / focalPosition 至少改一项"),
    "REP-02": ("版式占比过高", "P2", "同一版式签名占比 ≤ 35%；连续卡片页改用 comparison / timeline 等不同版式"),
    "REP-03": ("焦点位置连续重复", "P3", "焦点连续 3 页相同：至少一页把焦点移到另一侧，或改为 full"),
    "COL-01": ("颜色非法或未登记", "P2", "颜色只写 tokens.colors 的名字或 #RRGGBB；先补 token 再引用"),
    "COL-02": ("字号不在阶梯", "P3", "字号取 tokens.type 阶梯里的值，不要出现 17 pt 这类阶梯外字号"),
    "COL-03": ("字体族过多", "P2", "全 deck 字族 ≤ 3 个（display + body + mono）；装饰性字体并入 body"),
    "IMG-01": ("图片有效像素密度不足", "P2", "换更高分辨率的图，或缩小版面宽度：所需像素 = 版面英寸 × 110（阈值 qa.minImagePpi）；低于 80 ppi 报 P1"),
    "CNT-01": ("图片缺替代文本", "P1", "给 image 补 alt：描述图片承载的信息，不写文件名或“示意图”"),
    "CNT-02": ("图表缺单位或来源", "P1", "chart 必须填 unit 与 source，并把两者显示在图表附近"),
    "CNT-03": ("缺页码", "P3", "在每页（封面/封底除外）放 role=pageNumber 的页码元素"),
    "CNT-04": ("标题过长", "P2", "标题压到 ≤ 25 字；细节移到副标题或正文"),
    "STR-01": ("slide id 不连续", "P1", "slide id 必须从 s01 起连续递增且唯一"),
    "STR-02": ("页面数为零", "P0", "deck 至少要有 1 页 slides"),
    "THR-01": ("阈值键缺失", "P1", "在 schema/thresholds.json 补上该键（阈值文件是唯一真值，不要在脚本里写死数值）"),
}

#: 角色 → minPt 的阈值键
ROLE_MINPT_KEY = {
    "coverTitle": "h1",
    "h1": "h1",
    "h2": "h2",
    "h3": "h3",
    "lead": "body",
    "body": "body",
    "caption": "caption",
    "source": "caption",
    "pageNumber": "pageNumber",
    "kpi": "h2",
    "label": "caption",
    "quote": "h2",
}

#: 角色 → 每行中文字数上限（默认用 body 的 34 字/行）
ROLE_CHAR_LIMIT = {"h1": "h1", "coverTitle": "h1", "caption": "caption", "source": "caption", "body": "body"}

TEXT_KINDS = {"text"}
SHAPE_KINDS = {"rect", "ellipse", "polygon"}
CONTENT_ZROLES = {"content", "fgDecor"}


def _dig(mapping: Any, dotted: str) -> Any:
    """按 "a.b.c" 取嵌套字典值；任一层缺失返回 None。"""
    current = mapping
    for part in dotted.split("."):
        if not isinstance(current, dict) or part not in current:
            return None
        current = current[part]
    return current


class Thresholds:
    """阈值访问器：数值一律运行时来自 schema/thresholds.json。

    - ``num()``：该键属于阈值文件的标准内容；缺失时登记一条 THR-01（P1）并回退到规格值，
      而不是静默使用默认值。
    - ``opt_num()``：阈值文件未定义的补充常量（如标题/正文比区间），读得到就用，
      读不到用任务规格给定值，不报问题。
    """

    def __init__(self, raw: dict) -> None:
        self.raw = raw
        self.missing: Dict[str, str] = {}  # 键 → 回退说明；同一个键只登记一次

    def _miss(self, path: str, fallback_desc: str) -> None:
        if path not in self.missing:
            self.missing[path] = "%s（回退 %s）" % (path, fallback_desc)

    def num(self, path: str, fallback: float) -> float:
        value = _dig(self.raw, path)
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            self._miss(path, fallback)
            return float(fallback)
        return float(value)

    def opt_num(self, path: str, fallback: float) -> float:
        value = _dig(self.raw, path)
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            return float(fallback)
        return float(value)

    def seq(self, path: str, fallback: Sequence[Any]) -> List[Any]:
        value = _dig(self.raw, path)
        if not isinstance(value, list) or not value:
            self._miss(path, "、".join(str(v) for v in fallback))
            return list(fallback)
        return list(value)

    def flag(self, path: str, fallback: bool) -> bool:
        """布尔开关（如 charts.requireUnitAndSource）：读不到时用规格值，不登记缺失。"""
        value = _dig(self.raw, path)
        if isinstance(value, bool):
            return value
        return bool(fallback)


class DeckError(Exception):
    """deck 无法检查（退出码 2）。"""


@dataclass
class Issue:
    """一条质检问题。"""

    page: int
    slide_id: str
    severity: str
    check_id: str
    element: str
    message: str
    status: str = "待修复"

    @property
    def check_name(self) -> str:
        return CHECK_META.get(self.check_id, ("未知检查", "P3", ""))[0]

    @property
    def fix(self) -> str:
        return CHECK_META.get(self.check_id, ("", "", "按 schema 修正"))[2]

    def as_dict(self) -> dict:
        return {
            "page": self.page,
            "slideId": self.slide_id,
            "severity": self.severity,
            "checkId": self.check_id,
            "check": self.check_name,
            "element": self.element,
            "message": self.message,
            "fix": self.fix,
            "status": self.status,
        }


# ---------------------------------------------------------------------------
# 载入与令牌解析
# ---------------------------------------------------------------------------


VERBOSE = False  # --verbose 时打印 deckload 降级提示等诊断信息


def load_deck(path: Path) -> dict:
    """读取 deck.json。

    优先复用同目录 lib/deckload.py（它有 schema 校验与资产探测）；
    复用失败（文件缺失 / schema 不合法）时回退到内联 json 解析，
    保证"deck 有结构错误也能出质检报告"，而不是直接崩掉。
    """
    if not path.is_file():
        raise DeckError("找不到 deck 文件：%s" % path)

    deckload_path = SCRIPT_DIR / "lib" / "deckload.py"
    if deckload_path.is_file():
        try:
            module = _load_lib("qa_deckload", "deckload.py")
            loaded = module.load(str(path), check_assets=False)
            doc = getattr(loaded, "doc", None)
            if isinstance(doc, dict):
                missing = getattr(loaded, "missing_images", None)
                if missing:
                    # 只在本进程内使用，不写回 deck.json
                    doc["_qa_missing_images"] = [str(item) for item in missing]
                return doc
            if isinstance(loaded, tuple) and isinstance(loaded[0], dict):
                return loaded[0]
        except Exception as exc:  # schema 校验失败等：降级到内联解析继续检查
            if VERBOSE:
                print("提示：deckload 未能通过（%s），改用内置解析继续检查。" % exc, file=sys.stderr)

    try:
        with path.open("r", encoding="utf-8") as handle:
            deck = json.load(handle)
    except json.JSONDecodeError as exc:
        raise DeckError("deck JSON 解析失败：第 %d 行第 %d 列：%s" % (exc.lineno, exc.colno, exc.msg))
    except OSError as exc:
        raise DeckError("deck 文件读取失败：%s" % exc)
    if not isinstance(deck, dict):
        raise DeckError("deck 顶层必须是对象")
    return deck


def load_thresholds(path: Path) -> dict:
    """读取 thresholds.json，并校验本脚本用到的段落齐全。"""
    if not path.is_file():
        raise DeckError("找不到阈值文件：%s" % path)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise DeckError("阈值文件 JSON 解析失败：%s" % exc)
    required = ("canvas", "type", "density", "contrast", "repetition", "qa")
    missing = [key for key in required if key not in data]
    if missing:
        raise DeckError("阈值文件缺少字段：%s" % ", ".join(missing))
    return data


def validate_deck(deck: dict) -> None:
    """检查 deck 的最低结构要求（不完整则退出码 2）。"""
    for key in ("meta", "tokens", "slides"):
        if key not in deck:
            raise DeckError("deck 缺少必需字段：%s" % key)
    if not isinstance(deck["slides"], list):
        raise DeckError("deck.slides 必须是数组")
    if len(deck["slides"]) < 1:
        raise DeckError("deck 至少有 1 页（当前 0 页），无法质检")
    meta = deck["meta"]
    if not isinstance(meta, dict) or "size" not in meta:
        raise DeckError("deck.meta.size 缺失（应为 \"16:9\"/\"4:3\"/\"13.333x7.5\"）")
    tokens = deck["tokens"]
    if not isinstance(tokens, dict) or not isinstance(tokens.get("colors"), dict):
        raise DeckError("deck.tokens.colors 缺失")
    if "bg" not in tokens["colors"] or "fg" not in tokens["colors"]:
        raise DeckError("deck.tokens.colors 必须含 bg 与 fg")


def canvas_size(meta: dict, thresholds: dict) -> Tuple[float, float]:
    """解析 meta.size，返回 (宽, 高) 英寸。支持 16:9 / 4:3 / "13.333x7.5"。"""
    raw = str(meta.get("size", "16:9")).strip().lower().replace(" ", "")
    canvas = thresholds.get("canvas", {})
    if raw in ("16:9", "16x9"):
        entry = canvas.get("16:9", {})
        return float(entry.get("widthIn", 13.333)), float(entry.get("heightIn", 7.5))
    if raw in ("4:3", "4x3"):
        entry = canvas.get("4:3", {})
        return float(entry.get("widthIn", 10.0)), float(entry.get("heightIn", 7.5))
    match = re.match(r"^([0-9.]+)x([0-9.]+)$", raw)
    if match:
        return float(match.group(1)), float(match.group(2))
    raise DeckError("meta.size 无法识别：%r（应为 16:9 / 4:3 / 13.333x7.5）" % meta.get("size"))


def resolve_color_ref(ref: Any, tokens: dict) -> Tuple[Optional[str], bool]:
    """把 tokenRef 解析为 "#RRGGBB"。

    返回 (颜色, 是否合法)。合法 = 令牌存在，或本身就是合法 hex/rgb 字面量。
    """
    colors = tokens.get("colors") or {}
    if ref is None:
        return (None, True)
    if isinstance(ref, (tuple, list)):
        hex_value = colorlib.to_hex(ref)
        return (hex_value, hex_value is not None)
    text = str(ref).strip()
    if not text:
        return (None, True)
    if text.startswith("$"):
        name = text[1:]
        value = colors.get(name)
        hex_value = colorlib.to_hex(value)
        return (hex_value, hex_value is not None)
    if text in colors:
        return (colorlib.to_hex(colors[text]), colorlib.to_hex(colors[text]) is not None)
    hex_value = colorlib.to_hex(text)
    return (hex_value, hex_value is not None)


def color_or_none(ref: Any, tokens: dict, fallback: Optional[str] = None) -> Optional[str]:
    """解析颜色，非法时用 fallback。"""
    hex_value, ok = resolve_color_ref(ref, tokens)
    return hex_value if ok and hex_value else fallback


def normalize_text(text: Any) -> str:
    """把文本压成单行空格分隔，用于报告里的问题描述。"""
    return re.sub(r"\s+", " ", str(text or "")).strip()


def count_chars(text: str) -> int:
    """计字数：去掉所有空白后的字符数（中文按字、英文按字符）。"""
    return len(re.sub(r"\s+", "", str(text or "")))


def fmt_num(value: float, digits: int = 2) -> str:
    """数值格式化：去掉多余的 0，用于报告里的 in / pt。"""
    text = ("%." + str(digits) + "f") % float(value)
    return text.rstrip("0").rstrip(".") if "." in text else text


# ---------------------------------------------------------------------------
# 背景与文本样式解析
# ---------------------------------------------------------------------------


def slide_base_bg(slide: dict, tokens: dict) -> str:
    """页面底色：slide.bg → tokens.colors.bg → 白。"""
    bg = color_or_none(slide.get("bg"), tokens, None)
    if bg:
        return bg
    return color_or_none(tokens["colors"].get("bg"), tokens, "#FFFFFF") or "#FFFFFF"


def effective_bg(el: dict, slide: dict, tokens: dict) -> str:
    """元素背后的实际颜色：页面底色 + 叠在它下面的半透明/不透明色块 + 元素自身填充。

    只考虑被判定为背景层（zRole=bg/bgDecor）或几何填充块，且色块覆盖元素中心点。
    """
    base = slide_base_bg(slide, tokens)
    cx = float(el.get("x", 0) or 0) + float(el.get("w", 0) or 0) / 2.0
    cy = float(el.get("y", 0) or 0) + float(el.get("h", 0) or 0) / 2.0

    layers: List[Tuple[str, float]] = []
    for other in slide.get("elements") or []:
        if other is el or not isinstance(other, dict):
            continue
        kind = other.get("kind")
        zrole = str(other.get("zRole") or "content")
        if kind not in SHAPE_KINDS and kind not in ("image", "chart", "table"):
            continue
        if zrole not in ("bg", "bgDecor") and kind not in ("image", "chart", "table"):
            continue
        try:
            ox, oy = float(other.get("x", 0)), float(other.get("y", 0))
            ow, oh = float(other.get("w", 0)), float(other.get("h", 0))
        except (TypeError, ValueError):
            continue
        if not (ox <= cx <= ox + ow and oy <= cy <= oy + oh):
            continue
        if kind in ("image", "chart", "table"):
            layers.append(("#1A1A1A", 0.0))  # 图片/图表按"可能是深色"处理，不参与混合
            continue
        fill = color_or_none(other.get("fill"), tokens, None)
        if not fill:
            continue
        layers.append((fill, colorlib.parse_alpha(other.get("opacity", 1), 1.0)))

    for color, alpha in layers:
        if alpha <= 0:
            continue
        base = colorlib.to_hex(colorlib.blend_alpha(color, base, alpha)) or base

    own_fill = color_or_none(el.get("fill"), tokens, None)
    if own_fill and el.get("kind") in TEXT_KINDS:
        base = own_fill
    return base


def text_foreground(el: dict, tokens: dict) -> str:
    """文字实际显示颜色：element.color → tokens.colors.fg，并按 opacity 与背景混合。"""
    fg = color_or_none(el.get("color"), tokens, None)
    if not fg:
        fg = color_or_none(tokens["colors"].get("fg"), tokens, "#1A1A1A") or "#1A1A1A"
    return fg


def is_large_text(size_pt: float, weight: int, th: Thresholds) -> bool:
    """是否属于 WCAG 的「大字」：≥24 pt，或粗体 ≥18.66 pt（阈值来自 contrast.largeTextPt）。"""
    regular = th.num("contrast.largeTextPt.regular", 24)
    bold = th.num("contrast.largeTextPt.bold", 18.66)
    return size_pt >= regular or (weight >= 600 and size_pt >= bold)


# ---------------------------------------------------------------------------
# 质检主流程
# ---------------------------------------------------------------------------


@dataclass
class Context:
    """一次质检的全部上下文。"""

    deck: dict
    thresholds: dict
    th: Thresholds
    canvas_w: float
    canvas_h: float
    margin_x: float
    margin_y: float
    tol_overflow: float
    tol_overlap: float
    issues: List[Issue] = field(default_factory=list)
    stats: Dict[str, Any] = field(default_factory=dict)
    deck_dir: str = ""  # deck.json 所在目录；IMG-01 用它解析相对 src

    def add(self, page: int, slide_id: str, check_id: str, element: str, message: str, severity: Optional[str] = None) -> None:
        """追加一条问题（级别默认取 CHECK_META）。"""
        level = severity or CHECK_META.get(check_id, ("", "P3", ""))[1]
        self.issues.append(Issue(page=page, slide_id=slide_id, severity=level, check_id=check_id, element=element, message=message))


def build_context(deck: dict, thresholds: dict, deck_dir: str = "") -> Context:
    """由 deck + thresholds 构造上下文（含画布与安全边距，全部数值来自阈值文件）。"""
    width, height = canvas_size(deck["meta"], thresholds)
    th = Thresholds(thresholds)
    return Context(
        deck=deck,
        thresholds=thresholds,
        th=th,
        canvas_w=width,
        canvas_h=height,
        margin_x=th.num("canvas.safeMarginIn.x", 0.6),
        margin_y=th.num("canvas.safeMarginIn.y", 0.45),
        tol_overflow=th.num("qa.overflowToleranceIn", 0.02),
        tol_overlap=th.num("qa.overlapToleranceIn", 0.05),
        deck_dir=str(deck_dir or ""),
    )


def check_structure(ctx: Context) -> None:
    """STR-01 / STR-02：slide id 连续唯一、页数非零。"""
    slides = ctx.deck["slides"]
    if len(slides) < 1:  # validate_deck 已挡，这里保留防御
        ctx.add(0, "-", "STR-02", "pages", "页面数为 %d，至少需要 1 页" % len(slides), "P0")
        return
    seen: Dict[str, int] = {}
    for index, slide in enumerate(slides):
        slide_id = str(slide.get("id", ""))
        if slide_id in seen:
            ctx.add(index + 1, slide_id, "STR-01", "slide.id", "slide id 重复：%s 同时出现在第 %d 页与第 %d 页" % (slide_id, seen[slide_id], index + 1))
            continue
        seen[slide_id] = index + 1
        match = re.match(r"^s(\d{2,3})$", slide_id)
        if not match:
            ctx.add(index + 1, slide_id, "STR-01", "slide.id", "slide id 格式非法：%r（应为 s01、s02 …）" % slide_id)
            continue
        if int(match.group(1)) != index + 1:
            ctx.add(index + 1, slide_id, "STR-01", "slide.id", "slide id 不连续：第 %d 页应为 s%02d，实际为 %s" % (index + 1, index + 1, slide_id))


def check_geometry(ctx: Context, page: int, slide_id: str, el: dict, index: int) -> None:
    """GEO-01/02/03：画布越界、非正尺寸、安全边距。"""
    name = str(el.get("name") or el.get("role") or el.get("kind") or ("element%d" % index))
    try:
        x = float(el.get("x", 0))
        y = float(el.get("y", 0))
        w = float(el.get("w", 0))
        h = float(el.get("h", 0))
    except (TypeError, ValueError):
        ctx.add(page, slide_id, "GEO-02", name, "x/y/w/h 不是数字，无法定位")
        return
    if any(math.isnan(v) or math.isinf(v) for v in (x, y, w, h)):
        ctx.add(page, slide_id, "GEO-02", name, "x/y/w/h 含 NaN/Inf")
        return
    if w <= 0 or h <= 0:
        if el.get("kind") == "line":  # 直线用 x2/y2 表达端点，允许 0 厚度
            pass
        else:
            ctx.add(page, slide_id, "GEO-02", name, "尺寸非正：w=%.3fin h=%.3fin（必须 > 0）" % (w, h))
            return

    tol = ctx.tol_overflow
    right, bottom = x + w, y + h
    if x < -tol or y < -tol or right > ctx.canvas_w + tol or bottom > ctx.canvas_h + tol:
        sides = []
        if x < -tol:
            sides.append("左 %.2fin" % abs(x))
        if y < -tol:
            sides.append("上 %.2fin" % abs(y))
        if right > ctx.canvas_w + tol:
            sides.append("右 %.2fin" % (right - ctx.canvas_w))
        if bottom > ctx.canvas_h + tol:
            sides.append("下 %.2fin" % (bottom - ctx.canvas_h))
        ctx.add(page, slide_id, "GEO-01", name, "越出画布（%s；画布 %.3f×%.3f in）" % ("、".join(sides), ctx.canvas_w, ctx.canvas_h), "P0")
        return  # 已越画布，不再报安全边距

    zrole = str(el.get("zRole") or "content")
    role = str(el.get("role") or "")
    if zrole not in CONTENT_ZROLES or role == "decor":
        return  # 背景/装饰层允许出血
    if role == "pageNumber":
        return  # 页码定位在页脚边距内，属版式约定，不按安全边距判错
    inner_left, inner_top = ctx.margin_x, ctx.margin_y
    inner_right, inner_bottom = ctx.canvas_w - ctx.margin_x, ctx.canvas_h - ctx.margin_y
    breaches = []
    if x < inner_left - tol:
        breaches.append("左 %.2fin（应 ≥ %.2fin）" % (inner_left - x, inner_left))
    if y < inner_top - tol:
        breaches.append("上 %.2fin（应 ≥ %.2fin）" % (inner_top - y, inner_top))
    if right > inner_right + tol:
        breaches.append("右 %.2fin（应 ≤ %.3fin）" % (right - inner_right, inner_right))
    if bottom > inner_bottom + tol:
        breaches.append("下 %.2fin（应 ≤ %.3fin）" % (bottom - inner_bottom, inner_bottom))
    if breaches:
        ctx.add(page, slide_id, "GEO-03", name, "越出安全边距：" + "；".join(breaches))


def check_text_block(ctx: Context, page: int, slide_id: str, el: dict, index: int) -> None:
    """OVF-01/02/03/04 + TYP-01 + CON-01：单文本块的全部文本类检查。"""
    tokens = ctx.deck["tokens"]
    text = metricslib.element_text(el)
    if not text.strip():
        return
    name = str(el.get("name") or el.get("role") or ("text%d" % index))
    th = ctx.th

    report = metricslib.fit_report(el, tokens, line_em=th.num("type.lineBoxRatio", metricslib.LINE_EM))
    overflow = float(report["overflow_in"])
    box_h = float(report["box_h_in"])
    role = str(el.get("role") or "body")
    size_pt = float(report["size_pt"])
    weight = int(report["weight"])

    if overflow > ctx.tol_overflow:
        block_ratio = th.opt_num("qa.overflowRatioBlockP0", 0.30)
        if box_h > 0 and overflow / box_h > block_ratio:
            ctx.add(
                page,
                slide_id,
                "OVF-02",
                name,
                "文本严重溢出 %.2fin（needed %.2fin > box %.2fin，超出 %.0f%% > 上限 %.0f%%）"
                % (overflow, report["needed_h_in"], box_h, overflow / box_h * 100, block_ratio * 100),
                "P0",
            )
        else:
            ctx.add(
                page,
                slide_id,
                "OVF-01",
                name,
                "文本溢出 %.2fin（needed %.2fin > box %.2fin；%d 行）"
                % (overflow, report["needed_h_in"], box_h, int(report["lines"])),
            )

    max_chars = int(report["max_line_chars"])
    limit_key = ROLE_CHAR_LIMIT.get(role, "body")
    char_limit = int(th.num("type.maxCjkCharsPerLine." + limit_key, 34))
    if max_chars > char_limit:
        ctx.add(
            page,
            slide_id,
            "OVF-03",
            name,
            "每行 %d 字 > 上限 %d 字/行（role=%s）" % (max_chars, char_limit, role),
        )

    line_limit = th.opt_num("type.maxLinesPerTextBlock." + role, -1)
    if line_limit > 0 and int(report["lines"]) > int(line_limit):
        ctx.add(
            page,
            slide_id,
            "OVF-04",
            name,
            "段落 %d 行 > 上限 %d 行（role=%s）" % (int(report["lines"]), int(line_limit), role),
        )

    min_key = ROLE_MINPT_KEY.get(role)
    min_pt = th.num("type.minPt." + min_key, 0.0) if min_key else 0.0
    if min_pt > 0 and size_pt < min_pt - 1e-6:
        ctx.add(
            page,
            slide_id,
            "TYP-01",
            name,
            "字号 %.1fpt < %s 的 minPt %.0fpt" % (size_pt, min_key, min_pt),
        )

    bg = effective_bg(el, slide_of(ctx, slide_id), tokens)
    fg = text_foreground(el, tokens)
    opacity = colorlib.parse_alpha(el.get("opacity", 1), 1.0)
    if opacity < 1.0:
        fg = colorlib.to_hex(colorlib.blend_alpha(fg, bg, opacity)) or fg
    ratio = colorlib.contrast_ratio(fg, bg)
    if is_large_text(size_pt, weight, th):
        needed = th.num("contrast.largeText", 3.0)
    else:
        needed = th.num("contrast.bodyText", 4.5)
    if ratio < needed - 1e-6:
        ctx.add(
            page,
            slide_id,
            "CON-01",
            name,
            "对比度 %.2f:1 < %.1f:1（文字 %s / 背景 %s，%.0fpt%s）"
            % (ratio, needed, colorlib.to_hex(fg), colorlib.to_hex(bg), size_pt, "，大字" if needed < 4.5 else ""),
        )


_SLIDE_LOOKUP: Dict[int, dict] = {}


def slide_of(ctx: Context, slide_id: str) -> dict:
    """按 slide id 找 slide（供 effective_bg 使用）。"""
    for slide in ctx.deck["slides"]:
        if str(slide.get("id")) == slide_id:
            return slide
    return {}


def check_element_content(ctx: Context, page: int, slide_id: str, el: dict, index: int) -> None:
    """COL-01/02、CNT-01/02、CNT-04：颜色合法性、字号阶梯、图片 alt、图表单位来源、标题长度。"""
    tokens = ctx.deck["tokens"]
    kind = str(el.get("kind") or "")
    name = str(el.get("name") or el.get("role") or kind or ("element%d" % index))

    # 颜色合法性：所有出现的颜色引用都要能解析
    for key in ("color", "fill", "stroke"):
        if key in el and el[key] is not None:
            _, ok = resolve_color_ref(el[key], tokens)
            if not ok:
                ctx.add(page, slide_id, "COL-01", name, "%s 引用的颜色 %r 既不在 tokens.colors 也不是合法 #RRGGBB" % (key, el[key]))

    # 字号阶梯
    if kind in TEXT_KINDS and isinstance(el.get("size"), (int, float)):
        ladder = {
            float(entry["size"])
            for entry in (tokens.get("type") or {}).values()
            if isinstance(entry, dict) and isinstance(entry.get("size"), (int, float))
        }
        if ladder and float(el["size"]) not in ladder:
            ctx.add(page, slide_id, "COL-02", name, "字号 %.1fpt 不在 tokens.type 阶梯（%s pt）" % (float(el["size"]), "、".join(str(int(v)) for v in sorted(ladder))))

    if kind == "image":
        alt = str(el.get("alt") or "").strip()
        if not alt:
            ctx.add(page, slide_id, "CNT-01", name, "image 缺 alt（src=%s）" % normalize_text(el.get("src", "")))
        check_image_ppi(ctx, page, slide_id, el, name)

    if kind == "chart" and ctx.th.flag("charts.requireUnitAndSource", True):
        chart = el.get("chart") if isinstance(el.get("chart"), dict) else {}
        missing = [label for label, key in (("unit", "unit"), ("source", "source")) if not str(chart.get(key) or "").strip()]
        if missing:
            ctx.add(page, slide_id, "CNT-02", name, "chart 缺 %s（chartType=%s）" % (" 与 ".join(missing), el.get("chartType", "-")))

    role = str(el.get("role") or "")
    if kind in TEXT_KINDS and role in ("h1", "coverTitle"):
        max_title = int(ctx.th.opt_num("type.maxTitleChars", 25))
        chars = count_chars(metricslib.element_text(el))
        if chars > max_title:
            ctx.add(page, slide_id, "CNT-04", name, "标题 %d 字 > 上限 %d 字" % (chars, max_title))


def check_image_ppi(ctx: Context, page: int, slide_id: str, el: dict, name: str) -> None:
    """IMG-01：图片有效像素密度（像素宽 / 版面英寸）。

    数据来源是图片文件本身的像素尺寸，不是 deck.json 里写的数；
    文件缺失、不是图片、或读不出尺寸时静默跳过（缺图由 deckload 报）。
    """
    src = str(el.get("src") or "").strip()
    width_in = el.get("w")
    if not src or not isinstance(width_in, (int, float)) or float(width_in) <= 0:
        return
    src_path = Path(src)
    candidates = [src_path] if src_path.is_absolute() else [Path(ctx.deck_dir or ".") / src]
    path = next((c for c in candidates if c.is_file()), None)
    if path is None:
        return
    try:
        from PIL import Image  # 延迟导入：没有 Pillow 时不影响其它检查
        with Image.open(path) as img:
            px_w = float(img.size[0])
    except Exception:
        return
    ppi = px_w / float(width_in)
    warn_ppb = ctx.th.num("qa.minImagePpi", 110.0)
    block_ppb = ctx.th.num("qa.minImagePpiBlock", 80.0)
    if ppi < block_ppb - 1e-6:
        ctx.add(page, slide_id, "IMG-01", name,
                "有效 %.0f ppi < %.0f ppi（%d px 宽铺 %.2f in）——投影与打印都会糊"
                % (ppi, block_ppb, int(px_w), float(width_in)), severity="P1")
    elif ppi < warn_ppb - 1e-6:
        ctx.add(page, slide_id, "IMG-01", name,
                "有效 %.0f ppi < %.0f ppi（%d px 宽铺 %.2f in）——建议换图或缩小版面宽度"
                % (ppi, warn_ppb, int(px_w), float(width_in)))


def check_slide_density(ctx: Context, page: int, slide: dict) -> None:
    """DEN-01/02/03/04/05：元素数、文本块数、bullet 数、字数、留白比例。

    DEN-02 语义（与 Lead 约定、thresholds.density.textBlockExcludeRoles 一致）：
    只统计 kind=="text" 的元素，并排除 role ∈ textBlockExcludeRoles（caption/source/pageNumber）。
    label / kpi / quote 计入文本块——"标题 + 导语 + 眉标 + KPI 块"正好 4 个，
    第 5 个内容必须用 paragraphs/runs 合并进已有文本块。
    """
    slide_id = str(slide.get("id", ""))
    th = ctx.th
    elements = [el for el in (slide.get("elements") or []) if isinstance(el, dict)]
    exclude_roles = [str(role) for role in th.seq("density.textBlockExcludeRoles", ["caption", "source", "pageNumber"])]
    exclude_label = "/".join(exclude_roles)

    max_elements = int(th.num("density.maxElementsPerSlide", 14))
    if len(elements) > max_elements:
        ctx.add(page, slide_id, "DEN-01", "页面", "元素 %d 个 > 上限 %d 个" % (len(elements), max_elements))

    text_blocks = [
        el
        for el in elements
        if el.get("kind") in TEXT_KINDS
        and metricslib.element_text(el).strip()
        and str(el.get("role") or "") not in exclude_roles
    ]
    max_blocks = int(th.num("density.maxTextBlocksPerSlide", 4))
    if len(text_blocks) > max_blocks:
        ctx.add(
            page,
            slide_id,
            "DEN-02",
            "页面",
            "文本块 %d 个 > 上限 %d 个（已排除 role=%s）" % (len(text_blocks), max_blocks, exclude_label),
        )

    bullets = 0
    chars = 0
    for el in elements:  # 字数与 bullet 统计全页文本（含 caption/source），避免低估信息密度
        if el.get("kind") not in TEXT_KINDS:
            continue
        chars += count_chars(metricslib.element_text(el))
        for para in el.get("paragraphs") or []:
            if isinstance(para, dict) and para.get("bullet"):
                bullets += 1
    max_bullets = int(th.num("density.maxBulletItemsPerSlide", 6))
    if bullets > max_bullets:
        ctx.add(page, slide_id, "DEN-03", "页面", "bullet %d 条 > 上限 %d 条" % (bullets, max_bullets))
    max_chars = int(th.num("density.maxCharsPerSlide", 220))
    if chars > max_chars:
        ctx.add(page, slide_id, "DEN-04", "页面", "字数 %d 字 > 上限 %d 字" % (chars, max_chars))

    whitespace = _whitespace_ratio(ctx, slide)
    min_ws = th.num("density.minWhitespaceRatio", 0.25)
    if whitespace < min_ws - 1e-6:
        ctx.add(page, slide_id, "DEN-05", "页面", "留白比例 %.0f%% < 下限 %.0f%%（内容覆盖 %.0f%%）" % (whitespace * 100, min_ws * 100, (1 - whitespace) * 100))


def _whitespace_ratio(ctx: Context, slide: dict) -> float:
    """用 0.1 in 栅格估算元素并集覆盖率，返回留白比例（1 - 覆盖率）。

    背景/装饰层（zRole=bg/bgDecor）不计入，否则满版底色会让留白恒为 0。
    """
    step = 0.1
    cols = max(1, int(math.ceil(ctx.canvas_w / step)))
    rows = max(1, int(math.ceil(ctx.canvas_h / step)))
    covered = bytearray(cols * rows)
    for el in slide.get("elements") or []:
        if not isinstance(el, dict):
            continue
        zrole = str(el.get("zRole") or "content")
        if zrole in ("bg", "bgDecor"):
            continue
        try:
            x, y = float(el.get("x", 0)), float(el.get("y", 0))
            w, h = float(el.get("w", 0)), float(el.get("h", 0))
        except (TypeError, ValueError):
            continue
        x0, y0 = max(0, int(x / step)), max(0, int(y / step))
        x1, y1 = min(cols, int(math.ceil((x + w) / step))), min(rows, int(math.ceil((y + h) / step)))
        for row in range(y0, y1):
            base = row * cols
            for col in range(x0, x1):
                covered[base + col] = 1
    return 1.0 - (sum(covered) / float(cols * rows))


def check_text_overlap(ctx: Context, page: int, slide: dict) -> None:
    """OVL-01：两个文本块重叠（容差 0.05 in，仅 text-kind）。"""
    elements = [
        el
        for el in (slide.get("elements") or [])
        if isinstance(el, dict) and el.get("kind") in TEXT_KINDS and metricslib.element_text(el).strip()
    ]
    slide_id = str(slide.get("id", ""))
    for i in range(len(elements)):
        for j in range(i + 1, len(elements)):
            a, b = elements[i], elements[j]
            try:
                ax, ay = float(a["x"]), float(a["y"])
                aw, ah = float(a["w"]), float(a["h"])
                bx, by = float(b["x"]), float(b["y"])
                bw, bh = float(b["w"]), float(b["h"])
            except (KeyError, TypeError, ValueError):
                continue
            overlap_w = min(ax + aw, bx + bw) - max(ax, bx)
            overlap_h = min(ay + ah, by + bh) - max(ay, by)
            if overlap_w > ctx.tol_overlap and overlap_h > ctx.tol_overlap:
                name_a = str(a.get("name") or a.get("role") or "text%d" % i)
                name_b = str(b.get("name") or b.get("role") or "text%d" % j)
                ctx.add(
                    page,
                    slide_id,
                    "OVL-01",
                    "%s/%s" % (name_a, name_b),
                    "文本块重叠 %.2f×%.2fin（容差 %.2fin）" % (overlap_w, overlap_h, ctx.tol_overlap),
                )


def check_repetition(ctx: Context) -> None:
    """REP-01/02/03：相邻页版式签名重复、签名占比过高、焦点连续重复。"""
    slides = ctx.deck["slides"]
    fields = [str(field) for field in ctx.th.seq("repetition.signatureFields", ["layout", "gridTemplate", "focalPosition"])]
    signatures: List[Tuple[str, ...]] = []
    for slide in slides:
        signatures.append(tuple(str(slide.get(field) or "") for field in fields))

    # REP-01 相邻页同签名
    for index in range(1, len(slides)):
        if signatures[index] and signatures[index] == signatures[index - 1]:
            ctx.add(
                index + 1,
                str(slides[index].get("id", "")),
                "REP-01",
                "版式签名",
                "与上一页第 %d 页版式签名相同（%s）" % (index, " / ".join(v or "-" for v in signatures[index])),
            )

    # REP-02 签名占比（页数 < 4 时统计意义不足，跳过）
    if len(slides) >= 4:
        counts: Dict[Tuple[str, ...], List[int]] = {}
        for index, signature in enumerate(signatures):
            counts.setdefault(signature, []).append(index)
        max_share = ctx.th.num("repetition.maxSameLayoutShare", 0.35)
        for signature, pages in counts.items():
            share = len(pages) / float(len(slides))
            if share > max_share + 1e-9:
                ctx.add(
                    pages[0] + 1,
                    str(slides[pages[0]].get("id", "")),
                    "REP-02",
                    "版式签名",
                    "签名 %s 占 %d/%d 页 = %.0f%% > 上限 %.0f%%"
                    % (" / ".join(v or "-" for v in signature), len(pages), len(slides), share * 100, max_share * 100),
                )

    # REP-03 焦点位置连续 3 页相同
    run_start = 0
    for index in range(1, len(slides) + 1):
        current = slides[index].get("focalPosition") if index < len(slides) else None
        previous = slides[run_start].get("focalPosition")
        if index < len(slides) and current == previous:
            continue
        run_length = index - run_start
        if run_length >= 3 and previous not in (None, "", "none", "full"):
            ctx.add(
                index,
                str(slides[index - 1].get("id", "")),
                "REP-03",
                "focalPosition",
                "焦点位置 %s 连续 %d 页相同（第 %d–%d 页）" % (previous, run_length, run_start + 1, index),
            )
        run_start = index


def check_deck_consistency(ctx: Context) -> None:
    """TYP-02（标题正文比）、COL-03（字族数）以及页级颜色检查。"""
    tokens = ctx.deck["tokens"]
    type_tokens = tokens.get("type") or {}

    # TYP-02：标题字号 / 正文字号（区间由任务规格给定，thresholds.json 无对应键时用 opt 读取）
    ratio_min = ctx.th.opt_num("type.titleBodyRatioMin", 1.4)
    ratio_max = ctx.th.opt_num("type.titleBodyRatioMax", 2.2)
    title_size = None
    for key in ("h1", "coverTitle"):
        entry = type_tokens.get(key)
        if isinstance(entry, dict) and isinstance(entry.get("size"), (int, float)):
            title_size = float(entry["size"])
            break
    body_entry = type_tokens.get("body")
    body_size = float(body_entry["size"]) if isinstance(body_entry, dict) and isinstance(body_entry.get("size"), (int, float)) else None
    if title_size and body_size:
        ratio = title_size / body_size
        if ratio < ratio_min - 1e-9 or ratio > ratio_max + 1e-9:
            slide = ctx.deck["slides"][0]
            ctx.add(
                1,
                str(slide.get("id", "")),
                "TYP-02",
                "tokens.type",
                "标题/正文字号比 %.2f:1 不在 %.1f–%.1f:1（标题 %.0fpt / 正文 %.0fpt）" % (ratio, ratio_min, ratio_max, title_size, body_size),
            )

    # COL-03：字族数量
    max_families = int(ctx.th.opt_num("color.maxFontFamilies", 3))
    families: Dict[str, List[int]] = {}
    for index, slide in enumerate(ctx.deck["slides"]):
        for el in slide.get("elements") or []:
            if not isinstance(el, dict) or el.get("kind") not in TEXT_KINDS:
                continue
            if not metricslib.element_text(el).strip():
                continue
            family = metricslib.element_font_family(el, tokens)
            families.setdefault(family, []).append(index + 1)
    if len(families) > max_families:
        first_page = min(min(pages) for pages in families.values())
        ctx.add(
            first_page,
            str(ctx.deck["slides"][first_page - 1].get("id", "")),
            "COL-03",
            "tokens.fonts",
            "全 deck 出现 %d 个字族（%s），上限 %d 个" % (len(families), "、".join(sorted(families)), max_families),
        )


def check_page_number(ctx: Context, page: int, slide: dict) -> None:
    """CNT-03：缺页码（封面/封底除外）。"""
    layout = str(slide.get("layout") or "")
    if layout in ("cover", "closing"):
        return
    for el in slide.get("elements") or []:
        if not isinstance(el, dict):
            continue
        if str(el.get("role") or "") == "pageNumber":
            return
        text = normalize_text(metricslib.element_text(el))
        if re.fullmatch(r"P?\d{1,3}", text or ""):
            return
    ctx.add(page, str(slide.get("id", "")), "CNT-03", "页面", "缺页码（layout=%s）" % (layout or "-"))


def run_checks(deck: dict, thresholds: dict, deck_dir: str = "") -> Context:
    """跑完全部检查，返回 Context（含 issues 与统计）。"""
    ctx = build_context(deck, thresholds, deck_dir)
    check_structure(ctx)

    for page, slide in enumerate(deck["slides"], start=1):
        slide_id = str(slide.get("id", ""))
        elements = slide.get("elements") or []
        for index, el in enumerate(elements):
            if not isinstance(el, dict):
                continue
            check_geometry(ctx, page, slide_id, el, index)
            if el.get("kind") in TEXT_KINDS:
                check_text_block(ctx, page, slide_id, el, index)
            check_element_content(ctx, page, slide_id, el, index)
        check_slide_density(ctx, page, slide)
        check_text_overlap(ctx, page, slide)
        check_page_number(ctx, page, slide)

    check_repetition(ctx)
    check_deck_consistency(ctx)

    # THR-01：阈值键缺失不静默回退——逐键报 P1，指明缺哪个键（同键只报一次）
    if ctx.th.missing:
        first_id = str(deck["slides"][0].get("id", ""))
        for key in ctx.th.missing.values():
            ctx.add(1, first_id, "THR-01", "thresholds.json", "阈值键缺失：%s，请在 schema/thresholds.json 补上（脚本已按回退值继续检查）" % key)

    ctx.issues.sort(key=lambda item: (item.page, SEVERITY_RANK.get(item.severity, 9), item.check_id))
    counts = {level: 0 for level in SEVERITIES}
    for issue in ctx.issues:
        counts[issue.severity] = counts.get(issue.severity, 0) + 1
    ctx.stats = {
        "counts": counts,
        "total": len(ctx.issues),
        "pages": len(deck["slides"]),
        "exitCode": 1 if counts.get("P0", 0) > 0 else 0,
    }
    return ctx


# ---------------------------------------------------------------------------
# 输出
# ---------------------------------------------------------------------------


def render_console(ctx: Context) -> str:
    """控制台摘要：每条问题一行 + SUMMARY 行。"""
    lines = []
    for issue in ctx.issues:
        lines.append("%s %s %s %s %s" % (issue.severity, issue.slide_id, issue.check_id, issue.element, issue.message))
    counts = ctx.stats["counts"]
    lines.append(
        "SUMMARY P0=%d P1=%d P2=%d P3=%d issues=%d pages=%d exit=%d"
        % (counts.get("P0", 0), counts.get("P1", 0), counts.get("P2", 0), counts.get("P3", 0), ctx.stats["total"], ctx.stats["pages"], ctx.stats["exitCode"])
    )
    return "\n".join(lines)


def issue_category(check_id: str) -> str:
    """检查 ID → 类别（内容 / 设计 / 技术），对齐 templates/qa-report.md 的三选一。"""
    prefix = check_id.split("-")[0]
    if prefix in ("CNT",):
        return "内容"
    if prefix in ("DEN", "REP", "COL"):
        return "设计"
    return "技术"


def display_path(p: str) -> str:
    """把绝对路径转成便于进仓库的相对路径：优先相对仓库根（含 SKILL.md 的目录），其次相对当前目录。"""
    try:
        ap = Path(p).resolve()
    except OSError:
        return p
    for up in [ap.parent, *ap.parents]:
        if (up / "SKILL.md").exists():
            try:
                return ap.relative_to(up).as_posix()
            except ValueError:
                break
    try:
        return ap.relative_to(Path.cwd().resolve()).as_posix()
    except ValueError:
        return p


def render_markdown(ctx: Context, deck_path: str, thresholds_version: str) -> str:
    """生成 qa-report.md（9 列表格 + 分级统计行 + 结论行），对齐 templates/qa-report.md。"""
    counts = ctx.stats["counts"]
    generated = datetime.now().strftime("%Y-%m-%d %H:%M")
    blocked = counts.get("P0", 0) > 0
    verdict = "不通过（存在 P0，须修复并全量重渲后重跑）" if blocked else "通过（无 P0）"
    lines: List[str] = []
    lines.append("# QA 报告")
    lines.append("")
    lines.append("项目：%s ｜ 生成：%s ｜ 阈值版本：`schema/thresholds.json` v%s" % (display_path(deck_path), generated, thresholds_version))
    lines.append("")
    lines.append(
        "画布 %.3f × %.3f in ｜ 安全边距 %.2f / %.2f in ｜ 溢出容差 %.2f in ｜ 页数 %d"
        % (ctx.canvas_w, ctx.canvas_h, ctx.margin_x, ctx.margin_y, ctx.tol_overflow, ctx.stats["pages"])
    )
    lines.append("")
    lines.append("## 问题表")
    lines.append("")
    if ctx.issues:
        lines.append("| 编号 | 页码 | 类别 | 严重程度 | 检查项 | 问题描述 | 修复建议 | 状态 | 复查 |")
        lines.append("|---|---|---|---|---|---|---|---|---|")
        for number, issue in enumerate(ctx.issues, start=1):
            lines.append(
                "| Q-%02d | %s | %s | %s | %s | %s | %s | %s | — |"
                % (
                    number,
                    issue.slide_id or ("p%d" % issue.page),
                    issue_category(issue.check_id),
                    issue.severity,
                    issue.check_id,
                    issue.message.replace("|", "\\|"),
                    issue.fix.replace("|", "\\|"),
                    issue.status,
                )
            )
    else:
        lines.append("| — | — | — | — | — | 未发现结构/版面问题 | — | — | — |")
    lines.append("")
    lines.append(
        "**分级统计：P0=%d 条｜P1=%d 条｜P2=%d 条｜P3=%d 条｜合计 %d 条｜阻断=%s**"
        % (
            counts.get("P0", 0),
            counts.get("P1", 0),
            counts.get("P2", 0),
            counts.get("P3", 0),
            ctx.stats["total"],
            "是" if blocked else "否",
        )
    )
    lines.append("")
    lines.append("**结论：%s**" % verdict)
    lines.append("")
    lines.append("## 判定口径")
    lines.append("")
    lines.append("| 级别 | 判定标准（详见 schema/thresholds.json 与 workflows/06_quality-assurance.md） |")
    lines.append("|---|---|")
    lines.append("| P0 | 元素出画布、文字溢出 > 框高 30%、w/h ≤ 0：阻断交付 |")
    lines.append("| P1 | 溢出 > 0.02 in、字号低于下限、对比度不足、相邻页版式重复、缺 alt、图表缺 unit/source、slide id 不连续、阈值键缺失 |")
    lines.append("| P2 | 每行字数超限、标题正文比失衡、密度超限、单签名占比 > 35%、非法颜色、字族 > 3、标题 > 25 字 |")
    lines.append("| P3 | 段行数超限、文本块重叠、留白不足、字号不在阶梯、缺页码 |")
    lines.append("")
    lines.append("注：P1–P3 不阻断交付，但同一条检查 ID 在多页重复出现时按系统性问题处理——改 `tokens` 或版式，不逐页补。")
    lines.append("")
    return "\n".join(lines)


def render_json(ctx: Context, deck_path: str, thresholds_version: str) -> dict:
    """生成 qa.json 内容。"""
    counts = ctx.stats["counts"]
    return {
        "schemaVersion": QA_SCHEMA_VERSION,
        "deck": deck_path,
        "generatedAt": datetime.now().isoformat(timespec="seconds"),
        "thresholdsVersion": thresholds_version,
        "canvas": {"widthIn": ctx.canvas_w, "heightIn": ctx.canvas_h},
        "summary": {
            "P0": counts.get("P0", 0),
            "P1": counts.get("P1", 0),
            "P2": counts.get("P2", 0),
            "P3": counts.get("P3", 0),
            "total": ctx.stats["total"],
            "pages": ctx.stats["pages"],
            "exitCode": ctx.stats["exitCode"],
        },
        "issues": [issue.as_dict() for issue in ctx.issues],
    }


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    """命令行参数。"""
    parser = argparse.ArgumentParser(
        prog="qa_check.py",
        description="按 schema/thresholds.json 检查 deck.json（几何/溢出/字号/对比度/密度/重复/一致性/内容/结构）",
        epilog="退出码：0 无 P0；1 有 P0；2 输入不可读。示例：python scripts/qa_check.py --deck deck.json --out qa-report.md --json qa.json",
    )
    parser.add_argument("--deck", required=True, help="deck.json 路径")
    parser.add_argument("--out", default=None, help="Markdown 报告输出路径（省略则不写）")
    parser.add_argument("--json", dest="json_out", default=None, help="机器可读 JSON 输出路径（省略则不写）")
    parser.add_argument("--thresholds", default=str(DEFAULT_THRESHOLDS), help="阈值文件路径（默认 <repo>/schema/thresholds.json）")
    parser.add_argument("--verbose", action="store_true", help="打印 deckload 降级等诊断信息（默认静默）")
    parser.add_argument("--version", action="version", version="qa_check.py %s" % QA_SCHEMA_VERSION)
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    """入口：返回退出码。"""
    global VERBOSE
    args = parse_args(argv)
    VERBOSE = bool(args.verbose)
    deck_path = Path(args.deck).expanduser()
    thresholds_path = Path(args.thresholds).expanduser()

    try:
        thresholds = load_thresholds(thresholds_path)
        deck = load_deck(deck_path)
        validate_deck(deck)
        canvas_size(deck["meta"], thresholds)  # 提前暴露 meta.size 错误
    except DeckError as exc:
        print("ERROR %s" % exc)
        print("SUMMARY P0=0 P1=0 P2=0 P3=0 issues=0 pages=0 exit=2")
        return 2

    ctx = run_checks(deck, thresholds, str(deck_path.parent))
    print(render_console(ctx))

    deck_label = str(Path(args.deck))
    if args.out:
        out_path = Path(args.out).expanduser()
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(render_markdown(ctx, deck_label, str(thresholds.get("version", "unknown"))), encoding="utf-8")
        print("已写入 Markdown 报告：%s" % out_path, file=sys.stderr)
    if args.json_out:
        json_path = Path(args.json_out).expanduser()
        json_path.parent.mkdir(parents=True, exist_ok=True)
        payload = render_json(ctx, deck_label, str(thresholds.get("version", "unknown")))
        json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print("已写入 JSON 报告：%s" % json_path, file=sys.stderr)

    return int(ctx.stats["exitCode"])


if __name__ == "__main__":
    raise SystemExit(main())
