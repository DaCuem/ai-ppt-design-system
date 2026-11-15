#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""字体度量与文本占位估算（qa_check 判断"文字会不会溢出"的核心）。

本模块把 schema 里的字体族名映射到本机实际字体文件，用 PIL.ImageFont 取真实字宽，
再用实测的行距系数换算成英寸，从而在不渲染 PPTX 的前提下估算文本块所需高度。

只依赖标准库 + Pillow。不联网、不写文件。

尺寸换算约定（与 schema/thresholds.json 的 basis 一致）
    dpi = 96（版面估算用），px = pt * dpi / 72；英寸 = px / dpi。
    即 16 pt 字号 → 21.33 px；6 行 16 pt 正文在 1.35 行距下约占 1.35 in。

行距实测（2026-11 在本机用 LibreOffice 渲染管线实测，见 scripts/README.md）
    单行行盒 = 1.204 × 字号；行距倍率 lineHeight 直接乘在这个行盒上。
    实测对象：Microsoft YaHei / DengXian / SimHei / Noto Sans SC / Noto Serif SC /
    Lato / Georgia / Bahnschrift / Cascadia Mono / SimSun / Source Han Serif SC，
    20 pt、lineHeight=1.35 下行距全部为 1.625 × 字号（= 1.35 × 1.204），字体间偏差 < 1%。
"""

from __future__ import annotations

import os
import unicodedata
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple, Union

from PIL import ImageFont

# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------

#: 版面估算用 DPI（96 dpi 时 1 in = 96 px，16:9 画布 13.333 in = 1280 px）
DEFAULT_DPI = 96

#: 单行行盒相对字号的倍率（实测值，字体间偏差 < 1%）
LINE_EM = 1.204

#: 无法识别字体时的兜底行距倍率（略保守，宁可报溢出也不漏报）
LINE_EM_FALLBACK = 1.22

#: schema 文本元素默认内边距（in），与 deck.schema.json 的 paddingIn 默认值一致
DEFAULT_PADDING_IN = 0.04

#: 无法换算 pt→px 时的兜底正文字号（pt）
FALLBACK_SIZE_PT = 18.0

#: 行首禁则：这些字符不能出现在行首（避头点）
_NO_BREAK_BEFORE = set(
    "，。、；：？！）】》」』”’…·%,.;:?!)]}>〞〟"
    "》〉】〕〗〙〛〗"
)
#: 行尾禁则：这些字符不能出现在行尾（避尾点）
_NO_BREAK_AFTER = set("（【《「『“‘([{〈《〔〖〘〚")

#: 按角色回退的字号（pt），仅当 tokens.type 里没有该角色时使用
ROLE_FALLBACK_PT: Dict[str, float] = {
    "coverTitle": 64.0,
    "h1": 48.0,
    "h2": 32.0,
    "h3": 24.0,
    "lead": 22.0,
    "body": 18.0,
    "caption": 12.0,
    "source": 10.0,
    "pageNumber": 10.0,
    "kpi": 40.0,
    "label": 12.0,
    "quote": 24.0,
    "decor": 12.0,
}

#: 使用 display 字体族的角色；其余角色用 body 字体族
DISPLAY_ROLES = {"coverTitle", "h1", "h2", "h3", "quote"}

#: 每个 bullet 层级占用的额外缩进（in）
BULLET_INDENT_IN = 0.28
#: bullet 圆点的近似占位宽度（in，正文 16-20 pt 时的经验值）
BULLET_PREFIX_IN = 0.22

#: 字体搜索目录（Windows 优先，其次用户字体与 Linux/mac 常见路径）
FONT_DIRS: Tuple[Path, ...] = tuple(
    Path(p)
    for p in (
        os.environ.get("WINDIR", r"C:\Windows") + r"\Fonts",
        os.path.expanduser("~") + r"\AppData\Local\Microsoft\Windows\Fonts",
        "/usr/share/fonts",
        "/usr/local/share/fonts",
        os.path.expanduser("~") + "/.fonts",
        "/System/Library/Fonts",
    )
    if p and Path(p).is_dir()
)


def _norm_family(name: str) -> str:
    """归一化字体族名：小写、去掉空格/连字符/下划线与常见后缀。"""
    if not name:
        return ""
    text = str(name).strip().strip("'\"").lower()
    for suffix in ("-regular", "-bold", "-light", "-medium", " regular", " bold", " mt", " ms"):
        if text.endswith(suffix):
            text = text[: -len(suffix)]
    return "".join(ch for ch in text if ch.isalnum())


# ---------------------------------------------------------------------------
# 字体族 → 本机文件 映射表
# 结构：归一化族名 -> {weight: (文件名, ttc index, variable 名或 None), ...}
# ---------------------------------------------------------------------------

_FAMILY_FILES: Dict[str, Dict[int, Tuple[str, int, Optional[str]]]] = {
    # ---- 中文：黑体系 ----
    "microsoftyahei": {300: ("msyhl.ttc", 0, None), 400: ("msyh.ttc", 0, None), 700: ("msyhbd.ttc", 0, None), 900: ("msyhbd.ttc", 0, None)},
    "msyh": {400: ("msyh.ttc", 0, None), 700: ("msyhbd.ttc", 0, None)},
    "微软雅黑": {400: ("msyh.ttc", 0, None), 700: ("msyhbd.ttc", 0, None)},
    "microsoftyaheiui": {400: ("msyh.ttc", 1, None), 700: ("msyhbd.ttc", 1, None)},
    "微软雅黑ui": {400: ("msyh.ttc", 1, None), 700: ("msyhbd.ttc", 1, None)},
    "dengxian": {300: ("Dengl.ttf", 0, None), 400: ("Deng.ttf", 0, None), 700: ("Dengb.ttf", 0, None), 900: ("Dengb.ttf", 0, None)},
    "等线": {400: ("Deng.ttf", 0, None), 700: ("Dengb.ttf", 0, None)},
    "simhei": {400: ("simhei.ttf", 0, None), 700: ("simhei.ttf", 0, None)},
    "黑体": {400: ("simhei.ttf", 0, None), 700: ("simhei.ttf", 0, None)},
    "notosanssc": {300: ("NotoSansSC-VF.ttf", 0, "Light"), 400: ("NotoSansSC-VF.ttf", 0, "Regular"), 500: ("NotoSansSC-VF.ttf", 0, "Medium"), 700: ("NotoSansSC-VF.ttf", 0, "Bold"), 900: ("NotoSansSC-VF.ttf", 0, "Bold")},
    "notosanscjksc": {400: ("NotoSansSC-VF.ttf", 0, "Regular"), 700: ("NotoSansSC-VF.ttf", 0, "Bold")},
    "思源黑体": {400: ("NotoSansSC-VF.ttf", 0, "Regular"), 700: ("NotoSansSC-VF.ttf", 0, "Bold")},
    "sourcehansanssc": {400: ("NotoSansSC-VF.ttf", 0, "Regular"), 700: ("NotoSansSC-VF.ttf", 0, "Bold")},
    "harfbuzzsanssc": {400: ("NotoSansSC-VF.ttf", 0, "Regular"), 700: ("NotoSansSC-VF.ttf", 0, "Bold")},
    # ---- 中文：宋体系 ----
    "simsun": {400: ("simsun.ttc", 0, None), 700: ("simsunb.ttf", 0, None)},
    "宋体": {400: ("simsun.ttc", 0, None), 700: ("simsunb.ttf", 0, None)},
    "newsong": {400: ("simsun.ttc", 0, None)},
    "notoserifsc": {300: ("NotoSerifSC-VF.ttf", 0, "Light"), 400: ("NotoSerifSC-VF.ttf", 0, "Regular"), 600: ("NotoSerifSC-VF.ttf", 0, "SemiBold"), 700: ("NotoSerifSC-VF.ttf", 0, "Bold"), 900: ("NotoSerifSC-VF.ttf", 0, "Bold")},
    "思源宋体": {400: ("NotoSerifSC-VF.ttf", 0, "Regular"), 700: ("NotoSerifSC-VF.ttf", 0, "Bold")},
    "sourcehanserifsc": {900: ("Source Han Serif SC Heavy (TrueType).ttf", 0, None), 700: ("Source Han Serif SC Heavy (TrueType).ttf", 0, None), 400: ("NotoSerifSC-VF.ttf", 0, "Regular")},
    # ---- 西文 ----
    "lato": {300: ("Lato-Light.ttf", 0, None), 400: ("Lato-Regular.ttf", 0, None), 500: ("Lato-Medium.ttf", 0, None), 600: ("Lato-Semibold.ttf", 0, None), 700: ("Lato-Bold.ttf", 0, None), 900: ("Lato-Black.ttf", 0, None)},
    "latobold": {700: ("Lato-Bold.ttf", 0, None)},
    "latolight": {300: ("Lato-Light.ttf", 0, None)},
    "latoblack": {900: ("Lato-Black.ttf", 0, None)},
    "latosemibold": {600: ("Lato-Semibold.ttf", 0, None)},
    "latomedium": {500: ("Lato-Medium.ttf", 0, None)},
    "latohairline": {100: ("Lato-Hairline.ttf", 0, None)},
    "latothin": {200: ("Lato-Thin.ttf", 0, None)},
    "bahnschrift": {300: ("bahnschrift.ttf", 0, "Light"), 400: ("bahnschrift.ttf", 0, "Regular"), 600: ("bahnschrift.ttf", 0, "SemiBold"), 700: ("bahnschrift.ttf", 0, "Bold"), 900: ("bahnschrift.ttf", 0, "Bold")},
    "georgia": {400: ("georgia.ttf", 0, None), 700: ("georgiab.ttf", 0, None)},
    "cascadiamono": {400: ("CascadiaMono.ttf", 0, "Regular"), 700: ("CascadiaMono.ttf", 0, "Bold")},
    "cascadiacode": {400: ("CascadiaCode.ttf", 0, None), 700: ("CascadiaCode.ttf", 0, None)},
    "consolas": {400: ("consola.ttf", 0, None), 700: ("consolab.ttf", 0, None)},
    "arial": {400: ("arial.ttf", 0, None), 700: ("arialbd.ttf", 0, None)},
    "helvetica": {400: ("arial.ttf", 0, None), 700: ("arialbd.ttf", 0, None)},
    "calibri": {400: ("calibri.ttf", 0, None), 700: ("calibrib.ttf", 0, None)},
    "segoeui": {400: ("segoeui.ttf", 0, None), 700: ("segoeuib.ttf", 0, None)},
    "timesnewroman": {400: ("times.ttf", 0, None), 700: ("timesbd.ttf", 0, None)},
    "verdana": {400: ("verdana.ttf", 0, None), 700: ("verdanab.ttf", 0, None)},
    "tahoma": {400: ("tahoma.ttf", 0, None), 700: ("tahomabd.ttf", 0, None)},
    # ---- 通用族（按中英混排首选） ----
    "sansserif": {400: ("msyh.ttc", 0, None), 700: ("msyhbd.ttc", 0, None)},
    "systemui": {400: ("msyh.ttc", 0, None), 700: ("msyhbd.ttc", 0, None)},
    "serif": {400: ("NotoSerifSC-VF.ttf", 0, "Regular"), 700: ("NotoSerifSC-VF.ttf", 0, "Bold")},
    "monospace": {400: ("CascadiaMono.ttf", 0, "Regular"), 700: ("CascadiaMono.ttf", 0, "Bold")},
}

#: 找不到指定族时的回退顺序（先中文黑体，保证中文不出豆腐块）
FALLBACK_CHAIN: Tuple[str, ...] = (
    "microsoftyahei",
    "notosanssc",
    "dengxian",
    "simhei",
    "georgia",
    "arial",
)

#: 中文字形兜底字体（西文字体不含中文字形时使用）
CJK_FALLBACK_FAMILY = "microsoftyahei"

_WEIGHT_BUCKETS = (100, 200, 300, 400, 500, 600, 700, 800, 900)


@dataclass(frozen=True)
class FontSpec:
    """一个可加载的具体字体：文件路径 + ttc 索引 + 可变字体实例名。"""

    family: str          # 解析后的族名（报告与一致性检查用）
    requested: str        # 原始请求的族名
    path: str             # 绝对路径
    index: int = 0
    variation: Optional[str] = None
    weight: int = 400
    is_cjk: bool = True

    @property
    def key(self) -> Tuple[str, int, Optional[str], int]:
        """缓存键。"""
        return (self.path, self.index, self.variation, self.weight)

    @property
    def id(self) -> str:
        """稳定短标识，写入 fit_report。"""
        return "%s@%d%s" % (
            Path(self.path).name,
            self.index,
            (":" + self.variation) if self.variation else "",
        )


def _find_file(filename: str) -> Optional[str]:
    """在字体目录里查找文件（大小写不敏感）。"""
    for directory in FONT_DIRS:
        candidate = Path(directory) / filename
        if candidate.is_file():
            return str(candidate)
    # 退一步：在第一个字体目录里大小写不敏感匹配
    for directory in FONT_DIRS:
        try:
            for entry in Path(directory).iterdir():
                if entry.name.lower() == filename.lower():
                    return str(entry)
        except OSError:
            continue
    return None


def _pick_weight(mapping: Dict[int, Tuple[str, int, Optional[str]]], weight: int) -> Tuple[str, int, Optional[str]]:
    """在族内按字重取最接近的可用档位（向上取整优先，标题加粗更保险）。"""
    if weight in mapping:
        return mapping[weight]
    available = sorted(mapping)
    above = [w for w in available if w >= weight]
    if above:
        return mapping[above[0]]
    return mapping[available[-1]]


def _is_cjk_file(path: str) -> bool:
    """判断字体文件是否包含中文字形（按文件名白名单，保守判定）。"""
    name = Path(path).name.lower()
    return any(token in name for token in ("msyh", "deng", "simhei", "simsun", "noto sans sc", "notosanssc", "noto serif sc", "notoserifsc", "source han", "sourcehan", "微软", "hei", "song", "kai"))


@lru_cache(maxsize=256)
def resolve_font(family: Optional[str], weight: int = 400) -> FontSpec:
    """把字体族名解析为具体字体文件。

    解析失败时按 FALLBACK_CHAIN 回退，绝不抛异常（返回的 spec.requested 保留原请求族名）。
    例：resolve_font("Lato", 700) -> FontSpec(path=...Lato-Bold.ttf)
        resolve_font("Source Han Serif SC", 900) -> ...Source Han Serif SC Heavy (TrueType).ttf
    """
    requested = str(family).strip() if family else ""
    normalized = _norm_family(requested)
    weight = int(weight) if weight else 400
    weight = max(100, min(900, weight))

    if normalized and normalized in _FAMILY_FILES:
        filename, index, variation = _pick_weight(_FAMILY_FILES[normalized], weight)
        found = _find_file(filename)
        if found:
            return FontSpec(
                family=requested,
                requested=requested,
                path=found,
                index=index,
                variation=variation,
                weight=weight,
                is_cjk=_is_cjk_file(found),
            )

    # 回退链
    for fallback in FALLBACK_CHAIN:
        filename, index, variation = _pick_weight(_FAMILY_FILES[fallback], weight)
        found = _find_file(filename)
        if found:
            return FontSpec(
                family=fallback,
                requested=requested or fallback,
                path=found,
                index=index,
                variation=variation,
                weight=weight,
                is_cjk=_is_cjk_file(found),
            )

    raise FileNotFoundError(
        "找不到可用字体：请求 %r；已搜索目录 %s。请确认 C:\\Windows\\Fonts 下存在 msyh.ttc 等中文字体。"
        % (requested, ", ".join(str(d) for d in FONT_DIRS))
    )


@lru_cache(maxsize=64)
def _load_font(path: str, px: int, index: int, variation: Optional[str]) -> ImageFont.FreeTypeFont:
    """按像素字号加载字体（带缓存），可变字体按实例名取字重。"""
    font = ImageFont.truetype(path, size=px, index=index)
    if variation:
        try:
            font.set_variation_by_name(variation)
        except (OSError, ValueError, AttributeError):
            pass  # 非可变字体或名字不存在：忽略，按默认实例度量
    return font


def size_px(pt: float, dpi: int = DEFAULT_DPI) -> int:
    """pt → px（至少 1 px，避免 PIL 报错）。"""
    return max(1, int(round(float(pt) * dpi / 72.0)))


def is_cjk_char(ch: str) -> bool:
    """判断字符是否按全角中文处理（CJK 汉字、假名、谚文、全角符号、中文标点）。"""
    if not ch:
        return False
    code = ord(ch)
    if 0x3040 <= code <= 0x30FF:      # 平假名/片假名
        return True
    if 0x3400 <= code <= 0x4DBF:      # CJK 扩展 A
        return True
    if 0x4E00 <= code <= 0x9FFF:      # CJK 基本区
        return True
    if 0xF900 <= code <= 0xFAFF:      # CJK 兼容汉字
        return True
    if 0xAC00 <= code <= 0xD7AF:      # 谚文
        return True
    if 0x3000 <= code <= 0x303F:      # CJK 标点（、。〈〉《》「」等）
        return True
    if 0xFF00 <= code <= 0xFFEF:      # 全角形式
        return True
    if code in (0x2018, 0x2019, 0x201C, 0x201D, 0x2026, 0x2014, 0x00B7):
        return False  # 引号/省略号/破折号按西文宽度量，误差 < 0.2 em
    return False


def char_width_px(ch: str, spec: FontSpec, px: int) -> float:
    """单个字符的推进宽度（px）。中文走中文字体，西文走主字体，缺字形时互相兜底。"""
    if ch in ("\t",):
        ch = "    "
    font = _load_font(spec.path, px, spec.index, spec.variation)
    width = font.getlength(ch)
    if width <= 0 and ch.strip():
        # 主字体没有该字形（常见于西文字体遇到中文）：换中文字体再量
        cjk = resolve_font(CJK_FALLBACK_FAMILY, spec.weight)
        if cjk.path != spec.path:
            return _load_font(cjk.path, px, cjk.index, cjk.variation).getlength(ch)
    return width


@dataclass
class _Measurer:
    """一个字号下的测量器：拆主字体 / 中文字体两条通路，缓存字符宽度。"""

    spec: FontSpec
    px: int
    letter_spacing_px: float = 0.0
    cjk_spec: Optional[FontSpec] = None
    _cache: Dict[str, float] = field(default_factory=dict, repr=False)

    def __post_init__(self) -> None:
        if self.cjk_spec is None:
            self.cjk_spec = resolve_font(CJK_FALLBACK_FAMILY, self.spec.weight)

    def font_for(self, ch: str) -> FontSpec:
        """该字符应该用哪个字体量。"""
        if is_cjk_char(ch) and not self.spec.is_cjk:
            return self.cjk_spec or self.spec
        return self.spec

    def width(self, ch: str) -> float:
        """单字符宽度（含字距）。"""
        cached = self._cache.get(ch)
        if cached is None:
            cached = char_width_px(ch, self.font_for(ch), self.px) + self.letter_spacing_px
            self._cache[ch] = cached
        return cached

    def text_width(self, text: str) -> float:
        """整串宽度（px）。"""
        return sum(self.width(ch) for ch in text)


def _make_measurer(font: Union[str, FontSpec, None], pt: float, dpi: int, weight: int, letter_spacing_pt: float) -> _Measurer:
    """构造测量器；font 可以是路径、族名或 FontSpec。"""
    if isinstance(font, FontSpec):
        spec = font
    elif font and Path(str(font)).is_file():
        spec = FontSpec(family=Path(str(font)).stem, requested=str(font), path=str(font), weight=weight, is_cjk=_is_cjk_file(str(font)))
    else:
        spec = resolve_font(str(font) if font else "Microsoft YaHei", weight)
    px = size_px(pt, dpi)
    spacing_px = float(letter_spacing_pt or 0.0) * dpi / 72.0
    return _Measurer(spec=spec, px=px, letter_spacing_px=spacing_px)


def measure_text(
    text: str,
    font_path: Union[str, FontSpec, None] = None,
    pt: float = FALLBACK_SIZE_PT,
    dpi: int = DEFAULT_DPI,
    weight: int = 400,
    letter_spacing_pt: float = 0.0,
) -> float:
    """测量文本宽度，返回英寸（取最长一行的宽度）。

    参数
        text             文本；含换行时按最长行返回
        font_path        字体族名 / 字体文件路径 / FontSpec；None 用 Microsoft YaHei
        pt               字号（pt）
        dpi              版面 DPI，默认 96
        weight           字重，决定同族取哪个字重文件
        letter_spacing_pt 字距（pt），按字符累加
    """
    if text is None:
        return 0.0
    measurer = _make_measurer(font_path, pt, dpi, weight, letter_spacing_pt)
    lines = str(text).replace("\t", "    ").split("\n")
    widest = max((measurer.text_width(line) for line in lines), default=0.0)
    return widest / dpi


def measure_lines(text: str, font_path=None, pt: float = FALLBACK_SIZE_PT, dpi: int = DEFAULT_DPI, weight: int = 400, letter_spacing_pt: float = 0.0) -> List[float]:
    """逐行测量宽度（英寸），便于报告里定位是哪一行太长。"""
    if text is None:
        return []
    measurer = _make_measurer(font_path, pt, dpi, weight, letter_spacing_pt)
    return [measurer.text_width(line.replace("\t", "    ")) / dpi for line in str(text).split("\n")]


def can_break_before(text: str, i: int) -> bool:
    """判断能否在下标 i 之前断行（中文任意断、西文按词断、遵守中英文避头尾规则）。"""
    if i <= 0 or i >= len(text):
        return False
    prev, cur = text[i - 1], text[i]
    if cur in _NO_BREAK_BEFORE or prev in _NO_BREAK_AFTER:
        return False
    if prev.isspace():
        return True
    if is_cjk_char(prev) or is_cjk_char(cur):
        return True
    if prev == "-" and cur.isalnum():
        return True
    return False


def _wrap_paragraph(text: str, width_px: float, measurer: _Measurer) -> Tuple[List[str], int, int]:
    """单段折行，返回 (行列表, 硬断次数, 最长行字符数)。

    硬断 = 单个不可断单元（长英文单词/长数字串）本身超过行宽，被迫在词中断开。
    """
    text = text.replace("\t", "    ")
    if not text:
        return [""], 0, 0

    widths = [measurer.width(ch) for ch in text]
    n = len(text)
    lines: List[str] = []
    hard_breaks = 0
    start = 0
    while start < n:
        # 跳过行首空格
        while start < n and text[start] == " ":
            start += 1
        if start >= n:
            break
        acc = 0.0
        last_break = -1
        end = start
        while end < n:
            acc += widths[end]
            if acc > width_px and end > start:
                break
            if can_break_before(text, end + 1):
                last_break = end + 1
            end += 1
        if end >= n:
            lines.append(text[start:])
            break
        if last_break > start:
            cut = last_break
        else:
            cut = end if end > start else start + 1
            hard_breaks += 1
        lines.append(text[start:cut].rstrip())
        start = cut

    longest = max((cjk_equiv_chars(line) for line in lines), default=0)
    return lines or [""], hard_breaks, longest


def cjk_equiv_chars(text: str) -> int:
    """把一行折算成"中文字数"：中文 1 字、西文/数字 0.5 字，用于 maxCjkCharsPerLine 判定。"""
    if not text:
        return 0
    count = 0
    for ch in text:
        if ch.isspace():
            continue
        count += 2 if is_cjk_char(ch) else 1
    return int(round(count / 2.0 + 0.1))


def wrap_lines(
    text: str,
    width_in: float,
    font_path: Union[str, FontSpec, None] = None,
    pt: float = FALLBACK_SIZE_PT,
    dpi: int = DEFAULT_DPI,
    weight: int = 400,
    letter_spacing_pt: float = 0.0,
) -> List[str]:
    """按可用宽度折行，返回行文本列表（含显式换行切出的空行）。

    中文逐字可断（遵守避头尾），西文按词断；超长不可断单元会被硬断。
    width_in <= 0 时返回原文单行（由调用方按"框宽非法"报 P0）。
    """
    if text is None:
        return []
    if width_in <= 0:
        return [line for line in str(text).split("\n")]
    measurer = _make_measurer(font_path, pt, dpi, weight, letter_spacing_pt)
    width_px = width_in * dpi
    lines: List[str] = []
    for paragraph in str(text).split("\n"):
        para_lines, _, _ = _wrap_paragraph(paragraph, width_px, measurer)
        lines.extend(para_lines)
    return lines


def line_pitch_in(pt: float, line_height: float = 1.35, em_ratio: float = LINE_EM) -> float:
    """单行行距（英寸）：pt × lineHeight × 行盒系数 / 72。"""
    return max(0.0, float(pt) * float(line_height) * float(em_ratio) / 72.0)


def estimate_block_height(
    lines: Union[int, Sequence[str]],
    pt: float,
    line_height: float = 1.35,
    space_before_pt: float = 0.0,
    em_ratio: float = LINE_EM,
) -> float:
    """估算文本块高度（英寸）：行数 × 行距 + 段前距。

    lines 可以是行数（int）或行文本列表（len 即行数）。
    """
    n = lines if isinstance(lines, int) else len(list(lines))
    height = n * line_pitch_in(pt, line_height, em_ratio)
    height += float(space_before_pt or 0.0) / 72.0
    return max(0.0, height)


# ---------------------------------------------------------------------------
# 元素级：取文本 / 字号 / 字体 / 折行 / 溢出报告
# ---------------------------------------------------------------------------


def element_text(element: dict) -> str:
    """取出元素的纯文本：text / paragraphs[*].text 或 runs[*].text 拼接。"""
    if not isinstance(element, dict):
        return ""
    if element.get("text"):
        return str(element["text"])
    paragraphs = element.get("paragraphs")
    if isinstance(paragraphs, list) and paragraphs:
        parts: List[str] = []
        for para in paragraphs:
            if not isinstance(para, dict):
                continue
            if para.get("text"):
                parts.append(str(para["text"]))
            elif isinstance(para.get("runs"), list):
                parts.append("".join(str(run.get("text", "")) for run in para["runs"] if isinstance(run, dict)))
            else:
                parts.append("")
        return "\n".join(parts)
    runs = element.get("runs")
    if isinstance(runs, list) and runs:
        return "".join(str(run.get("text", "")) for run in runs if isinstance(run, dict))
    return ""


def element_paragraphs(element: dict) -> List[dict]:
    """把元素拆成段落列表：[{text, size, weight, bullet, level, spaceBeforePt, align}]。"""
    if not isinstance(element, dict):
        return []

    def _run_extremes(runs: Iterable[dict]) -> Tuple[Optional[float], Optional[int]]:
        sizes = [float(r["size"]) for r in runs if isinstance(r, dict) and isinstance(r.get("size"), (int, float))]
        weights = [int(r["weight"]) for r in runs if isinstance(r, dict) and isinstance(r.get("weight"), (int, float))]
        return (max(sizes) if sizes else None, max(weights) if weights else None)

    base = {
        "size": element.get("size"),
        "weight": element.get("weight"),
        "align": element.get("align", "left"),
    }
    paragraphs = element.get("paragraphs")
    if isinstance(paragraphs, list) and paragraphs:
        result = []
        for index, para in enumerate(paragraphs):
            if not isinstance(para, dict):
                continue
            text = str(para.get("text") or "")
            if not text and isinstance(para.get("runs"), list):
                text = "".join(str(run.get("text", "")) for run in para["runs"] if isinstance(run, dict))
            size, weight = _run_extremes(para.get("runs") or [])
            result.append(
                {
                    "index": index,
                    "text": text,
                    "size": size or base["size"],
                    "weight": weight or base["weight"],
                    "bullet": bool(para.get("bullet", False)),
                    "level": int(para.get("level", 0) or 0),
                    "spaceBeforePt": float(para.get("spaceBeforePt", 0) or 0),
                    "align": para.get("align", base["align"]),
                }
            )
        return result

    runs = element.get("runs")
    text = element_text(element)
    size, weight = _run_extremes(runs or [])
    if text == "" and not (isinstance(runs, list) and runs):
        return []
    chunks = text.split("\n")
    return [
        {
            "index": index,
            "text": chunk,
            "size": size or base["size"],
            "weight": weight or base["weight"],
            "bullet": False,
            "level": 0,
            "spaceBeforePt": 0.0,
            "align": base["align"],
        }
        for index, chunk in enumerate(chunks)
    ]


def resolve_type_scale(tokens: dict, role: str) -> dict:
    """从 tokens.type 取角色的字号阶梯；缺失时返回空字典。"""
    scale = (tokens or {}).get("type") or {}
    entry = scale.get(role)
    return entry if isinstance(entry, dict) else {}


def element_font_family(element: dict, tokens: dict) -> str:
    """元素的字体族名：element.font → tokens.fonts.display/body → 回退。"""
    if isinstance(element, dict) and element.get("font"):
        return str(element["font"])
    fonts = (tokens or {}).get("fonts") or {}
    role = str((element or {}).get("role") or "body")
    if role in DISPLAY_ROLES:
        return str(fonts.get("display") or fonts.get("body") or "Microsoft YaHei")
    return str(fonts.get("body") or fonts.get("display") or "Microsoft YaHei")


def element_size_pt(element: dict, tokens: dict) -> float:
    """元素字号（pt）：element.size → tokens.type[role].size → 角色回退值。"""
    if isinstance(element, dict) and isinstance(element.get("size"), (int, float)):
        return float(element["size"])
    role = str((element or {}).get("role") or "body")
    scale = resolve_type_scale(tokens, role)
    if isinstance(scale.get("size"), (int, float)):
        return float(scale["size"])
    return float(ROLE_FALLBACK_PT.get(role, FALLBACK_SIZE_PT))


def element_weight(element: dict, tokens: dict) -> int:
    """元素字重：element.weight → tokens.type[role].weight → 400。"""
    if isinstance(element, dict) and isinstance(element.get("weight"), (int, float)):
        return int(element["weight"])
    role = str((element or {}).get("role") or "body")
    scale = resolve_type_scale(tokens, role)
    if isinstance(scale.get("weight"), (int, float)):
        return int(scale["weight"])
    return 700 if role in DISPLAY_ROLES else 400


def element_line_height(element: dict, tokens: dict) -> float:
    """元素行距倍率：element.lineHeight → tokens.type[role].lineHeight → 1.35。"""
    if isinstance(element, dict) and isinstance(element.get("lineHeight"), (int, float)):
        return float(element["lineHeight"])
    role = str((element or {}).get("role") or "body")
    scale = resolve_type_scale(tokens, role)
    if isinstance(scale.get("lineHeight"), (int, float)):
        return float(scale["lineHeight"])
    return 1.35


def element_letter_spacing(element: dict, tokens: dict) -> float:
    """元素字距（pt）：element.letterSpacing → tokens.type[role].letterSpacing → 0。"""
    if isinstance(element, dict) and isinstance(element.get("letterSpacing"), (int, float)):
        return float(element["letterSpacing"])
    role = str((element or {}).get("role") or "body")
    scale = resolve_type_scale(tokens, role)
    if isinstance(scale.get("letterSpacing"), (int, float)):
        return float(scale["letterSpacing"])
    return 0.0


def fit_report(
    element: dict,
    tokens: Optional[dict] = None,
    dpi: int = DEFAULT_DPI,
    line_em: Optional[float] = None,
) -> dict:
    """估算一个文本元素是否溢出，返回可直接写进 QA 报告的字典型结果。

    返回键
        needed_h_in     文本实际需要的高度（含内边距）
        box_h_in        元素框高（schema 的 h）
        overflow_in     needed - box（负值表示有余量）
        overflow_ratio  overflow / box_h
        lines           视觉行数合计
        max_line_chars  最长行的"中文字数"（中文按 1、西文按 0.5 折算）
        available_w_in  可排版宽度（已扣除左右内边距、bullet 缩进）
        font / size_pt / line_height / weight / line_em
        paragraphs      每段的 {lines, chars, hard_breaks}
        hard_breaks     不可断长单元被硬断的次数（>0 说明该词需要手动改短）
        empty           元素无文本

    line_em 是行盒系数（行距 = 字号 × lineHeight × line_em / 72 英寸），默认取实测值
    ``LINE_EM``；调用方可传入 ``thresholds.type.lineBoxRatio``，让质检读数与阈值文件一致。
    """
    tokens = tokens or {}
    element = element or {}
    em = float(line_em) if line_em else LINE_EM
    padding = float(element.get("paddingIn", DEFAULT_PADDING_IN) or 0.0)
    box_h = float(element.get("h", 0) or 0)
    box_w = float(element.get("w", 0) or 0)
    available_w = box_w - 2 * padding

    family = element_font_family(element, tokens)
    base_size = element_size_pt(element, tokens)
    base_weight = element_weight(element, tokens)
    line_height = element_line_height(element, tokens)
    letter_spacing = element_letter_spacing(element, tokens)
    spec = resolve_font(family, base_weight)

    paragraphs = element_paragraphs(element)
    if not paragraphs:
        return {
            "needed_h_in": 0.0,
            "box_h_in": round(box_h, 4),
            "overflow_in": 0.0,
            "overflow_ratio": 0.0,
            "lines": 0,
            "max_line_chars": 0,
            "available_w_in": round(available_w, 4),
            "font": spec.family,
            "font_id": spec.id,
            "size_pt": round(base_size, 2),
            "line_height": round(line_height, 3),
            "weight": base_weight,
            "line_em": em,
            "paragraphs": [],
            "hard_breaks": 0,
            "empty": True,
        }

    needed = 0.0
    total_lines = 0
    max_chars = 0
    total_hard = 0
    para_reports: List[dict] = []
    max_size = max((float(p["size"] or base_size) for p in paragraphs), default=base_size)

    for para in paragraphs:
        size = float(para["size"] or base_size)
        weight = int(para["weight"] or base_weight)
        level = int(para["level"] or 0)
        width = available_w - level * BULLET_INDENT_IN
        if para["bullet"]:
            width -= BULLET_PREFIX_IN
        if width <= 0.01:
            width = 0.01
        lines, hard, longest = _wrap_paragraph(str(para["text"] or ""), width * dpi, _make_measurer(family, size, dpi, weight, letter_spacing))
        count = len(lines)
        para_needed = estimate_block_height(count, size, line_height, float(para["spaceBeforePt"]), em)
        needed += para_needed
        total_lines += count
        total_hard += hard
        max_chars = max(max_chars, longest)
        para_reports.append(
            {
                "index": para["index"],
                "lines": count,
                "chars": cjk_equiv_chars(str(para["text"] or "")),
                "hard_breaks": hard,
                "bullet": bool(para["bullet"]),
                "height_in": round(para_needed, 4),
            }
        )

    # 文本自"顶部内边距"开始排：占据高度 = 上内边距 + 行数 × 行距。
    # 底部内边距不增加所需高度（PowerPoint/LibreOffice 的底 inset 只影响自动缩放与换行），
    # 若按上下双边距相加会系统性多报 0.03–0.04 in——实测样例 deck 的标题框高恰好是
    # 「行距 + 上 inset」，双边距估算法会把 6 个正常文本块误判成溢出。
    needed += padding
    overflow = needed - box_h
    return {
        "needed_h_in": round(needed, 4),
        "box_h_in": round(box_h, 4),
        "overflow_in": round(overflow, 4),
        "overflow_ratio": round(overflow / box_h, 4) if box_h > 0 else 0.0,
        "lines": total_lines,
        "max_line_chars": max_chars,
        "available_w_in": round(available_w, 4),
        "font": spec.family,
        "font_id": spec.id,
        "size_pt": round(base_size, 2),
        "max_size_pt": round(max_size, 2),
        "line_height": round(line_height, 3),
        "weight": base_weight,
        "line_em": em,
        "paragraphs": para_reports,
        "hard_breaks": total_hard,
        "empty": False,
    }


def block_width_in(text: str, font_path=None, pt: float = FALLBACK_SIZE_PT, weight: int = 400, letter_spacing_pt: float = 0.0) -> float:
    """measure_text 的语义化别名：文本不折行时的宽度（in）。"""
    return measure_text(text, font_path, pt, DEFAULT_DPI, weight, letter_spacing_pt)


def available_families() -> List[str]:
    """列出本机可解析的字体族（诊断用，qa_check --fonts 打印）。"""
    found: List[str] = []
    for key, mapping in _FAMILY_FILES.items():
        filename, _, _ = _pick_weight(mapping, 400)
        if _find_file(filename):
            found.append(key)
    return sorted(set(found))


__all__ = [
    "DEFAULT_DPI",
    "LINE_EM",
    "LINE_EM_FALLBACK",
    "DEFAULT_PADDING_IN",
    "FALLBACK_SIZE_PT",
    "ROLE_FALLBACK_PT",
    "DISPLAY_ROLES",
    "BULLET_INDENT_IN",
    "BULLET_PREFIX_IN",
    "FontSpec",
    "resolve_font",
    "available_families",
    "size_px",
    "is_cjk_char",
    "char_width_px",
    "measure_text",
    "measure_lines",
    "wrap_lines",
    "can_break_before",
    "cjk_equiv_chars",
    "line_pitch_in",
    "estimate_block_height",
    "element_text",
    "element_paragraphs",
    "resolve_type_scale",
    "element_font_family",
    "element_size_pt",
    "element_weight",
    "element_line_height",
    "element_letter_spacing",
    "fit_report",
    "block_width_in",
]
