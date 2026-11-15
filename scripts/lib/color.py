#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""颜色工具（qa_check / contact_sheet 共用）。

公式来自 WCAG 2.2 SC 1.4.3（对比度）与 CSS Color 4 的 sRGB 相对亮度定义。
本模块只依赖标准库，不读写文件、不联网。

主要接口
    parse_hex(value)                  -> (r, g, b) 或 None
    srgb_luminance(color)             -> 0.0-1.0 相对亮度
    contrast_ratio(fg, bg)            -> 1.0-21.0
    blend_alpha(fg, bg, alpha)        -> (r, g, b)
    is_dark(color, threshold=0.18)    -> bool
    hue_distance(a, b)                -> 0.0-180.0 度
    color_name_zh(color)              -> 中文色名，如 "深蓝"

所有颜色入参接受 "#RRGGBB"、"#RGB"、非井号写法、"rgb(r, g, b)" 字符串、
以及 (r, g, b) / (r, g, b, a) 元组；alpha 取 0-1 浮点或 0-255 整数（>1 视为 0-255）。
"""

from __future__ import annotations

import re
from typing import Iterable, Optional, Sequence, Tuple, Union

RGB = Tuple[int, int, int]
ColorLike = Union[str, Sequence[float], None]

_HEX_RE = re.compile(r"^#?([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$")
_RGB_FN_RE = re.compile(
    r"^rgba?\(\s*([0-9.]+%?)\s*[, ]\s*([0-9.]+%?)\s*[, ]\s*([0-9.]+%?)\s*(?:[,/]\s*([0-9.]+%?)\s*)?\)$",
    re.IGNORECASE,
)

# WCAG 2.2 SC 1.4.3 阈值（与 schema/thresholds.json 的 contrast 段一致）
CONTRAST_BODY = 4.5
CONTRAST_LARGE = 3.0
LARGE_TEXT_PT_REGULAR = 24.0
LARGE_TEXT_PT_BOLD = 18.66


def _clamp_byte(value: float) -> int:
    """把数值夹到 0-255 后取整。"""
    if value != value:  # NaN
        return 0
    return max(0, min(255, int(round(value))))


def _channel_to_unit(value: float) -> float:
    """把 0-255 通道值转成 0.0-1.0。"""
    return _clamp_byte(value) / 255.0


def _parse_component(text: str) -> int:
    """解析 rgb() 里的单个分量，支持百分号。"""
    if text.endswith("%"):
        return _clamp_byte(float(text[:-1]) / 100.0 * 255.0)
    return _clamp_byte(float(text))


def parse_hex(value: ColorLike) -> Optional[RGB]:
    """解析颜色为 (r, g, b)。

    支持 "#RRGGBB" / "RRGGBB" / "#RGB" / "#RRGGBBAA"（忽略 AA，请改用 blend_alpha）、
    "rgb(1, 2, 3)" / "rgba(1,2,3,0.5)"，以及 (r, g, b) 序列。
    无法解析时返回 None（调用方据此报"颜色非法"），不抛异常。
    """
    if value is None:
        return None

    if isinstance(value, (tuple, list)):
        if len(value) < 3:
            return None
        try:
            return (_clamp_byte(float(value[0])), _clamp_byte(float(value[1])), _clamp_byte(float(value[2])))
        except (TypeError, ValueError):
            return None

    text = str(value).strip()
    if not text:
        return None

    match = _RGB_FN_RE.match(text)
    if match:
        try:
            return tuple(_parse_component(match.group(i)) for i in (1, 2, 3))  # type: ignore[return-value]
        except (TypeError, ValueError):
            return None

    match = _HEX_RE.match(text)
    if not match:
        return None
    digits = match.group(1)
    if len(digits) == 3:
        digits = "".join(ch * 2 for ch in digits)
    elif len(digits) == 8:
        digits = digits[:6]
    return (int(digits[0:2], 16), int(digits[2:4], 16), int(digits[4:6], 16))


def parse_alpha(value: object, default: float = 1.0) -> float:
    """解析 0-1 或 0-255 的 alpha；非法值回退 default。"""
    try:
        alpha = float(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return default
    if alpha > 1.0:
        alpha = alpha / 255.0
    return max(0.0, min(1.0, alpha))


def to_hex(color: ColorLike) -> Optional[str]:
    """转为 "#RRGGBB"；解析失败返回 None。"""
    rgb = parse_hex(color)
    if rgb is None:
        return None
    return "#%02X%02X%02X" % rgb


def _linearize(channel: int) -> float:
    """sRGB 通道线性化（WCAG 2.2 定义：<=0.03928 用 12.92 除法，否则 2.4 次幂）。"""
    c = _channel_to_unit(float(channel))
    if c <= 0.03928:
        return c / 12.92
    return ((c + 0.055) / 1.055) ** 2.4


def srgb_luminance(color: ColorLike) -> float:
    """WCAG 相对亮度 L ∈ [0, 1]。

    入参不可解析时返回 0.0（黑的亮度），调用方应先用 parse_hex 判空。
    """
    rgb = parse_hex(color)
    if rgb is None:
        return 0.0
    r, g, b = rgb
    return 0.2126 * _linearize(r) + 0.7152 * _linearize(g) + 0.0722 * _linearize(b)


def contrast_ratio(fg: ColorLike, bg: ColorLike) -> float:
    """WCAG 2.2 SC 1.4.3 对比度，返回 (L亮 + 0.05) / (L暗 + 0.05)，范围 1.0-21.0。

    不区分前景/背景顺序；不可解析的颜色按 0 亮度处理，因此调用方必须先校验 token。
    """
    lum_fg = srgb_luminance(fg)
    lum_bg = srgb_luminance(bg)
    lighter, darker = max(lum_fg, lum_bg), min(lum_fg, lum_bg)
    return (lighter + 0.05) / (darker + 0.05)


def blend_alpha(fg: ColorLike, bg: ColorLike, alpha: float = 1.0) -> RGB:
    """把半透明前景 fg 叠加到不透明背景 bg 上，返回混合后的实色（source-over）。

    alpha 取 0-1；>1 视为 0-255 整数。叠加公式：out = fg*a + bg*(1-a)。
    """
    fg_rgb = parse_hex(fg) or (0, 0, 0)
    bg_rgb = parse_hex(bg) or (255, 255, 255)
    a = parse_alpha(alpha)
    return tuple(_clamp_byte(fg_rgb[i] * a + bg_rgb[i] * (1.0 - a)) for i in range(3))  # type: ignore[return-value]


def composite(layers: Iterable[Tuple[ColorLike, float]], base: ColorLike = "#FFFFFF") -> RGB:
    """自下而上叠加多层半透明颜色，返回最终实色。

    layers 顺序为 从下到上；每项是 (color, alpha)。base 为最底层不透明色。
    """
    result = parse_hex(base) or (255, 255, 255)
    for color, alpha in layers:
        result = blend_alpha(color, result, alpha)
    return result


def is_dark(color: ColorLike, threshold: float = 0.18) -> bool:
    """按 WCAG 相对亮度判断是否"深色"。

    默认阈值 0.18 是 WCAG 定义的"深色背景"分界（相对亮度 < 0.18 需配浅色文字）。
    """
    return srgb_luminance(color) < threshold


def is_light(color: ColorLike, threshold: float = 0.18) -> bool:
    """is_dark 的反向判断。"""
    return not is_dark(color, threshold)


def rgb_to_hsl(color: ColorLike) -> Optional[Tuple[float, float, float]]:
    """转 HSL：h 为 0-360 度，s/l 为 0-1。解析失败返回 None。"""
    rgb = parse_hex(color)
    if rgb is None:
        return None
    r, g, b = (v / 255.0 for v in rgb)
    hi, lo = max(r, g, b), min(r, g, b)
    lightness = (hi + lo) / 2.0
    delta = hi - lo
    if delta == 0:
        return (0.0, 0.0, lightness)
    saturation = delta / (1.0 - abs(2.0 * lightness - 1.0)) if lightness not in (0.0, 1.0) else 0.0
    if hi == r:
        hue = ((g - b) / delta) % 6.0
    elif hi == g:
        hue = (b - r) / delta + 2.0
    else:
        hue = (r - g) / delta + 4.0
    return (hue * 60.0 % 360.0, max(0.0, min(1.0, saturation)), lightness)


def hue_of(color: ColorLike) -> Optional[float]:
    """色相角 0-360 度；灰度色返回 None（无有效色相）。"""
    hsl = rgb_to_hsl(color)
    if hsl is None:
        return None
    hue, saturation, _ = hsl
    if saturation < 0.08:
        return None
    return hue


def saturation_of(color: ColorLike) -> float:
    """HSL 饱和度 0-1；解析失败返回 0.0。"""
    hsl = rgb_to_hsl(color)
    return 0.0 if hsl is None else hsl[1]


def hue_distance(a: ColorLike, b: ColorLike) -> float:
    """两色的色相角距离，取环形最短距离，范围 0-180 度。

    任一色为灰度（饱和度 < 0.08）或不可解析时返回 0.0，表示"无色相差异"。
    """
    hue_a, hue_b = hue_of(a), hue_of(b)
    if hue_a is None or hue_b is None:
        return 0.0
    diff = abs(hue_a - hue_b) % 360.0
    return min(diff, 360.0 - diff)


def lightness_of(color: ColorLike) -> float:
    """HSL 亮度 0-1；解析失败返回 0.0。"""
    hsl = rgb_to_hsl(color)
    return 0.0 if hsl is None else hsl[2]


def rgb_distance(a: ColorLike, b: ColorLike) -> float:
    """sRGB 欧氏距离，范围 0-441.7；用于判断两色"是否肉眼可辨"。"""
    rgb_a, rgb_b = parse_hex(a), parse_hex(b)
    if rgb_a is None or rgb_b is None:
        return 0.0
    return sum((rgb_a[i] - rgb_b[i]) ** 2 for i in range(3)) ** 0.5


# 中文色名：先分有彩色/无彩色，再按色相带与亮度取 深/浅/中性。
_ACHROMATIC_ZH = (
    (0.10, "黑"),
    (0.26, "深灰"),
    (0.46, "灰"),
    (0.70, "中灰"),
    (0.90, "浅灰"),
    (1.01, "白"),
)

_HUE_BANDS_ZH = (
    (15.0, "红"),
    (45.0, "橙"),
    (68.0, "黄"),
    (85.0, "黄绿"),
    (150.0, "绿"),
    (175.0, "青绿"),
    (195.0, "青"),
    (240.0, "蓝"),
    (260.0, "靛蓝"),
    (290.0, "紫"),
    (330.0, "品红"),
    (345.0, "玫红"),
    (360.0, "红"),
)


def color_name_zh(color: ColorLike) -> str:
    """返回中文色名，用于报告里的可读描述，如 "深蓝"、"浅灰"、"橙"。

    判定顺序：HSL 饱和度 < 0.08 视为无彩色，只用 黑/灰/白 系列；
    否则按色相带取基色名，再按亮度前缀 深/浅（亮度 <0.30 加"深"，>0.72 加"浅"）。
    不可解析的颜色返回 "未知"。
    """
    hsl = rgb_to_hsl(color)
    if hsl is None:
        return "未知"
    hue, saturation, lightness = hsl

    if saturation < 0.08:
        for limit, name in _ACHROMATIC_ZH:
            if lightness < limit:
                return name
        return "白"

    base = "红"
    for limit, name in _HUE_BANDS_ZH:
        if hue < limit:
            base = name
            break

    if lightness < 0.30:
        return "深" + base
    if lightness > 0.72:
        return "浅" + base
    return base


def alpha_of(color: ColorLike) -> float:
    """从 (r,g,b,a) 元组里取 alpha；其它形式返回 1.0。"""
    if isinstance(color, (tuple, list)) and len(color) >= 4:
        return parse_alpha(color[3])
    return 1.0


__all__ = [
    "RGB",
    "ColorLike",
    "CONTRAST_BODY",
    "CONTRAST_LARGE",
    "LARGE_TEXT_PT_REGULAR",
    "LARGE_TEXT_PT_BOLD",
    "parse_hex",
    "parse_alpha",
    "to_hex",
    "srgb_luminance",
    "contrast_ratio",
    "blend_alpha",
    "composite",
    "is_dark",
    "is_light",
    "rgb_to_hsl",
    "hue_of",
    "saturation_of",
    "hue_distance",
    "lightness_of",
    "rgb_distance",
    "color_name_zh",
    "alpha_of",
]
