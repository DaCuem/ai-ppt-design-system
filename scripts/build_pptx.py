#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 deck.json 渲染成**原生可编辑** PPTX（唯一入口：deck.json -> .pptx）。

设计约束（决定了本文件为什么这样写）：

* 画面完全由 deck.json 的绝对坐标决定，单位英寸，1 in = 914400 EMU。所有换算走
  `Inches()` / `Pt()`，不写裸魔法数，保证 13.333x7.5 in 与 4:3 画布都能精确落位。
* 文本必须**同时**写 `a:latin` 与 `a:ea`。python-pptx 的 `font.name` 只写 latin，
  在 PowerPoint/LibreOffice 里中文会回退到宋体，与声明的字体不一致——这是本仓库
  所有版面预览「看着差不多但字不对」的根因，所以这里手写 `a:ea` 节点。
* 文本框的内边距必须显式覆盖。python-pptx 的默认 left/right margin 是 0.1 in，
  会让 x/w 精确坐标失效 ±0.1 in（约 1.3% 画布宽），所以先清零再写 `paddingIn`。
* 图表必须是原生 chart（可在 PowerPoint 里改数据），不能是图片。表格同理必须原生。
* 缺图不中断：跳过该元素、记入摘要与退出码，保证还能拿到一份可检查的 pptx。

CLI 契约（以 --help 为准）：

    python scripts/build_pptx.py --deck <deck.json> --out <out.pptx> [--check-only] [--quiet]

    退出码 0 = 成功；1 = deck 校验失败或资源缺失（缺图等）；2 = 参数/文件/依赖问题。
"""

from __future__ import annotations

import argparse
import os
import sys
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence, Tuple

# --- UTF-8 输出（Windows 控制台默认 GBK，中文与 emoji 会抛 UnicodeEncodeError）---
try:
    sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    sys.stderr.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
except Exception:  # pragma: no cover - 非 TTY / 旧 Python
    pass

_SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
if _SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, _SCRIPTS_DIR)

try:
    from lib import deckload  # type: ignore
except ImportError:  # 允许把 lib/ 直接加进 PYTHONPATH 或平铺导入
    import deckload  # type: ignore

try:
    from PIL import Image
except ImportError:  # pragma: no cover
    Image = None  # type: ignore

try:
    from lxml import etree
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData, XyChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
    from pptx.enum.dml import MSO_LINE
    from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
    from pptx.oxml.ns import qn
    from pptx.util import Emu, Inches, Pt
except ImportError as _exc:  # pragma: no cover - 依赖缺失时给出可操作提示
    sys.stderr.write(
        "依赖缺失：无法导入 python-pptx / Pillow / lxml —— 请用本机 Office 运行时自带的解释器运行：\n"
        "  \"D:\\Software\\DSH_Desktop\\DSH Desktop\\resources\\office-runtime\\primary-runtime\\"
        "dependencies\\python\\python.exe\" scripts/build_pptx.py --help\n"
        f"（原始错误：{_exc}）\n"
    )
    raise SystemExit(2)


# ---------------------------------------------------------------------------
# 常量映射
# ---------------------------------------------------------------------------

_ALIGN = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
    "justify": PP_ALIGN.JUSTIFY,
}
_ANCHOR = {
    "top": MSO_ANCHOR.TOP,
    "middle": MSO_ANCHOR.MIDDLE,
    "bottom": MSO_ANCHOR.BOTTOM,
}
_DASH = {
    "solid": MSO_LINE.SOLID,
    "dash": MSO_LINE.DASH,
    "dot": MSO_LINE.ROUND_DOT,
}
_CHART_TYPES = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "area": XL_CHART_TYPE.AREA,
    "pie": XL_CHART_TYPE.PIE,
    "donut": XL_CHART_TYPE.DOUGHNUT,
    "scatter": XL_CHART_TYPE.XY_SCATTER,
}
_AXISLESS_CHARTS = ("pie", "donut")

#: 无框无填充的表格样式：先剥掉主题样式，再逐格上色，避免默认蓝色斑马纹。
_TABLE_STYLE_NONE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"

#: 系列配色取用顺序（只取 tokens.colors 中真实存在的键）。
_SERIES_KEYS = ("primary", "accent", "secondary", "ok", "warn", "fgMuted")

_ICON_NAMES = (
    "circle", "dot", "oval", "square", "rect", "box", "triangle", "diamond",
    "plus", "add", "minus", "remove", "check", "tick", "chevron", "arrow", "bar", "rule",
)


class BuildError(Exception):
    """构建期错误（元素级），消息为中文，由 main 统一转成退出码 1。"""


# ---------------------------------------------------------------------------
# 构建报告
# ---------------------------------------------------------------------------


@dataclass
class BuildReport:
    """构建摘要：页数、逐页元素数、字体、缺图、警告。"""

    deck_path: str = ""
    out_path: str = ""
    out_bytes: int = 0
    page_count: int = 0
    per_page: List[Tuple[str, int]] = field(default_factory=list)
    fonts: List[str] = field(default_factory=list)
    missing_images: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    check_only: bool = False

    def note_font(self, name: str) -> None:
        if name and name not in self.fonts:
            self.fonts.append(name)

    def warn(self, message: str) -> None:
        if message not in self.warnings:
            self.warnings.append(message)

    def render(self) -> str:
        lines: List[str] = []
        lines.append("=== 构建摘要 ===")
        if self.check_only:
            lines.append("模式：--check-only（只校验，未写文件）")
        else:
            lines.append(f"输出文件：{self.out_path}（{self.out_bytes:,} 字节）")
        lines.append(f"页数：{self.page_count}")
        per_page = "  ".join(f"{sid}={count}" for sid, count in self.per_page) or "（无）"
        lines.append(f"逐页元素数：{per_page}")
        lines.append(f"使用字体（{len(self.fonts)}）：{'、'.join(self.fonts) or '（未声明）'}")
        if self.missing_images:
            lines.append(f"缺失图片（{len(self.missing_images)}）：")
            lines.extend(f"  - {p}" for p in self.missing_images)
        else:
            lines.append("缺失图片：无")
        if self.warnings:
            lines.append(f"警告（{len(self.warnings)}）：")
            lines.extend(f"  - {w}" for w in self.warnings)
        return "\n".join(lines)


# ---------------------------------------------------------------------------
# 底层 XML / 颜色工具
# ---------------------------------------------------------------------------


def _rgb(rgb: Tuple[int, int, int]) -> RGBColor:
    return RGBColor(rgb[0], rgb[1], rgb[2])


def _add_fill_alpha(x_pr: Any, opacity: float) -> None:
    """给 xPr（spPr / ln）下已有的 solidFill/srgbClr 写 a:alpha。

    PowerPoint 没有「形状整体透明度」，只有每个填充/线条自己的 alpha，
    因此 opacity 同时作用于填充与描边（这是 OOXML 的模型限制，非实现偷懒）。
    """
    if opacity >= 0.999:
        return
    solid = x_pr.find(qn("a:solidFill"))
    if solid is None:
        return
    clr = solid.find(qn("a:srgbClr"))
    if clr is None:
        return
    alpha = clr.find(qn("a:alpha"))
    if alpha is None:
        alpha = etree.SubElement(clr, qn("a:alpha"))
    alpha.set("val", str(max(0, min(100000, int(round(float(opacity) * 100000))))))


def _no_shadow(shape: Any) -> None:
    """去掉主题带来的投影。

    默认 autoshape 带 `<p:style>` 里的 effectRef，会在页面上投一圈阴影——
    这会让精确的 x/y/w/h 看起来偏移，也会在浅色背景上显脏。显式写空 effectLst。
    """
    try:
        shape.shadow.inherit = False
    except Exception:  # 个别容器（GraphicFrame）没有 shadow
        pass


def _set_ea_font(run: Any, typeface: str) -> None:
    """给 run 写字形属性补上 a:ea（东亚）字体，并保证 latin -> ea 的节点顺序。"""
    r_pr = run._r.get_or_add_rPr()
    latin = r_pr.find(qn("a:latin"))
    ea = r_pr.find(qn("a:ea"))
    if ea is None:
        ea = etree.Element(qn("a:ea"))
    if latin is not None:
        latin.addnext(ea)  # 覆盖已有位置，保证顺序
    else:
        r_pr.append(ea)
    ea.set("typeface", typeface)


def _set_attr(el: Any, name: str, value: str) -> None:
    el.set(name, value)


def _insert_pPr_child(p_pr: Any, child: Any) -> None:
    """按 DrawingML 顺序把 child 插到 a:pPr 的正确位置（buFont/buChar 必须在
    spcBef/spcAft 之后、tabLst/defRPr/extLst 之前）。"""
    for successor in ("a:tabLst", "a:defRPr", "a:extLst"):
        node = p_pr.find(qn(successor))
        if node is not None:
            node.addprevious(child)
            return
    p_pr.append(child)


def _blend(a: Tuple[int, int, int], b: Tuple[int, int, int], t: float) -> Tuple[int, int, int]:
    t = max(0.0, min(1.0, t))
    return tuple(int(round(a[i] * (1 - t) + b[i] * t)) for i in range(3))  # type: ignore[return-value]


def _luminance(rgb: Tuple[int, int, int]) -> float:
    """相对亮度（WCAG 2.2 定义），用于自动挑选表头文字颜色。"""
    channels = []
    for value in rgb:
        c = value / 255.0
        channels.append(c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4)
    return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2]


def _contrast_ratio(a: Tuple[int, int, int], b: Tuple[int, int, int]) -> float:
    la, lb = _luminance(a), _luminance(b)
    lighter, darker = max(la, lb), min(la, lb)
    return (lighter + 0.05) / (darker + 0.05)


# ---------------------------------------------------------------------------
# 构建器
# ---------------------------------------------------------------------------


class DeckBuilder:
    """把一份已校验的 Deck 逐页写成 pptx。状态只有 report，便于复用。"""

    def __init__(self, deck: "deckload.Deck", report: BuildReport):
        self.deck = deck
        self.tokens = deck.tokens
        self.report = report

    # -- 公共 -------------------------------------------------------------
    def build(self) -> Any:
        prs = Presentation()
        prs.slide_width = Inches(self.deck.size.width_in)
        prs.slide_height = Inches(self.deck.size.height_in)

        for slide_spec in self.deck.slides:
            slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白版式
            self._paint_slide_bg(slide, slide_spec)
            elements = slide_spec.get("elements") or []
            self.report.per_page.append((str(slide_spec.get("id", "?")), len(elements)))
            seen_content = False
            for index, element in enumerate(elements):
                path = f"$.slides[{len(self.report.per_page) - 1}].elements[{index}]"
                if str(element.get("zRole") or "content") == "bg" and seen_content:
                    self.report.warn(
                        f"{path}: zRole=bg 的元素排在内容元素之后，会盖住先前内容；"
                        "请把它移到 elements 数组首位（z 序 = 数组顺序）"
                    )
                if str(element.get("zRole") or "content") != "bg":
                    seen_content = True
                self._add_element(slide, element, path)
            self._write_notes(slide, slide_spec)

        self.report.page_count = len(prs.slides)
        return prs

    # -- 页面级 -----------------------------------------------------------
    def _paint_slide_bg(self, slide: Any, slide_spec: Dict[str, Any]) -> None:
        ref = self.deck.slide_bg(slide_spec)
        rgb = self.deck.color(ref, f"slide[{slide_spec.get('id')}].bg")
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = _rgb(rgb)

    def _write_notes(self, slide: Any, slide_spec: Dict[str, Any]) -> None:
        notes = slide_spec.get("notes")
        if notes is None:
            return
        try:
            slide.notes_slide.notes_text_frame.text = str(notes)
        except Exception as exc:  # pragma: no cover - 极少数模板没有 notesMaster
            self.report.warn(f"第 {slide_spec.get('id')} 页备注写入失败：{exc}")

    # -- 元素分发 ---------------------------------------------------------
    def _add_element(self, slide: Any, el: Dict[str, Any], path: str) -> None:
        kind = el.get("kind")
        if kind == "text":
            self._add_text(slide, el, path)
        elif kind in ("rect", "ellipse"):
            self._add_autoshape(slide, el, path, kind)
        elif kind == "line":
            self._add_line(slide, el, path, connector=True)
        elif kind == "polygon":
            self._add_polygon(slide, el, path)
        elif kind == "image":
            self._add_image(slide, el, path)
        elif kind == "chart":
            self._add_chart(slide, el, path)
        elif kind == "table":
            self._add_table(slide, el, path)
        elif kind == "icon":
            self._add_icon(slide, el, path)
        else:  # schema 已枚举，这里防御非校验路径（load_dict 被绕过时）
            raise BuildError(f"{path}.kind: 不支持的元素类型 {kind!r}")

    # -- 几何 -------------------------------------------------------------
    def _box(self, el: Dict[str, Any], path: str) -> Tuple[int, int, int, int]:
        try:
            x = float(el["x"])
            y = float(el["y"])
            w = float(el["w"])
            h = float(el["h"])
        except (KeyError, TypeError, ValueError) as exc:
            raise BuildError(f"{path}: 缺少或非法的几何字段 x/y/w/h（{exc}）") from exc
        if w <= 0 or h <= 0:
            raise BuildError(f"{path}: w={w:g} / h={h:g} 必须大于 0（单位英寸）")
        return Inches(x), Inches(y), Inches(w), Inches(h)

    # -- 文本 -------------------------------------------------------------
    def _add_text(self, slide: Any, el: Dict[str, Any], path: str) -> None:
        x, y, w, h = self._box(el, path)
        shape = slide.shapes.add_textbox(x, y, w, h)
        self._populate_text_frame(shape.text_frame, el, path)

    def _populate_text_frame(self, tf: Any, el: Dict[str, Any], path: str) -> None:
        """把 el 的 text / runs / paragraphs 写进 TextFrame（文本框与自选图形共用）。"""
        tokens = self.tokens
        role = el.get("role") or "body"
        scale = deckload.resolve_type(tokens, role, path=f"{path}.role")

        tf.word_wrap = True
        try:
            tf.auto_size = MSO_AUTO_SIZE.NONE
        except Exception:
            pass

        # 显式覆盖 python-pptx 的默认内边距（0.1 in / 0.05 in），否则精确坐标失效
        pad = el.get("paddingIn", 0.04)
        try:
            pad_in = float(pad) if pad is not None else 0.0
        except (TypeError, ValueError):
            raise BuildError(f"{path}.paddingIn: 必须是数字（英寸），实际是 {pad!r}") from None
        tf.margin_left = Inches(pad_in)
        tf.margin_right = Inches(pad_in)
        tf.margin_top = Inches(pad_in)
        tf.margin_bottom = Inches(pad_in)

        anchor_key = el.get("valign") or "top"
        if anchor_key not in _ANCHOR:
            raise BuildError(f"{path}.valign: 只能是 top/middle/bottom，实际是 {anchor_key!r}")
        tf.vertical_anchor = _ANCHOR[anchor_key]

        base_size = el.get("size", scale["size"])
        base_size = float(base_size) if base_size else float(scale["size"])
        base_weight = int(el.get("weight") or scale["weight"])
        base_color = self.deck.color(el.get("color") or "fg", f"{path}.color")
        base_font = deckload.resolve_font(tokens, el.get("font"))
        line_height = el.get("lineHeight", scale["lineHeight"])
        letter_spacing = el.get("letterSpacing", scale["letterSpacing"])
        caps = bool(el.get("caps", False))
        align_key = el.get("align") or "left"
        if align_key not in _ALIGN:
            raise BuildError(f"{path}.align: 只能是 left/center/right/justify，实际是 {align_key!r}")

        paragraphs = self._paragraph_specs(el, path)
        for i, spec in enumerate(paragraphs):
            if i == 0:
                para = tf.paragraphs[0]
            else:
                para = tf.add_paragraph()

            para_align = spec.get("align") or align_key
            if para_align not in _ALIGN:
                raise BuildError(f"{path}.paragraphs[{i}].align: 非法对齐方式 {para_align!r}")
            para.alignment = _ALIGN[para_align]
            if line_height:
                try:
                    para.line_spacing = float(line_height)
                except (TypeError, ValueError):
                    raise BuildError(f"{path}.lineHeight: 必须是倍数（如 1.35），实际是 {line_height!r}") from None
            space_before = spec.get("spaceBeforePt") or 0
            if space_before:
                para.space_before = Pt(float(space_before))
            level = int(spec.get("level") or 0)
            if spec.get("bullet"):
                self._apply_bullet(para, level)
            if level:
                para.level = min(level, 4)

            runs = spec.get("runs")
            if not runs:
                runs = [{"text": spec.get("text", "")}]
            for run_spec in runs:
                self._write_run(
                    para,
                    run_spec,
                    path,
                    default_size=base_size,
                    default_weight=base_weight,
                    default_color=base_color,
                    default_font=base_font,
                    letter_spacing=letter_spacing,
                    caps=caps,
                )

    def _paragraph_specs(self, el: Dict[str, Any], path: str) -> List[Dict[str, Any]]:
        """归一化三种文本形态为 [{text?, runs?, align?, bullet?, level?, spaceBeforePt?}]。"""
        if el.get("paragraphs"):
            specs = el["paragraphs"]
            if not isinstance(specs, list):
                raise BuildError(f"{path}.paragraphs: 必须是数组")
            return [dict(s) for s in specs if isinstance(s, dict)]

        if el.get("runs"):
            runs = el["runs"]
            if not isinstance(runs, list):
                raise BuildError(f"{path}.runs: 必须是数组")
            return [{"runs": [dict(r) for r in runs if isinstance(r, dict)]}]

        text = el.get("text")
        if text is None:
            text = ""
        if not isinstance(text, str):
            text = str(text)
        # 文本里的换行在 deck 里就是「多段」，不是软换行：这样行距/项目符号才可控。
        return [{"text": line} for line in text.split("\n")]

    def _write_run(
        self,
        para: Any,
        run_spec: Dict[str, Any],
        path: str,
        *,
        default_size: float,
        default_weight: int,
        default_color: Tuple[int, int, int],
        default_font: str,
        letter_spacing: float,
        caps: bool,
    ) -> None:
        run = para.add_run()
        run.text = str(run_spec.get("text", ""))

        size = run_spec.get("size") or default_size
        weight = int(run_spec.get("weight") or default_weight)
        if run_spec.get("color"):
            color = self.deck.color(run_spec["color"], f"{path}.runs.color")
        else:
            color = default_color
        font = deckload.resolve_font(self.tokens, run_spec.get("font") or default_font)

        try:
            run.font.size = Pt(float(size))
        except (TypeError, ValueError):
            raise BuildError(f"{path}: 字号 {size!r} 无法换算为 pt") from None
        # PPTX 对 CJK 只有粗/非粗两档，没有 100-900 的中间字重（已知限制，见 README）
        run.font.bold = weight >= 600
        if run_spec.get("italic"):
            run.font.italic = True
        run.font.color.rgb = _rgb(color)
        run.font.name = font
        _set_ea_font(run, font)

        r_pr = run._r.get_or_add_rPr()
        if letter_spacing:
            # DrawingML 的字符间距单位是 1/100 pt，只能写在 rPr 属性上
            r_pr.set("spc", str(int(round(float(letter_spacing) * 100))))
        if caps:
            r_pr.set("cap", "all")
        self.report.note_font(font)

    def _apply_bullet(self, para: Any, level: int) -> None:
        """写项目符号：marL/indent 控制缩进，buFont+buChar 控制字形。"""
        p_pr = para._p.get_or_add_pPr()
        p_pr.set("marL", str(int(Pt(18 + 18 * level))))
        p_pr.set("indent", str(-int(Pt(18))))
        bu_font = p_pr.find(qn("a:buFont"))
        if bu_font is None:
            bu_font = etree.Element(qn("a:buFont"))
            _insert_pPr_child(p_pr, bu_font)
        bu_font.set("typeface", "Arial")
        bu_char = p_pr.find(qn("a:buChar"))
        if bu_char is None:
            bu_char = etree.Element(qn("a:buChar"))
            _insert_pPr_child(p_pr, bu_char)
        bu_char.set("char", "•")

    # -- 形状 -------------------------------------------------------------
    def _apply_fill(self, shape: Any, el: Dict[str, Any], path: str) -> None:
        opacity = el.get("opacity", 1)
        fill_ref = el.get("fill")
        fill = shape.fill
        if fill_ref is None:
            fill.background()  # 显式 noFill：否则会继承主题 accent1 填充
            return
        rgb = self.deck.color(fill_ref, f"{path}.fill")
        fill.solid()
        fill.fore_color.rgb = _rgb(rgb)
        _add_fill_alpha(fill._xPr, float(opacity))  # type: ignore[attr-defined]

    def _apply_stroke(self, shape: Any, el: Dict[str, Any], path: str, *, default_ref: str = "fg") -> None:
        line = shape.line
        stroke_ref = el.get("stroke")
        if stroke_ref is None:
            line.fill.background()
            return
        rgb = self.deck.color(stroke_ref, f"{path}.stroke")
        line.color.rgb = _rgb(rgb)
        width = el.get("strokeWidthPt", 1)
        line.width = Pt(float(width))
        dash_key = el.get("dash") or "solid"
        if dash_key not in _DASH:
            raise BuildError(f"{path}.dash: 只能是 solid/dash/dot，实际是 {dash_key!r}")
        line.dash_style = _DASH[dash_key]
        _add_fill_alpha(line.fill._xPr, float(el.get("opacity", 1)))  # type: ignore[attr-defined]

    def _add_autoshape(self, slide: Any, el: Dict[str, Any], path: str, kind: str) -> None:
        x, y, w, h = self._box(el, path)
        corners = float(el.get("corners") or 0)
        if kind == "ellipse":
            shape_type = MSO_SHAPE.OVAL
            if corners:
                self.report.warn(f"{path}.corners: ellipse 不支持圆角，已忽略")
        else:
            shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if corners > 0 else MSO_SHAPE.RECTANGLE

        shape = slide.shapes.add_shape(shape_type, x, y, w, h)
        _no_shadow(shape)
        self._apply_fill(shape, el, path)
        self._apply_stroke(shape, el, path)

        if kind == "rect" and corners > 0:
            # roundRect 的 adj 是「圆角半径 / 短边」的比值，上限 0.5（半圆）
            short_side = min(float(el["w"]), float(el["h"]))
            try:
                shape.adjustments[0] = max(0.0, min(0.5, corners / short_side))
            except Exception:
                self.report.warn(f"{path}.corners: 圆角设置失败，已按直角渲染")

        if self._has_text(el):
            # 允许「矩形里带文字」，省掉一次对齐两个元素的坐标同步
            self._populate_text_frame(shape.text_frame, el, path)

    def _has_text(self, el: Dict[str, Any]) -> bool:
        for key in ("text", "runs", "paragraphs"):
            value = el.get(key)
            if value not in (None, "", []):
                return True
        return False

    def _add_line(self, slide: Any, el: Dict[str, Any], path: str, *, connector: bool) -> None:
        x_in, y_in = float(el["x"]), float(el["y"])
        x2 = el.get("x2")
        y2 = el.get("y2")
        x2_in = float(x2) if x2 is not None else x_in + float(el["w"])
        y2_in = float(y2) if y2 is not None else y_in + float(el["h"])
        begin_x, begin_y = Inches(x_in), Inches(y_in)
        end_x, end_y = Inches(x2_in), Inches(y2_in)
        if begin_x == end_x and begin_y == end_y:
            raise BuildError(f"{path}: line 的起点与终点重合（x={x_in:g}, y={y_in:g}）")

        shape = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, begin_x, begin_y, end_x, end_y)
        _no_shadow(shape)
        stroke_ref = el.get("stroke") or el.get("color") or "fg"
        rgb = self.deck.color(stroke_ref, f"{path}.stroke")
        shape.line.color.rgb = _rgb(rgb)
        shape.line.width = Pt(float(el.get("strokeWidthPt", 1)))
        dash_key = el.get("dash") or "solid"
        if dash_key not in _DASH:
            raise BuildError(f"{path}.dash: 只能是 solid/dash/dot，实际是 {dash_key!r}")
        shape.line.dash_style = _DASH[dash_key]
        _add_fill_alpha(shape.line.fill._xPr, float(el.get("opacity", 1)))

    def _freeform(self, slide: Any, points: Sequence[Tuple[float, float]], close: bool,
                  *, fill_ref: Optional[str], stroke_ref: Optional[str], el: Dict[str, Any],
                  path: str) -> Any:
        """按绝对英寸坐标画自由多边形/折线。

        FreeformBuilder 以起点为局部原点，且 convert_to_shape 的 origin 是「局部原点在
        页面的位置」——直接把绝对坐标当局部坐标、origin 传 0，形状就会落在绝对位置。
        （实测：2,2 -> 5,4 的三角形 left/top 正好 2.0/2.0，宽高 3.0/2.0）
        """
        start_x, start_y = Inches(points[0][0]), Inches(points[0][1])
        builder = slide.shapes.build_freeform(start_x, start_y, scale=1.0)
        builder.add_line_segments([(Inches(px), Inches(py)) for px, py in points[1:]], close=close)
        shape = builder.convert_to_shape()
        _no_shadow(shape)
        if fill_ref:
            shape.fill.solid()
            shape.fill.fore_color.rgb = _rgb(self.deck.color(fill_ref, f"{path}.fill"))
            _add_fill_alpha(shape.fill._xPr, float(el.get("opacity", 1)))  # type: ignore[attr-defined]
        else:
            shape.fill.background()
        if stroke_ref:
            shape.line.color.rgb = _rgb(self.deck.color(stroke_ref, f"{path}.stroke"))
            shape.line.width = Pt(float(el.get("strokeWidthPt", 1)))
        else:
            shape.line.fill.background()
        return shape

    def _add_polygon(self, slide: Any, el: Dict[str, Any], path: str) -> None:
        raw_points = el.get("points") or []
        if len(raw_points) < 3:
            raise BuildError(f"{path}.points: polygon 至少需要 3 个顶点，实际 {len(raw_points)} 个")
        points: List[Tuple[float, float]] = []
        for i, point in enumerate(raw_points):
            try:
                points.append((float(point["x"]), float(point["y"])))
            except (KeyError, TypeError, ValueError):
                raise BuildError(f"{path}.points[{i}]: 顶点必须是 {{\"x\": 数字, \"y\": 数字}}") from None
        self._freeform(
            slide, points, True,
            fill_ref=el.get("fill") or None,
            stroke_ref=el.get("stroke") or None,
            el=el, path=path,
        )

    # -- 图片 -------------------------------------------------------------
    def _add_image(self, slide: Any, el: Dict[str, Any], path: str) -> None:
        src = el.get("src")
        if not src:
            raise BuildError(f"{path}.src: image 元素缺少 src")
        abs_path = src if os.path.isabs(src) else os.path.join(self.deck.base_dir, src)
        abs_path = os.path.normpath(abs_path)
        if not os.path.isfile(abs_path):
            message = f"图片不存在：{abs_path}（{path}.src）"
            self.report.missing_images.append(abs_path)
            self.report.warn(message)
            return
        if Image is None:
            raise BuildError("Pillow 不可用，无法读取图片尺寸/裁切")
        try:
            with Image.open(abs_path) as image:
                src_w, src_h = image.size
        except Exception as exc:
            self.report.missing_images.append(f"{abs_path}（无法解码：{exc}）")
            self.report.warn(f"{path}.src: 图片无法解码，已跳过：{abs_path}（{exc}）")
            return
        if src_w <= 0 or src_h <= 0:
            self.report.missing_images.append(abs_path)
            self.report.warn(f"{path}.src: 图片尺寸非法 {src_w}x{src_h}，已跳过")
            return

        x_in, y_in = float(el["x"]), float(el["y"])
        w_in, h_in = float(el["w"]), float(el["h"])
        fit = el.get("fit") or "cover"
        if fit not in ("cover", "contain", "fill"):
            raise BuildError(f"{path}.fit: 只能是 cover/contain/fill，实际是 {fit!r}")

        crop = el.get("crop") or {}
        left, top, right, bottom = self._crop_fractions(crop, path)

        # 显式 crop 先扣掉，再按剩余区域做 cover/contain，避免两套裁切互相打架
        eff_w = src_w * (1 - left - right)
        eff_h = src_h * (1 - top - bottom)
        if eff_w <= 0 or eff_h <= 0:
            raise BuildError(f"{path}.crop: 裁切后没有剩余像素（l+r 或 t+b 达到 100%）")

        if fit == "contain":
            scale = min(w_in / (eff_w / 96.0), h_in / (eff_h / 96.0))
            draw_w = (eff_w / 96.0) * scale
            draw_h = (eff_h / 96.0) * scale
            draw_x = x_in + (w_in - draw_w) / 2
            draw_y = y_in + (h_in - draw_h) / 2
        elif fit == "cover":
            # 目标：让「可见源区域」的宽高比等于 box 的宽高比，再缩放到 box。
            # need_w = 保持全部有效高度时所需的源宽度（px）；多出来的部分两侧各裁一半。
            target_ratio = w_in / h_in
            need_w = eff_h * target_ratio
            if need_w <= eff_w:
                # 源图偏宽：裁掉左右。总横向裁切比例 = 1 - need_w/src_w
                extra = (1 - need_w / src_w) - (left + right)
                if extra > 1e-6:
                    left += extra / 2
                    right += extra / 2
            else:
                # 源图偏高：裁掉上下。need_h = 保持全部有效宽度时所需的源高度（px）
                need_h = eff_w / target_ratio
                extra = (1 - need_h / src_h) - (top + bottom)
                if extra > 1e-6:
                    top += extra / 2
                    bottom += extra / 2
            draw_x, draw_y, draw_w, draw_h = x_in, y_in, w_in, h_in
        else:  # fill：拉伸填满，不裁切（唯一允许变形的模式，需在 deck 里显式声明）
            draw_x, draw_y, draw_w, draw_h = x_in, y_in, w_in, h_in

        picture = slide.shapes.add_picture(abs_path, Inches(draw_x), Inches(draw_y), Inches(draw_w), Inches(draw_h))
        # python-pptx 的 crop_* 是「源图被裁掉的百分比」，与上面的比例定义一致
        picture.crop_left = max(0.0, min(1.0, left))
        picture.crop_right = max(0.0, min(1.0, right))
        picture.crop_top = max(0.0, min(1.0, top))
        picture.crop_bottom = max(0.0, min(1.0, bottom))
        _no_shadow(picture)
        alt = el.get("alt")
        if alt:
            try:
                picture._element.nvPicPr.cNvPr.set("descr", str(alt))
            except Exception:  # pragma: no cover
                pass

    def _crop_fractions(self, crop: Dict[str, Any], path: str) -> Tuple[float, float, float, float]:
        values: List[float] = []
        for key in ("l", "t", "r", "b"):
            raw = crop.get(key, 0) if isinstance(crop, dict) else 0
            try:
                value = float(raw or 0)
            except (TypeError, ValueError):
                raise BuildError(f"{path}.crop.{key}: 必须是 0-1 的小数，实际是 {raw!r}") from None
            if not -0.001 <= value <= 1.0001:
                raise BuildError(f"{path}.crop.{key}: 必须落在 0-1 区间，实际是 {value:g}")
            values.append(max(0.0, min(1.0, value)))
        left, top, right, bottom = values
        if left + right >= 0.999:
            raise BuildError(f"{path}.crop: l+r 达到 100%，没有剩余画面")
        if top + bottom >= 0.999:
            raise BuildError(f"{path}.crop: t+b 达到 100%，没有剩余画面")
        return left, top, right, bottom

    # -- 图表 -------------------------------------------------------------
    def _series_color(self, index: int) -> Tuple[int, int, int]:
        colors = self.tokens.get("colors") or {}
        keys = [key for key in _SERIES_KEYS if key in colors]
        if keys:
            return self.deck.color(keys[index % len(keys)], f"chart.series[{index}]")
        return (0x4A, 0x55, 0x68)

    def _add_chart(self, slide: Any, el: Dict[str, Any], path: str) -> None:
        x, y, w, h = self._box(el, path)
        spec = el.get("chart") or {}
        chart_type_key = el.get("chartType") or "column"
        if chart_type_key not in _CHART_TYPES:
            raise BuildError(
                f"{path}.chartType: 不支持 {chart_type_key!r}（可用：{'、'.join(_CHART_TYPES)}）"
            )
        series_specs = spec.get("series") or []
        if not series_specs:
            raise BuildError(f"{path}.chart.series: 至少需要 1 个系列")
        categories = [str(c) for c in (spec.get("categories") or [])]
        if chart_type_key in _AXISLESS_CHARTS and len(series_specs) > 1:
            raise BuildError(
                f"{path}.chart.series: {chart_type_key} 只支持 1 个系列，实际给了 {len(series_specs)} 个；"
                "多系列请改用 column/bar/line"
            )

        if chart_type_key == "scatter":
            chart_data: Any = XyChartData()
            for si, series in enumerate(series_specs):
                name = str(series.get("name") or f"系列 {si + 1}")
                values = [float(v) for v in (series.get("values") or [])]
                xs = self._scatter_x(categories, len(values))
                xy_series = chart_data.add_series(name)
                for xi, yi in zip(xs, values):
                    xy_series.add_data_point(xi, yi)
        else:
            chart_data = CategoryChartData()
            chart_data.categories = categories or [str(i + 1) for i in range(
                max(len([v for v in (s.get("values") or [])]) for s in series_specs)
            )]
            if not categories:
                self.report.warn(f"{path}.chart.categories: 未提供，已用序号 1..N 代替")
            for si, series in enumerate(series_specs):
                values = [float(v) for v in (series.get("values") or [])]
                if len(values) != len(chart_data.categories):
                    raise BuildError(
                        f"{path}.chart.series[{si}].values: 有 {len(values)} 个值，"
                        f"但 categories 有 {len(chart_data.categories)} 个；两者必须等长"
                    )
                chart_data.add_series(str(series.get("name") or f"系列 {si + 1}"), values)

        graphic_frame = slide.shapes.add_chart(_CHART_TYPES[chart_type_key], x, y, w, h, chart_data)
        chart = graphic_frame.chart
        self._style_chart(chart, el, spec, path, chart_type_key, len(series_specs))
        self._add_chart_caption(slide, el, spec, path, float(el["y"]) + float(el["h"]))

    def _scatter_x(self, categories: List[str], count: int) -> List[float]:
        """散点图的 X：categories 若能全部转成数字就用它，否则用 1..N。"""
        if categories and len(categories) == count:
            try:
                return [float(c) for c in categories]
            except (TypeError, ValueError):
                pass
        return [float(i + 1) for i in range(count)]

    def _style_chart(
        self,
        chart: Any,
        el: Dict[str, Any],
        spec: Dict[str, Any],
        path: str,
        chart_type_key: str,
        series_count: int,
    ) -> None:
        body_font = deckload.resolve_font(self.tokens, (self.tokens.get("fonts") or {}).get("body"))
        fg = self.deck.color("fg", f"{path}.chart")
        self.report.note_font(body_font)

        chart.has_title = False  # 否则 python-pptx 会用系列名自动生成一个标题
        chart.font.name = body_font
        chart.font.size = Pt(12)
        chart.font.color.rgb = _rgb(fg)
        self._style_chart_ea(chart, body_font)

        show_legend = bool(spec.get("showLegend", True))
        chart.has_legend = show_legend and series_count > 0
        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
            chart.legend.font.size = Pt(11)

        number_format = str(spec.get("numberFormat") or "General")
        plot = chart.plots[0]
        if chart_type_key in ("column", "bar"):
            try:
                plot.gap_width = 60
                plot.overlap = -8
            except Exception:
                pass

        series_list = list(plot.series)
        for index, series in enumerate(series_list):
            rgb = self._series_color(index)
            series_spec = (spec.get("series") or [])[index] if index < len(spec.get("series") or []) else {}
            if series_spec.get("color"):
                rgb = self.deck.color(series_spec["color"], f"{path}.chart.series[{index}].color")
            try:
                if chart_type_key in ("line", "scatter"):
                    series.format.line.color.rgb = _rgb(rgb)
                    series.format.line.width = Pt(2)
                    if chart_type_key == "line":
                        series.smooth = False
                else:
                    series.format.fill.solid()
                    series.format.fill.fore_color.rgb = _rgb(rgb)
            except Exception:
                pass

            if chart_type_key in _AXISLESS_CHARTS:
                # 饼图/环形图要按「点」上色，系列只是容器
                try:
                    for point_index in range(len(series.points)):
                        point = series.points[point_index]
                        point_rgb = self._series_color(point_index)
                        point.format.fill.solid()
                        point.format.fill.fore_color.rgb = _rgb(point_rgb)
                except Exception:
                    pass

        if spec.get("showValues"):
            try:
                plot.has_data_labels = True
                labels = plot.data_labels
                labels.show_value = True
                labels.number_format = number_format
                labels.number_format_is_linked = False
                labels.font.size = Pt(11)
                labels.font.color.rgb = _rgb(fg)
                if chart_type_key in ("column", "bar", "line"):
                    labels.position = XL_LABEL_POSITION.OUTSIDE_END
                elif chart_type_key in _AXISLESS_CHARTS:
                    labels.position = XL_LABEL_POSITION.CENTER
            except Exception as exc:  # pragma: no cover - 个别图表类型不支持数据标签定位
                self.report.warn(f"{path}.chart.showValues: 数据标签设置失败（{exc}）")

        if chart_type_key not in _AXISLESS_CHARTS:
            try:
                value_axis = chart.value_axis
                value_axis.has_major_gridlines = False
                value_axis.tick_labels.number_format = number_format
                value_axis.tick_labels.number_format_is_linked = False
                value_axis.tick_labels.font.size = Pt(11)
                value_axis.format.line.fill.background()
                category_axis = chart.category_axis
                category_axis.tick_labels.font.size = Pt(11)
                category_axis.format.line.color.rgb = _rgb(fg)
            except Exception as exc:  # pragma: no cover
                self.report.warn(f"{path}.chart: 坐标轴样式设置失败（{exc}）")

    def _style_chart_ea(self, chart: Any, font_name: str) -> None:
        """图表内文字（坐标轴、图例、标签）同样需要 a:ea，否则中文回退宋体。"""
        try:
            chart_space = chart._chartSpace
            tx_pr = chart_space.find(qn("c:txPr"))
            if tx_pr is None:
                chart_space.get_or_add_txPr()
                tx_pr = chart_space.find(qn("c:txPr"))
            if tx_pr is None:
                return
            for def_r_pr in tx_pr.iter(qn("a:defRPr")):
                latin = def_r_pr.find(qn("a:latin"))
                ea = def_r_pr.find(qn("a:ea"))
                if ea is None:
                    ea = etree.Element(qn("a:ea"))
                if latin is not None:
                    latin.addnext(ea)
                else:
                    def_r_pr.append(ea)
                ea.set("typeface", font_name)
        except Exception:  # pragma: no cover - 图表字体级联失败不影响主体
            pass

    def _add_chart_caption(self, slide: Any, el: Dict[str, Any], spec: Dict[str, Any],
                           path: str, top_in: float) -> None:
        """单位与数据来源必须显示在图表下方（禁止无单位数字与无来源数据）。"""
        unit = spec.get("unit")
        source = spec.get("source")
        if not unit and not source:
            return
        parts = []
        if unit:
            parts.append(f"单位：{unit}")
        if source:
            parts.append(f"来源：{source}")
        caption = " ｜ ".join(parts)

        scale = deckload.resolve_type(self.tokens, "caption")
        muted_ref = "fgMuted" if "fgMuted" in (self.tokens.get("colors") or {}) else "fg"
        box = slide.shapes.add_textbox(
            Inches(float(el["x"])), Inches(top_in + 0.04), Inches(float(el["w"])), Inches(0.3)
        )
        tf = box.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        for attr in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
            setattr(tf, attr, 0)
        tf.vertical_anchor = MSO_ANCHOR.TOP
        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.LEFT
        run = para.add_run()
        run.text = caption
        run.font.size = Pt(float(scale["size"]))
        body_font = deckload.resolve_font(self.tokens, (self.tokens.get("fonts") or {}).get("body"))
        run.font.name = body_font
        _set_ea_font(run, body_font)
        run.font.color.rgb = _rgb(self.deck.color(muted_ref, f"{path}.chart.caption"))
        self.report.note_font(body_font)

    # -- 表格 -------------------------------------------------------------
    def _add_table(self, slide: Any, el: Dict[str, Any], path: str) -> None:
        x, y, w, h = self._box(el, path)
        spec = el.get("table") or {}
        header = [str(c) for c in (spec.get("header") or [])]
        rows = [[str(c) for c in row] for row in (spec.get("rows") or [])]
        if not rows and not header:
            raise BuildError(f"{path}.table: 既没有 header 也没有 rows，无法建表")

        cols = max([len(header)] + [len(row) for row in rows]) if (header or rows) else 0
        if cols <= 0:
            raise BuildError(f"{path}.table: 列数为 0")
        for index, row in enumerate(rows):
            if len(row) < cols:
                rows[index] = row + [""] * (cols - len(row))
                self.report.warn(f"{path}.table.rows[{index}]: 只有 {len(row)} 列，少于表头 {cols} 列，已补空")
        header = header + [""] * (cols - len(header))

        total_rows = len(rows) + (1 if header and any(header) else 0)
        graphic_frame = slide.shapes.add_table(total_rows, cols, x, y, w, h)
        table = graphic_frame.table
        _no_shadow(graphic_frame)
        self._set_table_plain_style(table, graphic_frame)

        # 列宽：给了 colWidths 就按比例缩放到元素宽 w，保证不溢出
        widths = spec.get("colWidths") or []
        widths = [float(v) for v in widths if isinstance(v, (int, float))]
        if len(widths) >= cols:
            scale = float(el["w"]) / sum(widths[:cols]) if sum(widths[:cols]) else 1.0
            for index in range(cols):
                table.columns[index].width = Inches(widths[index] * scale)
        else:
            each = float(el["w"]) / cols
            for index in range(cols):
                table.columns[index].width = Inches(each)

        row_height = float(el["h"]) / total_rows
        for index in range(total_rows):
            table.rows[index].height = Inches(row_height)

        tokens = self.tokens
        fg = self.deck.color("fg", f"{path}.table")
        bg = self.deck.color(self.deck.slide_bg({}), f"{path}.table.bg")
        bg_alt = self.deck.color("bgAlt", f"{path}.table.bgAlt") if "bgAlt" in (tokens.get("colors") or {}) \
            else _blend(bg, fg, 0.06)
        header_fill_ref = spec.get("headerFill") or ("primary" if "primary" in (tokens.get("colors") or {}) else "fg")
        header_fill = self.deck.color(header_fill_ref, f"{path}.table.headerFill")
        header_text = bg if _contrast_ratio(header_fill, bg) >= _contrast_ratio(header_fill, fg) else fg
        accent = self.deck.color("accent", f"{path}.table.accent") if "accent" in (tokens.get("colors") or {}) else header_fill

        font_size = float(spec.get("fontSizePt") or 14)
        align_key = spec.get("align") or "left"
        if align_key not in ("left", "center", "right"):
            raise BuildError(f"{path}.table.align: 只能是 left/center/right，实际是 {align_key!r}")
        alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}[align_key]
        zebra = bool(spec.get("zebra", False))
        emphasis = spec.get("emphasisCol")
        emphasis_col = int(emphasis) if isinstance(emphasis, (int, float)) and not isinstance(emphasis, bool) else -1
        body_font = deckload.resolve_font(tokens, (tokens.get("fonts") or {}).get("body"))
        self.report.note_font(body_font)

        has_header = bool(header and any(header))
        for row_index in range(total_rows):
            is_header = has_header and row_index == 0
            fill_rgb = header_fill if is_header else (bg_alt if zebra and (row_index % 2 == 0) else bg)
            for col_index in range(cols):
                cell = table.cell(row_index, col_index)
                cell.fill.solid()
                cell.fill.fore_color.rgb = _rgb(fill_rgb)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.margin_left = Inches(0.08)
                cell.margin_right = Inches(0.08)
                cell.margin_top = Inches(0.02)
                cell.margin_bottom = Inches(0.02)

                text = header[col_index] if is_header else rows[row_index - (1 if has_header else 0)][col_index]
                tf = cell.text_frame
                tf.word_wrap = True
                para = tf.paragraphs[0]
                para.alignment = alignment
                run = para.add_run()
                run.text = text
                run.font.size = Pt(font_size)
                run.font.bold = is_header
                if is_header:
                    run.font.color.rgb = _rgb(header_text)
                elif emphasis_col == col_index:
                    run.font.color.rgb = _rgb(accent)
                    run.font.bold = True
                else:
                    run.font.color.rgb = _rgb(fg)
                run.font.name = body_font
                _set_ea_font(run, body_font)

    def _set_table_plain_style(self, table: Any, graphic_frame: Any) -> None:
        """剥掉主题表格样式（默认的蓝底斑马纹会盖住自定义配色）。"""
        try:
            table.first_row = False
            table.horz_banding = False
            tbl = graphic_frame._element.find(".//" + qn("a:tbl"))
            if tbl is None:
                return
            tbl_pr = tbl.find(qn("a:tblPr"))
            if tbl_pr is None:
                return
            style = tbl_pr.find(qn("a:tableStyleId"))
            if style is not None:
                style.text = _TABLE_STYLE_NONE
        except Exception:  # pragma: no cover
            pass

    # -- 图标（纯几何，不依赖图标字体） ------------------------------------
    def _add_icon(self, slide: Any, el: Dict[str, Any], path: str) -> None:
        name = str(el.get("icon") or "").strip().lower()
        x_in, y_in = float(el["x"]), float(el["y"])
        w_in, h_in = float(el["w"]), float(el["h"])
        if name not in _ICON_NAMES:
            self.report.warn(
                f"{path}.icon: 不支持的图标 {name!r}，已跳过（可用：{'、'.join(sorted(set(_ICON_NAMES)))}）"
            )
            return
        color_ref = el.get("stroke") or el.get("color") or el.get("fill") or "fg"
        rgb = self.deck.color(color_ref, f"{path}.icon")
        width = Pt(float(el.get("strokeWidthPt", 1.5)))

        if name in ("circle", "dot", "oval"):
            shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x_in), Inches(y_in), Inches(w_in), Inches(h_in))
            shape.fill.solid()
            shape.fill.fore_color.rgb = _rgb(rgb)
            shape.line.fill.background()
            _no_shadow(shape)
            return
        if name in ("square", "rect", "box"):
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x_in), Inches(y_in), Inches(w_in), Inches(h_in))
            shape.fill.solid()
            shape.fill.fore_color.rgb = _rgb(rgb)
            shape.line.fill.background()
            _no_shadow(shape)
            return
        if name in ("triangle", "diamond"):
            if name == "triangle":
                points = [(x_in + w_in / 2, y_in), (x_in + w_in, y_in + h_in), (x_in, y_in + h_in)]
            else:
                points = [
                    (x_in + w_in / 2, y_in), (x_in + w_in, y_in + h_in / 2),
                    (x_in + w_in / 2, y_in + h_in), (x_in, y_in + h_in / 2),
                ]
            self._freeform(slide, points, True, fill_ref=deckload.rgb_to_hex(rgb), stroke_ref=None, el=el, path=path)
            return
        if name in ("plus", "add", "minus", "remove"):
            self._add_line(slide, {"x": x_in, "y": y_in + h_in / 2, "w": w_in, "h": 0,
                                   "x2": x_in + w_in, "y2": y_in + h_in / 2,
                                   "stroke": deckload.rgb_to_hex(rgb), "strokeWidthPt": width.pt}, path)
            if name in ("plus", "add"):
                self._add_line(slide, {"x": x_in + w_in / 2, "y": y_in, "w": 0, "h": h_in,
                                       "x2": x_in + w_in / 2, "y2": y_in + h_in,
                                       "stroke": deckload.rgb_to_hex(rgb), "strokeWidthPt": width.pt}, path)
            return
        if name in ("chevron", "arrow"):
            points = [
                (x_in, y_in + h_in * 0.18), (x_in + w_in * 0.55, y_in + h_in * 0.18),
                (x_in + w_in, y_in + h_in * 0.5), (x_in + w_in * 0.55, y_in + h_in * 0.82),
                (x_in, y_in + h_in * 0.82), (x_in + w_in * 0.45, y_in + h_in * 0.5),
            ]
            self._freeform(slide, points, True, fill_ref=deckload.rgb_to_hex(rgb), stroke_ref=None, el=el, path=path)
            return
        if name in ("check", "tick"):
            points = [
                (x_in, y_in + h_in * 0.55), (x_in + w_in * 0.38, y_in + h_in * 0.92),
                (x_in + w_in, y_in + h_in * 0.12),
            ]
            self._freeform(slide, points, False, fill_ref=None, stroke_ref=deckload.rgb_to_hex(rgb), el=el, path=path)
            return
        # bar / rule：一条水平线
        self._add_line(slide, {"x": x_in, "y": y_in + h_in / 2, "w": w_in, "h": 0,
                               "x2": x_in + w_in, "y2": y_in + h_in / 2,
                               "stroke": deckload.rgb_to_hex(rgb), "strokeWidthPt": width.pt}, path)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="build_pptx.py",
        description="把 deck.json 构建成原生可编辑 PPTX（唯一入口）。返回 0 成功 / 1 校验或资源失败 / 2 参数或依赖错误。",
        epilog=(
            "示例：\n"
            "  python scripts/build_pptx.py --deck demo/smoke/deck.json --out demo/smoke/smoke.pptx\n"
            "  python scripts/build_pptx.py --deck deck.json --check-only\n"
            "说明：缺图不中断构建，但退出码为 1；--check-only 会在内存里完整构建一遍再丢弃，\n"
            "      因此它校验的是真实构建路径（含颜色令牌/字号回退/图表数据长度）。"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--deck", required=True, help="输入 deck.json 路径（相对或绝对）")
    parser.add_argument("--out", help="输出 .pptx 路径；--check-only 时可省略")
    parser.add_argument("--check-only", action="store_true", help="只校验（内存构建，不写文件）")
    parser.add_argument("--quiet", action="store_true", help="只打印一行结果，不打印逐页摘要")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)

    if not args.check_only and not args.out:
        sys.stderr.write("错误：缺少 --out（除非使用 --check-only）。\n")
        return 2
    if not os.path.isfile(args.deck):
        sys.stderr.write(f"错误：找不到 deck 文件：{os.path.abspath(args.deck)}\n")
        return 2

    try:
        deck = deckload.load(args.deck)
    except deckload.DeckValidationError as exc:
        sys.stderr.write(f"校验失败：{exc}\n")
        return 1
    except deckload.DeckError as exc:
        sys.stderr.write(f"读取失败：{exc}\n")
        return 1

    report = BuildReport(deck_path=os.path.abspath(args.deck), check_only=args.check_only)
    for warning in deck.warnings:
        report.warn(warning)

    try:
        builder = DeckBuilder(deck, report)
        presentation = builder.build()
    except BuildError as exc:
        sys.stderr.write(f"构建失败：{exc}\n")
        return 1
    except Exception as exc:  # pragma: no cover - 未预期错误也要给出可读信息
        import traceback

        sys.stderr.write(f"构建异常（这是实现缺陷，请连同下面的栈一起报）：{exc}\n")
        traceback.print_exc()
        return 1

    if not report.missing_images:
        # 只保留真正影响输出的缺图；deckload 已把路径写进 missing_images
        report.missing_images = list(deck.missing_images)

    if args.check_only:
        print(report.render())
        if report.missing_images:
            sys.stderr.write(f"错误：{len(report.missing_images)} 张图片缺失，无法通过校验。\n")
            return 1
        print("校验通过：deck 结构与资源均可用。")
        return 0

    out_path = os.path.abspath(args.out)
    out_dir = os.path.dirname(out_path)
    if out_dir:
        try:
            os.makedirs(out_dir, exist_ok=True)
        except OSError as exc:
            sys.stderr.write(f"错误：无法创建输出目录 {out_dir}（{exc}）\n")
            return 1
    try:
        presentation.save(out_path)
    except OSError as exc:
        sys.stderr.write(f"错误：写入 pptx 失败：{out_path}（{exc}）\n")
        return 1
    report.out_path = out_path
    try:
        report.out_bytes = os.path.getsize(out_path)
    except OSError:
        report.out_bytes = 0

    if args.quiet:
        print(f"{out_path} | {report.page_count} 页 | 缺图 {len(report.missing_images)} | 警告 {len(report.warnings)}")
    else:
        print(report.render())

    if report.missing_images:
        sys.stderr.write(
            f"错误：{len(report.missing_images)} 张图片缺失，PPTX 已生成但对应元素为空。"
            "补齐 src 指向的文件后重新构建。\n"
        )
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
