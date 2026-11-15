#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""contact_sheet.py —— 把一页一张的 PNG 预览拼成一张总览图（contact sheet）。

用途：渲染出全部页面 PNG 后，用一张图快速检查整套 deck 的节奏与版式分布
（是否连续三页同一种版式、章节页有没有视觉断点、页码/页脚是否每页都在）。

示例：
    python scripts/contact_sheet.py --dir previews --out contact-sheet.png
    python scripts/contact_sheet.py --dir previews --out contact-sheet.png --cols 3 --width 560
    python scripts/contact_sheet.py --dir previews --out sheet.png --cols 4 --title "季度汇报 v3"

每个格子：统一缩放到同一宽度、等比缩放、格间距 12 px；
左上角覆盖"序号 + 页名"标签块；底色按预览图整体明暗自动选深底/浅底。
只用标准库 + Pillow，不联网。
"""

from __future__ import annotations

import argparse
import math
import os
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

try:
    from PIL import Image, ImageDraw, ImageFont, ImageOps
except ImportError:  # pragma: no cover - 环境缺 Pillow 时给出可执行提示
    print("ERROR 缺少 Pillow，无法生成总览图。请使用仓库指定的 Python 运行时（自带 Pillow）。", file=sys.stderr)
    raise SystemExit(2)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

IMAGE_EXTS = (".png", ".jpg", ".jpeg", ".webp")
DEFAULT_GAP = 12
DEFAULT_COLS = 3
DEFAULT_CELL_WIDTH = 560
LIGHT_BG = (246, 246, 244)
DARK_BG = (18, 20, 22)
LIGHT_INK = (26, 28, 30)
DARK_INK = (238, 238, 236)

# 本机可用的中文字体（与 scripts/lib/metrics.py 的字体表保持一致的回退顺序）
_FONT_CANDIDATES = (
    (r"C:\Windows\Fonts\msyh.ttc", 0),
    (r"C:\Windows\Fonts\msyhbd.ttc", 0),
    (r"C:\Windows\Fonts\Deng.ttf", 0),
    (r"C:\Windows\Fonts\simhei.ttf", 0),
    (r"C:\Windows\Fonts\arial.ttf", 0),
)


def _natural_key(name: str) -> Tuple:
    """自然排序键：slide-2.png 排在 slide-10.png 前面。"""
    parts = re.split(r"(\d+)", name.lower())
    return tuple(int(part) if part.isdigit() else part for part in parts)


def _font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    """取一个能显示中文的字体；找不到就用 Pillow 默认字体。"""
    candidates = _FONT_CANDIDATES[1:] + _FONT_CANDIDATES[:1] if bold else _FONT_CANDIDATES
    for path, index in candidates:
        if not os.path.isfile(path):
            continue
        try:
            return ImageFont.truetype(path, size=size, index=index)
        except OSError:
            continue
    try:
        return ImageFont.load_default(size=size)
    except TypeError:  # Pillow < 10.1 的 load_default 没有 size 参数
        return ImageFont.load_default()


def collect_images(folder: Path, recursive: bool = False) -> List[Path]:
    """列出目录内的预览图，按页名自然排序。"""
    if not folder.is_dir():
        raise FileNotFoundError("找不到预览目录：%s" % folder)
    pattern = "**/*" if recursive else "*"
    files = [
        path
        for path in folder.glob(pattern)
        if path.is_file() and path.suffix.lower() in IMAGE_EXTS and not path.name.startswith("contact-sheet")
    ]
    files.sort(key=lambda p: _natural_key(p.name))
    return files


def _mean_luminance(image: Image.Image) -> float:
    """缩到 16×16 灰度取均值，判断整体明暗（0 黑 … 1 白）。"""
    small = image.convert("L").resize((16, 16), Image.BILINEAR)
    pixels = small.tobytes()  # 灰度图每像素 1 字节
    return (sum(pixels) / float(len(pixels))) / 255.0


def _prepare(image_path: Path, cell_width: int) -> Tuple[Image.Image, float]:
    """载入并等比缩放到统一宽度，返回 (图像, 平均亮度)。"""
    with Image.open(image_path) as raw:
        image = ImageOps.exif_transpose(raw)
        image = image.convert("RGB")
        if image.width <= 0 or image.height <= 0:
            raise ValueError("图像尺寸非法：%s" % image_path)
        scale = cell_width / float(image.width)
        height = max(1, int(round(image.height * scale)))
        resized = image.resize((cell_width, height), Image.LANCZOS)
    return resized, _mean_luminance(resized)


def _page_label(path: Path) -> str:
    """页名：去掉扩展名，去掉常见预览前缀（slide-01 → 01）。"""
    stem = path.stem
    return stem


def _draw_badge(
    canvas: Image.Image,
    draw: ImageDraw.ImageDraw,
    origin: Tuple[int, int],
    index: int,
    label: str,
    ink: Tuple[int, int, int],
    badge_size: int,
) -> None:
    """在单元左上角画"序号 + 页名"标签块（半透明圆角底）。"""
    x, y = origin
    font = _font(badge_size, bold=True)
    text = "%02d" % index
    number_w = draw.textlength(text, font=font)
    gap = max(4, int(badge_size * 0.4))
    name_w = draw.textlength(label, font=font)
    padding_x = max(6, int(badge_size * 0.5))
    padding_y = max(3, int(badge_size * 0.28))
    box_w = int(padding_x * 2 + number_w + gap + name_w)
    box_h = int(badge_size + padding_y * 2)
    radius = max(3, int(box_h * 0.28))
    overlay = Image.new("RGBA", (max(1, box_w), max(1, box_h)), (0, 0, 0, 0))
    odraw = ImageDraw.Draw(overlay)
    # 浅色字（深底）→ 深色标签板；深色字（浅底）→ 浅色标签板
    plate = (14, 16, 18, 214) if sum(ink) > 380 else (255, 255, 255, 224)
    odraw.rounded_rectangle([(0, 0), (box_w - 1, box_h - 1)], radius=radius, fill=plate)
    odraw.text((padding_x, padding_y - 1), text, font=font, fill=ink + (255,))
    odraw.text((padding_x + number_w + gap, padding_y - 1), label, font=font, fill=ink + (255,))
    canvas.paste(overlay, (x, y), overlay)


def build_contact_sheet(
    image_paths: Sequence[Path],
    out_path: Path,
    cols: int = DEFAULT_COLS,
    cell_width: int = DEFAULT_CELL_WIDTH,
    gap: int = DEFAULT_GAP,
    dark: Optional[bool] = None,
    labels: bool = True,
    title: Optional[str] = None,
) -> Dict[str, object]:
    """拼接总览图并保存，返回统计信息（列/行/单格尺寸/背景）。"""
    if not image_paths:
        raise ValueError("没有可用的预览图")
    if cols < 1:
        raise ValueError("--cols 必须 ≥ 1")
    if cell_width < 80:
        raise ValueError("--width 太小（建议 ≥ 320）")

    cells: List[Tuple[Image.Image, float, str]] = []
    for path in image_paths:
        image, luminance = _prepare(path, cell_width)
        cells.append((image, luminance, _page_label(path)))

    cell_height = max(image.height for image, _lum, _label in cells)
    rows = int(math.ceil(len(cells) / float(cols)))

    if dark is None:  # 全图平均亮度 < 0.5 视为深色 deck
        avg = sum(lum for _im, lum, _label in cells) / float(len(cells))
        dark = avg < 0.5
    bg = DARK_BG if dark else LIGHT_BG
    ink = DARK_INK if dark else LIGHT_INK
    border = (58, 60, 62) if dark else (214, 212, 206)

    badge_size = max(11, int(round(cell_width * 0.032)))
    title_font = _font(max(16, int(round(cell_width * 0.045))), bold=True)
    title_size = int(getattr(title_font, "size", 18) or 18)
    title_h = int(title_size * 2.0) if title else 0

    width = gap + cols * (cell_width + gap)
    height = gap + title_h + rows * (cell_height + gap)
    canvas = Image.new("RGB", (max(1, width), max(1, height)), bg)
    draw = ImageDraw.Draw(canvas)

    if title:
        draw.text((gap * 2, int(gap + title_size * 0.35)), title, font=title_font, fill=ink)

    for order, (image, _lum, label) in enumerate(cells):
        row, col = divmod(order, cols)
        x = gap + col * (cell_width + gap)
        y = gap + title_h + row * (cell_height + gap)
        # 图像在格内居中（宽度已统一，只补上下留白）
        paste_y = y + int((cell_height - image.height) / 2)
        canvas.paste(image, (x, paste_y))
        draw.rectangle([(x - 1, paste_y - 1), (x + cell_width, paste_y + image.height)], outline=border, width=1)
        if labels:
            _draw_badge(canvas, draw, (x + max(6, gap // 2), paste_y + max(6, gap // 2)), order + 1, label, ink, badge_size)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    suffix = out_path.suffix.lower()
    if suffix not in (".png", ".jpg", ".jpeg", ".webp"):
        out_path = out_path.with_suffix(".png")
    canvas.save(out_path, format="JPEG" if out_path.suffix.lower() in (".jpg", ".jpeg") else "PNG", quality=92)

    return {
        "count": len(cells),
        "cols": cols,
        "rows": rows,
        "cell": [cell_width, cell_height],
        "size": [canvas.width, canvas.height],
        "dark": bool(dark),
        "out": str(out_path),
    }


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    """命令行参数。"""
    parser = argparse.ArgumentParser(
        prog="contact_sheet.py",
        description="把一页一张的预览 PNG 拼成总览图（contact sheet），用于整体节奏检查",
        epilog="示例：python scripts/contact_sheet.py --dir previews --out contact-sheet.png --cols 3",
    )
    parser.add_argument("--dir", required=True, help="预览图目录（.png/.jpg/.webp）")
    parser.add_argument("--out", default="contact-sheet.png", help="输出路径（默认 contact-sheet.png）")
    parser.add_argument("--cols", type=int, default=DEFAULT_COLS, help="列数（默认 3）")
    parser.add_argument("--width", type=int, default=DEFAULT_CELL_WIDTH, dest="cell_width", help="单格宽度像素（默认 560）")
    parser.add_argument("--gap", type=int, default=DEFAULT_GAP, help="格间距像素（默认 12）")
    parser.add_argument("--title", default=None, help="总览图顶部标题（省略则不加）")
    parser.add_argument("--recursive", action="store_true", help="递归子目录")
    parser.add_argument("--dark", dest="dark", action="store_true", default=None, help="强制深色底")
    parser.add_argument("--light", dest="dark", action="store_false", help="强制浅色底")
    parser.add_argument("--no-labels", dest="labels", action="store_false", help="不画序号与页名")
    parser.add_argument("--version", action="version", version="contact_sheet.py 1.0.0")
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    """入口：返回退出码（0 成功，2 输入不可用）。"""
    args = parse_args(argv)
    folder = Path(args.dir).expanduser()
    try:
        paths = collect_images(folder, recursive=bool(args.recursive))
        if not paths:
            print("ERROR 目录内没有 PNG/JPEG/WebP 预览图：%s" % folder, file=sys.stderr)
            return 2
        info = build_contact_sheet(
            paths,
            Path(args.out).expanduser(),
            cols=args.cols,
            cell_width=args.cell_width,
            gap=max(0, args.gap),
            dark=args.dark,
            labels=bool(args.labels),
            title=args.title,
        )
    except (FileNotFoundError, ValueError) as exc:
        print("ERROR %s" % exc, file=sys.stderr)
        return 2
    except OSError as exc:
        print("ERROR 预览图读取或写出失败：%s" % exc, file=sys.stderr)
        return 2

    print(
        "已拼接 %d 张预览 → %s（%d 列 × %d 行，单格 %d×%d px，整图 %d×%d px，%s底）"
        % (
            info["count"],
            info["out"],
            info["cols"],
            info["rows"],
            info["cell"][0],
            info["cell"][1],
            info["size"][0],
            info["size"][1],
            "深" if info["dark"] else "浅",
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
