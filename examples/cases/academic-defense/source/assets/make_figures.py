# -*- coding: utf-8 -*-
"""生成学术答辩案例的两张示意图（纯本地 PIL 绘制，不下载任何网络素材）。

用法：
    python make_figures.py            # 在脚本所在目录写出两张 PNG
输出：
    fig-site-plan.png   1600x1000   场地平面示意：3 条街道 / 42 个测点（14x3）
    fig-section.png     1600x900    街道峡谷剖面与冬至正午太阳高度角

两图共用一套比例：平面图 2 px = 1 m（比例尺 0–200 m = 400 px，测点间距 25 m = 50 px）。
配色与 deck.json 的 tokens.colors 一致：墨底 #0E1620 / 面板 #1B2733 /
线条 #35485A / 强调 #C8A24A / 正文 #EDE6D8 / 次级 #8C99A6。
"""

from __future__ import annotations

import math
import random
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).resolve().parent

BG = (14, 22, 32)          # #0E1620
PANEL = (27, 39, 51)       # #1B2733
LINE = (53, 72, 90)        # #35485A
ACCENT = (200, 162, 74)    # #C8A24A
FG = (237, 230, 216)       # #EDE6D8
MUTED = (140, 153, 166)    # #8C99A6

FONT_CN = r"C:\Windows\Fonts\msyh.ttc"
FONT_MONO = r"C:\Windows\Fonts\CascadiaMono.ttf"

PX_PER_M = 2.0             # 平面图比例：2 px = 1 m


def cn(size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_CN, size)


def mono(size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_MONO, size)


def site_plan(out: Path) -> None:
    """场地平面：3 条街道（A 无遮荫 / B 单侧 / C 双侧）、42 个测点（14x3）。"""
    W, H = 1600, 1000
    img = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(img)

    # 地块边框
    d.rectangle([90, 130, W - 90, H - 150], outline=LINE, width=3)

    # 5x3 街区
    for x in (140, 420, 700, 980, 1260):
        for y in (200, 450, 700):
            d.rectangle([x, y, x + 220, y + 190], fill=PANEL, outline=LINE, width=3)

    # 街道中线：A、C 东西向，B 南北向
    a_y, b_x, c_y = 415, 330, 665
    d.line([110, a_y, W - 110, a_y], fill=ACCENT, width=7)
    d.line([b_x, 140, b_x, H - 160], fill=ACCENT, width=7)
    d.line([110, c_y, W - 110, c_y], fill=ACCENT, width=7)

    # 测点：每条街 14 个，间距 25 m = 50 px（A/C 从 x=180 起，B 从 y=170 起）
    rnd = random.Random(20260601)
    step = 25 * PX_PER_M
    pts = [(180 + i * step, a_y + rnd.uniform(-8, 8)) for i in range(14)]
    pts += [(b_x + rnd.uniform(-8, 8), 170 + i * step) for i in range(14)]
    pts += [(180 + i * step, c_y + rnd.uniform(-8, 8)) for i in range(14)]
    for px, py in pts:
        d.ellipse([px - 10, py - 10, px + 10, py + 10], fill=ACCENT, outline=BG, width=3)

    # 北向
    d.line([W - 190, 250, W - 190, 185], fill=FG, width=3)
    d.polygon([(W - 190, 168), (W - 203, 197), (W - 177, 197)], fill=FG)
    d.text((W - 190, 262), "N", font=mono(36), fill=FG, anchor="mt")

    # 街道标注（贴在中线端头外侧，避开测点）
    d.text((W - 108, a_y - 26), "A　无遮荫 24 m", font=cn(42), fill=ACCENT, anchor="rb")
    d.text((b_x + 22, 150), "B　单侧行道树 18 m", font=cn(42), fill=ACCENT, anchor="lt")
    d.text((W - 108, c_y - 26), "C　双侧连续树冠 21 m", font=cn(42), fill=ACCENT, anchor="rb")

    # 比例尺：200 m = 400 px
    sx, sy = 100, H - 92
    d.line([sx, sy, sx + 200 * PX_PER_M, sy], fill=FG, width=3)
    for k in (0, 100, 200):
        d.line([sx + k * PX_PER_M, sy - 12, sx + k * PX_PER_M, sy + 12], fill=FG, width=3)
    d.text((sx, sy + 24), "0", font=mono(34), fill=MUTED)
    d.text((sx + 100 * PX_PER_M, sy + 24), "100", font=mono(34), fill=MUTED, anchor="mt")
    d.text((sx + 200 * PX_PER_M + 14, sy + 24), "200 m", font=mono(34), fill=MUTED)

    # 图例（放在边框下方，避开街区与比例尺）
    lx, ly = 700, H - 96
    d.ellipse([lx, ly - 12, lx + 24, ly + 12], fill=ACCENT, outline=BG, width=3)
    d.text((lx + 42, ly - 26), "42 个实测点（每街 14 个）", font=cn(38), fill=FG)
    d.line([lx, ly + 52, lx + 24, ly + 52], fill=ACCENT, width=7)
    d.text((lx + 42, ly + 36), "对照街道 A / B / C", font=cn(38), fill=FG)

    img.save(out)
    print("wrote", out, out.stat().st_size, "bytes")


def section(out: Path) -> None:
    """剖面：街宽 24 m，左楼 45 m / 右楼 30 m，冬至正午太阳高度角 22.6°。"""
    W, H = 1600, 900
    img = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(img)

    ground_y = 745
    street_l, street_r = 560.0, 1200.0     # 640 px = 24 m
    k_v = 6.2                               # 竖向每米 6.2 px（压缩，避免超画布）
    h_l, h_r = 45, 30
    left_top = ground_y - h_l * k_v
    right_top = ground_y - h_r * k_v

    # 地面
    d.rectangle([90, ground_y, W - 90, ground_y + 16], fill=LINE)

    # 建筑
    d.rectangle([120, left_top, street_l, ground_y], fill=PANEL, outline=LINE, width=4)
    d.rectangle([street_r, right_top, 1510, ground_y], fill=PANEL, outline=LINE, width=4)

    # 阴影：右楼顶投影，30 / tan(22.6°) ≈ 72 m，落到左楼立面约 20 m 高处
    shadow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow)
    left_hit_y = ground_y - (h_r - 24 * math.tan(math.radians(22.6))) * k_v
    sd.polygon(
        [(street_r, right_top), (street_l, left_hit_y), (street_l, ground_y), (street_r, ground_y)],
        fill=(6, 10, 16, 185),
    )
    img = Image.alpha_composite(img.convert("RGBA"), shadow).convert("RGB")
    d = ImageDraw.Draw(img)

    # 行道树两棵（树冠压在街道上方）
    for cx in (700, 1030):
        d.ellipse([cx - 122, ground_y - 250, cx + 122, ground_y - 116], fill=(20, 32, 43),
                  outline=LINE, width=4)
        d.line([cx, ground_y - 128, cx, ground_y], fill=LINE, width=8)

    # 太阳与光线
    sun = (1478, 118)
    d.ellipse([sun[0] - 30, sun[1] - 30, sun[0] + 30, sun[1] + 30], fill=ACCENT)
    ang = math.radians(22.6)
    for k in range(4):
        off = (k - 1.5) * 46
        x1 = sun[0] - off * math.sin(ang) - 34
        y1 = sun[1] + off * math.cos(ang) + 34
        x2, y2 = x1 - math.cos(ang) * 700, y1 + math.sin(ang) * 700
        d.line([x1, y1, x2, y2], fill=ACCENT, width=3)

    # 角度标注：右楼屋顶处光线与水平面夹角
    ax, ay = street_r, right_top
    d.line([ax, ay, ax + 170, ay], fill=MUTED, width=3)
    d.arc([ax - 150, ay - 150, ax + 150, ay + 150], start=360 - 22.6, end=360, fill=ACCENT, width=4)
    d.text((ax + 96, ay - 96), "22.6°", font=mono(42), fill=ACCENT, anchor="lm")

    # 尺寸线
    def hdim(x1, x2, y, label):
        d.line([x1, y, x2, y], fill=MUTED, width=3)
        for x in (x1, x2):
            d.line([x, y - 14, x, y + 14], fill=MUTED, width=3)
        d.text(((x1 + x2) / 2, y - 16), label, font=cn(46), fill=MUTED, anchor="mb")

    def vdim(x, y_top, y_bot, label, anchor="lm"):
        d.line([x, y_top, x, y_bot], fill=MUTED, width=3)
        for y in (y_top, y_bot):
            d.line([x - 14, y, x + 14, y], fill=MUTED, width=3)
        d.text((x + 20, (y_top + y_bot) / 2), label, font=cn(46), fill=MUTED, anchor=anchor)

    hdim(street_l, street_r, ground_y + 104, "街宽 24 m")
    vdim(1470, right_top, ground_y, "30 m", anchor="rm")
    vdim(96, left_top, ground_y, "45 m", anchor="rm")

    # 图例
    d.rectangle([700, 104, 762, 138], fill=(9, 14, 21), outline=LINE, width=3)
    d.text((782, 98), "地面与立面遮荫区", font=cn(40), fill=FG)

    img.save(out)
    print("wrote", out, out.stat().st_size, "bytes")


if __name__ == "__main__":
    site_plan(HERE / "fig-site-plan.png")
    section(HERE / "fig-section.png")


