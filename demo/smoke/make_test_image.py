"""生成 demo/smoke 用的 1200×800 测试图（不依赖网络与外部素材）。

图片本身是一张「几何栅格卡」，用于验证 build_pptx.py 的 fit=cover 裁切是否正确：
四角各有标记、边框四周留 48 px 内边距，裁切若出错，边距会明显变窄或标记被切掉。

用法：
    python demo/smoke/make_test_image.py
"""

from __future__ import annotations

import os
import sys

from PIL import Image, ImageDraw

W, H = 1200, 800
BG = (21, 32, 48)
GRID = (44, 62, 86)
PRIMARY = (61, 123, 255)
ACCENT = (242, 178, 74)
FG = (240, 245, 250)


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    image = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(image)

    # 40 px 栅格：cover 裁切后仍应看到完整方格，用于肉眼判断是否被拉伸
    for x in range(0, W, 40):
        draw.line([(x, 0), (x, H)], fill=GRID, width=1)
    for y in range(0, H, 40):
        draw.line([(0, y), (W, y)], fill=GRID, width=1)

    # 四周 48 px 边框：裁切比例若算错，边框会缺边
    draw.rectangle([0, 0, W - 1, H - 1], outline=PRIMARY, width=6)
    draw.rectangle([48, 48, W - 49, H - 49], outline=ACCENT, width=3)

    # 中心三块色块（对角线），用于确认没有被翻转
    draw.rectangle([120, 560, 360, 720], fill=PRIMARY)
    draw.ellipse([480, 300, 720, 540], fill=ACCENT)
    draw.polygon([(760, 700), (1060, 700), (1060, 400)], fill=FG)

    # 四角标记：左上为圆形、右上为方形，便于判断方向
    draw.ellipse([70, 70, 130, 130], fill=ACCENT)
    draw.rectangle([W - 130, 70, W - 70, 130], fill=PRIMARY)

    # 尺寸标注（用默认位图字体，避免依赖特定字体文件）
    draw.text((60, 745), "1200 x 800  |  smoke test image", fill=FG)

    out_dir = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.join(out_dir, "assets", "test-image-1200x800.png")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    image.save(out_path, "PNG", optimize=True)
    print(f"已生成：{out_path}（{W}x{H}，{os.path.getsize(out_path):,} 字节）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
