# -*- coding: utf-8 -*-
"""把给定图片做成应用图标 app.ico（透明背景 + 多尺寸）

背景不是纯白而是米白纸色，所以用「与四角均值的颜色距离」判定，
而不是三通道阈值——阈值法会把 b=237 的背景判成前景。
"""
import sys
from collections import deque
from pathlib import Path

from PIL import Image

SRC = Path(sys.argv[1] if len(sys.argv) > 1 else "")
OUT = Path(sys.argv[2] if len(sys.argv) > 2 else "app.ico")

im = Image.open(SRC).convert("RGBA")
w, h = im.size
print(f"原始尺寸 {w}x{h}")
px = im.load()

# 0) 从四角取背景参考色
samples = [px[0, 0], px[w - 1, 0], px[0, h - 1], px[w - 1, h - 1]]
br = sum(s[0] for s in samples) / 4
bg_ = sum(s[1] for s in samples) / 4
bb = sum(s[2] for s in samples) / 4
print(f"背景参考色 ({br:.0f},{bg_:.0f},{bb:.0f})")


def dist(p):
    return max(abs(p[0] - br), abs(p[1] - bg_), abs(p[2] - bb))


def is_bg(p, tol=14):
    return p[3] > 0 and dist(p) <= tol


# 1) 从整条外边框 flood fill 抠背景
seen = bytearray(w * h)
dq = deque()
border = []
for x in range(w):
    border.append((x, 0)); border.append((x, h - 1))
for y in range(h):
    border.append((0, y)); border.append((w - 1, y))
for bx, by in border:
    if is_bg(px[bx, by]) and not seen[by * w + bx]:
        dq.append((bx, by))
        seen[by * w + bx] = 1

bg_count = 0
while dq:
    x, y = dq.popleft()
    px[x, y] = (255, 255, 255, 0)
    bg_count += 1
    for nx, ny in ((x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)):
        if 0 <= nx < w and 0 <= ny < h and not seen[ny * w + nx]:
            if is_bg(px[nx, ny]):
                seen[ny * w + nx] = 1
                dq.append((nx, ny))
print(f"抠掉背景像素 {bg_count} ({bg_count * 100 // (w * h)}%)")


# 2) 清边缘灰晕：接近背景色且紧邻透明区的像素一律透明
#    内部的白高光/眼镜片周围没有透明像素，因此不受影响
def near_transparent(x, y):
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            nx, ny = x + dx, y + dy
            if 0 <= nx < w and 0 <= ny < h and px[nx, ny][3] == 0:
                return True
    return False


for rnd in range(2):
    kill = []
    for y in range(h):
        for x in range(w):
            p = px[x, y]
            if p[3] > 0 and dist(p) <= 22 and near_transparent(x, y):
                kill.append((x, y))
    if not kill:
        break
    for x, y in kill:
        px[x, y] = (255, 255, 255, 0)
    print(f"第 {rnd + 1} 轮清灰晕 {len(kill)} 像素")

# 3) 裁到内容 bbox → 补成正方形
bbox = im.getbbox()
print(f"内容 bbox {bbox}")
im = im.crop(bbox)
cw, ch = im.size
side = max(cw, ch)
canvas = Image.new("RGBA", (side, side), (0, 0, 0, 0))
canvas.paste(im, ((side - cw) // 2, (side - ch) // 2), im)
im = canvas
print(f"正方形画布 {im.size}")

# 4) 256 基底，留 5% 内边距
PAD = 0.05
inner = int(256 * (1 - PAD * 2))
fig = im.resize((inner, inner), Image.LANCZOS)
base = Image.new("RGBA", (256, 256), (0, 0, 0, 0))
base.paste(fig, ((256 - inner) // 2, (256 - inner) // 2), fig)
for c in ((0, 0), (255, 0), (0, 255), (255, 255)):
    assert base.getpixel(c)[3] == 0, f"角点 {c} 不透明: {base.getpixel(c)}"
# 统计仍不透明的浅色像素（应为 0 或极少）
leak = sum(1 for y in range(256) for x in range(256)
           if base.getpixel((x, y))[3] > 0 and max(base.getpixel((x, y))[:3]) > 235
           and min(base.getpixel((x, y))[:3]) > 210)
print(f"四角透明校验通过，内容 {inner}px 居中；残留浅色像素 {leak}")

# 5) 写出 ico + 预览
sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (24, 24), (16, 16)]
base.save(OUT, format="ICO", sizes=sizes)
print(f"已写出 {OUT.resolve()}  {OUT.stat().st_size} 字节")

prev = OUT.with_name("app_icon_preview.png")
base.save(prev)
dark = Image.new("RGBA", (256, 256), (32, 32, 36, 255))
dark.alpha_composite(base)
dark.save(OUT.with_name("app_icon_dark.png"))
print("预览已写出（白底 + 深底各一张）")
