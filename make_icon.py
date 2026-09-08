# -*- coding: utf-8 -*-
"""ساخت آیکون جذب  ->  icon.ico"""
import math
from PIL import Image, ImageDraw, ImageFont
import arabic_reshaper
from bidi.algorithm import get_display

S = 1024
img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(img)

# ---- پس‌زمینهٔ گرادیان آبی -> فیروزه‌ای (قطری)
c1 = (28, 78, 128)      # 1C4E80
c2 = (34, 193, 195)     # 22C1C3
grad = Image.new("RGB", (S, S))
gd = ImageDraw.Draw(grad)
for y in range(S):
    for_x = y / S
    # افقی+عمودی برای حالت قطری
    t = for_x
    r = int(c1[0] + (c2[0] - c1[0]) * t)
    g = int(c1[1] + (c2[1] - c1[1]) * t)
    b = int(c1[2] + (c2[2] - c1[2]) * t)
    gd.line([(0, y), (S, y)], fill=(r, g, b))

# ماسک گوشه‌گرد
mask = Image.new("L", (S, S), 0)
md = ImageDraw.Draw(mask)
rad = int(S * 0.22)
md.rounded_rectangle([0, 0, S, S], radius=rad, fill=255)
img.paste(grad, (0, 0), mask)

d = ImageDraw.Draw(img)

# ---- قیف جذب (سفید، نیمه‌شفاف) وسط-بالا
cx = S // 2
funnel_top = int(S * 0.36)
funnel_w = int(S * 0.46)
funnel_h = int(S * 0.20)
neck_w = int(S * 0.085)
p = [
    (cx - funnel_w // 2, funnel_top),
    (cx + funnel_w // 2, funnel_top),
    (cx + neck_w // 2, funnel_top + funnel_h),
    (cx - neck_w // 2, funnel_top + funnel_h),
]
d.polygon(p, fill=(255, 255, 255, 235))
# لولهٔ خروجی
d.rectangle([cx - neck_w // 2, funnel_top + funnel_h,
             cx + neck_w // 2, funnel_top + funnel_h + int(S * 0.05)],
            fill=(255, 255, 255, 235))

# ---- نقطه‌های ورودی بالای قیف (متقاضیان)
dot_r = int(S * 0.032)
for i, dx in enumerate((-0.135, 0.0, 0.135)):
    x = cx + int(S * dx)
    y = funnel_top - int(S * 0.055)
    d.ellipse([x - dot_r, y - dot_r, x + dot_r, y + dot_r], fill=(255, 214, 92, 255))

# ---- سه میلهٔ صعودی زیر قیف (نمودار)
base_y = int(S * 0.80)
bw = int(S * 0.11)
gap = int(S * 0.05)
heights = [int(S * 0.11), int(S * 0.18), int(S * 0.26)]
start_x = cx - (3 * bw + 2 * gap) // 2
for i, h in enumerate(heights):
    x0 = start_x + i * (bw + gap)
    d.rounded_rectangle([x0, base_y - h, x0 + bw, base_y],
                        radius=int(bw * 0.28), fill=(255, 255, 255, 255))

# ---- نوشتهٔ «جذب»
txt = get_display(arabic_reshaper.reshape("جذب"))
font = None
for fp in (r"C:\Windows\Fonts\tahomabd.ttf", r"C:\Windows\Fonts\tahoma.ttf",
           r"C:\Windows\Fonts\arialbd.ttf"):
    try:
        font = ImageFont.truetype(fp, int(S * 0.185))
        break
    except OSError:
        continue
bb = d.textbbox((0, 0), txt, font=font)
tw, th = bb[2] - bb[0], bb[3] - bb[1]
d.text((cx - tw // 2 - bb[0], int(S * 0.055) - bb[1]), txt, font=font,
       fill=(255, 255, 255, 255))

# ---- ذخیره به‌صورت ico چندسایز
sizes = [256, 128, 64, 48, 32, 16]
img.save("icon.ico", sizes=[(s, s) for s in sizes])
img.resize((512, 512), Image.LANCZOS).save("icon_preview.png")
print("saved icon.ico + icon_preview.png")
