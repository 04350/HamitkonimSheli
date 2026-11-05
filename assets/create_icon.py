# -*- coding: utf-8 -*-
"""Generate a crisp multi-resolution Windows icon for the recipe library."""
from pathlib import Path
from PIL import Image, ImageDraw

SIZE = 1024
OUT = Path(__file__).resolve().parent / "app_icon.ico"
PNG = Path(__file__).resolve().parent / "app_icon.png"

# Draw at high resolution and downsample for clean edges at small icon sizes.
image = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
draw = ImageDraw.Draw(image)

# Deep blue rounded tile with a subtle vertical gradient.
top = (30, 75, 116)
bottom = (11, 35, 61)
for y in range(SIZE):
    t = y / (SIZE - 1)
    color = tuple(round(top[i] * (1 - t) + bottom[i] * t) for i in range(3))
    draw.line((0, y, SIZE, y), fill=color + (255,))
mask = Image.new("L", (SIZE, SIZE), 0)
ImageDraw.Draw(mask).rounded_rectangle((18, 18, 1006, 1006), radius=220, fill=255)
image.putalpha(mask)
draw = ImageDraw.Draw(image)

# Soft decorative glows.
draw.ellipse((600, -180, 1160, 380), fill=(34, 91, 125, 255))
draw.ellipse((-220, 600, 360, 1180), fill=(18, 65, 94, 255))

# Steam curls above the pot.
def bezier(p0, p1, p2, p3, count=50):
    points = []
    for n in range(count + 1):
        t = n / count
        u = 1 - t
        x = u**3*p0[0] + 3*u*u*t*p1[0] + 3*u*t*t*p2[0] + t**3*p3[0]
        y = u**3*p0[1] + 3*u*u*t*p1[1] + 3*u*t*t*p2[1] + t**3*p3[1]
        points.append((round(x), round(y)))
    return points

steam_color = (226, 244, 246, 255)
for curve in [
    ((380, 350), (310, 270), (440, 230), (390, 150)),
    ((510, 330), (590, 250), (470, 220), (535, 125)),
    ((650, 350), (720, 275), (610, 215), (690, 155)),
]:
    draw.line(bezier(*curve), fill=steam_color, width=30, joint="curve")

# Open cookbook at the bottom, behind the pot.
book_outline = (224, 181, 116, 255)
left_page = [(128, 698), (492, 760), (492, 914), (158, 842)]
right_page = [(532, 760), (896, 698), (866, 842), (532, 914)]
draw.polygon(left_page, fill=(250, 244, 224, 255), outline=book_outline)
draw.polygon(right_page, fill=(255, 249, 233, 255), outline=book_outline)
draw.line((512, 754, 512, 906), fill=(205, 171, 118, 255), width=18)
for y in (792, 830, 866):
    draw.line((205, y, 447, y + 42), fill=(189, 197, 197, 255), width=8)
    draw.line((577, y + 42, 819, y), fill=(189, 197, 197, 255), width=8)

# Pot handles.
draw.rounded_rectangle((205, 495, 310, 645), radius=48, fill=(12, 42, 67, 255), outline=(231, 242, 243, 255), width=16)
draw.rounded_rectangle((714, 495, 819, 645), radius=48, fill=(12, 42, 67, 255), outline=(231, 242, 243, 255), width=16)

# Main pot body.
draw.rounded_rectangle((270, 405, 754, 770), radius=112, fill=(14, 151, 157, 255), outline=(239, 247, 243, 255), width=22)
draw.rounded_rectangle((315, 470, 709, 705), radius=76, fill=(18, 169, 169, 255))
draw.arc((335, 485, 690, 690), 195, 330, fill=(115, 218, 210, 255), width=19)

# Lid, rim and warm golden knob.
draw.rounded_rectangle((238, 380, 786, 468), radius=44, fill=(245, 190, 89, 255), outline=(255, 238, 196, 255), width=18)
draw.ellipse((420, 335, 604, 415), fill=(255, 215, 130, 255), outline=(255, 244, 211, 255), width=16)
draw.rounded_rectangle((464, 300, 560, 360), radius=28, fill=(255, 215, 130, 255), outline=(255, 244, 211, 255), width=12)

# A clean front emblem on the pot: a small white recipe leaf.
draw.ellipse((435, 530, 589, 685), fill=(245, 250, 239, 255), outline=(255, 255, 255, 255), width=10)
draw.arc((465, 555, 559, 650), 205, 335, fill=(8, 126, 139, 255), width=13)
draw.line((512, 572, 512, 646), fill=(8, 126, 139, 255), width=12)
draw.ellipse((481, 594, 505, 618), fill=(242, 174, 75, 255))
draw.ellipse((520, 615, 544, 639), fill=(242, 174, 75, 255))

# Smooth antialiasing and produce a real multi-size ICO plus preview PNG.
image = image.resize((512, 512), Image.Resampling.LANCZOS)
image.save(PNG, format="PNG", optimize=True)
image.save(
    OUT,
    format="ICO",
    sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)],
)
print(f"Created {OUT} ({OUT.stat().st_size} bytes)")
