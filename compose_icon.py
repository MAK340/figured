"""Composite a crisp sigma onto a generated tile and export icon.png / icon.ico."""
import os, sys
from PIL import Image, ImageDraw, ImageFont

SRC = sys.argv[1] if len(sys.argv) > 1 else r"C:\SumSelect\icons\tile_emerald.png"
OUT = r"C:\SumSelect"
SIZE = 512

tile = Image.open(SRC).convert("RGBA").resize((SIZE, SIZE), Image.LANCZOS)

# keep only the tile: trim the flat background the model painted around it
bg = tile.getpixel((6, 6))
px = tile.load()
xs, ys = [], []
for y in range(0, SIZE, 2):
    for x in range(0, SIZE, 2):
        r, g, b, _ = px[x, y]
        if abs(r - bg[0]) + abs(g - bg[1]) + abs(b - bg[2]) > 60:
            xs.append(x); ys.append(y)
if xs:
    pad = 4
    box = (max(0, min(xs) - pad), max(0, min(ys) - pad),
           min(SIZE, max(xs) + pad), min(SIZE, max(ys) + pad))
    tile = tile.crop(box).resize((SIZE, SIZE), Image.LANCZOS)

# rounded-corner alpha mask so the icon sits cleanly on any background
mask = Image.new("L", (SIZE, SIZE), 0)
ImageDraw.Draw(mask).rounded_rectangle((0, 0, SIZE - 1, SIZE - 1), radius=int(SIZE * 0.22), fill=255)
tile.putalpha(mask)

# sigma, drawn with a real font so it stays sharp at 16px
draw = ImageDraw.Draw(tile)
font = None
for name in ("seguisb.ttf", "segoeuib.ttf", "arialbd.ttf"):
    try:
        font = ImageFont.truetype(name, int(SIZE * 0.56))
        break
    except OSError:
        continue
if font is None:
    font = ImageFont.load_default()
cx, cy = SIZE // 2, int(SIZE * 0.47)
draw.text((cx + 6, cy + 8), "\u03a3", font=font, fill=(0, 0, 0, 90), anchor="mm")
draw.text((cx, cy), "\u03a3", font=font, fill=(255, 255, 255, 235), anchor="mm")

tile.save(os.path.join(OUT, "icon.png"))
tile.save(os.path.join(OUT, "icon.ico"),
          sizes=[(16, 16), (20, 20), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print("wrote icon.png and icon.ico from", SRC)
