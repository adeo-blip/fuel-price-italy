"""Render dashboard.pdf (page 1) to after.png, cropped to the non-white content."""
import os
import fitz
from PIL import Image, ImageChops

out = os.environ["FUEL_OUTDIR"]
doc = fitz.open(os.path.join(out, "dashboard.pdf"))
pix = doc[0].get_pixmap(matrix=fitz.Matrix(3, 3))
raw = os.path.join(out, "dashboard_full.png")
pix.save(raw)
img = Image.open(raw).convert("RGB")
bbox = ImageChops.difference(img, Image.new("RGB", img.size, (255, 255, 255))).getbbox()
pad = 30
box = (max(0, bbox[0] - pad), max(0, bbox[1] - pad), min(img.width, bbox[2] + pad), min(img.height, bbox[3] + pad))
img.crop(box).save(os.path.join(out, "after.png"))
print("after.png", box)
