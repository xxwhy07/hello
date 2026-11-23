# -*- coding: utf-8 -*-
"""Compress SDK reference screenshots for the HTML + DOCX deliverables."""
import os
from PIL import Image

BASE = r"E:\HiSparkStudioSDK\fbb_ws63\vendor\HiHope_NearLink_DK_WS63E_V03\doc\media"
OUT = r"E:\deepseek harness\学习\实训六_蜂鸣器\assets"
os.makedirs(OUT, exist_ok=True)

# name, relative path, max width, format, quality
ITEMS = [
    ("hardware", r"doc\image-20240415144302434.png", 900, "JPEG", 84),
    ("new_folder", r"doc\image-20240801170551992.png", 900, "PNG", None),
    ("copy_sample", r"beep\image-20240415150723600.png", 900, "PNG", None),
    ("cmake", r"doc\image-20240801171056148.png", 900, "PNG", None),
    ("kconfig", r"doc\image-20240801180202806.png", 900, "PNG", None),
    ("menuconfig", r"doc\image-20240801171406113.png", 900, "PNG", None),
    ("support_beep", r"beep\image-20240415150847116.png", 900, "PNG", None),
    ("build_success", r"readme\image-20250307164622717.png", 1000, "JPEG", 88),
    ("project_config", r"readme\image-20250716164922699.png", 1000, "JPEG", 84),
    ("flash_start", r"doc\image-20240801174117545.png", 900, "PNG", None),
    ("flash_connecting", r"doc\image-20240801174230202.png", 900, "PNG", None),
    ("board_usb", r"readme\image-20240801173105245.png", 900, "JPEG", 82),
]

for name, rel, width, fmt, quality in ITEMS:
    src = os.path.join(BASE, rel)
    if not os.path.exists(src):
        print("MISSING", src)
        continue
    im = Image.open(src).convert("RGB")
    if im.width > width:
        im = im.resize((width, round(im.height * width / im.width)), Image.LANCZOS)
    ext = ".jpg" if fmt == "JPEG" else ".png"
    dst = os.path.join(OUT, name + ext)
    if fmt == "JPEG":
        im.save(dst, "JPEG", quality=quality, optimize=True, progressive=True)
    else:
        im.save(dst, "PNG", optimize=True)
    print(f"{name}{ext}: {im.width}x{im.height}, {os.path.getsize(dst)/1024:.1f} KB")
