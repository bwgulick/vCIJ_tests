"""
Build an animated GIF from a sequence of numbered TIFF images.

Edit the CONFIG section below, then run:  py make_gif.py
"""

import glob
import os
import re

from PIL import Image, ImageColor, ImageDraw, ImageFont

# ----------------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------------
SOURCE_DIR = r"c:\Users\bgulick\Desktop\16BMB_March_2026\e290730-Gulick\US\V1\Images"
OUTPUT_PATH = os.path.join(SOURCE_DIR, "sequence.gif")

FRAME_DURATION_MS = 2000  # how long each frame is shown, in milliseconds
LOOP = 0  # 0 = loop forever

SCALE = 0.5  # resize factor applied to each frame (1.0 = full resolution)

LABEL_HEIGHT = 60  # height in pixels of the label strip added above each image
LABEL_FONT_SIZE = 36
LABEL_BG_COLOR = "white"
LABEL_TEXT_COLOR = "black"
REPLACE_UNDERSCORES_WITH_SPACES = True

# Matches the trailing "_NNN.tif" used to order frames; stripped to build the label.
FILENAME_PATTERN = re.compile(r"_(\d{3})\.tif$", re.IGNORECASE)

# ----------------------------------------------------------------------------


def frame_index(filename):
    match = FILENAME_PATTERN.search(filename)
    if not match:
        raise ValueError(f"Filename does not match expected pattern: {filename}")
    return int(match.group(1))


def label_for(filename):
    label = FILENAME_PATTERN.sub("", filename)
    if REPLACE_UNDERSCORES_WITH_SPACES:
        label = label.replace("_", " ")
    return label


def load_font():
    try:
        return ImageFont.truetype("arial.ttf", LABEL_FONT_SIZE)
    except OSError:
        return ImageFont.load_default()


def build_frame(path, font):
    image = Image.open(path).convert("L")
    if SCALE != 1.0:
        new_size = (round(image.width * SCALE), round(image.height * SCALE))
        image = image.resize(new_size, Image.LANCZOS)

    bg = ImageColor.getcolor(LABEL_BG_COLOR, "L")
    fg = ImageColor.getcolor(LABEL_TEXT_COLOR, "L")

    canvas = Image.new("L", (image.width, image.height + LABEL_HEIGHT), color=bg)
    canvas.paste(image, (0, LABEL_HEIGHT))

    draw = ImageDraw.Draw(canvas)
    label = label_for(os.path.basename(path))
    bbox = draw.textbbox((0, 0), label, font=font)
    text_x = (canvas.width - (bbox[2] - bbox[0])) // 2
    text_y = (LABEL_HEIGHT - (bbox[3] - bbox[1])) // 2 - bbox[1]
    draw.text((text_x, text_y), label, font=font, fill=fg)

    return canvas.convert("P", palette=Image.ADAPTIVE)


def main():
    files = glob.glob(os.path.join(SOURCE_DIR, "*.tif"))
    files.sort(key=lambda f: frame_index(os.path.basename(f)))

    font = load_font()
    frames = [build_frame(f, font) for f in files]

    frames[0].save(
        OUTPUT_PATH,
        save_all=True,
        append_images=frames[1:],
        duration=FRAME_DURATION_MS,
        loop=LOOP,
        optimize=True,
    )

    total_seconds = FRAME_DURATION_MS * len(frames) / 1000
    print(f"Wrote {len(frames)} frames to {OUTPUT_PATH}")
    print(f"Frame duration: {FRAME_DURATION_MS} ms, total length: {total_seconds:.1f} s")


if __name__ == "__main__":
    main()
