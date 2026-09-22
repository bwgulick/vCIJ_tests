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
SOURCE_DIR = r"c:\Users\bgulick\Desktop\16BMB_March_2026\e290730-Gulick\US\Ni1\Strainimg"
OUTPUT_PATH = os.path.join(SOURCE_DIR, "Nickel.gif")

FRAME_DURATION_MS = 500  # default duration per frame, in milliseconds
LOOP = 0  # 0 = loop forever

# Variable framerate: slow down the trailing frames. Set SLOW_TAIL_COUNT = 0 to
# disable. Example: last 100 frames at 400 ms, the rest at FRAME_DURATION_MS.
SLOW_TAIL_COUNT = 0  # number of trailing frames to slow down
SLOW_TAIL_DURATION_MS = 400  # duration for those trailing frames

# Black-frame filtering: drop frames whose brightest pixel is <= this value
# (0-255). Set SKIP_BLACK_FRAMES = False to keep every frame. Raise the
# threshold above 0 to also drop near-black frames (e.g. 5 or 10).
SKIP_BLACK_FRAMES = True
BLACK_FRAME_THRESHOLD = 0

SCALE = 0.5  # resize factor applied to each frame (1.0 = full resolution)

LABEL_HEIGHT = 60  # height in pixels of the label strip added above each image
LABEL_FONT_SIZE = 36
LABEL_BG_COLOR = "white"
LABEL_TEXT_COLOR = "black"
REPLACE_UNDERSCORES_WITH_SPACES = True

# Trailing "_NNN.tif" counter, stripped to build the label. Any digit count.
# Frames are ordered by modification time (see main), not by this number, because
# the counter resets between runs and the middle token (otr_200/400/1000) is not
# monotonic -- so neither the trailing number nor an alphabetical sort is reliable.
FILENAME_PATTERN = re.compile(r"_(\d+)\.tif$", re.IGNORECASE)

# ----------------------------------------------------------------------------


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


def is_black_frame(path):
    """True if the image's brightest pixel is at or below the threshold."""
    with Image.open(path) as image:
        _, max_val = image.convert("L").getextrema()
    return max_val <= BLACK_FRAME_THRESHOLD


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
    # Acquisition order = modification time; filename breaks second-level ties.
    files.sort(key=lambda f: (os.path.getmtime(f), os.path.basename(f)))

    if SKIP_BLACK_FRAMES:
        kept = [f for f in files if not is_black_frame(f)]
        dropped = len(files) - len(kept)
        if dropped:
            print(f"Skipped {dropped} black frame(s) (threshold {BLACK_FRAME_THRESHOLD})")
        files = kept

    if not files:
        raise SystemExit("No frames left to write.")

    font = load_font()
    frames = [build_frame(f, font) for f in files]

    # Per-frame durations: default everywhere, longer for the trailing frames.
    durations = [FRAME_DURATION_MS] * len(frames)
    if SLOW_TAIL_COUNT > 0:
        for i in range(max(0, len(frames) - SLOW_TAIL_COUNT), len(frames)):
            durations[i] = SLOW_TAIL_DURATION_MS

    frames[0].save(
        OUTPUT_PATH,
        save_all=True,
        append_images=frames[1:],
        duration=durations,
        loop=LOOP,
        optimize=True,
    )

    total_seconds = sum(durations) / 1000
    print(f"Wrote {len(frames)} frames to {OUTPUT_PATH}")
    tail = min(SLOW_TAIL_COUNT, len(frames))
    if tail:
        print(
            f"Durations: {len(frames) - tail} frame(s) at {FRAME_DURATION_MS} ms, "
            f"last {tail} at {SLOW_TAIL_DURATION_MS} ms"
        )
    else:
        print(f"Frame duration: {FRAME_DURATION_MS} ms")
    print(f"Total length: {total_seconds:.1f} s")


if __name__ == "__main__":
    main()
