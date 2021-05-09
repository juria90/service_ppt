#!/usr/bin/env python3
"""
"""
import argparse
import glob
import numpy as np

# pip install Pillow
# or python -m pip install Pillow --update
from PIL import Image, ImageColor


def white_to_transparent(img):
    x = np.asarray(img).copy()
    x[:, :, 3] = (255 * (x[:, :, :3] != 255).any(axis=2)).astype(np.uint8)

    img = Image.fromarray(x)
    return img


def color_to_transparent(filename, toFilename, color):
    fromColor = (color[0], color[1], color[2], 255)  # change the opacity to 255
    toColor = (color[0], color[1], color[2], 0)  # change the opacity to 0

    img = Image.open(filename)
    img = img.convert("RGBA")

    changed = False
    if fromColor == (255, 255, 255, 255):
        img = white_to_transparent(img)
        changed = True
    else:
        pixdata = img.load()

        changed = False
        width, height = img.size
        for y in range(height):
            for x in range(width):
                if pixdata[x, y] == fromColor:
                    pixdata[x, y] = toColor
                    changed = True

    if changed:
        img.save(toFilename, "PNG")

    img.close()

    return changed


def parse_color(string):
    color = None

    try:
        color = ImageColor.getrgb(string)
    except ValueError as e:
        msg = "%r is not a valid color name or value" % string
        raise argparse.ArgumentTypeError(msg)

    return color


def parse_cmdline():
    parser = argparse.ArgumentParser(description="Convert a specific color into transparent color in image files.")
    parser.add_argument(
        "--color",
        type=parse_color,
        default=ImageColor.getrgb("white"),
        help="CSS3-style color to convert to transparent.",
    )
    parser.add_argument("filenames", nargs="+", help="Filenames to conver the color.")

    return parser


if __name__ == "__main__":
    parser = parse_cmdline()
    # args = parser.parse_args([r'C:\Users\juria\Desktop\OnlineService\output\Slide001.PNG'])
    args = parser.parse_args()

    for f in args.filenames:
        for f1 in glob.glob(f):
            print(f'Processing file "{f1}": ', end="")
            changed = color_to_transparent(f1, f1, args.color)
            print("Converted" if changed else "Skipped")
