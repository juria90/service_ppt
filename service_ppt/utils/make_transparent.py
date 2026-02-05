#!/usr/bin/env python3
"""Utility for making image backgrounds transparent.

This module provides functionality to convert specific colors in images to transparent,
useful for exporting PowerPoint shapes with transparent backgrounds.
"""

import argparse
import glob

import numpy as np
from PIL import Image, ImageColor


def white_to_transparent(img: Image.Image) -> Image.Image:
    """Convert white pixels to transparent in an image.

    :param img: PIL Image to process
    :returns: Image with white pixels converted to transparent
    """
    x = np.asarray(img).copy()
    x[:, :, 3] = (255 * (x[:, :, :3] != 255).any(axis=2)).astype(np.uint8)

    return Image.fromarray(x)


def color_to_transparent(filename: str | bytes, to_filename: str | bytes, color: tuple[int, int, int]) -> bool:
    """Convert a specific color to transparent in an image file.

    :param filename: Input image file path
    :param toFilename: Output image file path
    :param color: RGB color tuple to convert to transparent
    :returns: True if any pixels were changed, False otherwise
    """
    from_color = (color[0], color[1], color[2], 255)  # change the opacity to 255
    to_color = (color[0], color[1], color[2], 0)  # change the opacity to 0

    img = Image.open(filename)
    img = img.convert("RGBA")

    changed = False
    if from_color == (255, 255, 255, 255):
        img = white_to_transparent(img)
        changed = True
    else:
        pixdata = img.load()

        changed = False
        width, height = img.size
        for y in range(height):
            for x in range(width):
                if pixdata[x, y] == from_color:
                    pixdata[x, y] = to_color
                    changed = True

    if changed:
        img.save(to_filename, "PNG")

    img.close()

    return changed


def parse_color(string: str) -> tuple[int, int, int]:
    """Parse a color string into RGB tuple.

    :param string: Color name or value (CSS3-style)
    :returns: RGB tuple (r, g, b)
    :raises argparse.ArgumentTypeError: If color string is invalid
    """
    color = None

    try:
        color = ImageColor.getrgb(string)
    except ValueError:
        msg = f"{string!r} is not a valid color name or value"
        raise argparse.ArgumentTypeError(msg)

    return color


def parse_cmdline() -> argparse.ArgumentParser:
    """Create and configure command-line argument parser.

    :returns: Configured ArgumentParser instance
    """
    parser = argparse.ArgumentParser(description="Convert a specific color into transparent color in image files.")
    parser.add_argument("--color", type=parse_color, default=ImageColor.getrgb("white"), help="CSS3-style color to convert to transparent.")
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
