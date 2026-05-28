from __future__ import annotations

import argparse
from pathlib import Path

from PIL import Image

ICON_SIZES = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (24, 24), (16, 16)]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Convert a PNG image into a multi-size ICO file.")
    parser.add_argument("input_png", type=Path)
    parser.add_argument("output_ico", type=Path)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if not args.input_png.exists():
        raise FileNotFoundError(f"Input image not found: {args.input_png}")

    args.output_ico.parent.mkdir(parents=True, exist_ok=True)
    with Image.open(args.input_png) as image:
        rgba_image = image.convert("RGBA")
        rgba_image.save(args.output_ico, format="ICO", sizes=ICON_SIZES)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
