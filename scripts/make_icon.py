from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw


def main() -> None:
    project_root = Path(__file__).resolve().parent.parent
    assets_dir = project_root / "assets"
    assets_dir.mkdir(exist_ok=True)

    size = 512
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    draw.rounded_rectangle((32, 32, 480, 480), radius=112, fill=(24, 57, 43, 255))
    draw.rounded_rectangle((88, 92, 424, 420), radius=48, fill=(247, 244, 234, 255))
    draw.rounded_rectangle((132, 156, 380, 204), radius=24, fill=(44, 102, 73, 255))
    draw.rounded_rectangle((132, 236, 380, 284), radius=24, fill=(93, 134, 111, 255))
    draw.rounded_rectangle((132, 316, 280, 364), radius=24, fill=(210, 193, 151, 255))
    draw.polygon([(326, 332), (360, 368), (420, 286)], fill=(24, 57, 43, 255))
    draw.line((336, 344, 358, 366), fill=(247, 244, 234, 255), width=12)
    draw.line((358, 366, 410, 298), fill=(247, 244, 234, 255), width=12)

    png_path = assets_dir / "customer_notes_merger.png"
    ico_path = assets_dir / "customer_notes_merger.ico"
    icns_path = assets_dir / "customer_notes_merger.icns"

    image.save(png_path)
    image.save(ico_path, sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])
    image.save(icns_path)


if __name__ == "__main__":
    main()
