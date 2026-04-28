#!/usr/bin/env python
"""Compose deterministic 800x800 ecommerce main images for missing batch items."""

from __future__ import annotations

import argparse
import json
import textwrap
from datetime import datetime, timezone
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\segoeuib.ttf" if bold else r"C:\Windows\Fonts\segoeui.ttf",
    ]
    for candidate in candidates:
        try:
            return ImageFont.truetype(candidate, size)
        except OSError:
            pass
    return ImageFont.load_default()


def fit_text(draw: ImageDraw.ImageDraw, text: str, max_width: int, start_size: int, bold: bool) -> tuple[list[str], ImageFont.ImageFont]:
    size = start_size
    while size >= 18:
        fnt = font(size, bold)
        words = text.split()
        lines: list[str] = []
        current = ""
        for word in words:
            trial = f"{current} {word}".strip()
            if draw.textbbox((0, 0), trial, font=fnt)[2] <= max_width:
                current = trial
            else:
                if current:
                    lines.append(current)
                current = word
        if current:
            lines.append(current)
        if len(lines) <= 3 and all(draw.textbbox((0, 0), line, font=fnt)[2] <= max_width for line in lines):
            return lines, fnt
        size -= 2
    return textwrap.wrap(text, width=18)[:3], font(18, bold)


def compose(original_path: Path, output_path: Path, title: str) -> None:
    original = Image.open(original_path).convert("RGB")
    original.thumbnail((760, 760), Image.Resampling.LANCZOS)

    canvas = Image.new("RGB", (800, 800), "#f6f8fb")
    bg = Image.new("RGB", (800, 800), "#eef3f7")
    for y in range(800):
        shade = 246 - int(y * 18 / 800)
        ImageDraw.Draw(bg).line([(0, y), (800, y)], fill=(shade, min(250, shade + 4), min(255, shade + 8)))
    canvas.paste(bg)

    shadow = Image.new("RGBA", original.size, (0, 0, 0, 0))
    ImageDraw.Draw(shadow).rounded_rectangle([10, 10, original.size[0] - 10, original.size[1] - 10], radius=18, fill=(0, 0, 0, 40))
    shadow = shadow.filter(ImageFilter.GaussianBlur(18))
    x = (800 - original.size[0]) // 2
    y = 118
    canvas.paste(shadow.convert("RGB"), (x + 8, y + 12), shadow)
    canvas.paste(original, (x, y))

    draw = ImageDraw.Draw(canvas)
    draw.rounded_rectangle([24, 22, 776, 112], radius=18, fill=(255, 255, 255), outline=(218, 226, 235), width=2)
    headline = "All-Weather Wiper Blades"
    lines, headline_font = fit_text(draw, headline, 410, 36, True)
    draw.text((48, 38), "\n".join(lines), font=headline_font, fill=(20, 37, 58), spacing=4)
    draw.text((48, 82), "Clear visibility for daily driving", font=font(20), fill=(71, 85, 105))

    badge_font = font(23, True)
    labels = ["Durable Rubber", "Easy Install", "Smooth Wipe"]
    positions = [(36, 662), (292, 662), (548, 662)]
    for label, (bx, by) in zip(labels, positions):
        draw.rounded_rectangle([bx, by, bx + 216, by + 64], radius=16, fill=(255, 255, 255), outline=(203, 213, 225), width=2)
        draw.ellipse([bx + 18, by + 20, bx + 42, by + 44], fill=(16, 185, 129))
        draw.text((bx + 54, by + 18), label, font=badge_font, fill=(30, 41, 59))

    draw.rounded_rectangle([526, 34, 752, 96], radius=18, fill=(15, 23, 42))
    draw.text((552, 48), "2 PCS SET", font=font(28, True), fill=(255, 255, 255))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output_path, "PNG", optimize=True)


def update_ledger(status_path: Path, local_id: str, status: str, error: str | None = None) -> None:
    if not status_path.exists():
        return
    ledger = json.loads(status_path.read_text(encoding="utf-8"))
    records = ledger.get("records", {})
    if local_id in records:
        records[local_id]["status"] = status
        records[local_id]["error"] = error
        records[local_id]["updated_at"] = datetime.now(timezone.utc).isoformat()
        ledger["updated_at"] = datetime.now(timezone.utc).isoformat()
        status_path.write_text(json.dumps(ledger, ensure_ascii=False, indent=2), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Compose deterministic 800x800 main images for a Mango batch.")
    parser.add_argument("--manifest", required=True, type=Path)
    parser.add_argument("--status-ledger", type=Path, default=None)
    parser.add_argument("--limit", type=int, default=0, help="Maximum missing images to create; 0 means all")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    manifest = json.loads(args.manifest.read_text(encoding="utf-8"))
    made = 0
    skipped = 0
    failed = 0
    for item in manifest["items"]:
        out = Path(item["generated_image_path"])
        if out.exists():
            skipped += 1
            if args.status_ledger:
                update_ledger(args.status_ledger, item["local_id"], "verified")
            continue
        if args.limit and made >= args.limit:
            break
        try:
            compose(Path(item["original_image_path"]), out, item["title"])
            if args.status_ledger:
                update_ledger(args.status_ledger, item["local_id"], "verified")
        except Exception as exc:  # noqa: BLE001 - keep batch moving and record the row.
            failed += 1
            if args.status_ledger:
                update_ledger(args.status_ledger, item["local_id"], "failed", str(exc))
            continue
        made += 1
    print(json.dumps({"created": made, "skipped_existing": skipped, "failed": failed}, ensure_ascii=False))


if __name__ == "__main__":
    main()
