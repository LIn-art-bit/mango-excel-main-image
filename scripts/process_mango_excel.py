#!/usr/bin/env python
"""Prepare and build Mango main-image Excel batches.

This script intentionally does not call an image-generation API. It prepares
reference images and builds the final workbook after generated images named
<local_id>.png have been placed in the generated image directory.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import urllib.request
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image


REQUIRED_HEADERS = ("本地ID", "产品标题", "产品图片1")
HEADER_ROW = 2
FINAL_IMAGE_SIZE = (800, 800)
DEFAULT_OUTPUT_ROOT = Path(r"D:\MangoMainImageBatches")


@dataclass
class Item:
    row_number: int
    local_id: str
    title: str
    source_image_url: str
    original_image_path: str
    generated_image_path: str


def clean_cell(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def safe_id(value: str) -> str:
    value = clean_cell(value)
    return re.sub(r"[^0-9A-Za-z_-]", "_", value)


def workbook_paths(
    input_path: Path,
    output_path: Path | None = None,
    output_root: Path | None = None,
) -> dict[str, Path]:
    root = output_root if output_root is not None else DEFAULT_OUTPUT_ROOT
    base_dir = root / f"{input_path.stem}_main_image_batch"
    originals_dir = base_dir / "original_images"
    generated_dir = base_dir / "generated_images"
    staging_dir = base_dir / "staging"
    review_queue_dir = base_dir / "review_queue"
    manifest_path = base_dir / "manifest.json"
    status_path = base_dir / "status_ledger.json"
    log_path = base_dir / "process_log.json"
    if output_path is None:
        output_path = base_dir / f"{input_path.stem}_main_image_output.xlsx"
    return {
        "base_dir": base_dir,
        "originals_dir": originals_dir,
        "generated_dir": generated_dir,
        "staging_dir": staging_dir,
        "review_queue_dir": review_queue_dir,
        "manifest_path": manifest_path,
        "status_path": status_path,
        "log_path": log_path,
        "output_xlsx": output_path,
    }


def find_headers(ws) -> dict[str, int]:
    headers: dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        value = clean_cell(ws.cell(row=HEADER_ROW, column=col).value)
        if value in REQUIRED_HEADERS:
            headers[value] = col
    missing = [name for name in REQUIRED_HEADERS if name not in headers]
    if missing:
        raise ValueError(f"Missing required headers on row {HEADER_ROW}: {', '.join(missing)}")
    return headers


def iter_items(input_path: Path, limit: int | None, paths: dict[str, Path]) -> tuple[list[Item], list[dict]]:
    wb = load_workbook(input_path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    headers = find_headers(ws)
    items: list[Item] = []
    skips: list[dict] = []
    seen_ids: set[str] = set()

    for row in range(HEADER_ROW + 1, ws.max_row + 1):
        local_id = clean_cell(ws.cell(row=row, column=headers["本地ID"]).value)
        title = clean_cell(ws.cell(row=row, column=headers["产品标题"]).value)
        image_url = clean_cell(ws.cell(row=row, column=headers["产品图片1"]).value)

        if not (local_id or title or image_url):
            continue
        if not local_id or not title or not image_url:
            skips.append({"row_number": row, "reason": "missing_required_field", "local_id": local_id})
            continue
        normalized_id = safe_id(local_id)
        if not normalized_id:
            skips.append({"row_number": row, "reason": "invalid_local_id", "local_id": local_id})
            continue
        if normalized_id in seen_ids:
            skips.append({"row_number": row, "reason": "duplicate_local_id", "local_id": local_id})
            continue

        seen_ids.add(normalized_id)
        original_path = paths["originals_dir"] / f"{normalized_id}.jpg"
        generated_path = paths["generated_dir"] / f"{normalized_id}.png"
        items.append(
            Item(
                row_number=row,
                local_id=normalized_id,
                title=title,
                source_image_url=image_url,
                original_image_path=str(original_path.resolve()),
                generated_image_path=str(generated_path.resolve()),
            )
        )
        if limit and len(items) >= limit:
            break

    wb.close()
    return items, skips


def download_image(url: str, out_path: Path) -> tuple[bool, str | None]:
    if out_path.exists() and out_path.stat().st_size > 0:
        return True, None
    request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            data = response.read()
        if not data:
            return False, "empty_response"
        out_path.write_bytes(data)
        return True, None
    except Exception as exc:  # noqa: BLE001 - log exact failure for batch review.
        return False, str(exc)


def write_json(path: Path, payload: object) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def read_json(path: Path, default: object) -> object:
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return default


def summarize_ledger(ledger: dict) -> dict[str, int]:
    counts: dict[str, int] = {}
    records = ledger.get("records", {}) if isinstance(ledger, dict) else {}
    for record in records.values():
        status = record.get("status", "pending")
        counts[status] = counts.get(status, 0) + 1
    counts["total"] = len(records)
    counts["unfinished"] = sum(
        count for status, count in counts.items() if status not in ("verified", "total", "unfinished")
    )
    return counts


def sync_status_ledger(paths: dict[str, Path], items: list[Item], skips: list[dict]) -> dict:
    existing = read_json(paths["status_path"], {})
    previous_records = {}
    if isinstance(existing, dict) and isinstance(existing.get("records"), dict):
        previous_records = existing["records"]

    records = {}
    for item in items:
        previous = previous_records.get(item.local_id, {})
        current_status = previous.get("status", "pending") if isinstance(previous, dict) else "pending"
        generated_path = Path(item.generated_image_path)
        if current_status == "verified" and not generated_path.exists():
            current_status = "pending"

        records[item.local_id] = {
            "row_number": item.row_number,
            "local_id": item.local_id,
            "title": item.title,
            "source_image_url": item.source_image_url,
            "original_image_path": item.original_image_path,
            "generated_image_path": item.generated_image_path,
            "status": current_status,
            "error": previous.get("error") if isinstance(previous, dict) else None,
            "updated_at": (previous.get("updated_at") if isinstance(previous, dict) else None) or utc_now(),
        }

    ledger = {
        "batch_dir": str(paths["base_dir"].resolve()),
        "updated_at": utc_now(),
        "records": records,
        "skips": skips,
    }
    write_json(paths["status_path"], ledger)
    return ledger


def prepare(input_path: Path, limit: int | None, output_path: Path | None, output_root: Path | None) -> dict:
    paths = workbook_paths(input_path, output_path, output_root)
    for key in ("base_dir", "originals_dir", "generated_dir", "staging_dir", "review_queue_dir"):
        paths[key].mkdir(parents=True, exist_ok=True)

    items, skips = iter_items(input_path, limit, paths)
    downloads = []
    for item in items:
        ok, error = download_image(item.source_image_url, Path(item.original_image_path))
        downloads.append({"local_id": item.local_id, "ok": ok, "error": error})
    ledger = sync_status_ledger(paths, items, skips)

    payload = {
        "input_xlsx": str(input_path.resolve()),
        "limit": limit,
        "items": [asdict(item) for item in items],
        "skips": skips,
        "downloads": downloads,
        "originals_dir": str(paths["originals_dir"].resolve()),
        "generated_dir": str(paths["generated_dir"].resolve()),
        "staging_dir": str(paths["staging_dir"].resolve()),
        "review_queue_dir": str(paths["review_queue_dir"].resolve()),
        "output_xlsx": str(paths["output_xlsx"].resolve()),
        "status_ledger": str(paths["status_path"].resolve()),
        "status_counts": summarize_ledger(ledger),
    }
    write_json(paths["manifest_path"], payload)
    write_json(paths["log_path"], {"skips": skips, "downloads": downloads})
    payload["manifest_path"] = str(paths["manifest_path"].resolve())
    payload["log_path"] = str(paths["log_path"].resolve())
    return payload


def set_widths(ws, widths: Iterable[int]) -> None:
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width


def normalize_generated_image(path: Path) -> tuple[bool, str | None]:
    if not path.exists():
        return False, "missing"
    try:
        image = Image.open(path).convert("RGB")
        if image.size != FINAL_IMAGE_SIZE:
            image = image.resize(FINAL_IMAGE_SIZE, Image.Resampling.LANCZOS)
        image.save(path, "PNG", optimize=True)
        return True, None
    except Exception as exc:  # noqa: BLE001 - log exact failure for batch review.
        return False, str(exc)


def load_or_prepare(input_path: Path, limit: int | None, output_path: Path | None, output_root: Path | None) -> tuple[dict, dict[str, Path]]:
    paths = workbook_paths(input_path, output_path, output_root)
    if not paths["status_path"].exists() or not paths["manifest_path"].exists():
        prepare(input_path, limit, output_path, output_root)
    ledger = read_json(paths["status_path"], {})
    if not isinstance(ledger, dict) or "records" not in ledger:
        prepare(input_path, limit, output_path, output_root)
        ledger = read_json(paths["status_path"], {})
    return ledger if isinstance(ledger, dict) else {}, paths


def status(input_path: Path, limit: int | None, output_path: Path | None, output_root: Path | None) -> dict:
    ledger, paths = load_or_prepare(input_path, limit, output_path, output_root)
    records = ledger.get("records", {})
    unfinished = [
        record for record in records.values()
        if record.get("status") != "verified"
    ]
    return {
        "status_ledger": str(paths["status_path"].resolve()),
        "manifest_path": str(paths["manifest_path"].resolve()),
        "output_xlsx": str(paths["output_xlsx"].resolve()),
        "counts": summarize_ledger(ledger),
        "next_unfinished": unfinished[:10],
    }


def next_items(
    input_path: Path,
    limit: int | None,
    output_path: Path | None,
    output_root: Path | None,
    count: int,
) -> dict:
    ledger, paths = load_or_prepare(input_path, limit, output_path, output_root)
    records = ledger.get("records", {})
    unfinished = [
        record for record in records.values()
        if record.get("status") != "verified"
    ]
    return {
        "status_ledger": str(paths["status_path"].resolve()),
        "count": min(count, len(unfinished)),
        "items": unfinished[:count],
    }


def mark_item(
    input_path: Path,
    limit: int | None,
    output_path: Path | None,
    output_root: Path | None,
    local_id: str,
    item_status: str,
    error: str | None,
) -> dict:
    ledger, paths = load_or_prepare(input_path, limit, output_path, output_root)
    records = ledger.get("records", {})
    if local_id not in records:
        raise SystemExit(f"Unknown local_id in status ledger: {local_id}")
    if item_status == "verified":
        ok, normalize_error = normalize_generated_image(Path(records[local_id]["generated_image_path"]))
        if not ok:
            raise SystemExit(f"Cannot mark verified; generated image is not valid: {normalize_error}")
    records[local_id]["status"] = item_status
    records[local_id]["error"] = error
    records[local_id]["updated_at"] = utc_now()
    ledger["updated_at"] = utc_now()
    write_json(paths["status_path"], ledger)
    return {
        "status_ledger": str(paths["status_path"].resolve()),
        "local_id": local_id,
        "status": item_status,
        "counts": summarize_ledger(ledger),
    }


def build(input_path: Path, limit: int | None, output_path: Path | None, output_root: Path | None) -> dict:
    payload = prepare(input_path, limit, output_path, output_root)
    output_xlsx = Path(payload["output_xlsx"])
    wb = Workbook()
    ws = wb.active
    ws.title = "main_images"
    headers = ["本地ID", "产品标题", "产品图片1", "主图"]
    ws.append(headers)

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    link_font = Font(color="0563C1", underline="single")
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font

    missing_generated = []
    image_normalization = []
    for item_data in payload["items"]:
        generated_path = Path(item_data["generated_image_path"])
        normalized, normalize_error = normalize_generated_image(generated_path)
        image_normalization.append(
            {"local_id": item_data["local_id"], "ok": normalized, "error": normalize_error}
        )
        if not normalized:
            missing_generated.append(item_data["local_id"])
        ws.append(
            [
                item_data["local_id"],
                item_data["title"],
                item_data["source_image_url"],
                str(generated_path.resolve()),
            ]
        )
        current_row = ws.max_row
        for col in (3, 4):
            cell = ws.cell(row=current_row, column=col)
            cell.hyperlink = cell.value
            cell.font = link_font

    set_widths(ws, [24, 80, 70, 90])
    ws.freeze_panes = "A2"
    output_xlsx.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_xlsx)
    payload["missing_generated"] = missing_generated
    payload["image_normalization"] = image_normalization
    payload["output_xlsx"] = str(output_xlsx.resolve())
    write_json(
        Path(payload["log_path"]),
        {
            "skips": payload["skips"],
            "downloads": payload["downloads"],
            "missing_generated": missing_generated,
            "image_normalization": image_normalization,
        },
    )
    return payload


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Prepare/build Mango Excel main-image batches.")
    subparsers = parser.add_subparsers(dest="command", required=True)
    for command in ("prepare", "build", "status", "next", "mark"):
        sub = subparsers.add_parser(command)
        sub.add_argument("--input", required=True, type=Path, help="Input .xlsx file")
        sub.add_argument("--limit", type=int, default=3, help="Number of valid products to process; use 0 for all")
        sub.add_argument("--output", type=Path, default=None, help="Output .xlsx path")
        sub.add_argument("--output-root", type=Path, default=None, help="Root folder for batch outputs")
        if command == "next":
            sub.add_argument("--count", type=int, default=20, help="Number of unfinished products to return")
        if command == "mark":
            sub.add_argument("--local-id", required=True, help="Product local ID to update")
            sub.add_argument("--status", required=True, choices=("pending", "verified", "failed"), help="New status")
            sub.add_argument("--error", default=None, help="Failure/error note")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = args.input.resolve()
    if input_path.suffix.lower() != ".xlsx":
        raise SystemExit("Input must be a .xlsx file.")
    if not input_path.exists():
        raise SystemExit(f"Input not found: {input_path}")

    limit = None if args.limit == 0 else args.limit
    output_root = args.output_root.resolve() if args.output_root else None

    if args.command == "prepare":
        result = prepare(input_path, limit, args.output, output_root)
    elif args.command == "build":
        result = build(input_path, limit, args.output, output_root)
    elif args.command == "status":
        result = status(input_path, limit, args.output, output_root)
    elif args.command == "next":
        result = next_items(input_path, limit, args.output, output_root, args.count)
    else:
        result = mark_item(input_path, limit, args.output, output_root, args.local_id, args.status, args.error)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
