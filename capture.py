#!/usr/bin/env python3

import json
import os
import sys
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright

TZ = ZoneInfo("Asia/Bangkok")

SLOT_BY_CRON = {
    "0 0 * * *": "07:00",
    "0 1 * * *": "08:00",
    "0 2 * * *": "09:00",
    "0 5 * * *": "12:00",
    "0 6 * * *": "13:00",
    "0 12 * * *": "19:00",
    "0 13 * * *": "20:00",
}

ROOT = Path(__file__).resolve().parent
CFG = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
NOW = datetime.now(TZ)


def resolve_slot():
    slot = (os.getenv("SLOT") or "").strip()

    if not slot:
        slot = SLOT_BY_CRON.get(
            (os.getenv("SCHEDULE") or "").strip()
        )

    if slot not in CFG["anchors"]:
        print(
            f"ERROR: unknown slot={slot!r}; "
            f"known={list(CFG['anchors'])}",
            file=sys.stderr,
        )
        sys.exit(1)

    return slot


def anchor_of(img):
    anchor = img.anchor

    if isinstance(anchor, str):
        return anchor

    return (
        f"{get_column_letter(anchor._from.col + 1)}"
        f"{anchor._from.row + 1}"
    )


def drop_old_image(ws, cell):
    for img in list(ws._images):
        if anchor_of(img) == cell:
            ws._images.remove(img)


def row_height(ws, row):
    # Excel row height is in points; convert approximately to pixels.
    points = ws.row_dimensions[row].height or 15
    return points * 1.333


def calculate_image_height(ws, cell, next_cell=None):
    """
    Calculate image height from the space before the next anchor.
    This prevents images at A88/A94 or A214/A220 from overlapping.
    """
    current_row = ws[cell].row

    if next_cell:
        next_row = ws[next_cell].row
        available = sum(
            row_height(ws, row)
            for row in range(current_row, next_row)
        )
        return max(70, int(available - 8))

    return 300


def main():
    slot = resolve_slot()
    date = NOW.strftime("%Y-%m-%d")

    reports_dir = ROOT / "reports"
    shots_dir = ROOT / "screenshots" / date / slot.replace(":", "")

    reports_dir.mkdir(parents=True, exist_ok=True)
    shots_dir.mkdir(parents=True, exist_ok=True)

    report_path = reports_dir / f"{date}.xlsx"
    template_path = ROOT / "template.xlsx"

    source = report_path if report_path.exists() else template_path

    if not source.exists():
        print(f"ERROR: Excel file not found: {source}", file=sys.stderr)
        sys.exit(1)

    state_path = ROOT / "storage_state.json"

    if not state_path.exists():
        print(
            "ERROR: storage_state.json missing",
            file=sys.stderr,
        )
        sys.exit(1)

    wb = load_workbook(source)

    sheet_name = CFG["sheet"]

    if sheet_name not in wb.sheetnames:
        print(
            f"ERROR: worksheet {sheet_name!r} not found; "
            f"available={wb.sheetnames}",
            file=sys.stderr,
        )
        sys.exit(1)

    ws = wb[sheet_name]

    # Preserve the order from config.json.
    metrics = list(CFG["urls"].items())
    anchors = CFG["anchors"][slot]

    if len(metrics) != len(anchors):
        print(
            f"ERROR: URL count ({len(metrics)}) does not match "
            f"anchor count ({len(anchors)})",
            file=sys.stderr,
        )
        sys.exit(1)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)

        context = browser.new_context(
            storage_state=str(state_path),
            viewport={"width": 1600, "height": 1000},
            locale="en-US",
            ignore_https_errors=True
        )

        page = context.new_page()
        page.set_default_timeout(60_000)

        for index, ((metric, url), cell) in enumerate(
            zip(metrics, anchors),
            start=1
        ):
            print(f"[{slot}] image {index}/7: {metric} -> {cell}")

            page.goto(
                url,
                wait_until="domcontentloaded",
                timeout=60_000
            )

            if "/login" in page.url.lower():
                raise RuntimeError(
                    "Grafana session expired; update "
                    "GRAFANA_STORAGE_STATE_B64"
                )

            page.wait_for_timeout(8_000)

            png_path = shots_dir / f"{index:02d}_{metric}.png"
            page.screenshot(
                path=str(png_path),
                full_page=True
            )

            next_cell = (
                anchors[index]
                if index < len(anchors)
                else None
            )

            image_height = calculate_image_height(
                ws,
                cell,
                next_cell
            )

            drop_old_image(ws, cell)

            image = XLImage(str(png_path))
            image.width = 520
            image.height = image_height

            ws.add_image(image, cell)

        context.close()
        browser.close()

    ws[CFG["date_cell"]] = NOW.strftime("%d %b %Y")
    wb.save(report_path)

    print(f"OK -> {report_path}")


if __name__ == "__main__":
    main()
