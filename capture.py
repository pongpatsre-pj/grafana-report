#!/usr/bin/env python3
"""
Capture Grafana panels and insert into a pre-made Excel template.
Slot is chosen by the GitHub Actions schedule (UTC) or workflow_dispatch input.
"""
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
    "0 0 * * *":  "07:00",
    "0 1 * * *":  "08:00",
    "0 2 * * *":  "09:00",
    "0 5 * * *":  "12:00",
    "0 6 * * *":  "13:00",
    "0 12 * * *": "19:00",
    "0 13 * * *": "20:00",
}

ROOT = Path(__file__).parent
CFG = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
NOW = datetime.now(TZ)


def resolve_slot() -> str:
    slot = os.getenv("SLOT") or SLOT_BY_CRON.get(os.getenv("SCHEDULE", ""))
    if slot not in CFG["anchors"]:
        print(f"ERROR: unknown slot {slot!r}. "
              f"known={list(CFG['anchors'])} cron={os.getenv('SCHEDULE')!r}", file=sys.stderr)
        sys.exit(1)
    return slot


def anchor_of(img) -> str:
    a = img.anchor
    if isinstance(a, str):
        return a
    return f"{get_column_letter(a._from.col + 1)}{a._from.row + 1}"


def drop_old_image(ws, cell: str) -> None:
    """Remove any image already anchored at `cell` so re-runs don't duplicate."""
    for img in list(ws._images):
        if anchor_of(img) == cell:
            ws._images.remove(img)


def main() -> None:
    slot = resolve_slot()
    date = NOW.strftime("%Y-%m-%d")

    reports_dir = ROOT / "reports"
    shots_dir = ROOT / "screenshots" / date / slot.replace(":", "")
    reports_dir.mkdir(exist_ok=True)
    shots_dir.mkdir(parents=True, exist_ok=True)

    report_path = reports_dir / f"{date}.xlsx"
    source = report_path if report_path.exists() else (ROOT / "template.xlsx")
    if not source.exists():
        print(f"ERROR: template not found: {source}", file=sys.stderr)
        sys.exit(1)

    state = ROOT / "storage_state.json"
    if not state.exists():
        print("ERROR: storage_state.json missing (restore from secret)", file=sys.stderr)
        sys.exit(1)

    wb = load_workbook(source)
    ws = wb[CFG["sheet"]]
    anchors = CFG["anchors"][slot]

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(
            storage_state=str(state),
            viewport={"width": 1600, "height": 1000},
            locale="en-US",
        )
        page = ctx.new_page()
        page.set_default_timeout(60_000)

        for metric, url in CFG["urls"].items():
            if metric not in anchors:
                continue
            print(f"[{slot}] capturing {metric} ...")
            page.goto(url, wait_until="domcontentloaded", timeout=60_000)
            page.wait_for_timeout(8_000)          # let panels finish rendering

            if "/login" in page.url:
                print("ERROR: Grafana session expired — refresh the secret", file=sys.stderr)
                browser.close()
                sys.exit(2)

            png = shots_dir / f"{metric}.png"
            page.screenshot(path=str(png), full_page=True)

            cell = anchors[metric]
            drop_old_image(ws, cell)
            img = XLImage(str(png))
            img.width, img.height = 520, 300      # resize to fit the block
            ws.add_image(img, cell)

        browser.close()

    ws[CFG["date_cell"]] = NOW.strftime("%d %b %Y")
    wb.save(report_path)
    print(f"OK -> {report_path}")


if __name__ == "__main__":
    main()
