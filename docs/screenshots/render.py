#!/usr/bin/env python3
"""説明用 HTML を PNG にする。Google の実画面は撮影しない。"""
from __future__ import annotations
import subprocess, sys
from pathlib import Path

CHROME = "google-chrome"
ROOT = Path(__file__).resolve().parent
MOCKS = ROOT / "mocks"
JOBS = [
    ("input-before.html", "input-before.png", 1180, 560),
    ("dest-after.html", "dest-after.png", 1180, 500),
    ("log-after.html", "log-after.png", 1180, 420),
]

def render(html: Path, png: Path, w: int, h: int) -> None:
    cmd = [
        CHROME, "--headless=new", "--disable-gpu", "--no-sandbox",
        "--hide-scrollbars", "--force-device-scale-factor=1",
        f"--window-size={w},{h}", f"--screenshot={png}", html.resolve().as_uri(),
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    print(f"wrote {png}")

if __name__ == "__main__":
    for name, out, w, h in JOBS:
        render(MOCKS / name, ROOT / out, w, h)
