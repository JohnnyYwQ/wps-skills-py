#!/usr/bin/env python3
"""Write the embedded Windows PowerShell bridges for parser validation."""

import argparse
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
BRIDGE_DIR = ROOT / "bridge"
sys.path.insert(0, str(BRIDGE_DIR))

import wps_excel  # noqa: E402
import wps_ppt  # noqa: E402
import wps_word  # noqa: E402


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args(argv)
    args.output_dir.mkdir(parents=True, exist_ok=True)

    for name, script in (
        ("excel", wps_excel.PS_BRIDGE_SCRIPT),
        ("ppt", wps_ppt.PS_BRIDGE_SCRIPT),
        ("word", wps_word.PS_BRIDGE_SCRIPT),
    ):
        path = args.output_dir / f"wps_{name}.ps1"
        path.write_text(script, encoding="utf-8-sig")
        print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
