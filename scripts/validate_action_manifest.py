#!/usr/bin/env python3
"""Check the manifest against every public Windows ``Exec-*`` handler."""

import sys
from pathlib import Path

BRIDGE_DIR = Path(__file__).resolve().parents[1] / "bridge"
sys.path.insert(0, str(BRIDGE_DIR))

from action_catalog import (  # noqa: E402
    ActionCatalog,
    ActionManifestError,
    validate_windows_implementation_consistency,
)
import wps_excel  # noqa: E402
import wps_ppt  # noqa: E402
import wps_word  # noqa: E402


def main():
    try:
        catalog = ActionCatalog.from_path()
        validate_windows_implementation_consistency(
            catalog,
            {
                "excel": wps_excel.PS_BRIDGE_SCRIPT,
                "ppt": wps_ppt.PS_BRIDGE_SCRIPT,
                "word": wps_word.PS_BRIDGE_SCRIPT,
            },
        )
    except ActionManifestError as exc:
        print(f"{exc.code}: {exc}", file=sys.stderr)
        return 1
    print(f"Action manifest valid: {len(catalog.list())} contracts")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
