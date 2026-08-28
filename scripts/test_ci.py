#!/usr/bin/env python3
"""Run every repository check that does not require a real WPS installation."""

from pathlib import Path
import os
import subprocess
import sys


ROOT = Path(__file__).resolve().parents[1]


def main():
    child_env = os.environ.copy()
    child_env["PYTHONUTF8"] = "1"
    checks = (
        (
            "Action manifest",
            [sys.executable, str(ROOT / "scripts" / "validate_action_manifest.py")],
        ),
        (
            "Python unit tests",
            [
                sys.executable,
                "-m",
                "unittest",
                "discover",
                "-s",
                str(ROOT / "bridge"),
                "-p",
                "test_*.py",
            ],
        ),
    )

    failed = []
    for label, command in checks:
        print(f"\n=== {label} ===", flush=True)
        completed = subprocess.run(
            command,
            cwd=ROOT,
            env=child_env,
            check=False,
        )
        if completed.returncode:
            failed.append((label, completed.returncode))

    if failed:
        print("\nNon-WPS CI checks failed:", file=sys.stderr)
        for label, returncode in failed:
            print(f"- {label}: exit code {returncode}", file=sys.stderr)
        return 1

    print("\nNon-WPS CI checks passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
