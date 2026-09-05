#!/usr/bin/env python3
"""Repository entry point for Word Skill assembly."""

from pathlib import Path
import sys

MAIN_PYTHON = Path(__file__).resolve().parents[1] / "src" / "main" / "python"
sys.path.insert(0, str(MAIN_PYTHON))

from wps_skills.cli.build_word_skill import main  # noqa: E402


if __name__ == "__main__":
    raise SystemExit(main())
