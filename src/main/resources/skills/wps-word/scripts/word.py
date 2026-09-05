#!/usr/bin/env python3
"""Word Skill entry point, usable from its source tree or assembled distribution."""

from pathlib import Path
import sys


SKILL_ROOT = Path(__file__).resolve().parents[1]
SOURCES = [SKILL_ROOT / "runtime" / "src" / "main" / "python"]
if len(SKILL_ROOT.parents) > 2:
    SOURCES.append(SKILL_ROOT.parents[2] / "python")
for source in SOURCES:
    if (source / "wps_skills" / "word" / "skill.py").is_file():
        sys.path.insert(0, str(source))
        break
else:
    raise RuntimeError("Word Skill runtime is missing; use the complete assembled wps-word directory")

from wps_skills.word.skill import ActionFailed, SessionClientError, main, open_session  # noqa: E402,F401


if __name__ == "__main__":
    raise SystemExit(main())
