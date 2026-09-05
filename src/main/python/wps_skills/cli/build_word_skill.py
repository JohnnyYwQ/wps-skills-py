"""Assemble a standalone Word Skill without maintaining a second runtime copy."""

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import tempfile


MAIN = Path(__file__).resolve().parents[3]


def build_word_skill(destination):
    destination = Path(destination).resolve()
    if destination.exists():
        raise FileExistsError(f"Destination already exists: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="wps-word-build-", dir=destination.parent) as temporary:
        stage = Path(temporary) / "wps-word"
        ignore = shutil.ignore_patterns("__pycache__", "*.pyc", ".DS_Store")
        shutil.copytree(MAIN / "resources" / "skills" / "wps-word", stage, ignore=ignore)
        runtime = stage / "runtime" / "src" / "main"
        shutil.copytree(MAIN / "python" / "wps_skills", runtime / "python" / "wps_skills", ignore=ignore)
        shutil.copytree(MAIN / "resources" / "wps_skills", runtime / "resources" / "wps_skills", ignore=ignore)
        files = {
            path.relative_to(stage).as_posix(): hashlib.sha256(path.read_bytes()).hexdigest()
            for path in sorted(stage.rglob("*")) if path.is_file()
        }
        (stage / "runtime" / "files.sha256.json").write_text(
            json.dumps(files, indent=2) + "\n", encoding="utf-8"
        )
        stage.rename(destination)
    return destination


def main(argv=None):
    parser = argparse.ArgumentParser(description="Assemble the complete wps-word Skill directory")
    parser.add_argument("--output", type=Path, default=Path("build/skills/wps-word"))
    args = parser.parse_args(argv)
    try:
        print(build_word_skill(args.output))
    except FileExistsError as exc:
        parser.exit(2, str(exc) + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
