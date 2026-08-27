#!/usr/bin/env python3
"""Query the Windows WPS Action Contract catalog without starting the bridge."""

import argparse
import json
import sys
from pathlib import Path

BRIDGE_DIR = Path(__file__).resolve().parents[1] / "bridge"
sys.path.insert(0, str(BRIDGE_DIR))

from action_catalog import (  # noqa: E402
    ActionCatalog,
    ActionManifestError,
    ActionNotSupportedError,
    AmbiguousActionError,
    UnknownActionError,
)


def _parser():
    parser = argparse.ArgumentParser(description="查询 WPS Action Contract")
    commands = parser.add_subparsers(dest="command", required=True)

    list_parser = commands.add_parser("list", help="列出 Action 摘要")
    list_parser.add_argument("--app", choices=("bridge", "excel", "ppt", "word"))

    search_parser = commands.add_parser("search", help="搜索 Action 名称和描述")
    search_parser.add_argument("query")
    search_parser.add_argument("--app", choices=("bridge", "excel", "ppt", "word"))

    describe_parser = commands.add_parser("describe", help="读取完整 Action Contract")
    describe_parser.add_argument("action")
    describe_parser.add_argument("--app", choices=("bridge", "excel", "ppt", "word"))
    return parser


def _print_json(value):
    print(json.dumps(value, ensure_ascii=False, indent=2))


def main(argv=None, catalog=None):
    args = _parser().parse_args(argv)
    try:
        catalog = catalog or ActionCatalog.from_path()
        if args.command == "list":
            items = catalog.list(owner=args.app)
            _print_json({"schemaVersion": catalog.schema_version, "actions": items, "count": len(items)})
            return 0
        if args.command == "search":
            items = catalog.search(args.query, owner=args.app)
            _print_json({"query": args.query, "actions": items, "count": len(items)})
            return 0
        _print_json(catalog.describe(args.action, owner=args.app))
        return 0
    except AmbiguousActionError as exc:
        _print_json({
            "success": False,
            "code": exc.code,
            "action": exc.action,
            "candidateOwners": exc.owners,
        })
        return 2
    except ActionNotSupportedError as exc:
        _print_json({
            "success": False,
            "code": exc.code,
            "action": exc.action,
            "owner": exc.owner,
            "candidateOwners": exc.supported_owners,
        })
        return 2
    except UnknownActionError as exc:
        _print_json({"success": False, "code": exc.code, "error": str(exc)})
        return 1
    except ActionManifestError as exc:
        _print_json({"success": False, "code": exc.code, "error": str(exc)})
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
