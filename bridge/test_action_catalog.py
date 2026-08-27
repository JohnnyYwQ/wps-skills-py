import json
import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path
import sys

from action_catalog import ActionCatalog, AmbiguousActionError, UnknownActionError

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))
import actions


def _schema(properties=None, required=None):
    return {
        "type": "object",
        "properties": properties or {},
        "required": required or [],
        "additionalProperties": False,
    }


def _contract(owner, action, description, risk="read"):
    return {
        "owner": owner,
        "action": action,
        "description": description,
        "parameters": _schema(),
        "result": _schema(),
        "prerequisites": [],
        "risk": risk,
    }


class ActionCatalogTests(unittest.TestCase):
    def _catalog(self):
        return ActionCatalog({
            "schema_version": 1,
            "actions": [
                _contract("excel", "findReplace", "替换单元格文本。", "write"),
                _contract("word", "findReplace", "替换文档文本。", "write"),
                _contract("ppt", "addSlide", "新增 chart 幻灯片。", "write"),
            ],
        })

    def test_catalog_indexes_identity_and_queries_without_controller_modules(self):
        manifest = {
            "schema_version": 1,
            "actions": [
                _contract("excel", "findReplace", "替换单元格文本。", "write"),
                _contract("word", "findReplace", "替换文档文本。", "write"),
                _contract("ppt", "addSlide", "新增 chart 幻灯片。", "write"),
            ],
        }
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "action_manifest.json"
            path.write_text(json.dumps(manifest), encoding="utf-8")
            catalog = ActionCatalog.from_path(path)

        self.assertEqual("ppt", catalog.get("ppt", "addSlide")["owner"])
        self.assertEqual(["excel", "word"], catalog.owners_for("findReplace"))
        self.assertEqual(["ppt"], [item["owner"] for item in catalog.search("chart")])
        self.assertEqual(
            {"owner", "action", "description", "risk"},
            set(catalog.list(owner="ppt")[0]),
        )
        with self.assertRaises(AmbiguousActionError) as ambiguous:
            catalog.describe("findReplace")
        self.assertEqual(["excel", "word"], ambiguous.exception.owners)
        with self.assertRaises(UnknownActionError):
            catalog.describe("missing")

    def test_describe_without_owner_reports_ambiguous_candidates(self):
        output = StringIO()
        with redirect_stdout(output):
            status = actions.main(
                ["describe", "findReplace"], catalog=self._catalog()
            )

        self.assertEqual(2, status)
        self.assertEqual(
            {
                "success": False,
                "code": "AMBIGUOUS_ACTION",
                "action": "findReplace",
                "candidateOwners": ["excel", "word"],
            },
            json.loads(output.getvalue()),
        )


if __name__ == "__main__":
    unittest.main()
