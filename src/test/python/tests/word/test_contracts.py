import copy
import json
import unittest

from wps_skills.core.action_session import (
    ActionAddress,
    ActionContract,
    ApplicationContractSet,
    ContractValidationError,
)
from wps_skills.word.contracts import (
    WORD_TARGET_ACTION_INDEX,
    WORD_TARGET_CONTRACT_SET,
    WORD_PRODUCTION_CONTRACT_SET,
)


EXPECTED_ACTIONS = {
    "createDocument": ("establish", "write", "documentEstablishment"),
    "openDocument": ("establish", "write", "documentEstablishment"),
    "writeContent": ("required", "write", "bodyContent"),
    "inspectDocument": ("required", "read", "bodyContent"),
    "findContent": ("required", "read", "bodyContent"),
    "replaceContent": ("required", "destructive", "bodyContent"),
    "insertTable": ("required", "write", "embeddedContent"),
    "insertImage": ("required", "write", "embeddedContent"),
    "setHeaderFooter": ("required", "write", "pageStructure"),
    "setPageLayout": ("required", "write", "pageStructure"),
    "insertBreak": ("required", "write", "pageStructure"),
    "save": ("required", "write", "persistence"),
    "saveAs": ("required", "destructive", "persistence"),
    "exportPdf": ("required", "destructive", "persistence"),
}

EXPECTED_REQUIRED_PARAMS = {
    "createDocument": (),
    "openDocument": ("path",),
    "writeContent": ("anchor", "blocks"),
    "inspectDocument": ("scope", "limits"),
    "findContent": ("query", "limit"),
    "replaceContent": ("target", "replacement"),
    "insertTable": ("anchor", "data", "headerRow"),
    "insertImage": (
        "anchor",
        "source",
        "placement",
        "size",
        "alternativeText",
    ),
    "setHeaderFooter": ("sections", "updates"),
    "setPageLayout": ("sections", "layout"),
    "insertBreak": ("anchor", "type"),
    "save": (),
    "saveAs": ("outputPath", "overwritePolicy"),
    "exportPdf": ("outputPath", "overwritePolicy"),
}


def content_range(start=0, end=1, revision="revision-1"):
    return {"start": start, "end": end, "revision": revision}


def point(value):
    return {"value": value, "unit": "pt"}


VALID_PARAMS = {
    "createDocument": {},
    "openDocument": {"path": "C:/docs/report.docx"},
    "writeContent": {
        "anchor": {"kind": "documentEnd"},
        "blocks": ({
            "kind": "paragraph",
            "runs": ({"text": "Hello"},),
        },),
    },
    "inspectDocument": {
        "scope": {"kind": "document"},
        "limits": {
            "maxTextCharacters": 1024,
            "maxParagraphs": 32,
            "maxRuns": 128,
        },
    },
    "findContent": {
        "query": {
            "scope": {"kind": "document"},
            "text": "Hello",
            "caseSensitive": False,
            "wholeWord": True,
        },
        "limit": 20,
    },
    "replaceContent": {
        "target": {
            "kind": "query",
            "query": {
                "scope": {"kind": "document"},
                "text": "old",
                "caseSensitive": True,
                "wholeWord": True,
            },
            "expectedMatchCount": 1,
        },
        "replacement": {"kind": "text", "runs": ({"text": "new"},)},
    },
    "insertTable": {
        "anchor": {"kind": "documentEnd"},
        "data": (("A", "B"), ("1", "2")),
        "headerRow": True,
    },
    "insertImage": {
        "anchor": {"kind": "documentEnd"},
        "source": {"kind": "file", "path": "C:/images/logo.png"},
        "placement": {"kind": "inline"},
        "size": {"kind": "intrinsic"},
        "alternativeText": {"kind": "decorative"},
    },
    "setHeaderFooter": {
        "sections": {"kind": "all", "revision": "revision-1"},
        "updates": ({
            "area": "header",
            "variant": "primary",
            "operation": {"kind": "replace", "text": "Report"},
        },),
    },
    "setPageLayout": {
        "sections": {"kind": "all", "revision": "revision-1"},
        "layout": {"orientation": "landscape"},
    },
    "insertBreak": {
        "anchor": {"kind": "documentEnd"},
        "type": "page",
    },
    "save": {},
    "saveAs": {
        "outputPath": "C:/docs/output.docx",
        "overwritePolicy": "failIfExists",
    },
    "exportPdf": {
        "outputPath": "C:/docs/output.pdf",
        "overwritePolicy": "replaceExisting",
    },
}


EMPTY_STORIES = tuple(
    {
        "area": area,
        "variant": variant,
        "exists": False,
        "linkToPrevious": False,
        "text": "",
    }
    for area in ("header", "footer")
    for variant in ("primary", "firstPage", "evenPages")
)

EMPTY_SECTION = {
    "index": 0,
    "layout": {
        "orientation": "portrait",
        "margins": {
            "top": point(72),
            "right": point(72),
            "bottom": point(72),
            "left": point(72),
        },
    },
    "headerFooter": {
        "firstPageEnabled": False,
        "evenPagesEnabled": False,
        "stories": EMPTY_STORIES,
    },
}

EMPTY_STRUCTURE = {
    "paragraphCount": 0,
    "headingCount": 0,
    "tableCount": 0,
    "inlineImageCount": 0,
    "floatingImageCount": 0,
    "sectionCount": 1,
    "pageBreakCount": 0,
    "sectionBreakCount": 0,
    "sections": (EMPTY_SECTION,),
}

VALID_RESULTS = {
    "createDocument": {
        "revision": "revision-1",
        "documentState": {
            "persistenceState": "unsaved",
            "readOnly": False,
        },
    },
    "openDocument": {
        "revision": "revision-1",
        "artifact": {
            "path": "C:/docs/report.docx",
            "format": "docx",
            "sizeBytes": 10,
        },
        "documentState": {
            "persistenceState": "saved",
            "readOnly": False,
        },
    },
    "writeContent": {
        "revisionBefore": "revision-1",
        "revisionAfter": "revision-2",
        "range": content_range(revision="revision-2"),
    },
    "inspectDocument": {
        "revision": "revision-1",
        "scopeRange": content_range(0, 0),
        "returnedRange": content_range(0, 0),
        "text": "",
        "paragraphs": (),
        "structure": EMPTY_STRUCTURE,
        "documentState": {
            "persistenceState": "unsaved",
            "readOnly": False,
        },
        "truncated": False,
        "remainingRange": None,
    },
    "findContent": {
        "revision": "revision-1",
        "scopeRange": content_range(0, 1),
        "matches": (),
        "truncated": False,
        "remainingRange": None,
    },
    "replaceContent": {
        "revisionBefore": "revision-1",
        "revisionAfter": "revision-2",
        "matchedCount": 1,
        "changedCount": 1,
        "ranges": (content_range(0, 3, revision="revision-2"),),
    },
    "insertTable": {
        "revisionBefore": "revision-1",
        "revisionAfter": "revision-2",
        "table": {
            "range": content_range(revision="revision-2"),
            "rowCount": 2,
            "columnCount": 2,
            "headerRow": True,
            "data": (("A", "B"), ("1", "2")),
        },
    },
    "insertImage": {
        "revisionBefore": "revision-1",
        "revisionAfter": "revision-2",
        "image": {
            "kind": "inline",
            "range": content_range(revision="revision-2"),
            "embedded": True,
            "source": {
                "mediaType": "image/png",
                "byteLength": 100,
                "sha256": "a" * 64,
            },
            "size": {"width": point(72), "height": point(72)},
            "alternativeText": {"kind": "decorative"},
        },
    },
    "setHeaderFooter": {
        "selectedSectionCount": 1,
        "changedCount": 1,
        "revisionBefore": "revision-1",
        "revisionAfter": "revision-2",
        "stories": ({
            "sectionIndex": 0,
            "area": "header",
            "variant": "primary",
            "variantEnabled": True,
            "exists": True,
            "linkToPrevious": False,
            "text": "Report",
        },),
    },
    "setPageLayout": {
        "selectedSectionCount": 1,
        "changedCount": 1,
        "revisionBefore": "revision-1",
        "revisionAfter": "revision-2",
        "sections": ({
            "index": 0,
            "layout": {
                "orientation": "landscape",
                "margins": {
                    "top": point(72),
                    "right": point(72),
                    "bottom": point(72),
                    "left": point(72),
                },
            },
        },),
    },
    "insertBreak": {
        "revisionBefore": "revision-1",
        "revisionAfter": "revision-2",
        "break": {
            "type": "page",
            "range": content_range(revision="revision-2"),
            "sectionCountBefore": 1,
            "sectionCountAfter": 1,
        },
    },
    "save": {
        "revisionBefore": "revision-1",
        "revisionAfter": "revision-1",
        "artifact": {
            "path": "C:/docs/report.docx",
            "format": "docx",
            "sizeBytes": 10,
        },
        "documentState": {
            "persistenceState": "saved",
            "readOnly": False,
        },
    },
    "saveAs": {
        "revisionBefore": "revision-1",
        "revisionAfter": "revision-1",
        "artifact": {
            "path": "C:/docs/output.docx",
            "format": "docx",
            "sizeBytes": 10,
        },
        "documentState": {
            "persistenceState": "saved",
            "readOnly": False,
        },
        "replacedExisting": False,
    },
    "exportPdf": {
        "revisionBefore": "revision-1",
        "revisionAfter": "revision-1",
        "artifact": {
            "path": "C:/docs/output.pdf",
            "format": "pdf",
            "sizeBytes": 10,
        },
        "documentStateBefore": {
            "persistenceState": "modified",
            "readOnly": False,
        },
        "documentStateAfter": {
            "persistenceState": "modified",
            "readOnly": False,
        },
        "replacedExisting": True,
    },
}


class WordContractSetTests(unittest.TestCase):
    def test_production_set_advertises_every_completed_action_except_save_as(self):
        self.assertEqual(
            (
                "createDocument",
                "openDocument",
                "writeContent",
                "inspectDocument",
                "findContent",
                "replaceContent",
                "insertTable",
                "insertImage",
                "setHeaderFooter",
                "setPageLayout",
                "insertBreak",
                "save",
                "exportPdf",
            ),
            WORD_PRODUCTION_CONTRACT_SET.action_names,
        )

    def test_action_inventory_matches_the_approved_word_interface(self):
        self.assertEqual(tuple(EXPECTED_ACTIONS), WORD_TARGET_CONTRACT_SET.action_names)
        self.assertEqual(tuple(EXPECTED_ACTIONS), tuple(
            entry.action for entry in WORD_TARGET_ACTION_INDEX
        ))
        for action, expected in EXPECTED_ACTIONS.items():
            with self.subTest(action=action):
                contract = WORD_TARGET_CONTRACT_SET.resolve(action)
                self.assertEqual(expected[0], contract.binding_role)
                self.assertEqual(expected[1], contract.risk)
                self.assertEqual(expected[2], contract.category)

    def test_shared_validator_enforces_composite_and_semantic_keywords(self):
        contracts = ApplicationContractSet(
            application="word",
            format_validators={
                "absoluteDocxPath": lambda value: (
                    value.startswith("C:/")
                    and value.casefold().endswith(".docx")
                ),
            },
            contracts=(ActionContract(
                name="fixture",
                binding_role="none",
                risk="read",
                parameters={
                    "oneOf": (
                        {
                            "type": "object",
                            "properties": {
                                "kind": {"const": "range"},
                                "start": {"type": "integer", "minimum": 0},
                                "end": {"type": "integer", "minimum": 0},
                            },
                            "required": ("kind", "start", "end"),
                            "additionalProperties": False,
                            "x-ordered": ("start", "end"),
                        },
                        {
                            "type": "object",
                            "properties": {
                                "kind": {"const": "path"},
                                "path": {
                                    "type": "string",
                                    "minLength": 1,
                                    "format": "absoluteDocxPath",
                                },
                            },
                            "required": ("kind", "path"),
                            "additionalProperties": False,
                        },
                    ),
                },
                result={"type": "object"},
            ),),
            allow_incomplete=True,
        )

        contracts.validate_params(
            "fixture",
            {"kind": "range", "start": 1, "end": 2},
        )
        contracts.validate_params(
            "fixture",
            {"kind": "path", "path": "C:/docs/report.docx"},
        )
        invalid = (
            {"kind": "range", "start": 3, "end": 2},
            {"kind": "path", "path": "relative/report.docx"},
            {"kind": "path", "path": "C:/docs/report.pdf"},
        )
        for params in invalid:
            with self.subTest(params=params):
                with self.assertRaises(ContractValidationError):
                    contracts.validate_params("fixture", params)

    def test_every_contract_has_complete_metadata_and_a_derived_index_entry(self):
        self.assertEqual(len(EXPECTED_ACTIONS), len(WORD_TARGET_ACTION_INDEX))
        for contract, entry in zip(
            WORD_TARGET_CONTRACT_SET.contracts,
            WORD_TARGET_ACTION_INDEX,
        ):
            with self.subTest(action=contract.name):
                self.assertTrue(contract.purpose)
                self.assertTrue(contract.category)
                self.assertTrue(contract.verification)
                self.assertTrue(contract.prerequisites)
                self.assertTrue(contract.constraints)
                self.assertTrue(contract.examples)
                self.assertTrue(contract.stable_errors)
                self.assertIn(
                    "WORD_CAPABILITY_UNAVAILABLE",
                    contract.stable_errors,
                )
                self.assertEqual(
                    len(contract.stable_errors),
                    len(set(contract.stable_errors)),
                )
                self.assertEqual(
                    EXPECTED_REQUIRED_PARAMS[contract.name],
                    contract.required_parameters,
                )
                self.assertEqual(
                    {
                        "action",
                        "purpose",
                        "category",
                        "bindingRole",
                        "risk",
                        "requiredParams",
                    },
                    set(entry.to_wire()),
                )
                full = WORD_TARGET_CONTRACT_SET.resolve_wire(contract.name)
                self.assertEqual(
                    {"app": "word", "action": contract.name},
                    full["address"],
                )
                self.assertEqual(contract.purpose, full["purpose"])
                self.assertEqual(list(contract.stable_errors), full["errors"])
                self.assertEqual(
                    list(contract.prerequisites),
                    full["prerequisites"],
                )
                self.assertEqual(list(contract.constraints), full["constraints"])
                for example in contract.examples:
                    WORD_TARGET_CONTRACT_SET.validate_params(
                        contract.name,
                        example["params"],
                    )

    def test_all_approved_examples_validate_against_their_contracts(self):
        self.assertEqual(set(EXPECTED_ACTIONS), set(VALID_PARAMS))
        self.assertEqual(set(EXPECTED_ACTIONS), set(VALID_RESULTS))
        for action in EXPECTED_ACTIONS:
            with self.subTest(action=action, value="params"):
                WORD_TARGET_CONTRACT_SET.validate_params(action, VALID_PARAMS[action])
            with self.subTest(action=action, value="result"):
                WORD_TARGET_CONTRACT_SET.validate_result(
                    action,
                    VALID_RESULTS[action],
                    params=VALID_PARAMS[action],
                )

    def test_high_risk_invalid_shapes_fail_before_an_adapter_exists(self):
        invalid = {
            "createDocument": {"path": "C:/docs/report.docx"},
            "openDocument": {"path": "relative/report.docx"},
            "writeContent": {
                "anchor": {"kind": "documentEnd"},
                "blocks": (),
            },
            "inspectDocument": {
                "scope": {"kind": "document"},
                "limits": {
                    "maxTextCharacters": 0,
                    "maxParagraphs": 1,
                    "maxRuns": 1,
                },
            },
            "findContent": {
                **VALID_PARAMS["findContent"],
                "limit": 201,
            },
            "replaceContent": {
                "target": {
                    "kind": "range",
                    "range": content_range(1, 1),
                },
                "replacement": {"kind": "delete"},
            },
            "insertTable": {
                "anchor": {"kind": "documentEnd"},
                "data": (("A", "B"), ("1",)),
                "headerRow": False,
            },
            "insertImage": {
                **VALID_PARAMS["insertImage"],
                "source": {"kind": "file", "path": "C:/images/logo.gif"},
            },
            "setHeaderFooter": {
                "sections": {"kind": "all", "revision": "revision-1"},
                "updates": (
                    {
                        "area": "header",
                        "variant": "primary",
                        "operation": {"kind": "clear"},
                    },
                    {
                        "area": "header",
                        "variant": "primary",
                        "operation": {"kind": "clear"},
                    },
                ),
            },
            "setPageLayout": {
                "sections": {"kind": "all", "revision": "revision-1"},
                "layout": {},
            },
            "insertBreak": {
                "anchor": {"kind": "documentEnd"},
                "type": "line",
            },
            "save": {"outputPath": "C:/docs/output.docx"},
            "saveAs": {
                "outputPath": "C:/docs/output.pdf",
                "overwritePolicy": "failIfExists",
            },
            "exportPdf": {
                "outputPath": "C:/docs/output.docx",
                "overwritePolicy": "failIfExists",
            },
        }
        for action, params in invalid.items():
            with self.subTest(action=action):
                with self.assertRaises(ContractValidationError):
                    WORD_TARGET_CONTRACT_SET.validate_params(action, params)

    def test_contracts_expose_no_document_routing_fields(self):
        forbidden = {
            "documentId",
            "targetId",
            "binding",
            "lease",
            "activeDocument",
            "displayName",
            "registryKey",
            "candidate",
        }

        def property_names(schema):
            names = set(schema.get("properties", {}))
            for child in schema.get("properties", {}).values():
                names.update(property_names(child))
            if "items" in schema:
                names.update(property_names(schema["items"]))
            for keyword in ("oneOf", "anyOf"):
                for child in schema.get(keyword, ()):
                    names.update(property_names(child))
            return names

        for contract in WORD_TARGET_CONTRACT_SET.contracts:
            with self.subTest(action=contract.name):
                self.assertTrue(
                    forbidden.isdisjoint(property_names(contract.parameters))
                )

    def test_index_and_full_contracts_are_json_serializable(self):
        json.dumps([entry.to_wire() for entry in WORD_TARGET_ACTION_INDEX])
        for action in EXPECTED_ACTIONS:
            with self.subTest(action=action):
                encoded = json.dumps(WORD_TARGET_CONTRACT_SET.resolve_wire(action))
                self.assertIn(f'"action": "{action}"', encoded)

    def test_contract_schemas_are_deeply_immutable(self):
        schema = WORD_TARGET_CONTRACT_SET.resolve("replaceContent").parameters
        with self.assertRaises(TypeError):
            schema["oneOf"][0]["properties"]["extra"] = {"type": "string"}

    def test_non_fixture_contract_sets_require_complete_valid_contracts(self):
        incomplete = ActionContract(
            name="incomplete",
            binding_role="none",
            risk="read",
            parameters={"type": "object"},
            result={"type": "object"},
        )
        with self.assertRaisesRegex(ValueError, "missing purpose"):
            ApplicationContractSet(
                application="word",
                contracts=(incomplete,),
            )

        malformed = ActionContract(
            name="malformed",
            binding_role="none",
            risk="read",
            purpose="Malformed fixture",
            category="fixture",
            stable_errors=("INVALID_PARAMS",),
            verification="Validate it.",
            prerequisites=("Fixture prerequisite.",),
            constraints=("Fixture constraint.",),
            examples=({"params": {}},),
            parameters={"type": "object", "unknownKeyword": True},
            result={"type": "object"},
        )
        with self.assertRaisesRegex(ValueError, "unknown keyword"):
            ApplicationContractSet(
                application="word",
                contracts=(malformed,),
            )

    def test_every_object_schema_is_closed_and_errors_are_stable_codes(self):
        def assert_closed(schema, path):
            if schema.get("type") == "object":
                self.assertIs(
                    False,
                    schema.get("additionalProperties"),
                    msg=f"open object schema at {path}",
                )
            for name, child in schema.get("properties", {}).items():
                assert_closed(child, f"{path}.{name}")
            if "items" in schema:
                assert_closed(schema["items"], f"{path}[]")
            for keyword in ("oneOf", "anyOf"):
                for index, child in enumerate(schema.get(keyword, ())):
                    assert_closed(child, f"{path}.{keyword}[{index}]")

        for contract in WORD_TARGET_CONTRACT_SET.contracts:
            with self.subTest(action=contract.name):
                assert_closed(contract.parameters, f"{contract.name}.params")
                assert_closed(contract.result, f"{contract.name}.result")
                for code in contract.stable_errors:
                    self.assertRegex(code, r"^[A-Z][A-Z0-9_]*$")

    def test_batch_resolution_reports_every_requested_word_address(self):
        result = WORD_TARGET_CONTRACT_SET.batch_resolve((
            ActionAddress(app="word", action="createDocument"),
            ActionAddress(app="word", action="notAnAction"),
            ActionAddress(app="excel", action="createDocument"),
        ))

        self.assertEqual("partial", result["status"])
        self.assertEqual(
            ["resolved", "unresolved", "unresolved"],
            [item["status"] for item in result["items"]],
        )
        self.assertEqual(
            "createDocument",
            result["items"][0]["contract"]["address"]["action"],
        )
        self.assertEqual(
            "UNKNOWN_ACTION",
            result["items"][1]["error"]["code"],
        )
        self.assertEqual(
            "SESSION_APP_MISMATCH",
            result["items"][2]["error"]["code"],
        )

    def test_inspection_accepts_the_documented_normalized_structure_markers(self):
        result = dict(VALID_RESULTS["inspectDocument"])
        result["text"] = "cell\trow\nmanual\u2028break\fobject\ufffc"

        WORD_TARGET_CONTRACT_SET.validate_result(
            "inspectDocument",
            result,
            params=VALID_PARAMS["inspectDocument"],
        )

    def test_structured_text_limit_counts_payload_not_schema_tags(self):
        params = {
            "anchor": {"kind": "documentEnd"},
            "blocks": ({
                "kind": "text",
                "runs": tuple(
                    {"text": "a" * 32768}
                    for _ in range(8)
                ),
            },),
        }

        WORD_TARGET_CONTRACT_SET.validate_params("writeContent", params)

    def test_text_runs_support_distinct_western_and_east_asian_fonts(self):
        params = copy.deepcopy(VALID_PARAMS["writeContent"])
        params["blocks"][0]["runs"][0]["format"] = {
            "westernFontFamily": "Times New Roman",
            "eastAsiaFontFamily": "Microsoft YaHei",
        }

        WORD_TARGET_CONTRACT_SET.validate_params("writeContent", params)

        params["blocks"][0]["runs"][0]["format"][
            "fontFamily"
        ] = "Arial"
        with self.assertRaisesRegex(
            ContractValidationError,
            "cannot be combined",
        ):
            WORD_TARGET_CONTRACT_SET.validate_params(
                "writeContent",
                params,
            )

    def test_successful_image_dimensions_must_be_positive(self):
        result = copy.deepcopy(VALID_RESULTS["insertImage"])
        result["image"]["size"]["width"]["value"] = 0

        with self.assertRaises(ContractValidationError):
            WORD_TARGET_CONTRACT_SET.validate_result(
                "insertImage",
                result,
                params=VALID_PARAMS["insertImage"],
            )

    def test_table_coherence_uses_json_array_semantics(self):
        params = copy.deepcopy(VALID_PARAMS["insertTable"])
        params["data"] = [["A", "B"], ["1", "2"]]

        WORD_TARGET_CONTRACT_SET.validate_result(
            "insertTable",
            VALID_RESULTS["insertTable"],
            params=params,
        )

    def test_semantic_result_validation_uses_the_original_request(self):
        invalid = []

        inspect_result = copy.deepcopy(VALID_RESULTS["inspectDocument"])
        inspect_result["structure"]["sectionCount"] = 2
        invalid.append((
            "inspectDocument",
            VALID_PARAMS["inspectDocument"],
            inspect_result,
        ))

        find_params = copy.deepcopy(VALID_PARAMS["findContent"])
        find_params["limit"] = 1
        find_result = copy.deepcopy(VALID_RESULTS["findContent"])
        find_result["matches"] = (
            {"range": content_range(0, 1), "text": "a"},
            {"range": content_range(1, 2), "text": "a"},
        )
        invalid.append(("findContent", find_params, find_result))

        write_result = copy.deepcopy(VALID_RESULTS["writeContent"])
        write_result["range"]["revision"] = "wrong-revision"
        invalid.append((
            "writeContent",
            VALID_PARAMS["writeContent"],
            write_result,
        ))

        replace_result = copy.deepcopy(VALID_RESULTS["replaceContent"])
        replace_result["changedCount"] = 2
        invalid.append((
            "replaceContent",
            VALID_PARAMS["replaceContent"],
            replace_result,
        ))

        replacement_noop_result = copy.deepcopy(
            VALID_RESULTS["replaceContent"]
        )
        replacement_noop_result["revisionAfter"] = "revision-1"
        replacement_noop_result["changedCount"] = 0
        replacement_noop_result["ranges"] = (content_range(),)
        invalid.append((
            "replaceContent",
            VALID_PARAMS["replaceContent"],
            replacement_noop_result,
        ))

        empty_replacement_range_result = copy.deepcopy(
            VALID_RESULTS["replaceContent"]
        )
        empty_replacement_range_result["ranges"] = (
            content_range(0, 0, revision="revision-2"),
        )
        invalid.append((
            "replaceContent",
            VALID_PARAMS["replaceContent"],
            empty_replacement_range_result,
        ))

        save_as_result = copy.deepcopy(VALID_RESULTS["saveAs"])
        save_as_result["replacedExisting"] = True
        invalid.append((
            "saveAs",
            VALID_PARAMS["saveAs"],
            save_as_result,
        ))

        ranged_inspect_params = copy.deepcopy(
            VALID_PARAMS["inspectDocument"]
        )
        ranged_inspect_params["scope"] = {
            "kind": "range",
            "range": content_range(0, 1),
        }
        invalid.append((
            "inspectDocument",
            ranged_inspect_params,
            copy.deepcopy(VALID_RESULTS["inspectDocument"]),
        ))

        ranged_find_params = copy.deepcopy(VALID_PARAMS["findContent"])
        ranged_find_params["query"]["scope"] = {
            "kind": "range",
            "range": content_range(0, 2),
        }
        invalid.append((
            "findContent",
            ranged_find_params,
            copy.deepcopy(VALID_RESULTS["findContent"]),
        ))

        truncated_find_result = copy.deepcopy(VALID_RESULTS["findContent"])
        truncated_find_result["truncated"] = True
        truncated_find_result["remainingRange"] = content_range(0, 1)
        invalid.append((
            "findContent",
            VALID_PARAMS["findContent"],
            truncated_find_result,
        ))

        empty_suffix_inspect_result = copy.deepcopy(
            VALID_RESULTS["inspectDocument"]
        )
        empty_suffix_inspect_result["truncated"] = True
        empty_suffix_inspect_result["remainingRange"] = content_range(0, 0)
        invalid.append((
            "inspectDocument",
            VALID_PARAMS["inspectDocument"],
            empty_suffix_inspect_result,
        ))

        break_params = copy.deepcopy(VALID_PARAMS["insertBreak"])
        break_params["type"] = "sectionNextPage"
        invalid.append((
            "insertBreak",
            break_params,
            copy.deepcopy(VALID_RESULTS["insertBreak"]),
        ))

        incomplete_header_result = copy.deepcopy(
            VALID_RESULTS["setHeaderFooter"]
        )
        incomplete_header_result["selectedSectionCount"] = 2
        invalid.append((
            "setHeaderFooter",
            VALID_PARAMS["setHeaderFooter"],
            incomplete_header_result,
        ))

        incomplete_layout_result = copy.deepcopy(
            VALID_RESULTS["setPageLayout"]
        )
        incomplete_layout_result["selectedSectionCount"] = 2
        invalid.append((
            "setPageLayout",
            VALID_PARAMS["setPageLayout"],
            incomplete_layout_result,
        ))

        stale_write_params = copy.deepcopy(VALID_PARAMS["writeContent"])
        stale_write_params["anchor"] = {
            "kind": "after",
            "range": content_range(revision="revision-stale"),
        }
        invalid.append((
            "writeContent",
            stale_write_params,
            copy.deepcopy(VALID_RESULTS["writeContent"]),
        ))

        empty_write_result = copy.deepcopy(VALID_RESULTS["writeContent"])
        empty_write_result["range"]["end"] = 0
        invalid.append((
            "writeContent",
            VALID_PARAMS["writeContent"],
            empty_write_result,
        ))

        delete_params = copy.deepcopy(VALID_PARAMS["replaceContent"])
        delete_params["replacement"] = {"kind": "delete"}
        delete_noop_result = copy.deepcopy(VALID_RESULTS["replaceContent"])
        delete_noop_result["revisionAfter"] = "revision-1"
        delete_noop_result["changedCount"] = 0
        delete_noop_result["ranges"] = (content_range(),)
        invalid.append((
            "replaceContent",
            delete_params,
            delete_noop_result,
        ))

        exact_delete_params = {
            "target": {
                "kind": "range",
                "range": content_range(10, 20),
            },
            "replacement": {"kind": "delete"},
        }
        misplaced_delete_result = {
            "revisionBefore": "revision-1",
            "revisionAfter": "revision-2",
            "matchedCount": 1,
            "changedCount": 1,
            "ranges": (content_range(999, 999, revision="revision-2"),),
        }
        invalid.append((
            "replaceContent",
            exact_delete_params,
            misplaced_delete_result,
        ))

        wrong_table_result = copy.deepcopy(VALID_RESULTS["insertTable"])
        wrong_table_result["table"]["headerRow"] = False
        invalid.append((
            "insertTable",
            VALID_PARAMS["insertTable"],
            wrong_table_result,
        ))

        floating_image_params = copy.deepcopy(VALID_PARAMS["insertImage"])
        floating_image_params["placement"] = {
            "kind": "floating",
            "wrap": "square",
            "horizontal": {
                "relativeTo": "page",
                "offset": point(0),
            },
            "vertical": {
                "relativeTo": "page",
                "offset": point(0),
            },
        }
        invalid.append((
            "insertImage",
            floating_image_params,
            copy.deepcopy(VALID_RESULTS["insertImage"]),
        ))

        wrong_header_result = copy.deepcopy(
            VALID_RESULTS["setHeaderFooter"]
        )
        wrong_header_result["stories"][0]["text"] = "wrong"
        invalid.append((
            "setHeaderFooter",
            VALID_PARAMS["setHeaderFooter"],
            wrong_header_result,
        ))

        wrong_layout_result = copy.deepcopy(VALID_RESULTS["setPageLayout"])
        wrong_layout_result["sections"][0]["layout"][
            "orientation"
        ] = "portrait"
        invalid.append((
            "setPageLayout",
            VALID_PARAMS["setPageLayout"],
            wrong_layout_result,
        ))

        invalid_section_break = copy.deepcopy(VALID_RESULTS["insertBreak"])
        invalid_section_break["break"].update({
            "type": "sectionNextPage",
            "sectionCountAfter": 2,
            "followingSectionIndex": 999,
        })
        invalid.append((
            "insertBreak",
            break_params,
            invalid_section_break,
        ))

        empty_suffix_find_params = copy.deepcopy(VALID_PARAMS["findContent"])
        empty_suffix_find_params["limit"] = 1
        empty_suffix_find_result = copy.deepcopy(VALID_RESULTS["findContent"])
        empty_suffix_find_result.update({
            "scopeRange": content_range(0, 5),
            "matches": ({
                "range": content_range(0, 5),
                "text": "Hello",
            },),
            "truncated": True,
            "remainingRange": content_range(5, 5),
        })
        invalid.append((
            "findContent",
            empty_suffix_find_params,
            empty_suffix_find_result,
        ))

        wrong_open_result = copy.deepcopy(VALID_RESULTS["openDocument"])
        wrong_open_result["artifact"]["path"] = "C:/docs/other.docx"
        invalid.append((
            "openDocument",
            VALID_PARAMS["openDocument"],
            wrong_open_result,
        ))

        wrong_export_result = copy.deepcopy(VALID_RESULTS["exportPdf"])
        wrong_export_result["artifact"]["path"] = "C:/docs/other.pdf"
        invalid.append((
            "exportPdf",
            VALID_PARAMS["exportPdf"],
            wrong_export_result,
        ))

        wrong_save_as_result = copy.deepcopy(VALID_RESULTS["saveAs"])
        wrong_save_as_result["artifact"]["path"] = "C:/docs/other.docx"
        invalid.append((
            "saveAs",
            VALID_PARAMS["saveAs"],
            wrong_save_as_result,
        ))

        changed_during_save_result = copy.deepcopy(VALID_RESULTS["save"])
        changed_during_save_result["revisionAfter"] = "revision-2"
        invalid.append((
            "save",
            VALID_PARAMS["save"],
            changed_during_save_result,
        ))

        state_changing_export_result = copy.deepcopy(
            VALID_RESULTS["exportPdf"]
        )
        state_changing_export_result["documentStateAfter"][
            "persistenceState"
        ] = "saved"
        invalid.append((
            "exportPdf",
            VALID_PARAMS["exportPdf"],
            state_changing_export_result,
        ))

        paragraph_format = {
            "alignment": None,
            "lineSpacing": None,
            "spaceBeforePt": None,
            "spaceAfterPt": None,
            "leftIndentPt": None,
            "rightIndentPt": None,
            "firstLineIndentPt": None,
        }

        def paragraph(start, end, text):
            return {
                "kind": "paragraph",
                "range": content_range(start, end),
                "complete": True,
                "text": text,
                "runs": (),
                "format": paragraph_format,
            }

        unordered_inspect_result = copy.deepcopy(
            VALID_RESULTS["inspectDocument"]
        )
        unordered_inspect_result.update({
            "scopeRange": content_range(0, 4),
            "returnedRange": content_range(0, 4),
            "text": "abcd",
            "paragraphs": (
                paragraph(2, 4, "cd"),
                paragraph(0, 2, "ab"),
            ),
        })
        unordered_inspect_result["structure"] = copy.deepcopy(EMPTY_STRUCTURE)
        unordered_inspect_result["structure"]["paragraphCount"] = 2
        invalid.append((
            "inspectDocument",
            VALID_PARAMS["inspectDocument"],
            unordered_inspect_result,
        ))

        for action, params, result in invalid:
            with self.subTest(action=action):
                with self.assertRaises(ContractValidationError):
                    WORD_TARGET_CONTRACT_SET.validate_result(
                        action,
                        result,
                        params=params,
                    )

    def test_json_constants_and_host_paths_use_json_and_host_semantics(self):
        create_result = copy.deepcopy(VALID_RESULTS["createDocument"])
        create_result["documentState"]["readOnly"] = 0
        with self.assertRaises(ContractValidationError):
            WORD_TARGET_CONTRACT_SET.validate_result(
                "createDocument",
                create_result,
                params=VALID_PARAMS["createDocument"],
            )

        with self.assertRaises(ContractValidationError):
            WORD_TARGET_CONTRACT_SET.validate_params(
                "openDocument",
                {"path": "C:/docs/bad\0.docx"},
            )

        for path in (
            "C:/docs/\ud800report.docx",
            "C:/docs/report?.docx",
        ):
            with self.subTest(path=repr(path)):
                with self.assertRaises(ContractValidationError):
                    WORD_TARGET_CONTRACT_SET.validate_params(
                        "openDocument",
                        {"path": path},
                    )

    def test_link_to_previous_cannot_statically_select_section_zero(self):
        params = {
            "sections": {"kind": "all", "revision": "revision-1"},
            "updates": ({
                "area": "header",
                "variant": "primary",
                "operation": {"kind": "linkToPrevious"},
            },),
        }

        with self.assertRaisesRegex(
            ContractValidationError,
            "cannot target section 0",
        ):
            WORD_TARGET_CONTRACT_SET.validate_params("setHeaderFooter", params)

    def test_structured_write_rejects_object_and_structure_markers(self):
        for marker in ("\ufffc", "\u2028", "\u2029", "\x7f"):
            params = copy.deepcopy(VALID_PARAMS["writeContent"])
            params["blocks"][0]["runs"][0]["text"] = f"Hello{marker}"
            with self.subTest(marker=repr(marker)):
                with self.assertRaises(ContractValidationError):
                    WORD_TARGET_CONTRACT_SET.validate_params("writeContent", params)

        query_params = copy.deepcopy(VALID_PARAMS["findContent"])
        query_params["query"]["text"] = "Hello\u2029"
        with self.assertRaises(ContractValidationError):
            WORD_TARGET_CONTRACT_SET.validate_params("findContent", query_params)

        table_params = copy.deepcopy(VALID_PARAMS["insertTable"])
        table_params["data"] = (("Hello\u2029",),)
        with self.assertRaises(ContractValidationError):
            WORD_TARGET_CONTRACT_SET.validate_params("insertTable", table_params)

    def test_semantic_contract_results_require_the_original_params(self):
        with self.assertRaisesRegex(
            ContractValidationError,
            "original params are required",
        ):
            WORD_TARGET_CONTRACT_SET.validate_result(
                "save",
                VALID_RESULTS["save"],
            )


if __name__ == "__main__":
    unittest.main()
