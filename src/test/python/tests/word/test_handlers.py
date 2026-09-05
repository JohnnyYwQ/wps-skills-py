import unittest

from wps_skills.core.action_session import ControllerContext
from wps_skills.word.adapter import WriterBackendActionFailure
from wps_skills.word.contracts import WORD_PRODUCTION_CONTRACT_SET
from wps_skills.word.handlers import WORD_HANDLERS


def context():
    return ControllerContext(
        request_id="request-1",
        trace_id="trace-1",
        deadline_at=100,
    )


class RecordingBackend:
    def __init__(self, result=None, failure=None):
        self.result = result if result is not None else {"observed": True}
        self.failure = failure
        self.calls = []

    def invoke(self, document, operation, request_context):
        self.calls.append((document, operation, request_context))
        if self.failure is not None:
            raise self.failure
        return self.result


class WordHandlerTests(unittest.TestCase):
    def test_every_production_required_action_has_one_private_operation(self):
        expected = {
            "writeContent": "insert_structured_body_content",
            "inspectDocument": "read_revision_coherent_snapshot",
            "findContent": "find_literal_body_content",
            "replaceContent": "replace_body_content",
            "insertTable": "insert_plain_text_table",
            "insertImage": "insert_embedded_image",
            "setHeaderFooter": "update_header_footer_stories",
            "setPageLayout": "update_page_layout",
            "insertBreak": "insert_body_break",
            "save": "save_existing_artifact",
            "exportPdf": "export_pdf_artifact",
        }
        required = {
            contract.name
            for contract in WORD_PRODUCTION_CONTRACT_SET.contracts
            if contract.binding_role == "required"
        }

        self.assertEqual(set(expected), required)
        self.assertEqual(set(expected), set(WORD_HANDLERS))
        for action, operation_name in expected.items():
            with self.subTest(action=action):
                backend = RecordingBackend()
                params = {"opaque": action}
                result = WORD_HANDLERS[action](
                    backend,
                    object(),
                    params,
                    context(),
                )
                self.assertEqual("succeeded", result.outcome)
                operation = backend.calls[0][1]
                self.assertEqual(operation_name, operation.name)
                self.assertEqual(params, dict(operation.arguments))

    def test_write_content_uses_closed_structured_private_operation(self):
        backend = RecordingBackend({
            "revisionBefore": "revision-1",
            "revisionAfter": "revision-2",
            "range": {"start": 0, "end": 5, "revision": "revision-2"},
        })
        document = object()
        params = {
            "anchor": {"kind": "documentEnd"},
            "blocks": ({
                "kind": "paragraph",
                "runs": ({"text": "Hello"},),
            },),
        }

        result = WORD_HANDLERS["writeContent"](
            backend,
            document,
            params,
            context(),
        )

        self.assertEqual("succeeded", result.outcome)
        operation = backend.calls[0][1]
        self.assertEqual("insert_structured_body_content", operation.name)
        self.assertEqual(params["anchor"], operation.arguments["anchor"])
        self.assertEqual(params["blocks"], operation.arguments["blocks"])

    def test_inspect_uses_a_private_operation_against_the_exact_document(self):
        backend = RecordingBackend({"revision": "revision-1"})
        document = object()
        params = {
            "scope": {"kind": "document"},
            "limits": {
                "maxTextCharacters": 100,
                "maxParagraphs": 10,
                "maxRuns": 20,
            },
        }

        result = WORD_HANDLERS["inspectDocument"](
            backend,
            document,
            params,
            context(),
        )

        self.assertEqual("succeeded", result.outcome)
        self.assertEqual("revision-1", result.data["revision"])
        self.assertIs(document, backend.calls[0][0])
        operation = backend.calls[0][1]
        self.assertEqual("read_revision_coherent_snapshot", operation.name)
        self.assertEqual(params["scope"], operation.arguments["scope"])
        self.assertEqual(params["limits"], operation.arguments["limits"])

    def test_save_uses_no_locator_and_returns_backend_observations(self):
        backend = RecordingBackend({
            "revisionBefore": "revision-1",
            "revisionAfter": "revision-1",
        })
        document = object()

        result = WORD_HANDLERS["save"](
            backend,
            document,
            {},
            context(),
        )

        self.assertEqual("succeeded", result.outcome)
        operation = backend.calls[0][1]
        self.assertEqual("save_existing_artifact", operation.name)
        self.assertEqual({}, dict(operation.arguments))

    def test_closed_backend_failure_maps_to_controller_result(self):
        backend = RecordingBackend(failure=WriterBackendActionFailure(
            outcome="failed",
            code="DOCUMENT_READ_ONLY",
            message="read only",
            binding_disposition="unchanged",
        ))

        result = WORD_HANDLERS["save"](
            backend,
            object(),
            {},
            context(),
        )

        self.assertEqual("failed", result.outcome)
        self.assertEqual("DOCUMENT_READ_ONLY", result.error.code)
        self.assertEqual("usable", result.controller_state)
        self.assertEqual("unchanged", result.binding_disposition)

    def test_unprovable_backend_failure_breaks_the_controller(self):
        backend = RecordingBackend(failure=WriterBackendActionFailure(
            outcome="unknown",
            code="OUTPUT_WRITE_FAILED",
            message="save result unknown",
            binding_disposition="unprovable",
        ))

        result = WORD_HANDLERS["save"](
            backend,
            object(),
            {},
            context(),
        )

        self.assertEqual("unknown", result.outcome)
        self.assertEqual("broken", result.controller_state)
        self.assertEqual("unprovable", result.binding_disposition)

    def test_non_object_backend_result_fails_closed(self):
        backend = RecordingBackend(result="invalid")
        with self.assertRaisesRegex(TypeError, "must be an object"):
            WORD_HANDLERS["inspectDocument"](
                backend,
                object(),
                {
                    "scope": {"kind": "document"},
                    "limits": {
                        "maxTextCharacters": 100,
                        "maxParagraphs": 10,
                        "maxRuns": 20,
                    },
                },
                context(),
            )


if __name__ == "__main__":
    unittest.main()
