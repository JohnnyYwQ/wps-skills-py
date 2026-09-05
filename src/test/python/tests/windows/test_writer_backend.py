import unittest

from wps_skills.core.action_session import (
    ActionAddress,
    ActionRequest,
    ActionSession,
    ControllerContext,
    DocumentResourceCleanup,
    TraceContext,
)
from wps_skills.windows.writer_backend import (
    WindowsWriterBackend,
    WindowsWriterBridge,
    WindowsWriterDocument,
)
from wps_skills.word.adapter import (
    WordAdapter,
    WriterBackend,
    WriterBackendActionFailure,
    WriterBackendOperation,
    WriterDefiniteEstablishFailure,
    WriterUnprovableEstablishFailure,
)
from wps_skills.word.contracts import WORD_TARGET_CONTRACT_SET
from wps_skills.word.handlers import WORD_HANDLERS


def context():
    return ControllerContext(
        request_id="request-1",
        trace_id="trace-1",
        deadline_at=100,
    )


class RecordingBridge:
    def __init__(self, replies=None):
        self.replies = dict(replies or {})
        self.calls = []

    def execute(self, operation, arguments, request_context):
        self.calls.append((operation, dict(arguments), request_context))
        if operation == "prepare_existing_document":
            return self.replies.get(operation, {
                "preparationId": "preparation-1",
                "coordinationIdentity": "file-document-1",
            })
        if operation == "prepare_new_document":
            return self.replies.get(operation, {
                "preparationId": "preparation-1",
                "coordinationIdentity": "new-document-1",
            })
        reply = self.replies[operation]
        if isinstance(reply, BaseException):
            raise reply
        return reply

    def close(self):
        return True


class RecordingCoordinator:
    def __init__(self):
        self.guard = object()
        self.lease = object()
        self.calls = []

    def begin(self, coordination_identity, context):
        self.calls.append(("begin",))
        return self.guard

    def commit(self, guard, document, context):
        self.calls.append(("commit", guard, document))
        return self.lease

    def release(self, guard, document, lease, *, in_flight=False):
        self.calls.append((
            "release",
            guard,
            document,
            lease,
            in_flight,
        ))
        return DocumentResourceCleanup(state="released")


def request(action, params):
    return ActionRequest(
        address=ActionAddress(app="word", action=action),
        params=params,
    )


def trace(number):
    return TraceContext(trace_id=f"trace-{number}", trace_log=None)


def content_range(start=0, end=0):
    return {"start": start, "end": end, "revision": "revision-1"}


def point(value):
    return {"value": value, "unit": "pt"}


def open_document(backend, path="C:/docs/report.docx"):
    preparation = backend.prepare_open_document(path, context())
    return backend.open_document(preparation, path, context())


def create_document(backend):
    preparation = backend.prepare_create_document(context())
    return backend.create_document(preparation, context())


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


EMPTY_STRUCTURE = {
    "paragraphCount": 0,
    "headingCount": 0,
    "tableCount": 0,
    "inlineImageCount": 0,
    "floatingImageCount": 0,
    "sectionCount": 1,
    "pageBreakCount": 0,
    "sectionBreakCount": 0,
    "sections": ({
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
    },),
}


class WindowsWriterBackendTests(unittest.TestCase):
    def test_backend_and_bridge_satisfy_formal_seams(self):
        bridge = RecordingBridge()
        backend = WindowsWriterBackend(bridge=bridge)

        self.assertIsInstance(bridge, WindowsWriterBridge)
        self.assertIsInstance(backend, WriterBackend)

    def test_open_returns_exact_opaque_document_and_closed_observations(self):
        bridge = RecordingBridge({
            "acquire_existing_document": {
                "documentId": "document-1",
                "revision": "revision-1",
                "persistenceState": "saved",
                "readOnly": False,
                "artifactFormat": "docx",
                "artifactSizeBytes": 42,
            },
        })
        backend = WindowsWriterBackend(bridge=bridge)

        acquired = open_document(backend)

        self.assertIsInstance(acquired.document, WindowsWriterDocument)
        self.assertEqual("document-1", acquired.document.bridge_document_id)
        self.assertEqual("C:/docs/report.docx", acquired.document.authorized_path)
        self.assertEqual("revision-1", acquired.revision)
        self.assertEqual("saved", acquired.persistence_state)
        self.assertEqual("docx", acquired.artifact_format)
        self.assertEqual(42, acquired.artifact_size_bytes)
        self.assertEqual(
            (
                "acquire_existing_document",
                {"preparationId": "preparation-1"},
                context(),
            ),
            bridge.calls[1],
        )

    def test_create_returns_unsaved_document_without_locator(self):
        bridge = RecordingBridge({
            "acquire_new_document": {
                "documentId": "document-1",
                "revision": "revision-1",
                "persistenceState": "unsaved",
                "readOnly": False,
            },
        })
        backend = WindowsWriterBackend(bridge=bridge)

        acquired = create_document(backend)

        self.assertIsNone(acquired.document.authorized_path)
        self.assertEqual("unsaved", acquired.persistence_state)
        self.assertIsNone(acquired.artifact_format)
        self.assertIsNone(acquired.artifact_size_bytes)

    def test_liveness_and_operations_require_the_exact_bound_document(self):
        bridge = RecordingBridge({
            "acquire_existing_document": {
                "documentId": "document-1",
                "revision": "revision-1",
                "persistenceState": "saved",
                "readOnly": False,
                "artifactFormat": "docx",
                "artifactSizeBytes": 42,
            },
            "probe_bound_document": {"live": True},
            "read_revision_coherent_snapshot": {"revision": "revision-1"},
        })
        backend = WindowsWriterBackend(bridge=bridge)
        acquired = open_document(backend)

        self.assertTrue(backend.is_document_live(acquired.document))
        result = backend.invoke(
            acquired.document,
            WriterBackendOperation(
                "read_revision_coherent_snapshot",
                {"scope": {"kind": "document"}},
            ),
            context(),
        )

        self.assertEqual({"revision": "revision-1"}, result)
        self.assertEqual(
            {
                "documentId": "document-1",
                "authorizedPath": "C:/docs/report.docx",
                "operationArguments": {"scope": {"kind": "document"}},
            },
            bridge.calls[-1][1],
        )
        with self.assertRaisesRegex(ValueError, "another document"):
            backend.is_document_live(WindowsWriterDocument(
                bridge_document_id="document-1",
                authorized_path="C:/docs/report.docx",
            ))

    def test_definite_and_unprovable_acquisition_failures_remain_distinct(self):
        definite = WriterBackendActionFailure(
            outcome="failed",
            code="DOCUMENT_NOT_FOUND",
            message="missing",
            binding_disposition="unchanged",
        )
        backend = WindowsWriterBackend(bridge=RecordingBridge({
            "acquire_existing_document": definite,
        }))
        with self.assertRaises(WriterDefiniteEstablishFailure):
            open_document(backend, "C:/docs/missing.docx")

        unknown = WriterBackendActionFailure(
            outcome="unknown",
            code="RESPONSE_LOST",
            message="lost",
            binding_disposition="unprovable",
        )
        backend = WindowsWriterBackend(bridge=RecordingBridge({
            "acquire_new_document": unknown,
        }))
        with self.assertRaises(WriterUnprovableEstablishFailure):
            create_document(backend)

    def test_malformed_bridge_facts_fail_closed(self):
        backend = WindowsWriterBackend(bridge=RecordingBridge({
            "acquire_existing_document": {
                "documentId": "document-1",
                "revision": "revision-1",
                "persistenceState": "saved",
                "readOnly": False,
                "artifactFormat": "docx",
                "artifactSizeBytes": 42,
                "unexpected": True,
            },
        }))
        with self.assertRaisesRegex(TypeError, "invalid fields"):
            open_document(backend)

        bridge = RecordingBridge({
            "acquire_new_document": {
                "documentId": "document-1",
                "revision": "revision-1",
                "persistenceState": "unsaved",
                "readOnly": False,
            },
            "probe_bound_document": {"live": 1},
        })
        backend = WindowsWriterBackend(bridge=bridge)
        document = create_document(backend).document
        with self.assertRaisesRegex(TypeError, "must be Boolean"):
            backend.is_document_live(document)

    def test_backend_refuses_a_second_acquisition(self):
        bridge = RecordingBridge({
            "acquire_new_document": {
                "documentId": "document-1",
                "revision": "revision-1",
                "persistenceState": "unsaved",
                "readOnly": False,
            },
        })
        backend = WindowsWriterBackend(bridge=bridge)
        create_document(backend)

        with self.assertRaisesRegex(RuntimeError, "already bound"):
            create_document(backend)
        self.assertEqual(2, len(bridge.calls))

    def test_action_session_runs_open_write_inspect_and_save_on_one_document(self):
        bridge = RecordingBridge({
            "acquire_existing_document": {
                "documentId": "document-1",
                "revision": "revision-1",
                "persistenceState": "saved",
                "readOnly": False,
                "artifactFormat": "docx",
                "artifactSizeBytes": 42,
            },
            "probe_bound_document": {"live": True},
            "insert_structured_body_content": {
                "revisionBefore": "revision-1",
                "revisionAfter": "revision-2",
                "range": content_range(0, 6) | {"revision": "revision-2"},
            },
            "read_revision_coherent_snapshot": {
                "revision": "revision-2",
                "scopeRange": content_range(0, 6) | {"revision": "revision-2"},
                "returnedRange": content_range(0, 6) | {"revision": "revision-2"},
                "text": "Hello\n",
                "paragraphs": (),
                "structure": EMPTY_STRUCTURE,
                "documentState": {
                    "persistenceState": "saved",
                    "readOnly": False,
                },
                "truncated": False,
                "remainingRange": None,
            },
            "save_existing_artifact": {
                "revisionBefore": "revision-2",
                "revisionAfter": "revision-2",
                "artifact": {
                    "path": "C:/docs/report.docx",
                    "format": "docx",
                    "sizeBytes": 42,
                },
                "documentState": {
                    "persistenceState": "saved",
                    "readOnly": False,
                },
            },
        })
        backend = WindowsWriterBackend(bridge=bridge)
        coordinator = RecordingCoordinator()
        session = ActionSession(
            application="word",
            contracts=WORD_TARGET_CONTRACT_SET,
            adapter=WordAdapter(
                backend=backend,
                handlers=WORD_HANDLERS,
            ),
            coordinator=coordinator,
            session_id="session-1",
            request_id_factory=iter((
                "request-1",
                "request-2",
                "request-3",
                "request-4",
            )).__next__,
            clock=lambda: 0,
        )

        opened = session.execute(
            request("openDocument", {"path": "C:/docs/report.docx"}),
            trace(1),
        )
        written = session.execute(
            request("writeContent", {
                "anchor": {"kind": "documentEnd"},
                "blocks": ({
                    "kind": "paragraph",
                    "runs": ({"text": "Hello"},),
                },),
            }),
            trace(2),
        )
        inspected = session.execute(
            request("inspectDocument", {
                "scope": {"kind": "document"},
                "limits": {
                    "maxTextCharacters": 100,
                    "maxParagraphs": 10,
                    "maxRuns": 20,
                },
            }),
            trace(3),
        )
        saved = session.execute(request("save", {}), trace(4))
        cleanup = session.close()

        self.assertEqual("succeeded", opened.response.outcome)
        self.assertEqual("succeeded", written.response.outcome)
        self.assertEqual("succeeded", inspected.response.outcome)
        self.assertEqual("succeeded", saved.response.outcome)
        self.assertEqual("succeeded", cleanup.outcome)
        self.assertEqual(
            [
                "prepare_existing_document",
                "acquire_existing_document",
                "probe_bound_document",
                "insert_structured_body_content",
                "probe_bound_document",
                "read_revision_coherent_snapshot",
                "probe_bound_document",
                "save_existing_artifact",
            ],
            [call[0] for call in bridge.calls],
        )
        document = coordinator.calls[1][2]
        self.assertEqual("document-1", document.bridge_document_id)
        self.assertIs(document, coordinator.calls[2][2])


class WriterBackendActionFailureTests(unittest.TestCase):
    def test_failure_shape_is_closed(self):
        failure = WriterBackendActionFailure(
            outcome="unknown",
            code="OUTPUT_WRITE_FAILED",
            message="write result is unknown",
            binding_disposition="unprovable",
        )
        self.assertEqual("unknown", failure.outcome)
        self.assertEqual("unprovable", failure.binding_disposition)

        with self.assertRaisesRegex(ValueError, "cannot prove binding loss"):
            WriterBackendActionFailure(
                outcome="unknown",
                code="OUTPUT_WRITE_FAILED",
                message="unknown",
                binding_disposition="lost",
            )


if __name__ == "__main__":
    unittest.main()
