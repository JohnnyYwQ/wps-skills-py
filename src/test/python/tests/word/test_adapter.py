import unittest

from wps_skills.core.action_session import (
    ActionAddress,
    ActionRequest,
    ActionSession,
    ApplicationAdapter,
    ControllerCommand,
    ControllerContext,
    ControllerResult,
    DefiniteEstablishFailure,
    DocumentResourceCleanup,
    TraceContext,
    UnprovableEstablishFailure,
)
from wps_skills.word.contracts import WORD_TARGET_CONTRACT_SET
from wps_skills.word.adapter import (
    WordAdapter,
    WriterBackend,
    WriterBackendAcquisition,
    WriterBackendOperation,
    WriterBackendPreparation,
    WriterDefiniteEstablishFailure,
    WriterUnprovableEstablishFailure,
)


def command(action, params=None):
    return ControllerCommand(
        address=ActionAddress(app="word", action=action),
        params=params or {},
        context=ControllerContext(
            request_id="request-1",
            trace_id="trace-1",
            deadline_at=100,
        ),
    )


class RecordingWriterBackend:
    def __init__(self, document=None):
        self.document = document if document is not None else object()
        self.calls = []
        self.live = True
        self.create_failure = None
        self.open_failure = None

    def prepare_create_document(self, context):
        return WriterBackendPreparation(
            coordination_identity="new-document-1",
            state=self,
        )

    def prepare_open_document(self, path, context):
        return WriterBackendPreparation(
            coordination_identity="file-document-1",
            state=self,
        )

    def create_document(self, preparation, context):
        self.calls.append(("create_document", context))
        if self.create_failure is not None:
            raise self.create_failure
        return WriterBackendAcquisition(
            document=self.document,
            revision="revision-1",
            persistence_state="unsaved",
            read_only=False,
        )

    def open_document(self, preparation, path, context):
        self.calls.append(("open_document", path, context))
        if self.open_failure is not None:
            raise self.open_failure
        return WriterBackendAcquisition(
            document=self.document,
            revision="revision-1",
            persistence_state="saved",
            read_only=False,
            artifact_format="docx",
            artifact_size_bytes=10,
        )

    def is_document_live(self, document):
        self.calls.append(("is_document_live", document))
        return self.live

    def invoke(self, document, operation, context):
        self.calls.append(("invoke", document, operation, context))
        return {
            "operation": operation.name,
            "arguments": dict(operation.arguments),
        }

    def close(self):
        return True


class RecordingCoordinator:
    def __init__(self):
        self.guard = object()
        self.lease = object()

    def begin(self, coordination_identity, context):
        return self.guard

    def commit(self, guard, document, context):
        return self.lease

    def release(self, guard, document, lease, *, in_flight=False):
        return DocumentResourceCleanup(state="released")


class WordAdapterBoundaryTests(unittest.TestCase):
    def test_word_adapter_and_backend_satisfy_their_formal_seams(self):
        backend = RecordingWriterBackend()
        adapter = WordAdapter(backend=backend, handlers={})

        self.assertIsInstance(backend, WriterBackend)
        self.assertIsInstance(adapter, ApplicationAdapter)
        self.assertEqual("word", adapter.application)
        self.assertIs(WORD_TARGET_CONTRACT_SET, adapter.contracts)

    def test_establish_dispatches_create_and_open_without_forwarding_address(self):
        backend = RecordingWriterBackend()
        adapter = WordAdapter(backend=backend, handlers={})

        create_command = command("createDocument")
        created = adapter.establish(
            adapter.prepare_establish(create_command),
            create_command,
        )
        open_command = command(
            "openDocument",
            {"path": "C:/docs/report.docx"},
        )
        opened = adapter.establish(
            adapter.prepare_establish(open_command),
            open_command,
        )

        self.assertIs(backend.document, created.document)
        self.assertIs(backend.document, opened.document)
        self.assertEqual("unsaved", created.data["documentState"]["persistenceState"])
        self.assertEqual("C:/docs/report.docx", opened.data["artifact"]["path"])
        self.assertEqual(
            ["create_document", "open_document"],
            [call[0] for call in backend.calls],
        )
        self.assertEqual("C:/docs/report.docx", backend.calls[1][1])

    def test_required_handler_receives_exact_document_without_action_address(self):
        bound_document = object()
        other_document = object()
        backend = RecordingWriterBackend(document=other_document)
        observed = []

        def inspect_handler(backend, document, params, context):
            observed.append((document, params, context))
            operation = WriterBackendOperation(
                name="read_body_snapshot",
                arguments={"scope": params["scope"]},
            )
            data = backend.invoke(document, operation, context)
            return ControllerResult.succeeded(
                data=data,
                controller_state="usable",
                binding_disposition="unchanged",
            )

        adapter = WordAdapter(
            backend=backend,
            handlers={"inspectDocument": inspect_handler},
        )
        result = adapter.handle(
            bound_document,
            command("inspectDocument", {"scope": "body"}),
        )

        self.assertEqual("succeeded", result.outcome)
        self.assertIs(bound_document, observed[0][0])
        self.assertIs(bound_document, backend.calls[0][1])
        self.assertEqual("read_body_snapshot", backend.calls[0][2].name)
        self.assertNotEqual("inspectDocument", backend.calls[0][2].name)

    def test_missing_handler_reports_capability_unavailable_without_backend_call(self):
        backend = RecordingWriterBackend()
        adapter = WordAdapter(backend=backend, handlers={})

        result = adapter.handle(object(), command("save", {}))

        self.assertEqual("failed", result.outcome)
        self.assertEqual("WORD_CAPABILITY_UNAVAILABLE", result.error.code)
        self.assertEqual("usable", result.controller_state)
        self.assertEqual("unchanged", result.binding_disposition)
        self.assertEqual([], backend.calls)

    def test_liveness_requires_a_boolean_observation_for_the_exact_document(self):
        document = object()
        backend = RecordingWriterBackend()
        adapter = WordAdapter(backend=backend, handlers={})
        self.assertTrue(adapter.is_live(document))
        self.assertIs(document, backend.calls[0][1])

        backend.live = 1
        with self.assertRaisesRegex(TypeError, "must be Boolean"):
            adapter.is_live(document)

    def test_missing_handler_is_a_declared_nonterminal_action_failure_in_core(self):
        class ContractBackend(RecordingWriterBackend):
            def create_document(self, preparation, context):
                return WriterBackendAcquisition(
                    document=self.document,
                    revision="revision-1",
                    persistence_state="unsaved",
                    read_only=False,
                )

        session = ActionSession(
            application="word",
            contracts=WORD_TARGET_CONTRACT_SET,
            adapter=WordAdapter(
                backend=ContractBackend(),
                handlers={},
            ),
            coordinator=RecordingCoordinator(),
            session_id="session-1",
        )
        established = session.execute(
            ActionRequest(
                address=ActionAddress(
                    app="word",
                    action="createDocument",
                ),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        unavailable = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="save"),
                params={},
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )

        self.assertEqual("succeeded", established.response.outcome)
        self.assertEqual("failed", unavailable.response.outcome)
        self.assertEqual(
            "WORD_CAPABILITY_UNAVAILABLE",
            unavailable.response.error.code,
        )
        self.assertEqual("continue", unavailable.continuation)

    def test_backend_establish_failures_map_to_core_establish_failures(self):
        backend = RecordingWriterBackend()
        adapter = WordAdapter(backend=backend, handlers={})
        backend.open_failure = WriterDefiniteEstablishFailure(
            code="DOCUMENT_NOT_FOUND",
            message="missing document",
        )

        with self.assertRaisesRegex(
            DefiniteEstablishFailure,
            "missing document",
        ):
            open_command = command(
                "openDocument",
                {"path": "C:/docs/missing.docx"},
            )
            adapter.establish(
                adapter.prepare_establish(open_command),
                open_command,
            )

        partial = object()
        backend.create_failure = WriterUnprovableEstablishFailure(
            outcome="unknown",
            code="RESPONSE_LOST",
            message="creation response lost",
            partial_document=partial,
        )
        with self.assertRaises(UnprovableEstablishFailure) as raised:
            create_command = command("createDocument")
            adapter.establish(
                adapter.prepare_establish(create_command),
                create_command,
            )
        self.assertIs(partial, raised.exception.partial_document)

    def test_backend_operations_are_immutable_and_acquisition_is_closed(self):
        arguments = {"nested": {"values": [1]}}
        operation = WriterBackendOperation("read", arguments)
        acquisition = WriterBackendAcquisition(
            document=object(),
            revision="revision-1",
            persistence_state="unsaved",
            read_only=False,
        )
        arguments["nested"]["values"].append(2)

        self.assertEqual(
            (1,),
            operation.arguments["nested"]["values"],
        )
        self.assertEqual("revision-1", acquisition.revision)
        with self.assertRaises(TypeError):
            operation.arguments["new"] = True
        with self.assertRaises(AttributeError):
            acquisition.revision = "revision-2"
        with self.assertRaisesRegex(ValueError, "public Action name"):
            WriterBackendOperation("inspectDocument", {})

    def test_adapter_rejects_foreign_commands_and_establish_handlers(self):
        backend = RecordingWriterBackend()
        with self.assertRaisesRegex(ValueError, "invalid Word Action handler"):
            WordAdapter(
                backend=backend,
                handlers={"openDocument": lambda *args: None},
            )

        adapter = WordAdapter(backend=backend, handlers={})
        foreign = ControllerCommand(
            address=ActionAddress(app="excel", action="inspectDocument"),
            params={},
            context=command("inspectDocument").context,
        )
        with self.assertRaisesRegex(ValueError, "another application"):
            adapter.handle(object(), foreign)

    def test_missing_backend_method_fails_seam_construction(self):
        class IncompleteBackend:
            def create_document(self, context):
                return None

        with self.assertRaisesRegex(ValueError, "Writer Backend seam"):
            WordAdapter(backend=IncompleteBackend(), handlers={})


if __name__ == "__main__":
    unittest.main()
