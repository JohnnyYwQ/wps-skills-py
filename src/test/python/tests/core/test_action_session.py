import unittest
import threading
import time

from wps_skills.core.action_session import (
    AcquiredDocument,
    ActionAddress,
    ActionContract,
    ActionError,
    ActionRequest,
    ActionResponse,
    ActionSession,
    ActionTurn,
    ApplicationAdapter,
    ApplicationContractSet,
    ControllerResult,
    ControllerCommand,
    DocumentResourceCleanup,
    DefiniteEstablishFailure,
    PreparedDocumentAcquisition,
    ProcessCleanup,
    RequiredBindingFailure,
    TraceContext,
    UnprovableEstablishFailure,
)


class FakeWriterAdapter:
    application = "word"

    def __init__(
        self,
        *,
        document=None,
        events=None,
        establish_script=None,
        handler_script=None,
    ):
        self.calls = []
        self.document = document
        self.events = events if events is not None else []
        self.active_document = document
        self.live = True
        self.establish_script = list(establish_script or [])
        self.handler_script = list(handler_script or [])

    def prepare_establish(self, command):
        self.events.append("adapter.prepare_establish")
        return PreparedDocumentAcquisition(
            coordination_identity="fake-document-identity",
            application_state=self,
        )

    def establish(self, prepared, command):
        if prepared.application_state is not self:
            raise ValueError("another acquisition plan")
        self.calls.append(("establish", command))
        self.events.append("adapter.establish")
        if self.establish_script:
            result = self.establish_script.pop(0)
            if isinstance(result, BaseException):
                raise result
            return result
        return AcquiredDocument(
            document=self.document,
            data={"created": True},
        )

    def is_live(self, document):
        self.calls.append(("is_live", document))
        return self.live

    def handle(self, document, command):
        self.calls.append(("handle", document, command))
        if self.handler_script:
            result = self.handler_script.pop(0)
            if isinstance(result, BaseException):
                raise result
            return result
        return ControllerResult.succeeded(
            data={"text": "bound document"},
            controller_state="usable",
            binding_disposition="unchanged",
        )

    def handle_none(self, command):
        self.calls.append(("handle_none", command))
        return ControllerResult.succeeded(
            data={"status": "ready"},
            controller_state="usable",
            binding_disposition="unchanged",
        )

    def close(self):
        return True


class FakeDocumentCoordinator:
    def __init__(
        self,
        *,
        events=None,
        release_state="released",
        release_error=None,
        commit_error=None,
        begin_error=None,
        release_wait=None,
        begin_wait=None,
        begin_entered=None,
        release_entered=None,
    ):
        self.calls = []
        self.events = events if events is not None else []
        self.guard = object()
        self.lease = object()
        self.release_state = release_state
        self.release_error = release_error
        self.commit_error = commit_error
        self.begin_error = begin_error
        self.release_wait = release_wait
        self.begin_wait = begin_wait
        self.begin_entered = begin_entered
        self.release_entered = release_entered
        self.release_in_flight = []

    def begin(self, coordination_identity, context):
        self.calls.append(("begin",))
        self.events.append("coordinator.begin")
        if self.begin_entered is not None:
            self.begin_entered.set()
        if self.begin_wait is not None:
            self.begin_wait.wait(timeout=2)
        if self.begin_error is not None:
            raise self.begin_error
        return self.guard

    def commit(self, guard, document, context):
        self.calls.append(("commit", guard, document))
        self.events.append("coordinator.commit")
        if self.commit_error is not None:
            raise self.commit_error
        return self.lease

    def release(self, guard, document, lease, *, in_flight=False):
        self.calls.append(("release", guard, document, lease))
        self.events.append("coordinator.release")
        self.release_in_flight.append(in_flight)
        if self.release_entered is not None:
            self.release_entered.set()
        if self.release_wait is not None:
            self.release_wait.wait(timeout=2)
        if self.release_error is not None:
            raise self.release_error
        return DocumentResourceCleanup(state=self.release_state)


def word_contracts(*contracts):
    return ApplicationContractSet(
        application="word",
        contracts=contracts,
        allow_incomplete=True,
    )


class ActionSessionBindingTests(unittest.TestCase):
    def test_action_session_owns_and_reports_launcher_cleanup_once(self):
        class Launcher:
            def __init__(self):
                self.close_count = 0

            def close(self):
                self.close_count += 1
                return (ProcessCleanup(
                    pid=42,
                    cleanup_steps=("already_exited",),
                    released=True,
                ),)

        launcher = Launcher()
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=FakeWriterAdapter(),
            coordinator=FakeDocumentCoordinator(),
            session_id="session-1",
            launcher=launcher,
        )

        first = session.close()
        second = session.close()

        self.assertIs(first, second)
        self.assertEqual(1, launcher.close_count)
        self.assertEqual(42, first.cleanup.processes[0].pid)

    def test_action_response_rejects_mixed_or_incomplete_variants(self):
        address = ActionAddress(app="word", action="inspectDocument")
        trace = TraceContext(trace_id="trace-1", trace_log=None)
        with self.assertRaises(ValueError):
            ActionResponse(
                outcome="succeeded",
                address=address,
                session_id="session-1",
                trace=trace,
                data={},
                error=ActionError(code="IMPOSSIBLE", message="mixed variant"),
            )
        with self.assertRaises(ValueError):
            ActionResponse(
                outcome="failed",
                address=address,
                session_id="session-1",
                trace=trace,
            )

    def test_action_response_captures_immutable_data_and_emits_json_values(self):
        data = {"nested": {"ids": [1]}}
        response = ActionResponse(
            outcome="succeeded",
            address=ActionAddress(app="word", action="inspectDocument"),
            session_id="session-1",
            trace=TraceContext(trace_id="trace-1", trace_log=None),
            data=data,
        )
        data["nested"]["ids"].append(2)

        self.assertEqual([1], list(response.data["nested"]["ids"]))
        self.assertEqual(
            {"nested": {"ids": [1]}},
            response.to_wire()["data"],
        )

    def test_closed_core_types_reject_unknown_tags(self):
        with self.assertRaises(ValueError):
            ActionAddress(app="pages", action="inspectDocument")
        with self.assertRaises(ValueError):
            ControllerResult(
                outcome="maybe",
                controller_state="usable",
                binding_disposition="unchanged",
                error=ActionError(code="BAD", message="bad tag"),
            )
        with self.assertRaises(ValueError):
            ControllerResult(
                outcome="failed",
                controller_state="usable",
                binding_disposition="unchanged",
                error="bad",
            )
        response = ActionResponse(
            outcome="failed",
            address=ActionAddress(app="word", action="inspectDocument"),
            session_id="session-1",
            trace=TraceContext(trace_id="trace-1", trace_log=None),
            error=ActionError(code="FAILED", message="failed"),
        )
        with self.assertRaises(ValueError):
            ActionTurn(response=response, continuation="retry")
        with self.assertRaises(ValueError):
            DocumentResourceCleanup(state="maybe_released")

    def test_contract_set_rejects_duplicate_action_names(self):
        contract = ActionContract(
            name="inspectDocument",
            binding_role="required",
            risk="read",
            parameters={"type": "object"},
            result={"type": "object"},
        )
        with self.assertRaisesRegex(ValueError, "duplicate Action"):
            word_contracts(contract, contract)

    def test_contract_set_rejects_invalid_binding_role_and_risk(self):
        invalid_contracts = (
            ActionContract(
                name="optionalDocument",
                binding_role="optional",
                risk="read",
                parameters={"type": "object"},
                result={"type": "object"},
            ),
            ActionContract(
                name="unsafeRisk",
                binding_role="required",
                risk="execute",
                parameters={"type": "object"},
                result={"type": "object"},
            ),
        )
        for contract in invalid_contracts:
            with self.subTest(contract=contract.name):
                with self.assertRaises(ValueError):
                    word_contracts(contract)

    def test_session_construction_rejects_mismatched_contracts_or_adapter(self):
        contract = ActionContract(
            name="inspectDocument",
            binding_role="required",
            risk="read",
            parameters={"type": "object"},
            result={"type": "object"},
        )
        with self.assertRaisesRegex(ValueError, "Contract Set application"):
            ActionSession(
                application="word",
                contracts=ApplicationContractSet(
                    application="excel",
                    contracts=(contract,),
                    allow_incomplete=True,
                ),
                adapter=FakeWriterAdapter(),
                coordinator=FakeDocumentCoordinator(),
                session_id="session-1",
            )
        adapter = FakeWriterAdapter()
        adapter.application = "excel"
        with self.assertRaisesRegex(ValueError, "Adapter application"):
            ActionSession(
                application="word",
                contracts=word_contracts(contract),
                adapter=adapter,
                coordinator=FakeDocumentCoordinator(),
                session_id="session-1",
            )

        class IncompleteAdapter:
            application = "word"

        self.assertNotIsInstance(IncompleteAdapter(), ApplicationAdapter)
        with self.assertRaisesRegex(ValueError, "Application Adapter seam"):
            ActionSession(
                application="word",
                contracts=word_contracts(contract),
                adapter=IncompleteAdapter(),
                coordinator=FakeDocumentCoordinator(),
                session_id="session-1",
            )

    def test_fresh_session_rejects_required_action_without_dispatch(self):
        adapter = FakeWriterAdapter()
        coordinator = FakeDocumentCoordinator()
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="inspectDocument",
                binding_role="required",
                risk="read",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,
            session_id="session-1",
        )

        turn = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        self.assertEqual("failed", turn.response.outcome)
        self.assertEqual(
            "SESSION_DOCUMENT_NOT_BOUND",
            turn.response.error.code,
        )
        self.assertEqual("continue", turn.continuation)
        self.assertEqual([], adapter.calls)
        self.assertEqual([], coordinator.calls)

    def test_session_rejects_different_application_without_dispatch(self):
        adapter = FakeWriterAdapter()
        coordinator = FakeDocumentCoordinator()
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="inspectDocument",
                binding_role="required",
                risk="read",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,
            session_id="session-1",
        )

        turn = session.execute(
            ActionRequest(
                address=ActionAddress(app="excel", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        self.assertEqual("failed", turn.response.outcome)
        self.assertEqual("SESSION_APP_MISMATCH", turn.response.error.code)
        self.assertEqual("continue", turn.continuation)
        self.assertEqual([], adapter.calls)
        self.assertEqual([], coordinator.calls)

    def test_unknown_local_action_is_a_nonterminal_action_failure(self):
        adapter = FakeWriterAdapter()
        coordinator = FakeDocumentCoordinator()
        session = ActionSession(
            application="word",
            contracts=word_contracts(),
            adapter=adapter,
            coordinator=coordinator,
            session_id="session-1",
        )

        turn = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="notAnAction"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        self.assertEqual("failed", turn.response.outcome)
        self.assertEqual("UNKNOWN_ACTION", turn.response.error.code)
        self.assertEqual("continue", turn.continuation)
        self.assertEqual([], adapter.calls)
        self.assertEqual([], coordinator.calls)

    def test_none_action_executes_without_document_resources_and_preserves_unbound_state(self):
        adapter = FakeWriterAdapter()
        coordinator = FakeDocumentCoordinator()
        session = ActionSession(
            application="word",
            contracts=word_contracts(
                ActionContract(
                    name="status",
                    binding_role="none",
                    risk="read",
                    parameters={"type": "object"},
                    result={
                        "type": "object",
                        "properties": {"status": {"type": "string"}},
                        "required": ["status"],
                        "additionalProperties": False,
                    },
                ),
                ActionContract(
                    name="inspectDocument",
                    binding_role="required",
                    risk="read",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
            ),
            adapter=adapter,
            coordinator=coordinator,
            session_id="session-1",
        )

        status = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="status"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        required = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )

        self.assertEqual("succeeded", status.response.outcome)
        self.assertEqual({"status": "ready"}, status.response.data)
        self.assertEqual("continue", status.continuation)
        self.assertEqual("SESSION_DOCUMENT_NOT_BOUND", required.response.error.code)
        self.assertEqual([], coordinator.calls)
        self.assertEqual(
            1,
            len([call for call in adapter.calls if call[0] == "handle_none"]),
        )

    def test_successful_establish_commits_document_and_lease_before_response(self):
        events = []
        document = object()
        adapter = FakeWriterAdapter(document=document, events=events)
        coordinator = FakeDocumentCoordinator(events=events)
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,
            session_id="session-1",
        )

        turn = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        self.assertEqual("succeeded", turn.response.outcome)
        self.assertEqual({"created": True}, turn.response.data)
        self.assertEqual("continue", turn.continuation)
        self.assertEqual(
            [
                "adapter.prepare_establish",
                "coordinator.begin",
                "adapter.establish",
                "coordinator.commit",
            ],
            events,
        )
        self.assertEqual(
            [("commit", coordinator.guard, document)],
            [call for call in coordinator.calls if call[0] == "commit"],
        )

    def test_establishing_state_is_recorded_before_adapter_acquisition(self):
        events = []
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=FakeWriterAdapter(document=object(), events=events),
            coordinator=FakeDocumentCoordinator(events=events),
            session_id="session-1",
        )

        session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(
                trace_id="trace-1",
                trace_log=None,
                event_sink=events.append,
            ),
        )

        self.assertLess(
            events.index("binding.establishing"),
            events.index("adapter.prepare_establish"),
        )

    def test_required_request_rejects_all_document_routing_decoys(self):
        document = object()
        adapter = FakeWriterAdapter(document=document)
        session = ActionSession(
            application="word",
            contracts=word_contracts(
                ActionContract(
                    name="createDocument",
                    binding_role="establish",
                    risk="write",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
                ActionContract(
                    name="inspectDocument",
                    binding_role="required",
                    risk="read",
                    parameters={
                        "type": "object",
                        "properties": {},
                        "additionalProperties": False,
                    },
                    result={"type": "object"},
                ),
            ),
            adapter=adapter,
            coordinator=FakeDocumentCoordinator(),
            session_id="session-1",
        )
        session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        decoys = {
            "path": "C:/wrong.docx",
            "displayName": "wrong.docx",
            "activeDocument": True,
            "candidate": {"id": "candidate-1"},
            "registryKey": "registry-1",
            "documentId": "document-1",
        }
        for index, (name, value) in enumerate(decoys.items(), start=2):
            with self.subTest(decoy=name):
                turn = session.execute(
                    ActionRequest(
                        address=ActionAddress(
                            app="word",
                            action="inspectDocument",
                        ),
                        params={name: value},
                    ),
                    TraceContext(
                        trace_id=f"trace-{index}",
                        trace_log=None,
                    ),
                )
                self.assertEqual("failed", turn.response.outcome)
                self.assertEqual("INVALID_PARAMS", turn.response.error.code)
                self.assertEqual("continue", turn.continuation)
        self.assertEqual(
            [],
            [call for call in adapter.calls if call[0] == "handle"],
        )
        self.assertEqual(
            [],
            [call for call in adapter.calls if call[0] == "is_live"],
        )

    def test_required_action_rejects_illegal_controller_binding_matrix(self):
        document = object()
        adapter = FakeWriterAdapter(
            document=document,
            handler_script=[ControllerResult.succeeded(
                data={"text": "wrong transition"},
                controller_state="usable",
                binding_disposition="established",
            )],
        )
        session = ActionSession(
            application="word",
            contracts=word_contracts(
                ActionContract(
                    name="createDocument",
                    binding_role="establish",
                    risk="write",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
                ActionContract(
                    name="inspectDocument",
                    binding_role="required",
                    risk="read",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
            ),
            adapter=adapter,
            coordinator=FakeDocumentCoordinator(),
            session_id="session-1",
        )
        session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        turn = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )

        self.assertEqual("failed", turn.response.outcome)
        self.assertEqual("INVALID_RESULT", turn.response.error.code)
        self.assertEqual("terminate", turn.continuation)

    def test_read_only_invalid_success_data_maps_to_failed_without_losing_binding(self):
        document = object()
        adapter = FakeWriterAdapter(
            document=document,
            handler_script=[
                ControllerResult.succeeded(
                    data={},
                    controller_state="usable",
                    binding_disposition="unchanged",
                ),
                ControllerResult.succeeded(
                    data={"text": "still bound"},
                    controller_state="usable",
                    binding_disposition="unchanged",
                ),
            ],
        )
        session = ActionSession(
            application="word",
            contracts=word_contracts(
                ActionContract(
                    name="createDocument",
                    binding_role="establish",
                    risk="write",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
                ActionContract(
                    name="inspectDocument",
                    binding_role="required",
                    risk="read",
                    parameters={"type": "object"},
                    result={
                        "type": "object",
                        "properties": {"text": {"type": "string"}},
                        "required": ["text"],
                        "additionalProperties": False,
                    },
                ),
            ),
            adapter=adapter,
            coordinator=FakeDocumentCoordinator(),
            session_id="session-1",
        )
        session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        request = ActionRequest(
            address=ActionAddress(app="word", action="inspectDocument"),
            params={},
        )

        invalid = session.execute(
            request,
            TraceContext(trace_id="trace-2", trace_log=None),
        )
        valid = session.execute(
            request,
            TraceContext(trace_id="trace-3", trace_log=None),
        )

        self.assertEqual("failed", invalid.response.outcome)
        self.assertEqual("INVALID_RESULT", invalid.response.error.code)
        self.assertEqual("continue", invalid.continuation)
        self.assertEqual("succeeded", valid.response.outcome)
        self.assertEqual({"text": "still bound"}, valid.response.data)

    def test_invalid_establish_data_is_unknown_terminal_after_binding_commit(self):
        document = object()
        adapter = FakeWriterAdapter(
            document=document,
            establish_script=[AcquiredDocument(document=document, data={})],
        )
        coordinator = FakeDocumentCoordinator()
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={
                    "type": "object",
                    "properties": {"created": {"type": "boolean"}},
                    "required": ["created"],
                    "additionalProperties": False,
                },
            )),
            adapter=adapter,
            coordinator=coordinator,
            session_id="session-1",
        )

        turn = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        cleanup = session.close()

        self.assertEqual("unknown", turn.response.outcome)
        self.assertEqual("INVALID_RESULT", turn.response.error.code)
        self.assertEqual("terminate", turn.continuation)
        self.assertEqual("succeeded", cleanup.outcome)
        self.assertEqual(
            [("release", None, document, coordinator.lease)],
            [call for call in coordinator.calls if call[0] == "release"],
        )

    def test_required_handler_receives_bound_document_after_active_document_changes(self):
        bound_document = object()
        other_document = object()
        adapter = FakeWriterAdapter(document=bound_document)
        coordinator = FakeDocumentCoordinator()
        session = ActionSession(
            application="word",
            contracts=word_contracts(
                ActionContract(
                    name="createDocument",
                    binding_role="establish",
                    risk="write",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
                ActionContract(
                    name="inspectDocument",
                    binding_role="required",
                    risk="read",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
            ),
            adapter=adapter,
            coordinator=coordinator,
            session_id="session-1",
        )
        session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        adapter.active_document = other_document
        turn = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )

        self.assertEqual("succeeded", turn.response.outcome)
        self.assertEqual({"text": "bound document"}, turn.response.data)
        self.assertEqual("continue", turn.continuation)
        self.assertEqual(
            [bound_document],
            [call[1] for call in adapter.calls if call[0] == "handle"],
        )

    def test_adapter_receives_one_immutable_controller_command_per_action(self):
        document = object()
        adapter = FakeWriterAdapter(document=document)
        request_ids = iter(["request-1", "request-2"])
        session = ActionSession(
            application="word",
            contracts=word_contracts(
                ActionContract(
                    name="createDocument",
                    binding_role="establish",
                    risk="write",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
                ActionContract(
                    name="inspectDocument",
                    binding_role="required",
                    risk="read",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
            ),
            adapter=adapter,
            coordinator=FakeDocumentCoordinator(),
            session_id="session-1",
            request_id_factory=lambda: next(request_ids),
            clock=lambda: 100.0,
            action_timeout_seconds=120,
        )
        session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        required_params = {"nested": {"ids": [1]}}
        session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="inspectDocument"),
                params=required_params,
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )
        required_params["nested"]["ids"].append(2)

        establish_command = next(
            call[1] for call in adapter.calls if call[0] == "establish"
        )
        required_document, required_command = next(
            (call[1], call[2])
            for call in adapter.calls
            if call[0] == "handle"
        )
        self.assertIsInstance(establish_command, ControllerCommand)
        self.assertEqual(
            ActionAddress(app="word", action="createDocument"),
            establish_command.address,
        )
        self.assertEqual("request-1", establish_command.context.request_id)
        self.assertEqual("trace-1", establish_command.context.trace_id)
        self.assertEqual(220.0, establish_command.context.deadline_at)
        self.assertIs(document, required_document)
        self.assertEqual("request-2", required_command.context.request_id)
        self.assertEqual("trace-2", required_command.context.trace_id)
        self.assertEqual(220.0, required_command.context.deadline_at)
        self.assertEqual([1], list(required_command.params["nested"]["ids"]))
        with self.assertRaises(TypeError):
            required_command.params["nested"]["new"] = True

    def test_bound_session_rejects_second_establish_without_dispatch(self):
        adapter = FakeWriterAdapter(document=object())
        coordinator = FakeDocumentCoordinator()
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,
            session_id="session-1",
        )
        request = ActionRequest(
            address=ActionAddress(app="word", action="createDocument"),
            params={},
        )
        session.execute(
            request,
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        turn = session.execute(
            request,
            TraceContext(trace_id="trace-2", trace_log=None),
        )

        self.assertEqual("failed", turn.response.outcome)
        self.assertEqual(
            "SESSION_DOCUMENT_ALREADY_BOUND",
            turn.response.error.code,
        )
        self.assertEqual("continue", turn.continuation)
        establish_calls = [
            call for call in adapter.calls if call[0] == "establish"
        ]
        self.assertEqual(1, len(establish_calls))
        self.assertEqual(
            "createDocument",
            establish_calls[0][1].address.action,
        )
        self.assertEqual(
            1,
            len([call for call in coordinator.calls if call[0] == "begin"]),
        )

    def test_definite_establish_failure_releases_resources_and_allows_retry(self):
        document = object()
        adapter = FakeWriterAdapter(
            document=document,
            establish_script=[
                DefiniteEstablishFailure(
                    code="DOCUMENT_NOT_FOUND",
                    message="The requested document was not found",
                ),
                AcquiredDocument(
                    document=document,
                    data={"created": True},
                ),
            ],
        )
        coordinator = FakeDocumentCoordinator()
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,
            session_id="session-1",
        )
        request = ActionRequest(
            address=ActionAddress(app="word", action="createDocument"),
            params={},
        )

        failed = session.execute(
            request,
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        retried = session.execute(
            request,
            TraceContext(trace_id="trace-2", trace_log=None),
        )

        self.assertEqual("failed", failed.response.outcome)
        self.assertEqual("DOCUMENT_NOT_FOUND", failed.response.error.code)
        self.assertEqual("continue", failed.continuation)
        self.assertEqual("succeeded", retried.response.outcome)
        self.assertEqual("continue", retried.continuation)
        self.assertEqual(
            [("release", coordinator.guard, None, None)],
            [call for call in coordinator.calls if call[0] == "release"],
        )
        self.assertEqual(
            2,
            len([call for call in coordinator.calls if call[0] == "begin"]),
        )

    def test_definite_failure_with_quarantined_guard_is_terminal(self):
        coordinator = FakeDocumentCoordinator(release_state="quarantined")
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=FakeWriterAdapter(establish_script=[
                DefiniteEstablishFailure(
                    code="DOCUMENT_NOT_FOUND",
                    message="The requested document was not found",
                ),
            ]),
            coordinator=coordinator,
            session_id="session-1",
        )

        turn = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        cleanup = session.close()

        self.assertEqual("failed", turn.response.outcome)
        self.assertEqual(
            "DOCUMENT_BINDING_UNAVAILABLE",
            turn.response.error.code,
        )
        self.assertEqual("terminate", turn.continuation)
        self.assertEqual("failed", cleanup.outcome)
        self.assertEqual("quarantined", cleanup.cleanup.document_resources.state)
        self.assertEqual(
            1,
            len([call for call in coordinator.calls if call[0] == "release"]),
        )

    def test_close_does_not_repeat_guard_release_already_in_progress(self):
        release_entered = threading.Event()
        resume_release = threading.Event()
        coordinator = FakeDocumentCoordinator(
            release_wait=resume_release,
            release_entered=release_entered,
        )
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=FakeWriterAdapter(establish_script=[
                DefiniteEstablishFailure(
                    code="DOCUMENT_NOT_FOUND",
                    message="The requested document was not found",
                ),
            ]),
            coordinator=coordinator,
            session_id="session-1",
        )
        turns = []
        worker = threading.Thread(target=lambda: turns.append(session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )))
        worker.start()
        self.assertTrue(release_entered.wait(timeout=1))

        cleanup = session.close()
        resume_release.set()
        worker.join(timeout=2)

        self.assertFalse(worker.is_alive())
        self.assertEqual("failed", cleanup.outcome)
        self.assertEqual(
            "release_unconfirmed",
            cleanup.cleanup.document_resources.state,
        )
        self.assertEqual("terminate", turns[0].continuation)
        self.assertEqual(
            1,
            len([call for call in coordinator.calls if call[0] == "release"]),
        )

    def test_definite_guard_acquisition_failure_remains_unbound_without_release(self):
        coordinator = FakeDocumentCoordinator(
            begin_error=DefiniteEstablishFailure(
                code="DOCUMENT_LEASE_CONFLICT",
                message="Another Session owns this document",
            ),
        )
        adapter = FakeWriterAdapter(document=object())
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,
            session_id="session-1",
        )

        turn = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        cleanup = session.close()

        self.assertEqual("failed", turn.response.outcome)
        self.assertEqual("DOCUMENT_LEASE_CONFLICT", turn.response.error.code)
        self.assertEqual("continue", turn.continuation)
        self.assertEqual("not_acquired", cleanup.cleanup.document_resources.state)
        self.assertEqual([], adapter.calls)
        self.assertEqual(
            [],
            [call for call in coordinator.calls if call[0] == "release"],
        )

    def test_unknown_unprovable_establish_is_terminal_and_cleanup_releases_partial_guard(self):
        partial_document = object()
        adapter = FakeWriterAdapter(
            establish_script=[UnprovableEstablishFailure(
                outcome="unknown",
                code="RESPONSE_LOST",
                message="No reliable establish response was received",
                partial_document=partial_document,
            )],
        )
        coordinator = FakeDocumentCoordinator()
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,
            session_id="session-1",
        )

        turn = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        first_cleanup = session.close()
        second_cleanup = session.close()

        self.assertEqual("unknown", turn.response.outcome)
        self.assertEqual("RESPONSE_LOST", turn.response.error.code)
        self.assertEqual("terminate", turn.continuation)
        self.assertEqual("succeeded", first_cleanup.outcome)
        self.assertEqual("released", first_cleanup.cleanup.document_resources.state)
        self.assertIs(first_cleanup, second_cleanup)
        self.assertEqual(
            [("release", coordinator.guard, partial_document, None)],
            [call for call in coordinator.calls if call[0] == "release"],
        )

    def test_failed_unprovable_establish_normalizes_binding_error_and_terminates(self):
        partial_document = object()
        adapter = FakeWriterAdapter(
            establish_script=[UnprovableEstablishFailure(
                outcome="failed",
                code="OPEN_FAILED",
                message="The document may have opened without a provable binding",
                partial_document=partial_document,
            )],
        )
        coordinator = FakeDocumentCoordinator()
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,
            session_id="session-1",
        )

        turn = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        cleanup = session.close()

        self.assertEqual("failed", turn.response.outcome)
        self.assertEqual(
            "DOCUMENT_BINDING_UNAVAILABLE",
            turn.response.error.code,
        )
        self.assertEqual("terminate", turn.continuation)
        self.assertEqual("succeeded", cleanup.outcome)
        self.assertEqual(
            [("release", coordinator.guard, partial_document, None)],
            [call for call in coordinator.calls if call[0] == "release"],
        )

    def test_lease_commit_failure_is_terminal_and_cleanup_reclaims_partial_document(self):
        document = object()
        coordinator = FakeDocumentCoordinator(
            commit_error=RuntimeError("lease commit failed"),
        )
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=FakeWriterAdapter(document=document),
            coordinator=coordinator,
            session_id="session-1",
        )

        turn = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        cleanup = session.close()

        self.assertEqual("failed", turn.response.outcome)
        self.assertEqual(
            "DOCUMENT_BINDING_UNAVAILABLE",
            turn.response.error.code,
        )
        self.assertEqual("terminate", turn.continuation)
        self.assertEqual("succeeded", cleanup.outcome)
        self.assertEqual(
            [("release", coordinator.guard, document, None)],
            [call for call in coordinator.calls if call[0] == "release"],
        )

    def test_required_binding_loss_preserves_failure_and_terminates(self):
        document = object()
        adapter = FakeWriterAdapter(
            document=document,
            handler_script=[RequiredBindingFailure(
                disposition="lost",
                outcome="failed",
                code="DOCUMENT_CLOSED",
                message="The bound document is closed",
            )],
        )
        coordinator = FakeDocumentCoordinator()
        session = ActionSession(
            application="word",
            contracts=word_contracts(
                ActionContract(
                    name="createDocument",
                    binding_role="establish",
                    risk="write",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
                ActionContract(
                    name="inspectDocument",
                    binding_role="required",
                    risk="read",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
            ),
            adapter=adapter,
            coordinator=coordinator,
            session_id="session-1",
        )
        session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        turn = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )
        cleanup = session.close()

        self.assertEqual("failed", turn.response.outcome)
        self.assertEqual("DOCUMENT_CLOSED", turn.response.error.code)
        self.assertEqual("terminate", turn.continuation)
        self.assertEqual("succeeded", cleanup.outcome)
        self.assertEqual(
            [("release", None, document, coordinator.lease)],
            [call for call in coordinator.calls if call[0] == "release"],
        )

    def test_read_only_required_unprovable_unknown_maps_to_failed_and_terminates(self):
        document = object()
        adapter = FakeWriterAdapter(
            document=document,
            handler_script=[RequiredBindingFailure(
                disposition="unprovable",
                outcome="unknown",
                code="RESPONSE_LOST",
                message="The read response could not be recovered",
            )],
        )
        session = ActionSession(
            application="word",
            contracts=word_contracts(
                ActionContract(
                    name="createDocument",
                    binding_role="establish",
                    risk="write",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
                ActionContract(
                    name="inspectDocument",
                    binding_role="required",
                    risk="read",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
            ),
            adapter=adapter,
            coordinator=FakeDocumentCoordinator(),
            session_id="session-1",
        )
        session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        turn = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )

        self.assertEqual("failed", turn.response.outcome)
        self.assertEqual("RESPONSE_LOST", turn.response.error.code)
        self.assertEqual("terminate", turn.continuation)

    def test_closed_bound_document_prevents_required_handler_and_terminates(self):
        document = object()
        adapter = FakeWriterAdapter(document=document)
        session = ActionSession(
            application="word",
            contracts=word_contracts(
                ActionContract(
                    name="createDocument",
                    binding_role="establish",
                    risk="write",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
                ActionContract(
                    name="inspectDocument",
                    binding_role="required",
                    risk="read",
                    parameters={"type": "object"},
                    result={"type": "object"},
                ),
            ),
            adapter=adapter,
            coordinator=FakeDocumentCoordinator(),
            session_id="session-1",
        )
        session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )
        adapter.live = False

        turn = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )

        self.assertEqual("failed", turn.response.outcome)
        self.assertEqual("DOCUMENT_CLOSED", turn.response.error.code)
        self.assertEqual("terminate", turn.continuation)
        self.assertEqual(
            [],
            [call for call in adapter.calls if call[0] == "handle"],
        )
        self.assertEqual(
            [("is_live", document)],
            [call for call in adapter.calls if call[0] == "is_live"],
        )

    def test_bound_cleanup_reports_quarantine_failure_and_releases_once(self):
        document = object()
        adapter = FakeWriterAdapter(document=document)
        coordinator = FakeDocumentCoordinator(release_state="quarantined")
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,
            session_id="session-1",
        )
        session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        first = session.close()
        second = session.close()

        self.assertEqual("failed", first.outcome)
        self.assertEqual("SESSION_CLEANUP_INCOMPLETE", first.error.code)
        self.assertEqual("quarantined", first.cleanup.document_resources.state)
        self.assertIs(first, second)
        self.assertEqual(
            [("release", None, document, coordinator.lease)],
            [call for call in coordinator.calls if call[0] == "release"],
        )

    def test_cleanup_release_exception_is_reported_and_not_retried(self):
        document = object()
        adapter = FakeWriterAdapter(document=document)
        coordinator = FakeDocumentCoordinator(
            release_error=RuntimeError("coordinator unavailable"),
        )
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,
            session_id="session-1",
        )
        session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        first = session.close()
        second = session.close()

        self.assertEqual("failed", first.outcome)
        self.assertEqual(
            "release_unconfirmed",
            first.cleanup.document_resources.state,
        )
        self.assertEqual("SESSION_CLEANUP_INCOMPLETE", first.error.code)
        self.assertIs(first, second)
        self.assertEqual(
            1,
            len([call for call in coordinator.calls if call[0] == "release"]),
        )

    def test_close_during_establish_is_bounded_and_quarantines_in_flight_work(self):
        entered = threading.Event()
        resume = threading.Event()
        document = object()

        class BlockingAdapter(FakeWriterAdapter):
            def establish(self, prepared, command):
                self.calls.append(("establish", command))
                entered.set()
                resume.wait(timeout=2)
                return AcquiredDocument(
                    document=document,
                    data={"created": True},
                )

        coordinator = FakeDocumentCoordinator(release_state="quarantined")
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=BlockingAdapter(document=document),
            coordinator=coordinator,
            session_id="session-1",
        )
        turns = []
        worker = threading.Thread(target=lambda: turns.append(session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )))
        worker.start()
        self.assertTrue(entered.wait(timeout=1))

        started = time.monotonic()
        cleanup = session.close()
        elapsed = time.monotonic() - started
        resume.set()
        worker.join(timeout=2)

        self.assertLess(elapsed, 0.5)
        self.assertFalse(worker.is_alive())
        self.assertEqual("failed", cleanup.outcome)
        self.assertEqual("quarantined", cleanup.cleanup.document_resources.state)
        self.assertEqual([True], coordinator.release_in_flight)
        self.assertEqual("terminate", turns[0].continuation)
        self.assertIs(cleanup, session.close())
        self.assertEqual(
            1,
            len([call for call in coordinator.calls if call[0] == "release"]),
        )

    def test_close_while_guard_acquisition_blocks_retains_in_flight_fence(self):
        begin_entered = threading.Event()
        resume_begin = threading.Event()
        coordinator = FakeDocumentCoordinator(
            release_state="quarantined",
            begin_wait=resume_begin,
            begin_entered=begin_entered,
        )
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=FakeWriterAdapter(document=object()),
            coordinator=coordinator,
            session_id="session-1",
        )
        turns = []
        worker = threading.Thread(target=lambda: turns.append(session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )))
        worker.start()
        self.assertTrue(begin_entered.wait(timeout=1))

        cleanup = session.close()
        resume_begin.set()
        worker.join(timeout=2)

        self.assertFalse(worker.is_alive())
        self.assertEqual("failed", cleanup.outcome)
        self.assertEqual("quarantined", cleanup.cleanup.document_resources.state)
        self.assertEqual([True], coordinator.release_in_flight)
        self.assertEqual(
            [("release", None, None, None)],
            [call for call in coordinator.calls if call[0] == "release"],
        )
        self.assertEqual("terminate", turns[0].continuation)

    def test_cleanup_timeout_reports_unconfirmed_release_without_blocking(self):
        release = threading.Event()
        coordinator = FakeDocumentCoordinator(release_wait=release)
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=FakeWriterAdapter(document=object()),
            coordinator=coordinator,
            session_id="session-1",
            cleanup_timeout_seconds=0.02,
        )
        session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="createDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        started = time.monotonic()
        cleanup = session.close()
        elapsed = time.monotonic() - started
        release.set()

        self.assertLess(elapsed, 0.5)
        self.assertEqual("failed", cleanup.outcome)
        self.assertEqual(
            "release_unconfirmed",
            cleanup.cleanup.document_resources.state,
        )

    def test_unbound_cleanup_is_idempotent_and_stops_new_dispatch(self):
        adapter = FakeWriterAdapter()
        coordinator = FakeDocumentCoordinator()
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="inspectDocument",
                binding_role="required",
                risk="read",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=coordinator,
            session_id="session-1",
        )

        first = session.close()
        second = session.close()
        turn = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        self.assertEqual("succeeded", first.outcome)
        self.assertEqual("not_acquired", first.cleanup.document_resources.state)
        self.assertIs(first, second)
        self.assertEqual("failed", turn.response.outcome)
        self.assertEqual("RUNTIME_CLOSED", turn.response.error.code)
        self.assertEqual("terminate", turn.continuation)
        self.assertEqual([], adapter.calls)
        self.assertEqual([], coordinator.calls)

    def test_control_plane_error_is_not_rejected_by_action_error_whitelist(self):
        contracts = word_contracts(
            ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
                stable_errors=("APPLICATION_FAILURE",),
            ),
            ActionContract(
                name="inspectDocument",
                binding_role="required",
                risk="read",
                parameters={"type": "object"},
                result={"type": "object"},
                stable_errors=("APPLICATION_FAILURE",),
            ),
        )
        session = ActionSession(
            application="word",
            contracts=contracts,
            adapter=FakeWriterAdapter(document=object()),
            coordinator=FakeDocumentCoordinator(),
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
        self.assertEqual("succeeded", established.response.outcome)
        controller = session._runtime._controller
        controller.close()

        turn = session.execute(
            ActionRequest(
                address=ActionAddress(
                    app="word",
                    action="inspectDocument",
                ),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        self.assertEqual("failed", turn.response.outcome)
        self.assertEqual("RUNTIME_CLOSED", turn.response.error.code)
        self.assertEqual("terminate", turn.continuation)

    def test_adapter_cannot_impersonate_a_control_plane_error(self):
        contracts = word_contracts(
            ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
                stable_errors=("APPLICATION_FAILURE",),
            ),
            ActionContract(
                name="inspectDocument",
                binding_role="required",
                risk="read",
                parameters={"type": "object"},
                result={"type": "object"},
                stable_errors=("APPLICATION_FAILURE",),
            ),
        )
        adapter = FakeWriterAdapter(
            document=object(),
            handler_script=[ControllerResult.failed(
                error=ActionError(
                    code="RUNTIME_CLOSED",
                    message="forged by adapter",
                ),
                controller_state="usable",
                binding_disposition="unchanged",
            )],
        )
        session = ActionSession(
            application="word",
            contracts=contracts,
            adapter=adapter,
            coordinator=FakeDocumentCoordinator(),
            session_id="session-1",
        )
        session.execute(
            ActionRequest(
                address=ActionAddress(
                    app="word",
                    action="createDocument",
                ),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        turn = session.execute(
            ActionRequest(
                address=ActionAddress(
                    app="word",
                    action="inspectDocument",
                ),
                params={},
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )

        self.assertEqual("failed", turn.response.outcome)
        self.assertEqual("INVALID_RESULT", turn.response.error.code)
        self.assertEqual("terminate", turn.continuation)

    def test_semantic_result_failure_uses_the_action_risk_mapping(self):
        always_invalid = lambda params, result: "request/result mismatch"
        contracts = word_contracts(
            ActionContract(
                name="createDocument",
                binding_role="establish",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
            ),
            ActionContract(
                name="writeContent",
                binding_role="required",
                risk="write",
                parameters={"type": "object"},
                result={"type": "object"},
                semantic_validator=always_invalid,
            ),
            ActionContract(
                name="inspectDocument",
                binding_role="required",
                risk="read",
                parameters={"type": "object"},
                result={"type": "object"},
                semantic_validator=always_invalid,
            ),
        )
        session = ActionSession(
            application="word",
            contracts=contracts,
            adapter=FakeWriterAdapter(
                document=object(),
                handler_script=[
                    ControllerResult.succeeded(
                        data={},
                        controller_state="usable",
                        binding_disposition="unchanged",
                    ),
                    ControllerResult.succeeded(
                        data={},
                        controller_state="usable",
                        binding_disposition="unchanged",
                    ),
                ],
            ),
            coordinator=FakeDocumentCoordinator(),
            session_id="session-1",
        )
        session.execute(
            ActionRequest(
                address=ActionAddress(
                    app="word",
                    action="createDocument",
                ),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )

        write_turn = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="writeContent"),
                params={},
            ),
            TraceContext(trace_id="trace-2", trace_log=None),
        )
        read_turn = session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="inspectDocument"),
                params={},
            ),
            TraceContext(trace_id="trace-3", trace_log=None),
        )

        self.assertEqual("unknown", write_turn.response.outcome)
        self.assertEqual("INVALID_RESULT", write_turn.response.error.code)
        self.assertEqual("failed", read_turn.response.outcome)
        self.assertEqual("INVALID_RESULT", read_turn.response.error.code)

    def test_close_before_lazy_controller_creation_prevents_dispatch(self):
        request_id_entered = threading.Event()
        resume_request_id = threading.Event()

        def request_id_factory():
            request_id_entered.set()
            resume_request_id.wait(timeout=2)
            return "request-1"

        adapter = FakeWriterAdapter()
        session = ActionSession(
            application="word",
            contracts=word_contracts(ActionContract(
                name="status",
                binding_role="none",
                risk="read",
                parameters={"type": "object"},
                result={"type": "object"},
            )),
            adapter=adapter,
            coordinator=FakeDocumentCoordinator(),
            session_id="session-1",
            request_id_factory=request_id_factory,
        )
        turns = []
        worker = threading.Thread(target=lambda: turns.append(session.execute(
            ActionRequest(
                address=ActionAddress(app="word", action="status"),
                params={},
            ),
            TraceContext(trace_id="trace-1", trace_log=None),
        )))
        worker.start()
        self.assertTrue(request_id_entered.wait(timeout=1))

        cleanup = session.close()
        resume_request_id.set()
        worker.join(timeout=2)

        self.assertEqual("not_acquired", cleanup.cleanup.document_resources.state)
        self.assertFalse(worker.is_alive())
        self.assertEqual("failed", turns[0].response.outcome)
        self.assertEqual("RUNTIME_CLOSED", turns[0].response.error.code)
        self.assertEqual("terminate", turns[0].continuation)
        self.assertEqual([], adapter.calls)


if __name__ == "__main__":
    unittest.main()
