import unittest

from wps_skills.core.action_session import (
    ControllerContext,
    DefiniteEstablishFailure,
)
from wps_skills.windows.document_coordinator import WindowsDocumentCoordinator
from wps_skills.windows.writer_backend import WindowsWriterDocument
from wps_skills.word.adapter import WriterBackendActionFailure


def context():
    return ControllerContext(
        request_id="request-1",
        trace_id="trace-1",
        deadline_at=100,
    )


class RecordingBridge:
    def __init__(self, replies):
        self.replies = dict(replies)
        self.calls = []

    def execute(self, operation, arguments, request_context):
        self.calls.append((operation, dict(arguments), request_context))
        reply = self.replies[operation]
        if isinstance(reply, BaseException):
            raise reply
        return reply


class WindowsDocumentCoordinatorTests(unittest.TestCase):
    def test_guard_promotes_to_lease_and_releases_without_a_gap(self):
        bridge = RecordingBridge({
            "acquire_coordination_guard": {"guardId": "guard-1"},
            "commit_document_lease": {"leaseId": "lease-1"},
            "release_document_resources": {"state": "released"},
        })
        coordinator = WindowsDocumentCoordinator(bridge=bridge)
        document = WindowsWriterDocument(
            bridge_document_id="document-1",
            authorized_path="C:/docs/report.docx",
        )

        guard = coordinator.begin("file-identity-1", context())
        lease = coordinator.commit(guard, document, context())
        cleanup = coordinator.release(None, document, lease)

        self.assertEqual("released", cleanup.state)
        self.assertEqual(
            [
                "acquire_coordination_guard",
                "commit_document_lease",
                "release_document_resources",
            ],
            [call[0] for call in bridge.calls],
        )
        self.assertEqual("guard-1", bridge.calls[1][1]["guardId"])
        self.assertEqual("document-1", bridge.calls[1][1]["documentId"])
        self.assertEqual("lease-1", bridge.calls[2][1]["leaseId"])

    def test_conflict_is_a_definite_not_established_failure(self):
        bridge = RecordingBridge({
            "acquire_coordination_guard": WriterBackendActionFailure(
                outcome="failed",
                code="DOCUMENT_LEASE_CONFLICT",
                message="owned elsewhere",
                binding_disposition="unchanged",
            ),
        })
        coordinator = WindowsDocumentCoordinator(bridge=bridge)

        with self.assertRaises(DefiniteEstablishFailure) as raised:
            coordinator.begin("file-identity-1", context())

        self.assertEqual("DOCUMENT_LEASE_CONFLICT", raised.exception.code)

    def test_in_flight_cleanup_retains_quarantine_without_channel_reentry(self):
        bridge = RecordingBridge({
            "acquire_coordination_guard": {"guardId": "guard-1"},
        })
        coordinator = WindowsDocumentCoordinator(bridge=bridge)
        guard = coordinator.begin("file-identity-1", context())

        cleanup = coordinator.release(
            guard,
            None,
            None,
            in_flight=True,
        )

        self.assertEqual("quarantined", cleanup.state)
        self.assertEqual(1, len(bridge.calls))


if __name__ == "__main__":
    unittest.main()
