"""Cross-process Word document guard and Lease over the owned bridge."""

from dataclasses import dataclass

from wps_skills.core.action_session import (
    DefiniteEstablishFailure,
    DocumentResourceCleanup,
    UnprovableEstablishFailure,
)
from wps_skills.word.adapter import WriterBackendActionFailure
from wps_skills.windows.writer_backend import WindowsWriterDocument


@dataclass(frozen=True)
class WindowsDocumentGuard:
    guard_id: str
    coordination_identity: str


@dataclass(frozen=True)
class WindowsDocumentLease:
    lease_id: str
    guard: WindowsDocumentGuard
    document: WindowsWriterDocument


def _nonempty(value, description):
    if not isinstance(value, str) or not value:
        raise TypeError(f"{description} must be a non-empty string")
    return value


class WindowsDocumentCoordinator:
    """Keep the OS guard in the same process as the exact COM binding."""

    def __init__(self, *, bridge):
        self._bridge = bridge
        self._guard = None
        self._lease = None
        self._begin_active = False

    @staticmethod
    def _establish_failure(failure):
        if (
            failure.outcome == "failed"
            and failure.binding_disposition == "unchanged"
        ):
            raise DefiniteEstablishFailure(
                code=failure.code,
                message=failure.message,
            ) from failure
        raise UnprovableEstablishFailure(
            outcome=failure.outcome,
            code=failure.code,
            message=failure.message,
        ) from failure

    def begin(self, coordination_identity, context):
        coordination_identity = _nonempty(
            coordination_identity,
            "coordination identity",
        )
        if self._guard is not None or self._lease is not None:
            raise RuntimeError("document coordination is already active")
        self._begin_active = True
        try:
            try:
                result = self._bridge.execute(
                    "acquire_coordination_guard",
                    {"coordinationIdentity": coordination_identity},
                    context,
                )
            except WriterBackendActionFailure as failure:
                self._establish_failure(failure)
            if not isinstance(result, dict) or set(result) != {"guardId"}:
                raise TypeError("coordination guard result has invalid fields")
            guard = WindowsDocumentGuard(
                guard_id=_nonempty(result["guardId"], "guard id"),
                coordination_identity=coordination_identity,
            )
            self._guard = guard
            return guard
        finally:
            self._begin_active = False

    def commit(self, guard, document, context):
        if guard is not self._guard or not isinstance(
            document,
            WindowsWriterDocument,
        ):
            raise ValueError("another guard or document cannot be committed")
        try:
            result = self._bridge.execute(
                "commit_document_lease",
                {
                    "guardId": guard.guard_id,
                    "documentId": document.bridge_document_id,
                },
                context,
            )
        except WriterBackendActionFailure as failure:
            raise UnprovableEstablishFailure(
                outcome=failure.outcome,
                code=failure.code,
                message=failure.message,
                partial_document=document,
            ) from failure
        if not isinstance(result, dict) or set(result) != {"leaseId"}:
            raise TypeError("document Lease result has invalid fields")
        lease = WindowsDocumentLease(
            lease_id=_nonempty(result["leaseId"], "lease id"),
            guard=guard,
            document=document,
        )
        self._lease = lease
        return lease

    def release(self, guard, document, lease, *, in_flight=False):
        if all(value is None for value in (guard, document, lease)):
            if in_flight and self._begin_active:
                return DocumentResourceCleanup(state="quarantined")
            return DocumentResourceCleanup(state="not_acquired")
        if in_flight:
            return DocumentResourceCleanup(state="quarantined")
        if lease is not None:
            if lease is not self._lease:
                return DocumentResourceCleanup(state="release_unconfirmed")
            effective_guard = lease.guard
        else:
            effective_guard = guard
        if effective_guard is not self._guard:
            return DocumentResourceCleanup(state="release_unconfirmed")
        arguments = {
            "guardId": effective_guard.guard_id,
            "leaseId": lease.lease_id if lease is not None else "",
        }
        try:
            result = self._bridge.execute(
                "release_document_resources",
                arguments,
                None,
            )
        except Exception:
            return DocumentResourceCleanup(state="release_unconfirmed")
        if not isinstance(result, dict) or set(result) != {"state"}:
            return DocumentResourceCleanup(state="release_unconfirmed")
        state = result["state"]
        if state not in {"not_acquired", "released"}:
            return DocumentResourceCleanup(state="release_unconfirmed")
        self._guard = None
        self._lease = None
        return DocumentResourceCleanup(state=state)
