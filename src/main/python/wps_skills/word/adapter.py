"""Word Application Adapter and platform Writer Backend seam."""

from dataclasses import dataclass
from types import MappingProxyType
from typing import (
    Any,
    Callable,
    Mapping,
    Optional,
    Protocol,
    runtime_checkable,
)

from wps_skills.core.action_session import (
    AcquiredDocument,
    ActionError,
    ControllerCommand,
    ControllerContext,
    ControllerResult,
    DefiniteEstablishFailure,
    PreparedDocumentAcquisition,
    UnprovableEstablishFailure,
)
from wps_skills.word.contracts import WORD_TARGET_CONTRACT_SET


_WORD_ACTION_NAMES = frozenset(WORD_TARGET_CONTRACT_SET.action_names)


def _freeze_value(value):
    if isinstance(value, Mapping):
        return MappingProxyType({
            key: _freeze_value(item)
            for key, item in value.items()
        })
    if isinstance(value, (list, tuple)):
        return tuple(_freeze_value(item) for item in value)
    return value


@dataclass(frozen=True)
class WriterBackendOperation:
    """Backend-local operation produced by a Word Action handler."""

    name: str
    arguments: Mapping[str, Any]

    def __post_init__(self) -> None:
        if not isinstance(self.name, str) or not self.name:
            raise ValueError("Writer Backend operation name must be non-empty")
        if self.name in _WORD_ACTION_NAMES:
            raise ValueError(
                "Writer Backend operation must not reuse a public Action name"
            )
        if not isinstance(self.arguments, Mapping):
            raise ValueError("Writer Backend operation arguments must be an object")
        object.__setattr__(self, "arguments", _freeze_value(self.arguments))


@dataclass(frozen=True)
class WriterBackendAcquisition:
    """One exact document plus platform observations needed by establish."""

    document: Any
    revision: str
    persistence_state: str
    read_only: bool
    artifact_format: Optional[str] = None
    artifact_size_bytes: Optional[int] = None

    def __post_init__(self) -> None:
        if self.document is None:
            raise ValueError("Writer Backend acquisition requires a document")
        if not isinstance(self.revision, str) or not self.revision:
            raise ValueError("Writer Backend acquisition requires a revision")
        if self.persistence_state not in {"unsaved", "saved", "modified"}:
            raise ValueError("invalid Writer Backend persistence state")
        if not isinstance(self.read_only, bool):
            raise ValueError("Writer Backend read-only state must be Boolean")
        if (
            self.artifact_format is not None
            and self.artifact_format != "docx"
        ):
            raise ValueError("Writer establish artifacts must use docx")
        if (
            self.artifact_size_bytes is not None
            and (
                not isinstance(self.artifact_size_bytes, int)
                or isinstance(self.artifact_size_bytes, bool)
                or self.artifact_size_bytes < 1
            )
        ):
            raise ValueError("Writer establish artifact size must be positive")


@dataclass(frozen=True)
class WriterBackendPreparation:
    """Opaque platform preparation plus its cross-process identity."""

    coordination_identity: str
    state: Any

    def __post_init__(self) -> None:
        if (
            not isinstance(self.coordination_identity, str)
            or not self.coordination_identity
        ):
            raise ValueError("Writer coordination identity must be non-empty")


class WriterDefiniteEstablishFailure(Exception):
    """Backend proved that establish produced no document or binding effect."""

    def __init__(self, *, code: str, message: str):
        super().__init__(message)
        self.code = code
        self.message = message


class WriterUnprovableEstablishFailure(Exception):
    """Backend cannot prove the complete establish effect."""

    def __init__(
        self,
        *,
        outcome: str,
        code: str,
        message: str,
        partial_document: Any = None,
    ):
        super().__init__(message)
        if outcome not in {"failed", "unknown"}:
            raise ValueError("invalid Writer establish outcome")
        self.outcome = outcome
        self.code = code
        self.message = message
        self.partial_document = partial_document


class WriterBackendActionFailure(Exception):
    """Closed operation failure reported by a Writer Backend."""

    def __init__(
        self,
        *,
        outcome: str,
        code: str,
        message: str,
        binding_disposition: str,
    ):
        super().__init__(message)
        if outcome not in {"failed", "unknown"}:
            raise ValueError("invalid Writer Action failure outcome")
        if not isinstance(code, str) or not code:
            raise ValueError("Writer Action failure code must be non-empty")
        if not isinstance(message, str) or not message:
            raise ValueError("Writer Action failure message must be non-empty")
        if binding_disposition not in {
            "unchanged",
            "lost",
            "unprovable",
        }:
            raise ValueError("invalid Writer Action binding disposition")
        if outcome == "unknown" and binding_disposition == "lost":
            raise ValueError(
                "an unknown Writer Action cannot prove binding loss"
            )
        self.outcome = outcome
        self.code = code
        self.message = message
        self.binding_disposition = binding_disposition


@runtime_checkable
class WriterBackend(Protocol):
    """Platform seam used inside the Word Adapter implementation."""

    def prepare_create_document(
        self,
        context: ControllerContext,
    ) -> WriterBackendPreparation:
        ...

    def prepare_open_document(
        self,
        path: str,
        context: ControllerContext,
    ) -> WriterBackendPreparation:
        ...

    def create_document(
        self,
        preparation: WriterBackendPreparation,
        context: ControllerContext,
    ) -> WriterBackendAcquisition:
        ...

    def open_document(
        self,
        preparation: WriterBackendPreparation,
        path: str,
        context: ControllerContext,
    ) -> WriterBackendAcquisition:
        ...

    def is_document_live(self, document: Any) -> bool:
        ...

    def invoke(
        self,
        document: Any,
        operation: WriterBackendOperation,
        context: ControllerContext,
    ) -> Mapping[str, Any]:
        ...

    def close(self) -> bool:
        ...


WordActionHandler = Callable[
    [WriterBackend, Any, Mapping[str, Any], ControllerContext],
    ControllerResult,
]


class WordAdapter:
    """Core-facing Word Adapter; platform behavior stays behind WriterBackend."""

    application = "word"
    contracts = WORD_TARGET_CONTRACT_SET
    _ESTABLISH_ACTIONS = frozenset({"createDocument", "openDocument"})
    _REQUIRED_ACTIONS = frozenset(
        contract.name
        for contract in WORD_TARGET_CONTRACT_SET.contracts
        if contract.binding_role == "required"
    )

    def __init__(
        self,
        *,
        backend: WriterBackend,
        handlers: Mapping[str, WordActionHandler],
        contracts=WORD_TARGET_CONTRACT_SET,
    ):
        if not isinstance(backend, WriterBackend):
            raise ValueError("backend must satisfy the Writer Backend seam")
        if not isinstance(handlers, Mapping):
            raise ValueError("Word handlers must be a mapping")
        if getattr(contracts, "application", None) != "word":
            raise ValueError("Word contracts must belong to word")
        required_actions = frozenset(
            contract.name
            for contract in contracts.contracts
            if contract.binding_role == "required"
        )
        invalid = [
            name
            for name, handler in handlers.items()
            if (
                not isinstance(name, str)
                or not name
                or name not in required_actions
                or not callable(handler)
            )
        ]
        if invalid:
            raise ValueError("invalid Word Action handler mapping")
        missing = required_actions.difference(handlers)
        if contracts is not WORD_TARGET_CONTRACT_SET and missing:
            raise ValueError("registered Word contracts require every handler")
        self.contracts = contracts
        self._backend = backend
        self._handlers = MappingProxyType(dict(handlers))

    @staticmethod
    def _require_word_command(command: ControllerCommand) -> None:
        if not isinstance(command, ControllerCommand):
            raise ValueError("Word Adapter requires a Controller Command")
        if command.address.app != "word":
            raise ValueError("Word Adapter received another application")

    @staticmethod
    def _require_acquisition(acquisition) -> WriterBackendAcquisition:
        if not isinstance(acquisition, WriterBackendAcquisition):
            raise TypeError("Writer Backend returned an invalid acquisition")
        return acquisition

    @classmethod
    def _created(cls, acquisition) -> AcquiredDocument:
        acquisition = cls._require_acquisition(acquisition)
        if (
            acquisition.persistence_state != "unsaved"
            or acquisition.read_only
            or acquisition.artifact_format is not None
            or acquisition.artifact_size_bytes is not None
        ):
            raise TypeError("Writer Backend returned invalid create observations")
        return AcquiredDocument(
            document=acquisition.document,
            data={
                "revision": acquisition.revision,
                "documentState": {
                    "persistenceState": acquisition.persistence_state,
                    "readOnly": acquisition.read_only,
                },
            },
        )

    @classmethod
    def _opened(cls, acquisition, requested_path) -> AcquiredDocument:
        acquisition = cls._require_acquisition(acquisition)
        if (
            acquisition.persistence_state not in {"saved", "modified"}
            or acquisition.artifact_format != "docx"
            or acquisition.artifact_size_bytes is None
        ):
            raise TypeError("Writer Backend returned invalid open observations")
        return AcquiredDocument(
            document=acquisition.document,
            data={
                "revision": acquisition.revision,
                "artifact": {
                    "path": requested_path,
                    "format": acquisition.artifact_format,
                    "sizeBytes": acquisition.artifact_size_bytes,
                },
                "documentState": {
                    "persistenceState": acquisition.persistence_state,
                    "readOnly": acquisition.read_only,
                },
            },
        )

    @staticmethod
    def _translate_establish_failure(exc):
        if isinstance(exc, WriterDefiniteEstablishFailure):
            raise DefiniteEstablishFailure(
                code=exc.code,
                message=exc.message,
            ) from exc
        if isinstance(exc, WriterUnprovableEstablishFailure):
            raise UnprovableEstablishFailure(
                outcome=exc.outcome,
                code=exc.code,
                message=exc.message,
                partial_document=exc.partial_document,
            ) from exc
        raise exc

    def prepare_establish(
        self,
        command: ControllerCommand,
    ) -> PreparedDocumentAcquisition:
        self._require_word_command(command)
        try:
            if command.address.action == "createDocument":
                preparation = self._backend.prepare_create_document(
                    command.context
                )
            elif command.address.action == "openDocument":
                preparation = self._backend.prepare_open_document(
                    command.params["path"],
                    command.context,
                )
            else:
                raise DefiniteEstablishFailure(
                    code="INVALID_PARAMS",
                    message="Action is not a Word establish Action",
                )
        except (WriterDefiniteEstablishFailure, WriterUnprovableEstablishFailure) as exc:
            self._translate_establish_failure(exc)
        if not isinstance(preparation, WriterBackendPreparation):
            raise TypeError("Writer Backend returned an invalid preparation")
        return PreparedDocumentAcquisition(
            coordination_identity=preparation.coordination_identity,
            application_state=preparation,
        )

    def establish(
        self,
        prepared: PreparedDocumentAcquisition,
        command: ControllerCommand,
    ) -> AcquiredDocument:
        self._require_word_command(command)
        if not isinstance(prepared, PreparedDocumentAcquisition) or not isinstance(
            prepared.application_state,
            WriterBackendPreparation,
        ):
            raise TypeError("Word Adapter requires its prepared acquisition")
        preparation = prepared.application_state
        if preparation.coordination_identity != prepared.coordination_identity:
            raise TypeError("Word acquisition identities do not match")
        try:
            if command.address.action == "createDocument":
                acquisition = self._backend.create_document(
                    preparation,
                    command.context,
                )
                acquired = self._created(acquisition)
            elif command.address.action == "openDocument":
                acquisition = self._backend.open_document(
                    preparation,
                    command.params["path"],
                    command.context,
                )
                acquired = self._opened(acquisition, command.params["path"])
            else:
                raise DefiniteEstablishFailure(
                    code="INVALID_PARAMS",
                    message="Action is not a Word establish Action",
                )
        except (WriterDefiniteEstablishFailure, WriterUnprovableEstablishFailure) as exc:
            self._translate_establish_failure(exc)
        return acquired

    def is_live(self, document: Any) -> bool:
        live = self._backend.is_document_live(document)
        if not isinstance(live, bool):
            raise TypeError("Writer Backend liveness observation must be Boolean")
        return live

    def handle(
        self,
        document: Any,
        command: ControllerCommand,
    ) -> ControllerResult:
        self._require_word_command(command)
        handler = self._handlers.get(command.address.action)
        if handler is None:
            return ControllerResult.failed(
                error=ActionError(
                    code="WORD_CAPABILITY_UNAVAILABLE",
                    message=(
                        f"Word Action is not implemented: "
                        f"{command.address.action}"
                    ),
                ),
                controller_state="usable",
                binding_disposition="unchanged",
            )
        return handler(
            self._backend,
            document,
            command.params,
            command.context,
        )

    def handle_none(self, command: ControllerCommand) -> ControllerResult:
        self._require_word_command(command)
        raise ValueError("The Word Application Contract Set has no none Action")

    def close(self) -> bool:
        return self._backend.close()
