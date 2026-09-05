"""Windows Writer Backend over one Session-owned local automation bridge."""

from dataclasses import dataclass
from typing import Any, Mapping, Optional, Protocol, runtime_checkable

from wps_skills.core.action_session import ControllerContext
from wps_skills.word.adapter import (
    WriterBackendAcquisition,
    WriterBackendActionFailure,
    WriterBackendOperation,
    WriterBackendPreparation,
    WriterDefiniteEstablishFailure,
    WriterUnprovableEstablishFailure,
)


@runtime_checkable
class WindowsWriterBridge(Protocol):
    """Private synchronous seam to the Session-owned PowerShell bridge."""

    def execute(
        self,
        operation: str,
        arguments: Mapping[str, Any],
        context: Optional[ControllerContext],
    ) -> Mapping[str, Any]:
        ...

    def close(self) -> bool:
        ...


@dataclass(frozen=True)
class WindowsWriterPreparation:
    preparation_id: str
    coordination_identity: str
    authorized_path: Optional[str]

    def __post_init__(self) -> None:
        if not isinstance(self.preparation_id, str) or not self.preparation_id:
            raise ValueError("bridge preparation id must be non-empty")
        if (
            not isinstance(self.coordination_identity, str)
            or not self.coordination_identity
        ):
            raise ValueError("coordination identity must be non-empty")


@dataclass(frozen=True)
class WindowsWriterDocument:
    """Opaque in-process reference to the bridge's one exact COM document."""

    bridge_document_id: str
    authorized_path: Optional[str]

    def __post_init__(self) -> None:
        if (
            not isinstance(self.bridge_document_id, str)
            or not self.bridge_document_id
        ):
            raise ValueError("bridge document id must be non-empty")
        if self.authorized_path is not None and (
            not isinstance(self.authorized_path, str)
            or not self.authorized_path
        ):
            raise ValueError("authorized path must be non-empty when present")


def _require_mapping(value, description):
    if not isinstance(value, Mapping):
        raise TypeError(f"{description} must be an object")
    return value


def _require_exact_keys(value, expected, description):
    value = _require_mapping(value, description)
    if set(value) != set(expected):
        raise TypeError(f"{description} has invalid fields")
    return value


def _require_string(value, description):
    if not isinstance(value, str) or not value:
        raise TypeError(f"{description} must be a non-empty string")
    return value


def _require_boolean(value, description):
    if not isinstance(value, bool):
        raise TypeError(f"{description} must be Boolean")
    return value


class WindowsWriterBackend:
    """Validate bridge facts and preserve one exact bridge document reference."""

    _CREATE_FIELDS = frozenset({
        "documentId",
        "revision",
        "persistenceState",
        "readOnly",
    })
    _OPEN_FIELDS = _CREATE_FIELDS | frozenset({
        "artifactFormat",
        "artifactSizeBytes",
    })

    def __init__(self, *, bridge: WindowsWriterBridge):
        if not isinstance(bridge, WindowsWriterBridge):
            raise ValueError("bridge must satisfy the Windows Writer bridge seam")
        self._bridge = bridge
        self._preparation = None
        self._document = None

    def _execute(self, operation, arguments, context):
        result = self._bridge.execute(operation, arguments, context)
        return _require_mapping(result, "Windows Writer bridge result")

    def _begin_acquisition(self):
        if self._document is not None:
            raise RuntimeError("Windows Writer Backend is already bound")

    def _prepare(self, operation, arguments, context, *, path=None):
        self._begin_acquisition()
        try:
            result = _require_exact_keys(
                self._execute(operation, arguments, context),
                {"preparationId", "coordinationIdentity"},
                "Windows Writer preparation result",
            )
        except WriterBackendActionFailure as failure:
            if (
                failure.outcome == "failed"
                and failure.binding_disposition == "unchanged"
            ):
                self._preparation = None
                raise WriterDefiniteEstablishFailure(
                    code=failure.code,
                    message=failure.message,
                ) from failure
            raise WriterUnprovableEstablishFailure(
                outcome=failure.outcome,
                code=failure.code,
                message=failure.message,
            ) from failure
        state = WindowsWriterPreparation(
            preparation_id=_require_string(
                result["preparationId"],
                "bridge preparation id",
            ),
            coordination_identity=_require_string(
                result["coordinationIdentity"],
                "coordination identity",
            ),
            authorized_path=path,
        )
        preparation = WriterBackendPreparation(
            coordination_identity=state.coordination_identity,
            state=state,
        )
        self._preparation = preparation
        return preparation

    def prepare_create_document(self, context):
        return self._prepare(
            "prepare_new_document",
            {},
            context,
        )

    def prepare_open_document(self, path, context):
        if not isinstance(path, str) or not path:
            raise ValueError("open path must be non-empty")
        return self._prepare(
            "prepare_existing_document",
            {"path": path},
            context,
            path=path,
        )

    def _require_preparation(self, preparation, *, path):
        if preparation is not self._preparation:
            raise ValueError(
                "Windows Writer Backend received another preparation"
            )
        if not isinstance(preparation, WriterBackendPreparation) or not isinstance(
            preparation.state,
            WindowsWriterPreparation,
        ):
            raise ValueError("Windows Writer Backend preparation is invalid")
        if preparation.state.authorized_path != path:
            raise ValueError("Windows Writer preparation path does not match")
        return preparation.state

    @staticmethod
    def _acquisition(result, *, authorized_path, opened):
        fields = (
            WindowsWriterBackend._OPEN_FIELDS
            if opened
            else WindowsWriterBackend._CREATE_FIELDS
        )
        result = _require_exact_keys(
            result,
            fields,
            "Windows Writer acquisition result",
        )
        document = WindowsWriterDocument(
            bridge_document_id=_require_string(
                result["documentId"],
                "bridge document id",
            ),
            authorized_path=authorized_path,
        )
        revision = _require_string(result["revision"], "Content Revision")
        persistence_state = _require_string(
            result["persistenceState"],
            "persistence state",
        )
        read_only = _require_boolean(result["readOnly"], "read-only state")
        artifact_format = None
        artifact_size_bytes = None
        if opened:
            artifact_format = result["artifactFormat"]
            artifact_size_bytes = result["artifactSizeBytes"]
        return document, WriterBackendAcquisition(
            document=document,
            revision=revision,
            persistence_state=persistence_state,
            read_only=read_only,
            artifact_format=artifact_format,
            artifact_size_bytes=artifact_size_bytes,
        )

    def _establish(self, operation, arguments, context, *, path=None):
        self._begin_acquisition()
        try:
            result = self._execute(operation, arguments, context)
            document, acquisition = self._acquisition(
                result,
                authorized_path=path,
                opened=path is not None,
            )
        except WriterBackendActionFailure as failure:
            if (
                failure.outcome == "failed"
                and failure.binding_disposition == "unchanged"
            ):
                self._preparation = None
                raise WriterDefiniteEstablishFailure(
                    code=failure.code,
                    message=failure.message,
                ) from failure
            raise WriterUnprovableEstablishFailure(
                outcome=failure.outcome,
                code=failure.code,
                message=failure.message,
            ) from failure
        self._document = document
        return acquisition

    def create_document(self, preparation, context):
        state = self._require_preparation(preparation, path=None)
        return self._establish(
            "acquire_new_document",
            {"preparationId": state.preparation_id},
            context,
        )

    def open_document(self, preparation, path, context):
        if not isinstance(path, str) or not path:
            raise ValueError("open path must be non-empty")
        state = self._require_preparation(preparation, path=path)
        return self._establish(
            "acquire_existing_document",
            {"preparationId": state.preparation_id},
            context,
            path=path,
        )

    def _require_bound_document(self, document):
        if document is not self._document:
            raise ValueError(
                "Windows Writer Backend received another document"
            )
        return document

    def is_document_live(self, document):
        document = self._require_bound_document(document)
        result = self._execute(
            "probe_bound_document",
            {"documentId": document.bridge_document_id},
            None,
        )
        result = _require_exact_keys(
            result,
            {"live"},
            "Windows Writer liveness result",
        )
        return _require_boolean(result["live"], "document liveness")

    def invoke(self, document, operation, context):
        document = self._require_bound_document(document)
        if not isinstance(operation, WriterBackendOperation):
            raise ValueError("Windows Writer Backend requires an operation")
        arguments = {
            "documentId": document.bridge_document_id,
            "operationArguments": dict(operation.arguments),
        }
        if document.authorized_path is not None:
            arguments["authorizedPath"] = document.authorized_path
        return self._execute(operation.name, arguments, context)

    def close(self):
        return self._bridge.close()
