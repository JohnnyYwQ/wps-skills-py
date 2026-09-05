"""First real Word Action handlers behind the Word Adapter seam."""

from types import MappingProxyType
from typing import Mapping

from wps_skills.core.action_session import ActionError, ControllerResult
from wps_skills.word.adapter import (
    WriterBackendActionFailure,
    WriterBackendOperation,
)


def _failure_result(failure):
    factory = (
        ControllerResult.failed
        if failure.outcome == "failed"
        else ControllerResult.unknown
    )
    return factory(
        error=ActionError(
            code=failure.code,
            message=failure.message,
        ),
        controller_state=(
            "usable"
            if failure.binding_disposition == "unchanged"
            else "broken"
        ),
        binding_disposition=failure.binding_disposition,
    )


def _invoke(backend, document, operation, context):
    try:
        data = backend.invoke(document, operation, context)
    except WriterBackendActionFailure as failure:
        return _failure_result(failure)
    if not isinstance(data, Mapping):
        raise TypeError("Writer Backend Action result must be an object")
    return ControllerResult.succeeded(
        data=data,
        controller_state="usable",
        binding_disposition="unchanged",
    )


_PRIVATE_OPERATIONS = MappingProxyType({
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
})


def _handler(operation_name):
    def invoke(backend, document, params, context):
        return _invoke(
            backend,
            document,
            WriterBackendOperation(
                name=operation_name,
                arguments=params,
            ),
            context,
        )

    return invoke


WORD_HANDLERS = MappingProxyType({
    action: _handler(operation)
    for action, operation in _PRIVATE_OPERATIONS.items()
})
