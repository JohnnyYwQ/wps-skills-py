#!/usr/bin/env python3
"""Run one canonical application-scoped WPS Session Host."""

import argparse
import json
from pathlib import Path
import sys
import uuid

from wps_skills.core.trace_journal import JsonlTraceJournal
from wps_skills.host.session_host import SessionHost, SessionStartupError


MAIN_SOURCE_SET = Path(__file__).resolve().parents[3]
WRITER_BRIDGE_SCRIPT = (
    MAIN_SOURCE_SET
    / "resources"
    / "wps_skills"
    / "word"
    / "windows"
    / "writer_bridge.ps1"
)


def _build_word_session(
    *,
    session_id,
    debug_close_created_document=False,
):
    from wps_skills.core.action_session import ActionSession
    from wps_skills.windows.powershell_writer_bridge import JsonLineWriterBridgeTransport
    from wps_skills.windows.document_coordinator import WindowsDocumentCoordinator
    from wps_skills.windows.owned_process import WindowsOwnedProcessLauncher
    from wps_skills.windows.writer_backend import WindowsWriterBackend
    from wps_skills.windows.writer_runtime import LazyWindowsWriterBridge
    from wps_skills.word.adapter import WordAdapter
    from wps_skills.word.contracts import WORD_PRODUCTION_CONTRACT_SET
    from wps_skills.word.handlers import WORD_HANDLERS

    launcher = None
    try:
        launcher = WindowsOwnedProcessLauncher()
        bridge = LazyWindowsWriterBridge(
            launcher=launcher,
            script_path=WRITER_BRIDGE_SCRIPT,
            transport_factory=JsonLineWriterBridgeTransport,
            debug_close_created_document=debug_close_created_document,
        )
        backend = WindowsWriterBackend(bridge=bridge)
        adapter = WordAdapter(
            backend=backend,
            handlers=WORD_HANDLERS,
            contracts=WORD_PRODUCTION_CONTRACT_SET,
        )
        return ActionSession(
            application="word",
            contracts=WORD_PRODUCTION_CONTRACT_SET,
            adapter=adapter,
            coordinator=WindowsDocumentCoordinator(bridge=bridge),
            session_id=session_id,
            launcher=launcher,
        )
    except Exception as exc:
        if launcher is not None:
            try:
                launcher.close()
            except Exception:
                pass
        raise SessionStartupError(
            "the production Word Action Session could not be constructed"
        ) from exc


def _parser():
    parser = argparse.ArgumentParser(description="Discover Actions or run a WPS Action Session Host")
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--session", action="store_true")
    mode.add_argument("--index", action="store_true", help="Print the production Action Index without starting WPS")
    mode.add_argument("--resolve", nargs="+", metavar="ACTION", help="Resolve complete production Action Contracts without executing them")
    parser.add_argument(
        "--app",
        choices=("excel", "ppt", "word"),
        required=True,
    )
    parser.add_argument(
        "--debug-close-created-document",
        action="store_true",
        help=(
            "debug only: discard and close a document created by this "
            "Session during cleanup"
        ),
    )
    return parser


def _discover(args, output_stream, error_stream):
    if args.app != "word":
        error_stream.write(f"WPS_DISCOVERY_UNAVAILABLE app={args.app}\n")
        return 4
    from wps_skills.core.action_session import ActionAddress
    from wps_skills.word.contracts import WORD_PRODUCTION_CONTRACT_SET

    contracts = WORD_PRODUCTION_CONTRACT_SET
    if args.index:
        value = {
            "app": args.app,
            "actions": [entry.to_wire() for entry in contracts.action_index()],
        }
    else:
        value = contracts.batch_resolve([
            ActionAddress(app=args.app, action=action) for action in args.resolve
        ])
    output_stream.write(json.dumps(value, ensure_ascii=True, indent=2) + "\n")
    return 0 if args.index or value["status"] == "complete" else 2


def _production_session_factory(
    *,
    application,
    session_id,
    debug_close_created_document=False,
):
    if application == "word":
        return _build_word_session(
            session_id=session_id,
            debug_close_created_document=debug_close_created_document,
        )
    raise SessionStartupError(
        f"no production {application} Application Contract Set is installed"
    )


def main(
    argv=None,
    *,
    input_stream=None,
    output_stream=None,
    error_stream=None,
):
    parser = _parser()
    args = parser.parse_args(argv)
    if args.debug_close_created_document and not args.session:
        parser.error("--debug-close-created-document requires --session")
    input_stream = input_stream or sys.stdin
    output_stream = output_stream or sys.stdout
    error_stream = error_stream or sys.stderr
    if not args.session:
        return _discover(args, output_stream, error_stream)
    trace_journal = JsonlTraceJournal.default()
    host = SessionHost(
        session_factory=lambda **kwargs: _production_session_factory(
            **kwargs,
            debug_close_created_document=(
                args.debug_close_created_document
            ),
        ),
        session_id_factory=lambda: f"session-{uuid.uuid4().hex}",
        trace_log_factory=trace_journal.session_log,
        action_trace_factory=trace_journal.action_trace,
        session_event_sink=trace_journal.session_event,
    )
    return host.serve(
        application=args.app,
        input_stream=input_stream,
        output_stream=output_stream,
        error_stream=error_stream,
    )


if __name__ == "__main__":
    raise SystemExit(main())
