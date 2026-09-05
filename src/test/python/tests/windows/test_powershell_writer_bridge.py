import io
import json
import unittest

from wps_skills.core.action_session import ControllerContext
from wps_skills.windows.powershell_writer_bridge import (
    JsonLineWriterBridgeTransport,
    PowerShellWriterBridge,
    WriterBridgeTransport,
    WriterBridgeTransportError,
)
from wps_skills.windows.writer_backend import WindowsWriterBridge
from wps_skills.word.adapter import WriterBackendActionFailure


class RecordingTransport:
    def __init__(self, response):
        self.response = response
        self.calls = []

    def exchange(self, request, deadline_at):
        self.calls.append((request, deadline_at))
        response = dict(self.response)
        if response.get("requestId") == "<request>":
            response["requestId"] = request["requestId"]
        return response

    def close(self):
        return True


class RecordingInput(io.StringIO):
    def __init__(self):
        super().__init__()
        self.flush_count = 0

    def flush(self):
        self.flush_count += 1
        return super().flush()


class FakeProcess:
    def __init__(self, output="", return_code=None):
        self.stdin = RecordingInput()
        self.stdout = io.StringIO(output)
        self.return_code = return_code
        self.wait_calls = []

    def poll(self):
        return self.return_code

    def wait(self, timeout):
        self.wait_calls.append(timeout)
        self.return_code = 0
        return 0


class PowerShellWriterBridgeTests(unittest.TestCase):
    def test_bridge_and_transport_satisfy_their_seams(self):
        transport = RecordingTransport({
            "requestId": "<request>",
            "outcome": "succeeded",
            "data": {},
        })
        bridge = PowerShellWriterBridge(transport=transport)

        self.assertIsInstance(transport, WriterBridgeTransport)
        self.assertIsInstance(bridge, WindowsWriterBridge)

    def test_successful_exchange_uses_controller_correlation_and_deadline(self):
        transport = RecordingTransport({
            "requestId": "<request>",
            "outcome": "succeeded",
            "data": {"live": True},
        })
        bridge = PowerShellWriterBridge(transport=transport)
        request_context = ControllerContext(
            request_id="request-1",
            trace_id="trace-1",
            deadline_at=123.5,
        )

        result = bridge.execute(
            "probe_bound_document",
            {"nested": {"values": (1, 2)}},
            request_context,
        )

        self.assertEqual({"live": True}, result)
        request, deadline = transport.calls[0]
        self.assertEqual("request-1", request["requestId"])
        self.assertEqual("trace-1", request["traceId"])
        self.assertEqual("probe_bound_document", request["operation"])
        self.assertEqual(
            {"nested": {"values": [1, 2]}},
            request["arguments"],
        )
        self.assertEqual(123.5, deadline)

    def test_liveness_without_action_context_gets_internal_correlation(self):
        transport = RecordingTransport({
            "requestId": "<request>",
            "outcome": "succeeded",
            "data": {"live": True},
        })
        bridge = PowerShellWriterBridge(transport=transport)

        bridge.execute("probe_bound_document", {}, None)

        request, deadline = transport.calls[0]
        self.assertTrue(request["requestId"].startswith("liveness-"))
        self.assertEqual(request["requestId"], request["traceId"])
        self.assertIsNone(deadline)

    def test_closed_failure_becomes_backend_action_failure(self):
        transport = RecordingTransport({
            "requestId": "<request>",
            "outcome": "unknown",
            "error": {
                "code": "OUTPUT_WRITE_FAILED",
                "message": "save outcome is unknown",
            },
            "bindingDisposition": "unprovable",
        })
        bridge = PowerShellWriterBridge(transport=transport)

        with self.assertRaises(WriterBackendActionFailure) as raised:
            bridge.execute(
                "save_existing_artifact",
                {},
                ControllerContext("request-1", "trace-1", 10),
            )

        self.assertEqual("unknown", raised.exception.outcome)
        self.assertEqual("OUTPUT_WRITE_FAILED", raised.exception.code)
        self.assertEqual("unprovable", raised.exception.binding_disposition)

    def test_malformed_or_mismatched_response_is_unprovable(self):
        malformed = PowerShellWriterBridge(transport=RecordingTransport({
            "requestId": "<request>",
            "outcome": "succeeded",
            "data": {},
            "unexpected": True,
        }))
        with self.assertRaises(WriterBackendActionFailure) as malformed_error:
            malformed.execute("read", {}, None)
        self.assertEqual("unknown", malformed_error.exception.outcome)
        self.assertEqual("RESPONSE_LOST", malformed_error.exception.code)
        self.assertEqual(
            "unprovable",
            malformed_error.exception.binding_disposition,
        )

        mismatched = PowerShellWriterBridge(transport=RecordingTransport({
            "requestId": "another-request",
            "outcome": "succeeded",
            "data": {},
        }))
        with self.assertRaises(WriterBackendActionFailure) as mismatch_error:
            mismatched.execute("read", {}, None)
        self.assertEqual("unknown", mismatch_error.exception.outcome)
        self.assertEqual("RESPONSE_LOST", mismatch_error.exception.code)
        self.assertEqual(
            "unprovable",
            mismatch_error.exception.binding_disposition,
        )

    def test_transport_failure_is_unprovable(self):
        class FailedTransport:
            def exchange(self, request, deadline_at):
                raise WriterBridgeTransportError("channel failed")

            def close(self):
                return True

        bridge = PowerShellWriterBridge(transport=FailedTransport())

        with self.assertRaises(WriterBackendActionFailure) as raised:
            bridge.execute("save_existing_artifact", {}, None)

        self.assertEqual("unknown", raised.exception.outcome)
        self.assertEqual("RESPONSE_LOST", raised.exception.code)
        self.assertEqual("unprovable", raised.exception.binding_disposition)

    def test_non_json_arguments_are_rejected_before_transport(self):
        transport = RecordingTransport({
            "requestId": "<request>",
            "outcome": "succeeded",
            "data": {},
        })
        bridge = PowerShellWriterBridge(transport=transport)

        with self.assertRaisesRegex(ValueError, "only JSON values"):
            bridge.execute("read", {"bad": object()}, None)
        self.assertEqual([], transport.calls)


class JsonLineWriterBridgeTransportTests(unittest.TestCase):
    def test_exchange_writes_one_compact_record_and_reads_one_object(self):
        process = FakeProcess('{"requestId":"request-1","ok":true}\n')
        transport = JsonLineWriterBridgeTransport(
            process=process,
            clock=lambda: 10,
        )

        response = transport.exchange(
            {"requestId": "request-1", "value": "文档"},
            20,
        )

        self.assertEqual(
            {"requestId": "request-1", "ok": True},
            response,
        )
        self.assertEqual(
            {"requestId": "request-1", "value": "文档"},
            json.loads(process.stdin.getvalue()),
        )
        self.assertIn(r"\u6587\u6863", process.stdin.getvalue())
        self.assertNotIn("文档", process.stdin.getvalue())
        self.assertEqual(1, process.stdin.flush_count)

    def test_invalid_output_fails_closed(self):
        for output in (
            "[]\n",
            '{"a":1,"a":2}\n',
            '{"value":NaN}\n',
            "not-json\n",
            "",
        ):
            with self.subTest(output=output):
                transport = JsonLineWriterBridgeTransport(
                    process=FakeProcess(output),
                )
                with self.assertRaises(WriterBridgeTransportError):
                    transport.exchange({"requestId": "request-1"}, None)

    def test_dead_process_and_expired_deadline_do_not_dispatch(self):
        dead = FakeProcess(return_code=1)
        transport = JsonLineWriterBridgeTransport(process=dead)
        with self.assertRaisesRegex(
            WriterBridgeTransportError,
            "exited before dispatch",
        ):
            transport.exchange({}, None)
        self.assertEqual("", dead.stdin.getvalue())

        class TimeoutReader:
            def __init__(self, stream):
                pass

            def read(self, timeout_seconds):
                self.timeout_seconds = timeout_seconds
                raise TimeoutError

        process = FakeProcess()
        transport = JsonLineWriterBridgeTransport(
            process=process,
            clock=lambda: 10,
            reader_factory=TimeoutReader,
        )
        with self.assertRaisesRegex(
            WriterBridgeTransportError,
            "expired before dispatch",
        ):
            transport.exchange({}, 10)
        self.assertEqual("", process.stdin.getvalue())

    def test_close_sends_eof_once_and_waits_for_bridge_exit(self):
        process = FakeProcess()
        transport = JsonLineWriterBridgeTransport(
            process=process,
            close_timeout_seconds=1.5,
        )

        self.assertTrue(transport.close())
        self.assertTrue(transport.close())
        self.assertTrue(process.stdin.closed)
        self.assertEqual([1.5], process.wait_calls)


if __name__ == "__main__":
    unittest.main()
