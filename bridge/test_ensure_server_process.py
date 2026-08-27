import json
import os
from pathlib import Path
import socket
import subprocess
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = ROOT / "scripts"


def _unused_loopback_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


class EnsureServerProcessTests(unittest.TestCase):
    def test_concurrent_and_sequential_callers_converge_on_one_ready_bridge(self):
        port = _unused_loopback_port()
        worker = """
import json
import sys
import time
from pathlib import Path

scripts_dir, gate = sys.argv[1:]
sys.path.insert(0, scripts_dir)
while not Path(gate).exists():
    time.sleep(0.01)
import call
result = call._ensure_server()
print(json.dumps({
    "ok": result.ok,
    "disposition": result.disposition,
    "health": result.health,
}))
"""
        with tempfile.TemporaryDirectory() as tmp:
            gate = Path(tmp) / "go"
            environment = os.environ.copy()
            environment.update(
                {
                    "WPS_BRIDGE_PORT": str(port),
                    "WPS_BRIDGE_IDLE_SECONDS": "0",
                    "WPS_TRACE_DIR": tmp,
                }
            )
            callers = [
                subprocess.Popen(
                    [sys.executable, "-c", worker, str(SCRIPTS_DIR), str(gate)],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    env=environment,
                )
                for _ in range(2)
            ]
            gate.touch()
            results = []
            health = None
            try:
                for caller in callers:
                    stdout, stderr = caller.communicate(timeout=20)
                    self.assertEqual(0, caller.returncode, stderr)
                    results.append(json.loads(stdout))

                self.assertEqual([True, True], [result["ok"] for result in results])
                self.assertEqual(
                    {"self_started", "reused"},
                    {result["disposition"] for result in results},
                )
                instance_ids = {
                    result["health"]["instanceId"] for result in results
                }
                self.assertEqual(1, len(instance_ids))
                health = results[0]["health"]
                self.assertEqual(health["pid"], health["serverPid"])
                self.assertIn(health["launchedByPid"], {caller.pid for caller in callers})
                self.assertTrue(health["launchId"])
                self.assertEqual(str(ROOT), health["projectRoot"])
                self.assertEqual(str(ROOT / "bridge" / "server.py"), health["serverPath"])

                sequential = subprocess.run(
                    [sys.executable, "-c", worker, str(SCRIPTS_DIR), str(gate)],
                    capture_output=True,
                    text=True,
                    env=environment,
                    timeout=20,
                    check=False,
                )
                self.assertEqual(0, sequential.returncode, sequential.stderr)
                sequential_result = json.loads(sequential.stdout)
                self.assertTrue(sequential_result["ok"])
                self.assertEqual("reused", sequential_result["disposition"])
                self.assertEqual(
                    health["instanceId"],
                    sequential_result["health"]["instanceId"],
                )
            finally:
                if health is not None:
                    subprocess.run(
                        [sys.executable, str(SCRIPTS_DIR / "service.py"), "stop"],
                        capture_output=True,
                        text=True,
                        env=environment,
                        timeout=20,
                        check=False,
                    )


if __name__ == "__main__":
    unittest.main()
