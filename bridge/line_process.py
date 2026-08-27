"""Bounded cleanup for persistent line-RPC child processes."""

import subprocess


GRACEFUL_PROCESS_TIMEOUT_SECONDS = 2.0
FORCED_PROCESS_TIMEOUT_SECONDS = 1.0


def stop_line_process(
    process,
    graceful_timeout=GRACEFUL_PROCESS_TIMEOUT_SECONDS,
    forced_timeout=FORCED_PROCESS_TIMEOUT_SECONDS,
):
    """Stop a line-RPC child, preferring its protocol-level ``EXIT`` command."""
    if process is None:
        return

    try:
        running = process.poll() is None
    except Exception:
        running = True
    if not running:
        return

    try:
        if process.stdin:
            process.stdin.write("EXIT\n")
            process.stdin.flush()
    except Exception:
        pass

    try:
        process.wait(timeout=graceful_timeout)
        return
    except subprocess.TimeoutExpired:
        pass
    except Exception:
        pass

    try:
        process.terminate()
        process.wait(timeout=forced_timeout)
        return
    except subprocess.TimeoutExpired:
        pass
    except Exception:
        pass

    try:
        process.kill()
        process.wait(timeout=forced_timeout)
    except Exception:
        pass
