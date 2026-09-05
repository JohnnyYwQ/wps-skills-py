"""Word Skill client assembly; discovery and execution share production contracts."""

import os
from pathlib import Path
import sys

from wps_skills.client.session_client import ActionFailed, SessionClient, SessionClientError
from wps_skills.cli.call import main


def open_session(*, timeout=60, debug_close_created_document=False):
    """Return a context-managed Windows Word Session Client, not yet started."""
    if not isinstance(debug_close_created_document, bool):
        raise TypeError("debug_close_created_document must be Boolean")
    source = Path(__file__).resolve().parents[2]
    env = dict(os.environ)
    env["PYTHONPATH"] = str(source)
    env["PYTHONIOENCODING"] = "utf-8"
    command = [sys.executable, "-m", "wps_skills.cli.call", "--session", "--app", "word"]
    if debug_close_created_document:
        command.append("--debug-close-created-document")
    return SessionClient(command, application="word", env=env, timeout=timeout)
