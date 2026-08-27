"""Shared timing helpers for one bounded Action execution."""

import time


def bounded_timeout(timeout, deadline):
    """Return no more than ``timeout`` seconds without crossing ``deadline``."""
    if deadline is None:
        return timeout
    return max(0, min(timeout, deadline - time.monotonic()))
