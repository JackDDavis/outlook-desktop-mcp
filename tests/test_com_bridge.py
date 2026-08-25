import threading
import time

import pytest

from outlook_desktop_mcp.com_bridge import OutlookBridge


def test_start_waits_for_configured_timeout():
    bridge = OutlookBridge()

    def delayed_ready():
        time.sleep(0.05)
        bridge._ready.set()

    bridge._com_thread_main = delayed_ready
    bridge.start(timeout=1)
    bridge._thread.join(timeout=1)


def test_start_reports_configured_timeout():
    bridge = OutlookBridge()
    bridge._com_thread_main = lambda: threading.Event().wait(1)
    with pytest.raises(RuntimeError, match="within 0.01s"):
        bridge.start(timeout=0.01)
