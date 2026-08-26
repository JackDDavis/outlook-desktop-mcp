from types import SimpleNamespace

from outlook_desktop_mcp import server


class FakeBridge:
    def __init__(self):
        self.start_calls = []
        self.stopped = False

    def start(self, **kwargs):
        self.start_calls.append(kwargs)

    def stop(self):
        self.stopped = True


class FakeMcp:
    def __init__(self):
        self.settings = SimpleNamespace(host=None, port=None)
        self._tool_manager = SimpleNamespace(_tools={})
        self.transport = None

    def run(self, *, transport):
        self.transport = transport


def test_http_transport_starts_before_outlook_com_is_ready(monkeypatch):
    bridge = FakeBridge()
    mcp = FakeMcp()
    operation_manager = SimpleNamespace(shutdown=lambda: None)
    monkeypatch.setattr(server, "bridge", bridge)
    monkeypatch.setattr(server, "mcp", mcp)
    monkeypatch.setattr(server, "operation_manager", operation_manager)
    monkeypatch.setattr(
        "sys.argv",
        ["outlook-desktop-mcp", "--http", "--host", "127.0.0.1", "--port", "3721"],
    )

    server.main()

    assert bridge.start_calls == [
        {
            "wait_until_ready": False,
            "initialization_delay_seconds": 3.0,
        }
    ]
    assert mcp.transport == "sse"
    assert bridge.stopped is True
