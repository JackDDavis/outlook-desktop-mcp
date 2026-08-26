import asyncio
import sys
import threading
import time
from types import SimpleNamespace

import pytest

from outlook_desktop_mcp.com_bridge import BridgeTimeoutError, OutlookBridge
from outlook_desktop_mcp.utils.responses import tool_result


class FakeComError(Exception):
    def __init__(self, hresult):
        self.hresult = hresult
        super().__init__(hresult, "Server execution failed")


class FakeNamespace:
    DefaultStore = type("Store", (), {"DisplayName": "Test Store"})()
    CurrentUser = type("User", (), {"Name": "Test User"})()
    Stores = type("Stores", (), {"Count": 1})()


def start_fake_bridge(monkeypatch):
    bridge = OutlookBridge()
    monkeypatch.setattr(
        bridge,
        "_initialize_outlook",
        lambda: (object(), FakeNamespace()),
    )
    bridge.start(timeout=1)
    return bridge


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


def test_start_can_return_before_outlook_is_ready():
    bridge = OutlookBridge()
    release = threading.Event()
    bridge._com_thread_main = lambda: release.wait(1)

    started = time.monotonic()
    try:
        bridge.start(wait_until_ready=False)
        elapsed = time.monotonic() - started
    finally:
        release.set()
        bridge.stop()

    assert elapsed < 0.05


def test_nonblocking_start_can_delay_com_initialization():
    bridge = OutlookBridge()
    initialized = threading.Event()
    bridge._com_thread_main = initialized.set

    try:
        bridge.start(
            wait_until_ready=False,
            initialization_delay_seconds=0.05,
        )
        assert initialized.wait(0.01) is False
        assert initialized.wait(0.2) is True
    finally:
        bridge.stop()


def test_parallel_calls_are_serialized_with_queue_metrics(monkeypatch):
    bridge = start_fake_bridge(monkeypatch)

    def slow(_outlook, _namespace):
        time.sleep(0.05)
        return "slow"

    def fast(_outlook, _namespace):
        return "fast"

    async def run_calls():
        first = asyncio.create_task(
            bridge.call_with_metrics(slow, request_name="slow")
        )
        await asyncio.sleep(0.01)
        second = asyncio.create_task(
            bridge.call_with_metrics(fast, request_name="fast")
        )
        return await asyncio.gather(first, second)

    try:
        first, second = asyncio.run(run_calls())
    finally:
        bridge.stop()

    assert first.value == "slow"
    assert second.value == "fast"
    assert second.queue_wait_ms >= first.queue_wait_ms
    assert first.execution_ms >= 40


def test_health_snapshot_does_not_wait_for_active_request(monkeypatch):
    bridge = start_fake_bridge(monkeypatch)
    active = threading.Event()
    release = threading.Event()

    def blocked(_outlook, _namespace):
        active.set()
        release.wait(1)
        return "done"

    async def inspect_while_blocked():
        task = asyncio.create_task(
            bridge.call(blocked, timeout_seconds=1, request_name="blocked")
        )
        await asyncio.to_thread(active.wait, 1)
        started = time.monotonic()
        snapshot = bridge.health_snapshot()
        elapsed = time.monotonic() - started
        release.set()
        await task
        return snapshot, elapsed

    try:
        snapshot, elapsed = asyncio.run(inspect_while_blocked())
    finally:
        release.set()
        bridge.stop()

    assert elapsed < 0.05
    assert snapshot["active_request"]["name"] == "blocked"
    assert snapshot["thread_alive"] is True


def test_queued_timeout_cancels_request_before_side_effect(monkeypatch):
    bridge = start_fake_bridge(monkeypatch)
    active = threading.Event()
    release = threading.Event()
    side_effects = []

    def blocked(_outlook, _namespace):
        active.set()
        release.wait(1)

    def mutation(_outlook, _namespace):
        side_effects.append("ran")

    async def run_timeout():
        first = asyncio.create_task(
            bridge.call(blocked, timeout_seconds=1, request_name="blocked")
        )
        await asyncio.to_thread(active.wait, 1)
        with pytest.raises(BridgeTimeoutError, match="during queue"):
            await bridge.call(
                mutation,
                timeout_seconds=0.01,
                request_name="mutation",
            )
        release.set()
        await first
        await asyncio.sleep(0.05)

    try:
        asyncio.run(run_timeout())
    finally:
        release.set()
        bridge.stop()

    assert side_effects == []


def test_call_adds_native_queue_metadata(monkeypatch):
    bridge = start_fake_bridge(monkeypatch)

    def operation(_outlook, _namespace):
        return tool_result({"status": "ok"})

    try:
        result = asyncio.run(bridge.call(operation))
    finally:
        bridge.stop()

    assert result.structuredContent == {"status": "ok"}
    assert result.meta["queue_wait_ms"] >= 0
    assert result.meta["execution_ms"] >= 0


def test_idle_probe_is_rejected_without_queueing_behind_active_work(monkeypatch):
    bridge = start_fake_bridge(monkeypatch)
    active = threading.Event()
    release = threading.Event()
    probe_calls = []

    def blocked(_outlook, _namespace):
        active.set()
        release.wait(1)

    def probe(_outlook, _namespace):
        probe_calls.append("ran")

    async def run_probe():
        task = asyncio.create_task(
            bridge.call(blocked, timeout_seconds=1, request_name="blocked")
        )
        await asyncio.to_thread(active.wait, 1)
        result = await bridge.call_if_idle_with_metrics(
            probe,
            timeout_seconds=0.1,
            request_name="probe",
        )
        release.set()
        await task
        return result

    try:
        result = asyncio.run(run_probe())
    finally:
        release.set()
        bridge.stop()

    assert result is None
    assert probe_calls == []


def test_accounts_snapshot_is_copied(monkeypatch):
    bridge = start_fake_bridge(monkeypatch)
    accounts = [{"name": "one", "unread": 1}]

    try:
        bridge.update_accounts_snapshot(accounts)
        accounts[0]["unread"] = 99
        snapshot = bridge.health_snapshot()
    finally:
        bridge.stop()

    assert snapshot["accounts"] == [{"name": "one", "unread": 1}]
    assert snapshot["accounts_snapshot_age_ms"] >= 0


def test_invalid_timeout_environment(monkeypatch):
    monkeypatch.setenv("MCP_SINGLE_CALL_TIMEOUT_SECONDS", "0")

    with pytest.raises(ValueError, match="positive"):
        OutlookBridge()


def test_startup_timeout_environment_is_configurable(monkeypatch):
    monkeypatch.setenv("MCP_COM_STARTUP_TIMEOUT_SECONDS", "120")

    bridge = OutlookBridge()

    assert bridge.startup_timeout_seconds == 120


def test_timeout_on_stopped_bridge_does_not_leave_pending_request():
    bridge = OutlookBridge()

    async def run_call():
        with pytest.raises(BridgeTimeoutError, match="during queue"):
            await bridge.call(
                lambda _outlook, _namespace: None,
                timeout_seconds=0.01,
            )

    asyncio.run(run_call())

    assert bridge.health_snapshot()["queue_depth"] == 0


def test_call_fails_immediately_after_bridge_initialization_error():
    bridge = OutlookBridge()
    bridge._thread = threading.Thread()
    bridge._init_error = RuntimeError("COM initialization failed")

    started = time.monotonic()
    with pytest.raises(RuntimeError, match="COM initialization failed"):
        asyncio.run(bridge.call(lambda _outlook, _namespace: None))

    assert time.monotonic() - started < 0.05
    assert bridge.health_snapshot()["queue_depth"] == 0


def test_initialization_retries_transient_server_execution_failure(monkeypatch):
    bridge = OutlookBridge()
    attempts = []
    transient = FakeComError(0x80080005)

    def initialize():
        attempts.append("attempt")
        if len(attempts) == 1:
            raise transient
        return object(), FakeNamespace()

    monkeypatch.setattr(bridge, "_initialize_outlook", initialize)
    monkeypatch.setattr(bridge, "initialization_retry_seconds", 0.01)

    outlook, namespace = bridge._initialize_outlook_with_retry()

    assert outlook is not None
    assert namespace is not None
    assert attempts == ["attempt", "attempt"]
    assert bridge.health_snapshot()["last_failure"]["hresult"] == "0x80080005"


def test_initialization_only_attaches_to_running_outlook(monkeypatch):
    bridge = OutlookBridge()
    namespace = FakeNamespace()
    outlook = SimpleNamespace(GetNamespace=lambda _name: namespace)
    client = SimpleNamespace(
        GetActiveObject=lambda _name: outlook,
        Dispatch=lambda _name: pytest.fail("must not launch Outlook"),
    )
    monkeypatch.setitem(
        sys.modules,
        "win32com",
        SimpleNamespace(client=client),
    )
    monkeypatch.setitem(sys.modules, "win32com.client", client)

    connected_outlook, connected_namespace = bridge._initialize_outlook()

    assert connected_outlook is outlook
    assert connected_namespace is namespace
