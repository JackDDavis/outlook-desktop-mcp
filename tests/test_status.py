import asyncio
import time

from outlook_desktop_mcp import server
from outlook_desktop_mcp.com_bridge import BridgeCallResult, OutlookBridge


def base_snapshot():
    return {
        "process_instance_id": "test-process",
        "thread_alive": True,
        "queue_depth": 0,
        "active_request": None,
        "last_success_at": "2026-01-01T00:00:00+00:00",
        "last_success_age_ms": 100,
        "last_failure_at": None,
        "last_failure": None,
        "accounts": [],
        "accounts_snapshot_at": None,
        "accounts_snapshot_age_ms": None,
    }


class BusyStatusBridge:
    def __init__(self):
        self.snapshot = base_snapshot()
        self.snapshot["queue_depth"] = 1
        self.snapshot["active_request"] = {
            "name": "bulk_move",
            "started_at": "2026-01-01T00:00:00+00:00",
            "elapsed_ms": 90_000,
            "caller_timed_out": False,
        }
        self.probe_calls = 0

    async def call_if_idle_with_metrics(self, *_args, **_kwargs):
        self.probe_calls += 1
        return None

    def health_snapshot(self):
        return self.snapshot


class IdleStatusBridge:
    def __init__(self):
        self.snapshot = base_snapshot()

    async def call_if_idle_with_metrics(self, *_args, **_kwargs):
        return BridgeCallResult(
            value={
                "application_name": "Microsoft Outlook",
                "profile": "Test Profile",
                "accounts": [{"name": "Store", "unread": 2, "total": 10}],
            },
            queue_wait_ms=0,
            execution_ms=12,
            request_name="outlook_status_probe",
        )

    def update_accounts_snapshot(self, accounts):
        self.snapshot["accounts"] = accounts
        self.snapshot["accounts_snapshot_at"] = "now"
        self.snapshot["accounts_snapshot_age_ms"] = 0

    def health_snapshot(self):
        return self.snapshot


def test_status_returns_immediately_without_queueing_behind_busy_work(monkeypatch):
    bridge = BusyStatusBridge()
    monkeypatch.setattr(server, "bridge", bridge)
    monkeypatch.setattr(server, "_outlook_process_running", lambda: True)
    monkeypatch.setattr(
        server.operation_manager,
        "in_flight",
        lambda: 2,
    )

    started = time.monotonic()
    result = asyncio.run(server.outlook_status())
    elapsed = time.monotonic() - started
    payload = result.structuredContent

    assert elapsed < 0.1
    assert payload["com_probe"] == "skipped_busy"
    assert payload["com_state"] == "busy"
    assert payload["active_request"]["name"] == "bulk_move"
    assert payload["operations_in_flight"] == 2
    assert bridge.probe_calls == 1


def test_status_refreshes_accounts_when_idle(monkeypatch):
    monkeypatch.setattr(server, "bridge", IdleStatusBridge())
    monkeypatch.setattr(server, "_outlook_process_running", lambda: True)

    result = asyncio.run(server.outlook_status())
    payload = result.structuredContent

    assert payload["outlook_running"] is True
    assert payload["com_responsive"] is True
    assert payload["com_probe"] == "completed"
    assert payload["com_ping_ms"] == 12
    assert payload["profile"] == "Test Profile"
    assert payload["accounts"][0]["unread"] == 2


def test_status_does_not_probe_or_wait_when_bridge_is_stopped(monkeypatch):
    bridge = OutlookBridge()
    monkeypatch.setattr(server, "bridge", bridge)
    monkeypatch.setattr(server, "_outlook_process_running", lambda: False)

    started = time.monotonic()
    result = asyncio.run(server.outlook_status())
    elapsed = time.monotonic() - started
    payload = result.structuredContent

    assert elapsed < 0.1
    assert payload["com_probe"] == "skipped_bridge_stopped"
    assert payload["com_state"] == "bridge_stopped"
    assert payload["queue_depth"] == 0
