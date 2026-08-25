import asyncio
import json
import time

import pytest

from outlook_desktop_mcp import server
from outlook_desktop_mcp.com_bridge import OutlookBridge
from outlook_desktop_mcp.operations import (
    INTERRUPTED_GUIDANCE,
    OperationManager,
)
from outlook_desktop_mcp.utils.errors import error_details
from outlook_desktop_mcp.utils.formatting import (
    PR_INTERNET_MESSAGE_ID_UNICODE,
)
from tests.fakes import Collection, FakeComError, FakeMailItem, make_entry_id
from tests.test_bulk import (
    BatchingBridge,
    FakeStore,
    MappingNamespace,
    SequenceNamespace,
    SlowMailItem,
    poll_operation,
)
from tests.test_calendar_safety import fixture_environment
from tests.test_message_identity import (
    IdentityFolder,
    MovableMail,
    install,
)
from tests.test_search import install_search_environment
from tests.test_status import BusyStatusBridge, base_snapshot

pytestmark = pytest.mark.acceptance


def _payload(value):
    if hasattr(value, "content"):
        value = value.content[0].text
    return json.loads(value)


@pytest.mark.reliability_criterion("criterion_1")
def test_bounded_36_item_bulk_operation_polls_to_complete(monkeypatch, tmp_path):
    items = [
        SlowMailItem(
            make_entry_id(index + 1),
            properties={
                PR_INTERNET_MESSAGE_ID_UNICODE: (
                    f"<acceptance-{index:02}@example.com>"
                )
            },
        )
        for index in range(36)
    ]
    bridge = BatchingBridge(MappingNamespace(items))
    monkeypatch.setattr(server, "bridge", bridge)
    monkeypatch.setattr(
        server,
        "operation_manager",
        OperationManager(tmp_path, process_instance_id="acceptance-bulk"),
    )
    monkeypatch.setenv("MCP_OP_BUDGET_SECONDS", "0.01")

    initial = _payload(asyncio.run(server.bulk_mark_as_read(
        entry_ids=",".join(item.EntryID for item in items),
        count=36,
    )))

    assert initial["status"] == "in_progress"
    assert initial["status"] != "not_found"
    completed = poll_operation(initial["operation_id"])
    assert completed["status"] == "complete"
    assert completed["remaining"] == 0
    assert completed["summary"] == {
        "total": 36,
        "ok": 36,
        "failed": 0,
        "skipped": 0,
    }
    assert len(completed["results"]) == 36
    assert [row["id"] for row in completed["results"]] == [
        item.EntryID for item in items
    ]
    assert all(row["status"] == "ok" for row in completed["results"])


@pytest.mark.reliability_criterion("criterion_2")
def test_message_id_survives_move_resolution_and_degrades_explicitly(
    monkeypatch,
):
    message_id = "<acceptance-move@example.com>"
    old_entry_id = make_entry_id(1)
    new_entry_id = make_entry_id(2)
    moved = MovableMail(
        old_entry_id,
        moved_entry_id=new_entry_id,
        properties={PR_INTERNET_MESSAGE_ID_UNICODE: message_id},
    )
    archive = IdentityFolder("Archive")
    install(monkeypatch, [moved], archive=archive)

    move_payload = _payload(asyncio.run(server.move_email(
        message_id,
        target_folder="archive",
        folder="inbox",
    )))
    resolved = _payload(asyncio.run(server.read_email(
        entry_id=message_id,
        folder="archive",
    )))

    assert move_payload["old_entry_id"] == old_entry_id
    assert move_payload["new_entry_id"] == new_entry_id
    assert move_payload["message_id"] == message_id
    assert move_payload["id_stable"] is True
    assert resolved["entry_id"] == new_entry_id
    assert resolved["message_id"] == message_id
    assert resolved["id_stable"] is True

    stable = FakeMailItem(
        make_entry_id(3),
        properties={PR_INTERNET_MESSAGE_ID_UNICODE: "<shape@example.com>"},
    )
    absent = FakeMailItem(make_entry_id(4))
    install(monkeypatch, [stable, absent])
    listed = _payload(asyncio.run(server.list_emails(count=2)))
    assert listed[0]["id_stable"] is True
    assert listed[0]["message_id"] == "<shape@example.com>"
    assert listed[1]["id_stable"] is False
    assert listed[1]["message_id"] is None


class HungStatusBridge:
    def __init__(self):
        self.snapshot = base_snapshot()
        self.snapshot.update({
            "thread_alive": True,
            "last_success_at": None,
            "last_success_age_ms": None,
        })

    async def call_if_idle_with_metrics(self, *_args, **kwargs):
        assert kwargs["timeout_seconds"] == 5.0
        await asyncio.sleep(5.05)
        raise TimeoutError("synthetic hung bridge")

    def health_snapshot(self):
        return self.snapshot


@pytest.mark.reliability_criterion("criterion_3")
def test_status_degrades_within_six_seconds_for_stopped_hung_and_busy(
    monkeypatch,
):
    monkeypatch.setattr(server, "_outlook_process_running", lambda: False)

    stopped = OutlookBridge()
    monkeypatch.setattr(server, "bridge", stopped)
    started = time.monotonic()
    stopped_payload = asyncio.run(server.outlook_status()).structuredContent
    assert time.monotonic() - started < 6
    assert stopped_payload["com_responsive"] is False
    assert stopped_payload["com_state"] == "bridge_stopped"

    monkeypatch.setattr(server, "bridge", HungStatusBridge())
    started = time.monotonic()
    hung_payload = asyncio.run(server.outlook_status()).structuredContent
    assert time.monotonic() - started < 6
    assert hung_payload["com_responsive"] is False
    assert hung_payload["com_state"] == "unresponsive"
    assert hung_payload["com_probe"] == "failed"
    assert hung_payload["probe_error"]["code"] == "timeout"

    busy = BusyStatusBridge()
    busy.snapshot["last_success_age_ms"] = 90_000
    monkeypatch.setattr(server, "bridge", busy)
    started = time.monotonic()
    busy_payload = asyncio.run(server.outlook_status()).structuredContent
    assert time.monotonic() - started < 6
    assert busy_payload["com_responsive"] is False
    assert busy_payload["com_state"] == "busy"
    assert busy_payload["com_probe"] == "skipped_busy"
    assert busy_payload["active_request"]["elapsed_ms"] == 90_000
    assert busy_payload["active_request"]["name"] == "bulk_move"


@pytest.mark.reliability_criterion("criterion_4")
def test_representative_tool_errors_are_actionable(monkeypatch):
    representative_errors = [
        FakeComError(-2147023174),
        FakeComError(-2147352567),
        FakeComError(0x8004010F),
        FakeComError(-1, "sanitized detail"),
        TimeoutError("timed out"),
        ValueError("invalid"),
        LookupError("missing"),
        RuntimeError("unexpected"),
    ]
    for error in representative_errors:
        details = error_details(error)
        assert details["meaning"]
        assert details["suggested_action"]
        assert isinstance(details["retryable"], bool)

    fixture_environment(monkeypatch)
    result = asyncio.run(server.mcp._tool_manager.call_tool(
        "create_event",
        {
            "subject": "Invalid",
            "start": "2026-01-01 10:00",
            "end": "2026-01-01 09:00",
        },
        convert_result=True,
    ))
    assert result.isError is True
    assert result.structuredContent["error"]["meaning"]
    assert result.structuredContent["error"]["suggested_action"]
    assert result.structuredContent["error"]["retryable"] is False


class SlowListItems(Collection):
    def Sort(self, *_args):
        time.sleep(0.08)

    def Restrict(self, *_args):
        return self


class ListStore:
    DisplayName = "Acceptance Store"
    StoreID = "acceptance-store"

    def __init__(self):
        self.inbox = type(
            "Folder",
            (),
            {
                "Name": "Inbox",
                "Items": SlowListItems([FakeMailItem(make_entry_id())]),
                "Folders": Collection(),
            },
        )()

    def GetDefaultFolder(self, _folder_type):
        return self.inbox


class ListNamespace:
    CurrentUser = type("User", (), {"Name": "Acceptance User"})()

    def __init__(self):
        self.DefaultStore = ListStore()
        self.Stores = Collection([self.DefaultStore])
        self.Accounts = Collection()


@pytest.mark.reliability_criterion("criterion_5")
def test_parallel_list_calls_serialize_and_expose_queue_wait(monkeypatch):
    bridge = OutlookBridge()
    monkeypatch.setattr(
        bridge,
        "_initialize_outlook",
        lambda: (object(), ListNamespace()),
    )
    bridge.start(timeout=1)
    monkeypatch.setattr(server, "bridge", bridge)

    async def run_lists():
        return await asyncio.gather(
            server.list_emails(count=1),
            server.list_emails(count=1),
        )

    started = time.monotonic()
    try:
        first, second = asyncio.run(run_lists())
    finally:
        bridge.stop()
    elapsed = time.monotonic() - started

    assert elapsed >= 0.14
    assert _payload(first)[0]["entry_id"] == make_entry_id()
    assert _payload(second)[0]["entry_id"] == make_entry_id()
    waits = [first.meta["queue_wait_ms"], second.meta["queue_wait_ms"]]
    assert min(waits) >= 0
    assert max(waits) >= 50


@pytest.mark.reliability_criterion("criterion_6")
def test_calendar_0830_echoes_local_time_and_timezone(monkeypatch):
    fixture_environment(monkeypatch)

    payload = _payload(asyncio.run(server.create_event(
        "Acceptance local time",
        "2026-01-01 08:30",
        "2026-01-01 09:30",
    )))

    assert payload["start_local"] == "2026-01-01 08:30"
    assert payload["end_local"] == "2026-01-01 09:30"
    assert payload["timezone"]
    assert payload["interpreted_as"] == "local"


@pytest.mark.reliability_criterion("criterion_7")
def test_sender_modes_cover_exchange_and_capped_imap(monkeypatch):
    exchange_items = install_search_environment(
        monkeypatch,
        [FakeMailItem("exchange", subject="Acceptance")],
        account_type=0,
    )
    exchange = _payload(asyncio.run(server.search_emails(
        "Acceptance",
        sender="sender@example.com",
    )))
    assert exchange["filter_mode"] == "dasl"
    assert exchange["truncated"] is False
    assert "fromemail" in exchange_items.last_filter

    install_search_environment(
        monkeypatch,
        [
            FakeMailItem(str(index), subject="Acceptance")
            for index in range(1001)
        ],
        account_type=1,
    )
    imap = _payload(asyncio.run(server.search_emails(
        "Acceptance",
        sender="sender@example.com",
        count=2,
    )))
    assert imap["filter_mode"] == "client"
    assert imap["truncated"] is True
    assert len(imap["results"]) == 2


@pytest.mark.reliability_criterion("criterion_8")
def test_rerun_move_skips_missing_and_restart_interrupts_snapshot(
    monkeypatch,
    tmp_path,
):
    store = FakeStore()
    monkeypatch.setattr(
        server,
        "bridge",
        BatchingBridge(
            SequenceNamespace([FakeComError(0x8004010F)], store=store)
        ),
    )
    monkeypatch.setattr(
        server,
        "operation_manager",
        OperationManager(tmp_path / "move", process_instance_id="move"),
    )
    rerun = _payload(asyncio.run(server.bulk_move_emails(
        "archive",
        entry_ids=make_entry_id(),
    )))
    assert rerun["results"][0]["status"] == "skipped"
    assert rerun["results"][0]["reason"] == "not_found_in_source"

    first = OperationManager(
        tmp_path / "restart",
        process_instance_id="before-restart",
    )
    operation_id = first.create("bulk_move_emails", remaining=36)
    proven_rows = [{"id": "one", "status": "ok"}]
    first.append_results(
        operation_id,
        proven_rows,
        matched_count=36,
        remaining=35,
    )
    restarted = OperationManager(
        tmp_path / "restart",
        process_instance_id="after-restart",
    )
    interrupted = restarted.get(operation_id)
    assert interrupted["status"] == "interrupted"
    assert interrupted["results"] == proven_rows
    assert interrupted["remaining"] == 35
    assert interrupted["guidance"] == INTERRUPTED_GUIDANCE
