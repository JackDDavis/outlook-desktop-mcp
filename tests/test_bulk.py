import asyncio
import json
import time

from outlook_desktop_mcp import server
from outlook_desktop_mcp.operations import OperationManager
from outlook_desktop_mcp.utils.formatting import PR_INTERNET_MESSAGE_ID_UNICODE
from tests.fakes import Collection, FakeComError, FakeMailItem, make_entry_id


class FakeBridge:
    def __init__(self, namespace):
        self.namespace = namespace

    async def call(self, function, *args, **kwargs):
        return function(object(), self.namespace, *args, **kwargs)


class FakeStore:
    DisplayName = "Store"
    StoreID = "store-1"

    def __init__(self):
        self.archive = type(
            "Folder",
            (),
            {"Name": "Archive", "Folders": Collection()},
        )()
        self.root = type(
            "Root",
            (),
            {"Folders": Collection([self.archive])},
        )()

    def GetDefaultFolder(self, _folder_type):
        return self.archive

    def GetRootFolder(self):
        return self.root


class SequenceNamespace:
    def __init__(self, sequence, store=None):
        self.sequence = list(sequence)
        self.calls = 0
        self.DefaultStore = store
        self.Stores = Collection([store] if store else [])
        self.Accounts = Collection()

    def GetItemFromID(self, _entry_id, _store_id=None):
        value = self.sequence[min(self.calls, len(self.sequence) - 1)]
        self.calls += 1
        if isinstance(value, Exception):
            raise value
        return value


class MappingNamespace(SequenceNamespace):
    def __init__(self, items, store=None):
        super().__init__(items, store=store)
        self.items_by_id = {item.EntryID: item for item in items}

    def GetItemFromID(self, entry_id, _store_id=None):
        self.calls += 1
        return self.items_by_id[entry_id]


class SaveFails(FakeMailItem):
    def Save(self):
        raise FakeComError(0x80020009)


class MoveItem(FakeMailItem):
    def __init__(self, *args, new_entry_id="moved", **kwargs):
        super().__init__(*args, **kwargs)
        self.new_entry_id = new_entry_id

    def Move(self, _destination):
        return FakeMailItem(
            self.new_entry_id,
            subject=self.Subject,
            received_time=self.ReceivedTime,
            properties=self.PropertyAccessor.properties,
        )


class MoveRaises(FakeMailItem):
    def Move(self, _destination):
        raise FakeComError(0x80020009)


class BatchingBridge(FakeBridge):
    bulk_timeout_seconds = 5

    def __init__(self, namespace):
        super().__init__(namespace)
        self.batches = []

    async def call(self, function, *args, **kwargs):
        request_name = kwargs.pop("request_name", "")
        kwargs.pop("timeout_seconds", None)
        if request_name.endswith("_batch"):
            self.batches.append(list(args[0]))
        return function(object(), self.namespace, *args, **kwargs)


class SlowMailItem(FakeMailItem):
    def Save(self):
        time.sleep(0.005)
        super().Save()


class SortableCollection(Collection):
    def Sort(self, *_args):
        return None

    def Restrict(self, *_args):
        return self


class FilterStore:
    DisplayName = "Store"
    StoreID = "store-1"

    def __init__(self, items):
        self.inbox = type(
            "Folder",
            (),
            {
                "Name": "Inbox",
                "Folders": Collection(),
                "Items": SortableCollection(items),
            },
        )()
        self.root = type(
            "Root",
            (),
            {"Folders": Collection([self.inbox])},
        )()

    def GetDefaultFolder(self, _folder_type):
        return self.inbox

    def GetRootFolder(self):
        return self.root


def poll_operation(operation_id, timeout=5):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        payload = json.loads(asyncio.run(
            server.outlook_operation_status(operation_id)
        ))
        if payload["status"] != "in_progress":
            return payload
        time.sleep(0.01)
    raise AssertionError("operation did not finish")


def test_bulk_mark_refetches_once_and_echoes_fresh_identity(monkeypatch):
    entry_id = make_entry_id()
    message_id = "<bulk-refetch@example.com>"
    properties = {PR_INTERNET_MESSAGE_ID_UNICODE: message_id}
    stale = SaveFails(entry_id, subject="Stale subject", properties=properties)
    fresh = FakeMailItem(entry_id, subject="Fresh subject", properties=properties)
    namespace = SequenceNamespace([stale, fresh])
    monkeypatch.setattr(server, "bridge", FakeBridge(namespace))

    payload = json.loads(asyncio.run(server.bulk_mark_as_read(entry_ids=entry_id)))
    row = payload["results"][0]

    assert row["status"] == "ok"
    assert row["id"] == entry_id
    assert row["message_id"] == message_id
    assert row["id_stable"] is True
    assert row["subject"] == "Fresh subject"
    assert row["received_time"]
    assert row["action"] == "marked_as_read"
    assert row["error"] is None
    assert payload["summary"] == {
        "total": 1,
        "ok": 1,
        "failed": 0,
        "skipped": 0,
    }
    assert namespace.calls == 2


def test_bulk_mark_reports_structured_error_after_retry(monkeypatch):
    entry_id = make_entry_id()
    first = SaveFails(entry_id)
    second = SaveFails(entry_id)
    namespace = SequenceNamespace([first, second])
    monkeypatch.setattr(server, "bridge", FakeBridge(namespace))

    payload = json.loads(asyncio.run(server.bulk_mark_as_read(entry_ids=entry_id)))
    row = payload["results"][0]

    assert row["status"] == "failed"
    assert row["message_id"] is None
    assert row["id_stable"] is False
    assert row["error"]["code"] == "mapi_exception"
    assert payload["summary"]["failed"] == 1


def test_bulk_move_missing_source_is_idempotent_skip(monkeypatch):
    missing = FakeComError(0x8004010F)
    store = FakeStore()
    namespace = SequenceNamespace([missing], store=store)
    monkeypatch.setattr(server, "bridge", FakeBridge(namespace))

    payload = json.loads(asyncio.run(server.bulk_move_emails(
        "archive",
        entry_ids=make_entry_id(),
    )))
    row = payload["results"][0]

    assert row["status"] == "skipped"
    assert row["reason"] == "not_found_in_source"
    assert row["error"] is None
    assert payload["summary"]["skipped"] == 1


def test_bulk_move_returns_old_and_new_entry_ids(monkeypatch):
    old_entry_id = make_entry_id(1)
    new_entry_id = make_entry_id(2)
    message_id = "<bulk-move@example.com>"
    item = MoveItem(
        old_entry_id,
        new_entry_id=new_entry_id,
        properties={PR_INTERNET_MESSAGE_ID_UNICODE: message_id},
    )
    store = FakeStore()
    namespace = SequenceNamespace([item], store=store)
    monkeypatch.setattr(server, "bridge", FakeBridge(namespace))

    payload = json.loads(asyncio.run(server.bulk_move_emails(
        "archive",
        entry_ids=old_entry_id,
    )))
    row = payload["results"][0]

    assert row["status"] == "ok"
    assert row["id"] == old_entry_id
    assert row["entry_id"] == old_entry_id
    assert row["old_entry_id"] == old_entry_id
    assert row["new_entry_id"] == new_entry_id
    assert row["message_id"] == message_id
    assert row["id_stable"] is True


def test_bulk_move_reports_unconfirmed_when_old_id_disappears_after_error(
    monkeypatch,
):
    old_entry_id = make_entry_id()
    item = MoveRaises(old_entry_id)
    missing = FakeComError(0x8004010F)
    store = FakeStore()
    namespace = SequenceNamespace([item, item, missing], store=store)
    monkeypatch.setattr(server, "bridge", FakeBridge(namespace))

    payload = json.loads(asyncio.run(server.bulk_move_emails(
        "archive",
        entry_ids=old_entry_id,
    )))
    row = payload["results"][0]

    assert row["status"] == "skipped"
    assert row["reason"] == "moved_or_gone_unconfirmed"
    assert row["error"]["code"] == "not_found"


def test_short_budget_40_item_operation_polls_to_complete_in_ordered_batches(
    monkeypatch,
    tmp_path,
):
    items = [
        SlowMailItem(
            make_entry_id(index + 1),
            properties={
                PR_INTERNET_MESSAGE_ID_UNICODE: (
                    f"<chunk-{index:02}@example.com>"
                )
            },
        )
        for index in range(40)
    ]
    namespace = MappingNamespace(items)
    bridge = BatchingBridge(namespace)
    monkeypatch.setattr(server, "bridge", bridge)
    monkeypatch.setattr(
        server,
        "operation_manager",
        OperationManager(tmp_path, process_instance_id="test-process"),
    )
    monkeypatch.setenv("MCP_OP_BUDGET_SECONDS", "0.01")

    initial = json.loads(asyncio.run(server.bulk_mark_as_read(
        entry_ids=",".join(item.EntryID for item in items),
        count=40,
    )))

    assert initial["status"] == "in_progress"
    assert initial["remaining"] <= 40
    completed = poll_operation(initial["operation_id"])
    assert completed["status"] == "complete"
    assert completed["remaining"] == 0
    assert completed["summary"]["ok"] == 40
    assert [row["id"] for row in completed["results"]] == [
        item.EntryID for item in items
    ]
    assert [row["message_id"] for row in completed["results"]] == [
        f"<chunk-{index:02}@example.com>" for index in range(40)
    ]
    assert all(row["id_stable"] is True for row in completed["results"])
    assert bridge.batches == [
        [make_entry_id(index + 1) for index in range(offset, offset + 10)]
        for offset in range(0, 40, 10)
    ]


def test_filter_selected_work_over_ten_uses_ordered_sub_batches(
    monkeypatch,
    tmp_path,
):
    items = [FakeMailItem(make_entry_id(index + 1)) for index in range(25)]
    store = FilterStore(items)
    namespace = SequenceNamespace(items, store=store)
    bridge = BatchingBridge(namespace)
    monkeypatch.setattr(server, "bridge", bridge)
    monkeypatch.setattr(
        server,
        "operation_manager",
        OperationManager(tmp_path, process_instance_id="test-process"),
    )

    payload = json.loads(asyncio.run(server.bulk_mark_as_read(
        subject_contains="Subject",
        count=25,
    )))

    assert payload["summary"]["ok"] == 25
    assert bridge.batches == [
        [make_entry_id(index + 1) for index in range(0, 10)],
        [make_entry_id(index + 1) for index in range(10, 20)],
        [make_entry_id(index + 1) for index in range(20, 25)],
    ]


def test_explicit_identifiers_are_not_silently_capped_by_count(
    monkeypatch,
    tmp_path,
):
    items = [
        FakeMailItem(make_entry_id(index + 1))
        for index in range(60)
    ]
    monkeypatch.setattr(
        server,
        "bridge",
        BatchingBridge(MappingNamespace(items)),
    )
    monkeypatch.setattr(
        server,
        "operation_manager",
        OperationManager(tmp_path, process_instance_id="all-explicit-ids"),
    )

    payload = json.loads(asyncio.run(server.bulk_mark_as_read(
        entry_ids=",".join(item.EntryID for item in items),
        count=10,
    )))

    assert payload["summary"]["total"] == 60
    assert payload["summary"]["ok"] == 60
    assert [row["id"] for row in payload["results"]] == [
        item.EntryID for item in items
    ]
