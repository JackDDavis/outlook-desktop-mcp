import asyncio
import json

from outlook_desktop_mcp import server
from tests.fakes import Collection, FakeComError, FakeMailItem


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
        )


class MoveRaises(FakeMailItem):
    def Move(self, _destination):
        raise FakeComError(0x80020009)


def test_bulk_mark_refetches_once_and_echoes_fresh_identity(monkeypatch):
    stale = SaveFails("id-1", subject="Stale subject")
    fresh = FakeMailItem("id-1", subject="Fresh subject")
    namespace = SequenceNamespace([stale, fresh])
    monkeypatch.setattr(server, "bridge", FakeBridge(namespace))

    payload = json.loads(asyncio.run(server.bulk_mark_as_read(entry_ids="id-1")))
    row = payload["results"][0]

    assert row["status"] == "ok"
    assert row["id"] == "id-1"
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
    first = SaveFails("id-1")
    second = SaveFails("id-1")
    namespace = SequenceNamespace([first, second])
    monkeypatch.setattr(server, "bridge", FakeBridge(namespace))

    payload = json.loads(asyncio.run(server.bulk_mark_as_read(entry_ids="id-1")))
    row = payload["results"][0]

    assert row["status"] == "failed"
    assert row["error"]["code"] == "mapi_exception"
    assert payload["summary"]["failed"] == 1


def test_bulk_move_missing_source_is_idempotent_skip(monkeypatch):
    missing = FakeComError(0x8004010F)
    store = FakeStore()
    namespace = SequenceNamespace([missing], store=store)
    monkeypatch.setattr(server, "bridge", FakeBridge(namespace))

    payload = json.loads(asyncio.run(server.bulk_move_emails(
        "archive",
        entry_ids="old-id",
    )))
    row = payload["results"][0]

    assert row["status"] == "skipped"
    assert row["reason"] == "not_found_in_source"
    assert row["error"] is None
    assert payload["summary"]["skipped"] == 1


def test_bulk_move_returns_old_and_new_entry_ids(monkeypatch):
    item = MoveItem("old-id", new_entry_id="new-id")
    store = FakeStore()
    namespace = SequenceNamespace([item], store=store)
    monkeypatch.setattr(server, "bridge", FakeBridge(namespace))

    payload = json.loads(asyncio.run(server.bulk_move_emails(
        "archive",
        entry_ids="old-id",
    )))
    row = payload["results"][0]

    assert row["status"] == "ok"
    assert row["id"] == "old-id"
    assert row["entry_id"] == "old-id"
    assert row["new_entry_id"] == "new-id"


def test_bulk_move_reports_unconfirmed_when_old_id_disappears_after_error(
    monkeypatch,
):
    item = MoveRaises("old-id")
    missing = FakeComError(0x8004010F)
    store = FakeStore()
    namespace = SequenceNamespace([item, missing], store=store)
    monkeypatch.setattr(server, "bridge", FakeBridge(namespace))

    payload = json.loads(asyncio.run(server.bulk_move_emails(
        "archive",
        entry_ids="old-id",
    )))
    row = payload["results"][0]

    assert row["status"] == "skipped"
    assert row["reason"] == "moved_or_gone_unconfirmed"
    assert row["error"]["code"] == "not_found"
