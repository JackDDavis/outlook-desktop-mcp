import asyncio
import json
from datetime import datetime

import pytest
from mcp.server.fastmcp.exceptions import ToolError

from outlook_desktop_mcp import server


class Collection:
    def __init__(self, values=()):
        self.values = list(values)

    @property
    def Count(self):
        return len(self.values)

    def Item(self, index):
        return self.values[index - 1]

    def __iter__(self):
        return iter(self.values)


class FakeItems:
    def __init__(self, owner):
        self.owner = owner
        self.created = []

    @property
    def Count(self):
        return len(self.created)

    def Add(self, message_class):
        assert message_class == "IPM.Appointment"
        item = FakeItem(parent=self.owner)
        self.created.append(item)
        return item


class FakeFolder:
    def __init__(self, name, entry_id, *, path=None, children=()):
        self.Name = name
        self.EntryID = entry_id
        self.FolderPath = path or name
        self.DefaultItemType = server.OL_APPOINTMENT_ITEM
        self.Items = FakeItems(self)
        self.Folders = Collection(children)


class FakeStore:
    def __init__(self, display_name, store_id, default_calendar, root):
        self.DisplayName = display_name
        self.StoreID = store_id
        self.default_calendar = default_calendar
        self.root = root

    def GetDefaultFolder(self, folder_type):
        assert folder_type == server.OL_FOLDER_CALENDAR
        return self.default_calendar

    def GetRootFolder(self):
        return self.root


class FakeItem:
    Class = server._OL_CLASS_APPOINTMENT

    def __init__(self, entry_id="event-1", parent=None):
        self.EntryID = entry_id
        self.Parent = parent
        self.Subject = ""
        self.Start = datetime(2026, 1, 1, 9)
        self.End = datetime(2026, 1, 1, 10)
        self.Location = ""
        self.Body = ""
        self.AllDayEvent = False
        self.ReminderSet = False
        self.ReminderMinutesBeforeStart = 0
        self.MeetingStatus = 0
        self.saved = False

    def Save(self):
        self.saved = True

    def Move(self, target):
        self.Parent = target
        self.EntryID = f"{self.EntryID}-moved"
        return self


class FakeAccount:
    def __init__(self, store, address):
        self.DeliveryStore = store
        self.SmtpAddress = address
        self.DisplayName = address


class FakeNamespace:
    def __init__(self, stores, default_store, accounts=()):
        self.Stores = Collection(stores)
        self.DefaultStore = default_store
        self.Accounts = Collection(accounts)
        self.items = {}

    def GetItemFromID(self, entry_id, _store_id=None):
        return self.items[entry_id]


class FakeOutlook:
    def __init__(self, accounts=()):
        self.Session = type("Session", (), {"Accounts": Collection(accounts)})()
        self.create_item_calls = 0

    def CreateItem(self, _item_type):
        self.create_item_calls += 1
        return FakeItem()


class FakeBridge:
    def __init__(self, outlook, namespace):
        self.outlook = outlook
        self.namespace = namespace

    async def call(self, function, *args, **kwargs):
        return function(self.outlook, self.namespace, *args, **kwargs)


def fixture_environment(monkeypatch):
    primary = FakeFolder("Calendar", "primary-calendar")
    secondary = FakeFolder(
        "Projects", "projects-calendar", path="Calendar/Projects"
    )
    primary_root = FakeFolder(
        "Root", "primary-root", children=(primary, secondary)
    )
    primary_store = FakeStore(
        "primary@example.com", "primary-store", primary, primary_root
    )

    local = FakeFolder(
        "Calendar (This computer only)",
        "local-calendar",
    )
    local_root = FakeFolder("Root", "local-root", children=(local,))
    local_store = FakeStore(
        "local@example.com", "local-store", local, local_root
    )
    accounts = (
        FakeAccount(primary_store, "primary@example.com"),
        FakeAccount(local_store, "local@example.com"),
    )
    namespace = FakeNamespace(
        [primary_store, local_store], primary_store, accounts
    )
    outlook = FakeOutlook(accounts)
    monkeypatch.setattr(server, "bridge", FakeBridge(outlook, namespace))
    return outlook, namespace, primary, secondary, local


def test_account_resolution_prefers_exact_and_rejects_ambiguity(monkeypatch):
    outlook, namespace, *_ = fixture_environment(monkeypatch)
    assert server._require_store(namespace, "primary@example.com").StoreID == (
        "primary-store"
    )
    with pytest.raises(ValueError, match="ambiguous"):
        server._require_store(namespace, "example.com")
    assert outlook.create_item_calls == 0


def test_list_calendars_identifies_local_only(monkeypatch):
    fixture_environment(monkeypatch)
    payload = json.loads(asyncio.run(server.list_calendars("local@example.com")))
    assert payload == [{
        "account": "local@example.com",
        "calendar": "Calendar (This computer only)",
        "calendar_path": "Calendar (This computer only)",
        "is_default": True,
        "local_only": True,
        "entry_id": "local-calendar",
        "item_count": 0,
    }]


def test_create_event_blocks_local_only_before_side_effect(monkeypatch):
    outlook, _, _, _, local = fixture_environment(monkeypatch)
    with pytest.raises(ToolError, match="local-only"):
        asyncio.run(server.create_event(
            "Blocked",
            "2026-01-01 09:00",
            "2026-01-01 10:00",
            account="local@example.com",
        ))
    assert local.Items.created == []
    assert outlook.create_item_calls == 0


def test_create_event_directly_targets_selected_calendar(monkeypatch):
    outlook, _, primary, secondary, _ = fixture_environment(monkeypatch)
    payload = json.loads(asyncio.run(server.create_event(
        "Project work",
        "2026-01-01 09:00",
        "2026-01-01 10:00",
        account="primary@example.com",
        calendar="Projects",
    )))
    assert payload["calendar"] == "Projects"
    assert payload["local_only"] is False
    assert len(secondary.Items.created) == 1
    assert primary.Items.created == []
    assert outlook.create_item_calls == 0


def test_create_event_preserves_seconds(monkeypatch):
    _, _, primary, _, _ = fixture_environment(monkeypatch)
    asyncio.run(server.create_event(
        "Precise",
        "2026-01-01 09:00:59",
        "2026-01-01 10:00:31",
    ))
    item = primary.Items.created[0]
    assert item.Start == datetime(2026, 1, 1, 9, 0, 59)
    assert item.End == datetime(2026, 1, 1, 10, 0, 31)


def test_create_event_allows_explicit_local_only_opt_in(monkeypatch):
    _, _, _, _, local = fixture_environment(monkeypatch)
    payload = json.loads(asyncio.run(server.create_event(
        "Local",
        "2026-01-01 09:00",
        "2026-01-01 10:00",
        account="local@example.com",
        allow_local_only=True,
    )))
    assert payload["local_only"] is True
    assert len(local.Items.created) == 1


def test_invalid_interval_creates_nothing(monkeypatch):
    outlook, _, primary, _, _ = fixture_environment(monkeypatch)
    with pytest.raises(ToolError, match="later"):
        asyncio.run(server.create_event(
            "Invalid",
            "2026-01-01 10:00",
            "2026-01-01 09:00",
        ))
    assert primary.Items.created == []
    assert outlook.create_item_calls == 0


def test_aware_datetime_is_normalized_to_host_local_time():
    parsed = server._parse_date("2026-01-01T12:00:00+00:00")
    expected = datetime.fromisoformat(
        "2026-01-01T12:00:00+00:00"
    ).astimezone().replace(tzinfo=None)
    assert parsed == expected


def test_move_event_blocks_local_target_and_preserves_item(monkeypatch):
    _, namespace, primary, _, local = fixture_environment(monkeypatch)
    item = FakeItem(parent=primary)
    namespace.items[item.EntryID] = item
    with pytest.raises(ToolError, match="local-only"):
        asyncio.run(server.move_event(
            item.EntryID,
            "local@example.com",
            source_account="primary@example.com",
        ))
    assert item.Parent is primary
    payload = json.loads(asyncio.run(server.move_event(
        item.EntryID,
        "local@example.com",
        source_account="primary@example.com",
        allow_local_only=True,
    )))
    assert payload["old_entry_id"] == "event-1"
    assert payload["entry_id"] == "event-1-moved"
    assert item.Parent is local


def test_calendar_result_limit_is_explicit():
    with pytest.raises(ValueError, match="between 1 and 1000"):
        server._validate_result_count(1001)
