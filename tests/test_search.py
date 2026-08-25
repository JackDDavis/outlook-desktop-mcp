import asyncio
import json

from outlook_desktop_mcp import server
from tests.fakes import Collection, FakeComError, FakeMailItem


class FakeItems(Collection):
    def __init__(self, values=(), restrict_error=None):
        super().__init__(values)
        self.restrict_error = restrict_error
        self.last_filter = None

    def Sort(self, *_args):
        return None

    def Restrict(self, filter_string):
        self.last_filter = filter_string
        if self.restrict_error:
            raise self.restrict_error
        return self


class FakeFolder:
    def __init__(self, items):
        self.Items = items


class FakeStore:
    def __init__(self, folder):
        self.DisplayName = "Store"
        self.StoreID = "store-1"
        self.folder = folder

    def GetDefaultFolder(self, _folder_type):
        return self.folder


class FakeAccount:
    def __init__(self, store, account_type):
        self.DeliveryStore = store
        self.AccountType = account_type
        self.SmtpAddress = "store@example.com"


class FakeNamespace:
    def __init__(self, store, account):
        self.DefaultStore = store
        self.Stores = Collection([store])
        self.Accounts = Collection([account])


class FakeOutlook:
    def __init__(self, account):
        self.Session = type("Session", (), {"Accounts": Collection([account])})()


class FakeBridge:
    def __init__(self, outlook, namespace):
        self.outlook = outlook
        self.namespace = namespace

    async def call(self, function, *args, **kwargs):
        return function(self.outlook, self.namespace, *args, **kwargs)


def install_search_environment(monkeypatch, items, *, account_type, restrict_error=None):
    collection = FakeItems(items, restrict_error=restrict_error)
    store = FakeStore(FakeFolder(collection))
    account = FakeAccount(store, account_type)
    monkeypatch.setattr(
        server,
        "bridge",
        FakeBridge(FakeOutlook(account), FakeNamespace(store, account)),
    )
    return collection


def test_sender_search_uses_dasl_for_exchange(monkeypatch):
    item = FakeMailItem("one", subject="Quarterly report")
    items = install_search_environment(monkeypatch, [item], account_type=0)

    payload = json.loads(asyncio.run(server.search_emails(
        "Quarterly",
        sender="sender@example.com",
    )))

    assert payload["filter_mode"] == "dasl"
    assert payload["truncated"] is False
    assert payload["results"][0]["entry_id"] == "one"
    assert "fromemail" in items.last_filter


def test_sender_search_uses_client_filter_for_imap_and_reports_cap(monkeypatch):
    items = [
        FakeMailItem(str(index), subject="Quarterly report")
        for index in range(1001)
    ]
    install_search_environment(monkeypatch, items, account_type=1)

    payload = json.loads(asyncio.run(server.search_emails(
        "Quarterly",
        sender="sender@example.com",
        count=2,
    )))

    assert payload["filter_mode"] == "client"
    assert payload["truncated"] is True
    assert len(payload["results"]) == 2


def test_recognized_dasl_failure_falls_back_to_client(monkeypatch):
    item = FakeMailItem("one", subject="Quarterly report")
    install_search_environment(
        monkeypatch,
        [item],
        account_type=0,
        restrict_error=FakeComError(0x80040102),
    )

    payload = json.loads(asyncio.run(server.search_emails(
        "Quarterly",
        sender="sender@example.com",
    )))

    assert payload["filter_mode"] == "client"
    assert payload["results"][0]["entry_id"] == "one"


def test_search_without_sender_preserves_legacy_array(monkeypatch):
    item = FakeMailItem("one", subject="Quarterly report")
    install_search_environment(monkeypatch, [item], account_type=1)

    payload = json.loads(asyncio.run(server.search_emails("Quarterly")))

    assert isinstance(payload, list)
    assert payload[0]["entry_id"] == "one"
