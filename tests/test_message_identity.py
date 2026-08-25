import asyncio
import json
import re
from pathlib import Path

from outlook_desktop_mcp import server
from outlook_desktop_mcp.operations import OperationManager
from outlook_desktop_mcp.utils.formatting import (
    PR_INTERNET_MESSAGE_ID_ANSI,
    PR_INTERNET_MESSAGE_ID_UNICODE,
    format_email_full,
    format_email_summary,
)
from tests.fakes import Collection, FakeMailItem, make_entry_id


class IdentityItems(Collection):
    def __init__(self, values=(), filters=None):
        super().__init__(values)
        self.filters = filters if filters is not None else []

    def Sort(self, *_args):
        return None

    def Restrict(self, filter_string):
        self.filters.append(filter_string)
        match = re.fullmatch(r'@SQL="([^"]+)" = \'(.*)\'', filter_string)
        if not match:
            return self
        property_name, escaped_value = match.groups()
        expected = escaped_value.replace("''", "'")
        matches = []
        for item in self.values:
            try:
                actual = item.PropertyAccessor.GetProperty(property_name)
            except Exception:
                continue
            if str(actual) == expected:
                matches.append(item)
        return IdentityItems(matches, self.filters)


class IdentityFolder:
    def __init__(self, name, items=()):
        self.Name = name
        self.Items = IdentityItems(items)
        self.Folders = Collection()


class IdentityStore:
    DisplayName = "Store"
    StoreID = "store-1"

    def __init__(self, inbox, archive=None):
        self.inbox = inbox
        self.archive = archive or IdentityFolder("Archive")
        self.root = type(
            "Root",
            (),
            {"Folders": Collection([self.inbox, self.archive])},
        )()

    def GetDefaultFolder(self, folder_type):
        assert folder_type == server.OL_FOLDER_INBOX
        return self.inbox

    def GetRootFolder(self):
        return self.root


class IdentityNamespace:
    def __init__(self, store, entry_items=()):
        self.DefaultStore = store
        self.Stores = Collection([store])
        self.Accounts = Collection()
        self.entry_items = {item.EntryID: item for item in entry_items}
        self.entry_calls = []

    def GetItemFromID(self, entry_id, store_id=None):
        self.entry_calls.append((entry_id, store_id))
        if entry_id not in self.entry_items:
            raise LookupError(entry_id)
        return self.entry_items[entry_id]


class FakeBridge:
    def __init__(self, namespace):
        self.namespace = namespace

    async def call(self, function, *args, **kwargs):
        kwargs.pop("timeout_seconds", None)
        kwargs.pop("request_name", None)
        return function(object(), self.namespace, *args, **kwargs)


class MovableMail(FakeMailItem):
    def __init__(self, *args, moved_entry_id, **kwargs):
        super().__init__(*args, **kwargs)
        self.moved_entry_id = moved_entry_id

    def Move(self, destination):
        moved = FakeMailItem(
            self.moved_entry_id,
            subject=self.Subject,
            received_time=self.ReceivedTime,
            properties=self.PropertyAccessor.properties,
        )
        destination.Items.values.append(moved)
        return moved


class FakeAttachment:
    FileName = "report.txt"
    Size = 7

    def SaveAsFile(self, path):
        Path(path).write_text("content", encoding="utf-8")


class Sendable:
    def __init__(self):
        self.Body = "original"
        self.To = ""
        self.CC = ""
        self.BCC = ""
        self.sent = False

    def Send(self):
        self.sent = True


class ActionMail(FakeMailItem):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.Attachments = Collection([FakeAttachment()])
        self.reply = Sendable()
        self.reply_all = Sendable()
        self.forward = Sendable()
        self.Categories = ""

    def Reply(self):
        return self.reply

    def ReplyAll(self):
        return self.reply_all

    def Forward(self):
        return self.forward


def install(monkeypatch, items, *, entry_items=None, archive=None):
    inbox = IdentityFolder("Inbox", items)
    store = IdentityStore(inbox, archive=archive)
    namespace = IdentityNamespace(
        store,
        entry_items=items if entry_items is None else entry_items,
    )
    monkeypatch.setattr(server, "bridge", FakeBridge(namespace))
    return namespace, inbox


def test_message_identity_extraction_present_ansi_fallback_and_absent():
    unicode_item = FakeMailItem(
        make_entry_id(1),
        properties={
            PR_INTERNET_MESSAGE_ID_UNICODE: "<unicode@example.com>",
            PR_INTERNET_MESSAGE_ID_ANSI: "<ansi@example.com>",
        },
    )
    ansi_item = FakeMailItem(
        make_entry_id(2),
        properties={PR_INTERNET_MESSAGE_ID_ANSI: b"<ansi@example.com>"},
    )
    absent_item = FakeMailItem(make_entry_id(3))

    assert format_email_summary(unicode_item)["message_id"] == (
        "<unicode@example.com>"
    )
    assert format_email_full(ansi_item)["message_id"] == "<ansi@example.com>"
    absent = format_email_summary(absent_item)
    assert absent["message_id"] is None
    assert absent["id_stable"] is False


def test_identifier_validation_names_both_supported_formats(monkeypatch):
    install(monkeypatch, [])

    invalid = asyncio.run(server.mark_as_read("short-id"))
    ambiguous = asyncio.run(
        server.mark_as_read("<one@example.com> <two@example.com>")
    )

    for result in (invalid, ambiguous):
        assert result.isError is True
        error = result.structuredContent["error"]
        assert error["code"] == "validation_error"
        assert "Outlook EntryID" in error["message"]
        assert "Message-ID" in error["message"]
    assert "ambiguous" in ambiguous.structuredContent["error"]["message"]


def test_message_id_lookup_is_exact_escaped_and_rejects_duplicates(monkeypatch):
    message_id = "<o'hara@example.com>"
    item = FakeMailItem(
        make_entry_id(1),
        properties={PR_INTERNET_MESSAGE_ID_UNICODE: message_id},
    )
    _namespace, inbox = install(monkeypatch, [item])

    payload = json.loads(asyncio.run(server.read_email(entry_id=message_id)))

    assert payload["entry_id"] == item.EntryID
    assert payload["message_id"] == message_id
    assert inbox.Items.filters == [
        '@SQL="http://schemas.microsoft.com/mapi/proptag/0x1035001F" '
        "= '<o''hara@example.com>'"
    ]

    duplicate = FakeMailItem(
        make_entry_id(2),
        properties={PR_INTERNET_MESSAGE_ID_UNICODE: message_id},
    )
    _namespace, _inbox = install(monkeypatch, [item, duplicate])
    result = asyncio.run(server.read_email(entry_id=message_id))

    assert result.isError is True
    assert result.structuredContent["error"]["code"] == "validation_error"
    assert "2 items matched" in result.structuredContent["error"]["message"]


def test_message_id_lookup_falls_back_to_ansi_property(monkeypatch):
    message_id = "<ansi-only@example.com>"
    item = FakeMailItem(
        make_entry_id(1),
        properties={PR_INTERNET_MESSAGE_ID_ANSI: message_id},
    )
    namespace, inbox = install(monkeypatch, [item])

    payload = json.loads(asyncio.run(server.read_email(entry_id=message_id)))

    assert payload["entry_id"] == item.EntryID
    assert namespace.entry_calls == []
    assert [PR_INTERNET_MESSAGE_ID_UNICODE, PR_INTERNET_MESSAGE_ID_ANSI] == [
        re.search(r'"([^"]+)"', value).group(1)
        for value in inbox.Items.filters
    ]


def test_move_response_preserves_message_identity(monkeypatch):
    message_id = "<move-stable@example.com>"
    old_entry_id = make_entry_id(1)
    new_entry_id = make_entry_id(2)
    item = MovableMail(
        old_entry_id,
        moved_entry_id=new_entry_id,
        properties={PR_INTERNET_MESSAGE_ID_UNICODE: message_id},
    )
    archive = IdentityFolder("Archive")
    install(monkeypatch, [item], archive=archive)

    payload = json.loads(asyncio.run(server.move_email(
        message_id,
        target_folder="archive",
        folder="inbox",
    )))

    assert payload == {
        "status": "moved",
        "subject": "Subject",
        "target_folder": "archive",
        "old_entry_id": old_entry_id,
        "new_entry_id": new_entry_id,
        "message_id": message_id,
        "id_stable": True,
    }


def test_list_search_read_and_bulk_shapes_include_stable_identity(
    monkeypatch,
    tmp_path,
):
    stable_id = "<shape@example.com>"
    stable = FakeMailItem(
        make_entry_id(1),
        subject="Quarterly report",
        properties={PR_INTERNET_MESSAGE_ID_UNICODE: stable_id},
    )
    absent = FakeMailItem(make_entry_id(2), subject="Quarterly report")
    namespace, _inbox = install(monkeypatch, [stable, absent])

    listed = json.loads(asyncio.run(server.list_emails(count=2)))
    searched = json.loads(asyncio.run(server.search_emails("Quarterly", count=2)))
    read = json.loads(asyncio.run(server.read_email(entry_id=stable.EntryID)))

    for row in (listed[0], searched[0], read):
        assert row["message_id"] == stable_id
        assert row["id_stable"] is True
    assert listed[1]["message_id"] is None
    assert listed[1]["id_stable"] is False

    monkeypatch.setattr(
        server,
        "operation_manager",
        OperationManager(tmp_path, process_instance_id="identity-shapes"),
    )
    bulk = json.loads(asyncio.run(server.bulk_read_emails(
        entry_ids=f"{stable.EntryID},{absent.EntryID}",
    )))

    assert bulk["results"][0]["message_id"] == stable_id
    assert bulk["results"][0]["id_stable"] is True
    assert bulk["results"][0]["email"]["message_id"] == stable_id
    assert bulk["results"][1]["message_id"] is None
    assert bulk["results"][1]["id_stable"] is False
    assert namespace.entry_calls


def test_bulk_accepts_message_id_and_legacy_entry_id(
    monkeypatch,
    tmp_path,
):
    message_id = "<bulk-message@example.com>"
    by_message = FakeMailItem(
        make_entry_id(1),
        properties={PR_INTERNET_MESSAGE_ID_UNICODE: message_id},
    )
    by_entry = FakeMailItem(make_entry_id(2))
    namespace, _inbox = install(monkeypatch, [by_message, by_entry])
    monkeypatch.setattr(
        server,
        "operation_manager",
        OperationManager(tmp_path, process_instance_id="identity-bulk"),
    )

    payload = json.loads(asyncio.run(server.bulk_mark_as_read(
        entry_ids=f"{message_id},{by_entry.EntryID}",
    )))

    assert [row["status"] for row in payload["results"]] == ["ok", "ok"]
    assert payload["results"][0]["id"] == message_id
    assert payload["results"][0]["message_id"] == message_id
    assert payload["results"][0]["id_stable"] is True
    assert payload["results"][1]["id"] == by_entry.EntryID
    assert payload["results"][1]["entry_id"] == by_entry.EntryID
    assert payload["results"][1]["message_id"] is None
    assert payload["results"][1]["id_stable"] is False
    assert all(call[0] == by_entry.EntryID for call in namespace.entry_calls)


def test_item_targeting_tools_accept_message_ids_and_keep_generic_entry_ids(
    monkeypatch,
    tmp_path,
):
    message_id = "<actions@example.com>"
    mail = ActionMail(
        make_entry_id(1),
        properties={PR_INTERNET_MESSAGE_ID_UNICODE: message_id},
    )
    generic = type(
        "GenericItem",
        (),
        {
            "EntryID": make_entry_id(2),
            "Attachments": Collection([FakeAttachment()]),
        },
    )()
    namespace, _inbox = install(
        monkeypatch,
        [mail],
        entry_items=[mail, generic],
    )

    assert "Marked as read" in asyncio.run(server.mark_as_read(message_id))
    assert "Marked as unread" in asyncio.run(server.mark_as_unread(message_id))
    assert "Reply sent" in asyncio.run(server.reply_email(message_id, "reply"))
    assert "Forwarded" in asyncio.run(
        server.forward_email(message_id, "recipient@example.com")
    )
    attachments = json.loads(asyncio.run(server.list_attachments(message_id)))
    saved = json.loads(asyncio.run(server.save_attachment(
        message_id,
        save_directory=str(tmp_path),
    )))
    categorized = asyncio.run(server.set_category(message_id, "Follow-up"))
    generic_attachments = json.loads(asyncio.run(
        server.list_attachments(generic.EntryID)
    ))

    assert mail.UnRead is True
    assert mail.reply.sent is True
    assert mail.forward.sent is True
    assert attachments[0]["filename"] == "report.txt"
    assert Path(saved["path"]).read_text(encoding="utf-8") == "content"
    assert "Follow-up" in categorized
    assert generic_attachments[0]["filename"] == "report.txt"
    assert namespace.entry_calls == [(generic.EntryID, None)]
