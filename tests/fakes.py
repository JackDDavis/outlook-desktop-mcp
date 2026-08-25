"""Reusable hermetic Outlook-like fakes for reliability tests."""

from __future__ import annotations


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


def make_entry_id(number=1):
    return f"00000000{number:056X}"


class FakePropertyAccessor:
    def __init__(self, properties=None):
        self.properties = dict(properties or {})

    def GetProperty(self, name):
        if name not in self.properties:
            raise KeyError(name)
        return self.properties[name]


class FakeMailItem:
    Class = 43

    def __init__(
        self,
        entry_id,
        subject="Subject",
        received_time="2026-01-01 09:00:00",
        *,
        unread=True,
        properties=None,
    ):
        self.EntryID = entry_id
        self.Subject = subject
        self.ReceivedTime = received_time
        self.UnRead = unread
        self.SenderEmailAddress = "sender@example.com"
        self.SenderName = "Sender"
        self.To = ""
        self.CC = ""
        self.Body = ""
        self.Attachments = Collection()
        self.PropertyAccessor = FakePropertyAccessor(properties)
        self.saved = 0

    def Save(self):
        self.saved += 1


class FakeComError(Exception):
    def __init__(self, hresult, message="COM failure"):
        super().__init__(hresult, message, None, None)
        self.hresult = hresult
