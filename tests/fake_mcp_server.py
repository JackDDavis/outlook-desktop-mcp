"""Hermetic MCP server process used by transport acceptance tests."""

import argparse
import os
import time
from pathlib import Path

from outlook_desktop_mcp import server
from outlook_desktop_mcp.com_bridge import BridgeCallResult
from outlook_desktop_mcp.operations import OperationManager
from outlook_desktop_mcp.utils.formatting import (
    PR_INTERNET_MESSAGE_ID_UNICODE,
)
from outlook_desktop_mcp.utils.responses import with_meta
from tests.fakes import Collection, FakeMailItem, make_entry_id


class FakeItems(Collection):
    def Sort(self, *_args):
        return None

    def Restrict(self, filter_string):
        if "0x1035001F" not in filter_string:
            return self
        message_id = filter_string.rsplit(" = '", 1)[-1][:-1].replace("''", "'")
        return FakeItems([
            item
            for item in self.values
            if item.PropertyAccessor.properties.get(
                PR_INTERNET_MESSAGE_ID_UNICODE
            ) == message_id
        ])


class FakeFolder:
    def __init__(self, name, items=()):
        self.Name = name
        self.Items = FakeItems(items)
        self.Folders = Collection()


class FakeStore:
    DisplayName = "Hermetic Acceptance Store"
    StoreID = "hermetic-store"

    def __init__(self, items):
        self.inbox = FakeFolder("Inbox", items)
        self.archive = FakeFolder("Archive")
        self.root = type(
            "Root",
            (),
            {"Folders": Collection([self.inbox, self.archive])},
        )()

    def GetDefaultFolder(self, _folder_type):
        return self.inbox

    def GetRootFolder(self):
        return self.root


class FakeNamespace:
    CurrentProfileName = "Hermetic Acceptance"
    CurrentUser = type("User", (), {"Name": "Hermetic User"})()

    def __init__(self, items):
        self.DefaultStore = FakeStore(items)
        self.Stores = Collection([self.DefaultStore])
        self.Accounts = Collection()
        self.items = {item.EntryID: item for item in items}

    def GetItemFromID(self, entry_id, _store_id=None):
        if entry_id not in self.items:
            raise LookupError(entry_id)
        return self.items[entry_id]


class FakeOutlook:
    Name = "Microsoft Outlook (Hermetic)"
    Session = type("Session", (), {"Accounts": Collection()})()


class SlowMailItem(FakeMailItem):
    def Save(self):
        time.sleep(0.005)
        super().Save()


class FakeBridge:
    bulk_timeout_seconds = 5

    def __init__(self, namespace):
        self.outlook = FakeOutlook()
        self.namespace = namespace
        self.accounts = []

    def start(self, timeout=60):
        return None

    def stop(self):
        return None

    async def call(self, function, *args, **kwargs):
        request_name = kwargs.pop("request_name", function.__name__)
        kwargs.pop("timeout_seconds", None)
        started = time.monotonic()
        value = function(self.outlook, self.namespace, *args, **kwargs)
        execution_ms = round((time.monotonic() - started) * 1000)
        if hasattr(value, "content"):
            return with_meta(
                value,
                queue_wait_ms=0,
                execution_ms=execution_ms,
                request_name=request_name,
            )
        return value

    async def call_if_idle_with_metrics(
        self,
        function,
        *args,
        request_name,
        **kwargs,
    ):
        kwargs.pop("timeout_seconds", None)
        started = time.monotonic()
        value = function(self.outlook, self.namespace, *args, **kwargs)
        return BridgeCallResult(
            value=value,
            queue_wait_ms=0,
            execution_ms=round((time.monotonic() - started) * 1000),
            request_name=request_name,
        )

    def update_accounts_snapshot(self, accounts):
        self.accounts = [dict(account) for account in accounts]

    def health_snapshot(self):
        return {
            "process_instance_id": "hermetic-process",
            "thread_alive": True,
            "queue_depth": 0,
            "active_request": None,
            "last_success_at": None,
            "last_success_age_ms": None,
            "last_failure_at": None,
            "last_failure": None,
            "accounts": self.accounts,
            "accounts_snapshot_at": None,
            "accounts_snapshot_age_ms": None,
        }


def install_runtime():
    items = [
        SlowMailItem(
            make_entry_id(index + 1),
            subject=f"Hermetic acceptance {index + 1}",
            properties={
                PR_INTERNET_MESSAGE_ID_UNICODE: (
                    f"<hermetic-{index + 1:02}@example.com>"
                )
            },
        )
        for index in range(40)
    ]
    server.bridge = FakeBridge(FakeNamespace(items))
    state_dir = Path(os.environ["OUTLOOK_MCP_TEST_STATE_DIR"])
    server.operation_manager = OperationManager(
        state_dir,
        process_instance_id="hermetic-process",
    )
    server._outlook_process_running = lambda: True


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--http", action="store_true")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=3721)
    args = parser.parse_args()
    install_runtime()
    try:
        if args.http:
            server.mcp.settings.host = args.host
            server.mcp.settings.port = args.port
            server.mcp.run(transport="sse")
        else:
            server.mcp.run(transport="stdio")
    finally:
        server.operation_manager.shutdown()
        server.bridge.stop()


if __name__ == "__main__":
    main()
