import asyncio
import json

from tests.self_send import self_address_via_mcp, send_allowed


class Result:
    def __init__(self, value):
        self.content = [type("Text", (), {"text": json.dumps(value)})()]


class Session:
    def __init__(self, accounts):
        self.accounts = accounts

    async def call_tool(self, name, arguments):
        assert name == "list_accounts"
        assert arguments == {}
        return Result(self.accounts)


def test_send_requires_explicit_opt_in(monkeypatch):
    monkeypatch.delenv("OUTLOOK_MCP_ALLOW_SEND", raising=False)
    assert send_allowed() is False

    monkeypatch.setenv("OUTLOOK_MCP_ALLOW_SEND", "yes")
    assert send_allowed() is True


def test_mcp_self_address_uses_default_account_email():
    session = Session(
        [
            {
                "display_name": "Mailbox label",
                "email": "default@example.com",
                "is_default": True,
            },
            {
                "display_name": "Other label",
                "email": "other@example.com",
                "is_default": False,
            },
        ]
    )

    assert asyncio.run(self_address_via_mcp(session)) == "default@example.com"


def test_mcp_self_address_falls_back_to_first_available_email():
    session = Session(
        [
            {"display_name": "No address", "email": "", "is_default": False},
            {
                "display_name": "Mailbox label",
                "email": "available@example.com",
                "is_default": False,
            },
        ]
    )

    assert asyncio.run(self_address_via_mcp(session)) == "available@example.com"
