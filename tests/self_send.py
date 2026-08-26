"""Recipient policy for live COM/MCP validation tests.

Live validation tests must NEVER send mail or meeting invitations to an
address outside the local Outlook profile. The only permitted recipient is
the profile's own address ("self-send"), optionally overridden by the
OUTLOOK_MCP_SELF_SEND environment variable. If the self address cannot be
determined, the caller must SKIP the send step -- never fall back to a
placeholder address such as user@example.com.

Additionally, sending ANY email (including self-send) requires explicit
user approval before sending. Scripts default to skipping all send steps;
they only send when OUTLOOK_MCP_ALLOW_SEND is set to an affirmative value.
When an agent runs these scripts, in-session user approval is required in
addition to the environment flag.
"""
import os

ENV_OVERRIDE = "OUTLOOK_MCP_SELF_SEND"
ENV_ALLOW_SEND = "OUTLOOK_MCP_ALLOW_SEND"


def send_allowed():
    """Return True only when live sending is explicitly opted in.

    Sending requires explicit approval. Unattended runs (CI, validation)
    default to skipping all send and meeting-invite steps.
    """
    return os.environ.get(ENV_ALLOW_SEND, "").strip().lower() in ("1", "true", "yes")


def self_address(outlook):
    """Return the profile's own SMTP address for a live COM test, or None.

    ``outlook`` is a running ``Outlook.Application`` COM object. The address
    is derived at runtime from the session's accounts so no personal address
    is ever hardcoded in test code.
    """
    override = os.environ.get(ENV_OVERRIDE, "").strip()
    if override:
        return override
    try:
        accounts = outlook.Session.Accounts
        for i in range(1, accounts.Count + 1):
            address = (accounts.Item(i).SmtpAddress or "").strip()
            if address:
                return address
    except Exception:
        pass
    return None


async def self_address_via_mcp(session):
    """Return the default account's SMTP address via ``list_accounts``.

    ``session`` is an MCP ``ClientSession``. Returns the ``email`` of the
    account flagged ``is_default`` (falling back to the first account with
    an email address), or None if no SMTP address is available.
    """
    import json

    result = await session.call_tool("list_accounts", {})
    try:
        accounts = json.loads(result.content[0].text)
    except (ValueError, AttributeError, IndexError):
        return None
    if not isinstance(accounts, list) or not accounts:
        return None
    for account in accounts:
        if account.get("is_default"):
            return (account.get("email") or "").strip() or None
    for account in accounts:
        address = (account.get("email") or "").strip()
        if address:
            return address
    return None
