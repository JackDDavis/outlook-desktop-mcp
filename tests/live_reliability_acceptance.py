"""Explicitly opt-in acceptance against a real Windows Outlook profile."""

import asyncio
import json
import os
import shutil
import sys
import uuid
from pathlib import Path

import pytest
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


pytestmark = [
    pytest.mark.live_outlook,
    pytest.mark.skipif(
        os.name != "nt"
        or os.environ.get("OUTLOOK_MCP_RUN_LIVE_ACCEPTANCE") != "1",
        reason=(
            "Set OUTLOOK_MCP_RUN_LIVE_ACCEPTANCE=1 on Windows with "
            "Outlook Desktop (Classic) running"
        ),
    ),
]


REPO_ROOT = Path(__file__).parents[1]


def test_live_stdio_reads_and_cleans_unique_disposable_draft():
    import win32com.client

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    subject = f"outlook-desktop-mcp-live-{uuid.uuid4()}"
    draft = outlook.CreateItem(0)
    draft.Subject = subject
    draft.Body = "Disposable live acceptance item. Safe to delete."
    draft.Save()
    entry_id = draft.EntryID
    state_dir = REPO_ROOT / ".live-acceptance-state" / uuid.uuid4().hex
    state_dir.mkdir(parents=True)

    async def exercise():
        env = os.environ.copy()
        env["LOCALAPPDATA"] = str(state_dir)
        parameters = StdioServerParameters(
            command=sys.executable,
            args=["-m", "outlook_desktop_mcp.server"],
            cwd=str(REPO_ROOT),
            env=env,
        )
        async with stdio_client(parameters) as (read_stream, write_stream):
            async with ClientSession(read_stream, write_stream) as session:
                await session.initialize()
                tools = await session.list_tools()
                assert len(tools.tools) == 40
                result = await session.call_tool(
                    "read_email",
                    {"entry_id": entry_id, "folder": "drafts"},
                )
                payload = json.loads(result.content[0].text)
                assert payload["entry_id"] == entry_id
                assert payload["subject"] == subject

    try:
        asyncio.run(exercise())
    finally:
        try:
            disposable = namespace.GetItemFromID(entry_id)
            if disposable.Subject == subject:
                disposable.Delete()
        except Exception:
            pass
        shutil.rmtree(state_dir.parent, ignore_errors=True)
