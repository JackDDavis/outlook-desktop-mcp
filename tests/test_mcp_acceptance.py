import asyncio
import json
import os
import socket
import subprocess
import sys
import time
from contextlib import closing
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.sse import sse_client
from mcp.client.stdio import stdio_client

from tests.fakes import make_entry_id


REPO_ROOT = Path(__file__).parents[1]


def _state_env(state_dir):
    env = os.environ.copy()
    env.update({
        "OUTLOOK_MCP_TEST_STATE_DIR": str(state_dir),
        "LOCALAPPDATA": str(state_dir),
        "MCP_OP_BUDGET_SECONDS": "0.01",
        "PYTHONUNBUFFERED": "1",
    })
    return env


def _json_result(result):
    return json.loads(result.content[0].text)


def test_stdio_mcp_discovery_errors_polling_and_text(tmp_path):
    async def exercise():
        parameters = StdioServerParameters(
            command=sys.executable,
            args=["-m", "tests.fake_mcp_server"],
            cwd=str(REPO_ROOT),
            env=_state_env(tmp_path),
        )
        async with stdio_client(parameters) as (read_stream, write_stream):
            async with ClientSession(read_stream, write_stream) as session:
                await session.initialize()

                tools = await session.list_tools()
                assert len(tools.tools) == 40
                assert {
                    "outlook_status",
                    "outlook_operation_status",
                    "bulk_mark_as_read",
                } <= {tool.name for tool in tools.tools}

                invalid = await session.call_tool(
                    "mark_as_read",
                    {"entry_id": "short-id"},
                )
                assert invalid.isError is True
                invalid_payload = _json_result(invalid)
                assert invalid_payload["error"]["code"] == "validation_error"
                assert invalid.structuredContent == invalid_payload

                listed = await session.call_tool(
                    "list_emails",
                    {"count": 1},
                )
                listed_payload = _json_result(listed)
                assert isinstance(listed_payload, list)
                assert listed_payload[0]["subject"] == "Hermetic acceptance 1"
                assert listed.meta["queue_wait_ms"] == 0

                bulk = await session.call_tool(
                    "bulk_mark_as_read",
                    {
                        "entry_ids": ",".join(
                            make_entry_id(index + 1)
                            for index in range(36)
                        ),
                        "count": 36,
                    },
                )
                initial = _json_result(bulk)
                assert initial["status"] == "in_progress"
                statuses = [initial["status"]]

                deadline = time.monotonic() + 10
                while time.monotonic() < deadline:
                    polled = await session.call_tool(
                        "outlook_operation_status",
                        {"operation_id": initial["operation_id"]},
                    )
                    payload = _json_result(polled)
                    statuses.append(payload["status"])
                    if payload["status"] != "in_progress":
                        break
                    await asyncio.sleep(0.02)
                else:
                    raise AssertionError("operation did not complete")

                assert set(statuses) <= {"in_progress", "complete"}
                assert payload["status"] == "complete"
                assert payload["summary"]["ok"] == 36
                assert len(payload["results"]) == 36

    asyncio.run(exercise())


def _free_port():
    with closing(socket.socket()) as listener:
        listener.bind(("127.0.0.1", 0))
        return listener.getsockname()[1]


def _wait_for_port(process, port):
    deadline = time.monotonic() + 10
    while time.monotonic() < deadline:
        if process.poll() is not None:
            raise AssertionError("fake SSE server exited during startup")
        with closing(socket.socket()) as probe:
            probe.settimeout(0.1)
            if probe.connect_ex(("127.0.0.1", port)) == 0:
                return
        time.sleep(0.05)
    raise AssertionError("fake SSE server did not start")


def test_sse_transport_discovers_tools_without_outlook(tmp_path):
    port = _free_port()
    process = subprocess.Popen(
        [
            sys.executable,
            "-m",
            "tests.fake_mcp_server",
            "--http",
            "--host",
            "127.0.0.1",
            "--port",
            str(port),
        ],
        cwd=REPO_ROOT,
        env=_state_env(tmp_path),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    try:
        _wait_for_port(process, port)

        async def exercise():
            async with sse_client(
                f"http://127.0.0.1:{port}/sse",
                timeout=5,
            ) as (read_stream, write_stream):
                async with ClientSession(read_stream, write_stream) as session:
                    await session.initialize()
                    tools = await session.list_tools()
                    assert len(tools.tools) == 40
                    status = await session.call_tool("outlook_status")
                    assert _json_result(status)["com_state"] == "responsive"

        asyncio.run(exercise())
    finally:
        process.terminate()
        process.wait(timeout=10)
