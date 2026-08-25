"""Live SSE calendar safety smoke test.

Creates one disposable personal appointment, verifies local-only protection,
moves it across stores when a local-only calendar exists, moves it back, and
deletes it. No meeting invitations are sent.
"""
import asyncio
import json
import os
import sys
from contextlib import asynccontextmanager
from datetime import datetime, timedelta


def log(message):
    print(message, flush=True)


def parse_json(result):
    text = result.content[0].text
    return json.loads(text)


@asynccontextmanager
async def connect():
    from mcp.client.session import ClientSession

    proxy = os.environ.get("OUTLOOK_MCP_STDIO_PROXY")
    if proxy:
        from mcp import StdioServerParameters
        from mcp.client.stdio import stdio_client

        params = StdioServerParameters(command=sys.executable, args=[proxy])
        async with stdio_client(params) as streams:
            async with ClientSession(*streams) as session:
                yield session
        return

    from mcp.client.sse import sse_client

    async with sse_client("http://localhost:3721/sse") as streams:
        async with ClientSession(*streams) as session:
            yield session


async def run():
    subject = f"OutlookMCP v0.4 safety test {datetime.now():%Y%m%d%H%M%S}"
    current_entry_id = None
    current_account = ""
    passed = 0
    total = 0

    async with connect() as session:
            await session.initialize()
            tools = {tool.name for tool in (await session.list_tools()).tools}

            total += 1
            required = {
                "list_accounts", "list_calendars", "create_event",
                "get_event", "move_event", "delete_event", "search_events",
            }
            missing = required - tools
            assert not missing, f"Missing tools: {sorted(missing)}"
            passed += 1
            log("PASS tool discovery")

            accounts = parse_json(await session.call_tool("list_accounts", {}))
            primary = next(account for account in accounts if account["is_default"])
            primary_selector = primary["display_name"]
            local_target = None
            for account in accounts:
                calendars = parse_json(await session.call_tool(
                    "list_calendars",
                    {"account": account["display_name"]},
                ))
                local = next(
                    (calendar for calendar in calendars if calendar["local_only"]),
                    None,
                )
                if local:
                    local_target = (account["display_name"], local["calendar_path"])
                    break

            start = datetime.now() + timedelta(days=2)
            start = start.replace(hour=9, minute=0, second=0, microsecond=0)
            end = start + timedelta(hours=1)

            try:
                total += 1
                invalid = await session.call_tool("create_event", {
                    "subject": subject + " invalid",
                    "start": end.isoformat(),
                    "end": start.isoformat(),
                })
                assert invalid.isError, "Invalid interval was not an MCP error"
                passed += 1
                log("PASS validation error signaling")

                total += 1
                created = parse_json(await session.call_tool("create_event", {
                    "subject": subject,
                    "start": start.isoformat(),
                    "end": end.isoformat(),
                    "reminder_minutes": 0,
                }))
                current_entry_id = created["entry_id"]
                current_account = created["account"]
                assert created["account"] == primary_selector
                assert created["local_only"] is False
                passed += 1
                log("PASS destination-first creation")

                total += 1
                read_back = parse_json(await session.call_tool("get_event", {
                    "entry_id": current_entry_id,
                    "account": current_account,
                }))
                assert read_back["subject"] == subject
                passed += 1
                log("PASS create/read round trip")

                if local_target:
                    local_account, local_calendar = local_target
                    total += 1
                    blocked = await session.call_tool("move_event", {
                        "entry_id": current_entry_id,
                        "source_account": current_account,
                        "target_account": local_account,
                        "target_calendar": local_calendar,
                    })
                    assert blocked.isError, "Local-only move was not blocked"
                    passed += 1
                    log("PASS local-only default block")

                    total += 1
                    moved = parse_json(await session.call_tool("move_event", {
                        "entry_id": current_entry_id,
                        "source_account": current_account,
                        "target_account": local_account,
                        "target_calendar": local_calendar,
                        "allow_local_only": True,
                    }))
                    current_entry_id = moved["entry_id"]
                    current_account = local_account
                    assert moved["local_only"] is True
                    passed += 1
                    log("PASS explicit local-only opt-in")

                    total += 1
                    moved_back = parse_json(await session.call_tool("move_event", {
                        "entry_id": current_entry_id,
                        "source_account": current_account,
                        "target_account": primary_selector,
                    }))
                    current_entry_id = moved_back["entry_id"]
                    current_account = primary_selector
                    assert moved_back["local_only"] is False
                    passed += 1
                    log("PASS metadata-preserving return move")
            finally:
                if current_entry_id:
                    await session.call_tool("delete_event", {
                        "entry_id": current_entry_id,
                        "account": current_account,
                    })

            total += 1
            remaining = parse_json(await session.call_tool("search_events", {
                "query": subject,
                "start_date": (start - timedelta(days=1)).isoformat(),
                "end_date": (end + timedelta(days=1)).isoformat(),
                "count": 10,
            }))
            assert remaining == [], f"Test artifact remained: {remaining}"
            passed += 1
            log("PASS cleanup")

    log(f"Results: {passed}/{total} passed")
    return passed == total


if __name__ == "__main__":
    sys.exit(0 if asyncio.run(run()) else 1)
