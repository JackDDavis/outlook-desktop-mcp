"""
Stdio → SSE proxy for Outlook Desktop MCP server.

OpenClaw (WSL) launches this via the shared-tools Python. Connects to the
SSE server running non-elevated on Windows port 3721.

WSL can't reach Windows 127.0.0.1, but can reach the Windows host via the
default gateway IP (172.x.x.1). Uvicorn's host-header check requires
Host: 127.0.0.1:3721 regardless of actual IP used for the TCP connection.
"""
import asyncio
import sys
import os
import struct
import socket
from urllib.parse import urlparse

PORT = 3721


def get_windows_host_ip():
    """Return Windows host IP as seen from WSL2 (default gateway)."""
    try:
        with open("/proc/net/route") as f:
            for line in f.readlines()[1:]:
                fields = line.strip().split()
                if len(fields) >= 3 and fields[1] == "00000000":
                    return socket.inet_ntoa(struct.pack("<L", int(fields[2], 16)))
    except Exception:
        pass
    return "127.0.0.1"


def get_sse_url():
    try:
        with open("/proc/version") as f:
            if "microsoft" in f.read().lower():
                ip = get_windows_host_ip()
                return f"http://{ip}:{PORT}", f"127.0.0.1:{PORT}"
    except Exception:
        pass
    return f"http://127.0.0.1:{PORT}", f"127.0.0.1:{PORT}"


SSE_CONNECT_URL, HOST_HEADER = get_sse_url()
SSE_URL = SSE_CONNECT_URL + "/sse"


async def proxy():
    import httpx
    from httpx_sse import aconnect_sse

    message_endpoint = None
    response_queue = asyncio.Queue()

    headers = {"Host": HOST_HEADER}

    async def read_sse(client):
        nonlocal message_endpoint
        async with aconnect_sse(client, "GET", SSE_URL, headers=headers) as event_source:
            async for sse in event_source.aiter_sse():
                if sse.event == "endpoint":
                    endpoint = sse.data
                    if endpoint.startswith("/"):
                        endpoint = SSE_CONNECT_URL + endpoint
                    # Validate endpoint points to the expected server
                    parsed = urlparse(endpoint)
                    expected = urlparse(SSE_CONNECT_URL)
                    if parsed.netloc != expected.netloc:
                        raise ValueError(
                            f"SSE server returned unexpected endpoint host: {parsed.netloc}"
                        )
                    message_endpoint = endpoint
                elif sse.event == "message":
                    await response_queue.put(sse.data)

    async def write_responses():
        while True:
            data = await response_queue.get()
            sys.stdout.buffer.write((data + "\n").encode("utf-8"))
            sys.stdout.buffer.flush()

    async def read_stdin(client):
        # asyncio pipe readers can fail on Windows subprocess stdio handles.
        # Use a thread-backed blocking readline for cross-platform behavior.
        while True:
            line = await asyncio.to_thread(sys.stdin.readline)
            if not line:
                break
            line = line.strip()
            if not line:
                continue
            while message_endpoint is None:
                await asyncio.sleep(0.05)
            await client.post(
                message_endpoint,
                content=line,
                headers={"Content-Type": "application/json", "Host": HOST_HEADER},
            )

    async with httpx.AsyncClient(timeout=None) as client:
        await asyncio.gather(
            read_sse(client),
            write_responses(),
            read_stdin(client),
        )


if __name__ == "__main__":
    try:
        asyncio.run(proxy())
    except KeyboardInterrupt:
        pass
