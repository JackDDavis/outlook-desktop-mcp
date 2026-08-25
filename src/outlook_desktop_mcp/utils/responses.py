"""MCP-native response helpers with backwards-compatible text content."""

from __future__ import annotations

import json
from collections.abc import Mapping
from typing import Any

from mcp.types import CallToolResult, TextContent


def json_text(payload: Any) -> str:
    """Serialize a public tool payload consistently."""
    return json.dumps(payload, indent=2, default=str)


def tool_result(
    payload: Any,
    *,
    meta: Mapping[str, Any] | None = None,
    is_error: bool = False,
) -> CallToolResult:
    """Build a native MCP result without changing legacy text content."""
    text = payload if isinstance(payload, str) else json_text(payload)
    structured_content = payload if isinstance(payload, dict) else None
    return CallToolResult(
        content=[TextContent(type="text", text=text)],
        structuredContent=structured_content,
        isError=is_error,
        _meta=dict(meta) if meta else None,
    )


def with_meta(result: CallToolResult, **meta: Any) -> CallToolResult:
    """Return a copy of a result with merged protocol-native metadata."""
    merged = dict(result.meta or {})
    merged.update(meta)
    return result.model_copy(update={"meta": merged})


def bulk_summary(results: list[Mapping[str, Any]]) -> dict[str, int]:
    """Summarize standard per-item bulk statuses."""
    counts = {"ok": 0, "failed": 0, "skipped": 0}
    for result in results:
        status = result.get("status")
        if status in counts:
            counts[status] += 1
    return {"total": len(results), **counts}
