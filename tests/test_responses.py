import json

from outlook_desktop_mcp.utils.responses import (
    bulk_summary,
    tool_result,
    with_meta,
)


def test_tool_result_preserves_json_text_and_structured_content():
    payload = {"status": "ok", "items": [1, 2]}

    result = tool_result(payload, meta={"queue_wait_ms": 12})

    assert json.loads(result.content[0].text) == payload
    assert result.structuredContent == payload
    assert result.meta == {"queue_wait_ms": 12}
    assert result.isError is False


def test_array_result_preserves_legacy_array_text():
    payload = [{"entry_id": "one"}]

    result = tool_result(payload)

    assert json.loads(result.content[0].text) == payload
    assert result.structuredContent is None


def test_with_meta_merges_native_metadata():
    result = with_meta(tool_result("done", meta={"first": 1}), second=2)

    assert result.content[0].text == "done"
    assert result.meta == {"first": 1, "second": 2}


def test_bulk_summary_counts_standard_statuses():
    results = [
        {"status": "ok"},
        {"status": "ok"},
        {"status": "failed"},
        {"status": "skipped"},
    ]

    assert bulk_summary(results) == {
        "total": 4,
        "ok": 2,
        "failed": 1,
        "skipped": 1,
    }
