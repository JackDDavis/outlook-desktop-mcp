import json

import pytest

import stdio_sse_proxy as proxy


def test_proxy_filters_manifest_and_rejects_denied_calls():
    manifest = json.dumps({
        "jsonrpc": "2.0",
        "id": 1,
        "result": {
            "tools": [
                {"name": "list_emails"},
                {"name": "send_email"},
            ],
        },
    })
    filtered = json.loads(proxy.filter_response(manifest, {"send_email"}))
    assert filtered["result"]["tools"] == [{"name": "list_emails"}]

    request = json.dumps({
        "jsonrpc": "2.0",
        "id": 2,
        "method": "tools/call",
        "params": {"name": "send_email", "arguments": {}},
    })
    rejection = json.loads(proxy.check_request(request, {"send_email"}))
    assert rejection["id"] == 2
    assert rejection["error"]["code"] == -32601


def test_proxy_rejects_cross_host_message_endpoint(monkeypatch):
    monkeypatch.setattr(
        proxy,
        "SSE_CONNECT_URL",
        "http://127.0.0.1:3721",
    )
    assert proxy.validate_message_endpoint("/messages/one") == (
        "http://127.0.0.1:3721/messages/one"
    )
    with pytest.raises(ValueError, match="unexpected endpoint host"):
        proxy.validate_message_endpoint("http://example.com/messages/one")
