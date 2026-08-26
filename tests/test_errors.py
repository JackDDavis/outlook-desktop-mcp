from outlook_desktop_mcp.utils.errors import error_details, error_envelope
from tests.fakes import FakeComError


def test_rpc_unavailable_maps_to_actionable_envelope():
    details = error_details(FakeComError(-2147023174))

    assert details == {
        "code": "rpc_unavailable",
        "meaning": "Outlook is not responding",
        "likely_cause": "Outlook is hung, disconnected, or still initializing",
        "suggested_action": (
            "Call outlook_status(); reconnect the MCP server if "
            "com_responsive is false, then restart Outlook if the failure recurs"
        ),
        "retryable": True,
        "hresult": "0x800706BA",
    }


def test_mapi_exception_is_retryable():
    details = error_details(FakeComError(-2147352567))

    assert details["code"] == "mapi_exception"
    assert details["hresult"] == "0x80020009"
    assert details["retryable"] is True


def test_server_execution_failure_is_retryable():
    details = error_details(FakeComError(-2146959355))

    assert details["code"] == "server_execution_failed"
    assert details["hresult"] == "0x80080005"
    assert details["retryable"] is True


def test_outlook_not_registered_identifies_integrity_mismatch():
    details = error_details(FakeComError(-2147221021))

    assert details["code"] == "outlook_not_registered"
    assert details["hresult"] == "0x800401E3"
    assert details["retryable"] is True


def test_unknown_com_error_keeps_only_sanitized_message():
    details = error_details(FakeComError(-1, "  private   COM detail  "))

    assert details["code"] == "com_error"
    assert details["hresult"] == "0xFFFFFFFF"
    assert details["message"] == "private COM detail"
    assert details["retryable"] is False


def test_timeout_and_validation_have_non_com_codes():
    assert error_details(TimeoutError())["code"] == "timeout"
    assert error_envelope(ValueError("bad"))["error"]["code"] == "validation_error"
