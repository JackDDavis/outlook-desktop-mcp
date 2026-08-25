"""Sanitized error diagnostics for Outlook tool responses."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Any

_logger = logging.getLogger("outlook_desktop_mcp.errors")

_RPC_UNAVAILABLE = 0x800706BA
_MAPI_EXCEPTION = 0x80020009


@dataclass(frozen=True)
class ErrorDefinition:
    code: str
    meaning: str
    likely_cause: str
    suggested_action: str
    retryable: bool


_HRESULT_ERRORS = {
    _RPC_UNAVAILABLE: ErrorDefinition(
        code="rpc_unavailable",
        meaning="Outlook is not responding",
        likely_cause="Outlook is hung, disconnected, or still initializing",
        suggested_action=(
            "Call outlook_status(); reconnect the MCP server if "
            "com_responsive is false, then restart Outlook if the failure recurs"
        ),
        retryable=True,
    ),
    _MAPI_EXCEPTION: ErrorDefinition(
        code="mapi_exception",
        meaning="Outlook rejected the MAPI operation",
        likely_cause="The item reference became stale or the batch was too large",
        suggested_action=(
            "The item was re-fetched and retried once; if the failure persists, "
            "retry with a smaller batch"
        ),
        retryable=True,
    ),
}

_TIMEOUT_ERROR = ErrorDefinition(
    code="timeout",
    meaning="The Outlook operation exceeded its time budget",
    likely_cause="Outlook is busy, blocked by a dialog, or processing prior work",
    suggested_action="Call outlook_status() and poll any returned operation handle",
    retryable=True,
)

_NOT_FOUND_ERROR = ErrorDefinition(
    code="not_found",
    meaning="The requested Outlook item or folder was not found",
    likely_cause="The item moved, was deleted, or the selector is incorrect",
    suggested_action="Refresh the source folder and retry with a current identifier",
    retryable=False,
)

_VALIDATION_ERROR = ErrorDefinition(
    code="validation_error",
    meaning="The request parameters are invalid",
    likely_cause="A required selector is missing, ambiguous, or malformed",
    suggested_action="Correct the request using the tool parameter documentation",
    retryable=False,
)

_UNKNOWN_COM_ERROR = ErrorDefinition(
    code="com_error",
    meaning="Outlook COM returned an unrecognized error",
    likely_cause="Outlook rejected the operation for an unspecified reason",
    suggested_action="Review the sanitized message and Outlook state before retrying",
    retryable=False,
)

_UNEXPECTED_ERROR = ErrorDefinition(
    code="unexpected_error",
    meaning="An unexpected server error occurred",
    likely_cause="The server encountered an error outside the known Outlook paths",
    suggested_action="Review the server log and retry only after correcting the cause",
    retryable=False,
)


def _extract_hresult(error: Exception) -> int | None:
    hresult = getattr(error, "hresult", None)
    if isinstance(hresult, int):
        return hresult & 0xFFFFFFFF

    if error.args and isinstance(error.args[0], int):
        return error.args[0] & 0xFFFFFFFF
    return None


def _is_com_error(error: Exception) -> bool:
    if _extract_hresult(error) is not None:
        return True
    try:
        import pythoncom

        return isinstance(error, pythoncom.com_error)
    except ImportError:
        return _extract_hresult(error) is not None


def _sanitize_message(message: Any, max_length: int = 300) -> str:
    text = " ".join(str(message or "").split())
    return text[:max_length]


def _definition_for(error: Exception, hresult: int | None) -> ErrorDefinition:
    if isinstance(error, TimeoutError):
        return _TIMEOUT_ERROR
    if isinstance(error, (FileNotFoundError, KeyError, LookupError)):
        return _NOT_FOUND_ERROR
    if isinstance(error, ValueError):
        return _VALIDATION_ERROR
    if hresult is not None:
        return _HRESULT_ERRORS.get(hresult, _UNKNOWN_COM_ERROR)
    return _UNEXPECTED_ERROR


def error_details(error: Exception) -> dict[str, Any]:
    """Return a structured, sanitized diagnostic for an exception."""
    hresult = _extract_hresult(error)
    definition = _definition_for(error, hresult)
    details: dict[str, Any] = {
        "code": definition.code,
        "meaning": definition.meaning,
        "likely_cause": definition.likely_cause,
        "suggested_action": definition.suggested_action,
        "retryable": definition.retryable,
    }
    if hresult is not None:
        details["hresult"] = f"0x{hresult:08X}"

    if definition in {_UNKNOWN_COM_ERROR, _UNEXPECTED_ERROR}:
        message = _sanitize_message(error.args[1] if _is_com_error(error) and len(error.args) > 1 else error)
        if message:
            details["message"] = message

    _logger.debug(
        "Mapped %s to error code %s (HRESULT=%s)",
        type(error).__name__,
        definition.code,
        details.get("hresult"),
    )
    return details


def error_envelope(error: Exception) -> dict[str, dict[str, Any]]:
    """Wrap structured diagnostics in the public error response shape."""
    return {"error": error_details(error)}


def format_com_error(error: Exception) -> str:
    """Preserve the v0.4.0 text API while using structured diagnostics."""
    if isinstance(error, ValueError):
        return _sanitize_message(error)

    details = error_details(error)
    hresult = details.get("hresult")
    if hresult:
        return f"{details['meaning']} ({hresult})"
    if isinstance(error, TimeoutError):
        return details["meaning"]
    return "An unexpected error occurred."
