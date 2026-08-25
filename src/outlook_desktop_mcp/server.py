"""
Outlook Desktop MCP Server
===========================
Exposes Microsoft Outlook Desktop (Classic) as an MCP server over stdio.
Uses COM automation — no Microsoft Graph, no Entra app registration.
Just run this on Windows with Outlook open and you have a full email MCP server.

Entry point: python -m outlook_desktop_mcp.server
"""
import asyncio
import json
import logging
import os
import re
import subprocess
import sys
import time
from datetime import datetime, timedelta

from mcp.server.fastmcp import FastMCP

from outlook_desktop_mcp.com_bridge import OutlookBridge
from outlook_desktop_mcp.operations import OperationManager
from outlook_desktop_mcp.tools._folder_constants import (
    FOLDER_NAME_TO_ENUM,
    OL_APPOINTMENT_ITEM,
    OL_FOLDER_CALENDAR,
    OL_FOLDER_DRAFTS,
    OL_FOLDER_INBOX,
    OL_FOLDER_TASKS,
    OL_MAIL_ITEM,
    OL_MEETING,
    OL_MEETING_CANCELED,
    OL_OPTIONAL,
    OL_REQUIRED,
    OL_RESPONSE_ACCEPTED,
    OL_RESPONSE_DECLINED,
    OL_RESPONSE_TENTATIVE,
    OL_TASK_COMPLETE,
    OL_TASK_ITEM,
)
from outlook_desktop_mcp.utils.errors import (
    error_details,
    error_envelope,
)
from outlook_desktop_mcp.utils.formatting import (
    PR_INTERNET_MESSAGE_ID_ANSI,
    PR_INTERNET_MESSAGE_ID_UNICODE,
    format_email_full,
    format_email_summary,
    format_event_full,
    format_event_summary,
    format_message_identity,
    format_task_full,
    format_task_summary,
)
from outlook_desktop_mcp.utils.responses import tool_result

SERVER_VERSION = "0.5.0"
HEALTH_PROBE_TIMEOUT_SECONDS = 5.0
HEALTH_FRESH_SECONDS = 60.0

# --- Logging (all to stderr, stdout is reserved for MCP JSON-RPC) ---

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("outlook_desktop_mcp")


# --- Security helpers ---

def _safe_dasl(query: str) -> str:
    """Sanitize a string for use in a DASL LIKE filter value.
    Escapes SQL wildcards (% and _) so user input is treated as literals,
    then escapes quote characters required by DASL syntax.
    """
    query = query.replace("%", "[%]").replace("_", "[_]")
    return query.replace("'", "''").replace('"', '""')


# Outlook item Class constants (olObjectClass — distinct from olItemType used in CreateItem)
_OL_CLASS_MAIL = 43
_OL_CLASS_APPOINTMENT = 26
_OL_CLASS_TASK = 48


def _check_item_class(item, expected_class: int, label: str) -> str | None:
    """Return an error string if item is the wrong type, else None."""
    if item.Class != expected_class:
        return f"Error: Entry ID does not refer to a {label}."
    return None


def _tool_error(error: Exception):
    """Return an MCP-native structured error result."""
    return tool_result(error_envelope(error), is_error=True)


def _outlook_process_running() -> bool:
    """Check OUTLOOK.EXE without touching the serialized COM bridge."""
    if os.name != "nt":
        return False
    try:
        completed = subprocess.run(
            [
                "tasklist",
                "/FI",
                "IMAGENAME eq OUTLOOK.EXE",
                "/FO",
                "CSV",
                "/NH",
            ],
            capture_output=True,
            check=False,
            text=True,
            timeout=0.5,
        )
    except (OSError, subprocess.TimeoutExpired):
        return False
    return "OUTLOOK.EXE" in completed.stdout.upper()


def _collect_status_probe(outlook, namespace) -> dict:
    """Collect live Outlook status on the COM thread."""
    application_name = getattr(outlook, "Name", "Microsoft Outlook")
    profile = getattr(namespace, "CurrentProfileName", "") or ""
    accounts = []
    for index in range(namespace.Stores.Count):
        store = namespace.Stores.Item(index + 1)
        row = {
            "name": store.DisplayName,
            "store_id": store.StoreID,
            "unread": None,
            "total": None,
        }
        try:
            inbox = store.GetDefaultFolder(OL_FOLDER_INBOX)
            row["unread"] = inbox.UnReadItemCount
            row["total"] = inbox.Items.Count
        except Exception as error:  # noqa: BLE001 - report per-store degradation
            row["error"] = error_details(error)
        accounts.append(row)
    return {
        "application_name": application_name,
        "profile": profile,
        "accounts": accounts,
    }


# --- MCP Server ---


class OutlookFastMCP(FastMCP):
    """Keep legacy text outputs while tools opt into explicit structured results."""

    def tool(self, *args, **kwargs):
        kwargs.setdefault("structured_output", False)
        return super().tool(*args, **kwargs)


mcp = OutlookFastMCP(
    "outlook-desktop-mcp",
    instructions=(
        "This MCP server gives you full access to Microsoft Outlook Desktop on "
        "Windows via COM automation. It can send emails, read inbox messages, "
        "search across folders, mark messages as read/unread, move messages "
        "between folders (including archive), reply to emails, and list the "
        "complete folder hierarchy.\n\n"
        "All operations use the locally authenticated Outlook profile — no "
        "Microsoft Graph API, no Entra app registration, no OAuth tokens needed. "
        "The user's existing Outlook session handles all authentication.\n\n"
        "PREREQUISITE: Outlook Desktop (Classic) must be running. The new/modern "
        "Outlook (olk.exe) is NOT supported — only the classic OUTLOOK.EXE.\n\n"
        "AVAILABLE TOOL CATEGORIES:\n"
        "- Email: send, list, read, search, reply, mark read/unread, move, attachments\n"
        "- Calendar: list events, create appointments/meetings, update, delete, "
        "respond to invites, search events\n"
        "- Tasks: create, list, complete, update, delete to-do items\n"
        "- Categories: list and set color categories on any item\n"
        "- Rules: list and manage mail rules\n"
        "- Out of Office: check auto-reply status\n"
        "- Folders: list folder hierarchy with item counts"
    ),
)

bridge = OutlookBridge()
operation_manager = OperationManager(
    process_instance_id=bridge.health_snapshot()["process_instance_id"],
)


def _operations_in_flight() -> int:
    """Return the number of durable operations still running."""
    return operation_manager.in_flight()


# --- Helper: resolve store by account name ---

def _resolve_store(namespace, account: str = ""):
    """Resolve an account name to an Outlook Store object.

    If account is empty, returns DefaultStore.
    Otherwise does a case-insensitive substring match on Store.DisplayName.
    """
    if not account:
        return namespace.DefaultStore

    account_lower = account.lower().strip()
    exact = []
    partial = []
    address_by_store = {}
    try:
        for i in range(namespace.Accounts.Count):
            outlook_account = namespace.Accounts.Item(i + 1)
            address_by_store[outlook_account.DeliveryStore.StoreID] = (
                getattr(outlook_account, "SmtpAddress", "") or ""
            ).lower().strip()
    except Exception:
        pass
    for i in range(namespace.Stores.Count):
        store = namespace.Stores.Item(i + 1)
        display_name = store.DisplayName.lower().strip()
        aliases = {
            display_name,
            address_by_store.get(store.StoreID, ""),
        } - {""}
        if account_lower in aliases:
            exact.append(store)
        elif any(account_lower in alias for alias in aliases):
            partial.append(store)

    if len(exact) == 1:
        return exact[0]
    if len(exact) > 1:
        raise ValueError(f"Account selector '{account}' is ambiguous.")
    if len(partial) == 1:
        return partial[0]
    if len(partial) > 1:
        names = ", ".join(store.DisplayName for store in partial)
        raise ValueError(
            f"Account selector '{account}' is ambiguous. Matches: {names}"
        )
    return None


def _require_store(namespace, account: str = ""):
    """Resolve store, raising ValueError if not found."""
    store = _resolve_store(namespace, account)
    if store is None:
        raise ValueError(f"Account '{account}' not found. Use list_accounts to see available accounts.")
    return store


# --- Helper: resolve folder by name ---

def _walk_folders(parent, name_lower: str):
    """Recursively search subfolders of parent for a folder matching name_lower."""
    for i in range(parent.Folders.Count):
        try:
            f = parent.Folders.Item(i + 1)
            if f.Name.lower() == name_lower:
                return f
            found = _walk_folders(f, name_lower)
            if found:
                return found
        except Exception:
            continue
    return None


def _resolve_folder(namespace, folder_name: str, store=None):
    """Resolve a folder name to an Outlook MAPIFolder object.

    Resolution order:
    1. Slash-delimited path (e.g. "Inbox/Receipts") — traverse segment by segment
    2. Built-in Outlook folder enum (inbox, sent, deleted, etc.)
    3. Root-level folder name match (fast path)
    4. Recursive depth-first search of entire folder tree (fallback)
    """
    folder_name = folder_name.strip()
    store = store or namespace.DefaultStore

    # Email-address: treat as account name and return its inbox
    if "@" in folder_name:
        matched_store = _resolve_store(namespace, folder_name)
        if matched_store:
            return matched_store.GetDefaultFolder(OL_FOLDER_INBOX)

    # Slash-delimited path: traverse segment by segment
    if "/" in folder_name:
        parts = [p.strip() for p in folder_name.split("/")]
        current = _resolve_folder(namespace, parts[0], store)
        if current is None:
            return None
        for part in parts[1:]:
            part_lower = part.lower()
            found = None
            for i in range(current.Folders.Count):
                try:
                    f = current.Folders.Item(i + 1)
                    if f.Name.lower() == part_lower:
                        found = f
                        break
                except Exception:
                    continue
            if found is None:
                return None
            current = found
        return current

    folder_lower = folder_name.lower()

    # Built-in Outlook folders
    if folder_lower in FOLDER_NAME_TO_ENUM:
        return store.GetDefaultFolder(FOLDER_NAME_TO_ENUM[folder_lower])

    # Root-level search (fast path)
    root = store.GetRootFolder()
    for i in range(root.Folders.Count):
        try:
            f = root.Folders.Item(i + 1)
            if f.Name.lower() == folder_lower:
                return f
        except Exception:
            continue

    # Recursive fallback: search entire folder tree
    return _walk_folders(root, folder_lower)


def _parse_entry_ids(entry_ids: str) -> list[str]:
    """Parse a delimited list of Outlook EntryIDs or internet Message-IDs."""
    if not entry_ids.strip():
        return []
    return [part.strip() for part in re.split(r"[,\n;]+", entry_ids) if part.strip()]


_ENTRY_ID_RE = re.compile(r"00000000[0-9A-Fa-f]{56,}\Z")
_IDENTIFIER_FORMATS = (
    "an Outlook EntryID (a long hexadecimal value beginning 00000000) or "
    "an internet Message-ID (for example <local@domain> or another single "
    "header-like value containing @)"
)


def _classify_email_identifier(identifier: str) -> str:
    """Classify a strict Outlook EntryID or internet Message-ID."""
    value = (identifier or "").strip()
    if _ENTRY_ID_RE.fullmatch(value) and len(value) % 2 == 0:
        return "entry_id"

    if "@" in value:
        has_unsafe_separator = any(char in value for char in "\r\n,;")
        angle_count = (value.count("<"), value.count(">"))
        angles_valid = angle_count == (0, 0) or (
            angle_count == (1, 1)
            and value.startswith("<")
            and value.endswith(">")
            and not any(char.isspace() for char in value[1:-1])
        )
        if not has_unsafe_separator and angles_valid:
            return "message_id"
        raise ValueError(
            f"Identifier is ambiguous; provide exactly one {_IDENTIFIER_FORMATS}."
        )

    raise ValueError(f"Identifier must be {_IDENTIFIER_FORMATS}.")


def _message_id_restriction(property_name: str, message_id: str) -> str:
    escaped = message_id.replace("'", "''")
    return f'@SQL="{property_name}" = \'{escaped}\''


def _resolve_message_id_item(
    namespace,
    message_id: str,
    account: str,
    folder: str,
):
    if not folder.strip():
        raise ValueError(
            "A folder is required when resolving an internet Message-ID."
        )
    store = _require_store(namespace, account)
    target = _resolve_folder(namespace, folder, store)
    if not target:
        raise ValueError(f"Folder '{folder}' not found")

    for property_name in (
        PR_INTERNET_MESSAGE_ID_UNICODE,
        PR_INTERNET_MESSAGE_ID_ANSI,
    ):
        matches = target.Items.Restrict(
            _message_id_restriction(property_name, message_id)
        )
        if matches.Count > 1:
            raise ValueError(
                f"Message-ID '{message_id}' is ambiguous: {matches.Count} "
                f"items matched in folder '{folder}'. Specify a different "
                "account or folder."
            )
        if matches.Count == 1:
            item = matches.Item(1)
            if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
                raise ValueError(err)
            return item

    raise LookupError(
        f"No email with Message-ID '{message_id}' found in folder '{folder}'"
    )


def _resolve_email_item(
    namespace,
    identifier: str,
    account: str = "",
    folder: str = "inbox",
):
    """Resolve one mail item by EntryID or folder-scoped internet Message-ID."""
    identifier = identifier.strip()
    identifier_kind = _classify_email_identifier(identifier)
    if identifier_kind == "message_id":
        return _resolve_message_id_item(namespace, identifier, account, folder)

    if account:
        store = _require_store(namespace, account)
        item = namespace.GetItemFromID(identifier, store.StoreID)
    else:
        item = namespace.GetItemFromID(identifier)
    if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
        raise ValueError(err)
    return item


def _resolve_outlook_item(
    namespace,
    identifier: str,
    account: str = "",
    folder: str = "inbox",
):
    """Resolve Message-IDs as mail while preserving generic EntryID behavior."""
    identifier = identifier.strip()
    if _classify_email_identifier(identifier) == "message_id":
        return _resolve_message_id_item(namespace, identifier, account, folder)
    if account:
        store = _require_store(namespace, account)
        return namespace.GetItemFromID(identifier, store.StoreID)
    return namespace.GetItemFromID(identifier)


def _build_email_items(namespace, folder: str, unread_only: bool,
                       start_date: str, end_date: str, account: str):
    """Resolve a folder and apply date/unread restrictions."""
    store = _require_store(namespace, account)
    target = _resolve_folder(namespace, folder, store)
    if not target:
        raise ValueError(f"Folder '{folder}' not found")

    items = target.Items
    items.Sort("[ReceivedTime]", True)

    restrictions = []
    if unread_only:
        restrictions.append("[UnRead] = True")
    if start_date:
        start = _parse_date(start_date)
        restrictions.append(f"[ReceivedTime] >= '{start.strftime('%m/%d/%Y %H:%M')}'")
    if end_date:
        end = _parse_date(end_date)
        restrictions.append(f"[ReceivedTime] <= '{end.strftime('%m/%d/%Y %H:%M')}'")
    elif start_date:
        restrictions.append(
            f"[ReceivedTime] <= '{datetime.now().strftime('%m/%d/%Y %H:%M')}'"
        )

    if restrictions:
        items = items.Restrict(" AND ".join(restrictions))
    return items


def _matches_email_filters(item, sender: str, subject_contains: str,
                           body_contains: str) -> bool:
    """Apply in-memory substring filters to a MailItem."""
    if sender:
        sender_lower = sender.lower()
        sender_fields = " ".join([
            getattr(item, "SenderEmailAddress", "") or "",
            getattr(item, "SenderName", "") or "",
        ]).lower()
        if sender_lower not in sender_fields:
            return False

    if subject_contains:
        if subject_contains.lower() not in (item.Subject or "").lower():
            return False

    if body_contains:
        if body_contains.lower() not in (item.Body or "").lower():
            return False

    return True


def _account_type_by_store(outlook) -> dict[str, int]:
    result = {}
    for index in range(outlook.Session.Accounts.Count):
        try:
            account = outlook.Session.Accounts.Item(index + 1)
            result[account.DeliveryStore.StoreID] = account.AccountType
        except Exception:
            continue
    return result


def _requires_client_sender_filter(outlook, store) -> bool:
    # Outlook OlAccountType: Exchange=0, IMAP=1, POP3=2.
    return _account_type_by_store(outlook).get(store.StoreID) in {1, 2}


def _unsupported_sender_dasl(error: Exception) -> bool:
    hresult = error_details(error).get("hresult")
    return hresult in {
        "0x80040102",  # MAPI_E_NO_SUPPORT
        "0x8004010F",  # MAPI_E_NOT_FOUND / unavailable property
        "0x80070057",  # E_INVALIDARG from unsupported Restrict expressions
    }


def _email_in_date_range(item, start_date: str, end_date: str) -> bool:
    received = _parse_date(item.ReceivedTime)
    if start_date and received < _parse_date(start_date):
        return False
    if end_date and received > _parse_date(end_date):
        return False
    return True


def _client_search_emails(
    items,
    *,
    query: str,
    sender: str,
    start_date: str,
    end_date: str,
    count: int,
    include_body: bool,
) -> tuple[list[dict], bool]:
    items.Sort("[ReceivedTime]", True)
    scan_limit = min(items.Count, 1000)
    query_lower = query.lower()
    effective_end = end_date or (
        datetime.now().isoformat() if start_date else ""
    )
    results = []
    for index in range(scan_limit):
        item = items.Item(index + 1)
        haystack = f"{item.Subject or ''}\n{item.Body or ''}".lower()
        if query_lower not in haystack:
            continue
        if not _matches_email_filters(item, sender, "", ""):
            continue
        if not _email_in_date_range(item, start_date, effective_end):
            continue
        results.append(format_email_summary(item, include_body=include_body))
        if len(results) >= count:
            break
    return results, items.Count > scan_limit


def _has_bulk_filter(sender: str, subject_contains: str, body_contains: str,
                     unread_only: bool, start_date: str, end_date: str) -> bool:
    """Return True when a non-folder bulk selector is present."""
    return any([
        sender.strip(),
        subject_contains.strip(),
        body_contains.strip(),
        unread_only,
        start_date.strip(),
        end_date.strip(),
    ])


def _select_email_items(namespace, entry_ids: str = "", sender: str = "",
                        subject_contains: str = "", body_contains: str = "",
                        folder: str = "inbox", unread_only: bool = False,
                        start_date: str = "", end_date: str = "",
                        count: int = 50, account: str = ""):
    """Select email items by explicit EntryID list or folder-based filters."""
    count = min(max(1, count), 100)
    parsed_ids = _parse_entry_ids(entry_ids)
    items = []
    failures = []

    if parsed_ids:
        for identifier in parsed_ids[:count]:
            try:
                item = _resolve_email_item(
                    namespace,
                    identifier,
                    account,
                    folder,
                )
                if _matches_email_filters(item, sender, subject_contains, body_contains):
                    items.append((identifier, item))
                else:
                    failures.append({
                        **_bulk_identity(identifier, item),
                        "status": "skipped",
                        "reason": "filter_mismatch",
                        "error": None,
                    })
            except Exception as e:
                failures.append({
                    **_requested_bulk_identity(identifier),
                    "subject": None,
                    "received_time": None,
                    "status": "failed",
                    "error": error_details(e),
                })
        return items, failures

    mail_items = _build_email_items(namespace, folder, unread_only, start_date, end_date, account)
    scan_limit = min(mail_items.Count, 1000)

    for i in range(scan_limit):
        if len(items) >= count:
            break
        try:
            item = mail_items.Item(i + 1)
            if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
                failures.append({
                    "id": getattr(item, "EntryID", ""),
                    "entry_id": getattr(item, "EntryID", ""),
                    **format_message_identity(item),
                    "subject": getattr(item, "Subject", None),
                    "received_time": str(getattr(item, "ReceivedTime", "")) or None,
                    "status": "failed",
                    "error": error_details(ValueError(err)),
                })
                continue
            if _matches_email_filters(item, sender, subject_contains, body_contains):
                items.append((item.EntryID, item))
        except Exception as e:
            failures.append({
                "id": "",
                "entry_id": "",
                "message_id": None,
                "id_stable": False,
                "subject": None,
                "received_time": None,
                "status": "failed",
                "error": error_details(e),
            })

    return items, failures


def _bulk_identity(identifier: str, item) -> dict:
    return {
        "id": identifier,
        "entry_id": getattr(item, "EntryID", identifier),
        **format_message_identity(item),
        "subject": getattr(item, "Subject", None) or "(no subject)",
        "received_time": str(getattr(item, "ReceivedTime", "")) or None,
    }


def _requested_bulk_identity(identifier: str) -> dict:
    try:
        identifier_kind = _classify_email_identifier(identifier)
    except ValueError:
        identifier_kind = None
    return {
        "id": identifier,
        "entry_id": identifier if identifier_kind == "entry_id" else None,
        "message_id": identifier if identifier_kind == "message_id" else None,
        "id_stable": identifier_kind == "message_id",
    }


def _bulk_payload(results: list[dict], matched_count: int) -> str:
    summary = {
        "total": len(results),
        "ok": sum(row["status"] == "ok" for row in results),
        "failed": sum(row["status"] == "failed" for row in results),
        "skipped": sum(row["status"] == "skipped" for row in results),
    }
    return json.dumps({
        "results": results,
        "summary": summary,
        "matched_count": matched_count,
        "processed_count": len(results),
        "success_count": summary["ok"],
        "failure_count": summary["failed"],
    }, indent=2, default=str)


def _bulk_attempt(
    namespace,
    identifier: str,
    item,
    account: str,
    folder: str,
    action,
):
    identity = _bulk_identity(identifier, item)
    try:
        return action(item), identity, None, "ok"
    except Exception:
        try:
            live_item = _resolve_email_item(
                namespace,
                identifier,
                account,
                folder,
            )
        except Exception as refetch_error:
            return None, identity, refetch_error, "refetch_failed"
        identity = _bulk_identity(identifier, live_item)
        try:
            return action(live_item), identity, None, "retried"
        except Exception as retry_error:
            return None, identity, retry_error, "retry_failed"


OPERATION_BATCH_SIZE = 10
DEFAULT_OPERATION_BUDGET_SECONDS = 90.0


def _operation_budget_seconds() -> float:
    raw = os.environ.get("MCP_OP_BUDGET_SECONDS")
    if raw is None:
        return DEFAULT_OPERATION_BUDGET_SECONDS
    try:
        budget = float(raw)
    except ValueError as error:
        raise ValueError(
            "MCP_OP_BUDGET_SECONDS must be a positive number"
        ) from error
    if budget <= 0:
        raise ValueError("MCP_OP_BUDGET_SECONDS must be a positive number")
    return budget


def _initial_bulk_remaining(entry_ids: str, count: int) -> int:
    limit = min(max(1, count), 100)
    parsed_ids = _parse_entry_ids(entry_ids)
    return min(len(parsed_ids), limit) if parsed_ids else limit


async def _bulk_bridge_call(function, *args, request_name: str):
    kwargs = {}
    if hasattr(bridge, "bulk_timeout_seconds"):
        kwargs = {
            "timeout_seconds": bridge.bulk_timeout_seconds,
            "request_name": request_name,
        }
    return await bridge.call(function, *args, **kwargs)


def _prepare_bulk_email_operation(
    outlook,
    namespace,
    entry_ids,
    sender,
    subject_contains,
    body_contains,
    folder,
    unread_only,
    start_date,
    end_date,
    count,
    account,
    target_folder,
):
    if target_folder:
        store = _require_store(namespace, account)
        if not _resolve_folder(namespace, target_folder, store):
            raise ValueError(
                f"Target folder '{target_folder}' not found. "
                "Use list_folders to see available folders."
            )

    items, failures = _select_email_items(
        namespace,
        entry_ids=entry_ids,
        sender=sender,
        subject_contains=subject_contains,
        body_contains=body_contains,
        folder=folder,
        unread_only=unread_only,
        start_date=start_date,
        end_date=end_date,
        count=count,
        account=account,
    )
    if target_folder:
        for row in failures:
            if (row.get("error") or {}).get("code") == "not_found":
                row["status"] = "skipped"
                row["reason"] = "not_found_in_source"
                row["error"] = None
    return [identifier for identifier, _item in items], failures


def _bulk_resolution_failure(identifier: str, error: Exception) -> dict:
    return {
        **_requested_bulk_identity(identifier),
        "subject": None,
        "received_time": None,
        "status": "failed",
        "error": error_details(error),
    }


def _process_bulk_read_batch(
    outlook,
    namespace,
    identifiers,
    account,
    folder,
):
    results = []
    for identifier in identifiers:
        try:
            item = _resolve_email_item(namespace, identifier, account, folder)
        except Exception as error:
            results.append(_bulk_resolution_failure(identifier, error))
            continue
        email, identity, error, _attempt_state = _bulk_attempt(
            namespace,
            identifier,
            item,
            account,
            folder,
            format_email_full,
        )
        results.append({
            **identity,
            "status": "failed" if error else "ok",
            "email": email,
            "error": error_details(error) if error else None,
        })
    return results


def _process_bulk_mark_batch(
    outlook,
    namespace,
    identifiers,
    account,
    folder,
    unread,
):
    action_name = "marked_as_unread" if unread else "marked_as_read"

    def mark(item):
        item.UnRead = unread
        item.Save()
        return {"action": action_name}

    results = []
    for identifier in identifiers:
        try:
            item = _resolve_email_item(namespace, identifier, account, folder)
        except Exception as error:
            results.append(_bulk_resolution_failure(identifier, error))
            continue
        action, identity, error, _attempt_state = _bulk_attempt(
            namespace,
            identifier,
            item,
            account,
            folder,
            mark,
        )
        results.append({
            **identity,
            "status": "failed" if error else "ok",
            **(action or {}),
            "error": error_details(error) if error else None,
        })
    return results


def _process_bulk_move_batch(
    outlook,
    namespace,
    identifiers,
    account,
    folder,
    target_folder,
):
    store = _require_store(namespace, account)
    destination = _resolve_folder(namespace, target_folder, store)
    if not destination:
        raise ValueError(
            f"Target folder '{target_folder}' not found. "
            "Use list_folders to see available folders."
        )

    def move(item):
        old_entry_id = getattr(item, "EntryID", "")
        moved = item.Move(destination)
        return {
            "action": "moved",
            "target_folder": target_folder,
            "old_entry_id": old_entry_id,
            "new_entry_id": getattr(moved, "EntryID", ""),
            **format_message_identity(moved),
        }

    results = []
    for identifier in identifiers:
        try:
            item = _resolve_email_item(namespace, identifier, account, folder)
        except Exception as error:
            if error_details(error)["code"] == "not_found":
                row = _bulk_resolution_failure(identifier, error)
                row.update({
                    "status": "skipped",
                    "reason": "not_found_in_source",
                    "error": None,
                })
                results.append(row)
            else:
                results.append(_bulk_resolution_failure(identifier, error))
            continue
        action, identity, error, attempt_state = _bulk_attempt(
            namespace,
            identifier,
            item,
            account,
            folder,
            move,
        )
        if (
            error
            and error_details(error)["code"] == "not_found"
            and attempt_state == "refetch_failed"
        ):
            status = "skipped"
            reason = "moved_or_gone_unconfirmed"
            diagnostic = error_details(error)
        else:
            status = "failed" if error else "ok"
            reason = None
            diagnostic = error_details(error) if error else None
        results.append({
            **identity,
            "status": status,
            **(action or {}),
            "reason": reason,
            "error": diagnostic,
        })
    return results


async def _execute_bulk_operation(
    *,
    kind: str,
    selection_args: tuple,
    process_function,
    process_args: tuple,
    initial_remaining: int,
) -> str:
    started = time.monotonic()
    operation_id = operation_manager.create(
        kind,
        remaining=initial_remaining,
    )

    def runner(current_operation_id: str) -> None:
        async def run_batches() -> None:
            identifiers, failures = await _bulk_bridge_call(
                _prepare_bulk_email_operation,
                *selection_args,
                request_name=f"{kind}_select",
            )
            matched_count = len(identifiers)
            operation_manager.append_results(
                current_operation_id,
                failures,
                matched_count=matched_count,
                remaining=matched_count,
            )
            processed = 0
            for offset in range(0, matched_count, OPERATION_BATCH_SIZE):
                if operation_manager.should_stop():
                    operation_manager.interrupt(current_operation_id)
                    return
                batch = identifiers[offset:offset + OPERATION_BATCH_SIZE]
                rows = await _bulk_bridge_call(
                    process_function,
                    batch,
                    *process_args,
                    request_name=f"{kind}_batch",
                )
                processed += len(batch)
                operation_manager.append_results(
                    current_operation_id,
                    rows,
                    matched_count=matched_count,
                    remaining=matched_count - processed,
                )
            operation_manager.complete(current_operation_id)

        asyncio.run(run_batches())

    operation_manager.start(operation_id, runner)
    budget_remaining = max(
        0,
        _operation_budget_seconds() - (time.monotonic() - started),
    )
    payload = await asyncio.to_thread(
        operation_manager.wait,
        operation_id,
        budget_remaining,
    )
    if payload["status"] == "complete":
        return _bulk_payload(payload["results"], payload["matched_count"])
    return json.dumps(payload, indent=2, default=str)


# =====================================================================
# TOOL: outlook_status
# =====================================================================

@mcp.tool()
async def outlook_status():
    """Return immediate Outlook and bridge health diagnostics.

    This tool never queues behind active Outlook work. When the bridge is idle,
    it performs one bounded COM probe; otherwise it reports process, queue,
    active-request, and last-success state from the in-process health snapshot.
    """
    started = time.monotonic()
    outlook_running = _outlook_process_running()
    initial_snapshot = bridge.health_snapshot()
    probe_state = (
        "skipped_bridge_stopped"
        if not initial_snapshot["thread_alive"]
        else "skipped_busy"
    )
    probe_error = None
    com_ping_ms = None

    if initial_snapshot["thread_alive"]:
        try:
            probe = await bridge.call_if_idle_with_metrics(
                _collect_status_probe,
                timeout_seconds=HEALTH_PROBE_TIMEOUT_SECONDS,
                request_name="outlook_status_probe",
            )
        except Exception as error:  # noqa: BLE001 - status reports degraded state
            probe = None
            probe_state = "failed"
            probe_error = error_details(error)
    else:
        probe = None

    if probe is not None:
        probe_state = "completed"
        com_ping_ms = probe.execution_ms
        outlook_running = True
        bridge.update_accounts_snapshot(probe.value["accounts"])
        profile = probe.value["profile"]
        application_name = probe.value["application_name"]
    else:
        profile = ""
        application_name = "Microsoft Outlook"

    snapshot = bridge.health_snapshot()
    fresh_success = (
        snapshot["last_success_age_ms"] is not None
        and snapshot["last_success_age_ms"] <= HEALTH_FRESH_SECONDS * 1000
    )
    busy = snapshot["active_request"] is not None or snapshot["queue_depth"] > 0
    com_responsive = bool(
        snapshot["thread_alive"]
        and (probe_state == "completed" or fresh_success)
    )

    if probe_state == "completed":
        com_state = "responsive"
    elif not snapshot["thread_alive"]:
        com_state = "bridge_stopped"
    elif busy:
        com_state = "busy"
    else:
        com_state = "unresponsive"

    status_elapsed_ms = round((time.monotonic() - started) * 1000)
    payload = {
        "outlook_running": outlook_running,
        "com_responsive": com_responsive,
        "com_state": com_state,
        "com_probe": probe_state,
        "com_ping_ms": com_ping_ms,
        "application_name": application_name,
        "profile": profile,
        "version": SERVER_VERSION,
        "accounts": snapshot["accounts"],
        "accounts_snapshot_at": snapshot["accounts_snapshot_at"],
        "accounts_snapshot_age_ms": snapshot["accounts_snapshot_age_ms"],
        "accounts_stale": probe_state != "completed",
        "queue_depth": snapshot["queue_depth"],
        "active_request": snapshot["active_request"],
        "last_success_at": snapshot["last_success_at"],
        "last_success_age_ms": snapshot["last_success_age_ms"],
        "last_failure_at": snapshot["last_failure_at"],
        "last_failure": snapshot["last_failure"],
        "operations_in_flight": _operations_in_flight(),
        "status_elapsed_ms": status_elapsed_ms,
    }
    if probe_error:
        payload["probe_error"] = probe_error
    return tool_result(
        payload,
        meta={"queue_wait_ms": 0, "execution_ms": status_elapsed_ms},
    )


@mcp.tool()
async def outlook_operation_status(operation_id: str) -> str:
    """Poll a bounded bulk email operation by its operation ID.

    Returns accumulated standard bulk rows and one of: in_progress, complete,
    error, interrupted, or not_found. Interrupted and unknown operations
    include verification guidance and are never resumed automatically.
    """
    return json.dumps(
        operation_manager.get(operation_id),
        indent=2,
        default=str,
    )


# =====================================================================
# TOOL: list_accounts
# =====================================================================

@mcp.tool()
async def list_accounts() -> str:
    """List all Outlook accounts (stores) configured in the profile.

    Returns a JSON array of account objects with display_name, store_id,
    and is_default. Use the display_name (or a unique substring) as the
    'account' parameter in other tools to target a specific account.

    Returns:
        JSON array of account objects.
    """
    def _list(outlook, namespace):
        default_id = namespace.DefaultStore.StoreID
        addresses = {}
        for i in range(outlook.Session.Accounts.Count):
            try:
                account = outlook.Session.Accounts.Item(i + 1)
                store_id = account.DeliveryStore.StoreID
                address = getattr(account, "SmtpAddress", "") or ""
                if not address and "@" in account.DisplayName:
                    address = account.DisplayName
                addresses[store_id] = address
            except Exception:
                continue
        results = []
        for i in range(namespace.Stores.Count):
            store = namespace.Stores.Item(i + 1)
            results.append({
                "display_name": store.DisplayName,
                "email": addresses.get(store.StoreID, ""),
                "store_id": store.StoreID,
                "is_default": store.StoreID == default_id,
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list)
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 1: send_email
# =====================================================================

@mcp.tool()
async def send_email(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
    bcc: str = "",
    html_body: str = "",
    account: str = "",
) -> str:
    """Send an email using the user's Outlook account.

    Creates and sends an email immediately through the default Outlook profile.
    The email will appear in the user's Sent Items folder after sending.

    Args:
        to: One or more recipient email addresses, separated by semicolons.
            Example: "alice@example.com" or "alice@example.com; bob@example.com"
        subject: The email subject line.
        body: The plain-text body of the email. If html_body is also provided,
            both are set and Outlook will prefer the HTML version.
        cc: Optional. CC recipients, separated by semicolons.
        bcc: Optional. BCC recipients, separated by semicolons.
        html_body: Optional. HTML-formatted body. When provided, Outlook renders
            the email as HTML. The plain-text body serves as fallback.
        account: Optional. Account display name (or substring) to send from.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        A confirmation message with subject and recipients, or an error.
    """
    def _send(outlook, namespace, to, subject, body, cc, bcc, html_body, account):
        store = _require_store(namespace, account)
        mail = outlook.CreateItem(OL_MAIL_ITEM)
        # Set the sending account
        for acc in outlook.Session.Accounts:
            if acc.DeliveryStore.StoreID == store.StoreID:
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, acc))  # SendUsingAccount
                break
        mail.To = to
        mail.Subject = subject
        mail.Body = body
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        if html_body:
            mail.HTMLBody = html_body
        mail.Send()
        return f"Email sent: '{subject}' to {to}"

    try:
        return await bridge.call(_send, to, subject, body, cc, bcc, html_body, account)
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 1b: create_draft
# =====================================================================

@mcp.tool()
async def create_draft(
    subject: str,
    body: str,
    to: str = "",
    cc: str = "",
    bcc: str = "",
    html_body: str = "",
    account: str = "",
) -> str:
    """Create a draft email in the Drafts folder without sending it.

    The draft can be reviewed and sent manually by the user in Outlook, or
    located later via list_emails(folder="drafts"). Useful when the user
    wants to review an email before sending, or when the agent is not
    permitted to send directly.

    Args:
        subject: The email subject line.
        body: The plain-text body of the email. If html_body is also provided,
            both are set and Outlook will prefer the HTML version.
        to: Optional. One or more recipient email addresses, separated by
            semicolons. Can be left empty for the user to fill in later.
        cc: Optional. CC recipients, separated by semicolons.
        bcc: Optional. BCC recipients, separated by semicolons.
        html_body: Optional. HTML-formatted body. When provided, Outlook renders
            the email as HTML. The plain-text body serves as fallback.
        account: Optional. Account display name (or substring) to save the
            draft under. Default: primary account. Use list_accounts to see
            available accounts.

    Returns:
        JSON object with entry_id, subject, and account confirmation.
    """
    def _draft(outlook, namespace, subject, body, to, cc, bcc, html_body, account):
        store = _require_store(namespace, account)
        # Create draft in the target account's Drafts folder
        drafts_folder = store.GetDefaultFolder(OL_FOLDER_DRAFTS)
        mail = drafts_folder.Items.Add("IPM.Note")
        # Set the sending account
        for acc in outlook.Session.Accounts:
            if acc.DeliveryStore.StoreID == store.StoreID:
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, acc))  # SendUsingAccount
                break
        mail.Subject = subject
        mail.Body = body
        if to:
            mail.To = to
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        if html_body:
            mail.HTMLBody = html_body
        mail.Save()
        result = {
            "status": "Draft created",
            "entry_id": mail.EntryID,
            "subject": subject,
        }
        if to:
            result["to"] = to
        # Resolve account display name
        for acc in outlook.Session.Accounts:
            if acc.DeliveryStore.StoreID == store.StoreID:
                result["account"] = acc.DisplayName
                break
        return json.dumps(result)

    try:
        return await bridge.call(_draft, subject, body, to, cc, bcc, html_body, account)
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 2: list_emails
# =====================================================================

@mcp.tool()
async def list_emails(
    folder: str = "inbox",
    count: int = 10,
    unread_only: bool = False,
    include_body: bool = False,
    start_date: str = "",
    end_date: str = "",
    account: str = "",
) -> str:
    """List recent emails from a specified Outlook folder.

    Returns a JSON array of email summaries sorted by received time (newest
    first). Each summary includes entry_id, message_id, id_stable, subject,
    sender, sender_name, received_time, unread status, and attachment info.

    Use entry_id for direct Outlook lookup, or message_id with the matching
    account/folder for an identity that survives moves.

    Args:
        folder: The folder to list. Case-insensitive names: "inbox" (default),
            "sent"/"sentmail", "drafts", "deleted"/"trash", "junk"/"spam",
            "outbox", "archive", or any custom folder name visible in
            list_folders output.
        count: Maximum number of emails to return. Default 10, max recommended 50.
        unread_only: If true, only return unread emails. Default false.
        include_body: If true, include to, cc, and a ~300 char body preview
            for each email. Useful for triage without needing read_email
            follow-up calls. Default false.
        start_date: Optional. Only return emails received on or after this date.
            ISO 8601 format (e.g. "2026-03-10" or "2026-03-10 09:00").
        end_date: Optional. Only return emails received on or before this date.
            ISO 8601 format. Default: now (if start_date is provided).
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of email summary objects.
    """
    def _list(outlook, namespace, folder, count, unread_only, include_body, start_date, end_date, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            raise ValueError(f"Folder '{folder}' not found")

        items = target.Items
        items.Sort("[ReceivedTime]", True)

        # Build restriction filters
        restrictions = []
        if unread_only:
            restrictions.append("[UnRead] = True")
        if start_date:
            start = _parse_date(start_date)
            restrictions.append(f"[ReceivedTime] >= '{start.strftime('%m/%d/%Y %H:%M')}'")
        if end_date:
            end = _parse_date(end_date)
            restrictions.append(f"[ReceivedTime] <= '{end.strftime('%m/%d/%Y %H:%M')}'")
        elif start_date:
            # Default end to now when start is specified
            restrictions.append(f"[ReceivedTime] <= '{datetime.now().strftime('%m/%d/%Y %H:%M')}'")

        if restrictions:
            items = items.Restrict(" AND ".join(restrictions))

        results = []
        limit = min(count, items.Count)
        for i in range(limit):
            try:
                results.append(format_email_summary(items.Item(i + 1), include_body=include_body))
            except Exception:
                continue
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, folder, count, unread_only, include_body, start_date, end_date, account)
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 3: read_email
# =====================================================================

@mcp.tool()
async def read_email(
    entry_id: str = "",
    subject_search: str = "",
    folder: str = "inbox",
    account: str = "",
) -> str:
    """Read the full content of a specific email.

    Retrieves complete email details including body text, recipients, CC,
    and metadata. Provide EITHER entry_id (an Outlook EntryID or internet
    Message-ID) OR
    subject_search (finds most recent match by subject substring).

    Args:
        entry_id: An Outlook EntryID or internet Message-ID from list_emails or
            search_emails. Message-ID lookup is exact within folder/account.
        subject_search: Alternative to entry_id. A case-insensitive substring
            to search for in email subjects. Returns the most recent match.
        folder: Folder to search for subject_search or Message-ID lookup.
            Default "inbox".
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON object with full email details (entry_id, subject, sender,
        sender_name, received_time, unread, to, cc, body, attachment info,
        message_id, and id_stable).
    """
    def _read(outlook, namespace, entry_id, subject_search, folder, account):
        if entry_id:
            item = _resolve_email_item(namespace, entry_id, account, folder)
            return json.dumps(format_email_full(item), indent=2, default=str)

        if not subject_search:
            raise ValueError("Provide either entry_id or subject_search")

        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            raise ValueError(f"Folder '{folder}' not found")

        safe_query = _safe_dasl(subject_search)
        filter_str = (
            f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%'"
        )
        items = target.Items.Restrict(filter_str)
        items.Sort("[ReceivedTime]", True)
        if items.Count == 0:
            raise LookupError(f"No email found matching '{subject_search}'")

        return json.dumps(format_email_full(items.Item(1)), indent=2, default=str)

    try:
        return await bridge.call(_read, entry_id, subject_search, folder, account)
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 4: mark_as_read
# =====================================================================

@mcp.tool()
async def mark_as_read(
    entry_id: str,
    account: str = "",
    folder: str = "inbox",
) -> str:
    """Mark a specific email as read in Outlook.

    Changes the unread status to read, same as clicking on an email in Outlook.
    The change is persisted immediately and synced to the server.

    Args:
        entry_id: An Outlook EntryID or internet Message-ID.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.
        folder: Folder containing the email when entry_id is a Message-ID.
            Default "inbox".

    Returns:
        Confirmation message with the email subject, or an error.
    """
    def _mark(outlook, namespace, entry_id, account, folder):
        item = _resolve_email_item(namespace, entry_id, account, folder)
        subject = item.Subject
        item.UnRead = False
        item.Save()
        return f"Marked as read: '{subject}'"

    try:
        return await bridge.call(_mark, entry_id, account, folder)
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 5: mark_as_unread
# =====================================================================

@mcp.tool()
async def mark_as_unread(
    entry_id: str,
    account: str = "",
    folder: str = "inbox",
) -> str:
    """Mark a specific email as unread in Outlook.

    Restores a previously read email to unread status. Useful for flagging
    emails that need follow-up attention. Persisted immediately.

    Args:
        entry_id: An Outlook EntryID or internet Message-ID.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.
        folder: Folder containing the email when entry_id is a Message-ID.
            Default "inbox".

    Returns:
        Confirmation message with the email subject, or an error.
    """
    def _mark(outlook, namespace, entry_id, account, folder):
        item = _resolve_email_item(namespace, entry_id, account, folder)
        subject = item.Subject
        item.UnRead = True
        item.Save()
        return f"Marked as unread: '{subject}'"

    try:
        return await bridge.call(_mark, entry_id, account, folder)
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 6: move_email
# =====================================================================

@mcp.tool()
async def move_email(
    entry_id: str,
    target_folder: str = "archive",
    account: str = "",
    folder: str = "inbox",
) -> str:
    """Move an email to a different Outlook folder.

    Moves the specified email from its current location to the target folder.
    IMPORTANT: After moving, the email gets a NEW entry_id — the old one
    becomes invalid. Common use: archiving emails after processing.

    Args:
        entry_id: An Outlook EntryID or internet Message-ID.
        target_folder: Destination folder name. Default is "archive". Supports
            same names as list_emails: "archive", "inbox", "sent", "deleted"/
            "trash", "drafts", "junk"/"spam", or any custom folder name.
        account: Optional. Account display name (or substring) to resolve
            the target folder in. Default: primary account.
        folder: Source folder used for Message-ID lookup. Default "inbox".

    Returns:
        JSON confirmation with old/new EntryIDs and stable Message-ID.
    """
    def _move(outlook, namespace, entry_id, target_folder, account, folder):
        item = _resolve_email_item(namespace, entry_id, account, folder)
        subject = item.Subject
        old_entry_id = item.EntryID

        store = _require_store(namespace, account)
        dest = _resolve_folder(namespace, target_folder, store)
        if not dest:
            raise ValueError(
                f"Target folder '{target_folder}' not found. "
                "Use list_folders to see available folders."
            )

        moved = item.Move(dest)
        return json.dumps({
            "status": "moved",
            "subject": subject,
            "target_folder": target_folder,
            "old_entry_id": old_entry_id,
            "new_entry_id": moved.EntryID,
            **format_message_identity(moved),
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _move,
            entry_id,
            target_folder,
            account,
            folder,
        )
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 7: reply_email
# =====================================================================

@mcp.tool()
async def reply_email(
    entry_id: str,
    body: str,
    reply_all: bool = False,
    account: str = "",
    folder: str = "inbox",
) -> str:
    """Reply to an email in Outlook.

    Creates and sends a reply, preserving the original message thread.
    Use reply_all=True to reply to all recipients (sender + CC list).

    Args:
        entry_id: An Outlook EntryID or internet Message-ID.
        body: The reply message text. Prepended above the original message
            in the email thread.
        reply_all: If true, reply to all recipients (sender + all CC/To).
            If false (default), reply only to the sender.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.
        folder: Folder containing the email when entry_id is a Message-ID.
            Default "inbox".

    Returns:
        Confirmation indicating the reply was sent, or an error.
    """
    def _reply(outlook, namespace, entry_id, body, reply_all, account, folder):
        item = _resolve_email_item(namespace, entry_id, account, folder)
        subject = item.Subject
        reply_item = item.ReplyAll() if reply_all else item.Reply()
        reply_item.Body = body + "\n\n" + reply_item.Body
        reply_item.Send()
        return f"Reply sent to '{subject}' (reply_all={reply_all})"

    try:
        return await bridge.call(
            _reply,
            entry_id,
            body,
            reply_all,
            account,
            folder,
        )
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 7b: forward_email
# =====================================================================

@mcp.tool()
async def forward_email(
    entry_id: str,
    to: str,
    body: str = "",
    cc: str = "",
    bcc: str = "",
    account: str = "",
    folder: str = "inbox",
) -> str:
    """Forward an email to new recipients, preserving the original content and attachments.

    Creates and sends a forwarded copy of the email. The original message
    (headers, body, and all attachments) is included automatically.

    Args:
        entry_id: An Outlook EntryID or internet Message-ID.
        to: One or more recipient email addresses, separated by semicolons.
        body: Optional. Additional message to prepend above the forwarded content.
        cc: Optional. CC recipients, separated by semicolons.
        bcc: Optional. BCC recipients, separated by semicolons.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.
        folder: Folder containing the email when entry_id is a Message-ID.
            Default "inbox".

    Returns:
        Confirmation indicating the email was forwarded, or an error.
    """
    def _forward(
        outlook,
        namespace,
        entry_id,
        to,
        body,
        cc,
        bcc,
        account,
        folder,
    ):
        item = _resolve_email_item(namespace, entry_id, account, folder)
        subject = item.Subject
        fwd = item.Forward()
        fwd.To = to
        if cc:
            fwd.CC = cc
        if bcc:
            fwd.BCC = bcc
        if body:
            fwd.Body = body + "\n\n" + fwd.Body
        fwd.Send()
        return f"Forwarded '{subject}' to {to}"

    try:
        return await bridge.call(
            _forward,
            entry_id,
            to,
            body,
            cc,
            bcc,
            account,
            folder,
        )
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 8: list_folders
# =====================================================================

@mcp.tool()
async def list_folders(folder: str = "", max_depth: int = 3, account: str = "") -> str:
    """List mail folders in the user's Outlook mailbox.

    When called with no folder argument, lists top-level folders. Provide a
    folder name to drill into its subfolders — use this to browse the full
    folder tree step by step (e.g. first call with no folder to see top-level,
    then call with folder="Inbox" to see Inbox children, then
    folder="Inbox/Projects" to go deeper).

    Folder names from this output can be used directly in list_emails,
    move_email, search_emails, etc. Use slash-delimited paths for nested
    folders (e.g. "Inbox/Receipts/2026").

    Args:
        folder: Optional. Folder to list children of. Supports folder names
            ("Inbox"), slash paths ("Inbox/Receipts"), or built-in names
            ("sent", "drafts"). When empty, lists from the mailbox root.
        max_depth: How many levels deep to recurse below the starting folder.
            Default 3. Set to 1 to see only immediate children.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of folder objects with name, full_path, item_count,
        unread_count, and subfolders (if any).
    """
    def _list(outlook, namespace, folder, max_depth, account):
        max_depth = min(max(1, max_depth), 10)
        store = _require_store(namespace, account)

        if folder:
            start = _resolve_folder(namespace, folder, store)
            if not start:
                raise ValueError(f"Folder '{folder}' not found")
            base_path = folder
        else:
            start = store.GetRootFolder()
            base_path = ""

        def walk(f, depth, path_prefix):
            current_path = f"{path_prefix}/{f.Name}" if path_prefix else f.Name
            result = {
                "name": f.Name,
                "full_path": current_path,
                "item_count": f.Items.Count,
                "unread_count": f.UnReadItemCount,
            }
            if depth < max_depth:
                children = []
                for i in range(f.Folders.Count):
                    try:
                        child = f.Folders.Item(i + 1)
                        children.append(walk(child, depth + 1, current_path))
                    except Exception:
                        continue
                if children:
                    result["subfolders"] = children
            return result

        folders = []
        for i in range(start.Folders.Count):
            try:
                child = start.Folders.Item(i + 1)
                folders.append(walk(child, 1, base_path))
            except Exception:
                continue
        return json.dumps({"account": store.DisplayName, "folders": folders}, indent=2, default=str)

    try:
        return await bridge.call(_list, folder, max_depth, account)
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 9: search_emails
# =====================================================================

@mcp.tool()
async def search_emails(
    query: str,
    folder: str = "inbox",
    count: int = 10,
    include_body: bool = False,
    start_date: str = "",
    end_date: str = "",
    account: str = "",
    sender: str = "",
) -> str:
    """Search for emails in Outlook using text search.

    Searches email subjects and bodies using Outlook's DASL filter.
    Results are sorted by received time (newest first). Each result
    includes entry_id plus stable message_id/id_stable fields.

    Args:
        query: The search term (case-insensitive substring match).
            Examples: "budget report", "meeting notes", "quarterly".
        folder: Folder to search in. Default "inbox". Supports same
            names as list_emails.
        count: Maximum results to return. Default 10.
        include_body: If true, include to, cc, and a ~300 char body preview
            for each result. Default false.
        start_date: Optional. Only return emails received on or after this date.
            ISO 8601 format (e.g. "2026-03-10" or "2026-03-10 09:00").
        end_date: Optional. Only return emails received on or before this date.
            ISO 8601 format. Default: now (if start_date is provided).
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.
        sender: Optional. Case-insensitive substring match against sender email
            address or display name. Sender-filter responses include filter_mode
            and truncation diagnostics.

    Returns:
        JSON array of matching email summaries, or an error.
    """
    def _search(outlook, namespace, query, folder, count, include_body,
                start_date, end_date, account, sender):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            raise ValueError(f"Folder '{folder}' not found")

        if sender and _requires_client_sender_filter(outlook, store):
            results, truncated = _client_search_emails(
                target.Items,
                query=query,
                sender=sender,
                start_date=start_date,
                end_date=end_date,
                count=count,
                include_body=include_body,
            )
            return json.dumps({
                "results": results,
                "filter_mode": "client",
                "truncated": truncated,
            }, indent=2, default=str)

        safe_query = _safe_dasl(query)
        dasl_parts = [
            f"(\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%' OR "
            f"\"urn:schemas:httpmail:textdescription\" LIKE '%{safe_query}%')"
        ]
        if sender:
            safe_sender = _safe_dasl(sender)
            dasl_parts.append(
                f"(\"urn:schemas:httpmail:fromemail\" LIKE '%{safe_sender}%' OR "
                f"\"urn:schemas:httpmail:fromname\" LIKE '%{safe_sender}%')"
            )
        if start_date:
            start = _parse_date(start_date)
            dasl_parts.append(
                f"\"urn:schemas:httpmail:datereceived\" >= '{start.strftime('%m/%d/%Y %H:%M')}'"
            )
        if end_date:
            end = _parse_date(end_date)
            dasl_parts.append(
                f"\"urn:schemas:httpmail:datereceived\" <= '{end.strftime('%m/%d/%Y %H:%M')}'"
            )
        elif start_date:
            dasl_parts.append(
                f"\"urn:schemas:httpmail:datereceived\" <= '{datetime.now().strftime('%m/%d/%Y %H:%M')}'"
            )

        filter_str = "@SQL=" + " AND ".join(dasl_parts)
        try:
            items = target.Items.Restrict(filter_str)
        except Exception as error:
            if not sender or not _unsupported_sender_dasl(error):
                raise
            results, truncated = _client_search_emails(
                target.Items,
                query=query,
                sender=sender,
                start_date=start_date,
                end_date=end_date,
                count=count,
                include_body=include_body,
            )
            return json.dumps({
                "results": results,
                "filter_mode": "client",
                "truncated": truncated,
            }, indent=2, default=str)
        items.Sort("[ReceivedTime]", True)

        results = []
        limit = min(count, items.Count)
        for i in range(limit):
            try:
                results.append(format_email_summary(items.Item(i + 1), include_body=include_body))
            except Exception:
                continue
        payload = (
            {
                "results": results,
                "filter_mode": "dasl",
                "truncated": False,
            }
            if sender
            else results
        )
        return json.dumps(payload, indent=2, default=str)

    try:
        return await bridge.call(
            _search, query, folder, count, include_body, start_date, end_date,
            account, sender,
        )
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# BULK EMAIL TOOLS
# =====================================================================

@mcp.tool()
async def bulk_read_emails(
    entry_ids: str = "",
    sender: str = "",
    subject_contains: str = "",
    body_contains: str = "",
    folder: str = "inbox",
    unread_only: bool = False,
    start_date: str = "",
    end_date: str = "",
    count: int = 10,
    account: str = "",
) -> str:
    """Read multiple emails in one tool call.

    Use this when you need more than one full email body at a time. You can
    either provide explicit Outlook EntryIDs/internet Message-IDs or select
    emails from a folder using filters such as sender, subject_contains,
    unread_only, and date range.

    Args:
        entry_ids: Optional. Comma/semicolon/newline-delimited Outlook EntryIDs
            or internet Message-IDs to read.
        sender: Optional. Case-insensitive substring match against sender email
            address or display name.
        subject_contains: Optional. Case-insensitive substring match in subject.
        body_contains: Optional. Case-insensitive substring match in body text.
        folder: Folder to search, and the scope for Message-ID lookup.
            Default "inbox".
        unread_only: If true, only include unread messages.
        start_date: Optional. Only include emails received on or after this date.
        end_date: Optional. Only include emails received on or before this date.
        count: Maximum number of matching emails to return. Default 10, max 100.
        account: Optional. Account display name (or substring) to target.

    Returns:
        JSON object with selection summary and a results array containing full
        email details or per-item errors.
    """
    try:
        return await _execute_bulk_operation(
            kind="bulk_read_emails",
            selection_args=(
                entry_ids,
                sender,
                subject_contains,
                body_contains,
                folder,
                unread_only,
                start_date,
                end_date,
                count,
                account,
                "",
            ),
            process_function=_process_bulk_read_batch,
            process_args=(account, folder),
            initial_remaining=_initial_bulk_remaining(entry_ids, count),
        )
    except Exception as e:
        return _tool_error(e)


@mcp.tool()
async def bulk_mark_as_read(
    entry_ids: str = "",
    sender: str = "",
    subject_contains: str = "",
    body_contains: str = "",
    folder: str = "inbox",
    unread_only: bool = False,
    start_date: str = "",
    end_date: str = "",
    count: int = 50,
    account: str = "",
) -> str:
    """Mark multiple emails as read.

    Provide explicit Outlook EntryIDs/internet Message-IDs or select emails
    with filters. Message-ID lookup is scoped by folder/account. For safety,
    when entry_ids is omitted you must provide at least one real selector.
    """
    try:
        if not entry_ids.strip() and not _has_bulk_filter(
            sender,
            subject_contains,
            body_contains,
            unread_only,
            start_date,
            end_date,
        ):
            raise ValueError(
                "Provide entry_ids or at least one filter for bulk_mark_as_read."
            )
        return await _execute_bulk_operation(
            kind="bulk_mark_as_read",
            selection_args=(
                entry_ids,
                sender,
                subject_contains,
                body_contains,
                folder,
                unread_only,
                start_date,
                end_date,
                count,
                account,
                "",
            ),
            process_function=_process_bulk_mark_batch,
            process_args=(account, folder, False),
            initial_remaining=_initial_bulk_remaining(entry_ids, count),
        )
    except Exception as e:
        return _tool_error(e)


@mcp.tool()
async def bulk_mark_as_unread(
    entry_ids: str = "",
    sender: str = "",
    subject_contains: str = "",
    body_contains: str = "",
    folder: str = "inbox",
    unread_only: bool = False,
    start_date: str = "",
    end_date: str = "",
    count: int = 50,
    account: str = "",
) -> str:
    """Mark multiple emails as unread.

    Provide explicit Outlook EntryIDs/internet Message-IDs or select emails
    with filters. Message-ID lookup is scoped by folder/account. For safety,
    when entry_ids is omitted you must provide at least one real selector.
    """
    try:
        if not entry_ids.strip() and not _has_bulk_filter(
            sender,
            subject_contains,
            body_contains,
            unread_only,
            start_date,
            end_date,
        ):
            raise ValueError(
                "Provide entry_ids or at least one filter for bulk_mark_as_unread."
            )
        return await _execute_bulk_operation(
            kind="bulk_mark_as_unread",
            selection_args=(
                entry_ids,
                sender,
                subject_contains,
                body_contains,
                folder,
                unread_only,
                start_date,
                end_date,
                count,
                account,
                "",
            ),
            process_function=_process_bulk_mark_batch,
            process_args=(account, folder, True),
            initial_remaining=_initial_bulk_remaining(entry_ids, count),
        )
    except Exception as e:
        return _tool_error(e)


@mcp.tool()
async def bulk_move_emails(
    target_folder: str,
    entry_ids: str = "",
    sender: str = "",
    subject_contains: str = "",
    body_contains: str = "",
    folder: str = "inbox",
    unread_only: bool = False,
    start_date: str = "",
    end_date: str = "",
    count: int = 50,
    account: str = "",
) -> str:
    """Move multiple emails to a destination folder.

    Provide explicit Outlook EntryIDs/internet Message-IDs or select emails
    with filters. Message-ID lookup is scoped by source folder/account. For
    safety, when entry_ids is omitted you must provide a real selector.
    """
    try:
        if not entry_ids.strip() and not _has_bulk_filter(
            sender,
            subject_contains,
            body_contains,
            unread_only,
            start_date,
            end_date,
        ):
            raise ValueError(
                "Provide entry_ids or at least one filter for bulk_move_emails."
            )
        return await _execute_bulk_operation(
            kind="bulk_move_emails",
            selection_args=(
                entry_ids,
                sender,
                subject_contains,
                body_contains,
                folder,
                unread_only,
                start_date,
                end_date,
                count,
                account,
                target_folder,
            ),
            process_function=_process_bulk_move_batch,
            process_args=(account, folder, target_folder),
            initial_remaining=_initial_bulk_remaining(entry_ids, count),
        )
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# CALENDAR TOOLS
# =====================================================================


# --- Calendar helpers ---

MAX_CALENDAR_RESULTS = 1000
LOCAL_ONLY_MARKER = "(this computer only)"


def _parse_date(date_str: str | datetime) -> datetime:
    """Parse ISO 8601 and normalize aware values to the Windows local timezone."""
    if isinstance(date_str, datetime):
        parsed = date_str
        if parsed.tzinfo is not None:
            parsed = parsed.astimezone().replace(tzinfo=None)
        return parsed
    if not date_str or not date_str.strip():
        raise ValueError("Date/time value must not be empty.")
    value = date_str.strip().replace("Z", "+00:00")
    try:
        parsed = datetime.fromisoformat(value)
    except ValueError as exc:
        raise ValueError(
            f"Invalid ISO 8601 date/time: '{date_str}'."
        ) from exc
    if parsed.tzinfo is not None:
        parsed = parsed.astimezone().replace(tzinfo=None)
    return parsed


def _has_explicit_timezone(value: str | datetime) -> bool:
    if isinstance(value, datetime):
        return value.tzinfo is not None
    text = value.strip().replace("Z", "+00:00").replace("z", "+00:00")
    try:
        return datetime.fromisoformat(text).tzinfo is not None
    except ValueError:
        return False


def _local_timezone_name(value: datetime) -> str:
    is_dst = time.localtime(value.timestamp()).tm_isdst > 0
    index = 1 if is_dst and len(time.tzname) > 1 else 0
    return time.tzname[index] or "local"


def _calendar_time_echo(
    start_value: str | datetime,
    end_value: str | datetime,
    *,
    all_day: bool = False,
    input_values: tuple[str | datetime, ...] = (),
) -> dict:
    start_local = _parse_date(start_value)
    end_local = _parse_date(end_value)
    supplied = tuple(value for value in input_values if value not in ("", None))
    if all_day:
        interpreted_as = "all-day"
    elif any(_has_explicit_timezone(value) for value in supplied):
        interpreted_as = "explicit offset"
    else:
        interpreted_as = "local"
    return {
        "start_local": start_local.strftime("%Y-%m-%d %H:%M"),
        "end_local": end_local.strftime("%Y-%m-%d %H:%M"),
        "timezone": _local_timezone_name(start_local),
        "interpreted_as": interpreted_as,
    }


def _validate_calendar_interval(start: str, end: str,
                                all_day: bool = False) -> tuple[datetime, datetime]:
    start_dt = _parse_date(start)
    end_dt = _parse_date(end)
    if end_dt <= start_dt:
        raise ValueError("Event end must be later than event start.")
    if all_day:
        midnight = datetime.min.time()
        if start_dt.time() != midnight or end_dt.time() != midnight:
            raise ValueError(
                "All-day event start and end must be date-aligned at midnight."
            )
    return start_dt, end_dt


def _validate_result_count(count: int) -> int:
    if count < 1 or count > MAX_CALENDAR_RESULTS:
        raise ValueError(
            f"count must be between 1 and {MAX_CALENDAR_RESULTS}."
        )
    return count


def _calendar_path(folder) -> str:
    path = getattr(folder, "FolderPath", "") or getattr(folder, "Name", "")
    return str(path).lstrip("\\").replace("\\", "/")


def _is_local_only_calendar(folder) -> bool:
    text = f"{getattr(folder, 'Name', '')} {_calendar_path(folder)}".lower()
    return LOCAL_ONLY_MARKER in text


def _calendar_candidates(store):
    default = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
    candidates = []
    seen = set()

    def add(folder):
        key = getattr(folder, "EntryID", None) or _calendar_path(folder).lower()
        if key in seen:
            return
        seen.add(key)
        candidates.append(folder)

    add(default)

    def walk(parent):
        for i in range(parent.Folders.Count):
            try:
                child = parent.Folders.Item(i + 1)
                if getattr(child, "DefaultItemType", None) == OL_APPOINTMENT_ITEM:
                    add(child)
                walk(child)
            except Exception:
                continue

    walk(store.GetRootFolder())
    return candidates


def _resolve_calendar(namespace, account: str = "", calendar: str = "",
                      allow_local_only: bool = True):
    store = _require_store(namespace, account)
    default = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
    if not calendar.strip():
        selected = default
    else:
        selector = calendar.lower().strip().replace("\\", "/")
        exact = []
        partial = []
        for candidate in _calendar_candidates(store):
            names = {
                str(candidate.Name).lower().strip(),
                _calendar_path(candidate).lower().strip(),
            }
            if selector in names:
                exact.append(candidate)
            elif any(selector in name for name in names):
                partial.append(candidate)
        matches = exact or partial
        unique = {
            getattr(candidate, "EntryID", _calendar_path(candidate)): candidate
            for candidate in matches
        }
        if len(unique) != 1:
            if not unique:
                raise ValueError(
                    f"Calendar '{calendar}' not found in account "
                    f"'{store.DisplayName}'. Use list_calendars first."
                )
            names = ", ".join(candidate.Name for candidate in unique.values())
            raise ValueError(
                f"Calendar selector '{calendar}' is ambiguous. Matches: {names}"
            )
        selected = next(iter(unique.values()))

    local_only = _is_local_only_calendar(selected)
    if local_only and not allow_local_only:
        raise ValueError(
            f"Calendar '{selected.Name}' is local-only and does not synchronize. "
            "Set allow_local_only=true to opt in explicitly."
        )
    return store, selected, (
        getattr(selected, "EntryID", None) == getattr(default, "EntryID", None)
    ), local_only


def _calendar_context(store, folder, is_default: bool,
                      local_only: bool) -> dict:
    return {
        "account": store.DisplayName,
        "calendar": folder.Name,
        "calendar_path": _calendar_path(folder),
        "is_default": is_default,
        "local_only": local_only,
    }


# =====================================================================
# TOOL 9b: list_calendars
# =====================================================================

@mcp.tool()
async def list_calendars(account: str = "") -> str:
    """List calendar folders for an Outlook account.

    Returns each calendar's name/path, whether it is the account default,
    item count, and whether Outlook marks it as local-only. Use this before
    calendar writes when synchronization matters.

    Args:
        account: Optional account display name, email, or unique substring.
            Default: primary account.
    """
    def _list(outlook, namespace, account):
        store = _require_store(namespace, account)
        default = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
        default_id = getattr(default, "EntryID", None)
        results = []
        for folder in _calendar_candidates(store):
            results.append({
                **_calendar_context(
                    store,
                    folder,
                    getattr(folder, "EntryID", None) == default_id,
                    _is_local_only_calendar(folder),
                ),
                "entry_id": getattr(folder, "EntryID", ""),
                "item_count": folder.Items.Count,
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, account)
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 10: list_events
# =====================================================================

@mcp.tool()
async def list_events(
    start_date: str = "",
    end_date: str = "",
    count: int = 20,
    include_body: bool = False,
    account: str = "",
    calendar: str = "",
    include_meta: bool = False,
) -> str:
    """List upcoming calendar events from Outlook.

    Returns a JSON array of event summaries within a date range, sorted by
    start time. Includes recurring event occurrences. Each summary has
    entry_id, subject, start, end, duration, location, organizer, attendees,
    and status info.

    Use entry_id from results with get_event, update_event, delete_event,
    or respond_to_meeting.

    Args:
        start_date: Start of date range in ISO 8601 format (e.g. "2026-02-25"
            or "2026-02-25 09:00"). Default: now.
        end_date: End of date range. Default: 7 days from start_date.
        count: Maximum number of events to return. Default 20.
        include_body: If true, include a ~300 char body preview and categories
            for each event. Default false.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.
        calendar: Optional calendar name/path. Default: the account's default
            calendar. Use list_calendars to discover calendars.
        include_meta: If true, return an object containing events, count, and
            truncated. Default false preserves the JSON-array response.

    Returns:
        JSON array of event summary objects.
    """
    def _list(outlook, namespace, start_date, end_date, count, include_body,
              account, calendar, include_meta):
        count = _validate_result_count(count)
        _, selected, _, _ = _resolve_calendar(
            namespace, account, calendar, allow_local_only=True
        )
        items = selected.Items

        # CRITICAL ORDER: Sort BEFORE IncludeRecurrences BEFORE Restrict
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        start = _parse_date(start_date) if start_date else datetime.now()
        end = _parse_date(end_date) if end_date else start + timedelta(days=7)

        restrict = (
            f"[Start] >= '{start.strftime('%m/%d/%Y %H:%M')}' "
            f"AND [Start] <= '{end.strftime('%m/%d/%Y %H:%M')}'"
        )
        filtered = items.Restrict(restrict)

        results = []
        truncated = False
        for item in filtered:
            if len(results) >= count:
                truncated = True
                break
            try:
                results.append(format_event_summary(item, include_body=include_body))
            except Exception:
                continue

        payload = (
            {"events": results, "count": len(results), "truncated": truncated}
            if include_meta else results
        )
        return json.dumps(payload, indent=2, default=str)

    try:
        return await bridge.call(
            _list, start_date, end_date, count, include_body, account,
            calendar, include_meta,
        )
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 11: get_event
# =====================================================================

@mcp.tool()
async def get_event(entry_id: str, account: str = "") -> str:
    """Read the full details of a specific calendar event.

    Retrieves complete event information including body/description,
    attendees, recurrence status, reminders, and response status.

    Args:
        entry_id: The unique Outlook EntryID of the event. Get this from
            list_events or search_events results.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON object with full event details.
    """
    def _get(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item"):
            raise ValueError(err)
        return json.dumps(format_event_full(item), indent=2, default=str)

    try:
        return await bridge.call(_get, entry_id, account)
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 12: create_event
# =====================================================================

@mcp.tool()
async def create_event(
    subject: str,
    start: str,
    end: str,
    location: str = "",
    body: str = "",
    all_day: bool = False,
    reminder_minutes: int = 15,
    account: str = "",
    calendar: str = "",
    allow_local_only: bool = False,
) -> str:
    """Create a personal calendar appointment (no attendees).

    Creates and saves an appointment on the user's calendar. This is a
    personal event — no meeting invitations are sent. Use create_meeting
    instead if you need to invite attendees.

    Args:
        subject: The event title.
        start: Start time in ISO 8601 format. Examples: "2026-02-25 14:00",
            "2026-02-25T14:00:00". For all-day events, use just the date:
            "2026-02-25".
        end: End time in ISO 8601 format. For all-day events, use the next
            day: "2026-02-26".
        location: Optional. Event location (e.g. "Conference Room A",
            "Microsoft Teams Meeting").
        body: Optional. Description or notes for the event.
        all_day: If true, creates an all-day event. Default false.
        reminder_minutes: Minutes before the event to show a reminder.
            Default 15. Set to 0 to disable reminder.
        account: Optional. Account display name (or substring) to create
            the event in. Default: primary account.
        calendar: Optional calendar name/path within the selected account.
            Default: the account's default calendar.
        allow_local_only: Explicitly permit a calendar marked
            "(This computer only)". Default false.

    Returns:
        Confirmation with event subject and entry_id, or an error.
    """
    def _create(outlook, namespace, subject, start, end, location, body,
                all_day, reminder_minutes, account, calendar, allow_local_only):
        start_dt, end_dt = _validate_calendar_interval(start, end, all_day)
        store, selected, is_default, local_only = _resolve_calendar(
            namespace, account, calendar, allow_local_only=allow_local_only
        )
        appt = selected.Items.Add("IPM.Appointment")
        appt.Subject = subject
        appt.Start = start_dt
        appt.End = end_dt
        if location:
            appt.Location = location
        if body:
            appt.Body = body
        appt.AllDayEvent = all_day
        if reminder_minutes > 0:
            appt.ReminderSet = True
            appt.ReminderMinutesBeforeStart = reminder_minutes
        else:
            appt.ReminderSet = False
        appt.Save()
        return json.dumps({
            "status": "created",
            "subject": appt.Subject,
            "start": str(appt.Start),
            "end": str(appt.End),
            "entry_id": appt.EntryID,
            **_calendar_time_echo(
                appt.Start,
                appt.End,
                all_day=all_day,
                input_values=(start, end),
            ),
            **_calendar_context(store, selected, is_default, local_only),
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _create, subject, start, end, location, body, all_day,
            reminder_minutes, account, calendar, allow_local_only,
        )
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 13: create_meeting
# =====================================================================

@mcp.tool()
async def create_meeting(
    subject: str,
    start: str,
    end: str,
    required_attendees: str,
    location: str = "",
    body: str = "",
    optional_attendees: str = "",
    account: str = "",
    calendar: str = "",
    allow_local_only: bool = False,
) -> str:
    """Create a meeting and send invitations to attendees.

    Creates a calendar meeting and immediately sends meeting requests to
    all specified attendees. The meeting will appear on the organizer's
    calendar and attendees will receive an invitation they can accept,
    decline, or tentatively accept.

    Args:
        subject: The meeting title.
        start: Start time in ISO 8601 format (e.g. "2026-02-25 14:00").
        end: End time in ISO 8601 format (e.g. "2026-02-25 15:00").
        required_attendees: Required attendee email addresses, separated by
            semicolons. Example: "alice@example.com; bob@example.com"
        location: Optional. Meeting location (e.g. "Teams", "Room 301").
        body: Optional. Meeting description or agenda.
        optional_attendees: Optional. Optional attendee emails, separated
            by semicolons.
        account: Optional. Account display name (or substring) to send from.
            Default: primary account. Use list_accounts to see available accounts.
        calendar: Optional calendar name/path. Meetings must use the selected
            account's default calendar.
        allow_local_only: Explicitly permit a local-only default calendar.
            Default false.

    Returns:
        Confirmation that the meeting was created and invitations sent.
    """
    def _create(outlook, namespace, subject, start, end, required_attendees,
                location, body, optional_attendees, account, calendar,
                allow_local_only):
        start_dt, end_dt = _validate_calendar_interval(start, end)
        store, selected, is_default, local_only = _resolve_calendar(
            namespace, account, calendar, allow_local_only=allow_local_only
        )
        if not is_default:
            raise ValueError(
                "Meetings must be created in the selected account's default "
                "calendar."
            )
        appt = selected.Items.Add("IPM.Appointment")
        # Set sending account
        for acc in outlook.Session.Accounts:
            if acc.DeliveryStore.StoreID == store.StoreID:
                appt._oleobj_.Invoke(*(64209, 0, 8, 0, acc))
                break
        appt.Subject = subject
        appt.Start = start_dt
        appt.End = end_dt
        appt.MeetingStatus = OL_MEETING
        if location:
            appt.Location = location
        if body:
            appt.Body = body

        for addr in required_attendees.split(";"):
            addr = addr.strip()
            if addr:
                recip = appt.Recipients.Add(addr)
                recip.Type = OL_REQUIRED

        if optional_attendees:
            for addr in optional_attendees.split(";"):
                addr = addr.strip()
                if addr:
                    recip = appt.Recipients.Add(addr)
                    recip.Type = OL_OPTIONAL

        appt.Recipients.ResolveAll()
        appt.Send()
        return json.dumps({
            "status": "sent",
            "subject": subject,
            "entry_id": getattr(appt, "EntryID", ""),
            "start": str(appt.Start),
            "end": str(appt.End),
            **_calendar_time_echo(
                appt.Start,
                appt.End,
                input_values=(start, end),
            ),
            **_calendar_context(store, selected, is_default, local_only),
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _create, subject, start, end, required_attendees, location, body,
            optional_attendees, account, calendar, allow_local_only,
        )
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 14: update_event
# =====================================================================

@mcp.tool()
async def update_event(
    entry_id: str,
    subject: str = "",
    start: str = "",
    end: str = "",
    location: str = "",
    body: str = "",
    account: str = "",
) -> str:
    """Update an existing calendar event.

    Modifies properties of an appointment or meeting. Only the fields you
    provide will be updated — omitted fields remain unchanged. For meetings
    you organize, attendees will receive an update notification.

    Args:
        entry_id: The unique Outlook EntryID of the event to update.
        subject: Optional. New event title.
        start: Optional. New start time in ISO 8601 format.
        end: Optional. New end time in ISO 8601 format.
        location: Optional. New location.
        body: Optional. New description/notes.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with updated event details, or an error.
    """
    def _update(outlook, namespace, entry_id, subject, start, end, location, body, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item"):
            raise ValueError(err)
        next_start = _parse_date(start) if start else _parse_date(item.Start)
        next_end = _parse_date(end) if end else _parse_date(item.End)
        _validate_calendar_interval(
            next_start.isoformat(), next_end.isoformat(),
            bool(getattr(item, "AllDayEvent", False)),
        )
        if subject:
            item.Subject = subject
        if start:
            item.Start = next_start
        if end:
            item.End = next_end
        if location:
            item.Location = location
        if body:
            item.Body = body
        item.Save()
        return json.dumps({
            "status": "updated",
            "subject": item.Subject,
            "start": str(item.Start),
            "end": str(item.End),
            "location": item.Location or "",
            "entry_id": item.EntryID,
            **_calendar_time_echo(
                item.Start,
                item.End,
                all_day=bool(getattr(item, "AllDayEvent", False)),
                input_values=(start, end),
            ),
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _update, entry_id, subject, start, end, location, body, account,
        )
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 15: delete_event
# =====================================================================

@mcp.tool()
async def delete_event(entry_id: str, account: str = "") -> str:
    """Delete a calendar event or cancel a meeting.

    For personal appointments, the event is simply deleted. For meetings
    you organized, this cancels the meeting and sends cancellation notices
    to all attendees. For meetings you received, this declines and removes
    the event from your calendar.

    Args:
        entry_id: The unique Outlook EntryID of the event to delete/cancel.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the event subject, or an error.
    """
    def _delete(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item"):
            raise ValueError(err)
        subject = item.Subject
        meeting_status = item.MeetingStatus

        # If this is a meeting we organized, cancel it (sends notices)
        if meeting_status == OL_MEETING:
            item.MeetingStatus = OL_MEETING_CANCELED
            item.Send()
            return f"Meeting canceled: '{subject}' (cancellation sent to attendees)"

        # Otherwise just delete
        item.Delete()
        return f"Event deleted: '{subject}'"

    try:
        return await bridge.call(_delete, entry_id, account)
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 15b: move_event
# =====================================================================

@mcp.tool()
async def move_event(
    entry_id: str,
    target_account: str,
    source_account: str = "",
    target_calendar: str = "",
    allow_local_only: bool = False,
) -> str:
    """Move a calendar event between Outlook account stores.

    Moves the existing AppointmentItem rather than recreating it, preserving
    meeting metadata, attendees, response state, recurrence, reminders, and
    other properties that create_event cannot reproduce.

    Args:
        entry_id: The unique Outlook EntryID of the event to move.
        target_account: Account display name (or unique substring) whose
            default Calendar should receive the event.
        source_account: Optional source account display name (or substring).
            Provide this when moving across stores so the EntryID is resolved
            against the correct source store.
        target_calendar: Optional target calendar name/path. Default: the
            target account's default calendar.
        allow_local_only: Explicitly permit a local-only target calendar.
            Default false.

    Returns:
        JSON confirmation with the moved event's new EntryID and target account.
    """
    def _move(outlook, namespace, entry_id, target_account, source_account,
              target_calendar, allow_local_only):
        if source_account:
            source_store = _require_store(namespace, source_account)
            item = namespace.GetItemFromID(entry_id, source_store.StoreID)
        else:
            source_store = None
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item"):
            raise ValueError(err)

        target_store, selected, is_default, local_only = _resolve_calendar(
            namespace, target_account, target_calendar,
            allow_local_only=allow_local_only,
        )
        current_parent = getattr(item, "Parent", None)
        if (
            current_parent is not None
            and getattr(current_parent, "EntryID", None)
            == getattr(selected, "EntryID", None)
        ):
            raise ValueError("Event is already in the target calendar.")

        old_entry_id = item.EntryID
        subject = item.Subject
        moved = item.Move(selected)
        moved_parent = getattr(moved, "Parent", None)
        if (
            moved_parent is not None
            and getattr(moved_parent, "EntryID", None)
            != getattr(selected, "EntryID", None)
        ):
            raise RuntimeError("Outlook did not place the event in the target calendar.")
        return json.dumps({
            "status": "moved",
            "subject": subject,
            "old_entry_id": old_entry_id,
            "entry_id": moved.EntryID,
            "source_account": (
                source_store.DisplayName if source_store is not None else ""
            ),
            **_calendar_context(target_store, selected, is_default, local_only),
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _move, entry_id, target_account, source_account,
            target_calendar, allow_local_only,
        )
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 16: respond_to_meeting
# =====================================================================

@mcp.tool()
async def respond_to_meeting(
    entry_id: str,
    response: str,
    account: str = "",
) -> str:
    """Respond to a meeting invitation (accept, decline, or tentative).

    Sends your response to the meeting organizer. The meeting will be
    added to (or updated on) your calendar accordingly.

    Args:
        entry_id: The unique Outlook EntryID of the meeting to respond to.
            Get this from list_events or search_events.
        response: Your response. Must be one of: "accept", "decline",
            or "tentative".
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation of your response, or an error.
    """
    def _respond(outlook, namespace, entry_id, response, account):
        response_map = {
            "accept": OL_RESPONSE_ACCEPTED,
            "decline": OL_RESPONSE_DECLINED,
            "tentative": OL_RESPONSE_TENTATIVE,
        }
        response_lower = response.lower().strip()
        if response_lower not in response_map:
            raise ValueError(
                "response must be 'accept', 'decline', or 'tentative'. "
                f"Got: '{response}'"
            )

        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item"):
            raise ValueError(err)
        subject = item.Subject
        response_item = item.Respond(response_map[response_lower])
        response_item.Send()
        return f"Responded '{response_lower}' to meeting: '{subject}'"

    try:
        return await bridge.call(_respond, entry_id, response, account)
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TOOL 17: search_events
# =====================================================================

@mcp.tool()
async def search_events(
    query: str,
    start_date: str = "",
    end_date: str = "",
    count: int = 10,
    include_body: bool = False,
    account: str = "",
    calendar: str = "",
    include_meta: bool = False,
) -> str:
    """Search for calendar events by keyword.

    Searches event subjects within a date range. Results are sorted by
    start time. Includes recurring event occurrences.

    Args:
        query: The search term (case-insensitive substring match on subject).
            Examples: "standup", "review", "1:1".
        start_date: Start of search range in ISO 8601 format. Default: 30
            days ago.
        end_date: End of search range. Default: 30 days from now.
        count: Maximum results to return. Default 10.
        include_body: If true, include a ~300 char body preview and categories
            for each result. Default false.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.
        calendar: Optional calendar name/path. Default: the account's default
            calendar. Use list_calendars to discover calendars.
        include_meta: If true, return an object containing events, count, and
            truncated. Default false preserves the JSON-array response.

    Returns:
        JSON array of matching event summaries.
    """
    def _search(outlook, namespace, query, start_date, end_date, count,
                include_body, account, calendar, include_meta):
        count = _validate_result_count(count)
        _, selected, _, _ = _resolve_calendar(
            namespace, account, calendar, allow_local_only=True
        )
        items = selected.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        start = _parse_date(start_date) if start_date else datetime.now() - timedelta(days=30)
        end = _parse_date(end_date) if end_date else datetime.now() + timedelta(days=30)

        restrict = (
            f"[Start] >= '{start.strftime('%m/%d/%Y %H:%M')}' "
            f"AND [Start] <= '{end.strftime('%m/%d/%Y %H:%M')}'"
        )
        filtered = items.Restrict(restrict)

        query_lower = query.lower()
        results = []
        truncated = False
        for item in filtered:
            if query_lower in (item.Subject or "").lower():
                if len(results) >= count:
                    truncated = True
                    break
                try:
                    results.append(format_event_summary(item, include_body=include_body))
                except Exception:
                    continue

        payload = (
            {"events": results, "count": len(results), "truncated": truncated}
            if include_meta else results
        )
        return json.dumps(payload, indent=2, default=str)

    try:
        return await bridge.call(
            _search, query, start_date, end_date, count, include_body,
            account, calendar, include_meta,
        )
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# TASK TOOLS
# =====================================================================

@mcp.tool()
async def list_tasks(
    include_completed: bool = False,
    count: int = 20,
    include_body: bool = False,
    account: str = "",
) -> str:
    """List tasks from the Outlook Tasks folder.

    Returns a JSON array of task summaries sorted by due date. Each task
    includes entry_id, subject, status, percent_complete, due_date,
    importance, and categories.

    Args:
        include_completed: If true, include completed tasks. Default false
            (only pending/in-progress tasks).
        count: Maximum number of tasks to return. Default 20.
        include_body: If true, include a ~300 char body/notes preview for
            each task. Default false.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of task summary objects.
    """
    def _list(outlook, namespace, include_completed, count, include_body, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
        items = folder.Items
        items.Sort("[DueDate]")

        if not include_completed:
            items = items.Restrict("[Complete] = False")

        results = []
        limit = min(count, items.Count)
        for i in range(limit):
            try:
                results.append(format_task_summary(items.Item(i + 1), include_body=include_body))
            except Exception:
                continue
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, include_completed, count, include_body, account)
    except Exception as e:
        return _tool_error(e)


@mcp.tool()
async def get_task(entry_id: str, account: str = "") -> str:
    """Read the full details of a specific task.

    Args:
        entry_id: The unique Outlook EntryID of the task.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON object with full task details including body.
    """
    def _get(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        return json.dumps(format_task_full(item), indent=2, default=str)

    try:
        return await bridge.call(_get, entry_id, account)
    except Exception as e:
        return _tool_error(e)


@mcp.tool()
async def create_task(
    subject: str,
    body: str = "",
    due_date: str = "",
    importance: str = "normal",
    reminder_minutes: int = 0,
    account: str = "",
) -> str:
    """Create a new task in Outlook.

    Args:
        subject: The task title.
        body: Optional. Task description or notes.
        due_date: Optional. Due date in ISO 8601 format (e.g. "2026-03-01").
        importance: Optional. "low", "normal" (default), or "high".
        reminder_minutes: Optional. Minutes before due date to remind.
            Default 0 (no reminder).
        account: Optional. Account display name (or substring) to create
            the task in. Default: primary account.

    Returns:
        Confirmation with task subject and entry_id.
    """
    def _create(outlook, namespace, subject, body, due_date, importance,
                reminder_minutes, account):
        task = outlook.CreateItem(OL_TASK_ITEM)
        # Move to correct store's tasks folder if account specified
        if account:
            store = _require_store(namespace, account)
            tasks_folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
            task.Move(tasks_folder)
            task = namespace.GetItemFromID(task.EntryID)
        task.Subject = subject
        if body:
            task.Body = body
        if due_date:
            task.DueDate = due_date
        imp_map = {"low": 0, "normal": 1, "high": 2}
        task.Importance = imp_map.get(importance.lower(), 1)
        if reminder_minutes > 0:
            task.ReminderSet = True
            task.ReminderMinutesBeforeStart = reminder_minutes
        else:
            task.ReminderSet = False
        task.Save()
        return json.dumps({
            "status": "created",
            "subject": task.Subject,
            "entry_id": task.EntryID,
            "due_date": str(task.DueDate) if due_date else None,
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _create, subject, body, due_date, importance, reminder_minutes,
            account,
        )
    except Exception as e:
        return _tool_error(e)


@mcp.tool()
async def complete_task(entry_id: str, account: str = "") -> str:
    """Mark a task as complete.

    Sets the task status to complete and percent to 100%.

    Args:
        entry_id: The unique Outlook EntryID of the task.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the task subject.
    """
    def _complete(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_TASK, "task item"):
            raise ValueError(err)
        item.Status = OL_TASK_COMPLETE
        item.PercentComplete = 100
        item.Save()
        return f"Task completed: '{item.Subject}'"

    try:
        return await bridge.call(_complete, entry_id, account)
    except Exception as e:
        return _tool_error(e)


@mcp.tool()
async def delete_task(entry_id: str, account: str = "") -> str:
    """Delete a task from Outlook.

    Args:
        entry_id: The unique Outlook EntryID of the task to delete.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the task subject.
    """
    def _delete(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_TASK, "task item"):
            raise ValueError(err)
        subject = item.Subject
        item.Delete()
        return f"Task deleted: '{subject}'"

    try:
        return await bridge.call(_delete, entry_id, account)
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# ATTACHMENT TOOLS
# =====================================================================

@mcp.tool()
async def list_attachments(
    entry_id: str,
    account: str = "",
    folder: str = "inbox",
) -> str:
    """List all attachments on an email or calendar event.

    Args:
        entry_id: The EntryID of an email/event, or an email Message-ID.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.
        folder: Folder containing the email when entry_id is a Message-ID.
            Default "inbox".

    Returns:
        JSON array of attachment objects with index, filename, and size.
    """
    def _list(outlook, namespace, entry_id, account, folder):
        item = _resolve_outlook_item(namespace, entry_id, account, folder)
        results = []
        for i in range(item.Attachments.Count):
            att = item.Attachments.Item(i + 1)
            results.append({
                "index": i + 1,
                "filename": att.FileName,
                "size": att.Size,
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, entry_id, account, folder)
    except Exception as e:
        return _tool_error(e)


@mcp.tool()
async def save_attachment(
    entry_id: str,
    attachment_index: int = 1,
    save_directory: str = "",
    account: str = "",
    folder: str = "inbox",
) -> str:
    """Save an attachment from an email or event to disk.

    Downloads the specified attachment to a local directory.

    Args:
        entry_id: The EntryID of an email/event, or an email Message-ID.
        attachment_index: Which attachment to save (1-based index). Default 1
            (first attachment). Use list_attachments to see available indices.
        save_directory: Directory to save the file to. Default: user's
            Downloads folder.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.
        folder: Folder containing the email when entry_id is a Message-ID.
            Default "inbox".

    Returns:
        The full file path where the attachment was saved, or an error.
    """
    def _save(
        outlook,
        namespace,
        entry_id,
        attachment_index,
        save_directory,
        account,
        folder,
    ):
        item = _resolve_outlook_item(namespace, entry_id, account, folder)
        if attachment_index < 1 or item.Attachments.Count < attachment_index:
            raise ValueError(
                f"Only {item.Attachments.Count} attachment(s), "
                f"requested index {attachment_index}"
            )

        att = item.Attachments.Item(attachment_index)
        if not save_directory:
            save_directory = os.path.join(os.path.expanduser("~"), "Downloads")

        # Resolve to real path before creating
        save_directory = os.path.realpath(save_directory)
        os.makedirs(save_directory, exist_ok=True)

        # Strip path separators and dangerous characters from filename
        safe_name = os.path.basename(att.FileName)
        safe_name = re.sub(r'[^\w\.\-_ ]', '_', safe_name)
        if not safe_name:
            safe_name = "attachment"

        save_path = os.path.join(save_directory, safe_name)

        # Ensure final path is still inside the intended directory
        if not os.path.realpath(save_path).startswith(save_directory + os.sep) and \
           os.path.realpath(save_path) != save_directory:
            return "Error: Attachment filename would escape the target directory."

        att.SaveAsFile(save_path)
        return json.dumps({
            "status": "saved",
            "filename": safe_name,
            "path": save_path,
            "size": att.Size,
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _save,
            entry_id,
            attachment_index,
            save_directory,
            account,
            folder,
        )
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# CATEGORY TOOLS
# =====================================================================

@mcp.tool()
async def list_categories(account: str = "") -> str:
    """List all available Outlook categories.

    Returns the color categories configured in the user's Outlook profile.
    These can be applied to emails, events, tasks, and other items.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of category objects with name and color index.
    """
    def _list(outlook, namespace, account):
        # Categories are profile-wide, not per-store, but we accept the param for consistency
        results = []
        for i in range(namespace.Categories.Count):
            cat = namespace.Categories.Item(i + 1)
            results.append({"name": cat.Name, "color": cat.Color})
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, account)
    except Exception as e:
        return _tool_error(e)


@mcp.tool()
async def set_category(
    entry_id: str,
    categories: str,
    account: str = "",
    folder: str = "inbox",
) -> str:
    """Set categories on an email, event, or task.

    Replaces any existing categories on the item. Use comma-separated
    values for multiple categories.

    Args:
        entry_id: The EntryID of any item, or an email Message-ID.
        categories: Category name(s), comma-separated. Example:
            "Important" or "Work, Follow-up". Use an empty string to
            clear all categories.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.
        folder: Folder containing the email when entry_id is a Message-ID.
            Default "inbox".

    Returns:
        Confirmation with the item subject and applied categories.
    """
    def _set(outlook, namespace, entry_id, categories, account, folder):
        item = _resolve_outlook_item(namespace, entry_id, account, folder)
        item.Categories = categories
        item.Save()
        return (
            f"Categories set on '{item.Subject}': "
            f"'{item.Categories or '(none)'}'"
        )

    try:
        return await bridge.call(
            _set,
            entry_id,
            categories,
            account,
            folder,
        )
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# RULES TOOLS
# =====================================================================

@mcp.tool()
async def list_rules(account: str = "") -> str:
    """List all mail rules in Outlook.

    Returns the configured inbox rules with their names and enabled status.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of rule objects with name, enabled status, and index.
    """
    def _list(outlook, namespace, account):
        store = _require_store(namespace, account)
        rules = store.GetRules()
        results = []
        for i in range(rules.Count):
            rule = rules.Item(i + 1)
            results.append({
                "index": i + 1,
                "name": rule.Name,
                "enabled": bool(rule.Enabled),
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, account)
    except Exception as e:
        return _tool_error(e)


@mcp.tool()
async def toggle_rule(
    rule_name: str,
    enabled: bool,
    account: str = "",
) -> str:
    """Enable or disable a mail rule by name.

    CAUTION: This modifies live mail rules immediately. Confirm the rule name
    with list_rules before calling.

    Args:
        rule_name: The exact name of the rule to toggle. Use list_rules
            to see available rule names.
        enabled: True to enable the rule, False to disable it.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        Confirmation with the rule name and new status.
    """
    def _toggle(outlook, namespace, rule_name, enabled, account):
        store = _require_store(namespace, account)
        rules = store.GetRules()
        for i in range(rules.Count):
            rule = rules.Item(i + 1)
            if rule.Name == rule_name:
                logger.warning(
                    "toggle_rule: setting rule '%s' enabled=%s", rule_name, enabled
                )
                rule.Enabled = enabled
                rules.Save()
                status = "enabled" if enabled else "disabled"
                return f"Rule '{rule_name}' {status}"
        raise LookupError(
            f"Rule '{rule_name}' not found. Use list_rules to see available rules."
        )

    try:
        return await bridge.call(_toggle, rule_name, enabled, account)
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# OUT OF OFFICE TOOLS
# =====================================================================

@mcp.tool()
async def get_out_of_office(account: str = "") -> str:
    """Check the current Out of Office (auto-reply) status.

    Returns whether Out of Office is currently enabled.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON object with the OOF status.
    """
    def _get(outlook, namespace, account):
        store = _require_store(namespace, account)
        try:
            prop_tag = "http://schemas.microsoft.com/mapi/proptag/0x661D000B"
            oof_state = store.PropertyAccessor.GetProperty(prop_tag)
            return json.dumps({
                "out_of_office": bool(oof_state),
                "status": "on" if oof_state else "off",
            }, indent=2)
        except Exception:
            return json.dumps({
                "out_of_office": None,
                "status": "unknown",
                "note": "Could not read OOF property. Check Outlook settings directly.",
            }, indent=2)

    try:
        return await bridge.call(_get, account)
    except Exception as e:
        return _tool_error(e)


# =====================================================================
# Entry point
# =====================================================================

def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--http", action="store_true", help="Run HTTP/SSE transport instead of stdio")
    parser.add_argument("--host", default="0.0.0.0", help="Host for HTTP transport (default: 0.0.0.0)")
    parser.add_argument("--port", type=int, default=3721, help="Port for HTTP transport (default: 3721)")
    parser.add_argument(
        "--deny",
        default="",
        help="Comma-separated list of tool names to remove (e.g. --deny send_email,reply_email)",
    )
    args = parser.parse_args()

    # Apply tool deny list before starting
    if args.deny:
        denied = [t.strip() for t in args.deny.split(",") if t.strip()]
        available = set(mcp._tool_manager._tools.keys())
        unknown = [t for t in denied if t not in available]
        if unknown:
            parser.error(f"Unknown tool(s) in --deny: {', '.join(unknown)}")
        for tool_name in denied:
            del mcp._tool_manager._tools[tool_name]
        logger.info("Denied %d tool(s): %s", len(denied), ", ".join(denied))

    logger.info("Starting Outlook Desktop MCP server...")
    bridge.start()

    if args.http:
        mcp.settings.host = args.host
        mcp.settings.port = args.port
        logger.info("COM bridge ready. Starting MCP SSE transport on port %d...", args.port)
        try:
            mcp.run(transport="sse")
        finally:
            operation_manager.shutdown()
            bridge.stop()
    else:
        logger.info("COM bridge ready. Starting MCP stdio transport...")
        try:
            mcp.run(transport="stdio")
        finally:
            operation_manager.shutdown()
            bridge.stop()


if __name__ == "__main__":
    main()
