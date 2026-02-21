"""Helpers for extracting and formatting Outlook MailItem data."""
import re


def truncate(text: str, max_length: int = 2000) -> str:
    if len(text) <= max_length:
        return text
    return text[:max_length] + "\n... [truncated]"


def strip_html(html: str) -> str:
    text = re.sub(r"<[^>]+>", "", html)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def format_email_summary(item) -> dict:
    """Extract key fields from an Outlook MailItem into a dict."""
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "sender": getattr(item, "SenderEmailAddress", "unknown"),
        "sender_name": getattr(item, "SenderName", "unknown"),
        "received_time": str(item.ReceivedTime),
        "unread": bool(item.UnRead),
        "has_attachments": bool(item.Attachments.Count > 0),
        "attachment_count": item.Attachments.Count,
    }


def format_email_full(item, body_max_length: int = 5000) -> dict:
    """Extract full email details including body."""
    result = format_email_summary(item)
    result["to"] = item.To or ""
    result["cc"] = item.CC or ""
    result["body"] = truncate(item.Body or "", body_max_length)
    return result
