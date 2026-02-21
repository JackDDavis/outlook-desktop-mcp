"""
Outlook Desktop MCP - Phase 1 COM Validation
=============================================
Standalone script that validates all Outlook COM operations before building
the MCP layer. Requires Classic Outlook (Desktop) to be running.

Run: .venv\\Scripts\\python tests\\phase1_com_test.py
"""
import sys
import time

def log(msg):
    print(msg, file=sys.stderr)


def test_connect():
    """Test 1: Connect to running Outlook via COM."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    log(f"  Connected to store: {namespace.DefaultStore.DisplayName}")
    log(f"  Current user: {namespace.CurrentUser.Name}")
    return outlook, namespace


def test_list_folders(namespace):
    """Test 2: Enumerate default folders and their item counts."""
    folder_map = {
        "Inbox": 6, "Sent Mail": 5, "Deleted Items": 3,
        "Drafts": 16, "Calendar": 9, "Tasks": 13, "Junk": 23,
    }
    for name, enum_val in folder_map.items():
        try:
            folder = namespace.GetDefaultFolder(enum_val)
            log(f"  {name}: {folder.Items.Count} items, {folder.UnReadItemCount} unread")
        except Exception as e:
            log(f"  {name}: SKIP ({e})")


def test_read_inbox(namespace, count=5):
    """Test 3: Read top N emails from Inbox (newest first)."""
    inbox = namespace.GetDefaultFolder(6)
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    actual = min(count, items.Count)
    for i in range(actual):
        item = items.Item(i + 1)  # 1-indexed
        log(f"  [{i+1}] {item.Subject}")
        log(f"       From: {item.SenderEmailAddress}")
        log(f"       Date: {item.ReceivedTime}")
        log(f"       EntryID: {item.EntryID[:40]}...")
        log(f"       UnRead: {item.UnRead}")
    if actual == 0:
        log("  (inbox is empty)")


def test_filter_unread(namespace):
    """Test 4: Use Restrict to find unread emails in Inbox."""
    inbox = namespace.GetDefaultFolder(6)
    unread = inbox.Items.Restrict("[UnRead] = True")
    log(f"  Unread count: {unread.Count}")
    if unread.Count > 0:
        first = unread.Item(1)
        log(f"  First unread: {first.Subject}")
    return unread.Count


def test_send_email(outlook):
    """Test 5: Create and send a test email to self."""
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    mail.To = "user@example.com"
    mail.Subject = "Outlook Desktop MCP - COM Test"
    mail.Body = (
        "This is an automated test from the Outlook Desktop MCP COM validation.\n"
        f"Sent at: {time.strftime('%Y-%m-%d %H:%M:%S')}\n"
        "\nIf you received this, COM send_email works."
    )
    mail.Send()
    log("  Email sent to user@example.com")


def test_mark_read_unread(namespace):
    """Test 6: Find first unread, mark as read, then restore to unread."""
    inbox = namespace.GetDefaultFolder(6)
    unread = inbox.Items.Restrict("[UnRead] = True")
    if unread.Count == 0:
        log("  SKIP: No unread emails to test with")
        return
    item = unread.Item(1)
    subject = item.Subject
    log(f"  Target: '{subject}'")
    log(f"  Marking as read...")
    item.UnRead = False
    item.Save()
    log(f"  UnRead is now: {item.UnRead}")
    log(f"  Restoring to unread...")
    item.UnRead = True
    item.Save()
    log(f"  UnRead is now: {item.UnRead}")


def test_move_to_archive(namespace):
    """Test 7: Find Archive folder and move oldest inbox email there, then move it back."""
    inbox = namespace.GetDefaultFolder(6)
    root = namespace.DefaultStore.GetRootFolder()

    # Find archive folder
    archive = None
    for i in range(root.Folders.Count):
        folder = root.Folders.Item(i + 1)
        if folder.Name.lower() == "archive":
            archive = folder
            break

    if archive is None:
        log("  Archive folder not found. Creating it...")
        archive = root.Folders.Add("Archive")

    log(f"  Archive folder: {archive.Name} ({archive.Items.Count} items)")

    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    if items.Count == 0:
        log("  SKIP: Inbox is empty")
        return

    # Move the oldest email (last in descending sort) to avoid touching important mail
    item = items.Item(items.Count)
    subject = item.Subject
    log(f"  Moving oldest inbox email: '{subject}' -> Archive")
    moved_item = item.Move(archive)
    log(f"  Archive now has {archive.Items.Count} items")

    # Move it back to inbox
    log(f"  Moving it back to Inbox...")
    moved_item.Move(inbox)
    log(f"  Restored to Inbox")


def test_search(namespace, keyword="test"):
    """Test 8: Search Inbox for emails containing keyword in Subject using DASL."""
    inbox = namespace.GetDefaultFolder(6)
    safe_keyword = keyword.replace("'", "''")
    filter_str = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{safe_keyword}%'"
    results = inbox.Items.Restrict(filter_str)
    log(f"  Found {results.Count} emails matching '{keyword}' in subject")
    if results.Count > 0:
        first = results.Item(1)
        log(f"  First match: {first.Subject}")


def main():
    log("=" * 60)
    log("Outlook Desktop MCP - Phase 1 COM Validation")
    log("=" * 60)
    log("")

    # Test 1 must succeed for all others to run
    log("--- Test 1: Connect to Outlook COM ---")
    try:
        outlook, namespace = test_connect()
        log("  PASS")
    except Exception as e:
        log(f"  FAIL: {e}")
        log("")
        log("Is Outlook Desktop (Classic) running?")
        sys.exit(1)

    tests = [
        ("List Default Folders", lambda: test_list_folders(namespace)),
        ("Read Inbox (top 5)", lambda: test_read_inbox(namespace)),
        ("Filter Unread Emails", lambda: test_filter_unread(namespace)),
        ("Send Test Email", lambda: test_send_email(outlook)),
        ("Mark Read/Unread Cycle", lambda: test_mark_read_unread(namespace)),
        ("Move to Archive & Back", lambda: test_move_to_archive(namespace)),
        ("Search by Subject", lambda: test_search(namespace)),
    ]

    passed = 1  # Test 1 already passed
    total = len(tests) + 1

    for name, fn in tests:
        log(f"\n--- Test {passed + 1}: {name} ---")
        try:
            fn()
            passed += 1
            log("  PASS")
        except Exception as e:
            log(f"  FAIL: {e}")

    log("")
    log("=" * 60)
    log(f"Results: {passed}/{total} passed")
    log("=" * 60)

    import pythoncom
    pythoncom.CoUninitialize()

    sys.exit(0 if passed == total else 1)


if __name__ == "__main__":
    main()
