"""Outlook OlDefaultFolders enum values and friendly-name mapping."""

OL_FOLDER_INBOX = 6
OL_FOLDER_SENT_MAIL = 5
OL_FOLDER_DELETED_ITEMS = 3
OL_FOLDER_DRAFTS = 16
OL_FOLDER_CALENDAR = 9
OL_FOLDER_TASKS = 13
OL_FOLDER_JUNK = 23
OL_FOLDER_OUTBOX = 4

FOLDER_NAME_TO_ENUM = {
    "inbox": OL_FOLDER_INBOX,
    "sent": OL_FOLDER_SENT_MAIL,
    "sentmail": OL_FOLDER_SENT_MAIL,
    "deleted": OL_FOLDER_DELETED_ITEMS,
    "trash": OL_FOLDER_DELETED_ITEMS,
    "drafts": OL_FOLDER_DRAFTS,
    "calendar": OL_FOLDER_CALENDAR,
    "tasks": OL_FOLDER_TASKS,
    "junk": OL_FOLDER_JUNK,
    "spam": OL_FOLDER_JUNK,
    "outbox": OL_FOLDER_OUTBOX,
}

OL_MAIL_ITEM = 0
