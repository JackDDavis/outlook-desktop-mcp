# Message-ID capability spike — 2026-08-25

The probe enumerated six Outlook stores without recording account names,
addresses, subjects, bodies, or Message-ID values.

## Results

| Store | Account type | Mail sampled | Message-ID coverage | Sample truncated |
|---|---:|---:|---:|---|
| 1 | Exchange | 702 | 100% | no |
| 2 | Exchange | 604 | 100% | no |
| 3 | No account mapping | 0 | n/a | no |
| 4 | IMAP | 977 | 100% | yes, inbox exceeded 1000 items |
| 5 | IMAP | 668 | 100% | no |
| 6 | IMAP | 989 | 100% | yes, inbox exceeded 1000 items |

`MAPIFolder.FindItemMessageID` is not exposed by Outlook COM on any connected
store. The implementation must not depend on that method.

An exact folder-scoped DASL restriction on `PR_INTERNET_MESSAGE_ID`
(`0x1035001F`) resolved a known item in every non-empty store. Controlled
move-and-restore checks passed on all three stores that expose an Archive
folder, covering both Exchange and IMAP. The other two non-empty IMAP stores
passed lookup in place but do not expose an Archive folder.

## Implementation decision

Resolve `message_id` with an exact DASL `Items.Restrict` query in the selected
folder. Continue using `GetItemFromID` for Outlook EntryIDs. Report
`message_id: null` and `id_stable: false` when the property is absent.

The two inboxes above 1000 items require an unbounded or paged release
acceptance probe before claiming exhaustive per-store coverage. The bounded
spike establishes 100% coverage in the sampled window, not the entire inbox.
