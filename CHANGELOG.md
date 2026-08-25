# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Added
- `outlook_status` for queue-independent Outlook process, COM bridge, account,
  and active-request diagnostics
- Structured error envelopes with stable codes, meanings, likely causes,
  suggested actions, retryability, and sanitized HRESULTs
- Queue wait and execution timing support in the serialized COM bridge
- Local time, timezone, and input interpretation echoes on calendar writes
- Sender filtering on `search_emails`, with DASL on Exchange/MAPI accounts
  and bounded client-side fallback on IMAP/POP accounts
- Standard bulk result rows with requested ID, subject, received time, status,
  structured error, and aggregate ok/failed/skipped summary

### Changed
- COM requests now have configurable 30-second single-call and 90-second bulk
  time budgets
- Queued requests that expire before execution are canceled instead of running
  later with an unknown side effect
- Calendar and non-calendar tool failures use the same MCP-native error shape
- Sender-filter responses report `filter_mode` and whether the 1000-item
  client-side scan window was truncated
- Bulk item failures are re-fetched from Outlook and retried once; repeated
  moves of items already absent from the source report idempotent skips
- Package version bumped to 0.5.0

## [0.4.0] — 2026-08-11

### Added
- `list_calendars` with default-calendar, folder-path, item-count, and
  local-only metadata
- `move_event` for metadata-preserving cross-store calendar moves
- Explicit `calendar` selectors and `allow_local_only` write opt-in
- Optional `include_meta` listing responses with truncation status
- Hermetic fake-COM calendar safety tests

### Changed
- Calendar writes resolve and validate their destination before creating a COM
  item
- `create_event` creates directly in the selected calendar rather than saving
  in the primary store and moving afterward
- Timezone-aware ISO values are converted to Windows local time consistently
- Calendar interval validation is shared across create, meeting, and update
- Calendar failures now surface as MCP-native tool errors
- Calendar list/search limits are explicit and support up to 1000 results
- Account resolution prefers exact display-name/email matches and rejects
  ambiguous substrings
- COM bridge startup now allows 60 seconds, matching operation timeouts and
  avoiding false startup failures while Outlook is busy initializing
- COM bridge startup attaches to the running Outlook COM object before falling
  back to `Dispatch`, avoiding hangs while a visible Outlook session is active
- Package version bumped to 0.4.0

### Fixed
- **`create_event` COM 0x80020009 error** — date strings are now normalized through `_parse_date().strftime()` before being assigned to Outlook COM, accepting ISO 8601 variants (with T/Z separators) that COM would otherwise reject
- **`create_event` cross-account placement** — replaced `CreateItem()` + unsaved `Move()` pattern with `CreateItem()` → set properties → `Save()` → `Move()`; moving an unsaved item caused Outlook to silently land it in the wrong store
- **`_resolve_folder` email-address routing** — folder names containing `@` are now resolved as account names (returning that account's inbox), so `list_emails(folder="user@gmail.com")` works without needing a separate `account` parameter
- **`list_folders` missing account context** — response now includes `"account"` field (store display name) so agents know which mailbox the folder tree belongs to

### Added
- `tests/sse_calendar_test.py` — live SSE integration test for `create_event`, `get_event`, `list_folders` account context, and `delete_event`

## [0.3.0] — 2025-05-16

### Added
- **Multi-account support** — `account` parameter on all tools; `list_accounts` tool
- **Bulk operations** — `bulk_mark_as_read`, `bulk_mark_as_unread`, `bulk_move_emails`, `bulk_read_emails`
- **Drafts** — `create_draft` tool (saves to account-specific Drafts folder)
- **Forward** — `forward_email` tool
- **Efficiency** — `include_body` parameter on list/search tools returns ~300-char body preview inline, eliminating N+1 follow-up calls
- **Agent permission controls** — `--deny` flag on both server and SSE proxy for per-agent tool restrictions
- **Security** — SSRF validation in SSE proxy, sanitized COM error output, path traversal guards
- **COM reliability** — 60-second COM bridge timeout, improved error handling
- SECURITY.md, CODE_OF_CONDUCT.md, CONTRIBUTING.md overhaul, GitHub issue/PR templates

### Changed
- Tool count: 29 → 36
- Version bump to 0.3.0
- pyproject.toml: added AGPL-3.0 license classifier, dev dependencies, upstream link

## [0.2.2] — 2025-04 (pre-fork baseline)

- Initial fork from [Aanerud/outlook-desktop-mcp](https://github.com/Aanerud/outlook-desktop-mcp)
- 29 tools for Windows COM-based Outlook automation
- SSE proxy for WSL-to-Windows bridging
