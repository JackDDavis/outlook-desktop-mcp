# Changelog

All notable changes to this project will be documented in this file.

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
