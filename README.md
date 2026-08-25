# outlook-desktop-mcp

[![PyPI](https://img.shields.io/pypi/v/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Python](https://img.shields.io/pypi/pyversions/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Platform](https://img.shields.io/badge/platform-Windows-blue)]()
[![License: AGPL-3.0](https://img.shields.io/badge/license-AGPL--3.0-blue)](LICENSE)

> **Fork of [Aanerud/outlook-desktop-mcp](https://github.com/Aanerud/outlook-desktop-mcp)** with
> multi-account support, bulk operations, agent permission controls (`--deny`), and efficiency
> improvements (`include_body`). See upstream for macOS support.

**Turn your running Outlook Desktop into an MCP server with 39 tools — including full multi-account support.** No Microsoft Graph API, no Entra app registration, no OAuth tokens — just your local Outlook and the authentication you already have.

Works with Claude Code, Claude Desktop, GitHub Copilot, OpenAI Codex, OpenClaw, and any MCP-compatible agent. Send emails, manage your calendar, create tasks, handle attachments, and more — across every email account configured in Outlook, with a single `account` parameter to target any of them.

## Quick Start

**1. Install** (requires Python 3.12+ on Windows):

```bash
pip install outlook-desktop-mcp
```

**2. Register with your agent:**

<details>
<summary>Claude Code</summary>

```bash
claude mcp add outlook-desktop -- outlook-desktop-mcp
```
</details>

<details>
<summary>Claude Desktop</summary>

Add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook-desktop": {
      "command": "outlook-desktop-mcp"
    }
  }
}
```
</details>

<details>
<summary>GitHub Copilot (VS Code)</summary>

Add to `.vscode/mcp.json` or user MCP settings:

```json
{
  "servers": {
    "outlook-desktop": {
      "type": "stdio",
      "command": "outlook-desktop-mcp"
    }
  }
}
```
</details>

<details>
<summary>OpenAI Codex</summary>

```bash
codex mcp add --name outlook-desktop --command outlook-desktop-mcp
```
</details>

<details>
<summary>OpenClaw (SSE mode — recommended)</summary>

Start the server in SSE mode, then add to your OpenClaw MCP config:

```bash
outlook-desktop-mcp --http --host 127.0.0.1 --port 3721
```

```json
{
  "mcpServers": {
    "outlook-desktop": {
      "type": "sse",
      "url": "http://127.0.0.1:3721/sse"
    }
  }
}
```
</details>

**3. Open Outlook Desktop (Classic) and start a session.** That's it — 39 tools are available immediately.

## Alternative Transport: stdio → SSE Proxy (recommended for elevated/WSL edge cases)

If direct stdio launch of `outlook-desktop-mcp.exe` fails in your client context (for example, elevated agent process, Windows handle quirks, or WSL routing), run the server as SSE on Windows and connect through the included stdio proxy:

- `stdio_sse_proxy.py` keeps MCP client transport as stdio
- The proxy forwards requests to the running SSE endpoint (`http://127.0.0.1:3721/sse`)
- Endpoint origin validation prevents the SSE server from redirecting requests externally

Example Claude MCP entry:

```json
{
  "mcpServers": {
    "outlook-desktop": {
      "type": "stdio",
      "command": "C:\\tools\\outlook-desktop-mcp\\.venv\\Scripts\\python.exe",
      "args": [
        "C:\\tools\\outlook-desktop-mcp\\stdio_sse_proxy.py"
      ]
    }
  }
}
```

Start the Outlook server in SSE mode:

```bash
outlook-desktop-mcp --http --host 0.0.0.0 --port 3721
```

This route preserves full tool functionality (`list_folders`, `list_emails`, etc.) while avoiding direct COM startup issues in some client runtimes.

## Tool Restrictions (--deny)

Both the server and the proxy support `--deny` to remove tools, enabling flexible permission levels.

### Per-agent restrictions (proxy-level, recommended)

When multiple agents share one SSE server, use `--deny` on the **proxy** to give each agent different permissions. The proxy strips denied tools from the manifest and rejects denied tool calls before they reach the server.

```json
{
  "mcpServers": {
    "outlook-desktop": {
      "type": "stdio",
      "command": "python",
      "args": [
        "stdio_sse_proxy.py",
        "--deny", "send_email,reply_email"
      ]
    }
  }
}
```

Common presets:
- **Compose-only** (can draft, can't send): `--deny send_email,reply_email`
- **Read-only**: `--deny send_email,reply_email,create_draft,move_email,bulk_move_emails,...`
- **Full access**: omit `--deny`

### Global restrictions (server-level)

Use `--deny` on the **server** to restrict all clients universally. Tool names are validated at startup — typos cause an immediate error.

```bash
outlook-desktop-mcp --http --deny send_email,reply_email
```

## How It Works

```
Claude Code / Claude Desktop / GitHub Copilot / Codex / OpenClaw / Any MCP Client
    |
    | stdio (JSON-RPC)
    v
outlook-desktop-mcp (Python)
    |
    | COM automation via Outlook Object Model (MSOUTL.OLB)
    v
Outlook Desktop (Classic) — OUTLOOK.EXE
    |
    | Your existing authenticated session
    v
Exchange Online / Microsoft 365 / On-Premises Exchange
```

The server uses Windows COM automation to talk directly to the running `OUTLOOK.EXE` process. It inherits whatever authentication Outlook already has — your M365 account, on-prem Exchange, or even personal Outlook.com accounts. No additional credentials or API keys are needed.

Internally, the server runs a dedicated COM thread (Single-Threaded Apartment) that holds the `Outlook.Application` object. The async MCP event loop dispatches tool calls to this thread via a queue, keeping COM threading rules respected and the MCP protocol non-blocking.

## Requirements

- **Windows** — COM automation is Windows-only
- **Outlook Desktop (Classic)** — the `OUTLOOK.EXE` that comes with Microsoft 365 / Office. The new "modern" Outlook (`olk.exe`) does **not** support COM
- **Python 3.12+**
- **Outlook must be running** when the MCP server starts

## Available Tools (39)

All tool descriptions are optimized for LLM tool discovery — Claude understands exactly how to use each one, what arguments to pass, and what to expect back.

Most tools accept an optional `account` parameter to target a specific Outlook account (e.g., `"work@company.com"`). If omitted, the default account is used.

Calendar tools also accept an optional `calendar` name/path where relevant.
Call `list_calendars` before writing when synchronization matters. Outlook
folders named `Calendar (This computer only)` are local-only and do not sync to
the provider; writes to them are blocked unless `allow_local_only=true` is
explicitly supplied.

List tools (`list_emails`, `search_emails`, `list_events`, `search_events`, `list_tasks`) accept an optional `include_body` parameter. When true, results include a ~300 char body preview, To/CC, and categories inline — eliminating the need for follow-up detail calls during triage.

### Health and accounts (2 tools)

| Tool | Description |
|------|-------------|
| `outlook_status` | Return immediate Outlook process, COM bridge, queue, active-request, and cached per-account health |
| `list_accounts` | List all accounts configured in Outlook with display name and email |

### Email (15 tools)

| Tool | Description |
|------|-------------|
| `send_email` | Send an email with To/CC/BCC, plain text or HTML body |
| `create_draft` | Create a draft email in Drafts without sending (for user review) |
| `list_emails` | List recent emails from any folder, with optional body preview |
| `read_email` | Read full email content by entry ID or subject search |
| `search_emails` | Full-text search across email subjects and bodies |
| `reply_email` | Reply or reply-all, preserving the conversation thread |
| `forward_email` | Forward an email with attachments to new recipients |
| `mark_as_read` | Mark a specific email as read |
| `mark_as_unread` | Mark a specific email as unread |
| `move_email` | Move an email to Archive, Trash, or any folder |
| `list_folders` | Browse the complete folder hierarchy with item counts |
| `bulk_mark_as_read` | Mark multiple emails as read in a single call |
| `bulk_mark_as_unread` | Mark multiple emails as unread in a single call |
| `bulk_move_emails` | Move multiple emails to a folder in a single call |
| `bulk_read_emails` | Read full content of multiple emails in a single call |

### Calendar (10 tools)

| Tool | Description |
|------|-------------|
| `list_calendars` | List calendar folders with default and local-only status |
| `list_events` | List upcoming events with recurring occurrence support |
| `get_event` | Read full event details by entry ID |
| `create_event` | Create a personal appointment directly in a selected calendar |
| `create_meeting` | Create a meeting and send invitations to attendees |
| `update_event` | Modify an existing event's subject, time, location, etc. |
| `delete_event` | Delete an appointment or cancel a meeting (sends notices) |
| `move_event` | Move an existing event between account/calendar stores without recreating it |
| `respond_to_meeting` | Accept, decline, or tentatively accept a meeting invite |
| `search_events` | Search calendar events by keyword within a date range |

#### Calendar safety and time semantics

- Omitting `account` continues to target Outlook's primary account.
- `create_event` resolves the account and calendar before creating anything,
  then creates directly in that folder. It no longer creates in the primary
  calendar and moves afterward.
- Local-only calendar writes require `allow_local_only=true`.
- Naive ISO values are interpreted as the Windows host's local Outlook time.
  Offset/Z values are converted to host-local time before COM assignment.
- Calendar write responses echo `start_local`, `end_local`, `timezone`, and
  `interpreted_as` so callers can verify Outlook's wall-clock interpretation.
- Event end must be later than start. All-day values must be aligned to
  midnight; use the next date as the exclusive end for a one-day event.
- Calendar validation and COM failures are MCP tool errors (`isError=true`),
  not successful strings beginning with `Error`.
- `list_events` and `search_events` accept `count` from 1 through 1000. Set
  `include_meta=true` to receive `{events, count, truncated}` instead of a
  bare array.

### Reliability and recovery

Tool failures return a structured error with `code`, `meaning`,
`likely_cause`, `suggested_action`, and `retryable`. Known Outlook HRESULTs
also include a sanitized `hresult`; raw COM internals remain in server logs.

`outlook_status()` does not wait behind queued COM work. When Outlook is busy,
it immediately reports the active request, queue depth, and age of the last
successful COM call. It submits a five-second live COM probe only when the
bridge is idle.

For `rpc_unavailable` (`0x800706BA`):

1. Call `outlook_status`.
2. If `com_responsive` is false, reconnect the MCP server.
3. If the failure recurs, use `restart.ps1` or restart Outlook Desktop.

Parallel tool calls are safe: the server serializes Outlook COM access. A
response may include native `_meta.queue_wait_ms` and `_meta.execution_ms`;
queue waits over ten seconds include an Outlook-busy note.

Example:

```json
{
  "account": "work@example.com",
  "calendar": "Calendar/Projects",
  "subject": "Planning",
  "start": "2026-08-12T09:00:00-04:00",
  "end": "2026-08-12T10:00:00-04:00"
}
```

### Tasks (5 tools)

| Tool | Description |
|------|-------------|
| `list_tasks` | List pending or completed tasks, sorted by due date |
| `get_task` | Read full task details including body and completion status |
| `create_task` | Create a new task with subject, due date, importance |
| `complete_task` | Mark a task as complete (100%) |
| `delete_task` | Remove a task |

### Attachments (2 tools)

| Tool | Description |
|------|-------------|
| `list_attachments` | List all attachments on an email or calendar event |
| `save_attachment` | Download an attachment to a local directory |

### Categories (2 tools)

| Tool | Description |
|------|-------------|
| `list_categories` | List all available color categories in Outlook |
| `set_category` | Set or clear categories on any email, event, or task |

### Rules (2 tools)

| Tool | Description |
|------|-------------|
| `list_rules` | List all mail rules with enabled/disabled status |
| `toggle_rule` | Enable or disable a mail rule by name |

### Out of Office (1 tool)

| Tool | Description |
|------|-------------|
| `get_out_of_office` | Check whether Out of Office auto-reply is on or off |

## Install from Source

```bash
git clone https://github.com/JackDDavis/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python -m venv .venv
.venv\Scripts\activate
pip install pywin32 "mcp[cli]" -e .
python .venv\Scripts\pywin32_postinstall.py -install
```

Register from source using the launcher script:

```bash
claude mcp add outlook-desktop -- powershell.exe -Command "& 'C:\path\to\outlook-desktop-mcp\outlook-desktop-mcp.cmd' mcp"
```

## Usage Examples

Once registered, just talk to Claude naturally:

- *"Show me my 10 most recent inbox emails"*
- *"Read the email from Taylor about MLADS"*
- *"Send an email to alice@example.com about the project update"*
- *"What's on my calendar this week?"*
- *"Create a meeting with bob@example.com tomorrow at 2pm for 30 minutes"*
- *"Save the attachment from that email to my Downloads folder"*
- *"Create a task to review the quarterly report, due Friday, high importance"*
- *"Mark that email as read and move it to archive"*
- *"What categories do I have? Set this email to 'Follow-up'"*
- *"List my mail rules"*
- *"Am I set as Out of Office?"*
- *"What accounts do I have in Outlook?"*
- *"Show unread emails from my work account"*
- *"Mark all those emails as read"*

## Why Not Microsoft Graph?

| | Microsoft Graph | outlook-desktop-mcp |
|---|---|---|
| Entra app registration | Required | Not needed |
| Admin consent | Required for mail permissions | Not needed |
| OAuth token management | You handle refresh tokens | Not needed |
| Tenant configuration | Required | Not needed |
| Works offline / cached | No | Yes (reads from OST cache) |
| Setup time | 30-60 minutes | 2 minutes |
| Auth requirement | **Your own OAuth flow** | **Outlook is open** |

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for the branching strategy and development setup.

## Project Structure

```
outlook-desktop-mcp/
  src/outlook_desktop_mcp/
    server.py              # MCP server + tool definitions
    com_bridge.py          # Observable serialized COM bridge
    tools/
      _folder_constants.py # Outlook enums and constants
    utils/
      formatting.py        # Email, event, and task data extraction
      errors.py            # Structured, sanitized diagnostics
      responses.py         # MCP-native response and metadata helpers
  tests/
    phase1_com_test.py     # Email COM validation
    phase3_mcp_test.py     # Email MCP test
    calendar_com_test.py   # Calendar COM validation
    calendar_mcp_test.py   # Calendar MCP test
    extras_com_test.py     # Tasks/attachments/categories/rules/OOF COM test
    extras_mcp_test.py     # Tasks/attachments/categories/rules/OOF MCP test
  outlook-desktop-mcp.cmd  # Windows launcher script
  stdio_sse_proxy.py       # WSL stdio-to-SSE proxy (with endpoint validation)
  pyproject.toml
```

## License

See [LICENSE](LICENSE) file.
