# outlook-desktop-mcp

[![PyPI](https://img.shields.io/pypi/v/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Python](https://img.shields.io/pypi/pyversions/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Platform](https://img.shields.io/badge/platform-Windows-blue)]()

**Turn your running Outlook Desktop into an MCP server.** No Microsoft Graph API, no Entra app registration, no OAuth tokens — just your local Outlook and the authentication you already have.

Any MCP client (Claude Code, Claude Desktop, etc.) can then send emails, read your inbox, search messages, manage folders, and more — all through your existing Outlook session.

## Quick Start

**1. Install** (requires Python 3.12+ on Windows):

```bash
pip install outlook-desktop-mcp
```

**2. Register with Claude Code:**

```bash
claude mcp add outlook-desktop -- outlook-desktop-mcp
```

**3. Open Outlook Desktop (Classic) and start a Claude Code session.** That's it. Nine email tools are now available.

## How It Works

```
Claude Code / Claude Desktop / Any MCP Client
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

## Available Tools

All tool descriptions are optimized for LLM tool discovery — Claude understands exactly how to use each one, what arguments to pass, and what to expect back.

| Tool | Description |
|------|-------------|
| `send_email` | Send an email with To/CC/BCC, plain text or HTML body |
| `list_emails` | List recent emails from any folder, with optional unread filter |
| `read_email` | Read full email content by entry ID or subject search |
| `search_emails` | Full-text search across email subjects and bodies |
| `reply_email` | Reply or reply-all, preserving the conversation thread |
| `mark_as_read` | Mark a specific email as read |
| `mark_as_unread` | Mark a specific email as unread |
| `move_email` | Move an email to Archive, Trash, or any folder |
| `list_folders` | Browse the complete folder hierarchy with item counts |

## Install from Source

```bash
git clone https://github.com/Aanerud/outlook-desktop-mcp.git
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
- *"Search my inbox for emails about the budget report"*
- *"Mark that email as read and move it to archive"*
- *"Reply to that email saying I'll join the meeting"*
- *"List all my mail folders"*

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

## Extending

The architecture is designed to grow. The COM bridge is shared across tool modules, so adding new capabilities is straightforward.

**Planned modules:**

- **Calendar** — create events, check availability, manage meetings
- **Contacts** — search address book, resolve recipients
- **Tasks** — manage to-do items

To add a new module, create `@mcp.tool()` functions in `server.py` that call through the existing `bridge.call()` pattern. Contributions welcome.

## Project Structure

```
outlook-desktop-mcp/
  src/outlook_desktop_mcp/
    server.py              # MCP server + all tool definitions
    com_bridge.py          # Async-to-COM threading bridge
    tools/
      _folder_constants.py # Outlook folder enum values
    utils/
      formatting.py        # Email data extraction helpers
      errors.py            # COM error formatting
  tests/
    phase1_com_test.py     # Standalone COM validation
    phase3_mcp_test.py     # MCP protocol integration test
  outlook-desktop-mcp.cmd  # Windows launcher script
  pyproject.toml
```

## License

See [LICENSE](LICENSE) file.
