# Contributing to outlook-desktop-mcp

> This is a fork of [Aanerud/outlook-desktop-mcp](https://github.com/Aanerud/outlook-desktop-mcp)
> with multi-account support, bulk operations, security hardening, and
> agent permission controls. See the upstream repo for the original project.

## How to Contribute

1. **Fork** this repo to your GitHub account
2. **Clone** your fork locally
3. **Create a feature branch** from `main`:
   ```bash
   git checkout main
   git pull origin main
   git checkout -b feature/my-change
   ```
4. **Make your changes**, commit, push to your fork
5. **Open a PR** into `main`
6. PRs are reviewed and merged; releases are published to PyPI from `main`

## Development Setup

Requires Windows with Outlook Desktop (Classic) running.

```bash
git clone https://github.com/<your-username>/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python -m venv .venv
.venv\Scripts\activate
pip install pywin32 "mcp[cli]" -e .
python .venv\Scripts\pywin32_postinstall.py -install
```

### Optional dev tools

```bash
pip install ruff pytest
```

## Testing

With Outlook Desktop (Classic) open:

```bash
# COM validation (no MCP layer)
outlook-desktop-mcp.cmd test

# MCP protocol test
.venv\Scripts\python tests\phase3_mcp_test.py
```

> **Note**: Tests require a running Outlook instance with at least one
> configured account. They perform read-only operations against your mailbox.

## Code Style

- Format and lint with [ruff](https://docs.astral.sh/ruff/) (no config needed — defaults are fine)
- All public tool functions need detailed docstrings — LLMs use them for tool discovery
- Keep COM operations in inner `def` functions passed to `bridge.call()`
- Log to stderr only (`stdout` is reserved for MCP JSON-RPC)

## Adding New Tools

1. Define the COM function that does the work (receives `outlook, namespace` as first args)
2. Add an `@mcp.tool()` async handler in `server.py` that calls `bridge.call(your_function, ...)`
3. Write a detailed docstring — this is what LLMs see during tool discovery
4. Add the `account: str = ""` parameter for multi-account support
5. Add test coverage
6. Update the `instructions` string if adding a new capability category
7. Update the tool table in `README.md`

## Security

- Never log or return raw COM exception details to clients (use `format_com_error`)
- Validate file paths in tools that write to disk (see `save_attachment`)
- See [SECURITY.md](SECURITY.md) for the vulnerability reporting policy
