# Security Policy

## Supported Versions

| Version | Supported          |
|---------|--------------------|
| 0.3.x   | ✅ Current          |
| < 0.3   | ❌ No longer supported |

## Reporting a Vulnerability

This project handles email data through COM automation, so security issues
are taken seriously.

**Please do NOT open a public GitHub issue for security vulnerabilities.**

Instead, report vulnerabilities privately:

1. **GitHub Security Advisory** (preferred): Go to the
   [Security tab](https://github.com/JackDDavis/outlook-desktop-mcp/security/advisories)
   and click "Report a vulnerability."
2. **Email**: Send details to the maintainer via GitHub's private messaging.

### What to include

- Description of the vulnerability
- Steps to reproduce
- Potential impact
- Suggested fix (if any)

### Response timeline

- **Acknowledgement**: Within 48 hours
- **Initial assessment**: Within 1 week
- **Fix or mitigation**: Depends on severity, targeting 2 weeks for critical issues

## Security Architecture

- The MCP server uses Windows COM automation to talk to `OUTLOOK.EXE` — it
  inherits whatever authentication Outlook already has.
- No credentials, API keys, or tokens are stored or transmitted by this project.
- The SSE transport binds to `0.0.0.0` by default (configurable via `--host`).
  For local-only access, use `--host 127.0.0.1`.
- The stdio-to-SSE proxy validates endpoint origins to prevent SSRF.
- The `--deny` flag (server and proxy) can restrict tool access per agent.
- Path traversal protection is applied to attachment save operations.
- COM error details are sanitized before being returned to clients.
