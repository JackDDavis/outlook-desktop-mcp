#!/usr/bin/env pwsh
# Restart Outlook Desktop MCP server(s)
#
# Run this when experiencing COM connection errors in Claude Code or OpenClaw.
# Outlook Classic (OUTLOOK.EXE) must be running before you use this script.
#
# What this script does:
#   1. Kills any stale outlook-desktop-mcp / python MCP server processes
#   2. Verifies Outlook is running
#   3. Restarts the Task Scheduler SSE server (used by OpenClaw)
#   4. Runs a COM smoke test
#   5. Tells you what to do next in each agent

$ErrorActionPreference = 'Stop'

$scriptDir  = $PSScriptRoot
$python     = Join-Path $scriptDir ".venv\Scripts\python.exe"
$comTest    = Join-Path $scriptDir "tests\phase1_com_test.py"
$taskName   = "\OpenClaw\OutlookMCPServer"

Write-Host ""
Write-Host "=== Outlook MCP Restart ===" -ForegroundColor Cyan

# --- Step 1: Kill stale MCP processes ---
Write-Host ""
Write-Host "[1] Killing stale MCP processes..."
$killed = 0

Get-Process -Name "outlook-desktop-mcp" -ErrorAction SilentlyContinue | ForEach-Object {
    Stop-Process -Id $_.Id -Force
    $killed++
}
Get-Process -Name "python" -ErrorAction SilentlyContinue | Where-Object {
    try { $_.MainModule.FileName -like "*outlook-desktop-mcp*" } catch { $false }
} | ForEach-Object {
    Stop-Process -Id $_.Id -Force
    $killed++
}

if ($killed -gt 0) {
    Write-Host "  Killed $killed stale process(es)." -ForegroundColor Yellow
} else {
    Write-Host "  No stale processes found." -ForegroundColor Gray
}

# --- Step 2: Check Outlook is running ---
Write-Host ""
Write-Host "[2] Checking Outlook..."
$outlook = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
if (-not $outlook) {
    Write-Host ""
    Write-Host "  OUTLOOK.EXE is not running." -ForegroundColor Red
    Write-Host "  -> Open Outlook Classic and wait for it to load the inbox, then re-run this script." -ForegroundColor Yellow
    exit 1
}
Write-Host "  Outlook is running (PID $($outlook.Id))." -ForegroundColor Green

# --- Step 3: Restart Task Scheduler SSE server (for OpenClaw) ---
Write-Host ""
Write-Host "[3] Restarting SSE server (OpenClaw)..."
try {
    schtasks /End /TN $taskName 2>$null | Out-Null
    Start-Sleep -Seconds 1
    $result = schtasks /Run /TN $taskName 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  Task '$taskName' started." -ForegroundColor Green
    } else {
        Write-Host "  Warning: Could not start task: $result" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  Warning: Task Scheduler error: $_" -ForegroundColor Yellow
}

# Give the SSE server a moment to initialize COM
Start-Sleep -Seconds 3

# --- Step 4: COM smoke test ---
Write-Host ""
Write-Host "[4] Running COM smoke test..."
$testOutput = & $python $comTest 2>&1
$exitCode = $LASTEXITCODE

if ($exitCode -eq 0) {
    Write-Host "  COM connection OK." -ForegroundColor Green
    Write-Host ""
    Write-Host "=== Next steps ===" -ForegroundColor Cyan
    Write-Host "  Claude Code : run /mcp to reconnect the MCP server" -ForegroundColor White
    Write-Host "  OpenClaw    : SSE server is running on port 3721 -- restart the agent or reconnect" -ForegroundColor White
} else {
    Write-Host "  COM smoke test FAILED (exit $exitCode)." -ForegroundColor Red
    Write-Host $testOutput
    Write-Host ""
    Write-Host "  -> Close Outlook completely, reopen it, wait for the inbox to fully load, then re-run this script." -ForegroundColor Yellow
    exit 1
}
