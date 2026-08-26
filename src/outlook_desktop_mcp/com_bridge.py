"""
COM Threading Bridge
====================
Runs all Outlook COM calls on a dedicated STA (Single-Threaded Apartment)
thread so the async MCP event loop never touches COM objects directly.
"""

from __future__ import annotations

import asyncio
import logging
import os
import queue
import threading
import time
import uuid
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import UTC, datetime
from typing import Any

logger = logging.getLogger("outlook_desktop_mcp.com_bridge")

DEFAULT_SINGLE_TIMEOUT_SECONDS = 30.0
DEFAULT_BULK_TIMEOUT_SECONDS = 90.0


def _positive_env_float(name: str, default: float) -> float:
    raw = os.environ.get(name)
    if raw is None:
        return default
    try:
        value = float(raw)
    except ValueError as error:
        raise ValueError(f"{name} must be a positive number") from error
    if value <= 0:
        raise ValueError(f"{name} must be a positive number")
    return value


def _utc_now() -> str:
    return datetime.now(UTC).isoformat()


@dataclass
class BridgeRequest:
    name: str
    function: Callable[..., Any]
    args: tuple[Any, ...]
    kwargs: dict[str, Any]
    timeout_seconds: float
    enqueued_monotonic: float = field(default_factory=time.monotonic)
    enqueued_at: str = field(default_factory=_utc_now)
    started_monotonic: float | None = None
    started_at: str | None = None
    completed_monotonic: float | None = None
    completed_at: str | None = None
    result: Any = None
    error: Exception | None = None
    cancelled: bool = False
    caller_timed_out: bool = False
    pending: bool = False
    state_lock: threading.Lock = field(default_factory=threading.Lock)
    completion_event: threading.Event = field(default_factory=threading.Event)

    @property
    def queue_wait_ms(self) -> int:
        end = self.started_monotonic or time.monotonic()
        return max(0, round((end - self.enqueued_monotonic) * 1000))

    @property
    def execution_ms(self) -> int | None:
        if self.started_monotonic is None:
            return None
        end = self.completed_monotonic or time.monotonic()
        return max(0, round((end - self.started_monotonic) * 1000))


@dataclass(frozen=True)
class BridgeCallResult:
    value: Any
    queue_wait_ms: int
    execution_ms: int
    request_name: str


class BridgeTimeoutError(TimeoutError):
    def __init__(self, request_name: str, timeout_seconds: float, phase: str):
        self.request_name = request_name
        self.timeout_seconds = timeout_seconds
        self.phase = phase
        super().__init__(
            f"Outlook COM request '{request_name}' timed out during {phase} "
            f"after {timeout_seconds:g} seconds"
        )


class OutlookBridge:
    """Manage serialized Outlook COM access and observable bridge state."""

    def __init__(self):
        self._thread: threading.Thread | None = None
        self._request_queue: queue.Queue[BridgeRequest | None] = queue.Queue()
        self._outlook = None
        self._namespace = None
        self._ready = threading.Event()
        self._shutdown = threading.Event()
        self._init_error: Exception | None = None
        self._state_lock = threading.Lock()
        self._active_request: BridgeRequest | None = None
        self._pending_count = 0
        self._last_success_at: str | None = None
        self._last_success_monotonic: float | None = None
        self._last_failure_at: str | None = None
        self._last_failure: dict[str, Any] | None = None
        self._accounts_snapshot: list[dict[str, Any]] = []
        self._accounts_snapshot_at: str | None = None
        self._accounts_snapshot_monotonic: float | None = None
        self._process_instance_id = str(uuid.uuid4())
        self.single_timeout_seconds = _positive_env_float(
            "MCP_SINGLE_CALL_TIMEOUT_SECONDS",
            DEFAULT_SINGLE_TIMEOUT_SECONDS,
        )
        self.bulk_timeout_seconds = _positive_env_float(
            "MCP_BULK_CALL_TIMEOUT_SECONDS",
            DEFAULT_BULK_TIMEOUT_SECONDS,
        )

    def start(self, timeout: float = 60, *, wait_until_ready: bool = True):
        """Start the COM thread, optionally without blocking server startup."""
        self._thread = threading.Thread(
            target=self._com_thread_main,
            daemon=True,
            name="outlook-com",
        )
        self._thread.start()
        if not wait_until_ready:
            return
        if not self._ready.wait(timeout=timeout):
            if self._init_error:
                raise self._init_error
            raise RuntimeError(
                f"Outlook COM thread failed to initialize within {timeout}s. "
                "Is Outlook Desktop (Classic) running?"
            )
        if self._init_error:
            raise self._init_error

    def _initialize_outlook(self):
        import pythoncom
        import win32com.client

        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
            logger.debug("Attached to the running Outlook COM instance")
        except pythoncom.com_error:
            outlook = win32com.client.Dispatch("Outlook.Application")
            logger.debug("Started a new Outlook COM automation instance")
        namespace = outlook.GetNamespace("MAPI")
        return outlook, namespace

    def _com_thread_main(self):
        import pythoncom

        pythoncom.CoInitialize()
        try:
            self._outlook, self._namespace = self._initialize_outlook()
            store_name = self._namespace.DefaultStore.DisplayName
            user_name = self._namespace.CurrentUser.Name
            logger.debug(
                "COM thread ready. Store: %s, User: %s",
                store_name,
                user_name,
            )
            self._ready.set()

            while not self._shutdown.is_set():
                try:
                    request = self._request_queue.get(timeout=0.5)
                except queue.Empty:
                    continue
                if request is None:
                    break

                with request.state_lock:
                    if request.cancelled:
                        self._release_pending(request)
                        request.completed_monotonic = time.monotonic()
                        request.completed_at = _utc_now()
                        request.completion_event.set()
                        continue
                    request.started_monotonic = time.monotonic()
                    request.started_at = _utc_now()
                    with self._state_lock:
                        if request.pending:
                            self._pending_count -= 1
                            request.pending = False
                        self._active_request = request

                try:
                    request.result = request.function(
                        self._outlook,
                        self._namespace,
                        *request.args,
                        **request.kwargs,
                    )
                    with self._state_lock:
                        self._last_success_at = _utc_now()
                        self._last_success_monotonic = time.monotonic()
                except Exception as error:  # noqa: BLE001 - isolate COM failures per request
                    request.error = error
                    self._record_failure(error)
                finally:
                    request.completed_monotonic = time.monotonic()
                    request.completed_at = _utc_now()
                    with self._state_lock:
                        self._active_request = None
                    request.completion_event.set()
        except Exception as error:  # noqa: BLE001 - report COM initialization failures
            self._init_error = error
            self._record_failure(error)
            self._ready.set()
            logger.error("COM thread init failed: %s", error)
        finally:
            pythoncom.CoUninitialize()

    def _record_failure(self, error: Exception):
        from outlook_desktop_mcp.utils.errors import error_details

        with self._state_lock:
            self._last_failure_at = _utc_now()
            self._last_failure = error_details(error)

    def health_snapshot(self) -> dict[str, Any]:
        """Return current bridge state without enqueueing a COM request."""
        now = time.monotonic()
        with self._state_lock:
            active = self._active_request
            last_success_age_ms = (
                round((now - self._last_success_monotonic) * 1000)
                if self._last_success_monotonic is not None
                else None
            )
            accounts_age_ms = (
                round((now - self._accounts_snapshot_monotonic) * 1000)
                if self._accounts_snapshot_monotonic is not None
                else None
            )
            return {
                "process_instance_id": self._process_instance_id,
                "thread_alive": bool(self._thread and self._thread.is_alive()),
                "queue_depth": self._pending_count,
                "active_request": (
                    {
                        "name": active.name,
                        "started_at": active.started_at,
                        "elapsed_ms": active.execution_ms,
                        "caller_timed_out": active.caller_timed_out,
                    }
                    if active
                    else None
                ),
                "last_success_at": self._last_success_at,
                "last_success_age_ms": last_success_age_ms,
                "last_failure_at": self._last_failure_at,
                "last_failure": dict(self._last_failure) if self._last_failure else None,
                "accounts": [dict(account) for account in self._accounts_snapshot],
                "accounts_snapshot_at": self._accounts_snapshot_at,
                "accounts_snapshot_age_ms": accounts_age_ms,
            }

    def update_accounts_snapshot(self, accounts: list[dict[str, Any]]):
        """Update cached account health data from a successful COM call."""
        with self._state_lock:
            self._accounts_snapshot = [dict(account) for account in accounts]
            self._accounts_snapshot_at = _utc_now()
            self._accounts_snapshot_monotonic = time.monotonic()

    def is_idle(self) -> bool:
        with self._state_lock:
            return self._active_request is None and self._pending_count == 0

    def _enqueue(self, request: BridgeRequest, *, only_if_idle: bool = False) -> bool:
        with self._state_lock:
            if not only_if_idle:
                if self._init_error is not None:
                    raise self._init_error
                if self._thread is not None and not self._thread.is_alive():
                    raise RuntimeError("Outlook COM bridge is not running")
            if only_if_idle and (
                not self._thread
                or not self._thread.is_alive()
                or not self._ready.is_set()
                or self._init_error is not None
                or self._active_request is not None
                or self._pending_count > 0
            ):
                return False
            self._pending_count += 1
            request.pending = True
            self._request_queue.put(request)
            return True

    def _release_pending(self, request: BridgeRequest):
        with self._state_lock:
            if request.pending:
                self._pending_count -= 1
                request.pending = False

    async def _wait_for_request(self, request: BridgeRequest) -> BridgeCallResult:
        signaled = await asyncio.to_thread(
            request.completion_event.wait,
            request.timeout_seconds,
        )
        if not signaled:
            with request.state_lock:
                if request.started_monotonic is None:
                    request.cancelled = True
                    self._release_pending(request)
                    phase = "queue"
                else:
                    request.caller_timed_out = True
                    phase = "execution"
            raise BridgeTimeoutError(
                request.name,
                request.timeout_seconds,
                phase,
            )

        if request.cancelled:
            raise BridgeTimeoutError(
                request.name,
                request.timeout_seconds,
                "queue",
            )
        if request.error is not None:
            raise request.error
        return BridgeCallResult(
            value=request.result,
            queue_wait_ms=request.queue_wait_ms,
            execution_ms=request.execution_ms or 0,
            request_name=request.name,
        )

    async def call_with_metrics(
        self,
        function,
        *args,
        timeout_seconds: float | None = None,
        request_name: str | None = None,
        **kwargs,
    ) -> BridgeCallResult:
        """Run a COM function and return its value plus queue/execution timing."""
        timeout = (
            self.single_timeout_seconds
            if timeout_seconds is None
            else timeout_seconds
        )
        if timeout <= 0:
            raise ValueError("timeout_seconds must be positive")

        request = BridgeRequest(
            name=request_name or getattr(function, "__name__", "outlook_call"),
            function=function,
            args=args,
            kwargs=kwargs,
            timeout_seconds=timeout,
        )
        self._enqueue(request)
        return await self._wait_for_request(request)

    async def call_if_idle_with_metrics(
        self,
        function,
        *args,
        timeout_seconds: float,
        request_name: str,
        **kwargs,
    ) -> BridgeCallResult | None:
        """Run a probe only if no COM request is active or queued."""
        if timeout_seconds <= 0:
            raise ValueError("timeout_seconds must be positive")
        request = BridgeRequest(
            name=request_name,
            function=function,
            args=args,
            kwargs=kwargs,
            timeout_seconds=timeout_seconds,
        )
        if not self._enqueue(request, only_if_idle=True):
            return None
        return await self._wait_for_request(request)

    async def call(
        self,
        function,
        *args,
        timeout_seconds: float | None = None,
        request_name: str | None = None,
        **kwargs,
    ):
        """Run a COM function and preserve legacy values while adding native metadata."""
        result = await self.call_with_metrics(
            function,
            *args,
            timeout_seconds=timeout_seconds,
            request_name=request_name,
            **kwargs,
        )
        try:
            from mcp.types import CallToolResult

            from outlook_desktop_mcp.utils.responses import with_meta

            if isinstance(result.value, CallToolResult):
                meta = {
                    "queue_wait_ms": result.queue_wait_ms,
                    "execution_ms": result.execution_ms,
                }
                if result.queue_wait_ms > 10_000:
                    meta["note"] = (
                        "Outlook is busy; prior operations were still running"
                    )
                return with_meta(result.value, **meta)
        except ImportError:
            pass
        return result.value

    def stop(self):
        """Signal the COM thread to shut down."""
        self._shutdown.set()
        self._request_queue.put(None)
        if self._thread:
            self._thread.join(timeout=5)
