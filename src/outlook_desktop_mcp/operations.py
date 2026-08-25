"""Durable state for bounded, pollable Outlook operations."""

from __future__ import annotations

import json
import os
import threading
import uuid
from collections.abc import Callable
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any

TERMINAL_STATUSES = {"complete", "error", "interrupted"}
INTERRUPTED_GUIDANCE = (
    "Remaining work did not execute — verify the folder and re-run; "
    "idempotent skip makes re-run safe."
)
NOT_FOUND_GUIDANCE = (
    "Operation state is unknown. Verify the folder before re-running; the "
    "operation may have expired or its snapshot may be unreadable."
)


def _default_storage_dir() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / "outlook-desktop-mcp" / "operations"
    return Path.home() / "AppData" / "Local" / "outlook-desktop-mcp" / "operations"


def _summary(rows: list[dict[str, Any]]) -> dict[str, int]:
    return {
        "total": len(rows),
        "ok": sum(row.get("status") == "ok" for row in rows),
        "failed": sum(row.get("status") == "failed" for row in rows),
        "skipped": sum(row.get("status") == "skipped" for row in rows),
    }


class OperationManager:
    """Thread-safe operation records backed by atomic JSON snapshots."""

    def __init__(
        self,
        storage_dir: str | Path | None = None,
        *,
        ttl_hours: float = 72,
        process_instance_id: str | None = None,
        now: Callable[[], datetime] | None = None,
    ):
        self.storage_dir = Path(storage_dir) if storage_dir else _default_storage_dir()
        self.ttl = timedelta(hours=ttl_hours)
        self.process_instance_id = process_instance_id or str(uuid.uuid4())
        self._now = now or (lambda: datetime.now(UTC))
        self._lock = threading.RLock()
        self._records: dict[str, dict[str, Any]] = {}
        self._events: dict[str, threading.Event] = {}
        self._threads: dict[str, threading.Thread] = {}
        self._closing = threading.Event()
        self.storage_dir.mkdir(parents=True, exist_ok=True)
        self._load()

    def _timestamp(self) -> str:
        return self._now().astimezone(UTC).isoformat()

    def _is_expired(self, record: dict[str, Any]) -> bool:
        try:
            expires_at = datetime.fromisoformat(record["expires_at"])
        except (KeyError, TypeError, ValueError):
            return True
        return expires_at <= self._now().astimezone(UTC)

    def _touch(self, record: dict[str, Any]) -> None:
        now = self._now().astimezone(UTC)
        record["updated_at"] = now.isoformat()
        record["expires_at"] = (now + self.ttl).isoformat()
        record["summary"] = _summary(record["results"])
        record["processed_count"] = len(record["results"])
        record["success_count"] = record["summary"]["ok"]
        record["failure_count"] = record["summary"]["failed"]

    def _snapshot_path(self, operation_id: str) -> Path:
        return self.storage_dir / f"{operation_id}.json"

    def _persist(self, record: dict[str, Any]) -> None:
        path = self._snapshot_path(record["operation_id"])
        temp_path = path.with_name(
            f".{path.name}.{uuid.uuid4().hex}.tmp"
        )
        try:
            with temp_path.open("w", encoding="utf-8", newline="\n") as stream:
                json.dump(record, stream, indent=2, sort_keys=True, default=str)
                stream.write("\n")
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temp_path, path)
        finally:
            try:
                temp_path.unlink()
            except FileNotFoundError:
                pass

    def _quarantine(self, path: Path) -> None:
        quarantine = path.with_name(
            f"{path.name}.corrupt-{uuid.uuid4().hex}"
        )
        try:
            os.replace(path, quarantine)
        except OSError:
            pass

    def _load(self) -> None:
        for path in self.storage_dir.glob("*.json"):
            try:
                with path.open("r", encoding="utf-8") as stream:
                    record = json.load(stream)
                if not isinstance(record, dict):
                    raise ValueError("operation snapshot must be an object")
                operation_id = record["operation_id"]
                if not isinstance(operation_id, str) or not operation_id:
                    raise ValueError("operation snapshot has no operation_id")
                if self._is_expired(record):
                    try:
                        path.unlink()
                    except OSError:
                        pass
                    continue
                record.setdefault("results", [])
                record.setdefault("matched_count", 0)
                record.setdefault("remaining", 0)
                status = record.get("status")
                if (
                    status not in TERMINAL_STATUSES
                    and record.get("process_instance_id")
                    != self.process_instance_id
                ):
                    record["status"] = "interrupted"
                    record["guidance"] = INTERRUPTED_GUIDANCE
                    self._touch(record)
                    self._persist(record)
                self._records[operation_id] = record
                event = threading.Event()
                if record.get("status") in TERMINAL_STATUSES:
                    event.set()
                self._events[operation_id] = event
            except (OSError, ValueError, TypeError, KeyError, json.JSONDecodeError):
                self._quarantine(path)

    def create(self, kind: str, *, remaining: int = 0) -> str:
        operation_id = str(uuid.uuid4())
        created_at = self._timestamp()
        record = {
            "operation_id": operation_id,
            "kind": kind,
            "status": "in_progress",
            "process_instance_id": self.process_instance_id,
            "created_at": created_at,
            "updated_at": created_at,
            "expires_at": created_at,
            "results": [],
            "summary": _summary([]),
            "matched_count": 0,
            "processed_count": 0,
            "success_count": 0,
            "failure_count": 0,
            "remaining": max(0, int(remaining)),
            "error": None,
        }
        with self._lock:
            self._touch(record)
            self._records[operation_id] = record
            self._events[operation_id] = threading.Event()
            self._persist(record)
        return operation_id

    def start(self, operation_id: str, runner: Callable[[str], None]) -> None:
        def run() -> None:
            try:
                runner(operation_id)
            except Exception as error:  # noqa: BLE001 - durable worker boundary
                self.fail(operation_id, error)
            finally:
                with self._lock:
                    self._threads.pop(operation_id, None)

        with self._lock:
            if self._closing.is_set():
                self.interrupt(operation_id)
                return
            thread = threading.Thread(
                target=run,
                daemon=True,
                name=f"outlook-operation-{operation_id[:8]}",
            )
            self._threads[operation_id] = thread
            thread.start()

    def should_stop(self) -> bool:
        return self._closing.is_set()

    def append_results(
        self,
        operation_id: str,
        rows: list[dict[str, Any]],
        *,
        matched_count: int,
        remaining: int,
    ) -> None:
        with self._lock:
            record = self._records.get(operation_id)
            if not record or record["status"] != "in_progress":
                return
            record["results"].extend(rows)
            record["matched_count"] = max(0, int(matched_count))
            record["remaining"] = max(0, int(remaining))
            self._touch(record)
            self._persist(record)

    def complete(self, operation_id: str) -> None:
        self._transition(operation_id, "complete", remaining=0)

    def fail(self, operation_id: str, error: Exception | str) -> None:
        self._transition(operation_id, "error", error=str(error))

    def interrupt(self, operation_id: str) -> None:
        self._transition(
            operation_id,
            "interrupted",
            guidance=INTERRUPTED_GUIDANCE,
        )

    def _transition(self, operation_id: str, status: str, **updates: Any) -> None:
        with self._lock:
            record = self._records.get(operation_id)
            if not record or record["status"] in TERMINAL_STATUSES:
                return
            record["status"] = status
            record.update(updates)
            self._touch(record)
            self._persist(record)
            self._events[operation_id].set()

    def _not_found(self, operation_id: str) -> dict[str, Any]:
        return {
            "operation_id": operation_id,
            "status": "not_found",
            "results": [],
            "summary": _summary([]),
            "matched_count": 0,
            "processed_count": 0,
            "success_count": 0,
            "failure_count": 0,
            "remaining": None,
            "guidance": NOT_FOUND_GUIDANCE,
        }

    def get(self, operation_id: str) -> dict[str, Any]:
        with self._lock:
            record = self._records.get(operation_id)
            if not record:
                return self._not_found(operation_id)
            if self._is_expired(record):
                self._records.pop(operation_id, None)
                self._events.pop(operation_id, None)
                try:
                    self._snapshot_path(operation_id).unlink()
                except OSError:
                    pass
                return self._not_found(operation_id)
            return json.loads(json.dumps(record, default=str))

    def wait(self, operation_id: str, timeout: float) -> dict[str, Any]:
        with self._lock:
            event = self._events.get(operation_id)
        if event is None:
            return self._not_found(operation_id)
        event.wait(max(0, timeout))
        return self.get(operation_id)

    def in_flight(self) -> int:
        with self._lock:
            return sum(
                record.get("status") == "in_progress"
                for record in self._records.values()
            )

    def shutdown(self, timeout: float = 5) -> None:
        self._closing.set()
        with self._lock:
            threads = list(self._threads.items())
        deadline = datetime.now().timestamp() + max(0, timeout)
        for _operation_id, thread in threads:
            remaining = max(0, deadline - datetime.now().timestamp())
            thread.join(remaining)
        with self._lock:
            active_ids = [
                operation_id
                for operation_id, record in self._records.items()
                if record.get("status") == "in_progress"
            ]
        for operation_id in active_ids:
            self.interrupt(operation_id)
