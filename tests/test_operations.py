import json
from datetime import UTC, datetime, timedelta

from outlook_desktop_mcp.operations import (
    INTERRUPTED_GUIDANCE,
    NOT_FOUND_GUIDANCE,
    OperationManager,
)


class Clock:
    def __init__(self):
        self.value = datetime(2026, 8, 25, tzinfo=UTC)

    def now(self):
        return self.value


def test_atomic_persistence_and_ttl(tmp_path):
    clock = Clock()
    manager = OperationManager(
        tmp_path,
        process_instance_id="process-a",
        now=clock.now,
    )
    operation_id = manager.create("bulk_mark_as_read", remaining=2)
    manager.append_results(
        operation_id,
        [{"id": "one", "status": "ok"}],
        matched_count=2,
        remaining=1,
    )
    manager.complete(operation_id)

    snapshot_path = tmp_path / f"{operation_id}.json"
    assert json.loads(snapshot_path.read_text())["status"] == "complete"
    assert not list(tmp_path.glob("*.tmp"))
    assert OperationManager(
        tmp_path,
        process_instance_id="process-b",
        now=clock.now,
    ).get(operation_id)["status"] == "complete"

    clock.value += timedelta(hours=73)
    expired = OperationManager(
        tmp_path,
        process_instance_id="process-c",
        now=clock.now,
    ).get(operation_id)
    assert expired["status"] == "not_found"
    assert expired["guidance"] == NOT_FOUND_GUIDANCE


def test_startup_converts_previous_process_work_to_interrupted(tmp_path):
    first = OperationManager(tmp_path, process_instance_id="process-a")
    operation_id = first.create("bulk_move_emails", remaining=30)
    first.append_results(
        operation_id,
        [{"id": "one", "status": "ok"}],
        matched_count=40,
        remaining=39,
    )

    restarted = OperationManager(tmp_path, process_instance_id="process-b")
    payload = restarted.get(operation_id)

    assert payload["status"] == "interrupted"
    assert payload["results"] == [{"id": "one", "status": "ok"}]
    assert payload["remaining"] == 39
    assert payload["guidance"] == INTERRUPTED_GUIDANCE


def test_corrupt_snapshot_is_quarantined_without_crashing(tmp_path):
    corrupt = tmp_path / "broken.json"
    corrupt.write_text("{not-json", encoding="utf-8")

    manager = OperationManager(tmp_path, process_instance_id="process-a")

    assert manager.get("broken")["status"] == "not_found"
    assert not corrupt.exists()
    assert list(tmp_path.glob("broken.json.corrupt-*"))


def test_unknown_operation_is_not_found(tmp_path):
    payload = OperationManager(tmp_path).get("unknown")

    assert payload["status"] == "not_found"
    assert payload["results"] == []
    assert payload["remaining"] is None
    assert payload["guidance"] == NOT_FOUND_GUIDANCE


def test_clean_shutdown_terminalizes_active_snapshot(tmp_path):
    manager = OperationManager(tmp_path, process_instance_id="process-a")
    operation_id = manager.create("bulk_read_emails", remaining=10)

    manager.shutdown(timeout=0)

    snapshot = json.loads(
        (tmp_path / f"{operation_id}.json").read_text(encoding="utf-8")
    )
    assert snapshot["status"] == "interrupted"
    assert snapshot["remaining"] == 10
    assert snapshot["guidance"] == INTERRUPTED_GUIDANCE
