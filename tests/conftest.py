import json
from pathlib import Path


_CRITERION_DESCRIPTIONS = {
    "criterion_1": "bounded bulk operation polling",
    "criterion_2": "stable message identity and degradation shape",
    "criterion_3": "bounded degraded status diagnostics",
    "criterion_4": "actionable structured tool errors",
    "criterion_5": "serialized parallel calls with queue metrics",
    "criterion_6": "calendar local-time echo and timezone",
    "criterion_7": "Exchange DASL and capped IMAP sender modes",
    "criterion_8": "idempotent move rerun and restart interruption",
}
_EXPECTED_CRITERIA = set(_CRITERION_DESCRIPTIONS)
_RELIABILITY_CRITERIA = {}


def pytest_configure(config):
    _RELIABILITY_CRITERIA.clear()


def pytest_collection_modifyitems(config, items):
    for item in items:
        marker = item.get_closest_marker("reliability_criterion")
        if marker:
            _RELIABILITY_CRITERIA[item.nodeid] = {
                "criterion": marker.args[0],
                "passed": False,
            }


def pytest_runtest_logreport(report):
    if report.when != "call":
        return
    entry = _RELIABILITY_CRITERIA.get(report.nodeid)
    if entry is not None:
        entry["passed"] = report.passed


def pytest_sessionfinish(session, exitstatus):
    collected = _RELIABILITY_CRITERIA
    criteria = {
        entry["criterion"]: entry["passed"]
        for entry in collected.values()
    }
    if set(criteria) != _EXPECTED_CRITERIA:
        return

    artifact = {
        "schema_version": 1,
        "suite": "reliability-v0.5.0-hermetic-acceptance",
        "passed": exitstatus == 0 and all(criteria.values()),
        "criteria": {
            criterion: {
                "description": _CRITERION_DESCRIPTIONS[criterion],
                "passed": criteria[criterion],
            }
            for criterion in sorted(criteria)
        },
    }
    path = Path(__file__).parent / "artifacts" / "reliability_acceptance.json"
    path.parent.mkdir(exist_ok=True)
    path.write_text(
        json.dumps(artifact, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
