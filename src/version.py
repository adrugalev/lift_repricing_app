from __future__ import annotations

REPRICING_VERSION_DATE = "05.06.2026"
REPRICING_VERSION_REVISION = 10


def repricing_version_label() -> str:
    return f"Версия {REPRICING_VERSION_REVISION} от {REPRICING_VERSION_DATE}"
