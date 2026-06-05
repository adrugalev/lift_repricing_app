from __future__ import annotations

from collections.abc import Iterable
from typing import Any

import pandas as pd
from pydantic import ValidationError

from .models import LiftInput, ProjectParams

INPUT_COLUMNS = [
    "№",
    "Лифт",
    "Цена лифта для клиента, CNY",
    "Количество остановок",
    "Монтаж за 1 остановку без наценки, RUB",
]


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, float) and pd.isna(value)) or str(value).strip() == ""


def optional_int(value: Any) -> int | None:
    if is_blank(value):
        return None
    numeric = pd.to_numeric(value, errors="coerce")
    if pd.isna(numeric):
        return None
    return int(numeric)


def validate_params(raw: dict[str, Any]) -> tuple[ProjectParams | None, list[str]]:
    try:
        return ProjectParams(**raw), []
    except ValidationError as exc:
        return None, [f"Параметры проекта: {error['msg']}" for error in exc.errors()]


def validate_lifts(raw_lifts: Iterable[dict[str, Any]]) -> tuple[list[LiftInput], list[str]]:
    lifts: list[LiftInput] = []
    errors: list[str] = []
    for idx, raw in enumerate(raw_lifts, start=1):
        try:
            lifts.append(LiftInput(**raw))
        except ValidationError as exc:
            for error in exc.errors():
                field = ".".join(str(part) for part in error["loc"])
                errors.append(f"Строка {idx}: {field}: {error['msg']}")
    if not lifts and not errors:
        errors.append("Добавьте хотя бы один лифт для расчета.")
    return lifts, errors


def dataframe_to_lift_dicts(df: pd.DataFrame) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        if all(is_blank(row.get(column)) for column in INPUT_COLUMNS[1:]):
            continue
        number = row.get("№")
        rows.append(
            {
                "number": optional_int(number),
                "lift_name": row.get("Лифт") or "",
                "original_lift_price_cny": row.get("Цена лифта для клиента, CNY"),
                "stops": row.get("Количество остановок"),
                "installation_price_per_stop_rub": row.get(
                    "Монтаж за 1 остановку без наценки, RUB"
                ),
            }
        )
    return rows
