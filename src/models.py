from __future__ import annotations

from enum import StrEnum

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


class TransferMethod(StrEnum):
    PERCENT = "Процент"
    FIXED = "Фиксированная сумма"

    @classmethod
    def from_raw(cls, value: object) -> "TransferMethod":
        if isinstance(value, TransferMethod):
            return value
        text = str(value or "").strip()
        aliases = {
            "Процент": cls.PERCENT,
            "%": cls.PERCENT,
            "Фиксированная сумма": cls.FIXED,
            "Сумма/ост.": cls.FIXED,
            "Сумма": cls.FIXED,
        }
        if text not in aliases:
            raise ValueError("Метод переноса должен быть 'Процент' или 'Фиксированная сумма'.")
        return aliases[text]


class LiftInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)

    number: int | None = Field(default=None, ge=1)
    lift_name: str = Field(default="")
    original_lift_price_cny: float = Field(ge=0)
    stops: int = Field(gt=0)
    installation_price_per_stop_rub: float = Field(ge=0)

    @field_validator("lift_name")
    @classmethod
    def lift_name_must_not_be_empty(cls, value: str) -> str:
        return value or "Без названия"


class ProjectParams(BaseModel):
    model_config = ConfigDict(use_enum_values=False)

    exchange_rate_rub_per_cny: float = Field(gt=0)
    installation_markup: float = Field(ge=0)
    transfer_method: TransferMethod = TransferMethod.PERCENT
    transfer_share: float = Field(default=0.0, ge=0, le=1)
    fixed_transfer_per_stop_rub: float = Field(default=0.0, ge=0)

    @field_validator("transfer_method", mode="before")
    @classmethod
    def normalize_transfer_method(cls, value: object) -> TransferMethod:
        return TransferMethod.from_raw(value)

    @model_validator(mode="after")
    def method_specific_values_are_present(self) -> "ProjectParams":
        if self.transfer_method == TransferMethod.PERCENT and self.transfer_share is None:
            raise ValueError("Для метода 'Процент' укажите долю переноса.")
        if self.transfer_method == TransferMethod.FIXED and self.fixed_transfer_per_stop_rub is None:
            raise ValueError("Для фиксированного метода укажите сумму переноса за 1 остановку.")
        return self


class LiftResult(BaseModel):
    number: int | None
    lift_name: str
    original_lift_price_cny: float
    stops: int
    installation_price_per_stop_rub: float
    original_installation_total_rub: float
    installation_per_stop_with_markup_rub: float
    transfer_per_stop_rub: float
    transfer_total_rub: float
    transferred_to_lift_cny: float
    new_lift_price_cny: float
    new_installation_total_rub: float
    new_installation_per_stop_rub: float
    control_cny: float
    status: str


class ProjectSummary(BaseModel):
    total_original_lifts_cny: float
    total_new_lifts_cny: float
    lift_price_delta_cny: float
    total_original_installation_rub: float
    total_new_installation_rub: float
    total_transferred_rub: float
    total_transferred_cny: float
    total_project_before_cny: float
    total_project_after_cny: float
    project_control_cny: float
    status: str

