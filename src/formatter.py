from __future__ import annotations


CNY_NUMBER_FORMAT = '#,##0 "CNY"'
RUB_NUMBER_FORMAT = '#,##0 ₽'
PERCENT_NUMBER_FORMAT = "0.0%"


def _format_money(value: float, suffix: str) -> str:
    rounded = round(value)
    return f"{rounded:,.0f}".replace(",", " ") + f" {suffix}"


def format_cny(value: float) -> str:
    return _format_money(value, "CNY")


def format_rub(value: float) -> str:
    return _format_money(value, "₽")

