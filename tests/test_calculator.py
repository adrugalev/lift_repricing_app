from __future__ import annotations

import pytest

from src.calculator import calculate_lift, calculate_project
from src.models import LiftInput, ProjectParams, TransferMethod


@pytest.fixture
def excel_percent_params() -> ProjectParams:
    return ProjectParams(
        exchange_rate_rub_per_cny=11.13,
        installation_markup=0.0526,
        transfer_method=TransferMethod.PERCENT,
        transfer_share=0.2,
        fixed_transfer_per_stop_rub=0,
    )


def test_percent_method_matches_excel_group_row_11(excel_percent_params: ProjectParams) -> None:
    lift = LiftInput(
        number=1,
        lift_name="L1",
        original_lift_price_cny=180_200,
        stops=9,
        installation_price_per_stop_rub=215_000,
    )

    result = calculate_lift(lift, excel_percent_params)

    assert result.original_installation_total_rub == pytest.approx(2_036_781)
    assert result.installation_per_stop_with_markup_rub == pytest.approx(226_309)
    assert result.transfer_per_stop_rub == pytest.approx(45_261.8)
    assert result.transfer_total_rub == pytest.approx(407_356.2)
    assert result.transferred_to_lift_cny == pytest.approx(36_599.83827493261)
    assert result.new_lift_price_cny == pytest.approx(216_799.8382749326)
    assert result.new_installation_total_rub == pytest.approx(1_629_424.8)
    assert result.control_cny == pytest.approx(0, abs=0.01)
    assert result.status == "OK"


def test_fixed_method_caps_transfer_by_installation_with_markup() -> None:
    params = ProjectParams(
        exchange_rate_rub_per_cny=10,
        installation_markup=0.1,
        transfer_method=TransferMethod.FIXED,
        transfer_share=0,
        fixed_transfer_per_stop_rub=1_000_000,
    )
    lift = LiftInput(
        number=1,
        lift_name="L1",
        original_lift_price_cny=100_000,
        stops=2,
        installation_price_per_stop_rub=100_000,
    )

    result = calculate_lift(lift, params)

    assert result.installation_per_stop_with_markup_rub == pytest.approx(110_000)
    assert result.transfer_per_stop_rub == pytest.approx(110_000)
    assert result.new_installation_total_rub == pytest.approx(0)
    assert result.status == "OK, перенос ограничен монтажом"


def test_currency_reserve_uses_protective_transfer_rate() -> None:
    params = ProjectParams(
        exchange_rate_rub_per_cny=10,
        installation_markup=0.1,
        currency_reserve_share=0.05,
        transfer_method=TransferMethod.PERCENT,
        transfer_share=0.2,
        fixed_transfer_per_stop_rub=0,
    )
    lift = LiftInput(
        number=1,
        lift_name="L1",
        original_lift_price_cny=100_000,
        stops=2,
        installation_price_per_stop_rub=100_000,
    )

    result = calculate_lift(lift, params)

    assert params.transfer_exchange_rate_rub_per_cny == pytest.approx(9.5)
    assert result.transfer_total_rub == pytest.approx(44_000)
    assert result.transferred_to_lift_cny == pytest.approx(44_000 / 9.5)
    assert result.new_lift_price_cny == pytest.approx(104_631.57894736843)
    assert result.control_cny == pytest.approx(0, abs=0.01)
    assert result.status == "OK"


def test_project_summary_preserves_total_project_value(excel_percent_params: ProjectParams) -> None:
    lifts = [
        LiftInput(
            number=1,
            lift_name="L1",
            original_lift_price_cny=180_200,
            stops=9,
            installation_price_per_stop_rub=215_000,
        ),
        LiftInput(
            number=2,
            lift_name="L2",
            original_lift_price_cny=153_000,
            stops=8,
            installation_price_per_stop_rub=215_000,
        ),
    ]

    results, summary = calculate_project(lifts, excel_percent_params)

    assert len(results) == 2
    assert summary.total_transferred_rub == pytest.approx(769_450.6)
    assert summary.total_transferred_cny == pytest.approx(69_133.0278526505)
    assert summary.total_new_installation_rub == pytest.approx(3_077_802.4)
    assert summary.project_control_cny == pytest.approx(0, abs=0.01)
    assert summary.status == "OK"
