from __future__ import annotations

from collections.abc import Iterable

from .models import LiftInput, LiftResult, ProjectParams, ProjectSummary, TransferMethod

CONTROL_EPS_CNY = 0.01


def calculate_lift(lift: LiftInput, params: ProjectParams) -> LiftResult:
    installation_per_stop_with_markup = (
        lift.installation_price_per_stop_rub * (1 + params.installation_markup)
    )
    original_installation_total = lift.stops * installation_per_stop_with_markup

    if params.transfer_method == TransferMethod.PERCENT:
        transfer_per_stop = installation_per_stop_with_markup * params.transfer_share
        capped = False
    else:
        transfer_per_stop = min(
            params.fixed_transfer_per_stop_rub,
            installation_per_stop_with_markup,
        )
        capped = params.fixed_transfer_per_stop_rub > installation_per_stop_with_markup

    transfer_total = lift.stops * transfer_per_stop
    transferred_to_lift_cny = transfer_total / params.transfer_exchange_rate_rub_per_cny
    new_lift_price_cny = lift.original_lift_price_cny + transferred_to_lift_cny
    new_installation_total = original_installation_total - transfer_total
    new_installation_per_stop = new_installation_total / lift.stops

    project_before = (
        lift.original_lift_price_cny
        + original_installation_total / params.transfer_exchange_rate_rub_per_cny
    )
    project_after = new_lift_price_cny + new_installation_total / params.transfer_exchange_rate_rub_per_cny
    control_cny = project_after - project_before

    if new_installation_total < -CONTROL_EPS_CNY:
        status = "Ошибка: монтаж стал отрицательным"
    elif abs(control_cny) >= CONTROL_EPS_CNY:
        status = "Проверьте контроль"
    elif capped:
        status = "OK, перенос ограничен монтажом"
    else:
        status = "OK"

    return LiftResult(
        number=lift.number,
        lift_name=lift.lift_name,
        original_lift_price_cny=lift.original_lift_price_cny,
        stops=lift.stops,
        installation_price_per_stop_rub=lift.installation_price_per_stop_rub,
        original_installation_total_rub=original_installation_total,
        installation_per_stop_with_markup_rub=installation_per_stop_with_markup,
        transfer_per_stop_rub=transfer_per_stop,
        transfer_total_rub=transfer_total,
        transferred_to_lift_cny=transferred_to_lift_cny,
        new_lift_price_cny=new_lift_price_cny,
        new_installation_total_rub=new_installation_total,
        new_installation_per_stop_rub=new_installation_per_stop,
        control_cny=control_cny,
        status=status,
    )


def calculate_project(
    lifts: Iterable[LiftInput],
    params: ProjectParams,
) -> tuple[list[LiftResult], ProjectSummary]:
    results = [calculate_lift(lift, params) for lift in lifts]

    total_original_lifts_cny = sum(row.original_lift_price_cny for row in results)
    total_new_lifts_cny = sum(row.new_lift_price_cny for row in results)
    total_original_installation_rub = sum(row.original_installation_total_rub for row in results)
    total_new_installation_rub = sum(row.new_installation_total_rub for row in results)
    total_transferred_rub = sum(row.transfer_total_rub for row in results)
    total_transferred_cny = sum(row.transferred_to_lift_cny for row in results)
    total_project_before_cny = (
        total_original_lifts_cny + total_original_installation_rub / params.transfer_exchange_rate_rub_per_cny
    )
    total_project_after_cny = (
        total_new_lifts_cny + total_new_installation_rub / params.transfer_exchange_rate_rub_per_cny
    )
    project_control_cny = total_project_after_cny - total_project_before_cny

    status = "OK" if abs(project_control_cny) < CONTROL_EPS_CNY else "Проверьте"

    return results, ProjectSummary(
        total_original_lifts_cny=total_original_lifts_cny,
        total_new_lifts_cny=total_new_lifts_cny,
        lift_price_delta_cny=total_new_lifts_cny - total_original_lifts_cny,
        total_original_installation_rub=total_original_installation_rub,
        total_new_installation_rub=total_new_installation_rub,
        total_transferred_rub=total_transferred_rub,
        total_transferred_cny=total_transferred_cny,
        total_project_before_cny=total_project_before_cny,
        total_project_after_cny=total_project_after_cny,
        project_control_cny=project_control_cny,
        status=status,
    )
