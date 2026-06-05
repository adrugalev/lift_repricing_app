from __future__ import annotations

from html import escape
from io import BytesIO

import pandas as pd
import streamlit as st

from src.calculator import calculate_project
from src.excel_io import (
    export_repriced_pricing_excel,
    load_pricing_from_excel,
    results_to_dataframe,
    summary_to_dataframe,
)
from src.formatter import format_cny, format_rub
from src.models import ProjectParams, TransferMethod
from src.validation import INPUT_COLUMNS, dataframe_to_lift_dicts, validate_lifts, validate_params
from src.version import repricing_version_label


MONEY_COLUMNS = [
    "Исходная цена лифта, CNY",
    "Монтаж за 1 остановку без наценки, RUB",
    "Монтаж всего до переноса, RUB",
    "Монтаж за 1 остановку с наценкой, RUB",
    "Перенос за 1 остановку, RUB",
    "Сумма переноса, RUB",
    "Перенесено в лифт, CNY",
    "Новая цена лифта, CNY",
    "Новый монтаж, RUB",
    "Новый монтаж за 1 остановку, RUB",
]

INPUT_COLUMN_CONFIG = {
    "№": st.column_config.NumberColumn("№", min_value=1, step=1, width=60, format="%d"),
    "Лифт": st.column_config.TextColumn("Лифт", width=90),
    "Цена лифта для клиента, CNY": st.column_config.NumberColumn(
        "Цена CNY",
        min_value=0,
        step=1000,
        width=120,
        format="%d",
    ),
    "Количество остановок": st.column_config.NumberColumn(
        "Ост.",
        min_value=1,
        step=1,
        width=70,
        format="%d",
    ),
    "Монтаж за 1 остановку без наценки, RUB": st.column_config.NumberColumn(
        "Монтаж ₽",
        min_value=0,
        step=1000,
        width=120,
        format="%d",
    ),
}

RESULT_DISPLAY_COLUMNS = {
    "№": "№",
    "Лифт": "Лифт",
    "Исходная цена лифта, CNY": "Лифт до|переноса|CNY",
    "Количество остановок": "Ост.",
    "Монтаж всего до переноса, RUB": "Монтаж до|переноса|₽",
    "Перенос за 1 остановку, RUB": "Перенос за|остановку|₽",
    "Сумма переноса, RUB": "Перенос|всего|₽",
    "Перенесено в лифт, CNY": "Перенос|в лифт|CNY",
    "Новая цена лифта, CNY": "Лифт после|переноса|CNY",
    "Новый монтаж, RUB": "Монтаж после|переноса|₽",
    "Новый монтаж за 1 остановку, RUB": "Монтаж после|переноса за 1 ост.|₽",
}

RESULT_CNY_COLUMNS = [
    "Исходная цена лифта, CNY",
    "Перенесено в лифт, CNY",
    "Новая цена лифта, CNY",
]

RESULT_RUB_COLUMNS = [
    "Монтаж всего до переноса, RUB",
    "Перенос за 1 остановку, RUB",
    "Сумма переноса, RUB",
    "Новый монтаж, RUB",
    "Новый монтаж за 1 остановку, RUB",
]

RESULT_COLUMN_CONFIG = {
    "№": st.column_config.NumberColumn("№", width=52, format="%d"),
    "Лифт": st.column_config.TextColumn("Лифт", width=70),
    "До CNY": st.column_config.TextColumn("До CNY", width=120),
    "Ост.": st.column_config.NumberColumn("Ост.", width=58, format="%d"),
    "Монт. до ₽": st.column_config.TextColumn("Монт. до ₽", width=120),
    "Р/ост. ₽": st.column_config.TextColumn("Р/ост. ₽", width=95),
    "Перенос ₽": st.column_config.TextColumn("Перенос ₽", width=110),
    "+CNY": st.column_config.TextColumn("+CNY", width=105),
    "После CNY": st.column_config.TextColumn("После CNY", width=120),
    "Монт. после ₽": st.column_config.TextColumn("Монт. после ₽", width=125),
    "Монт. после/ост. ₽": st.column_config.TextColumn("Монт. после/ост. ₽", width=125),
    "Статус": st.column_config.TextColumn("Статус", width=75),
}


def default_params() -> ProjectParams:
    return ProjectParams(
        exchange_rate_rub_per_cny=11.13,
        installation_markup=0.0526,
        currency_reserve_share=0.0,
        transfer_method=TransferMethod.PERCENT,
        transfer_share=0.2,
        fixed_transfer_per_stop_rub=0.0,
    )


def sync_param_widgets(params: ProjectParams) -> None:
    st.session_state.exchange_rate_input = float(params.exchange_rate_rub_per_cny)
    st.session_state.installation_markup_input = float(params.installation_markup)
    st.session_state.currency_reserve_input = float(params.currency_reserve_share * 100)
    st.session_state.method_input = params.transfer_method.value
    st.session_state.transfer_share_input = float(params.transfer_share)
    st.session_state.fixed_transfer_input = float(params.fixed_transfer_per_stop_rub)


def ensure_param_widgets_initialized(params: ProjectParams) -> None:
    defaults = {
        "exchange_rate_input": float(params.exchange_rate_rub_per_cny),
        "installation_markup_input": float(params.installation_markup),
        "currency_reserve_input": float(params.currency_reserve_share * 100),
        "method_input": params.transfer_method.value,
        "transfer_share_input": float(params.transfer_share),
        "fixed_transfer_input": float(params.fixed_transfer_per_stop_rub),
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def default_input_dataframe() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "№": 1,
                "Лифт": "L1",
                "Цена лифта для клиента, CNY": 0.0,
                "Количество остановок": 1,
                "Монтаж за 1 остановку без наценки, RUB": 0.0,
            }
        ],
        columns=INPUT_COLUMNS,
    )


def format_exchange_rate(value: float) -> str:
    return f"{value:.2f}".replace(".", ",")


def format_percent(value: float) -> str:
    return f"{value * 100:.1f}%".replace(".", ",")


def show_summary_metrics(summary: object, params: ProjectParams) -> None:
    metrics = [
        ("Лифты", format_cny(summary.total_original_lifts_cny), format_cny(summary.total_new_lifts_cny), format_cny(summary.lift_price_delta_cny)),
        ("Монтаж", format_rub(summary.total_original_installation_rub), format_rub(summary.total_new_installation_rub), f"-{format_rub(summary.total_transferred_rub)}"),
    ]
    body_rows = "".join(
        f"<tr><td>{label}</td><td>{before}</td><td>{after}</td><td>{delta}</td></tr>"
        for label, before, after, delta in metrics
    )
    body_rows += (
        f'<tr><td>Курс из расценки</td><td colspan="3">'
        f"{format_exchange_rate(params.exchange_rate_rub_per_cny)} RUB за 1 CNY</td></tr>"
        f'<tr><td>Валютный резерв</td><td colspan="3">{format_percent(params.currency_reserve_share)}</td></tr>'
        f'<tr><td>Курс переноса</td><td colspan="3">'
        f"{format_exchange_rate(params.transfer_exchange_rate_rub_per_cny)} RUB за 1 CNY</td></tr>"
    )
    st.markdown(
        f"""
        <style>
          .summary-table {{
            width: 720px;
            max-width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
            font-size: 13px;
            line-height: 1.25;
            margin-bottom: 14px;
          }}
          .summary-table th,
          .summary-table td {{
            border: 1px solid #e5e7eb;
            padding: 7px 10px;
            white-space: nowrap;
          }}
      .summary-table th {{
        background: #f8fafc;
        color: #64748b;
        font-weight: 600;
        text-align: center;
          }}
          .summary-table td {{
            text-align: right;
            font-weight: 600;
          }}
          .summary-table td:first-child {{
            text-align: left;
            font-weight: 600;
          }}
          .summary-table td[colspan] {{
            text-align: left;
            color: #334155;
          }}
        </style>
        <table class="summary-table">
          <thead><tr><th>Показатель</th><th>До</th><th>После</th><th>Изменение</th></tr></thead>
          <tbody>{body_rows}</tbody>
        </table>
        """,
        unsafe_allow_html=True,
    )


def round_result_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    rounded = df.copy()
    for column in MONEY_COLUMNS:
        if column in rounded:
            rounded[column] = rounded[column].round(0).astype("Int64")
    if "Проверка / статус" in rounded:
        rounded["Проверка / статус"] = rounded["Проверка / статус"].astype(str)
    return rounded


def format_number(value: float) -> str:
    if pd.isna(value):
        return ""
    return f"{round(float(value)):,}".replace(",", " ")


def compact_result_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    compact = df[list(RESULT_DISPLAY_COLUMNS)].copy()
    for column in RESULT_CNY_COLUMNS:
        compact[column] = compact[column].apply(lambda value: format_number(float(value)))
    for column in RESULT_RUB_COLUMNS:
        compact[column] = compact[column].apply(lambda value: format_number(float(value)))
    compact = compact.rename(columns=RESULT_DISPLAY_COLUMNS)
    return compact


def compact_input_dataframe(df: pd.DataFrame, installation_markup: float) -> pd.DataFrame:
    compact = df.copy()
    stops = pd.to_numeric(compact["Количество остановок"], errors="coerce")
    installation_per_stop_with_markup = (
        pd.to_numeric(compact["Монтаж за 1 остановку без наценки, RUB"], errors="coerce")
        * (1 + installation_markup)
    )
    compact["Монтаж за 1 остановку с наценкой, RUB"] = installation_per_stop_with_markup
    compact["Монтаж всего с наценкой, RUB"] = stops * installation_per_stop_with_markup
    compact["Цена лифта для клиента, CNY"] = pd.to_numeric(
        compact["Цена лифта для клиента, CNY"], errors="coerce"
    ).apply(lambda value: format_number(float(value)) if pd.notna(value) else "")
    compact["Монтаж за 1 остановку с наценкой, RUB"] = compact[
        "Монтаж за 1 остановку с наценкой, RUB"
    ].apply(lambda value: format_number(float(value)) if pd.notna(value) else "")
    compact["Монтаж всего с наценкой, RUB"] = compact[
        "Монтаж всего с наценкой, RUB"
    ].apply(lambda value: format_number(float(value)) if pd.notna(value) else "")
    compact = compact.drop(columns=["Монтаж за 1 остановку без наценки, RUB"])
    return compact.rename(
        columns={
            "Цена лифта для клиента, CNY": "Цена лифта с наценкой|CNY",
            "Количество остановок": "Остановки",
            "Монтаж за 1 остановку с наценкой, RUB": "Монтаж за остановку|с наценкой, ₽",
            "Монтаж всего с наценкой, RUB": "Монтаж за лифт|с наценкой, ₽",
        }
    )


def render_compact_table(df: pd.DataFrame) -> str:
    column_widths = {
        "№": "4%",
        "Лифт": "6%",
        "Лифт до|переноса|CNY": "10%",
        "Ост.": "5%",
        "Монтаж до|переноса|₽": "10%",
        "Перенос за|остановку|₽": "9%",
        "Перенос|всего|₽": "9%",
        "Перенос|в лифт|CNY": "8%",
        "Лифт после|переноса|CNY": "10%",
        "Монтаж после|переноса|₽": "10%",
        "Монтаж после|переноса за 1 ост.|₽": "11%",
    }
    header_cells = "".join(
        f'<th style="width: {column_widths.get(column, "auto")}">{"<br>".join(escape(part) for part in str(column).split("|"))}</th>'
        for column in df.columns
    )
    body_rows = []
    for _, row in df.iterrows():
        cells = "".join(f"<td>{escape(str(value))}</td>" for value in row)
        body_rows.append(f"<tr>{cells}</tr>")
    return f"""
    <style>
      .result-table {{
        width: 1120px;
        max-width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
        font-size: 12px;
        line-height: 1.25;
      }}
      .result-table th,
      .result-table td {{
        border: 1px solid #e5e7eb;
        padding: 6px 8px;
        overflow: hidden;
        text-overflow: clip;
      }}
      .result-table th {{
        background: #f8fafc;
        color: #64748b;
        font-weight: 600;
        line-height: 1.15;
        white-space: normal;
        text-align: center;
      }}
      .result-table td {{
        text-align: right;
        white-space: nowrap;
      }}
      .result-table th:nth-child(9),
      .result-table td:nth-child(9),
      .result-table th:nth-child(10),
      .result-table td:nth-child(10),
      .result-table th:nth-child(11),
      .result-table td:nth-child(11) {{
        font-weight: 700;
        color: #111827;
      }}
      .result-table th:nth-child(9),
      .result-table th:nth-child(10),
      .result-table th:nth-child(11) {{
        background: #dcfce7;
        color: #166534;
      }}
      .result-table td:nth-child(9),
      .result-table td:nth-child(10),
      .result-table td:nth-child(11) {{
        background: #f0fdf4;
      }}
      .result-table td:nth-child(2) {{
        text-align: left;
      }}
    </style>
    <table class="result-table">
      <thead><tr>{header_cells}</tr></thead>
      <tbody>{"".join(body_rows)}</tbody>
    </table>
    """


def render_input_table(df: pd.DataFrame) -> str:
    column_widths = {
        "№": "6%",
        "Лифт": "12%",
        "Цена лифта с наценкой|CNY": "22%",
        "Остановки": "12%",
        "Монтаж за остановку|с наценкой, ₽": "24%",
        "Монтаж за лифт|с наценкой, ₽": "24%",
    }
    header_cells = "".join(
        f'<th style="width: {column_widths.get(column, "auto")}">{"<br>".join(escape(part) for part in str(column).split("|"))}</th>'
        for column in df.columns
    )
    body_rows = []
    for _, row in df.iterrows():
        cells = "".join(f"<td>{escape(str(value))}</td>" for value in row)
        body_rows.append(f"<tr>{cells}</tr>")
    return f"""
    <style>
      .input-table {{
        width: 760px;
        max-width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
        font-size: 12px;
        line-height: 1.2;
        margin-bottom: 14px;
      }}
      .input-table th,
      .input-table td {{
        border: 1px solid #e5e7eb;
        padding: 5px 8px;
        overflow: hidden;
        text-overflow: clip;
      }}
      .input-table th {{
        background: #f8fafc;
        color: #64748b;
        font-weight: 600;
        line-height: 1.15;
        white-space: normal;
        text-align: center;
      }}
      .input-table td:nth-child(1) {{
        text-align: right;
      }}
      .input-table td:nth-child(2) {{
        text-align: left;
      }}
      .input-table td:nth-child(4) {{
        text-align: right;
      }}
      .input-table td {{
        text-align: right;
        white-space: nowrap;
      }}
    </style>
    <table class="input-table">
      <thead><tr>{header_cells}</tr></thead>
      <tbody>{"".join(body_rows)}</tbody>
    </table>
    """


def format_summary_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    formatted = df.copy()

    def format_row(row: pd.Series) -> str:
        value = row["Значение"]
        currency = row["Валюта"]
        if currency == "CNY":
            return format_cny(float(value))
        if currency == "RUB":
            return format_rub(float(value))
        if currency == "RATE":
            return f"{format_exchange_rate(float(value))} RUB за 1 CNY"
        if currency == "PERCENT":
            return format_percent(float(value))
        return str(value)

    formatted["Значение"] = formatted.apply(format_row, axis=1)
    return formatted[["Показатель", "Значение"]]


def table_height(row_count: int, max_rows: int = 12) -> int:
    visible_rows = min(max(row_count, 1), max_rows)
    return 39 + visible_rows * 35


def main() -> None:
    st.set_page_config(page_title="Групповая переоценка лифтов EPSS", layout="wide")
    st.markdown(
        """
        <style>
        .block-container {
            max-width: 1120px;
            padding-left: 2rem;
            padding-right: 2rem;
        }
        div[data-testid="stVerticalBlock"]:has(> div .compact-input-table) {
            max-width: 760px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.title("Групповая переоценка лифтов EPSS")
    st.caption(repricing_version_label())

    with st.sidebar:
        st.header("Файлы")
        pricing_file = st.file_uploader(
            "Расценка или расчет стоимости",
            type=["xlsx"],
            help=(
                "Например: файл с листом 'Стоимость'. Таблица лифтов будет заполнена "
                "из колонок Продукт, Кол-во этажей, Общая стоимость для клиента, "
                "Стоимость монтажа за этаж."
            ),
        )

    if "input_df" not in st.session_state:
        st.session_state.input_df = default_input_dataframe()
    if "loaded_pricing_file_name" not in st.session_state:
        st.session_state.loaded_pricing_file_name = None
    if "pricing_file_bytes" not in st.session_state:
        st.session_state.pricing_file_bytes = None
    if "params" not in st.session_state:
        st.session_state.params = default_params()
    ensure_param_widgets_initialized(st.session_state.params)

    if pricing_file and pricing_file.name != st.session_state.loaded_pricing_file_name:
        pricing_file_bytes = pricing_file.getvalue()
        input_df, imported_params, warnings = load_pricing_from_excel(BytesIO(pricing_file_bytes))
        st.session_state.input_df = input_df if not input_df.empty else default_input_dataframe()
        if imported_params:
            st.session_state.params = imported_params
            sync_param_widgets(imported_params)
        st.session_state.loaded_pricing_file_name = pricing_file.name
        st.session_state.pricing_file_bytes = pricing_file_bytes
        st.success(f"Загружено строк из файла расценки: {len(st.session_state.input_df)}")
        for warning in warnings:
            st.warning(warning)

    st.subheader("Таблица лифтов")
    st.markdown(
        render_input_table(
            compact_input_dataframe(
                st.session_state.input_df,
                float(st.session_state.installation_markup_input),
            )
        ),
        unsafe_allow_html=True,
    )

    editor_key = f"lifts_editor_{st.session_state.loaded_pricing_file_name or 'manual'}"
    with st.expander("Редактировать таблицу лифтов"):
        edited_df = st.data_editor(
            st.session_state.input_df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            height=table_height(len(st.session_state.input_df), max_rows=10),
            column_config=INPUT_COLUMN_CONFIG,
            key=editor_key,
        )
        st.session_state.input_df = edited_df

    st.subheader("Параметры переноса")
    method_options = [TransferMethod.PERCENT.value, TransferMethod.FIXED.value]
    method_col, reserve_col, value_col = st.columns([1, 1, 3])
    with method_col:
        method = st.selectbox(
            "Метод",
            method_options,
            index=method_options.index(st.session_state.method_input),
            key="method_input",
        )
    with reserve_col:
        currency_reserve_percent = st.number_input(
            "Валютный резерв, %",
            min_value=0.0,
            max_value=99.0,
            step=0.5,
            format="%.1f",
            help=(
                "Защитный запас от курсовой разницы при переносе рублевого монтажа в цену лифта в CNY. "
                "Курс переноса считается как: курс из расценки × (1 - валютный резерв). "
                "Пример: курс 11,00 RUB/CNY и резерв 5% дают курс переноса 10,45 RUB/CNY. "
                "Тогда одна и та же сумма монтажа в RUB переносится в большее количество CNY."
            ),
            key="currency_reserve_input",
        )

    if TransferMethod.from_raw(method) == TransferMethod.PERCENT:
        with value_col:
            transfer_share = st.slider(
                "Доля переноса",
                min_value=0.0,
                max_value=1.0,
                step=0.01,
                key="transfer_share_input",
            )
        fixed_transfer = float(st.session_state.fixed_transfer_input)
    else:
        with value_col:
            fixed_transfer = st.number_input(
                "Фиксированный перенос на 1 остановку, RUB",
                min_value=0.0,
                step=1000.0,
                format="%.0f",
                key="fixed_transfer_input",
            )
        transfer_share = float(st.session_state.transfer_share_input)

    raw_params = {
        "exchange_rate_rub_per_cny": st.session_state.exchange_rate_input,
        "installation_markup": st.session_state.installation_markup_input,
        "currency_reserve_share": currency_reserve_percent / 100,
        "transfer_method": method,
        "transfer_share": transfer_share,
        "fixed_transfer_per_stop_rub": fixed_transfer,
    }

    params, param_errors = validate_params(raw_params)
    lift_dicts = dataframe_to_lift_dicts(edited_df)
    lifts, lift_errors = validate_lifts(lift_dicts)

    if param_errors or lift_errors:
        st.warning("Данные неполные или содержат ошибки.")
        for error in [*param_errors, *lift_errors]:
            st.error(error)
        return

    assert params is not None
    st.session_state.params = params
    results, summary = calculate_project(lifts, params)

    st.subheader("Результирующая таблица")
    result_df = results_to_dataframe(results)
    st.markdown(render_compact_table(compact_result_dataframe(result_df)), unsafe_allow_html=True)

    st.subheader("Сводные показатели")
    show_summary_metrics(summary, params)

    with st.expander("Подробная сводка"):
        st.dataframe(
            format_summary_dataframe(summary_to_dataframe(summary, params)),
            use_container_width=True,
            hide_index=True,
            height=table_height(14, max_rows=14),
        )

    if st.session_state.pricing_file_bytes:
        output = export_repriced_pricing_excel(
            BytesIO(st.session_state.pricing_file_bytes),
            results,
            params,
        )
        source_name = st.session_state.loaded_pricing_file_name or "pricing.xlsx"
        output_name = source_name.replace(".xlsx", "_переоценка_формулы.xlsx")
        st.caption("Экспорт: формульная расценка, версия XML-patch v3.")
        st.download_button(
            "Скачать обновленную расценку",
            data=output,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.download_button(
            "Скачать обновленную расценку",
            data=b"",
            file_name="pricing_repriced.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=True,
            help="Сначала загрузите файл расценки.",
        )


if __name__ == "__main__":
    main()
