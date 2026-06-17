import datetime
import logging
from typing import Any, Callable, Optional

import pandas as pd
from pandas import DataFrame

from src.utils import base_dir, quantize_decimal

path_to_log = base_dir + "/logs/module_reports.log"
reports_logger = logging.getLogger(__name__)
if not reports_logger.handlers:
    file_handler = logging.FileHandler(path_to_log, mode="w")
    file_formatter = logging.Formatter("%(asctime)s - %(module)s - %(levelname)s - %(message)s")
    file_handler.setFormatter(file_formatter)
    reports_logger.addHandler(file_handler)
    reports_logger.setLevel(logging.DEBUG)


def writing_to_file(function: Callable) -> Callable:
    def inner(*args: Any, **kwargs: Any) -> Any:
        result = function(*args, **kwargs)
        print("Записываю отчёт в output_data/ ...")
        result.to_excel(
            excel_writer=f"{base_dir}/output_data/{function.__name__}.xlsx",
            index=False,
            float_format="%.2f",
            engine="openpyxl",
        )
        print("Отчёт создан")
        return result

    return inner


@writing_to_file
def spending_by_workday(transactions: pd.DataFrame, date: Optional[str] = None) -> DataFrame:
    """
    Функция принимает транзакции в виде DataFrame и опциональную дату в виде строки, а возвращает данные по средним
    тратам в виде DataFrame.
    """

    reports_logger.info("Запуск <spending_by_workday>")
    if date is None:
        date = "31.12.2021"
    date_obj = datetime.datetime.strptime(date, "%d.%m.%Y")
    date_full_obj = date_obj.replace(hour=23, minute=59, second=59)
    date_border = date_obj - datetime.timedelta(days=89)

    spent_all_df = transactions.loc[transactions.Итого < 0]
    spent_indexed_df = spent_all_df.reset_index(drop=True)
    spent_indexed_df["Итого"] = spent_indexed_df["Итого"].astype(str).transform(quantize_decimal)

    reports_logger.info("Создание датафрейма операций за 3 месяца")
    sector_date_df = spent_indexed_df[
        (spent_indexed_df["Дата_операции"] <= pd.Timestamp(date_full_obj))
        & (pd.Timestamp(date_border) <= spent_indexed_df["Дата_операции"])
    ]

    only_day_df = sector_date_df[["Дата_операции", "Итого"]]
    only_day_df.rename(columns={"Дата_операции": "Дата_операций"}, inplace=True)
    only_day_df["Дата_операций"] = only_day_df["Дата_операций"].apply(lambda x: x.date())
    if not pd.api.types.is_datetime64_any_dtype(only_day_df["Дата_операций"]):
        only_day_df["Дата_операций"] = pd.to_datetime(only_day_df["Дата_операций"])
    only_day_df["День_недели"] = only_day_df["Дата_операций"].dt.day_name()
    only_day_df["Итого"] = only_day_df["Итого"].apply(lambda x: abs(x))

    reports_logger.info("Формирование статистики по тратам в выходной/рабочий день")
    grouped_df = only_day_df.groupby(["Дата_операций", "День_недели"]).agg({"Итого": "mean"})
    grouped_df.reset_index(level=[0, 1], drop=False, inplace=True)
    grouped_df["Итого"] = grouped_df["Итого"].astype(str).transform(quantize_decimal)
    grouped_df["Дата_операций"] = grouped_df["Дата_операций"].astype(str)

    reports_logger.info("Завершение <spending_by_workday>")
    return grouped_df
