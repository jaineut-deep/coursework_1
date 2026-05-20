import datetime
import pandas as pd
from pandas import DataFrame
from src.utils import file_write_path, quantize_decimal
from typing import Optional


def spending_by_workday(transactions: pd.DataFrame, date: Optional[str] = None) -> DataFrame:
    """
     Функция принимает транзакции в виде DataFrame и опциональную дату в виде строки, а возвращает данные по средним
     тратам в виде DataFrame.
    """

    if date is None:
        date = "31.12.2021"
    date_obj = datetime.datetime.strptime(date, "%d.%m.%Y")
    date_full_obj = date_obj.replace(hour=23, minute=59, second=59)
    date_border = date_obj - datetime.timedelta(days=89)

    spent_all_df = transactions.loc[transactions.Итого < 0]
    spent_indexed_df = spent_all_df.reset_index(drop=True)
    spent_indexed_df["Итого"] = spent_indexed_df["Итого"].astype(str).transform(quantize_decimal)

    sector_date_df = spent_indexed_df[(spent_indexed_df["Дата_операции"] <= pd.Timestamp(date_full_obj)) &
                                     (pd.Timestamp(date_border) <= spent_indexed_df["Дата_операции"])]

    only_day_df = sector_date_df[["Дата_операции", "Итого"]]
    only_day_df.rename(columns={'Дата_операции': 'Дата_операций'}, inplace=True)
    only_day_df["Дата_операций"] = only_day_df["Дата_операций"].apply(lambda x: x.date())
    if not pd.api.types.is_datetime64_any_dtype(only_day_df["Дата_операций"]):
        only_day_df["Дата_операций"] = pd.to_datetime(only_day_df["Дата_операций"])
    only_day_df["День_недели"] = only_day_df["Дата_операций"].dt.day_name()
    only_day_df["Итого"] = only_day_df["Итого"].apply(lambda x: abs(x))

    grouped_df = only_day_df.groupby(["Дата_операций", "День_недели"]).agg({"Итого": "mean"})
    grouped_df.reset_index(level=[0, 1], drop=False, inplace=True)
    grouped_df["Итого"] = grouped_df["Итого"].astype(str).transform(quantize_decimal)

    return grouped_df


if __name__ == "__main__":
    print(spending_by_workday(pd.read_excel(file_write_path)))
