import json
from src.utils import get_span_operations, get_frame_expenses, get_frame_income, get_currency_stock_prices


def aggregate_json(chosen_date: str, range_date: str = "M") -> str:
    """
    Функция принимает дату и обозначение временного диапазона в виде строки, а возвращает json-строку.
    """

    sector_dates_df = get_span_operations(chosen_date, range_date)
    expenses_dict = get_frame_expenses(sector_dates_df)
    income_dict = get_frame_income(sector_dates_df)
    curr_stock_prices = get_currency_stock_prices(chosen_date)
    events_response = {"expenses": expenses_dict, "income": income_dict,
                       "currency_rates": curr_stock_prices["currency_rates"],
                       "stock_prices": curr_stock_prices["stock_prices"]}
    json_data = json.dumps(obj=events_response, ensure_ascii=False, indent=4)
    return json_data
