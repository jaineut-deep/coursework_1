import json
import logging

from src.utils import base_dir, get_currency_stock_prices, get_frame_expenses, get_frame_income, get_span_operations

path_to_log = base_dir + "/logs/module_views.log"
views_logger = logging.getLogger(__name__)
if not views_logger.handlers:
    file_handler = logging.FileHandler(path_to_log, mode="w")
    file_formatter = logging.Formatter("%(asctime)s - %(module)s - %(levelname)s - %(message)s")
    file_handler.setFormatter(file_formatter)
    views_logger.addHandler(file_handler)
    views_logger.setLevel(logging.DEBUG)


def aggregate_json(chosen_date: str, range_date: str = "M") -> str:
    """
    Функция принимает дату и обозначение временного диапазона в виде строки, а возвращает json-строку.
    """

    views_logger.info("Запуск <aggregate_json>")
    sector_dates_df = get_span_operations(chosen_date, range_date)
    expenses_dict = get_frame_expenses(sector_dates_df)
    income_dict = get_frame_income(sector_dates_df)
    curr_stock_prices = get_currency_stock_prices(chosen_date)

    views_logger.info("Формирование общей статистики по транзакциям")
    events_response = {
        "expenses": expenses_dict,
        "income": income_dict,
        "currency_rates": curr_stock_prices["currency_rates"],
        "stock_prices": curr_stock_prices["stock_prices"],
    }
    json_data = json.dumps(obj=events_response, ensure_ascii=False, indent=4)

    views_logger.info("Завершение <aggregate_json>")
    return json_data
