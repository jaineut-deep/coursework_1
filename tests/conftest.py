import datetime
import json
import os
import pandas as pd
import pytest
from decimal import Decimal
from pandas import DataFrame

trans_data_path = os.path.dirname(os.path.dirname(__file__)) + "/data/operations.xlsx"
refined_data_path = os.path.dirname(os.path.dirname(__file__)) + "/buffer_data/refined.xlsx"
agg_private_transfer_path = os.path.dirname(os.path.dirname(__file__)) + "/materials_testing/agg_private_transfer.json"
safe_currency_path = os.path.dirname(os.path.dirname(__file__)) + "/materials_testing/safe_currency.json"


@pytest.fixture
def get_empty_input() -> str:
    return ""


@pytest.fixture
def get_no_answer() -> str:
    return "N"


@pytest.fixture
def get_week_range() -> str:
    return "W"


@pytest.fixture
def get_month_range() -> str:
    return "M"


@pytest.fixture
def get_data_string_july() -> str:
    return "09.07.2019"


@pytest.fixture
def get_early_date() -> str:
    return "31.03.2018"


@pytest.fixture
def get_non_format_date() -> str:
    return "20/08/2020"


@pytest.fixture
def get_feb_wrong_day() -> str:
    return "29.02.2018"


@pytest.fixture
def get_outside_range_date() -> str:
    return "01.01.2025"


@pytest.fixture
def get_none_input_list() -> list:
    prepare_rub_df = pd.read_excel(refined_data_path)
    return [prepare_rub_df, None]


@pytest.fixture
def get_float_number() -> float:
    return 3.7567


@pytest.fixture
def get_decimal_quantized() -> Decimal:
    return Decimal('3.76')


@pytest.fixture
def get_stock_curr_dict() -> dict:
    curr_stock_prices = {
        "currency_rates": [{"currency": "USD", "rate": 63.65}, {"currency": "EUR", "rate": 71.6}],
        "stock_prices": [
            {"stock": "AAPL", "price": 50.31},
            {"stock": "AMZN", "price": 99.42},
            {"stock": "GOOGL", "price": 56.21},
            {"stock": "MSFT", "price": 136.46},
            {"stock": "TSLA", "price": 15.34},
        ],
    }
    return curr_stock_prices


@pytest.fixture
def get_cny_currency_str() -> str:
    return "CNY"


@pytest.fixture
def get_currency_string() -> str:
    return "EUR"


@pytest.fixture
def get_stock_string() -> str:
    return "AAPL"


@pytest.fixture
def get_date_time_obj() -> datetime.datetime:
    return datetime.datetime(2019, 9, 11)


@pytest.fixture
def get_date_april_obj() -> datetime.datetime:
    return datetime.datetime(2019, 4, 24)


@pytest.fixture
def get_date_april_ready() -> datetime.datetime:
    return datetime.datetime(2019, 4, 26)


@pytest.fixture
def get_trial_df() -> DataFrame:
    trial_df = pd.read_excel(refined_data_path)
    return trial_df.tail(600)


@pytest.fixture
def get_output_agg_json() -> str:
    output_json = (
        "{\n"
        '    "expenses": {\n'
        '        "total_amount": 1080,\n'
        '        "main": [\n'
        "            {\n"
        '                "category": "Супермаркеты",\n'
        '                "amount": 385\n'
        "            },\n"
        "            {\n"
        '                "category": "Фастфуд",\n'
        '                "amount": 384\n'
        "            },\n"
        "            {\n"
        '                "category": "Транспорт",\n'
        '                "amount": 227\n'
        "            },\n"
        "            {\n"
        '                "category": "Образование",\n'
        '                "amount": 84\n'
        "            },\n"
        "            {\n"
        '                "category": "Остальное",\n'
        '                "amount": 0\n'
        "            }\n"
        "        ],\n"
        '        "transfers_and_cash": [\n'
        "            {\n"
        '                "category": "Переводы",\n'
        '                "amount": 0\n'
        "            },\n"
        "            {\n"
        '                "category": "Наличные",\n'
        '                "amount": 0\n'
        "            }\n"
        "        ]\n"
        "    },\n"
        '    "income": {\n'
        '        "total_amount": 0,\n'
        '        "main": []\n'
        "    },\n"
        '    "currency_rates": [\n'
        "        {\n"
        '            "currency": "USD",\n'
        '            "rate": 63.65\n'
        "        },\n"
        "        {\n"
        '            "currency": "EUR",\n'
        '            "rate": 71.6\n'
        "        }\n"
        "    ],\n"
        '    "stock_prices": [\n'
        "        {\n"
        '            "stock": "AAPL",\n'
        '            "price": 50.31\n'
        "        },\n"
        "        {\n"
        '            "stock": "AMZN",\n'
        '            "price": 99.42\n'
        "        },\n"
        "        {\n"
        '            "stock": "GOOGL",\n'
        '            "price": 56.21\n'
        "        },\n"
        "        {\n"
        '            "stock": "MSFT",\n'
        '            "price": 136.46\n'
        "        },\n"
        "        {\n"
        '            "stock": "TSLA",\n'
        '            "price": 15.34\n'
        "        }\n"
        "    ]\n"
        "}"
    )
    return output_json


@pytest.fixture
def get_private_trans_json() -> str:
    with open(agg_private_transfer_path, "r", encoding="utf-8") as file:
        data = json.load(file)
    json_data = json.dumps(data, ensure_ascii=False, indent=4)
    return json_data
