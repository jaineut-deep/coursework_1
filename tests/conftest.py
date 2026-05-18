import datetime
import os
import pytest

trans_data_path = os.path.dirname(os.path.dirname(__file__)) + "/data/operations.xlsx"
refined_data_path = os.path.dirname(os.path.dirname(__file__)) + "/refined.xlsx"


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
def get_currency_string() -> str:
    return "EUR"


@pytest.fixture
def get_stock_string() -> str:
    return "AAPL"


@pytest.fixture
def get_date_time_obj() -> datetime.datetime:
    return datetime.datetime(2019, 9, 11)
