import datetime
import os
from decimal import Decimal
from unittest.mock import patch

import pandas as pd
import pytest
import requests
from dotenv import load_dotenv
from pandas import DataFrame
from pandas.testing import assert_frame_equal
from pytest import CaptureFixture

from src.utils import (file_write_path, filter_choice_spending_workday, filter_user_choice, get_choice_menu,
                       get_currency_rate, get_currency_stock_prices, get_frame_expenses, get_frame_income,
                       get_stock_price, quantize_decimal)


def get_year_df() -> DataFrame:
    entire_df = pd.read_excel(file_write_path)
    return entire_df[3204:4947]  # 2019-12-29 22:28:13 / 2019-01-02 14:09:11


def get_month_df() -> DataFrame:
    entire_df = pd.read_excel(file_write_path)
    return entire_df[3998:4093]  # 2019-07-16 16:30:10 / 2019-07-01 12:40:51


def get_week_df() -> DataFrame:
    entire_df = pd.read_excel(file_write_path)
    return entire_df[4714:4730]  # 2019-03-09 19:47:28 / 2019-03-04 09:48:14


@pytest.mark.parametrize(
    "mock_data, chosen_date, chosen_range, expected_df",
    [
        (get_year_df(), "29.12.2019", "Y", get_year_df()),
        (get_year_df(), "16.07.2019", "M", get_month_df()),
        (get_year_df(), "09.03.2019", "W", get_week_df()),
    ],
)
def test_span_operations(mock_data: DataFrame, chosen_date: str, chosen_range: str, expected_df: DataFrame) -> None:
    with patch("src.utils.get_frame_operations", return_value=mock_data):
        from src.utils import get_span_operations

        result = get_span_operations(chosen_date, chosen_range)
        assert all(result == expected_df)


@pytest.mark.parametrize(
    "primary_df, output_dict",
    [
        (
            get_year_df(),
            {
                "total_amount": 1574796,
                "main": [
                    {"category": "Образование", "amount": 309557},
                    {"category": "Другое", "amount": 90205},
                    {"category": "Супермаркеты", "amount": 84307},
                    {"category": "Дом и ремонт", "amount": 83470},
                    {"category": "Различные товары", "amount": 57034},
                    {"category": "Фастфуд", "amount": 50161},
                    {"category": "Ж/д билеты", "amount": 45017},
                    {"category": "Остальное", "amount": 155158},
                ],
                "transfers_and_cash": [
                    {"category": "Переводы", "amount": 400702},
                    {"category": "Наличные", "amount": 299186},
                ],
            },
        ),
        (
            get_month_df(),
            {
                "total_amount": 66076,
                "main": [
                    {"category": "Фастфуд", "amount": 3428},
                    {"category": "Супермаркеты", "amount": 3140},
                    {"category": "Ж/д билеты", "amount": 1944},
                    {"category": "Образование", "amount": 1193},
                    {"category": "Связь", "amount": 1047},
                    {"category": "Аптеки", "amount": 929},
                    {"category": "Транспорт", "amount": 572},
                    {"category": "Остальное", "amount": 824},
                ],
                "transfers_and_cash": [
                    {"category": "Переводы", "amount": 50000},
                    {"category": "Наличные", "amount": 3000},
                ],
            },
        ),
        (
            get_week_df(),
            {
                "total_amount": 5509,
                "main": [
                    {"category": "Супермаркеты", "amount": 2134},
                    {"category": "Аптеки", "amount": 1500},
                    {"category": "Одежда и обувь", "amount": 850},
                    {"category": "Книги", "amount": 396},
                    {"category": "Фастфуд", "amount": 379},
                    {"category": "Связь", "amount": 250},
                    {"category": "Остальное", "amount": 0},
                ],
                "transfers_and_cash": [{"category": "Переводы", "amount": 0}, {"category": "Наличные", "amount": 0}],
            },
        ),
    ],
)
def test_frame_expenses(primary_df: DataFrame, output_dict: dict) -> None:
    result = get_frame_expenses(primary_df)
    assert result == output_dict


@pytest.mark.parametrize(
    "income_df, final_dict",
    [
        (
            get_year_df(),
            {
                "total_amount": 1725053,
                "main": [
                    {"category": "Пополнения", "amount": 873959},
                    {"category": "Переводы", "amount": 516835},
                    {"category": "Зарплата", "amount": 313200},
                    {"category": "Бонусы", "amount": 18009},
                    {"category": "Ж/д билеты", "amount": 2052},
                    {"category": "Транспорт", "amount": 700},
                    {"category": "Сервис", "amount": 299},
                ],
            },
        ),
        (
            get_month_df(),
            {
                "total_amount": 34183,
                "main": [
                    {"category": "Пополнения", "amount": 32950},
                    {"category": "Ж/д билеты", "amount": 785},
                    {"category": "Бонусы", "amount": 447},
                ],
            },
        ),
        (get_month_df()[-49:-40], {"total_amount": 0, "main": []}),
    ],
)
def test_frame_income(income_df: DataFrame, final_dict: dict) -> None:
    result = get_frame_income(income_df)
    assert result == final_dict


def test_currencies_stocks(get_data_string_july: str, get_stock_curr_dict: dict) -> None:
    result = get_currency_stock_prices(get_data_string_july)
    assert result == get_stock_curr_dict


def test_get_connection_error(get_date_time_obj: datetime.datetime, get_currency_string: str) -> None:
    load_dotenv()
    stated_date_obj = get_date_time_obj.date()
    stated_next_date_obj = stated_date_obj + datetime.timedelta(days=1)
    stated_next_date_str = stated_next_date_obj.strftime("%Y-%m-%d")
    with patch("src.utils.requests.get") as mock_get:
        mock_get.side_effect = requests.exceptions.ConnectionError("Network problem")

        result = get_currency_rate(get_date_time_obj, get_currency_string)
        assert result == "Connection failed"
        mock_get.assert_called_once_with(
            f'https://api.twelvedata.com/time_series?apikey={os.getenv("API_KEY")}&interval=1day&symbol='
            f"{get_currency_string}/RUB&format=JSON&start_date={get_date_time_obj}&type=stock&dp=2&end_date="
            f"{stated_next_date_str} 23:59:00"
        )


def test_get_timeout_error(get_date_time_obj: datetime.datetime, get_currency_string: str) -> None:
    load_dotenv()
    stated_date_obj = get_date_time_obj.date()
    stated_next_date_obj = stated_date_obj + datetime.timedelta(days=1)
    stated_next_date_str = stated_next_date_obj.strftime("%Y-%m-%d")
    with patch("src.utils.requests.get") as mock_get:
        mock_get.side_effect = requests.exceptions.Timeout("Request timeout")

        result = get_currency_rate(get_date_time_obj, get_currency_string)
        assert result == "Request timed out"
        mock_get.assert_called_once_with(
            f'https://api.twelvedata.com/time_series?apikey={os.getenv("API_KEY")}&interval=1day&symbol='
            f"{get_currency_string}/RUB&format=JSON&start_date={get_date_time_obj}&type=stock&dp=2&end_date="
            f"{stated_next_date_str} 23:59:00"
        )


def test_get_http_error(get_date_time_obj: datetime.datetime, get_currency_string: str) -> None:
    load_dotenv()
    stated_date_obj = get_date_time_obj.date()
    stated_next_date_obj = stated_date_obj + datetime.timedelta(days=1)
    stated_next_date_str = stated_next_date_obj.strftime("%Y-%m-%d")
    with patch("src.utils.requests.get") as mock_get:
        mock_response = mock_get.return_value
        mock_response.status_code = 404
        mock_response.raise_for_status.side_effect = requests.exceptions.HTTPError(
            "404 Client Error", response=mock_response
        )

        result = get_currency_rate(get_date_time_obj, get_currency_string)
        assert result == "HTTP error: 404"
        mock_get.assert_called_once_with(
            f'https://api.twelvedata.com/time_series?apikey={os.getenv("API_KEY")}&interval=1day&symbol='
            f"{get_currency_string}/RUB&format=JSON&start_date={get_date_time_obj}&type=stock&dp=2&end_date="
            f"{stated_next_date_str} 23:59:00"
        )


@pytest.mark.parametrize(
    "target_date, target_currency, mock_value",
    [
        (datetime.datetime(2019, 4, 25), "CNY", {"values": [{"close": "9.59"}]}),
        (datetime.datetime(2019, 4, 26), "CNY", {"values": [{"close": "9.62"}]}),
    ],
)
def test_get_cny_currency(target_date: datetime.datetime, target_currency: str, mock_value: dict) -> None:
    with patch("src.utils.json.load") as mock_json_load:
        mock_json_load.return_value = mock_value

        result = get_currency_rate(target_date, target_currency)
        mock_json_load.assert_called_once()
        assert result == mock_value["values"][0]["close"]


def test_stock_connection_error(get_date_time_obj: datetime.datetime, get_stock_string: str) -> None:
    load_dotenv()
    stated_date_obj = get_date_time_obj.date()
    stated_next_date_obj = stated_date_obj + datetime.timedelta(days=1)
    stated_next_date_str = stated_next_date_obj.strftime("%Y-%m-%d")
    with patch("src.utils.requests.get") as mock_get:
        mock_get.side_effect = requests.exceptions.ConnectionError("Network problem")

        result = get_stock_price(get_date_time_obj, get_stock_string)
        assert result == "Connection failed"
        mock_get.assert_called_once_with(
            f'https://api.twelvedata.com/time_series?apikey={os.getenv("API_KEY")}&interval=1day&symbol='
            f"{get_stock_string}&format=JSON&start_date={get_date_time_obj}&type=stock&dp=2&end_date="
            f"{stated_next_date_str} 23:59:00"
        )


def test_stock_timeout_error(get_date_time_obj: datetime.datetime, get_stock_string: str) -> None:
    load_dotenv()
    stated_date_obj = get_date_time_obj.date()
    stated_next_date_obj = stated_date_obj + datetime.timedelta(days=1)
    stated_next_date_str = stated_next_date_obj.strftime("%Y-%m-%d")
    with patch("src.utils.requests.get") as mock_get:
        mock_get.side_effect = requests.exceptions.Timeout("Request timeout")

        result = get_stock_price(get_date_time_obj, get_stock_string)
        assert result == "Request timed out"
        mock_get.assert_called_once_with(
            f'https://api.twelvedata.com/time_series?apikey={os.getenv("API_KEY")}&interval=1day&symbol='
            f"{get_stock_string}&format=JSON&start_date={get_date_time_obj}&type=stock&dp=2&end_date="
            f"{stated_next_date_str} 23:59:00"
        )


def test_stock_http_error(get_date_time_obj: datetime.datetime, get_stock_string: str) -> None:
    load_dotenv()
    stated_date_obj = get_date_time_obj.date()
    stated_next_date_obj = stated_date_obj + datetime.timedelta(days=1)
    stated_next_date_str = stated_next_date_obj.strftime("%Y-%m-%d")
    with patch("src.utils.requests.get") as mock_get:
        mock_response = mock_get.return_value
        mock_response.status_code = 404
        mock_response.raise_for_status.side_effect = requests.exceptions.HTTPError(
            "404 Client Error", response=mock_response
        )

        result = get_stock_price(get_date_time_obj, get_stock_string)
        assert result == "HTTP error: 404"
        mock_get.assert_called_once_with(
            f'https://api.twelvedata.com/time_series?apikey={os.getenv("API_KEY")}&interval=1day&symbol='
            f"{get_stock_string}&format=JSON&start_date={get_date_time_obj}&type=stock&dp=2&end_date="
            f"{stated_next_date_str} 23:59:00"
        )


@pytest.mark.parametrize(
    "chosen_date, chosen_range, expected_tuple",
    [
        ("29.12.2019", "Y", ("29.12.2019", "Y")),
        ("16.07.2019", "M", ("16.07.2019", "M")),
        ("09.03.2019", "W", ("09.03.2019", "W")),
    ],
)
def test_filter_user_choice(chosen_date: str, chosen_range: str, expected_tuple: tuple) -> None:
    with patch("builtins.input") as mock_input:
        mock_input.side_effect = [chosen_date, chosen_range]
        result = filter_user_choice()
        expected = expected_tuple

        assert mock_input.call_count == 2

        first_call = mock_input.call_args_list[0]
        args, kwargs = first_call
        assert args == ()

        second_call = mock_input.call_args_list[1]
        args, kwargs = second_call
        assert args == ()
        assert result == expected


@pytest.mark.usefixtures("capsys")
def test_invalid_input_date(
    capsys: CaptureFixture[str], get_non_format_date: str, get_month_range: str, get_no_answer: str
) -> None:
    inputs = [get_non_format_date, get_month_range, get_no_answer]

    with patch("builtins.input", side_effect=inputs):
        with pytest.raises(SystemExit) as exc_info:
            filter_user_choice()

    assert exc_info.value.code == 0
    captured = capsys.readouterr()
    assert "Указанная дата не существует или неверный формат даты" in captured.out


@pytest.mark.usefixtures("capsys")
def test_outrange_input_date(
    capsys: CaptureFixture[str], get_outside_range_date: str, get_week_range: str, get_no_answer: str
) -> None:
    inputs = [get_outside_range_date, get_week_range, get_no_answer]

    with patch("builtins.input", side_effect=inputs):
        with pytest.raises(SystemExit) as exc_info:
            filter_user_choice()

    assert exc_info.value.code == 0
    captured = capsys.readouterr()
    assert "Указанная дата недоступна или не выбран диапазон дат" in captured.out


def test_filter_spending_workday(get_empty_input: str, get_none_input_list: list) -> None:
    user_input = ""
    with patch("builtins.input", return_value=user_input):
        result = filter_choice_spending_workday()

    expected = get_none_input_list[0]
    assert_frame_equal(result[0], expected)


@pytest.mark.usefixtures("capsys")
def test_invalid_input_choice_workday(
    capsys: CaptureFixture[str], get_non_format_date: str, get_no_answer: str
) -> None:
    inputs = [get_non_format_date, get_no_answer]

    with patch("builtins.input", side_effect=inputs):
        with pytest.raises(SystemExit) as exc_info:
            filter_choice_spending_workday()

    assert exc_info.value.code == 0
    captured = capsys.readouterr()
    assert (
        "Введите дату вида 'dd.mm.YYYY':\n"
        "\n"
        "Указанной даты не существует или неверный формат даты\n" in captured.out
    )


@pytest.mark.usefixtures("capsys")
def test_unreachable_input_choice_workday(
    capsys: CaptureFixture[str], get_outside_range_date: str, get_no_answer: str
) -> None:
    inputs = [get_outside_range_date, get_no_answer]

    with patch("builtins.input", side_effect=inputs):
        with pytest.raises(SystemExit) as exc_info:
            filter_choice_spending_workday()

    assert exc_info.value.code == 0
    captured = capsys.readouterr()
    assert "Введите дату вида 'dd.mm.YYYY':\n" "\n" "Указанная дата недоступна\n" in captured.out


@pytest.mark.parametrize(
    "chosen_num, output_num",
    [
        ("1", "1"),
        ("2", "2"),
        ("3", "3"),
    ],
)
def test_user_choice_menu(chosen_num: str, output_num: str) -> None:
    with patch("builtins.input") as mock_input:
        mock_input.return_value = chosen_num
        result = get_choice_menu()
        expected = output_num

        assert mock_input.call_count == 1

        first_call = mock_input.call_args_list[0]
        args, kwargs = first_call
        assert args == ()
        assert result == expected


def test_get_quantize_decimal(get_float_number: float, get_decimal_quantized: Decimal) -> None:
    result = quantize_decimal(get_float_number)

    assert result == get_decimal_quantized
