import datetime
import json
import os
import pandas as pd
import requests
import sys
from requests.exceptions import ConnectionError, Timeout, HTTPError
from decimal import Decimal, ROUND_HALF_UP, ROUND_HALF_EVEN
from dotenv import load_dotenv
from pandas import DataFrame

past_currency = os.path.dirname(os.path.dirname(__file__)) + "/safe_currency.json"
user_settings_path = os.path.dirname(os.path.dirname(__file__)) + "/user_settings.json"
data_path = os.path.dirname(os.path.dirname(__file__)) + "/data/operations.xlsx"


def get_frame_operations() -> DataFrame:
    """
    Функция возвращает отфильтрованные данные всех расходов и поступлений по данным о транзакциях клиента.
    """

    all_operations_df = pd.read_excel(data_path)
    all_operations_df.rename(columns=lambda x: x.replace(" ", "_") if isinstance(x, str) else x, inplace=True)
    executed_operations_df = all_operations_df[all_operations_df["Статус"] == "OK"]
    executed_operations_df.drop(
        ["Дата_платежа", "Номер_карты", "Статус", "MCC", "Бонусы_(включая_кэшбэк)", "Округление_на_инвесткопилку"],
        axis=1,
        inplace=True,
    )
    executed_operations_df["Дата_операции"] = pd.to_datetime(executed_operations_df["Дата_операции"], dayfirst=True)
    indexed_executed_df = executed_operations_df.reset_index(drop=True)
    inner_transfer_index_list = []
    for idx in range(len(indexed_executed_df) - 2):
        if (
            (indexed_executed_df.iloc[idx, 1] != indexed_executed_df.iloc[(idx + 1), 1])
            and (abs(indexed_executed_df.iloc[idx, 1]) == abs(indexed_executed_df.iloc[(idx + 1), 1]))
            and (indexed_executed_df.iloc[idx, 7] == indexed_executed_df.iloc[(idx + 1), 7])
        ):
            inner_transfer_index_list.extend([idx, (idx + 1)])
    charge_payment_df = indexed_executed_df.loc[~(indexed_executed_df.index.isin(inner_transfer_index_list))]
    indexed_df = charge_payment_df.reset_index(drop=True)
    indexed_df["Сумма_операции"] = indexed_df["Сумма_операции"].transform(
        lambda x: Decimal(str(x)).quantize(Decimal("0.00"))   # type: ignore
    )
    indexed_df["Сумма_платежа"] = indexed_df["Сумма_платежа"].transform(
        lambda x: Decimal(str(x)).quantize(Decimal("0.00"))   # type: ignore
    )
    indexed_df["Сумма_операции_с_округлением"] = indexed_df["Сумма_операции_с_округлением"].transform(
        lambda x: Decimal(str(x)).quantize(Decimal("0.00"))  # type: ignore
    )
    mask_currency_ruble = (indexed_df["Валюта_операции"] != "RUB") & (indexed_df["Валюта_платежа"] == "RUB")
    mask_ruble_currency = (indexed_df["Валюта_операции"] == "RUB") & (indexed_df["Валюта_платежа"] != "RUB")
    mask_ruble_positive = (
        (indexed_df["Валюта_операции"] == "RUB")
        & (indexed_df["Валюта_платежа"] == "RUB")
        & (indexed_df["Сумма_операции"] > 0)
    )
    mask_ruble_negative = (
        (indexed_df["Валюта_операции"] == "RUB")
        & (indexed_df["Валюта_платежа"] == "RUB")
        & (indexed_df["Сумма_операции"] < 0)
    )
    indexed_df.loc[mask_currency_ruble, "Итого"] = indexed_df.loc[mask_currency_ruble, "Сумма_платежа"]
    indexed_df.loc[mask_ruble_currency, "Итого"] = indexed_df.loc[mask_ruble_currency, "Сумма_операции"]
    indexed_df.loc[mask_ruble_positive, "Итого"] = indexed_df.loc[mask_ruble_positive, "Сумма_операции_с_округлением"]
    indexed_df.loc[mask_ruble_negative, "Итого"] = indexed_df["Сумма_операции_с_округлением"].where(
        cond=~mask_ruble_negative, other=lambda x: x * (-1)
    )
    for idx in range(len(indexed_df) - 1):
        if (indexed_df.loc[idx, "Валюта_операции"] != "RUB") & (indexed_df.loc[idx, "Валюта_платежа"] != "RUB"):
            date_obj = indexed_df["Дата_операции"].iloc[idx].to_pydatetime()
            rate = Decimal(get_currency_rate(date_obj, str(indexed_df.loc[idx, "Валюта_операции"])))
            indexed_df.loc[idx, "Итого"] = (   # type: ignore
                rate * Decimal(str(indexed_df.loc[idx, "Сумма_операции"]))
            ).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP)
    return indexed_df


def get_span_operations(spec_date: str, spec_range: str) -> DataFrame:
    """
    Функция принимает дату и обозначение диапазона дат в виде строк, а возвращает данные по транзакциям клиента за
    указанный диапазон дат.
    """

    primary_df = get_frame_operations()
    spec_date_obj = datetime.datetime.strptime(spec_date, "%d.%m.%Y")
    spec_date_format = spec_date_obj.replace(spec_date_obj.year, spec_date_obj.month, spec_date_obj.day)
    spec_date_full = spec_date_format.replace(hour=23, minute=59, second=59)
    days_kit = {
        "W": (spec_date_format - datetime.timedelta(days=(spec_date_format.isoweekday() - 1))),
        "M": spec_date_format.replace(spec_date_format.year, spec_date_format.month, 1),
        "Y": spec_date_format.replace(spec_date_format.year, 1, 1),
        "ALL": primary_df["Дата_операции"].tail(1).iloc[0].to_pydatetime(),
    }
    date_border = days_kit[spec_range]
    at_date_df = primary_df.loc[primary_df["Дата_операции"] <= spec_date_full]
    return at_date_df.loc[at_date_df["Дата_операции"] >= pd.Timestamp(date_border)]


def get_frame_expenses(operations_df: DataFrame) -> dict:
    """
    Функция принимает данные по транзакциям, а возвращает данные о сумме трат всего и по категориям за определенный
    период времени.
    """

    selected_expenses_df = operations_df.loc[operations_df.Итого < 0]
    selected_indexed_df = selected_expenses_df.reset_index(drop=True)
    selected_indexed_df["Итого"] = selected_indexed_df["Итого"].transform(
        lambda x: Decimal(str(x)).quantize(Decimal("0.00")))   # type: ignore
    total_sum_df = selected_indexed_df.pivot_table(index=["Категория"], values="Итого", aggfunc="sum")
    total_sum_df.sort_values(by="Итого", ascending=True, inplace=True)
    expenses_abs_df = total_sum_df.apply(pd.Series.abs)
    expenses_total = int(
        Decimal(float(expenses_abs_df.apply(pd.Series.sum).values[0])).quantize(Decimal("1."), ROUND_HALF_EVEN)
    )
    expenses_all_dict = expenses_abs_df.to_dict()["Итого"]
    remit_value = int(expenses_all_dict.pop("Переводы", Decimal("0")).quantize(Decimal("1."), ROUND_HALF_EVEN))
    cash_value = int(expenses_all_dict.pop("Наличные", Decimal("0")).quantize(Decimal("1."), ROUND_HALF_EVEN))
    trans_cash_list = [{"category": "Переводы", "amount": remit_value}, {"category": "Наличные", "amount": cash_value}]
    trans_cash_list.sort(key=lambda x: (x["amount"], x["category"]), reverse=True)
    expenses_collection = list(expenses_all_dict.items())
    expenses_list_rounded = list(
        map(lambda x: (x[0], int(x[1].quantize(Decimal("1."), ROUND_HALF_EVEN))), expenses_collection)
    )
    main_expenses_list = [{"category": key, "amount": value} for key, value in expenses_list_rounded[:7]]
    expenses_remaining = (
        sum([twain[1] for twain in expenses_collection[7:]]) if len(expenses_collection) > 7 else Decimal("0")
    )
    expenses_remaining_rounded = int(expenses_remaining.quantize(Decimal("1."), ROUND_HALF_EVEN))
    main_expenses_list.append({"category": "Остальное", "amount": expenses_remaining_rounded})
    expenses_category_dict = {
        "total_amount": expenses_total,
        "main": main_expenses_list,
        "transfers_and_cash": trans_cash_list,
    }
    return expenses_category_dict


def get_frame_income(operations_selected_df: DataFrame) -> dict:
    """
    Функция принимает данные по транзакциям, а возвращает данные о сумме поступлений всего и по категориям за
    определенный период времени.
    """

    selected_expenses_df = operations_selected_df.loc[operations_selected_df.Итого > 0]
    selected_indexed_df = selected_expenses_df.reset_index(drop=True)
    selected_indexed_df["Итого"] = selected_indexed_df["Итого"].transform(
        lambda x: Decimal(str(x)).quantize(Decimal("0.00")))   # type: ignore
    total_sum_df = selected_indexed_df.pivot_table(index=["Категория"], values="Итого", aggfunc="sum")
    if total_sum_df.shape == (0, 0):
        return {"total_amount": 0, "main": []}
    else:
        total_sum_df.sort_values(by="Итого", ascending=False, inplace=True)
        income_abs_df = total_sum_df.apply(pd.Series.abs)
        income_total = int(
            Decimal(float(income_abs_df.apply(pd.Series.sum).values[0])).quantize(Decimal("1."), ROUND_HALF_EVEN)
        )
        income_all_dict = income_abs_df.to_dict()["Итого"]
        income_collection = list(income_all_dict.items())
        income_list_rounded = list(
            map(lambda x: (x[0], int(x[1].quantize(Decimal("1."), ROUND_HALF_EVEN))), income_collection)
        )
        main_income_list = [{"category": key, "amount": value} for key, value in income_list_rounded]
        income_category_dict = {
            "total_amount": income_total,
            "main": main_income_list,
        }
        return income_category_dict


def get_currency_stock_prices(user_date: str) -> dict:
    """
    Функция принимает дату в виде строки, а возвращает словарь с данными по курсам валют и акций за указанную дату.
    """

    with open(user_settings_path) as cur_data:
        curr_stock_dict = json.load(cur_data)
    spec_date_obj = datetime.datetime.strptime(user_date, "%d.%m.%Y")
    spec_date_format = spec_date_obj.replace(spec_date_obj.year, spec_date_obj.month, spec_date_obj.day)
    currencies_list = []
    for idx in curr_stock_dict["user_currencies"]:
        currencies_list.append({"currency": idx, "rate": float(get_currency_rate(spec_date_format, idx))})
    stocks_list = []
    for idx in curr_stock_dict["user_stocks"]:
        stocks_list.append({"stock": idx, "price": float(get_stock_price(spec_date_format, idx))})
    currency_stock = {"currency_rates": currencies_list, "stock_prices": stocks_list}
    return currency_stock


def get_currency_rate(stated_date: datetime.datetime, stated_currency: str) -> str:
    """
    Функция принимает дату и валюту в виде строки для конвертации, а возвращает курс валюты в рублях.
    """

    with open(past_currency) as cur_data:
        past_data = json.load(cur_data)
    load_dotenv()
    twelvedata_api_key = os.getenv("API_KEY")
    stated_date_obj = stated_date.date()
    stated_next_date_obj = stated_date_obj + datetime.timedelta(days=1)
    stated_date_str = stated_date_obj.strftime("%Y-%m-%d")
    stated_next_date_str = stated_next_date_obj.strftime("%Y-%m-%d")
    try:
        response = requests.get(
            f"https://api.twelvedata.com/time_series?apikey={twelvedata_api_key}&interval=1day&symbol="
            f"{stated_currency}/RUB&format=JSON&start_date={stated_date_str} 00:00:00&type=stock&dp=2&end_date="
            f"{stated_next_date_str} 23:59:00"
        )
        if response.status_code == 404:
            raise HTTPError(f"HTTP error: {response.status_code}")
    except ConnectionError:
        return "Connection failed"
    except HTTPError as err:
        return f"{err}"
    except Timeout:
        return "Request timed out"
    status_code = response.status_code
    result_value = "0.01"
    if status_code == 200:
        currency_values = response.json().get("values", past_data["values"])
        for day_price in currency_values:
            if day_price["datetime"] == stated_date_str:
                result_value = day_price["close"]
            elif day_price["datetime"] == stated_next_date_str:
                result_value = day_price["close"]
            else:
                result_value = "0.02"
        return result_value
    else:
        past_cur_values = past_data["values"]
        for day_price in past_cur_values:
            if day_price["datetime"] == stated_date_str:
                result_value = day_price["close"]
            elif day_price["datetime"] == stated_next_date_str:
                result_value = day_price["close"]
        return result_value


def get_stock_price(indicated_date: datetime.datetime, indicated_stock: str) -> str:
    """
    Функция принимает дату и акцию в виде строки для конвертации, а возвращает курс акции в долларах(USD).
    """

    load_dotenv()
    twelvedata_api_key = os.getenv("API_KEY")
    indicated_date_obj = indicated_date.date()
    stated_next_date_obj = indicated_date_obj + datetime.timedelta(days=1)
    stated_date_str = indicated_date_obj.strftime("%Y-%m-%d")
    stated_next_date_str = stated_next_date_obj.strftime("%Y-%m-%d")
    try:
        response = requests.get(
            f"https://api.twelvedata.com/time_series?apikey={twelvedata_api_key}&interval=1day&symbol="
            f"{indicated_stock}&format=JSON&start_date={stated_date_str} 00:00:00&type=stock&dp=2&end_date="
            f"{stated_next_date_str} 23:59:00"
        )
        if response.status_code == 404:
            raise HTTPError(f"HTTP error: {response.status_code}")
    except ConnectionError:
        return "Connection failed"
    except HTTPError as err:
        return f"{err}"
    except Timeout:
        return "Request timed out"
    status_code = response.status_code
    stock_string = "0.01"
    if status_code == 200:
        currency_values = response.json().get("values", [{"datetime": stated_date_str, "close": "0.00"}])
        for day_price in currency_values:
            if day_price["datetime"] == stated_date_str:
                stock_string = day_price["close"]
            elif day_price["datetime"] == stated_next_date_str:
                stock_string = day_price["close"]
            else:
                stock_string = "0.02"
        return stock_string
    else:
        return stock_string


def filter_user_choice() -> tuple[str, str]:
    """
    Функция возвращает кортеж строк: дату и обозначение диапазона дат.
    """

    while True:
        print(f"Введите дату вида 'dd.mm.YYYY':\n")
        user_date = input()
        print(
            """
        Выберите диапазон дат:
        W -- данные за текущую неделю
        M -- данные за текущий месяц
        Y -- данные за текущий год
        ALL -- данные за всё время
        """
        )
        user_range = input()
        term_dates = ()
        try:
            datetime.datetime.strptime(user_date, "%d.%m.%Y")

            if (datetime.datetime.strptime(user_date, "%d.%m.%Y") and
                    (user_range.upper() in ["W", "M", "Y", "ALL"])):
                term_dates = (user_date, user_range.upper())
                break
        except ValueError:
            print("Указанной даты не существует")
        more_input = input(f"Неверный формат введенных данных\nПопробуете снова? Y/N\n")
        if more_input.upper() == "N":
            sys.exit(0)
        elif more_input.upper() == "Y":
            continue
        else:
            continue
    return term_dates
