import datetime
import json
import logging
import os
import sys
from decimal import ROUND_HALF_EVEN, ROUND_HALF_UP, Decimal
from typing import Any, List

import pandas as pd
import requests
from dotenv import load_dotenv
from pandas import DataFrame
from requests.exceptions import ConnectionError, HTTPError, Timeout

base_dir = os.path.dirname(os.path.dirname(__file__))
past_currency = base_dir + "/materials_testing/safe_currency.json"
user_settings_path = base_dir + "/user_settings.json"
data_path = base_dir + "/data/operations.xlsx"
file_write_path = base_dir + "/buffer_data/refined.xlsx"
path_to_log = base_dir + "/logs/module_utils.log"
utils_logger = logging.getLogger(__name__)
if not utils_logger.handlers:
    file_handler = logging.FileHandler(path_to_log, mode="w")
    file_formatter = logging.Formatter("%(asctime)s - %(module)s - %(levelname)s - %(message)s")
    file_handler.setFormatter(file_formatter)
    utils_logger.addHandler(file_handler)
    utils_logger.setLevel(logging.DEBUG)


def get_frame_operations() -> DataFrame:
    """
    Функция возвращает отфильтрованные данные всех расходов и поступлений по данным о транзакциях клиента.
    """

    utils_logger.info(
        f"Начало работы функции <get_frame_operations> для нормализации начальной таблицы из файла <{data_path}>"
    )
    all_operations_df = pd.read_excel(data_path)
    utils_logger.info(f"Создан рабочий датафрейм <{all_operations_df}>")
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
    utils_logger.info("Начинаем приведение типов данных в рабочем датафрейме")
    for idx in range(len(indexed_executed_df) - 2):
        if (
            (indexed_executed_df.iloc[idx, 1] != indexed_executed_df.iloc[(idx + 1), 1])
            and (
                abs(indexed_executed_df["Сумма_операции"].astype(float).iloc[idx])
                == abs(indexed_executed_df["Сумма_операции"].astype(float).iloc[idx + 1])
            )
            and (indexed_executed_df.iloc[idx, 7] == indexed_executed_df.iloc[(idx + 1), 7])
        ):
            inner_transfer_index_list.extend([idx, (idx + 1)])
    charge_payment_df = indexed_executed_df.loc[~(indexed_executed_df.index.isin(inner_transfer_index_list))]
    indexed_df = charge_payment_df.reset_index(drop=True)
    indexed_df["Сумма_операции"] = indexed_df["Сумма_операции"].astype(str).transform(quantize_decimal)

    indexed_df["Сумма_платежа"] = indexed_df["Сумма_платежа"].astype(str).transform(quantize_decimal)

    indexed_df["Сумма_операции_с_округлением"] = (
        indexed_df["Сумма_операции_с_округлением"].astype(str).transform(quantize_decimal)
    )

    utils_logger.info("Типы данных в отдельных столбцах приведены")
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
    utils_logger.info(
        "Начало создания столбца 'Итого' с рублевой суммой по операциям в валюте с учётом дневного курса"
    )
    for idx in range(len(indexed_df) - 1):
        if (indexed_df.loc[idx, "Валюта_операции"] != "RUB") & (indexed_df.loc[idx, "Валюта_платежа"] != "RUB"):
            date_obj = indexed_df["Дата_операции"].iloc[idx].to_pydatetime()
            rate = Decimal(get_currency_rate(date_obj, str(indexed_df.loc[idx, "Валюта_операции"])))
            indexed_df.loc[idx, "Итого"] = (  # type: ignore
                rate * Decimal(str(indexed_df.loc[idx, "Сумма_операции"]))
            ).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP)
    utils_logger.info("Столбец 'Итого' с рублевой суммой добавлен")
    indexed_df.to_excel(excel_writer=file_write_path, index=False)
    utils_logger.info("<get_frame_operations> закончила свою работу")
    return indexed_df


def get_span_operations(spec_date: str, spec_range: str) -> DataFrame:
    """
    Функция принимает дату и обозначение диапазона дат в виде строк, а возвращает данные по транзакциям клиента за
    указанный диапазон дат.
    """

    utils_logger.info("Начало работы <get_span_operations> по нахождению транзакций определенного периода")
    primary_df = (
        pd.read_excel(file_write_path)
        if (os.path.isfile(file_write_path) and (os.path.splitext(file_write_path)[1].lower() == ".xlsx"))
        else get_frame_operations()
    )
    primary_df["Итого"] = primary_df["Итого"].astype(str).transform(quantize_decimal)

    spec_date_obj = datetime.datetime.strptime(spec_date, "%d.%m.%Y")
    spec_date_full = spec_date_obj.replace(hour=23, minute=59, second=59)
    utils_logger.info("Задание границ периода дат транзакций")
    days_kit = {
        "W": (spec_date_obj - datetime.timedelta(days=(spec_date_obj.isoweekday() - 1))),
        "M": spec_date_obj.replace(spec_date_obj.year, spec_date_obj.month, 1),
        "Y": spec_date_obj.replace(spec_date_obj.year, 1, 1),
        "ALL": primary_df["Дата_операции"].tail(1).iloc[0].to_pydatetime(),
    }
    date_border = days_kit[spec_range]
    at_dates_df = primary_df[
        (primary_df["Дата_операции"] <= pd.Timestamp(spec_date_full))
        & (pd.Timestamp(date_border) <= primary_df["Дата_операции"])
    ]
    utils_logger.info("Датафрейм за определенный период сформирован")

    utils_logger.info("<get_span_operations> закончила свою работу")
    return at_dates_df


def get_frame_expenses(operations_df: DataFrame) -> dict:
    """
    Функция принимает данные по транзакциям, а возвращает данные о сумме трат всего и по категориям за определенный
    период времени.
    """

    utils_logger.info("Начало работы <get_frame_expenses> для формирования статистики о тратах")
    selected_expenses_df = operations_df.loc[operations_df.Итого < 0]
    selected_indexed_df = selected_expenses_df.reset_index(drop=True)
    selected_indexed_df["Итого"] = selected_indexed_df["Итого"].transform(
        lambda x: Decimal(str(x)).quantize(Decimal("0.00"))  # type: ignore
    )
    total_sum_df = selected_indexed_df.pivot_table(index=["Категория"], values="Итого", aggfunc="sum")
    utils_logger.info("Создана сводная таблица с видами трат за отдельный период")

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

    utils_logger.info("Округление значений по категориям до двух знаков после запятой")
    expenses_list_rounded = list(
        map(lambda x: (x[0], int(x[1].quantize(Decimal("1."), ROUND_HALF_EVEN))), expenses_collection)
    )
    main_expenses_list = [{"category": key, "amount": value} for key, value in expenses_list_rounded[:7]]
    expenses_remaining = (
        sum([twain[1] for twain in expenses_collection[7:]]) if len(expenses_collection) > 7 else Decimal("0")
    )
    expenses_remaining_rounded = int(expenses_remaining.quantize(Decimal("1."), ROUND_HALF_EVEN))
    main_expenses_list.append({"category": "Остальное", "amount": expenses_remaining_rounded})
    utils_logger.info("Создание словаря со статистикой по тратам за данный период")
    expenses_category_dict = {
        "total_amount": expenses_total,
        "main": main_expenses_list,
        "transfers_and_cash": trans_cash_list,
    }

    utils_logger.info("<get_frame_expenses> завершила свою работу")
    return expenses_category_dict


def get_frame_income(operations_selected_df: DataFrame) -> dict:
    """
    Функция принимает данные по транзакциям, а возвращает данные о сумме поступлений всего и по категориям за
    определенный период времени.
    """

    utils_logger.info("Начало работы <get_frame_income> для сбора статистики о поступлениях")
    selected_expenses_df = operations_selected_df.loc[operations_selected_df.Итого > 0]
    selected_indexed_df = selected_expenses_df.reset_index(drop=True)
    selected_indexed_df["Итого"] = selected_indexed_df["Итого"].astype(str).transform(quantize_decimal)

    total_sum_df = selected_indexed_df.pivot_table(index=["Категория"], values="Итого", aggfunc="sum")
    if total_sum_df.shape == (0, 0):
        utils_logger.info("Поступления в данный период отсутствуют")
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

        utils_logger.info("В данный период были поступления")
        income_category_dict = {
            "total_amount": income_total,
            "main": main_income_list,
        }

        utils_logger.info("Завершение работы <get_frame_income>")
        return income_category_dict


def get_currency_stock_prices(user_date: str) -> dict:
    """
    Функция принимает дату в виде строки, а возвращает словарь с данными по курсам валют и акций за указанную дату.
    """

    utils_logger.info("Начало работы <get_currency_stock_prices> для отображения курсов валют и акций на дату")
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
    utils_logger.info("Создание отчета по валютам и акциям за указанную дату")

    currency_stock = {"currency_rates": currencies_list, "stock_prices": stocks_list}

    utils_logger.info("Завершение <get_currency_stock_prices>")
    return currency_stock


def get_currency_rate(stated_date: datetime.datetime, stated_currency: str) -> str:
    """
    Функция принимает дату и валюту в виде строки для конвертации, а возвращает курс валюты в рублях.
    """

    utils_logger.info("Запуск <get_currency_rate>")
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
    except ConnectionError as err:
        utils_logger.exception("Обнаружена ошибка: %s", err)
        return "Connection failed"
    except HTTPError as err:
        utils_logger.exception("Обнаружена ошибка: %s", err)
        return f"{err}"
    except Timeout as err:
        utils_logger.exception("Обнаружена ошибка: %s", err)
        return "Request timed out"
    status_code = response.status_code
    result_value = "0.01"
    if status_code == 200:
        utils_logger.info("Получение данных с сервера")
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
        utils_logger.info("Получение данных из файла")
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

    utils_logger.info("Запуск <get_stock_price>")
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
    except ConnectionError as err:
        utils_logger.exception("Обнаружена ошибка: %s", err)
        return "Connection failed"
    except HTTPError as err:
        utils_logger.exception("Обнаружена ошибка: %s", err)
        return f"{err}"
    except Timeout as err:
        utils_logger.exception("Обнаружена ошибка: %s", err)
        return "Request timed out"
    status_code = response.status_code
    stock_string = "0.01"
    if status_code == 200:
        utils_logger.info("Получение данных с сервера")
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


def filter_user_choice() -> List[str]:
    """
    Функция возвращает кортеж строк: пользовательские дату и обозначение диапазона дат.
    """

    while True:
        print("Введите дату вида 'dd.mm.YYYY':\n")
        user_date = input()
        print("""
        Выберите диапазон дат:
        W -- данные за неделю на которую приходится дата
        M -- данные за месяц на который приходится дата
        Y -- данные за год на который приходится дата
        ALL -- данные за всё время до указанной даты
        """)
        user_range = input()
        term_dates = ["", ""]
        right_edge_date = datetime.datetime(2021, 12, 31)
        left_edge_date = datetime.datetime(2018, 1, 1)
        try:
            trial_date = datetime.datetime.strptime(user_date, "%d.%m.%Y")
            if not trial_date:
                raise ValueError
            if (right_edge_date < trial_date) or (trial_date < left_edge_date):
                raise IndexError
            if datetime.datetime.strptime(user_date, "%d.%m.%Y") and (user_range.upper() in ["W", "M", "Y", "ALL"]):
                term_dates = [user_date, user_range.upper()]
                break
            else:
                raise IndexError
        except ValueError:
            print("Указанная дата не существует или неверный формат даты")
            if input("Попробуете снова? Y/N\n").upper() == "N":
                sys.exit(0)
            else:
                continue
        except IndexError:
            print("Указанная дата недоступна или не выбран диапазон дат")
            if input("Попробуете снова? Y/N\n").upper() == "N":
                sys.exit(0)
            else:
                continue
    return term_dates


def filter_choice_spending_workday() -> List[DataFrame | Any]:
    """
    Функция возвращает список содержащий Dataframe и пользовательский ввод даты в виде строки.
    """

    while True:
        print("Введите дату вида 'dd.mm.YYYY' или пропустите для установления даты по умолчанию:\n")
        user_date = input()
        if user_date == "":
            break
        right_edge_date = datetime.datetime(2021, 12, 31)
        left_edge_date = datetime.datetime(2018, 3, 31)
        try:
            trial_date = datetime.datetime.strptime(user_date, "%d.%m.%Y")
            if not trial_date:
                raise ValueError
            if (right_edge_date < trial_date) or (trial_date < left_edge_date):
                raise IndexError

            if trial_date:
                break
        except ValueError:
            print("Указанной даты не существует или неверный формат даты")
            if input("Попробуете снова? Y/N\n").upper() == "N":
                sys.exit(0)
            else:
                continue
        except IndexError:
            print("Указанная дата недоступна")
            if input("Попробуете снова? Y/N\n").upper() == "N":
                sys.exit(0)
            else:
                continue
    operations_df = (
        pd.read_excel(file_write_path)
        if (os.path.isfile(file_write_path) and (os.path.splitext(file_write_path)[1].lower() == ".xlsx"))
        else get_frame_operations()
    )
    if user_date == "":
        return [operations_df, None]
    else:
        return [operations_df, user_date]


def get_choice_menu() -> str:
    """
    Функция возвращает ввод пользователя в виде строки.
    """

    greeting = """
            Привет! В этом разделе можно управлять статистикой и
            формировать отчёты о транзакциях"""
    print(greeting)

    while True:
        print("""
            Выберите необходимый пункт меню:

            1. Общая статистика по транзакциям за определённый период

            2. Сервисы по категориям

            3. Отчеты по расходам/пополнениям

            0. Выход
            """)

        user_choice = input()
        if not user_choice.isdigit():
            print(f"Статус выбора {user_choice} недоступен.")
            continue
        elif int(user_choice) in [1, 2, 3, 0]:
            break

    return user_choice


def quantize_decimal(x: Any) -> Decimal:
    """
    Функция принимает подготовленное значение с любым типом данных и возвращает число типа Decimal.
    """

    return Decimal(str(x)).quantize(Decimal("0.00"))


if __name__ == "__main__":
    print(get_frame_operations())
