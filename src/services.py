import json
import logging
from decimal import ROUND_HALF_EVEN, Decimal

import pandas as pd

from src.utils import base_dir, data_path, quantize_decimal

path_to_log = base_dir + "/logs/module_services.log"
services_logger = logging.getLogger(__name__)
if not services_logger.handlers:
    file_handler = logging.FileHandler(path_to_log, mode="w")
    file_formatter = logging.Formatter("%(asctime)s - %(module)s - %(levelname)s - %(message)s")
    file_handler.setFormatter(file_formatter)
    services_logger.addHandler(file_handler)
    services_logger.setLevel(logging.DEBUG)


def aggregate_private_transfer() -> str:
    """
    Функция возвращает строку со всеми транзакциями, которые относятся к переводам физлицам.
    """

    services_logger.info("Запуск <aggregate_private_transfer>")
    transfer_df = pd.read_excel(data_path)
    transfer_df.rename(columns=lambda x: x.replace(" ", "_") if isinstance(x, str) else x, inplace=True)
    transactions_df = transfer_df[transfer_df["Статус"] == "OK"]

    transactions_df["Сумма_операции"] = transactions_df["Сумма_операции"].astype(str).transform(quantize_decimal)

    mask_descript = transactions_df["Описание"].str.contains(r"^[А-Я]{1}[а-я]{3,}\s[А-Я].$", regex=True, na=False)
    mask_category = transactions_df["Категория"] == "Переводы"

    services_logger.info("Создание датафрейма по переводам физлицам")
    private_df = transactions_df.loc[(mask_descript & mask_category), ["Дата_операции", "Описание", "Сумма_операции"]]
    private_df["Дата_операции"] = private_df["Дата_операции"].astype(str)

    private_index_df = private_df.set_index(keys=["Дата_операции", "Описание"])
    private_abs_df = private_index_df.apply(pd.Series.abs)
    total_private_amount = int(
        Decimal(float(private_abs_df.apply(pd.Series.sum).values[0])).quantize(Decimal("1."), ROUND_HALF_EVEN)
    )
    private_abs_df = private_abs_df.astype(float)

    private_all_dict = private_abs_df.to_dict()["Сумма_операции"]
    private_collection = list(private_all_dict.items())
    private_list = list(map(lambda x: {"datetime": x[0][0], "payee": x[0][1], "amount": x[1]}, private_collection))

    services_logger.info("Формирование статистики по переводам физлицам")
    private_orderly_dict = {"outgoing_transfers": {"total_amount": total_private_amount, "main": private_list}}
    private_transfer_json = json.dumps(obj=private_orderly_dict, ensure_ascii=False, indent=4)

    services_logger.info("Завершение <aggregate_private_transfer>")
    return private_transfer_json
