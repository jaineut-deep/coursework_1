import os
import sys

from pandas import DataFrame

from src.reports import spending_by_workday
from src.services import aggregate_private_transfer
from src.utils import (
    file_write_path,
    filter_choice_spending_workday,
    filter_user_choice,
    get_choice_menu,
    get_frame_operations,
)
from src.views import aggregate_json


def main() -> str | None | DataFrame:
    """Функция предоставляет пользовательский интерфейс и возвращает данные по операциям и формирует отсчёты в
    зависимости от ввода пользовательских данных и выбранной функциональности.
    """

    if os.path.isfile(file_write_path) and (os.path.splitext(file_write_path)[1].lower() == ".xlsx"):
        pass
    else:
        transactions_df = get_frame_operations()
        transactions_df.to_excel(excel_writer=file_write_path, index=False)

    user_num = get_choice_menu()

    if int(user_num) == 1:
        print("Выбрано <<Общая статистика по транзакциям>>")
        data_for_statistics = filter_user_choice()
        return aggregate_json(data_for_statistics[0], data_for_statistics[1])
    elif int(user_num) == 2:
        print("Выбрано <<Сервисы по категориям>>")
        return aggregate_private_transfer()
    elif int(user_num) == 3:
        print("Выбрано <<Отчеты по расходам/пополнениям>>")
        data_for_report = filter_choice_spending_workday()
        return spending_by_workday(data_for_report[0], data_for_report[1])  # type: ignore
    else:
        data_for_bye = ["Выбрано <<Выход>>. Завершение...", ""]
        print(data_for_bye[0])
        sys.exit(0)


if __name__ == "__main__":
    print(main())
