import pandas as pd
from pandas import DataFrame
from pytest import CaptureFixture
from src.reports import writing_to_file, spending_by_workday, report_dir
from src.utils import file_write_path

from tests.conftest import get_trial_df, get_early_date


def test_write_report(capsys: CaptureFixture[str]) -> None:
    @writing_to_file
    def get_small_df() -> DataFrame:
        sector_df = pd.read_excel(file_write_path)
        return sector_df[["Дата_операции", "Сумма_операции", "Описание"]][:25]

    result = get_small_df()
    captured = capsys.readouterr()
    assert captured.out == f"Записываю отчёт в output_data/ ...\nОтчёт создан\n"
    assert type(result) == pd.DataFrame

    report_file_df = pd.read_excel(f"{report_dir}/output_data/get_small_df.xlsx")
    assert report_file_df.shape == (25, 3)


def test_spending_from_last(get_trial_df: DataFrame, get_early_date: str) -> None:
    result = spending_by_workday(get_trial_df, get_early_date)
    assert result.shape == (84, 3)
