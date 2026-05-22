from src.views import aggregate_json


def test_agg_json(get_data_string_july: str, get_week_range: str, get_output_agg_json: str) -> None:
    result = aggregate_json(get_data_string_july, get_week_range)
    assert result == get_output_agg_json
