from src.services import aggregate_private_transfer


def test_agg_private_trans(get_private_trans_json) -> None:
    result = aggregate_private_transfer()
    assert result == get_private_trans_json
