import pytest

from tr_export import config


def test_defaults_and_paths():
    s = config.load({"TR_PHONE": "+352000000000", "TR_HOME": "/x"})
    assert (s.account, s.output_name, s.lookback_days, s.drive_file_id) == ("Trade Republic", "trade-republic.csv", 45, "")
    assert str(s.cookies_file) == "/x/session/cookies.txt"
    assert str(s.output_file) == "/x/data/trade-republic.csv"
    assert s.tz.key == "Europe/Brussels"


@pytest.mark.parametrize("env, message", [
    ({}, "TR_PHONE"),
    ({"TR_PHONE": "352000"}, "TR_PHONE"),
    ({"TR_PHONE": "+352000", "TR_LOOKBACK_DAYS": "x"}, "TR_LOOKBACK_DAYS"),
    ({"TR_PHONE": "+352000", "TR_LOOKBACK_DAYS": "0"}, "TR_LOOKBACK_DAYS"),
])
def test_invalid_settings_name_the_setting(env, message):
    with pytest.raises(ValueError, match=message):
        config.load(env)
