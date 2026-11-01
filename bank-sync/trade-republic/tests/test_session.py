import pytest

from fakes import factory
from tr_export import config, session


@pytest.fixture
def settings(tmp_path):
    return config.load({"TR_PHONE": "+352000000000", "TR_HOME": str(tmp_path)})


def test_resume_uses_the_v2_web_login_and_the_saved_cookies(settings):
    tr = session.resume(settings, factory(resumes=True))
    assert tr.kwargs["use_v2_login"] is True and tr.kwargs["save_cookies"] is True
    assert tr.kwargs["cookies_file"] == str(settings.cookies_file)
    assert tr.kwargs["waf_token"] is None
    assert tr.calls == ["resume"]


def test_scheduled_run_never_logs_in(settings):
    with pytest.raises(session.LoginNeeded):
        session.resume(settings, factory(resumes=False))


def test_login_confirms_in_the_app_and_asks_only_for_the_pin(settings):
    said = []
    tr = session.login(settings, factory(resumes=False), ask_pin=lambda: " 1234 ",
                       ask_code=lambda _: pytest.fail("no code needed"), say=said.append)
    assert tr.kwargs["pin"] == "1234"
    assert tr.calls == ["initiate", ("complete", None)]
    assert any("Confirm the login" in s for s in said)


def test_login_with_an_authenticator_app_asks_for_its_code(settings):
    tr = session.login(settings, factory(resumes=False, needs_authenticator=True), ask_pin=lambda: "1234",
                       ask_code=lambda _: "123456", say=lambda _: None)
    assert tr.calls == ["initiate", ("complete", "123456")]


def test_login_keeps_a_session_that_still_works(settings):
    tr = session.login(settings, factory(resumes=True), ask_pin=lambda: pytest.fail("no PIN needed"))
    assert tr.calls == ["resume"]
