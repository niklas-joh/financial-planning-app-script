"""Trade Republic web login through pytr (the app.traderepublic.com flow, v2: you confirm in the app).

The scheduled run only resumes the saved web session (`session/cookies.txt`); it never asks for anything.
When the session has expired it raises `LoginNeeded`, and you run `login.sh` once: it asks for your PIN
(never stored), sends a confirmation to the Trade Republic app and saves the new session. Your phone stays
logged in (no device reset).
"""

from __future__ import annotations

import getpass
from typing import Callable

from pytr.api import TradeRepublicApi

from .config import Settings

# The scheduled run never logs in, so it has no PIN; pytr needs some value to build the client.
_NO_PIN = "0000"


class LoginNeeded(RuntimeError):
    """The saved web session is missing or expired: run login.sh."""


def _api(settings: Settings, pin: str, factory: Callable[..., TradeRepublicApi]) -> TradeRepublicApi:
    settings.session_dir.mkdir(parents=True, exist_ok=True)
    return factory(
        phone_no=settings.phone,
        pin=pin,
        save_cookies=True,
        cookies_file=str(settings.cookies_file),
        use_v2_login=True,
        waf_token=None,
    )


def resume(settings: Settings, factory: Callable[..., TradeRepublicApi] = TradeRepublicApi) -> TradeRepublicApi:
    """The saved session, or LoginNeeded."""
    tr = _api(settings, _NO_PIN, factory)
    if not tr.resume_websession():
        raise LoginNeeded("Trade Republic session missing or expired")
    return tr


def login(
    settings: Settings,
    factory: Callable[..., TradeRepublicApi] = TradeRepublicApi,
    ask_pin: Callable[[], str] = lambda: getpass.getpass("Trade Republic PIN (hidden, not stored): "),
    ask_code: Callable[[str], str] = input,
    say: Callable[[str], None] = print,
) -> TradeRepublicApi:
    """Resumes the saved session, else logs in interactively and saves the new one."""
    try:
        return resume(settings, factory)
    except LoginNeeded:
        pass
    tr = _api(settings, ask_pin().strip(), factory)
    tr.initiate_weblogin()
    if tr.weblogin_needs_authenticator:
        tr.complete_weblogin(ask_code("Code from your authenticator app: ").strip())
    else:
        say("Confirm the login in your Trade Republic app (waiting up to a few minutes)...")
        tr.complete_weblogin(None)
    if settings.cookies_file.exists():
        settings.cookies_file.chmod(0o600)
    say("Logged in; the session is saved for the scheduled runs.")
    return tr
