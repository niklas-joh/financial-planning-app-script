"""Fakes for pytr's client and Drive's HTTP session (no network)."""

from __future__ import annotations

import hashlib


class FakeApi:
    instances: list["FakeApi"] = []

    def __init__(self, resumes=True, needs_authenticator=False, **kwargs):
        self.kwargs = kwargs
        self._resumes = resumes
        self.weblogin_needs_authenticator = needs_authenticator
        self.calls: list = []
        FakeApi.instances.append(self)

    def resume_websession(self):
        self.calls.append("resume")
        return self._resumes

    def initiate_weblogin(self):
        self.calls.append("initiate")
        return 120

    def complete_weblogin(self, code=None):
        self.calls.append(("complete", code))


def factory(resumes=True, needs_authenticator=False):
    def make(**kwargs):
        return FakeApi(resumes=resumes, needs_authenticator=needs_authenticator, **kwargs)
    return make


class Response:
    def __init__(self, status_code, body=None):
        self.status_code = status_code
        self._body = body or {}

    def json(self):
        return self._body


class FakeDrive:
    def __init__(self, current: bytes | None = b"", get_status=200, patch_status=200):
        self.current = current
        self.get_status = get_status
        self.patch_status = patch_status
        self.patches: list[bytes] = []

    def get(self, url, params=None):
        md5 = hashlib.md5(self.current).hexdigest() if self.current is not None else None
        return Response(self.get_status, {"md5Checksum": md5})

    def patch(self, url, params=None, data=None, headers=None):
        assert params["uploadType"] == "media" and headers["Content-Type"] == "text/csv"
        self.patches.append(data)
        if self.patch_status == 200:
            self.current = data
        return Response(self.patch_status)
