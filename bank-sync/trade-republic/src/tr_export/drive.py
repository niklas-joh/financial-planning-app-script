"""Replaces the content of one existing CSV file in the Bank Sync CSV folder on Google Drive.

The file is created once by you (upload the first `trade-republic.csv` to the folder) and shared with the
service account as Editor. The job only ever updates that file's content, so it stays yours (a service
account has no Drive storage of its own) and Bank Sync reads it like any other export. Nothing is uploaded
when the content has not changed.
"""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any

SCOPES = ("https://www.googleapis.com/auth/drive",)
_FILES = "https://www.googleapis.com/drive/v3/files/"
_UPLOAD = "https://www.googleapis.com/upload/drive/v3/files/"


class DriveError(RuntimeError):
    """The upload failed; the message names the HTTP status only."""


def session(credentials: Path) -> Any:
    """An authorised requests session for the service account."""
    from google.auth.transport.requests import AuthorizedSession
    from google.oauth2 import service_account

    creds = service_account.Credentials.from_service_account_file(str(credentials), scopes=list(SCOPES))
    return AuthorizedSession(creds)


def upload(http: Any, file_id: str, content: bytes) -> str:
    """'updated' or 'unchanged'. Raises DriveError (e.g. 404 when the file is not shared with the service account)."""
    meta = http.get(_FILES + file_id, params={"fields": "md5Checksum", "supportsAllDrives": "true"})
    if meta.status_code != 200:
        raise DriveError(f"reading the Drive file failed (HTTP {meta.status_code}); is it shared with the service account?")
    if meta.json().get("md5Checksum") == hashlib.md5(content).hexdigest():  # noqa: S324 - Drive's checksum, not security
        return "unchanged"
    r = http.patch(
        _UPLOAD + file_id,
        params={"uploadType": "media", "supportsAllDrives": "true"},
        data=content,
        headers={"Content-Type": "text/csv"},
    )
    if r.status_code != 200:
        raise DriveError(f"updating the Drive file failed (HTTP {r.status_code})")
    return "updated"
