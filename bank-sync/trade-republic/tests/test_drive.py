import pytest

from fakes import FakeDrive
from tr_export import drive


def test_uploads_only_when_the_content_changed():
    http = FakeDrive(current=b"old")
    assert drive.upload(http, "file-1", b"new") == "updated"
    assert drive.upload(http, "file-1", b"new") == "unchanged"
    assert http.patches == [b"new"]


def test_a_file_not_shared_with_the_service_account_names_the_status_only():
    with pytest.raises(drive.DriveError, match=r"HTTP 404.*shared with the service account"):
        drive.upload(FakeDrive(get_status=404), "file-1", b"x")
    with pytest.raises(drive.DriveError, match=r"HTTP 403"):
        drive.upload(FakeDrive(current=b"old", patch_status=403), "file-1", b"x")
