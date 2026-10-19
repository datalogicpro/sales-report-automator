from dataclasses import dataclass
from pathlib import Path
from unittest.mock import patch

import pytest
import streamlit
from streamlit.testing.v1 import AppTest

from app import UPLOAD_CHOICE, make_report
from generate_report import ReportError

APP_FILE = str(Path(__file__).parent.parent / "app.py")
SAMPLE_FILE = Path(__file__).parent.parent / "sample_csv_data.csv"

# The first run imports pandas and builds the report, which can take a while
TIMEOUT = 30


@dataclass
class FakeUpload:
    """Stands in for Streamlit's uploaded file, which AppTest can't create."""

    name: str
    data: bytes

    def getvalue(self) -> bytes:
        return self.data


def run_with_upload(upload: FakeUpload | None) -> AppTest:
    with patch.object(streamlit, "file_uploader", return_value=upload):
        app = AppTest.from_file(APP_FILE, default_timeout=TIMEOUT).run()
        app.radio[0].set_value(UPLOAD_CHOICE).run()
    return app


def test_make_report_on_sample_file() -> None:
    result = make_report(SAMPLE_FILE.name, SAMPLE_FILE.read_bytes())

    assert result.raw_rows == 618
    assert len(result.clean) == 600
    # Excel files are zip files, which start with "PK"
    assert result.report.startswith(b"PK")
    assert list(result.summaries.by_month.columns) == [
        "Revenue",
        "Change vs Prior Month (%)",
    ]


def test_make_report_bad_file() -> None:
    with pytest.raises(ReportError, match="missing these columns"):
        make_report("sales.csv", b"region\nWest\n")


def test_page_shows_sample_report() -> None:
    app = AppTest.from_file(APP_FILE, default_timeout=TIMEOUT).run()

    assert not app.exception
    assert [metric.value for metric in app.metric] == ["618", "600", "18"]


def test_page_waits_for_upload() -> None:
    app = run_with_upload(None)

    assert not app.exception
    assert app.info[0].value.startswith("Upload a .csv")
    assert not app.metric


def test_page_builds_uploaded_file() -> None:
    app = run_with_upload(FakeUpload("sales.csv", SAMPLE_FILE.read_bytes()))

    assert not app.exception
    assert [metric.value for metric in app.metric] == ["618", "600", "18"]


def test_page_shows_file_errors() -> None:
    app = run_with_upload(FakeUpload("sales.csv", b"region\nWest\n"))

    assert not app.exception
    assert "missing these columns" in app.error[0].value
    assert not app.metric
