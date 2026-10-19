"""Web page for the Sales Report Automator.

Run it with:  uv run streamlit run app.py
"""

import tempfile
from io import BytesIO
from pathlib import Path
from typing import NamedTuple

import pandas as pd
import streamlit as st
from pandas.io.formats.style import Styler

from generate_report import (
    REQUIRED_COLUMNS,
    ReportError,
    Summaries,
    build_summaries,
    clean_data,
    format_summaries,
    load_data,
    write_report,
)

SAMPLE_FILE = Path(__file__).parent / "sample_csv_data.csv"
SAMPLE_CHOICE = "Try the sample file"
UPLOAD_CHOICE = "Upload my own file"
UPLOAD_TYPES = ["csv", "tsv", "txt", "xlsx"]
REPORT_NAME = "sales_report.xlsx"
EXCEL_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
CHANGE_COLUMN = "Change vs Prior Month (%)"


class Result(NamedTuple):
    report: bytes
    raw_rows: int
    clean: pd.DataFrame
    summaries: Summaries


def make_report(file_name: str, data: bytes) -> Result:
    """Clean an uploaded file and build the Excel report in memory."""
    suffix = Path(file_name).suffix.lower()
    with tempfile.TemporaryDirectory() as folder:
        path = Path(folder) / f"upload{suffix}"
        path.write_bytes(data)
        raw = load_data(path)

    raw_rows = len(raw)
    clean = clean_data(raw)
    summaries = build_summaries(clean)

    output = BytesIO()
    write_report(clean, summaries, output)
    return Result(output.getvalue(), raw_rows, clean, format_summaries(summaries))


def format_table(table: pd.DataFrame) -> Styler:
    """Show money as dollars, change as a percent, and a dash for no value."""
    formats = {"Revenue": "${:,.2f}"}
    if CHANGE_COLUMN in table.columns:
        formats[CHANGE_COLUMN] = "{:+.1f}%"
    return table.style.format(formats, na_rep="—")


def show_result(result: Result) -> None:
    rows, clean_rows, duplicates = st.columns(3)
    rows.metric("Rows in the file", result.raw_rows)
    clean_rows.metric("Clean rows", len(result.clean))
    duplicates.metric("Duplicates removed", result.raw_rows - len(result.clean))

    st.download_button(
        "Download the Excel report",
        result.report,
        file_name=REPORT_NAME,
        mime=EXCEL_TYPE,
        type="primary",
    )

    monthly, products, regions, clean = st.tabs(
        ["Monthly", "Products", "Regions", "Clean Data"]
    )
    for tab, table in [
        (monthly, result.summaries.by_month),
        (products, result.summaries.by_product),
        (regions, result.summaries.by_region),
    ]:
        with tab:
            st.bar_chart(table["Revenue"])
            st.dataframe(format_table(table), placeholder="—")
    with clean:
        st.dataframe(result.clean, hide_index=True)


def main() -> None:
    st.set_page_config(page_title="Sales Report Automator", layout="wide")
    st.title("Sales Report Automator")
    st.write(
        "Turns a messy sales export into a clean, formatted Excel report "
        "with summaries and charts."
    )

    choice = st.radio("Sales file", [SAMPLE_CHOICE, UPLOAD_CHOICE], horizontal=True)

    if choice == SAMPLE_CHOICE:
        data = SAMPLE_FILE.read_bytes()
        name = SAMPLE_FILE.name
        st.download_button(
            "Download the messy sample file", data, file_name=name, mime="text/csv"
        )
    else:
        columns = ", ".join(f"`{column}`" for column in REQUIRED_COLUMNS)
        st.caption(f"The file needs these columns: {columns}")
        uploaded = st.file_uploader("Upload a sales file", type=UPLOAD_TYPES)
        if uploaded is None:
            st.info("Upload a .csv, .tsv, .txt, or .xlsx file to build the report.")
            return
        data = uploaded.getvalue()
        name = uploaded.name

    try:
        result = make_report(name, data)
    except ReportError as exc:
        st.error(str(exc))
        return

    show_result(result)


if __name__ == "__main__":
    main()
