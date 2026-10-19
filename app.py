"""Web page for the Sales Report Automator.

Run it with:  uv run streamlit run app.py
"""

import tempfile
from io import BytesIO
from pathlib import Path
from typing import Any, NamedTuple

import altair as alt
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
# The dark blue used in the Excel report and the page theme
REPORT_COLOR = "#1F4E78"

# How rows of data look on the page: readable names, dates without a time,
# and money as dollars
DATA_COLUMNS = {
    "order_id": st.column_config.TextColumn("Order ID"),
    "order_date": st.column_config.DatetimeColumn("Order Date", format="YYYY-MM-DD"),
    "region": st.column_config.TextColumn("Region"),
    "product": st.column_config.TextColumn("Product"),
    "quantity": st.column_config.NumberColumn("Quantity"),
    "unit_price": st.column_config.NumberColumn("Unit Price", format="dollar"),
    "revenue": st.column_config.NumberColumn("Revenue", format="dollar"),
    "sales_rep": st.column_config.TextColumn("Sales Rep"),
}


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


def product_chart(by_product: pd.DataFrame) -> alt.Chart:
    """Sideways bars, biggest first, with every product's full name."""
    # Each bar gets its own 40-pixel row, and no label is shortened or skipped
    product_axis = alt.Axis(labelLimit=400, labelOverlap=False)
    # Altair's type hints don't say what .properties() returns, so name it here
    chart: alt.Chart = (
        alt.Chart(by_product.reset_index())
        .mark_bar(color=REPORT_COLOR)
        .encode(
            x=alt.X("Revenue:Q", title="Revenue"),
            y=alt.Y("Product:N", sort="-x", title=None, axis=product_axis),
        )
        .properties(height=alt.Step(40))
    )
    return chart


def column_settings(table: pd.DataFrame) -> dict[str, Any]:
    """The column settings for the columns this table has."""
    return {name: setting for name, setting in DATA_COLUMNS.items() if name in table}


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
    summaries = result.summaries
    with monthly:
        # Bars in table order, so months stay in date order
        st.bar_chart(summaries.by_month["Revenue"], color="primary", sort=False)
        st.dataframe(format_table(summaries.by_month), placeholder="—")
    with products:
        st.altair_chart(product_chart(summaries.by_product))
        st.dataframe(format_table(summaries.by_product), placeholder="—")
    with regions:
        # Bars in table order, biggest first
        st.bar_chart(summaries.by_region["Revenue"], color="primary", sort=False)
        st.dataframe(format_table(summaries.by_region), placeholder="—")
    with clean:
        st.dataframe(
            result.clean,
            hide_index=True,
            placeholder="—",
            column_config=column_settings(result.clean),
        )


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
