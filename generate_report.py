"""Generate a formatted Excel sales report from a raw sales file.

Usage:  python generate_report.py sample_csv_data.csv

Part 1: loading the file (handles CSV, TSV, semicolon, pipe, and Excel).
Part 2: cleaning the regions.
Part 3: cleaning the money columns.
Part 4: cleaning the dates.
Part 5: cleaning the product names.
Part 6: removing duplicate rows.
Part 7: filling in blanks.
Part 8: building the summary tables.
Part 9: writing the Excel report.
"""

import argparse
import logging
import sys
from pathlib import Path
from typing import IO, NamedTuple

import pandas as pd

logger = logging.getLogger(__name__)

REQUIRED_COLUMNS = [
    "order_date",
    "region",
    "product",
    "quantity",
    "unit_price",
    "revenue",
    "sales_rep",
]

TEXT_SUFFIXES = (".csv", ".tsv", ".txt")
EXCEL_SUFFIXES = (".xlsx",)


class ReportError(Exception):
    """A problem with the input file that the person running the report can fix."""


class Summaries(NamedTuple):
    by_month: pd.DataFrame
    by_product: pd.DataFrame
    by_region: pd.DataFrame


def detect_separator(path: Path) -> str:
    """Work out which character separates the columns.

    Only the header line is checked. Data rows can contain stray commas
    (like "$2,790.00"), which would fool a guess based on the whole file.
    """
    with open(path, encoding="utf-8-sig") as f:
        header = f.readline()
    candidates = [",", "\t", ";", "|"]
    return max(candidates, key=header.count)


def check_columns(df: pd.DataFrame) -> None:
    """Stop with a clear message if any required column is missing."""
    missing = [column for column in REQUIRED_COLUMNS if column not in df.columns]
    if missing:
        raise ReportError(f"The file is missing these columns: {', '.join(missing)}")


def load_data(path: str | Path) -> pd.DataFrame:
    """Read a raw sales file into a DataFrame.

    Everything is read as text on purpose. The cleaning step decides how to
    turn each column into dates and numbers, so values like "$2,790.00"
    are not silently mangled while loading.
    """
    path = Path(path)
    if not path.exists():
        raise ReportError(f"File not found: {path}")

    suffix = path.suffix.lower()
    if suffix not in TEXT_SUFFIXES + EXCEL_SUFFIXES:
        raise ReportError(
            f"Unsupported file type {suffix!r}. Use .csv, .tsv, .txt, or .xlsx."
        )

    try:
        if suffix in EXCEL_SUFFIXES:
            df = pd.read_excel(path, dtype=str)
        else:
            df = pd.read_csv(path, sep=detect_separator(path), dtype=str)
    except ValueError as exc:
        raise ReportError(f"Could not read the file: {exc}") from exc

    # Column names like " Region " or "REGION" still count as "region"
    df.columns = df.columns.str.strip().str.lower()
    check_columns(df)

    if df.empty:
        raise ReportError("The file has no data rows.")

    return df


REGIONS = {"nor": "North", "sou": "South", "eas": "East", "wes": "West"}


def clean_regions(df: pd.DataFrame) -> pd.DataFrame:
    """Make every region spelling match: WEST, " west ", Wes -> West."""
    first_three = df["region"].str.strip().str.lower().str[:3]
    df["region"] = first_three.map(REGIONS)
    return df


MONEY_COLUMNS = ["unit_price", "revenue"]


def clean_money(df: pd.DataFrame) -> pd.DataFrame:
    """Turn "$2,790.00" and "315.00" into real numbers. Blanks stay blank."""
    for column in MONEY_COLUMNS:
        text = df[column].str.strip().str.replace("$", "", regex=False)
        text = text.str.replace(",", "", regex=False)
        df[column] = pd.to_numeric(text, errors="coerce")
    return df


DATE_FORMATS = ["%Y-%m-%d", "%m/%d/%Y", "%d-%b-%Y", "%B %d, %Y"]


def clean_dates(df: pd.DataFrame) -> pd.DataFrame:
    """Turn all four date styles into real dates that Python can sort and group."""
    text = df["order_date"].str.strip()
    dates = pd.Series(pd.NaT, index=df.index)
    for date_format in DATE_FORMATS:
        dates = dates.fillna(pd.to_datetime(text, format=date_format, errors="coerce"))
    df["order_date"] = dates
    return df


def clean_products(df: pd.DataFrame) -> pd.DataFrame:
    """Make product names match: "laptop stand", "Laptop  Stand" -> "Laptop Stand"."""
    text = df["product"].str.strip().str.replace(r"\s+", " ", regex=True)
    key = text.str.lower()
    # Each product's most common spelling becomes its official name.
    official = text.groupby(key).agg(lambda names: names.mode().iloc[0])
    df["product"] = key.map(official)
    return df


def remove_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """Drop rows that are exact copies of another row."""
    before = len(df)
    df = df.drop_duplicates().reset_index(drop=True)
    logger.info("Removed %d duplicate rows. %d rows left.", before - len(df), len(df))
    return df


def fill_blanks(df: pd.DataFrame) -> pd.DataFrame:
    """Fill blanks that can be worked out from other columns.

    revenue = quantity x unit_price, so if only one of those three is
    blank, it can be calculated from the other two.
    """
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce")

    no_revenue = (
        df["revenue"].isna() & df["quantity"].notna() & df["unit_price"].notna()
    )
    df.loc[no_revenue, "revenue"] = df["quantity"] * df["unit_price"]

    no_price = df["unit_price"].isna() & df["quantity"].notna() & df["revenue"].notna()
    df.loc[no_price, "unit_price"] = df["revenue"] / df["quantity"]

    no_quantity = (
        df["quantity"].isna() & df["unit_price"].notna() & df["revenue"].notna()
    )
    df.loc[no_quantity, "quantity"] = (df["revenue"] / df["unit_price"]).round()

    df["sales_rep"] = df["sales_rep"].fillna("Unassigned")
    return df


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Run every cleaning step, in order."""
    df = clean_regions(df)
    df = clean_money(df)
    df = clean_dates(df)
    df = clean_products(df)
    df = remove_duplicates(df)
    return fill_blanks(df)


def find_missing(df: pd.DataFrame) -> pd.DataFrame:
    """Rows still missing a required value after cleaning, and which values.

    These can't be worked out from the other columns, so a person needs to
    look them up in the original records.
    """
    blanks = df[REQUIRED_COLUMNS].isna()
    needs_review = blanks.any(axis=1)
    missing = blanks.apply(lambda row: ", ".join(row.index[row]), axis=1)
    review = df[needs_review].assign(missing=missing[needs_review])
    return review.reset_index(drop=True)


def build_summaries(df: pd.DataFrame) -> Summaries:
    """Make the summary tables that will go into the report."""
    months = df["order_date"].dt.to_period("M")
    by_month = df.groupby(months)["revenue"].sum().to_frame()
    by_month["change_pct"] = (by_month["revenue"].pct_change() * 100).round(1)

    by_product = df.groupby("product")["revenue"].sum().sort_values(ascending=False)
    by_region = df.groupby("region")["revenue"].sum().sort_values(ascending=False)

    return Summaries(
        by_month.round(2),
        by_product.to_frame().round(2),
        by_region.to_frame().round(2),
    )


def format_summaries(summaries: Summaries) -> Summaries:
    """Give the summary tables the titles shown in the report."""
    by_month = summaries.by_month.copy()
    by_month.index = by_month.index.astype(str)
    by_month.index.name = "Month"
    by_month.columns = ["Revenue", "Change vs Prior Month (%)"]

    by_product = summaries.by_product.copy()
    by_product.index.name = "Product"
    by_product.columns = ["Revenue"]

    by_region = summaries.by_region.copy()
    by_region.index.name = "Region"
    by_region.columns = ["Revenue"]

    return Summaries(by_month, by_product, by_region)


OUTPUT_FILE = Path("output/sales_report.xlsx")


def write_report(
    df: pd.DataFrame, summaries: Summaries, output: Path | IO[bytes]
) -> None:
    """Save a formatted Excel report: one sheet per summary, with charts.

    The output can be a file path or an open binary file, such as the
    in-memory file the web page sends back for download.
    """
    if isinstance(output, Path):
        output.parent.mkdir(parents=True, exist_ok=True)

    tables = format_summaries(summaries)

    with pd.ExcelWriter(
        output, engine="xlsxwriter", datetime_format="yyyy-mm-dd"
    ) as writer:
        workbook = writer.book
        header = workbook.add_format(
            {"bold": True, "bg_color": "#1F4E78", "font_color": "white", "border": 1}
        )
        money = workbook.add_format({"num_format": "$#,##0.00"})
        percent = workbook.add_format({"num_format": "0.0"})

        sheets = [
            ("Monthly", tables.by_month, "column"),
            ("Products", tables.by_product, "bar"),
            ("Regions", tables.by_region, "column"),
        ]
        for name, table, chart_type in sheets:
            table.to_excel(writer, sheet_name=name)
            sheet = writer.sheets[name]

            for col, title in enumerate([table.index.name, *table.columns]):
                sheet.write(0, col, title, header)
            sheet.set_column(0, 0, 28)
            sheet.set_column(1, 1, 16, money)
            if name == "Monthly":
                sheet.set_column(2, 2, 26, percent)

            chart = workbook.add_chart({"type": chart_type})
            last_row = len(table)
            label = f"Revenue by {str(table.index.name).lower()}"
            chart.add_series(
                {
                    "name": label,
                    "categories": [name, 1, 0, last_row, 0],
                    "values": [name, 1, 1, last_row, 1],
                    "fill": {"color": "#1F4E78"},
                }
            )
            chart.set_title({"name": label})
            chart.set_legend({"none": True})
            if chart_type == "bar":
                chart.set_y_axis({"reverse": True})  # biggest at the top
            chart.set_size({"width": 620, "height": 340})
            sheet.insert_chart("E2", chart)

        df.to_excel(writer, sheet_name="Clean Data", index=False)
        data_sheet = writer.sheets["Clean Data"]
        for col, title in enumerate(df.columns):
            data_sheet.write(0, col, title, header)
        data_sheet.set_column(0, len(df.columns) - 1, 18)
        data_sheet.freeze_panes(1, 0)
        data_sheet.autofilter(0, 0, len(df), len(df.columns) - 1)


def build_report(input_path: str | Path, output: Path | IO[bytes]) -> pd.DataFrame:
    """Load, clean, summarize, and save. Returns the clean data."""
    df = load_data(input_path)
    logger.info("Loaded %d rows and %d columns", len(df), len(df.columns))

    df = clean_data(df)
    logger.info("%d rows need review", len(find_missing(df)))
    write_report(df, build_summaries(df), output)
    return df


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Build a sales report from a raw file."
    )
    parser.add_argument(
        "input_file", help="path to the raw sales file (.csv, .tsv, .txt, .xlsx)"
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=OUTPUT_FILE,
        help=f"where to save the report (default: {OUTPUT_FILE})",
    )
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(message)s")

    try:
        build_report(args.input_file, args.output)
    except ReportError as exc:
        logger.error("%s", exc)
        return 1

    logger.info("Report saved to %s", args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
