"""Generate a formatted Excel sales report from a raw sales file.

Usage:  python generate_report.py sample_data/sales_raw.csv

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
import sys
from pathlib import Path

import pandas as pd


def detect_separator(path):
    """Work out which character separates the columns.

    Only the header line is checked. Data rows can contain stray commas
    (like "$2,790.00"), which would fool a guess based on the whole file.
    """
    with open(path, encoding="utf-8-sig") as f:
        header = f.readline()
    candidates = [",", "\t", ";", "|"]
    return max(candidates, key=header.count)


def load_data(path):
    """Read a raw sales file into a DataFrame.

    Everything is read as text on purpose. The cleaning step decides how to
    turn each column into dates and numbers, so values like "$2,790.00"
    are not silently mangled while loading.
    """
    path = Path(path)
    if not path.exists():
        sys.exit(f"File not found: {path}")

    if path.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(path, dtype=str)

    return pd.read_csv(path, sep=detect_separator(path), dtype=str)


REGIONS = {"nor": "North", "sou": "South", "eas": "East", "wes": "West"}


def clean_regions(df):
    """Make every region spelling match: WEST, " west ", Wes -> West."""
    first_three = df["region"].str.strip().str.lower().str[:3]
    df["region"] = first_three.map(REGIONS)
    return df


MONEY_COLUMNS = ["unit_price", "revenue"]


def clean_money(df):
    """Turn "$2,790.00" and "315.00" into real numbers. Blanks stay blank."""
    for column in MONEY_COLUMNS:
        text = df[column].str.strip().str.replace("$", "", regex=False)
        text = text.str.replace(",", "", regex=False)
        df[column] = pd.to_numeric(text, errors="coerce")
    return df


DATE_FORMATS = ["%Y-%m-%d", "%m/%d/%Y", "%d-%b-%Y", "%B %d, %Y"]


def clean_dates(df):
    """Turn all four date styles into real dates that Python can sort and group."""
    text = df["order_date"].str.strip()
    dates = pd.Series(pd.NaT, index=df.index)
    for date_format in DATE_FORMATS:
        dates = dates.fillna(pd.to_datetime(text, format=date_format, errors="coerce"))
    df["order_date"] = dates
    return df


def clean_products(df):
    """Make every product name match: "laptop stand", "Laptop  Stand" -> "Laptop Stand"."""
    text = df["product"].str.strip().str.replace(r"\s+", " ", regex=True)
    key = text.str.lower()
    # Each product's most common spelling becomes its official name.
    official = text.groupby(key).agg(lambda names: names.mode().iloc[0])
    df["product"] = key.map(official)
    return df


def remove_duplicates(df):
    """Drop rows that are exact copies of another row."""
    before = len(df)
    df = df.drop_duplicates().reset_index(drop=True)
    print(f"Removed {before - len(df)} duplicate rows. {len(df)} rows left.")
    return df


def fill_blanks(df):
    """Fill blanks that can be worked out from other columns.

    revenue = quantity x unit_price, so if only one of those three is
    blank, it can be calculated from the other two.
    """
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce")

    no_revenue = df["revenue"].isna() & df["quantity"].notna() & df["unit_price"].notna()
    df.loc[no_revenue, "revenue"] = df["quantity"] * df["unit_price"]

    no_price = df["unit_price"].isna() & df["quantity"].notna() & df["revenue"].notna()
    df.loc[no_price, "unit_price"] = df["revenue"] / df["quantity"]

    no_quantity = df["quantity"].isna() & df["unit_price"].notna() & df["revenue"].notna()
    df.loc[no_quantity, "quantity"] = (df["revenue"] / df["unit_price"]).round()

    df["sales_rep"] = df["sales_rep"].fillna("Unassigned")
    return df


def build_summaries(df):
    """Make the summary tables that will go into the report."""
    by_month = df.groupby(df["order_date"].dt.to_period("M"))["revenue"].sum().to_frame()
    by_month["change_pct"] = (by_month["revenue"].pct_change() * 100).round(1)

    by_product = df.groupby("product")["revenue"].sum().sort_values(ascending=False).to_frame()
    by_region = df.groupby("region")["revenue"].sum().sort_values(ascending=False).to_frame()

    return by_month.round(2), by_product.round(2), by_region.round(2)


OUTPUT_FILE = "output/sales_report.xlsx"


def write_report(df, by_month, by_product, by_region, output_path=OUTPUT_FILE):
    """Save a formatted Excel report: one sheet per summary, with charts."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    by_month = by_month.copy()
    by_month.index = by_month.index.astype(str)
    by_month.index.name = "Month"
    by_month.columns = ["Revenue", "Change vs Prior Month (%)"]
    by_product.index.name = "Product"
    by_product.columns = ["Revenue"]
    by_region.index.name = "Region"
    by_region.columns = ["Revenue"]

    with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        workbook = writer.book
        header = workbook.add_format(
            {"bold": True, "bg_color": "#1F4E78", "font_color": "white", "border": 1}
        )
        money = workbook.add_format({"num_format": "$#,##0.00"})
        percent = workbook.add_format({"num_format": "0.0"})

        sheets = [
            ("Monthly", by_month, "column"),
            ("Products", by_product, "bar"),
            ("Regions", by_region, "column"),
        ]
        for name, table, chart_type in sheets:
            table.to_excel(writer, sheet_name=name)
            sheet = writer.sheets[name]

            for col, title in enumerate([table.index.name] + list(table.columns)):
                sheet.write(0, col, title, header)
            sheet.set_column(0, 0, 28)
            sheet.set_column(1, 1, 16, money)
            if name == "Monthly":
                sheet.set_column(2, 2, 26, percent)

            chart = workbook.add_chart({"type": chart_type})
            last_row = len(table)
            chart.add_series(
                {
                    "name": f"Revenue by {table.index.name.lower()}",
                    "categories": [name, 1, 0, last_row, 0],
                    "values": [name, 1, 1, last_row, 1],
                    "fill": {"color": "#1F4E78"},
                }
            )
            chart.set_title({"name": f"Revenue by {table.index.name.lower()}"})
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

    return output_path


def main():
    parser = argparse.ArgumentParser(description="Build a sales report from a raw file.")
    parser.add_argument("input_file", help="path to the raw sales file (.csv, .tsv, .txt, .xlsx)")
    args = parser.parse_args()

    df = load_data(args.input_file)
    print(f"Loaded {len(df)} rows and {len(df.columns)} columns")

    df = clean_regions(df)
    df = clean_money(df)
    df = clean_dates(df)                                                                  
    df = clean_products(df)
    df = remove_duplicates(df)
    df = fill_blanks(df)

    by_month, by_product, by_region = build_summaries(df)
    saved_to = write_report(df, by_month, by_product, by_region)
    print(f"Report saved to {saved_to}")


if __name__ == "__main__":
    main()