# Sales Report Automator

Turns a messy sales export into a clean, formatted Excel report in seconds.

## The problem

Many small businesses rebuild the same sales report by hand every week. The raw exports are messy: dates in different formats, misspelled regions, prices stored as text, duplicate rows, and blank cells. Cleaning that by hand takes hours and is easy to get wrong.

## What it does

- **Reads many file types:** CSV, tab-separated, semicolon-separated, pipe-separated, and Excel (.xlsx). It detects the separator automatically.
- **Cleans the data:**
  - Makes region names consistent (`WEST`, ` west `, `Wes` all become `West`)
  - Turns text like `$2,790.00` into real numbers
  - Converts four different date styles into one real date format
  - Makes product names consistent
  - Removes duplicate rows
  - Fills in blanks that can be calculated (revenue = quantity x price) and labels missing sales reps as "Unassigned"
- **Builds summaries:** revenue by month (with change from the previous month), by product, and by region.
- **Saves a formatted Excel report** with a chart on each summary sheet.

## Results on the sample file

| | Before | After |
|---|---|---|
| Rows | 618 | 600 (18 duplicates removed) |
| Region spellings | 20 | 4 |
| Product spellings | 24 | 8 |
| Date formats | 4 | 1 |
| Blank quantities | 13 | 1 |
| Blank prices | 51 | 4 |
| Blank revenue amounts | 58 | 5 |
| Blank sales reps | 23 | 0 |

The few blanks left are rows missing two of the three numbers (quantity, price, revenue), so they can't be calculated.

## How to run

1. Install the libraries:

   ```
   pip install -r requirements.txt
   ```

2. Run the program on a raw sales file:

   ```
   python generate_report.py sample_csv_data.csv
   ```

3. Open the finished report at `output/sales_report.xlsx`.

## The report

The Excel file has four sheets:

- **Monthly:** revenue by month, percent change from the previous month, and a chart
- **Products:** revenue by product, biggest first, with a chart
- **Regions:** revenue by region, with a chart
- **Clean Data:** all cleaned rows, with filters

## Built with

Python, pandas, xlsxwriter, and openpyxl.