# Changelog

All notable changes to this project are listed here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project uses [Semantic Versioning](https://semver.org/).

## [1.1.0] - 2026-10-19

### Added

- A Needs Review sheet in the Excel report. It lists the rows still missing a value that can't be calculated, with a column naming what's missing.
- A Needs Review tab and a "Rows needing review" count on the web page.

### Changed

- The summary tables on the web page show dollar amounts and percents, and a dash where there is no earlier month to compare with.
- The charts on the web page use the report's dark blue. Months stay in date order, and regions and products are biggest first.
- The product chart on the web page is sideways and shows every product's full name.
- The Clean Data tab on the web page shows readable column names, dates without a time, and dollar amounts.
- The Clean Data sheet in the Excel report shows prices and revenue as dollars, and the product column is wide enough for long names.
- The file checks that run before each commit run through uv, like Ruff and mypy.

## [1.0.0] - 2026-10-19

### Added

- A web page: try the sample file or upload your own, see the summaries and charts, and download the Excel report.
- A check for the 7 required columns, with a message naming any that are missing. Column names still match with different capitals or extra spaces.
- Clear error messages for a file that is missing, unreadable, empty, or the wrong type.
- An `--output` option to choose where the report is saved.
- Logging for each step, in place of printed messages.
- Tests with 100% coverage, Ruff, mypy, pip-audit, and pre-commit.
- GitHub Actions for the checks and tests, and Dependabot.

### Changed

- Dependencies are managed with uv in `pyproject.toml` and `uv.lock`, in place of `requirements.txt`.

### Removed

- Support for old `.xls` Excel files, which needed a library the project didn't include. `.xlsx` files still work.
